#!/usr/bin/env python3
"""ACI 자동화 실행기.

  python3 run.py inventory/example.yaml            # dry-run (기본): 보낼 JSON만 출력
  python3 run.py inventory/example.yaml --apply    # 실제 APIC 에 적용
  python3 run.py inventory/example.yaml --validate # 입력 파일 검증만
  python3 run.py inventory/aci.xlsx                  # 엑셀: static port 만 dry-run
  python3 run.py inventory/aci.xlsx --vlan 100       # VLAN 100 의 static port 만
  python3 run.py inventory/aci.xlsx --init-settings  # 엑셀에 Settings 시트 추가

엑셀 입력은 기본으로 static port 만 넣는다 (VRF/BD/AP/EPG 는 이미 APIC 에 있다고 본다).
--build-all 을 붙이면 테넌트 시트로 VRF/BD/AP/EPG/Contract 까지 만든다.

인자 없이 실행하면(더블클릭 포함) 엑셀 파일과 작업을 메뉴로 고르고, 끝나면 창이 바로 닫히지 않는다.
"""

import argparse
import getpass
import glob
import logging
import os
import re
import sys
import traceback

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
# 폐쇄망용: lib 폴더가 있으면 거기서도 모듈을 찾음 (pip install -r requirements.txt --target lib)
# 설치된 게 있으면 그걸 먼저 씀
sys.path.append(os.path.join(HERE, "lib"))


def main(argv=None):
    from aci_auto import builders, config, excel
    from aci_auto.client import ApicClient, ApicError

    ap = argparse.ArgumentParser(description="APIC REST API 로 ACI 오브젝트 생성")
    ap.add_argument("inventory", help="입력 엑셀(.xlsx) 또는 YAML 파일")
    mode = ap.add_mutually_exclusive_group()
    mode.add_argument("--apply", action="store_true", help="실제 APIC 에 적용 (없으면 dry-run)")
    mode.add_argument("--validate", action="store_true", help="입력 파일 검증만 하고 끝냄")
    mode.add_argument("--init-settings", action="store_true", help="엑셀에 Settings 시트를 기본값으로 추가")
    ap.add_argument("--tenant", action="append", help="이 테넌트만 처리 (여러 번 지정 가능)")
    ap.add_argument("--vlan", type=int, action="append",
                    help="이 VLAN 의 EPG 와 관련 BD/VRF/Contract 만 처리 (여러 번 지정 가능)")
    ap.add_argument("--build-all", action="store_true",
                    help="엑셀 입력에서 static port 뿐 아니라 VRF/BD/AP/EPG/Contract 도 생성")
    ap.add_argument("-v", "--verbose", action="store_true")
    args = ap.parse_args(argv)

    logging.basicConfig(level=logging.DEBUG if args.verbose else logging.INFO,
                        format="%(levelname)s %(message)s")

    if args.init_settings:
        if excel.init_settings(args.inventory):
            print(f"{args.inventory} 에 Settings 시트를 추가했습니다. 값을 채운 뒤 다시 실행하세요.")
        else:
            print("Settings 시트가 이미 있습니다.")
        return 0

    try:
        cfg = config.load(args.inventory)
    except config.ConfigError as e:
        print(e, file=sys.stderr)
        return 2
    print(f"입력 파일 검증 통과: 테넌트 {len(cfg['tenants'])}개")
    if args.validate:
        return 0

    tenants = cfg["tenants"]
    if args.tenant:
        tenants = [t for t in tenants if t["name"] in args.tenant]
        missing = set(args.tenant) - {t["name"] for t in tenants}
        if missing:
            print(f"입력 파일에 없는 테넌트: {', '.join(sorted(missing))}", file=sys.stderr)
            return 2
    if args.vlan:
        tenants = [excel.only_vlans(t, args.vlan) for t in tenants]
        tenants = [t for t in tenants if t["app_profiles"]]
        if not tenants:
            print(f"VLAN {args.vlan} 에 해당하는 EPG 가 없습니다", file=sys.stderr)
            return 2

    apic = cfg["apic"]
    if args.apply and not apic.get("password"):
        # APIC_PASSWORD 환경변수가 없으면 실행할 때 물어본다 (화면에 안 보임)
        apic["password"] = getpass.getpass(f"APIC 비밀번호 ({apic['username']}@{apic['host']}): ")
    client = ApicClient(
        host=apic["host"],
        username=apic["username"],
        password=apic.get("password"),
        verify_ssl=apic.get("verify_ssl", True),
        timeout=apic.get("timeout", 30),
        proxy=apic.get("proxy"),
        dry_run=not args.apply,
    )
    if client.dry_run:
        print("=== DRY-RUN: APIC 에 접속하지 않습니다 ===")

    static_only = bool(cfg.get("excel")) and not args.build_all
    try:
        with client:
            if static_only:
                return _static_ports(client, tenants)
            for tn in tenants:
                print(f"\n--- 테넌트 {tn['name']}: {_summary(tn)} ---")
                client.post_mo(builders.tenant(tn))
    except ApicError as e:
        print(f"APIC 오류: {e}", file=sys.stderr)
        return 1

    print(f"\n완료: 요청 {len(client.sent)}건" + (" (dry-run)" if client.dry_run else " 적용"))
    return 0


def _static_ports(client, tenants):
    from aci_auto import builders, excel
    """이미 있는 EPG 에 static port(fvRsPathAtt) 만 추가한다."""
    jobs = [j for tn in tenants for j in excel.static_port_jobs(tn)]
    if not jobs:
        print("static port 시트에 넣을 행이 없습니다.")
        return 0
    print(f"\n--- static port: EPG {len(jobs)}개, 포트 {sum(len(p) for _, p in jobs)}건 ---")

    # 없는 EPG 에 POST 하면 BD 없는 빈 EPG 가 생기므로, 적용 전에 전부 확인하고 하나라도 없으면 중단
    if client.dry_run:
        print("[dry-run] 적용할 때는 먼저 EPG 가 APIC 에 있는지 확인하고, 없으면 아무것도 넣지 않습니다")
    else:
        missing = [dn for dn, _ in jobs if int(client.get(f"/api/mo/{dn}.json")["totalCount"]) == 0]
        if missing:
            print("APIC 에 없는 EPG 가 있어 중단합니다 (아무것도 넣지 않음):", file=sys.stderr)
            for dn in missing:
                print(f"  - {dn}", file=sys.stderr)
            return 1
        bad = _resolve_policy_paths(client, jobs)
        if bad:
            print("PC/vPC 포트 경로를 APIC 에서 찾지 못해 중단합니다 (아무것도 넣지 않음):", file=sys.stderr)
            for b in bad:
                print(f"  - {b}", file=sys.stderr)
            return 1
        conflicts = _vlan_conflicts(client, jobs)
        if conflicts:
            print("다른 EPG 가 이미 같은 포트에 같은 VLAN 을 쓰고 있어 중단합니다 (아무것도 넣지 않음):",
                  file=sys.stderr)
            for c in conflicts:
                print(f"  - {c}", file=sys.stderr)
            return 1

    for dn, paths in jobs:
        client.post_mo({"fvAEPg": {"attributes": {"dn": dn},
                                   "children": [builders.static_path(p) for p in paths]}}, dn=dn)
    print(f"\n완료: 요청 {len(client.sent)}건" + (" (dry-run)" if client.dry_run else " 적용"))
    return 0


def _resolve_policy_paths(client, jobs):
    """PC/vPC(Policy Group 이름) 포트는 APIC 에서 실제 경로(fabricPathEp)를 찾아 tdn 으로 넣는다.
    Target Node 에 노드를 하나만 적어도 그 노드가 포함된 vPC 경로를 찾는다."""
    bad, cache = [], {}
    for _, paths in jobs:
        for p in paths:
            if p["port"].startswith("eth"):
                continue
            name = p["port"]
            if name not in cache:
                res = client.get("/api/node/class/fabricPathEp.json",
                                 params={"query-target-filter": f'eq(fabricPathEp.name,"{name}")'})
                cache[name] = [m["fabricPathEp"]["attributes"]["dn"] for m in res.get("imdata", [])]
            nodes = [int(x) for x in str(p["node"]).split("-")]
            pod = f"topology/pod-{p.get('pod', 1)}/"
            hits = [dn for dn in cache[name] if dn.startswith(pod)
                    and all(n in [int(x) for x in re.findall(r"\d+", dn.split("/")[2])] for n in nodes)]
            if len(hits) == 1:
                p["tdn"] = hits[0]
            elif not hits:
                bad.append(f"node {p['node']} {name}: 이 이름의 PC/vPC 포트가 없습니다 "
                           f"(APIC 에서 찾은 경로: {', '.join(cache[name]) or '없음'})")
            else:
                bad.append(f"node {p['node']} {name}: 경로가 여러 개라 고를 수 없습니다 ({', '.join(hits)})")
    return bad


def _vlan_conflicts(client, jobs):
    """포트별로 APIC 에 이미 있는 static port 를 조회해, 같은 VLAN 이 다른 EPG 에 붙어 있으면 알려준다."""
    from aci_auto import builders
    wanted = {}  # 포트 tDn -> [(encap, 넣을 EPG dn, 행 설명)]
    for dn, paths in jobs:
        for p in paths:
            att = builders.static_path(p)["fvRsPathAtt"]["attributes"]
            wanted.setdefault(att["tDn"], []).append((att["encap"], dn, f"node {p['node']} {p['port']} vlan {p['vlan']}"))
    out = []
    for tdn, items in wanted.items():
        res = client.get("/api/node/class/fvRsPathAtt.json",
                         params={"query-target-filter": f'eq(fvRsPathAtt.tDn,"{tdn}")'})
        for mo in res.get("imdata", []):
            a = mo["fvRsPathAtt"]["attributes"]
            owner = a["dn"].split("/rspathAtt-")[0]
            for encap, dn, desc in items:
                if a.get("encap") == encap and owner != dn:
                    out.append(f"{desc}: 이미 {owner} 에 붙어 있음")
    return out


def _summary(tn):
    epgs = [e for a in tn.get("app_profiles") or [] for e in a.get("epgs") or []]
    paths = sum(len(e.get("static_paths") or []) for e in epgs)
    return (f"VRF {len(tn.get('vrfs') or [])}, BD {len(tn.get('bds') or [])}, "
            f"AP {len(tn.get('app_profiles') or [])}, EPG {len(epgs)}, "
            f"Contract {len(tn.get('contracts') or [])}, static port {paths}")


class _Back(Exception):
    """메뉴에서 b 를 눌러 이전 단계로."""


def _ask(prompt, choices, default, back=True):
    """choices 안의 값이 나올 때까지 다시 묻는다. q 는 종료, b 는 이전 단계."""
    extra = "b: 뒤로, q: 종료" if back else "q: 종료"
    while True:
        v = input(f"{prompt} [{default}] ({extra}): ").strip().lower() or default
        if v in choices:
            return v
        if v == "q":
            raise KeyboardInterrupt
        if v == "b" and back:
            raise _Back
        print(f"  '{v}' 는 없는 선택입니다. {', '.join(choices)} 중에서 고르세요.")


def interactive():
    """인자 없이 실행했을 때: 엑셀 파일과 작업을 고르게 하고 argv 를 만든다."""
    files = sorted(set(glob.glob(os.path.join(HERE, "*.xlsx")) +
                       glob.glob(os.path.join(HERE, "inventory", "*.xlsx"))))
    files = [f for f in files if not os.path.basename(f).startswith("~$")]  # 엑셀 임시파일 제외
    if not files:
        print(f"엑셀 파일이 없습니다. {HERE} 또는 inventory 폴더에 .xlsx 를 넣어 주세요.")
        raise KeyboardInterrupt
    step, path = "file", None
    while True:
        try:
            if step == "file":
                print("\n엑셀 파일:")
                for i, f in enumerate(files, 1):
                    print(f"  {i}) {os.path.relpath(f, HERE)}")
                n = _ask("번호 선택", [str(i) for i in range(1, len(files) + 1)], "1", back=False)
                path, step = files[int(n) - 1], "job"
            elif step == "job":
                print(f"\n[{os.path.basename(path)}] 작업:")
                print("  1) 검증만\n  2) dry-run (APIC 에 안 보내고 JSON 만 출력)\n  3) 실제 적용 (--apply)")
                job = _ask("번호 선택", ["1", "2", "3"], "2")
                if job != "3":
                    return [path] + {"1": ["--validate"], "2": []}[job]
                step = "confirm"
            else:
                if _ask("정말 APIC 에 적용할까요? (y/n)", ["y", "n"], "n") == "y":
                    return [path, "--apply"]
                print("취소했습니다.")
                step = "job"
        except _Back:
            step = {"job": "file", "confirm": "job"}[step]


def _run(argv, double_click):
    """main 을 실행하고 어떤 오류가 나도 창이 닫히지 않게 메시지만 보여준다."""
    try:
        return main(argv)
    except ImportError as e:
        req = os.path.join(HERE, "requirements.txt")
        print(f"필요한 모듈이 없습니다: {e}")
        if double_click and input("지금 설치할까요? (Y/n): ").strip().lower() in ("", "y"):
            import subprocess
            ok = subprocess.call([sys.executable, "-m", "pip", "install", "-r", req]) == 0
            print("\n설치 완료. run.py 를 다시 실행하세요." if ok else "\n설치에 실패했습니다. 위 메시지를 확인하세요.")
        else:
            print(f"명령 프롬프트에서 실행하세요: \"{sys.executable}\" -m pip install -r \"{req}\"")
    except KeyboardInterrupt:
        print("\n중단했습니다.")
    except SystemExit as e:
        return e.code
    except Exception:
        traceback.print_exc()
    return 1


if __name__ == "__main__":
    if len(sys.argv) > 1:
        sys.exit(_run(None, False))
    # 더블클릭: 메뉴 -> 실행 -> 다시 메뉴. q 를 입력하면 끝
    try:
        while True:
            _run(interactive(), True)
            if input("\n엔터: 처음 메뉴로 / q: 종료 > ").strip().lower() == "q":
                break
    except (KeyboardInterrupt, EOFError):
        print("\n종료합니다.")
