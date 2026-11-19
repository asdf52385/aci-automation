"""엑셀(xlsx) 입력 읽기.

시트를 YAML 과 같은 tenants 구조(dict)로 바꾼다. 검증은 config.validate 가 이어서 한다.
열은 위치가 아니라 1행 머리글 이름으로 찾으므로 열 순서를 바꿔도 된다.

테넌트 시트 (한 행 = 서브넷 하나, BD 에 서브넷이 여럿이면 같은 BD 를 여러 행에 적음)
  VRF | Bridge Domain | Unicast Route | Subnet | Encap Vlan | Application Profile | Contract | Endpoint Group

static port 시트 (한 행 = 포트 하나에 VLAN 하나)
  AP | Endpoint Group | Target Node | Port | Vlan | Mode
  AP / Endpoint Group 이 비어 있으면 Vlan 으로 테넌트 시트에서 찾는다 (VLOOKUP 과 같음).

Settings 시트 (항목 | 값 | 설명): APIC 접속 정보, 테넌트 이름, 기본 옵션.
비밀번호는 시트에 적지 않고 APIC_PASSWORD 환경변수로 넘긴다.
"""

import os
import re

import openpyxl

TENANT_COLS = {
    "vrf": "VRF",
    "bd": "Bridge Domain",
    "unicast": "Unicast Route",
    "subnet": "Subnet",
    "vlan": "Encap Vlan",
    "ap": "Application Profile",
    "contract": "Contract",
    "epg": "Endpoint Group",
}
DEPLOY_COLS = {
    "ap": "AP",
    "epg": "Endpoint Group",
    "node": "Target Node",
    "port": "Port",
    "vlan": "Vlan",
    "mode": "Mode",
}
# 엑셀 Mode 값 -> APIC fvRsPathAtt mode
MODES = {"tagged": "regular", "regular": "regular", "trunk": "regular",
         "untagged": "untagged", "access": "untagged",
         "native": "native", "802.1p": "native"}
PORT_RE = re.compile(r"^eth\d+/\d+(/\d+)?$")
# PC/vPC 는 Port 칸에 Interface Policy Group 이름을 적는다 (예: vpc_leaf101_102)
POLICY_RE = re.compile(r"^[A-Za-z0-9_.:-]{1,64}$")
NODE_PAIR_RE = re.compile(r"^(\d+)\s*[-,/~_ ]\s*(\d+)$")

DEFAULTS = {
    "pod": 1,
    "immediacy": "immediate",
    "phys_domain": None,
    "subnet_scope": ["private"],
    "l2_bd": {"arp_flooding": True, "l2_unknown_unicast": "flood"},
    "contract": {"role": "both", "filter": "default", "scope": "context"},
}


SETTINGS_SHEET = "Settings"
# (항목, 기본값, 설명) - Settings 시트를 새로 만들 때 이 순서로 채운다
SETTINGS = [
    ("apic_host", "https://apic.example.local", "APIC 주소"),
    ("apic_username", "admin", "APIC 계정 (비밀번호는 APIC_PASSWORD 환경변수)"),
    ("verify_ssl", False, "사설 인증서면 false"),
    ("timeout", 30, "요청 제한 시간(초)"),
    ("proxy", None, "프록시 주소. SecureCRT SOCKS 포워딩이면 socks5h://127.0.0.1:포트, HTTP 프록시면 http://IP:포트. 비우면 직접 접속"),
    ("tenant", "TN-EXAMPLE", "EPG 가 있는 테넌트 이름 (EPG 경로 uni/tn-<이름>/ap-<AP>/epg-<EPG> 에 필요)"),
    ("deploy_sheet", "Deploy_int_excute", "넣을 static port 시트"),
    ("tenant_sheet", "tenant_all_s", "AP/EPG 가 비어 있을 때 Vlan 으로 찾아볼 시트 (VLOOKUP 과 같음)"),
    ("pod", 1, "포트 경로 topology/pod-<번호>/paths-<노드>/pathep-[eth1/1] 에 필요"),
    ("immediacy", "immediate", "static port 배포 시점: immediate(바로 leaf 에 배포) / lazy(On Demand, 단말이 보일 때 배포)"),
]


def read_settings(path):
    """xlsx 의 Settings 시트를 config 구조({apic, excel})로 바꾼다. (cfg, 문제 목록)"""
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
    except (OSError, ValueError, KeyError) as e:
        return None, [f"엑셀 파일을 열 수 없습니다: {path} ({e})"]
    if SETTINGS_SHEET not in wb.sheetnames:
        wb.close()
        return None, [f"'{SETTINGS_SHEET}' 시트가 없습니다. "
                      f"python3 run.py {path} --init-settings 로 만들 수 있습니다"]
    known = {k for k, _, _ in SETTINGS}
    s, p = {}, []
    for rn, row in enumerate(wb[SETTINGS_SHEET].iter_rows(min_row=2, values_only=True), start=2):
        key = _cell(row[0]) if row else None
        if key is None:
            continue
        if key not in known:
            p.append(f"{SETTINGS_SHEET} {rn}행: 모르는 항목 '{key}'")
            continue
        s[key] = _cell(row[1]) if len(row) > 1 else None
    wb.close()
    defaults = {k: v for k, v, _ in SETTINGS}
    get = lambda k: s.get(k, defaults[k])  # noqa: E731

    def as_bool(k):
        v = _bool(get(k))
        if v is None:
            p.append(f"{SETTINGS_SHEET}: {k} 는 true/false 입니다")
        return v

    cfg = {
        "apic": {
            "host": get("apic_host"),
            "username": get("apic_username"),
            "verify_ssl": as_bool("verify_ssl"),
            "timeout": _int(get("timeout")) or 30,
            "proxy": get("proxy"),
        },
        "excel": {
            "file": os.path.abspath(path),
            "tenant": get("tenant"),
            "tenant_sheet": get("tenant_sheet"),
            "deploy_sheet": get("deploy_sheet"),
            "pod": _int(get("pod")) or 1,
            "immediacy": get("immediacy"),
        },
    }
    return cfg, p


def init_settings(path):
    """Settings 시트가 없으면 기본값으로 만들어 넣는다. 다른 시트는 건드리지 않는다."""
    wb = openpyxl.load_workbook(path)
    if SETTINGS_SHEET in wb.sheetnames:
        return False
    ws = wb.create_sheet(SETTINGS_SHEET, 0)
    ws.append(["항목", "값", "설명"])
    for key, val, desc in SETTINGS:
        ws.append([key, "true" if val is True else "false" if val is False else val, desc])
    ws.column_dimensions["A"].width = 20
    ws.column_dimensions["B"].width = 30
    ws.column_dimensions["C"].width = 70
    wb.save(path)
    return True


def load(opts, base_dir="."):
    """settings(YAML 의 excel: 섹션)를 받아 (tenant dict, 문제 목록) 을 돌려준다."""
    p = []
    opts = {**DEFAULTS, **(opts or {})}
    contract_opts = {**DEFAULTS["contract"], **(opts.get("contract") or {})}
    for key in ("file", "tenant", "tenant_sheet"):
        if not opts.get(key):
            p.append(f"excel.{key}: 값이 필요합니다")
    if contract_opts["role"] not in ("both", "provider", "consumer"):
        p.append("excel.contract.role 은 both / provider / consumer 입니다")
    if p:
        return None, p

    path = opts["file"]
    if not os.path.isabs(path):
        path = os.path.join(base_dir, path)
    try:
        # data_only: 수식 셀은 엑셀이 마지막으로 저장한 값을 읽는다
        wb = openpyxl.load_workbook(path, data_only=True)
    except (OSError, ValueError, KeyError) as e:
        return None, [f"엑셀 파일을 열 수 없습니다: {path} ({e})"]

    tenant = {"name": opts["tenant"], "vrfs": [], "bds": [], "app_profiles": [], "contracts": []}
    by_vlan = _tenant_sheet(p, wb, opts, contract_opts, tenant)
    if opts.get("deploy_sheet"):
        _deploy_sheet(p, wb, opts, tenant, by_vlan)
    wb.close()
    return tenant, p


def epg_dn(tenant, ap, epg):
    return f"uni/tn-{tenant}/ap-{ap}/epg-{epg}"


def static_port_jobs(tenant):
    """static port 가 있는 EPG 마다 (EPG dn, 포트 목록)."""
    return [(epg_dn(tenant["name"], ap["name"], e["name"]), e["static_paths"])
            for ap in tenant.get("app_profiles") or []
            for e in ap["epgs"] if e.get("static_paths")]


def only_vlans(tenant, vlans):
    """지정한 VLAN 의 EPG 와 그것이 쓰는 BD / VRF / Contract 만 남긴 테넌트."""
    vlans = set(vlans)
    aps, bds, ctrs = [], set(), set()
    for ap in tenant.get("app_profiles") or []:
        epgs = [e for e in ap["epgs"] if e.get("encap_vlan") in vlans]
        if epgs:
            aps.append({**ap, "epgs": epgs})
            for e in epgs:
                bds.add(e["bd"])
                ctrs.update(e.get("provide", []) + e.get("consume", []))
    bd_list = [b for b in tenant.get("bds") or [] if b["name"] in bds]
    vrfs = {b["vrf"] for b in bd_list}
    return {**tenant,
            "vrfs": [v for v in tenant.get("vrfs") or [] if v["name"] in vrfs],
            "bds": bd_list,
            "app_profiles": aps,
            "contracts": [c for c in tenant.get("contracts") or [] if c["name"] in ctrs]}


# ---------- 테넌트 시트 ----------

def _tenant_sheet(p, wb, opts, contract_opts, tenant):
    sheet = opts["tenant_sheet"]
    rows = _rows(p, wb, sheet, TENANT_COLS)
    vrfs, bds, aps, ctrs = {}, {}, {}, {}
    epg_of = {}     # epg 이름 -> (행번호, 값)
    by_vlan = {}    # vlan -> (ap, epg)

    for rn, r in rows:
        where = f"{sheet} {rn}행"
        missing = [TENANT_COLS[k] for k in ("vrf", "bd", "ap", "epg") if not r[k]]
        if missing:
            p.append(f"{where}: {', '.join(missing)} 값이 비었습니다")
            continue
        unicast = _bool(r["unicast"])
        if r["unicast"] is not None and unicast is None:
            p.append(f"{where}: Unicast Route 는 true/false 입니다 (값: {r['unicast']})")
            continue
        vlan = _int(r["vlan"])
        if r["vlan"] is not None and vlan is None:
            p.append(f"{where}: Encap Vlan '{r['vlan']}' 은 숫자여야 합니다")
            continue

        vrfs.setdefault(r["vrf"], {"name": r["vrf"]})

        bd = bds.get(r["bd"])
        if bd is None:
            bd = bds[r["bd"]] = {"name": r["bd"], "vrf": r["vrf"], "subnets": [], "_row": rn}
            if unicast is not None:
                bd["unicast_routing"] = unicast
            if unicast is False:
                bd.update(opts["l2_bd"] or {})
        elif bd["vrf"] != r["vrf"] or bd.get("unicast_routing") != unicast:
            p.append(f"{where}: BD '{r['bd']}' 의 VRF/Unicast Route 가 {bd['_row']}행과 다릅니다")
        if r["subnet"]:
            bd["subnets"].append({"ip": r["subnet"], "scope": list(opts["subnet_scope"])})

        same = (r["ap"], r["bd"], vlan, r["contract"])
        if r["epg"] in epg_of:
            prev_rn, prev = epg_of[r["epg"]]
            if prev != same:
                p.append(f"{where}: EPG '{r['epg']}' 의 AP/BD/Vlan/Contract 가 {prev_rn}행과 다릅니다")
            continue
        epg_of[r["epg"]] = (rn, same)

        if vlan is not None:
            if vlan in by_vlan and by_vlan[vlan][1] != r["epg"]:
                p.append(f"{where}: Vlan {vlan} 이 EPG '{by_vlan[vlan][1]}' 에도 쓰였습니다")
            by_vlan.setdefault(vlan, (r["ap"], r["epg"]))

        epg = {"name": r["epg"], "bd": r["bd"], "encap_vlan": vlan, "static_paths": []}
        if opts.get("phys_domain"):
            epg["domains"] = [opts["phys_domain"]]
        if r["contract"]:
            ctrs.setdefault(r["contract"], {"name": r["contract"],
                                            "scope": contract_opts["scope"],
                                            "filters": [contract_opts["filter"]]})
            if contract_opts["role"] in ("both", "provider"):
                epg["provide"] = [r["contract"]]
            if contract_opts["role"] in ("both", "consumer"):
                epg["consume"] = [r["contract"]]
        aps.setdefault(r["ap"], {"name": r["ap"], "epgs": []})["epgs"].append(epg)

    for bd in bds.values():
        bd.pop("_row")
    tenant["vrfs"] = list(vrfs.values())
    tenant["bds"] = list(bds.values())
    tenant["app_profiles"] = list(aps.values())
    tenant["contracts"] = list(ctrs.values())
    return by_vlan


# ---------- static port 시트 ----------

def _deploy_sheet(p, wb, opts, tenant, by_vlan):
    sheet = opts["deploy_sheet"]
    epgs = {(ap["name"], e["name"]): e for ap in tenant["app_profiles"] for e in ap["epgs"]}
    for rn, r in _rows(p, wb, sheet, DEPLOY_COLS):
        where = f"{sheet} {rn}행"
        vlan = _int(r["vlan"])
        if vlan is None:
            p.append(f"{where}: Vlan '{r['vlan']}' 이 비었거나 숫자가 아닙니다")
            continue
        node = _node(r["node"])
        if node is None:
            p.append(f"{where}: Target Node '{r['node']}' 형식이 잘못됐습니다 (예: 101, vPC 는 101-102)")
            continue
        raw_port = str(r["port"] or "").strip()
        if PORT_RE.match(raw_port.replace(" ", "").lower()):
            port = raw_port.replace(" ", "").lower()
            if isinstance(node, str):
                p.append(f"{where}: eth 포트에는 노드를 하나만 적습니다 (Target Node '{r['node']}')")
                continue
        elif POLICY_RE.match(raw_port):
            port = raw_port  # PC/vPC Policy Group 이름 (대소문자 유지)
        else:
            p.append(f"{where}: Port '{r['port']}' 형식이 잘못됐습니다 (예: eth1/1 또는 vPC Policy Group 이름)")
            continue
        mode = MODES.get(str(r["mode"] or "tagged").lower())
        if mode is None:
            p.append(f"{where}: Mode '{r['mode']}' 는 tagged / untagged / native 입니다")
            continue

        found = by_vlan.get(vlan)
        if found is None:
            p.append(f"{where}: Vlan {vlan} 이 {opts['tenant_sheet']} 시트에 없습니다")
            continue
        # 셀에 값(또는 VLOOKUP 결과)이 있으면 Vlan 으로 찾은 것과 같은지 확인
        given = (r["ap"], r["epg"])
        if any(v and not str(v).startswith("#") for v in given) and given != found:
            p.append(f"{where}: AP/EPG {given} 가 Vlan {vlan} 의 {found} 와 다릅니다")
            continue
        epgs[found]["static_paths"].append(
            {"pod": opts["pod"], "node": node, "port": port, "vlan": vlan, "mode": mode,
             "immediacy": opts.get("immediacy") or "immediate"})


# ---------- 공통 ----------

def _rows(p, wb, sheet, cols):
    """(행번호, {키: 값}) 목록. 빈 행은 건너뛴다.

    머리글 행은 위쪽 20행 안에서 찾는다. 머리글이 없으면(데이터만 붙여 넣은 경우)
    열 순서가 양식과 같다고 보고 A열부터 차례로 읽는다.
    """
    if sheet not in wb.sheetnames:
        p.append(f"시트 '{sheet}' 가 없습니다 (있는 시트: {', '.join(wb.sheetnames)})")
        return []
    rows = list(wb[sheet].iter_rows(values_only=True))
    want = {k: _norm(t) for k, t in cols.items()}

    best, best_idx = None, {}
    for hn, row in enumerate(rows[:20]):
        names = [_norm(h) for h in row]
        idx = {k: names.index(t) for k, t in want.items() if t in names}
        if len(idx) > len(best_idx):
            best, best_idx = hn, idx
    if len(best_idx) == len(cols):
        start, idx = best + 1, best_idx
    elif len(best_idx) >= 2:
        missing = [cols[k] for k in cols if k not in best_idx]
        p.append(f"시트 '{sheet}' {best + 1}행 머리글에 {', '.join(missing)} 열이 없습니다 "
                 f"(필요한 머리글: {' | '.join(cols.values())})")
        return []
    else:
        start, idx = 0, {k: i for i, k in enumerate(cols)}

    out = []
    for rn, row in enumerate(rows[start:], start=start + 1):
        vals = {k: _cell(row[i]) if i < len(row) else None for k, i in idx.items()}
        if any(v is not None for v in vals.values()):
            out.append((rn, vals))
    return out


def _norm(v):
    """머리글 비교용: 대소문자, 공백, 밑줄 무시."""
    if not isinstance(v, str):
        return None
    return re.sub(r"[\s_\u00a0]+", "", v).lower() or None


def _node(v):
    """101 -> 101, '102-101' -> '101-102' (vPC, 작은 번호 먼저)."""
    n = _int(v)
    if n is not None:
        return n
    m = NODE_PAIR_RE.match(str(v or "").strip())
    if m:
        a, b = sorted(int(x) for x in m.groups())
        return f"{a}-{b}"
    return None


def _cell(v):
    if isinstance(v, str):
        v = v.strip()
        return v or None
    return v


def _bool(v):
    if isinstance(v, bool) or v is None:
        return v
    return {"true": True, "yes": True, "y": True, "1": True,
            "false": False, "no": False, "n": False, "0": False}.get(str(v).lower())


def _int(v):
    if isinstance(v, bool):
        return None
    if isinstance(v, int):
        return v
    if isinstance(v, float) and v.is_integer():
        return int(v)
    if isinstance(v, str) and v.isdigit():
        return int(v)
    return None
