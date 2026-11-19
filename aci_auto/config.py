"""입력 파일(YAML, 또는 YAML 이 가리키는 엑셀) 읽기와 검증.

검증에 실패하면 문제 목록을 한 번에 모아 ConfigError 로 알려준다.
"""

import ipaddress
import os
import re


NAME_RE = re.compile(r"^[A-Za-z0-9_.:-]{1,64}$")
SUBNET_SCOPES = {"private", "public", "shared"}
L2_UNKNOWN = {"proxy", "flood"}
CONTRACT_SCOPES = {"context", "tenant", "global", "application-profile"}
PATH_MODES = {"regular", "untagged", "native"}
PORT_RE = re.compile(r"^eth\d+/\d+(/\d+)?$")
POLICY_RE = re.compile(r"^[A-Za-z0-9_.:-]{1,64}$")  # PC/vPC Policy Group 이름
NODE_RE = re.compile(r"^\d+-\d+$")                   # vPC 노드 쌍

# 다음 단계에서 채울 섹션 (지금은 리스트인지만 확인)
RESERVED_SECTIONS = ("l3outs", "filters")


class ConfigError(Exception):
    def __init__(self, problems):
        self.problems = problems
        super().__init__("입력 파일 오류:\n" + "\n".join(f"  - {p}" for p in problems))


def load(path):
    from . import excel
    if path.lower().endswith((".xlsx", ".xlsm")):
        cfg, problems = excel.read_settings(path)
        if problems:
            raise ConfigError(problems)
    else:
        import yaml  # YAML 입력일 때만 필요 (엑셀만 쓰면 PyYAML 설치 불필요)
        with open(path, encoding="utf-8") as f:
            cfg = yaml.safe_load(f) or {}
    if cfg.get("excel"):
        tenant, problems = excel.load(cfg["excel"], os.path.dirname(os.path.abspath(path)))
        if problems:
            raise ConfigError(problems)
        cfg["tenants"] = (cfg.get("tenants") or []) + [tenant]
    validate(cfg)
    apic = cfg.setdefault("apic", {})
    apic.setdefault("password", os.environ.get("APIC_PASSWORD"))
    return cfg


def validate(cfg):
    p = []

    apic = cfg.get("apic")
    if not isinstance(apic, dict):
        p.append("apic: 섹션이 필요합니다 (host, username)")
    else:
        for key in ("host", "username"):
            if not apic.get(key):
                p.append(f"apic.{key}: 값이 필요합니다")

    tenants = cfg.get("tenants")
    if not isinstance(tenants, list) or not tenants:
        p.append("tenants: 테넌트를 1개 이상 적어야 합니다")
        tenants = []

    paths = {}  # (pod, node, port, vlan) -> 처음 쓴 EPG
    for ti, tn in enumerate(tenants):
        where = f"tenants[{ti}]"
        if not isinstance(tn, dict):
            p.append(f"{where}: 항목 형식이 잘못됐습니다")
            continue
        _name(p, where, tn)
        where = f"tenant {tn.get('name', ti)}"

        vrfs = _list(p, where, tn, "vrfs")
        bds = _list(p, where, tn, "bds")
        aps = _list(p, where, tn, "app_profiles")
        ctrs = _list(p, where, tn, "contracts")
        for sec in RESERVED_SECTIONS:
            _list(p, where, tn, sec)

        vrf_names = _unique_names(p, f"{where}.vrfs", vrfs)
        bd_names = _unique_names(p, f"{where}.bds", bds)
        _unique_names(p, f"{where}.app_profiles", aps)
        ctr_names = _unique_names(p, f"{where}.contracts", ctrs)

        for c in ctrs:
            if c.get("scope", "context") not in CONTRACT_SCOPES:
                p.append(f"{where} Contract {c.get('name')}: scope 는 {sorted(CONTRACT_SCOPES)} 중 하나입니다")

        for v in vrfs:
            if "enforce" in v and not isinstance(v["enforce"], bool):
                p.append(f"{where} VRF {v.get('name')}: enforce 는 true/false 입니다")

        for bd in bds:
            w = f"{where} BD {bd.get('name')}"
            vrf = bd.get("vrf")
            if not vrf:
                p.append(f"{w}: vrf 가 필요합니다")
            elif bd.get("vrf_tenant") not in (None, tn.get("name"), "common"):
                # APIC 은 이름으로 VRF 를 찾을 때 자기 테넌트 다음 common 만 본다
                p.append(f"{w}: vrf_tenant 는 자기 테넌트나 common 만 됩니다")
            elif bd.get("vrf_tenant") in (None, tn.get("name")) and vrf not in vrf_names:
                p.append(f"{w}: VRF '{vrf}' 가 이 테넌트에 없습니다 (다른 테넌트면 vrf_tenant 지정)")
            for key in ("unicast_routing", "arp_flooding"):
                if key in bd and not isinstance(bd[key], bool):
                    p.append(f"{w}: {key} 는 true/false 입니다")
            if bd.get("l2_unknown_unicast", "proxy") not in L2_UNKNOWN:
                p.append(f"{w}: l2_unknown_unicast 는 proxy 또는 flood 입니다")
            for sn in bd.get("subnets") or []:
                _subnet(p, w, sn)

        for ap in aps:
            w = f"{where} AP {ap.get('name')}"
            epgs = _list(p, w, ap, "epgs")
            _unique_names(p, f"{w}.epgs", epgs)
            for epg in epgs:
                we = f"{w} EPG {epg.get('name')}"
                if not epg.get("bd"):
                    p.append(f"{we}: bd 가 필요합니다")
                elif epg["bd"] not in bd_names:
                    p.append(f"{we}: BD '{epg['bd']}' 가 이 테넌트에 없습니다")
                for key in ("provide", "consume"):
                    for c in _list(p, we, epg, key):
                        if c not in ctr_names:
                            p.append(f"{we}: {key} Contract '{c}' 가 이 테넌트 contracts 에 없습니다")
                for sp in _list(p, we, epg, "static_paths"):
                    _static_path(p, we, sp, paths)

    names = [t.get("name") for t in tenants if isinstance(t, dict)]
    for n in {n for n in names if names.count(n) > 1}:
        p.append(f"tenants: 이름 '{n}' 이 중복됩니다")
    if p:
        raise ConfigError(p)


def _name(p, where, obj):
    name = obj.get("name")
    if not name:
        p.append(f"{where}: name 이 필요합니다")
    elif not NAME_RE.match(str(name)):
        p.append(f"{where}: 이름 '{name}' 에 허용되지 않는 문자가 있거나 64자를 넘습니다")


def _list(p, where, obj, key):
    val = obj.get(key)
    if val is None:
        return []
    if not isinstance(val, list):
        p.append(f"{where}.{key}: 리스트여야 합니다")
        return []
    return val


def _unique_names(p, where, items):
    seen = set()
    for i, it in enumerate(items):
        if not isinstance(it, dict):
            p.append(f"{where}[{i}]: 항목 형식이 잘못됐습니다")
            continue
        _name(p, f"{where}[{i}]", it)
        n = it.get("name")
        if n in seen:
            p.append(f"{where}: 이름 '{n}' 이 중복됩니다")
        seen.add(n)
    return seen


def _subnet(p, where, sn):
    ip = sn.get("ip") if isinstance(sn, dict) else None
    try:
        iface = ipaddress.ip_interface(ip)
        if iface.ip == iface.network.network_address and iface.network.prefixlen < iface.max_prefixlen:
            p.append(f"{where}: subnet {ip} 는 네트워크 주소입니다. 게이트웨이 IP를 적으세요 (예: 10.0.0.1/24)")
    except (ValueError, TypeError):
        p.append(f"{where}: subnet ip '{ip}' 형식이 잘못됐습니다 (예: 10.0.0.1/24)")
        return
    scope = sn.get("scope", ["private"])
    if isinstance(scope, str):
        scope = [scope]
    bad = set(scope) - SUBNET_SCOPES
    if bad:
        p.append(f"{where}: subnet {ip} scope {sorted(bad)} 는 허용되지 않습니다 (private/public/shared)")
    if {"private", "public"} <= set(scope):
        p.append(f"{where}: subnet {ip} 는 private 과 public 을 동시에 가질 수 없습니다")


def _static_path(p, where, sp, seen):
    if not isinstance(sp, dict):
        p.append(f"{where}: static_paths 항목 형식이 잘못됐습니다")
        return
    node, vlan, port = sp.get("node"), sp.get("vlan"), str(sp.get("port", ""))
    w = f"{where} static {node}/{port} vlan {vlan}"
    if not (isinstance(node, int) or (isinstance(node, str) and NODE_RE.match(node))):
        p.append(f"{w}: node 는 숫자(예: 101) 또는 vPC 노드 쌍(예: 101-102)입니다")
    if not (PORT_RE.match(port) or POLICY_RE.match(port)):
        p.append(f"{w}: port 는 eth1/1 형식이나 PC/vPC Policy Group 이름입니다")
    if not isinstance(vlan, int) or not 1 <= vlan <= 4094:
        p.append(f"{w}: vlan 은 1~4094 입니다")
    if sp.get("immediacy") not in (None, "immediate", "lazy"):
        p.append(f"{w}: immediacy 는 immediate 또는 lazy 입니다")
    if sp.get("mode", "regular") not in PATH_MODES:
        p.append(f"{w}: mode 는 regular / untagged / native 입니다")
    key = (sp.get("pod", 1), node, port, vlan)
    if key in seen:
        p.append(f"{w}: 같은 포트·VLAN 이 {seen[key]} 에도 있습니다")
    seen.setdefault(key, where)
