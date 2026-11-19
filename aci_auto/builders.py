"""YAML 항목을 APIC MO(JSON) 페이로드로 바꾸는 함수들.

Tenant / VRF / BD / AP·EPG (static port 포함) / Contract
L3Out 빌더는 다음 단계에서 추가한다.
"""


def mo(cls, attrs, children=None):
    """APIC MO 한 개. 값이 None 인 속성은 빼서 APIC 기본값을 쓰게 한다."""
    body = {"attributes": {k: v for k, v in attrs.items() if v is not None}}
    if children:
        body["children"] = children
    return {cls: body}


def yn(value):
    return None if value is None else ("yes" if value else "no")


def vrf(spec):
    enforce = spec.get("enforce")
    return mo("fvCtx", {
        "name": spec["name"],
        "descr": spec.get("description"),
        "pcEnfPref": None if enforce is None else ("enforced" if enforce else "unenforced"),
    })


def subnet(spec):
    scope = spec.get("scope", ["private"])
    if isinstance(scope, str):
        scope = [scope]
    return mo("fvSubnet", {
        "ip": spec["ip"],
        "descr": spec.get("description"),
        "scope": ",".join(scope),
    })


def bd(spec):
    children = [mo("fvRsCtx", {"tnFvCtxName": spec["vrf"]})]
    children += [subnet(s) for s in spec.get("subnets") or []]
    return mo("fvBD", {
        "name": spec["name"],
        "descr": spec.get("description"),
        "unicastRoute": yn(spec.get("unicast_routing")),
        "arpFlood": yn(spec.get("arp_flooding")),
        "unkMacUcastAct": spec.get("l2_unknown_unicast"),
    }, children)


def path_dn(spec):
    """포트 경로. eth -> paths-노드, PC -> paths-노드/Policy Group, vPC -> protpaths-노드1-노드2/Policy Group.
    APIC 에서 실제 경로를 조회해 넣은 tdn 이 있으면 그걸 쓴다."""
    if spec.get("tdn"):
        return spec["tdn"]
    pod, node = spec.get("pod", 1), spec["node"]
    kind = "protpaths" if isinstance(node, str) and "-" in node else "paths"
    return f"topology/pod-{pod}/{kind}-{node}/pathep-[{spec['port']}]"


def static_path(spec):
    return mo("fvRsPathAtt", {
        "tDn": path_dn(spec),
        "encap": f"vlan-{spec['vlan']}",
        "mode": spec.get("mode"),
        "instrImedcy": spec.get("immediacy"),
    })


def epg(spec):
    children = [mo("fvRsBd", {"tnFvBDName": spec["bd"]})]
    children += [mo("fvRsDomAtt", {"tDn": f"uni/phys-{d}"}) for d in spec.get("domains") or []]
    children += [mo("fvRsProv", {"tnVzBrCPName": c}) for c in spec.get("provide") or []]
    children += [mo("fvRsCons", {"tnVzBrCPName": c}) for c in spec.get("consume") or []]
    children += [static_path(sp) for sp in spec.get("static_paths") or []]
    return mo("fvAEPg", {"name": spec["name"], "descr": spec.get("description")}, children)


def app_profile(spec):
    return mo("fvAp", {"name": spec["name"], "descr": spec.get("description")},
              [epg(e) for e in spec.get("epgs") or []])


def contract(spec):
    """subject 하나에 filter 를 건다. filter 이름은 자기 테넌트, 없으면 common 에서 찾는다."""
    filters = spec.get("filters") or ["default"]
    subj = mo("vzSubj", {"name": spec.get("subject", "default")},
              [mo("vzRsSubjFiltAtt", {"tnVzFilterName": f}) for f in filters])
    return mo("vzBrCP", {
        "name": spec["name"],
        "descr": spec.get("description"),
        "scope": spec.get("scope"),
    }, [subj])


def tenant(spec):
    """테넌트 트리 전체 (VRF, BD, Contract, AP/EPG)."""
    name = spec["name"]
    children = [vrf(v) for v in spec.get("vrfs") or []]
    children += [bd(b) for b in spec.get("bds") or []]
    children += [contract(c) for c in spec.get("contracts") or []]
    children += [app_profile(a) for a in spec.get("app_profiles") or []]
    return mo("fvTenant", {"name": name, "descr": spec.get("description")}, children)
