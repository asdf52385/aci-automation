"""APIC REST API 로그인/세션 모듈.

- aaaLogin 으로 로그인하고 APIC-cookie 토큰을 세션에 보관한다.
- 토큰 만료 전에 aaaRefresh 로 자동 갱신한다.
- dry_run=True 이면 APIC 에 접속하지 않고 보낼 요청(JSON)만 출력한다.
"""

import json
import logging
import time

import requests
import urllib3

log = logging.getLogger(__name__)


class ApicError(Exception):
    """APIC 이 에러를 돌려줬을 때."""


class ApicClient:
    # 만료 이 초 전부터는 요청 전에 토큰을 갱신한다
    REFRESH_MARGIN = 60

    def __init__(self, host, username, password=None, verify_ssl=True,
                 timeout=30, dry_run=False, printer=print, proxy=None):
        self.base = host.rstrip("/")
        if not self.base.startswith(("http://", "https://")):
            self.base = "https://" + self.base
        self.username = username
        self.password = password
        self.timeout = timeout
        self.dry_run = dry_run
        self.printer = printer
        self.sent = []  # dry-run 에서 보냈을 요청 기록 (method, path, payload)

        self.session = requests.Session()
        self.session.verify = verify_ssl
        if proxy:
            # 예: http://10.0.0.1:8080 또는 http://user:pass@10.0.0.1:8080
            if "://" not in proxy:
                proxy = "http://" + proxy
            self.session.proxies = {"http": proxy, "https": proxy}
            self.session.trust_env = False  # 환경변수 프록시 설정보다 이 값을 우선
        if not verify_ssl:
            urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

        self._token = None
        self._expires_at = 0.0

    # ---------- 로그인 ----------

    def login(self):
        if self.dry_run:
            self.printer(f"[dry-run] POST {self.base}/api/aaaLogin.json (user={self.username})")
            return
        if not self.password:
            raise ApicError("비밀번호가 없습니다. YAML의 apic.password 또는 APIC_PASSWORD 환경변수를 지정하세요.")
        body = {"aaaUser": {"attributes": {"name": self.username, "pwd": self.password}}}
        data = self._raw("POST", "/api/aaaLogin.json", body)
        self._store_token(data, "aaaLogin")
        log.info("APIC 로그인 성공: %s", self.base)

    def refresh(self):
        data = self._raw("GET", "/api/aaaRefresh.json")
        self._store_token(data, "aaaLogin")
        log.debug("APIC 토큰 갱신")

    def logout(self):
        if self.dry_run or not self._token:
            return
        body = {"aaaUser": {"attributes": {"name": self.username}}}
        try:
            self._raw("POST", "/api/aaaLogout.json", body)
        finally:
            self._token = None
            self.session.cookies.clear()

    def __enter__(self):
        self.login()
        return self

    def __exit__(self, *exc):
        self.logout()

    def _store_token(self, data, cls):
        attrs = data["imdata"][0][cls]["attributes"]
        self._token = attrs["token"]
        self._expires_at = time.time() + int(attrs.get("refreshTimeoutSeconds", 600))
        self.session.cookies.set("APIC-cookie", self._token)

    def _ensure_session(self):
        if self.dry_run:
            return
        if not self._token:
            self.login()
        elif time.time() > self._expires_at - self.REFRESH_MARGIN:
            try:
                self.refresh()
            except ApicError:
                self.login()

    # ---------- 요청 ----------

    def get(self, path, params=None):
        """조회. dry-run 에서도 APIC 에 접속하지 않으므로 빈 결과를 돌려준다."""
        if self.dry_run:
            self.printer(f"[dry-run] GET {self.base}{path} params={params or {}}")
            return {"totalCount": "0", "imdata": []}
        self._ensure_session()
        return self._raw("GET", path, params=params)

    def post(self, path, payload):
        """생성/변경. dry-run 이면 페이로드만 출력한다."""
        self.sent.append(("POST", path, payload))
        if self.dry_run:
            self.printer(f"[dry-run] POST {self.base}{path}")
            self.printer(json.dumps(payload, indent=2, ensure_ascii=False))
            return {"totalCount": "0", "imdata": []}
        self._ensure_session()
        return self._raw("POST", path, payload)

    def post_mo(self, payload, dn="uni"):
        """MO 트리를 /api/mo/{dn}.json 에 POST."""
        return self.post(f"/api/mo/{dn}.json", payload)

    def _raw(self, method, path, body=None, params=None):
        url = self.base + path
        try:
            resp = self.session.request(method, url, json=body, params=params, timeout=self.timeout)
        except requests.RequestException as e:
            raise ApicError(f"{method} {url} 연결 실패: {e}") from e
        try:
            data = resp.json()
        except ValueError:
            raise ApicError(f"{method} {url} -> HTTP {resp.status_code}, JSON 아님: {resp.text[:200]}")
        if resp.status_code >= 400:
            raise ApicError(f"{method} {url} -> HTTP {resp.status_code}: {_error_text(data)}")
        return data


def _error_text(data):
    try:
        attrs = data["imdata"][0]["error"]["attributes"]
        return f"[{attrs.get('code')}] {attrs.get('text')}"
    except (KeyError, IndexError, TypeError):
        return json.dumps(data)[:300]
