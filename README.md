# ACI static port 자동화 (엑셀 → APIC REST API)

엑셀 시트에 적은 포트를 **이미 있는 EPG**에 static port(`fvRsPathAtt`)로 넣어 줍니다.
VRF / BD / AP / EPG / Contract는 건드리지 않습니다.

## 준비물

- Python 3.8 이상 (Windows는 python.org 설치본이면 됨)
- 필요한 모듈 설치 (requests, openpyxl, PySocks, PyYAML)

  ```
  pip install -r requirements.txt
  ```

  모듈이 없는 상태에서 `run.py`를 더블클릭하면 설치할지 물어봅니다.

### 폐쇄망 PC에서 쓰기

인터넷이 되는 PC에서 모듈을 `lib` 폴더로 받은 뒤, 프로젝트 폴더째 복사하면 설치 없이 실행됩니다.
(`run.py`가 `lib` 폴더를 자동으로 찾습니다.)

```
pip install -r requirements.txt --target lib
```

## 실행

`run.py`를 **더블클릭**합니다.

```
=== APIC 로그인: apic#인증서버\계정 @ https://ipaddr (proxy socks5h://ip:port) ===
비밀번호 (그냥 엔터: 로그인 없이 진행):
로그인 성공. 창을 닫기 전까지 비밀번호를 다시 묻지 않습니다.

엑셀 파일:
  1) inventory\대상엑셀파일.xlsx
번호 선택 [1] (q: 종료):

[대상엑셀파일.xlsx] 작업:
  1) 검증만
  2) dry-run (APIC 에 안 보내고 JSON 만 출력)
  3) 실제 적용 (--apply)
번호 선택 [2] (b: 뒤로, q: 종료):
```

- 처음에 APIC 로그인을 한 번 해 보고, 비밀번호는 창을 닫을 때까지 메모리에만 기억합니다 (파일 저장 안 함).
  틀리면 그 자리에서 다시 묻고, 엔터만 치면 로그인 없이 검증/dry-run만 할 수 있습니다.
- 엑셀은 `run.py` 옆이나 `inventory` 폴더의 `.xlsx`를 자동으로 찾습니다.
- 잘못 입력하면 다시 묻고, `b`는 이전 단계, `q`는 종료입니다. 작업이 끝나면 처음 메뉴로 돌아갑니다.
- **실제 적용 전에 항상 dry-run으로 먼저 확인**하는 것을 권장합니다.

명령 프롬프트에서 직접 실행할 수도 있습니다.

```bat
python run.py inventory\대상엑셀파일.xlsx --validate   :: 검증만
python run.py inventory\대상엑셀파일.xlsx              :: dry-run
python run.py inventory\대상엑셀파일.xlsx --apply      :: 실제 적용
python run.py inventory\대상엑셀파일.xlsx --vlan 100   :: VLAN 100 행만
python run.py 새파일.xlsx --init-settings                    :: Settings 시트 추가
```

## 엑셀 구성

| 시트 | 내용 |
|---|---|
| `Settings` | APIC 접속 정보와 옵션 (아래 표) |
| 테넌트 시트 (예: `tenant_all_s`) | VRF, Bridge Domain, Unicast Route, Subnet, **Encap Vlan**, **Application Profile**, Contract, **Endpoint Group**. Vlan으로 AP/EPG를 찾는 데 사용 |
| `Deploy_int_excute` | 넣을 포트 목록: **AP, Endpoint Group, Target Node, Port, Vlan, Mode** |

### Settings 시트

| 항목 | 예시 | 설명 |
|---|---|---|
| apic_host | `https://ipaddr` | APIC 주소 (GUI 주소와 같음) |
| apic_username | `apic#인증서버\계정` | 로그인 도메인이 있으면 `apic#도메인\아이디`. 로컬 계정은 `apic#fallback\admin` |
| verify_ssl | `false` | 사설 인증서면 false |
| timeout | `30` | 요청 제한 시간(초) |
| proxy | `socks5h://ip:port` | SecureCRT SOCKS 포워딩을 거칠 때. HTTP 프록시는 `http://IP:포트`. 비우면 직접 접속 |
| tenant | `TENANT` | EPG가 있는 테넌트 이름 |
| deploy_sheet | `Deploy_int_excute` | 넣을 포트 시트 이름 |
| tenant_sheet | `tenant_all_s` | Vlan → AP/EPG를 찾을 시트 이름 |
| pod | `1` | pod 번호 |
| immediacy | `immediate` | `immediate`(바로 leaf에 배포) / `lazy`(On Demand, 단말이 학습될 때 배포). 행이 없으면 immediate |

- 비밀번호는 시트에 적지 않습니다. 실행할 때 묻거나 `APIC_PASSWORD` 환경변수를 씁니다.
- 열은 머리글 이름으로 찾으므로 순서가 바뀌어도 되고, 머리글이 없으면 양식 순서(A~F)대로 읽습니다.

### Deploy_int_excute 작성법

| AP | Endpoint Group | Target Node | Port | Vlan | Mode |

- **AP / Endpoint Group은 비워 두는 것을 권장**합니다. Vlan으로 테넌트 시트에서 찾습니다.
  값을 적으면 Vlan으로 찾은 EPG와 같은지 확인하고, 다르면 오류로 멈춥니다.
- **Port**: 일반 포트는 `eth1/1`, PC/vPC는 Interface Policy Group 이름 (예: `vpc_leaf101_102`).
- **Target Node**: 일반 포트는 노드 하나. vPC는 `101-102`처럼 두 노드 (하나만 적어도 적용 시 APIC에서 vPC 경로를 찾음).
- **Mode**
  - `tagged` → Trunk (regular)
  - `native` → Access 802.1P. **같은 포트에 tagged와 함께 untagged VLAN을 쓸 때는 이걸 사용**
  - `untagged` → Access untagged. 그 포트에 VLAN 하나만 쓸 때만 사용
  - 한 포트에 native/untagged는 하나만 가능

## 적용할 때 자동으로 하는 확인 (`--apply`)

아래 중 하나라도 걸리면 **아무것도 넣지 않고** 문제 목록만 보여 줍니다.

1. 엑셀 검증: Vlan 숫자, 포트 형식, Mode 값, Vlan과 AP/EPG 일치, 같은 포트·VLAN 중복
2. 시트의 EPG가 APIC에 모두 있는지 (없는 EPG에 넣으면 BD 없는 빈 EPG가 생기므로)
3. PC/vPC 이름으로 실제 포트 경로(`protpaths-...`)를 APIC에서 찾을 수 있는지
4. 같은 포트에 같은 VLAN이 **다른 EPG**에 이미 붙어 있지 않은지 (공인/사설 VLAN 충돌 등)

확인을 통과하면 EPG마다 한 번씩 POST합니다. POST는 추가/변경만 하고 기존 static port를 지우지 않습니다.
같은 시트를 다시 적용하면 이미 있는 포트는 Mode/immediacy만 엑셀 값으로 갱신됩니다.

## 자주 나오는 오류

| 메시지 | 원인 / 조치 |
|---|---|
| `연결 실패 ... timed out` | PC에서 APIC로 연결 안 됨. SecureCRT 터널이 켜져 있는지, `proxy` 값이 맞는지 확인 |
| `HTTP 401 ... Authentication DENIED` | 계정/비밀번호 또는 로그인 도메인 오류. `apic_username` 형식 확인 |
| `머리글 ... 열이 없습니다` | 시트 머리글 이름이 다름. `AP, Endpoint Group, Target Node, Port, Vlan, Mode` |
| `AP/EPG ... 가 Vlan N 의 ... 와 다릅니다` | 그 행의 Vlan과 AP/EPG가 서로 안 맞음. AP/EPG 칸을 비우면 Vlan 기준으로 찾음 |
| `Vlan N 이 ... 시트에 없습니다` | 테넌트 시트에 그 VLAN의 EPG가 없음 |
| `이미 ... 에 붙어 있음` | 같은 포트·VLAN을 다른 EPG가 사용 중. VLAN이나 포트를 바꿔야 함 |
| `[100] Validation failed ... encapsulation` | 위와 같은 충돌을 APIC가 거부한 것 |
| Deployed EPG에 안 보임 | `immediacy`가 lazy(On Demand)면 단말이 붙어야 배포됨. 그래도 안 보이면 EPG Faults에서 `invalid-vlan`/`invalid-path` 확인 (Domain·VLAN Pool·AAEP 설정) |

## 폴더 구조

```
run.py                 실행기 (로그인 → 메뉴 → 검증/dry-run/적용)
aci_auto/client.py     APIC 로그인, 토큰 자동 갱신, 프록시, dry-run 출력
aci_auto/excel.py      엑셀 읽기 (Settings / 테넌트 시트 / static port 시트)
aci_auto/config.py     입력 검증
aci_auto/builders.py   APIC JSON(MO) 생성
requirements.txt       필요한 모듈 목록
lib/                   (선택) 폐쇄망용 모듈 묶음, git 에는 올리지 않음
inventory/             엑셀 파일 위치 (git 에는 올리지 않음)
.gitignore             실제 운영 엑셀(*.xlsx) 제외
```

## 참고

- 엑셀에는 APIC 주소, 계정, 테넌트, 서비스 대역이 들어 있으므로 git에 올리지 않습니다 (`.gitignore`).
- 신규 구축용으로 테넌트 시트 전체(VRF/BD/AP/EPG/Contract)를 만드는 `--build-all` 옵션과
  YAML 입력(`inventory/example.yaml`)도 남아 있지만, 평소 작업에는 쓰지 않습니다.
