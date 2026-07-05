# 반도체 수출·가격·장비 사이클 트래커 (semicon_tracker)

한국 반도체 업황을 **수출(물량·금액) → 가격(ASP·Spot·Contract) → 장비(capex)** 세 축으로
한 워크북에 모으는 심층 모델. 옆 `export_data_all`(8섹터 breadth 모니터)과 달리, 여기는
반도체 한 섹터를 가장 깊게 파는 **depth 모델**이다.

투자 관점의 한 줄: **관세청 수출 데이터는 메모리 업체(삼성전자·SK하이닉스) 실적·업황의
가장 빠른 공개 선행지표**다. 분기 실적이 나오기 전에 월별(심지어 10·20일 잠정치로 당월
중간에) 방향을 읽는 것이 목적.

---

## 1. 데이터 소스 — 2계층 구조

### 계층 A: 관세청 TRASS (자동·무료 API, `getNitemtradeList`)
스크립트가 키 하나로 자동 수집. `data/cache_memory.csv`, `data/cache_ict.csv` 에 누적.

**(A-1) 메모리 상세 — HS 854232 (`cache_memory.csv`, 2019-01~)**
국가 × HS 10자리 단위 원천 데이터. 핵심 필드:

| 필드 | 의미 | 분석에서의 쓰임 |
|---|---|---|
| `exp_usd` | 수출액 (USD) | 물량×가격이 섞인 금액. 업황 모멘텀의 기본 |
| `exp_wgt` | 수출 중량 (kg) | **물량 프록시**. 금액과 분리해 가격/물량 기여 분해 |
| `imp_usd`/`imp_wgt` | 수입액·중량 | 재수입(후공정 위탁), 역수입 흐름 참고 |
| `asp_usd_kg` | = exp_usd / exp_wgt | **단가(ASP) 프록시**. 아래 §3 참조 |
| `country_code` | 향지(ISO-2) | 수요처·환적 경로 (HK/TW/VN 의미는 §3) |

**HS 10자리 제품 분류 (`PRODUCT_MAP`) — 메모리 안에서 무엇을 보는가:**

| HS10 | 분류 | 의미 |
|---|---|---|
| `8542321010` | DRAM | 범용/서버 D램 |
| `8542321030` | NAND Flash | 낸드 |
| `8542321020` | SRAM | |
| `8542321090` | Other Memory | 기타 메모리 |
| `8542322000` | Hybrid IC | 하이브리드 집적회로 |
| **`8542323000`** | **HBM/첨단패키징 프록시** | "복합구조칩 집적회로" = 적층·복합구조. **HBM·2.5D/3D 첨단패키징의 대용**. ⚠️ 순수 HBM만은 아님(다이 적층 제품 일반 포함) → 방향성 지표로만 |
| `8542324000` | MCOs | 복합부품 집적회로 |

**(A-2) MOTIE ICT 수출 재현 (`cache_ict.csv`)**
산업부(MOTIE)가 매월 발표하는 ICT 수출 동향을 관세청 원천으로 직접 재현. `ICT_HS` 묶음:

| 품목 | HS 묶음 | 의미 |
|---|---|---|
| 반도체 | `8541` + `8542` | 8541=개별소자(다이오드·트랜지스터·포토셀 등) + 8542=집적회로(IC). 합 = 반도체 전체 |
| 디스플레이 | `8524` | HS2022 신설, 평판디스플레이 모듈(OLED/LCD) |
| 휴대폰 | `851713` | 스마트폰 포함 셀룰러폰 |

→ 공식 발표 전에 같은 숫자를 미리 계산. `exp_musd`(백만달러), MoM/YoY 자동 산출.

### 계층 B: 수동 입력 CSV (유료·보도자료 — 무료 API 없음)
`data/`에 해당 CSV가 없으면 스크립트가 **예시 1행짜리 템플릿을 자동 생성**한다. 그
아래에 실제 행을 계속 붙이면 누적·중복제거·최신순 정렬된다(예시행은 로드 시 자동 무시).
**이제 이 3종(Spot/Contract/Billings)은 아래 계층 C 자동 프록시로 대체되므로 선택사항**이다.
(TrendForce 유료 구독이 있으면 직접 붙여 넣어 교차검증용으로 쓰면 된다.)

| 파일 | 내용 | 핵심 컬럼 | 의미 |
|---|---|---|---|
| `trendforce_spot.csv` | DRAM/NAND **현물가** | `date, product, price_usd, chg_pct` | 소량·단기 거래가. 변동성 크고 **사이클 전환 선행** |
| `trendforce_contract.csv` | DRAM/NAND **계약가** | `yyyymm, product, price_usd, mom_pct` | 대형 고객 월간 협상가. **실제 매출 단가에 직결** |
| `semi_billings.csv` | SEMI 장비 **빌링/북투빌** | `yyyymm, region, billings_musd, bookings_musd, book_to_bill` | 장비 투자(capex) 사이클. **공급 선행** |
| `trass_speed.csv` | 관세청 **10/20일 잠정치** | `date_label, item, exp_musd, yoy_pct, note` | 당월 마감 전 **속보**(§3) |

> `motie_ict.csv` 는 **스크립트가 읽지 않는 참고용 고아 파일**이다(공식 MOTIE 발표치 1행).
> A-2의 관세청 재현치와 공식 발표를 손으로 대조할 때 쓰는 용도로 보인다. 자동 반영을
> 원하면 `CSV_SOURCES`에 등록하면 된다.

### 계층 C: 자동 대체 소스 (수동 CSV를 무료·자동으로 대체)
유료 TrendForce/SEMI 대신, 같은 신호를 잡는 무료 소스를 코드가 자동 수집한다.

| 소스 | 대체 대상 | 출처 | 의미 |
|---|---|---|---|
| **반도체장비 수입(HS 8486)** | SEMI Billings | 관세청(자동) | 한국이 ASML·TEL·AMAT 장비를 수입 = 국내 팹 **capex 프록시**. `EQUIP_HS` |
| **메모리 ASP(USD/kg) + 지수** | Spot/Contract | 관세청(자동) | DRAM/NAND/HBM 수출단가 = 메모리 **가격 프록시**. 최초=100 지수화 |
| **미국 반도체 PPI** | 외부 가격 검증 | FRED(자동) | `PCU334413334413`, API키 불필요 CSV. `FRED_SERIES` |

- 장비수입은 **향국 분해**까지 자동 → 네덜란드(ASML)·일본(TEL)·미국(AMAT/LAM) 공급망 추적.
- ⚠️ capex 프록시는 '인도(billings)' 측면만 — bookings가 없어 **book-to-bill은 산출 불가**.
  대신 수입액 YoY·3개월 이동평균으로 capex 사이클을 본다.
- ⚠️ 8486은 디스플레이 장비(848630)를 제외한 반도체 코어(848620)+부품(848690) 기준.

---

## 2. 출력 워크북 시트 가이드 (`semicon_tracker.xlsx`)

**Dashboard** (최신월 스냅샷 6블록, 당월 미집계는 제외):
- `[1] 메모리 수출 국가별` — 향지별 수출액·비중·ASP·MoM·YoY (KEY_COUNTRIES 우선)
- `[2] 제품 믹스` — DRAM/NAND/HBM프록시 등 제품별 최신월
- `[3] HBM/첨단패키징 프록시 향지별` — 8542323000만 떼어 향지별
- `[4] ICT 수출(관세청 재현)` — 반도체/디스플레이/휴대폰 백만달러·YoY
- `[5] 반도체장비 수입 = capex 프록시` — 장비 수입액·3M평균·YoY (계층 C)
- `[6] 외부 가격지표 (FRED)` — 미국 반도체 PPI 등 최신값·YoY

**상세 시계열 시트:**
- `국가별_월간` / `제품별_월간` / `HBM프록시_향지별` — 전체 월 시계열
- `MOTIE_ICT_월간` — ICT 재현 시계열
- `DRAM_NAND_HBM_ASP_관세청` — DRAM·NAND·**HBM** ASP(USD/kg)+지수 = **Spot/Contract 대체 가격 프록시**
- `장비수입_월간_capex` — 반도체장비(8486) 수입 = **SEMI Billings 대체 capex 프록시**
- `장비수입_향국` — 장비 수입 공급망(ASML·TEL·AMAT 향국별)
- `FRED_미국_반도체_PPI` — 외부 가격지수 (무료 자동)
- `원본_10자리상세` — 가공 전 원천(감사·재집계용)
- `DRAM_NAND_Spot` / `_Contract` / `SEMI_Billings` — (선택) 수동 입력 시 표시
- `관세청_10_20일_잠정` / `DRAM_NAND_Spot` / `DRAM_NAND_Contract` / `SEMI_Billings` — 수동 CSV 반영

조건부 서식: MoM/YoY 열은 -30%(빨강)~0(흰)~+30%(초록) 컬러스케일.

---

## 3. 핵심 지표 해석 (도메인 노트)

**ASP(USD/kg) 프록시 — 왜 가격 사이클을 잡나**
메모리는 금액(`exp_usd`)과 중량(`exp_wgt`)이 같이 신고된다. 그 비율 USD/kg이 사실상
수출 단가다. 방향이 메모리 가격 사이클(상승/하락)을 따라간다. HBM은 가볍고 비싸
USD/kg이 압도적으로 높다(예: HBM프록시 향지 HK 한 달 ≈ $49,000/kg vs 범용 D램 수천/kg).
⚠️ **한계**: ① 믹스 변화(HBM 비중↑)만으로 블렌디드 ASP가 올라 칩값이 안 변해도 상승해
보임 ② 중량에 패키지·기판 포함 → 순수 실리콘 가격 아님 ③ 중량 분모가 작으면 노이즈 큼.
→ **절대수준보다 추세·전환점**으로 읽을 것.

**Spot vs Contract — 리드/래그**
현물(Spot)은 소량·단기라 민감하게 먼저 움직이고, 계약가(Contract)는 대형 고객 월간
협상가라 실적 매출단가에 직결된다. 보통 **Spot이 Contract를 수 주~수개월 선행** → Spot
반등이 Contract 반등(=실적 개선)의 선행 신호.

**Book-to-Bill — capex/공급 선행**
`book_to_bill = bookings / billings`. >1이면 수주가 매출을 초과 = 장비 투자 확장 국면.
장비 투자는 6~12개월 뒤 캐파(공급)로 이어지므로 **공급 사이클 선행지표**. North America
기준치는 미국 장비사(AMAT·LAM·KLA) 비중이 커 글로벌 capex의 대표 프록시.

**향지(국가)의 진짜 의미**
- `HK 홍콩` — 중국 본토 재수출 **환적 허브**. 실수요는 중국 → **HK+CN 합산**으로 봐야 정확
- `TW 대만` — TSMC/OSAT 패키징·중화권 세트 수요
- `VN 베트남` — 삼성 후공정·세트 조립 거점
- `SG 싱가포르` — 환적·유통 허브
- `US 미국` — 데이터센터·AI 직수요(HBM 관련 모멘텀)

**관세청 10/20일 잠정치 — 당월 속보**
관세청이 매월 11일·21일경 1~10일·1~20일 누계 수출 잠정치를 발표한다. 월 마감 전 반도체
수출 YoY로 **업황·분기 실적 방향을 가장 빠르게** 포착하는 용도(`trass_speed.csv`).

---

## 4. 실행 / 운영

VS Code에서 파일 상단 상수만 고치고 Run. `SERVICE_KEY`는 8섹터 트래커와 동일 키 사용 가능.

| MODE | 동작 | 용도 |
|---|---|---|
| `"full"` | `MONTHS`개월(2019-01~ ≈ 90) 전체 호출 → 캐시 **덮어쓰기** | 최초 1회 |
| `"update"` | 최근 2개월만 호출 → 캐시에 **누적** | 매월 |

**월간 루틴**
1. `MODE="update"` Run → 신규 월이 `cache_memory.csv`/`cache_ict.csv`에 누적, 엑셀 재생성.
2. (선택) `trendforce_*` / `semi_billings` / `trass_speed` CSV에 최신 행 추가.
3. 엑셀 Dashboard에서 메모리 수출 YoY·ASP·HBM 향지·book-to-bill 점검.

CLI: `python semicon_tracker.py --mode update`  (키 없이 검증: `--demo`)

---

## 5. 폴더 구조
```
semi/
├─ semicon_tracker.py    # 엔진(수집·가공·엑셀)
├─ semicon_tracker.xlsx  # 출력
├─ README.md             # 이 문서
└─ data/
    ├─ cache_memory.csv      # [자동] 메모리 854232 상세 (원천 캐시)
    ├─ cache_ict.csv         # [자동] ICT 재현 캐시
    ├─ trendforce_spot.csv   # [수동] 현물가
    ├─ trendforce_contract.csv  # [수동] 계약가
    ├─ semi_billings.csv     # [수동] 장비 빌링/북투빌
    ├─ trass_speed.csv       # [수동] 10/20일 잠정치
    └─ motie_ict.csv         # [참고] 공식 MOTIE 발표치 (스크립트 미사용)
```

---

## 6. 데이터 주의점 (gotchas)
- **당월은 미집계**: 최신월(예: 진행 중인 달)은 수출이 0에 가까워 YoY가 -100%처럼 찍힌다.
  대시보드 스냅샷이 당월을 가리키면 직전 완성월로 해석할 것.
  *(8섹터 트래커는 대시보드 기준월에서 당월을 자동 제외하도록 보정됨 — 이 파일도 동일 보정을
  원하면 적용 가능.)*
- **HBM 프록시(8542323000)는 순수 HBM이 아님** — 적층·복합구조 칩 일반 포함. 수준이 아니라
  추세·향지 변화로 읽기.
- **ASP는 중량 기반** → 믹스·패키지 무게에 왜곡. §3 한계 참조.
- **HK는 환적** → 중국 실수요와 합산 해석.
- 캐시(`data/cache_*.csv`)는 원천이다. 임의 편집 금지. 재생성하려면 `full` 모드.
