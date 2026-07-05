# export-tracker

관세청 TRASS(무역통계) API로 국내 상장 섹터의 **수출 트렌드를 자동 수집·집계**하는 도구 모음.
넓게 훑는 **섹터 breadth 모니터**와, 반도체를 깊게 보는 **심층 트래커** 두 갈래로 구성된다.

## 구성

| 폴더 | 내용 |
|------|------|
| [`export_data_all/`](export_data_all/) | 15개 섹터 수출 트래커 (엔진 · 섹터 정의 · 캐시 · 엑셀 출력) |
| [`export_data_all/semi/`](export_data_all/semi/) | 반도체 심층 트래커 (FRED·TrendForce·SEMI 등 교차검증 포함) |

각 폴더에 상세 README가 있다 → [섹터 트래커 README](export_data_all/README.md) · [반도체 트래커 README](export_data_all/semi/README.md)

## 대상 섹터 (15개)
**기존 8:** K뷰티화장품 · 이차전지 · 자동차 · 철강 · 석유화학 · 조선 · 의료기기 · K푸드
**확장 7:** 방산 · 전력기기·변압기 · 제약·바이오 · 엔터·K콘텐츠 · 태양광·신재생 · 기계·건설기계 · 타이어

## 빠른 시작

```bash
# 1) 의존성 설치
pip install pandas numpy openpyxl requests

# 2) API 없이 캐시만으로 마스터 엑셀 생성 (즉시)
cd export_data_all
python sector_tracker.py --mode cache
```

- 결과물: `export_data_all/output/sector_tracker.xlsx`
- 저장소에 캐시 CSV가 포함돼 있어 **API 키 없이 `cache` 모드로 바로 실행**된다.
- 신규 데이터를 직접 수집하려면(`full`/`update` 모드) 관세청 공공데이터포털 서비스키가 필요하다.

## 실행 모드

| MODE | 동작 | API 호출 |
|------|------|----------|
| `full`   | 최초 전체 수집 → 캐시 저장 → 엑셀 | 많음 (~30~45분) |
| `update` | 매월 최근 2개월만 수집해 캐시 누적 → 엑셀 | 적음 (~1분) |
| `cache`  | **API 없이** 캐시만 읽어 마스터 엑셀 생성 | 없음 (즉시) |

## API 키 설정
`full`/`update` 모드는 관세청 [공공데이터포털](https://www.data.go.kr) 서비스키가 필요하다.
`export_data_all/sector_tracker.py`(및 `semi/semicon_tracker.py`) 상단의 `SERVICE_KEY` 상수에 발급키를 입력한다.

> ⚠️ 공개 저장소에 실제 키를 커밋하지 말 것. 개인 키를 넣은 뒤에는 push하지 않거나, 환경변수로 분리해 사용하는 것을 권장한다.
