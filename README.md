# 판매/재고/수급 현황 PWA (stock-supply)

품명별 **일별 판매 · 일별 재고 · 수급(생산/수입) 일정**을 한 화면에서 보는 직원 공유용 웹앱.
갱신은 `update.bat` 더블클릭 한 번 (= 바탕화면 **재고수급 업데이트**).

```
판매  01)실적dashboard\RAW_상품별유형별 실적_출력.xlsx     ← 파워쿼리로 갱신 (사용자)
재고  ★Jay\13.생산&재고\AA.재고_쿼리\창고별 재고조회_*.xlsx  ← 일자별 저장 (사용자)
생산  14)Notion\01)제품생산계획 스페이스\일자별_자재계획_*.xlsx ← 최신 파일 자동 선택
수입  구글시트 '수입상품 입고일정' 종합 탭                 ← 자동 다운로드 (로그인 불필요)
코드집 ★Jay\06.상품\▥상품코드집_운영.xlsx                  ← 정렬·유형
      │
      ▼  update.bat → build_data.py (검산) → data/data.json → git push
GitHub Pages  https://yim0790.github.io/stock-supply/
```

## 최초 1회 (내 PC)

1. Python·Git 설치 확인 (`python --version`, `git --version`). 생산계획_PWA를 쓰고 있으면 이미 있음
2. https://github.com/new → Repository name `stock-supply` · **Private** · README 체크 해제 → Create
   (비공개 저장소 Pages는 GitHub Pro 필요 — production-plan과 같은 계정이면 이미 충족)
3. `config.cmd` 확인 (GH_USER / GH_REPO=stock-supply / GH_MAIL) — 이미 채워져 있음
4. `setup_github.bat` 더블클릭 (로그인 창 뜨면 로그인 후 한 번 더)
5. 저장소 → Settings → Pages → Source: Deploy from a branch → Branch **main / (root)** → Save
6. `make_shortcut.bat` 더블클릭 → 바탕화면에 **재고수급 업데이트**, `직원배포용\판매재고수급 현황.url` 생성

## 매번 (약 1~2분)

1. 판매 RAW 파워쿼리 새로고침 → 저장
2. 오늘 재고 파일을 AA.재고_쿼리 폴더에 저장
3. 바탕화면 **재고수급 업데이트** 더블클릭 → 검산표 확인 → `DONE`
4. 1분 뒤 직원 화면 새로고침

## 화면 규칙

- 표시 품목 = 상품코드집(제품명 B열) ∪ 판매실적 품명 중 **최근 12개월 판매·현재고·수급 하나라도 있는 것**. 코드집 행 순서로 정렬
- 회전일 = 현재고 ÷ 최근 90일 일평균 판매. 15일 미만 빨강, 현재고 0 = **품절**
- 재고 탭: 전일 대비 **+30 이상 → 숫자 빨간 볼드**(입고), 0 → 회색
- 입고 열 **[일정]** = 오늘 이후 생산계획(계획수량) 또는 수입일정 있음 → 클릭 시 팝업
- 수입 일자 = 김포공장 입고일 우선, 없으면 ETA `(ETA)` 표기. 오늘 −31일부터 표시(지난 건 회색 "지남")
- 판매 셀 클릭 → 분석채널·거래처·수량

## 스크립트가 멈추는 경우 (update_log.txt 확인)

- 판매 RAW / 코드집 파일 없음 · 헤더 이름 바뀜
- 검산 불일치 (판매 합계 · 재고 TOTAL · 생산 합계) — 원본을 확인하기 전엔 배포하지 않음
- 구글시트 다운로드 실패 → 멈추지 않고 `cache\import_latest.xlsx`(마지막 성공본)로 진행, 로그에 [주의]

## 폴더

```
06_Stock-Supply/
├ build_data.py · index.html · data/data.json · sw.js · manifest.json · icons/
├ config.cmd · update.bat · setup_github.bat · make_shortcut.bat · make_shortcut.py
├ cache/   재고 파일 파싱 캐시 + 수입 시트 마지막 성공본 (git 제외)
├ docs/    계획서 · 체크리스트 · 컨텍스트노트 (git 제외)
└ 목업/    승인용 PC 목업 (git 제외)
```
