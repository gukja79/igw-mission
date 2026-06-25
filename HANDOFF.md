# igw-mission 핸드오프

선교 오프라인 회계 PWA. 오프라인 우선(IndexedDB) + Google Sheets 동기화.
스택: GitHub Pages(정적) + Apps Script(독립형 웹앱) + vanilla JS. Claude Code 터미널로 파일·git, Apps Script는 브라우저 편집기에서 붙여넣고 재배포.

- 라이브: https://gukja79.github.io/igw-mission/
- repo: github.com/gukja79/igw-mission  (main)
- Apps Script exec URL: …/macros/s/AKfycbygbwjiL_zkHzgKwXlasJ4a1WpHZ8MlVGk6raBJtGwSc5UMUExC53FTL1hbpY74TZBt/exec
- SECRET: 29540027  (앱 SYNC.secret 과 동일)
- 운영 범위: **현재 몽골 팀만 운영.** 나머지 6팀(라오스1/중국1/인도네시아1/말레이시아1/캄보디아1/스리랑카)은 SHEETS 에 ID 미설정 상태(`여기에-…`).

---

## 작업 흐름 (변함없음)
- 설계·검토 = Claude 채팅 / 파일수정·git = Claude Code 터미널 / Apps Script = 브라우저 편집기 붙여넣기 후 재배포
- Apps Script 재배포: 편집기 저장 → **배포 관리 → 편집(연필) → 버전: 새 버전 → 배포** (URL 동일). "새 배포"는 URL 이 바뀌므로 금지.
- 프런트 변경 시 sw.js CACHE 버전 bump. 라이브 확인은 `?v=N` 쿼리로 캐시 우회. 서비스워커는 "한 번 방문해 새 sw 받고 다음 방문에 적용" 구조라 폰은 **앱 두 번 닫았다 열기**가 필요.
- 커밋: 한 토픽 = 한 커밋, 한글 메시지. diff 건별 확인 후 커밋.

---

## 시트 구조 (대상 탭 '수입지출내역서', 데이터 12행부터)
A 일자 · B 계정과목 · C 지원금/자부담 · D 내역 · E 수입원화(수식) · F 지출원화(수식) ·
G 사용금액 · H 사용통화 · I 영수증(O/X) · J 비고 · K 사진링크 ·
**L id(숨김)** · **M 영수증번호(숨김)**

- 정렬 보기 탭 `수입지출내역서01` 은 A:K 만 FILTER → L·M 은 안 딸려옴(숨김 OK).
- **L = entry id**: 수정·삭제·재시도·머지의 행 식별 단일 키.
- **M = 영수증번호**: 서버 단일 발급 `max(M)+1`. **숫자 서식 강제 필수**(날짜 서식 셀이면 3을 날짜로 오인 → epoch 오염). 코드가 setNumberFormat("0") 로 방어.

---

## ✅ Phase 2 완료 (이번 세션) — 앱↔시트 양방향 동기화

### 서버 (Code.gs)
1. **pull** (`action:"pull"`): 시트 전체를 앱 거울용 JSON 으로 반환. + **보정**: id(L) 비고 A·B·C·D·G·H 가 다 찬 '완성행'은 이때 L=id, M=영수증번호 자동 부여(위→아래). 누가 웹 입력했든 pull 한 번이면 빠짐없이 번호 박힘. 멱등(이미 채워진 행은 재부여 안 함).
2. **영수증번호 서버 단일 발급**: create 가 `nextReceiptNo_`(=max(M)+1) 로 발급. **LockService** 로 직렬화(동시 입력 번호충돌 방지). 사진 파일명도 서버 발급번호 사용. 앱이 보낸 receiptNo 는 무시.
3. **pull 보정(backfill)**: 위 1번에 포함. 트리거 대신 pull 시점 보정 방식 선택 — 이유: 회계 계정이 여럿(mong@… 등)이라 설치형 onEdit 트리거를 각 계정에 깔 수 없음. pull 보정은 누가 입력했든 누락 0, 설치 부담 0.
4. **hotfix**: M열 날짜서식 오염 방어. write 양쪽 `setNumberFormat("0")` + read 는 `receiptNoOf_()` 로 Date 오염값 → null.

검증 완료: 몽골 pull → 12·13·14 backfill(번호 2·3·4), 15(앱입력)=1, 멱등 backfilled:0, M13 정수 정상.

### 앱 (index.html)
- **머지 코어**: `pullNow()` (pull→머지→렌더, 시트로 push 안 함) + `mergeFromSheet_()`. 변환 헬퍼 `sheetRowToEntry_`/`applySheetValues_`/`sheetDiffers_`. type 은 krwIn>0 이면 수입.
- **머지 규칙 (확정)**:
  - 시트만 있음 → 신규 추가(syncState:"완료", syncedRow, sheetPhoto)
  - 양쪽 + 로컬 "완료" → **시트값으로 갱신**(A). 단 로컬 사진(base64)·생성시각 보존.
  - 양쪽 + 미동기화(대기/수정대기/삭제대기) → **로컬 유지**. 시트와 값 다르면 `mirrorFlag:"시트와다름"` 표시(자동 안 덮음).
  - 로컬만 + "완료" → **보존 + `mirrorFlag:"시트에없음"`**(B: 시트에서 삭제된 행). 자동삭제 안 함(데이터 손실 방지).
  - 로컬만 + 미동기화 → 보존(곧 push), 표시 없음.
- **UI**: "내려받기" 버튼(동기화 옆) / mirror 배지(노랑) / 원화 강조(외화행은 원화=메인 크기·굵게, 외화=같은 크기·굵기없이 흐리게) / **⋯ 액션메뉴**(수정·삭제, 카드 본체 탭은 무반응) / **사진 라이트박스**(앱에서 올린 b64 썸네일 탭→전체화면, 탭해서 닫기) / 비고는 배지 아닌 **하단 📝 메모 줄** / 깨진 b64 onerror 가드는 repo 에 미반영(필요시 추가).
- **사진 정책**: 웹은 사진 못 찍음 → 사진은 앱 입력 행에만 존재. 앱 직접올림=썸네일+라이트박스 / 내려받은 시트행(드묾)=`📎 사진 보기` 링크(private 라 권한자만), 사진 없으면 "사진 없음" 표시.

### 배포
- sw.js CACHE **v7 → v10**. 최종 커밋 `01ad1bf`.
- 주요 커밋: bc2d935(pull) · f26a0ea(번호발급+락) · ec39f68(backfill) · (hotfix) · d72d2d9(머지코어) · e95a396(버튼·배지·사진링크) · f86879f(⋯·라이트박스) · 01ad1bf(비고→메모)

---

## 🔧 남은 잔손질 (다음에, 전부 사소)
- [ ] Code.gs doPost 라우팅 주석: `// pull (읽기 전용)` → 실제론 보정도 함, 문구만 갱신 (동작 무관)
- [ ] 나눔고딕 폰트 적용 — 단, 오프라인 우선 앱이라 **웹폰트 sw 캐싱과 함께** 설계해야 함(안 하면 오프라인서 시스템폰트 폴백). 별도 작업으로.
- [ ] `<meta apple-mobile-web-app-capable>` deprecated 경고 → `<meta name="mobile-web-app-capable">` 추가
- [ ] 깨진 b64 onerror 폴백 가드(현재 repo 미반영). 아이폰 IndexedDB b64 손상 재발 시 추가 검토.

## ⏳ 7팀 확장 시 (지금은 몽골만)
- [ ] 각 팀 시트에 **L·M 열 추가**(숨김) + M열 숫자 서식
- [ ] Code.gs SHEETS 에 팀별 스프레드시트 ID 입력 (현재 `여기에-…`)
- [ ] 각 팀 시트 조건 확인: I3/I6/I9 환율, E12/F12 환산 수식, K11 '사진' 머리글
- [ ] 앱 SYNC 에 팀별 url·secret (전 팀 동일 값)

## 알아둘 점
- 새 버전 배포해도 폰은 "두 번 열어야 최신". 회계들에게도 동일. (원하면 '새 버전 안내 후 새로고침' 기능 추후 추가 가능)
- 사진 원본은 드라이브에 안전. 폰 IndexedDB 의 b64 가 깨져도 데이터 손실 아님(표시만 문제).
- 시트에서 행 삭제 → 다음 pull 때 그 폰 앱엔 `시트에없음` 표시로 남음(보존). 폰에서도 지우려면 카드 ⋯→삭제.
