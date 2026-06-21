# igw-mission — 선교 오프라인 회계 (핸드오버)

## 개요
선교여행용 **오프라인 회계 입력 PWA**. 각 팀 회계가 본인 폰에서 오프라인으로
일자·계정과목·지원금/자부담·내역·사용통화·사용금액·영수증을 입력 → 인터넷 연결 시
각 팀 구글시트(`수입지출내역서`)에 자동으로 행이 추가됨. 원화 환산은 **시트 수식**이 계산.
"PC에서 구글시트에 직접 입력하던 것"을 오프라인 모바일로 대체하는 도구.

## 스택
- 프론트: vanilla HTML/CSS/JS + IndexedDB(로컬 저장) + Service Worker(오프라인 셸)
- 백엔드: 팀별 구글시트 + Apps Script 웹앱(doPost)
- 호스팅: GitHub Pages (gukja79.github.io)

## 파일
- `index.html`    — 앱 본체 (입력 / 내역 / 환율 화면, 동기화 로직)
- `sw.js`         — 서비스워커 (오프라인 셸 캐시). **파일 수정 시 맨 위 CACHE 버전 올리기**
- `manifest.json` — PWA 설치 정보
- `icon-192.png`, `icon-512.png` — 앱 아이콘
- `Code.gs`       — Apps Script (각 팀 시트에 **수동으로 붙여 배포** — Pages 대상 아님, 참고용 보관)

## 현재 상태
- [완료] 1단계 — 오프라인 입력 + 폰 저장(사진 포함) + 영수증 O/X + 내역 목록
- [완료] 2단계 — 연결 시 구글시트 자동 반영 (한 건씩, 끊기면 이어서, id 중복 방지)
- [완료] 환율 설정 + 입력 시 원화 환산(≈) 실시간 표시
- [남음] 사진 → 구글 드라이브 업로드 + 비고에 링크 (행 반영 검증 후 진행)
- [최후순위] 환전 거래내역 시트로 기준환율 자동 산출 (시트 쪽 작업)

## 코드 안 설정값
- `TEAMS`: 라오스1 · 몽골 · 중국1 · 인도네시아1 · 말레이시아1 · 캄보디아1 · 스리랑카 (7팀, 설정됨)
- `SYNC[팀]`: 각 팀 시트 Apps Script 배포 후 `url`·`secret` 입력 (현재 비어 있음)
- `CURRENCIES`: 원화 / 달러 / 현지화1 / 현지화2 (시트 H열 값과 동일, label만 통화명으로 변경 가능)

## 팀 시트 연결 (팀마다 1회)
1. 시트 `E12`/`F12`에 환산 수식 입력 (아래 참조)
2. `I3`=달러환율, `I6`=현지화1환율, `I9`=현지화2환율 확인
3. 확장 프로그램 → Apps Script에 `Code.gs` 붙이고 `SECRET`/`SHEET_NAME`/`FIRST_DATA_ROW` 조정
4. 배포 → 새 배포 → 웹 앱(실행: 나 / 액세스: 모든 사용자) → URL 복사
5. `index.html`의 `SYNC[팀]`에 `url`·`secret` 입력 (secret은 `Code.gs`와 동일하게)

### 시트 수식 (수입/지출 자동 분기)
```
E12  =IF(LEFT($B12,2)="수입",IFERROR(ROUND(IF($H12="원화",$G12,IF($H12="달러",$I$3*$G12,IF($H12="현지화1",$G12*$I$6,IF($H12="현지화2",$G12*$I$9,"")))),0),""),"")
F12  =IF(LEFT($B12,2)="수입","",IFERROR(ROUND(IF($H12="원화",$G12,IF($H12="달러",$I$3*$G12,IF($H12="현지화1",$G12*$I$6,IF($H12="현지화2",$G12*$I$9,"")))),0),""))
```

## 배포 (GitHub Pages)
레포가 아직 없으면 (gh CLI 사용):
```
gh repo create gukja79/igw-mission --public --source=. --remote=origin --push
```
또는 수동:
```
git init
git add .
git commit -m "선교 오프라인 회계: 1·2단계 + 환율 표시"
git branch -M main
git remote add origin https://github.com/gukja79/igw-mission.git
git push -u origin main
```
→ GitHub Settings → Pages → Branch: `main` / 루트 → `https://gukja79.github.io/igw-mission/`

## 테스트 체크리스트
- 팀 선택 → "환율"에서 환율 입력 → 지출/수입 입력 시 원화 환산(≈) 표시되는지
- 내역 탭 → "동기화" → 시트에 행 추가 + E/F열 원화 자동 계산되는지
- 비행기 모드로 껐다 켜도 앱이 열리고 데이터가 남는지 / 연결되면 자동 동기화되는지
- 사진 2~3장 첨부 → 영수증 자동 O / × 로 제거 가능한지
