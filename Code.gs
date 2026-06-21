/* =====================================================================
   선교 오프라인 회계 — 시트 수신용 Apps Script (팀 시트마다 1개)

   설치 방법
   1) 팀 구글시트 열기 → 확장 프로그램 → Apps Script
   2) 이 코드 전체 붙여넣기
   3) 아래 SECRET 을 팀별 비밀키로 바꾸기 (앱의 SYNC[팀].secret 과 동일하게)
   4) 배포 → 새 배포 → 유형: 웹 앱
        - 실행 계정: 나
        - 액세스 권한: 모든 사용자
   5) 나온 웹앱 URL 을 앱(index.html) SYNC[팀].url 에 입력

   ※ 시트 조건: 대상 탭에 아래 셀이 있어야 함
     - I3 = 달러 기준환율 / I6 = 현지화1 환율 / I9 = 현지화2 환율
     - E12, F12 에 환산 수식(수입/지출 자동 분기) — 핸드오버 메모의 수식 붙여넣기
   ===================================================================== */

const SECRET = "CHANGE-ME-팀별-비밀키";   // ← 앱과 동일하게 변경
const SHEET_NAME = "수입지출내역서";        // ← 대상 탭 이름 (팀 시트마다 확인)
const FIRST_DATA_ROW = 12;                  // ← 데이터 시작 행 (몽골팀 기준 12)

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    if (body.secret !== SECRET) return json({ ok: false, error: "인증 실패" });

    const entry = body.entry;
    if (!entry || !entry.id) return json({ ok: false, error: "빈 요청" });

    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName(SHEET_NAME);
    if (!sh) return json({ ok: false, error: "시트를 찾을 수 없음: " + SHEET_NAME });

    // --- 중복 방지 로그 (이미 반영된 id면 건너뜀) ---
    let log = ss.getSheetByName("_synced");
    if (!log) { log = ss.insertSheet("_synced"); log.hideSheet(); }
    const logLast = log.getLastRow();
    const ids = logLast > 0 ? log.getRange(1, 1, logLast, 1).getValues().flat().map(String) : [];
    if (ids.indexOf(String(entry.id)) !== -1) return json({ ok: true, dup: true });

    // --- 마지막 데이터 행 다음 행 찾기 (B열=계정과목 기준) ---
    const maxRows = sh.getMaxRows();
    const colB = sh.getRange(FIRST_DATA_ROW, 2, maxRows - FIRST_DATA_ROW + 1, 1).getValues();
    let last = FIRST_DATA_ROW - 1;
    for (let i = 0; i < colB.length; i++) {
      if (String(colB[i][0]).trim() !== "") last = FIRST_DATA_ROW + i;
    }
    const row = last + 1;

    // --- 입력 컬럼 기록 (E·F 원화 환산은 시트 수식이 계산) ---
    sh.getRange(row, 1).setValue(entry.date || "");      // A 일자
    sh.getRange(row, 2).setValue(entry.account || "");    // B 계정과목
    sh.getRange(row, 3).setValue(entry.fund || "");       // C 지원금/자부담
    sh.getRange(row, 4).setValue(entry.desc || "");       // D 내역
    sh.getRange(row, 7).setValue(entry.amount);           // G 사용 금액
    sh.getRange(row, 8).setValue(entry.currency || "");   // H 사용 통화
    sh.getRange(row, 9).setValue(entry.receipt || "");    // I 영수증
    sh.getRange(row, 10).setValue(entry.note || "");      // J 비고

    // --- E·F 환산 수식을 12행에서 복사해 내려서 채움 (수입/지출 자동 분기) ---
    sh.getRange(FIRST_DATA_ROW, 5, 1, 2).copyTo(sh.getRange(row, 5, 1, 2));

    log.appendRow([String(entry.id)]);
    return json({ ok: true, row: row });

  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

// 연결 확인용 (브라우저에서 URL 열면 보임)
function doGet() {
  return json({ ok: true, msg: "선교 회계 수신 서버 작동 중" });
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
