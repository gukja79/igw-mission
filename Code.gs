/* =====================================================================
   선교 오프라인 회계 — 시트 수신용 Apps Script (독립형, 전 팀 공용 1개)

   왜 독립형인가
   팀 시트 소유자가 따로 있고 운영자는 '편집자'라, 시트에 묶인(bound) 스크립트 대신
   운영자 계정의 '독립' 스크립트가 시트를 ID로 열어 기록합니다(openById).
   → 배포 1번, URL 1개로 7개 팀 모두 처리. 소유권 문제 없음.

   설치
   1) script.google.com → 새 프로젝트
   2) 이 코드 전체 붙여넣기
   3) SECRET 변경 + SHEETS 에 팀별 스프레드시트 ID 입력
   4) 배포 → 새 배포 → 유형: 웹 앱 (실행: 나 / 액세스: 모든 사용자) → URL 복사
      (최초 배포 시 '스프레드시트 접근' 권한 승인 화면이 한 번 뜸 → 허용)
   5) index.html SYNC 의 해당 팀 칸에 이 URL·secret 입력 (전 팀 동일 값)

   각 팀 시트 조건: 대상 탭에
     - I3 = 달러 기준환율 / I6 = 현지화1 환율 / I9 = 현지화2 환율
     - E12, F12 에 환산 수식(수입/지출 자동 분기) — 핸드오버 수식
   ===================================================================== */

const SECRET = "CHANGE-ME-공용-비밀키";   // ← 앱 SYNC.secret 과 동일하게
const SHEET_NAME = "수입지출내역서";        // 대상 탭 이름 (전 팀 공통)
const FIRST_DATA_ROW = 12;                  // 데이터 시작 행 (전 팀 공통)

// 팀 → 스프레드시트 ID  (시트 URL의 /d/ 와 /edit 사이 문자열)
const SHEETS = {
  "몽골":        "1xs9W_J6I1mlMDimLvrPKWLZOYeD_5NqgY86-qJByIhc",
  "라오스1":     "여기에-라오스1-시트ID",
  "중국1":       "여기에-중국1-시트ID",
  "인도네시아1": "여기에-인도네시아1-시트ID",
  "말레이시아1": "여기에-말레이시아1-시트ID",
  "캄보디아1":   "여기에-캄보디아1-시트ID",
  "스리랑카":    "여기에-스리랑카-시트ID"
};

function doPost(e){
  try{
    const body = JSON.parse(e.postData.contents);
    if(body.secret !== SECRET) return json({ ok:false, error:"인증 실패" });

    const entry = body.entry;
    if(!entry || !entry.id) return json({ ok:false, error:"빈 요청" });

    const sheetId = SHEETS[entry.team];
    if(!sheetId || sheetId.indexOf("여기에") === 0)
      return json({ ok:false, error:"팀 시트 ID 미설정: " + entry.team });

    const ss = SpreadsheetApp.openById(sheetId);
    const sh = ss.getSheetByName(SHEET_NAME);
    if(!sh) return json({ ok:false, error:"탭을 찾을 수 없음: " + SHEET_NAME });

    // 중복 방지 로그 (스프레드시트별 _synced 탭)
    let log = ss.getSheetByName("_synced");
    if(!log){ log = ss.insertSheet("_synced"); log.hideSheet(); }
    const logLast = log.getLastRow();
    const ids = logLast > 0 ? log.getRange(1,1,logLast,1).getValues().flat().map(String) : [];
    if(ids.indexOf(String(entry.id)) !== -1) return json({ ok:true, dup:true });

    // 마지막 데이터 행 다음 (B열=계정과목 기준)
    const maxRows = sh.getMaxRows();
    const colB = sh.getRange(FIRST_DATA_ROW, 2, maxRows - FIRST_DATA_ROW + 1, 1).getValues();
    let last = FIRST_DATA_ROW - 1;
    for(let i = 0; i < colB.length; i++){
      if(String(colB[i][0]).trim() !== "") last = FIRST_DATA_ROW + i;
    }
    const row = last + 1;

    // 입력 컬럼 기록 (E·F 원화 환산은 시트 수식이 계산)
    sh.getRange(row, 1).setValue(entry.date || "");      // A 일자
    sh.getRange(row, 2).setValue(entry.account || "");    // B 계정과목
    sh.getRange(row, 3).setValue(entry.fund || "");       // C 지원금/자부담
    sh.getRange(row, 4).setValue(entry.desc || "");       // D 내역
    sh.getRange(row, 7).setValue(entry.amount);           // G 사용 금액
    sh.getRange(row, 8).setValue(entry.currency || "");   // H 사용 통화
    sh.getRange(row, 9).setValue(entry.receipt || "");    // I 영수증
    sh.getRange(row, 10).setValue(entry.note || "");      // J 비고

    // E·F 환산 수식을 12행에서 복사해 내려 채움 (수입/지출 자동 분기)
    sh.getRange(FIRST_DATA_ROW, 5, 1, 2).copyTo(sh.getRange(row, 5, 1, 2));

    log.appendRow([String(entry.id)]);
    return json({ ok:true, team:entry.team, row:row });

  }catch(err){
    return json({ ok:false, error:String(err) });
  }
}

// 연결 확인용 (브라우저에서 URL 열면 보임)
function doGet(){
  return json({ ok:true, msg:"선교 회계 수신 서버(공용) 작동 중" });
}

function json(obj){
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
