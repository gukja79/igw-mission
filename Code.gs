/* =====================================================================
   선교 오프라인 회계 — 시트 수신용 Apps Script (독립형, 전 팀 공용 1개)

   왜 독립형인가
   팀 시트 소유자가 따로 있고 운영자는 '편집자'라, 시트에 묶인(bound) 스크립트 대신
   운영자 계정의 '독립' 스크립트가 시트를 ID로 열어 기록합니다(openById).
   → 배포 1번, URL 1개로 7개 팀 모두 처리. 소유권 문제 없음.

   하는 일
   - 앱이 보낸 한 건을 대상 시트 '수입지출내역서' 맨 아래에 추가
   - A 일자는 '진짜 날짜값'으로 기록 (정렬 보기 탭이 날짜순으로 깔끔히 정렬되도록)
   - E·F 원화 환산은 시트 수식이 계산 (12행 수식 복사)
   - 영수증 사진은 운영자(나) 드라이브에 저장하고 그 링크를 K열에 기록
     (J열 비고는 회계 메모용으로 비워둠)

   설치 / 재배포
   1) script.google.com → 이 코드 전체 붙여넣기
   2) SECRET 변경 + SHEETS 에 팀별 스프레드시트 ID 입력
   3) 배포 → 새 배포(또는 기존 배포 '편집 → 새 버전') → 웹 앱 (실행: 나 / 액세스: 모든 사용자)
      ※ 이번 버전부터 '드라이브 접근' 권한 승인 화면이 새로 한 번 뜸 → 허용 (사진 저장용)
   4) index.html SYNC 의 해당 팀 칸에 이 URL·secret 입력 (전 팀 동일 값)

   각 팀 시트 조건: 대상 탭에
     - I3 = 달러 기준환율 / I6 = 현지화1 환율 / I9 = 현지화2 환율
     - E12, F12 에 환산 수식(수입/지출 자동 분기)
     - K11 머리글 '사진' (사진 링크가 들어갈 칸)
   ===================================================================== */

const SECRET = "29540027";                  // ← 앱 SYNC.secret 과 동일
const SHEET_NAME = "수입지출내역서";        // 대상 탭 이름 (전 팀 공통)
const FIRST_DATA_ROW = 12;                  // 데이터 시작 행 (전 팀 공통)

// 영수증 사진 — 운영자(나) 드라이브에 저장
const PARENT_FOLDER_ID = "1J9cpOBhZKsa18NFBQ9Me0ajXfJ7NEFlT";  // 이 폴더 안에 팀 폴더 생성. 비우면("") 내 드라이브 최상단에 PHOTO_PARENT 폴더를 자동 생성.
const PHOTO_PARENT = "선교회계_영수증";     // PARENT_FOLDER_ID 가 빈 경우에만 사용
const PHOTO_SHARE  = "private";             // "private"=나만 열람 / "view"=링크 가진 사람 보기 가능

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

// L열에 entry id 기록 → 수정·삭제·재시도 모두 id 로 행을 찾는다
// (행 식별을 단일 키로. 정렬 보기 탭 FILTER 가 A:K 라 L열은 안 딸려옴)
const ID_COL = 12;
const RECEIPTNO_COL = 13; // [Phase 2] M열=영수증번호(숨김). 서버 발급 max(M)+1 기준 + pull 이 읽어 표시.

function doPost(e){
  try{
    const body = JSON.parse(e.postData.contents);
    if(body.secret !== SECRET) return json({ ok:false, error:"인증 실패" });

    // [Phase 2] 내려받기(pull): 시트 전체를 앱 거울용 JSON 으로 반환 + 미완성행 L·M 보정(backfill)
    if(body.action === "pull") return pullAll_(body);

    // 사진만 다시 올리기 (행은 그대로, K열만 교체)
    if(body.retryPhotos && body.entry && body.entry.id) return retryPhotos_(body.entry);

    // Phase 1 — 수정/삭제 (행은 L열 id 로 찾음)
    if(body.action === "update") return updateEntry_(body);
    if(body.action === "delete") return deleteEntry_(body);

    const entry = body.entry;
    if(!entry || !entry.id) return json({ ok:false, error:"빈 요청" });

    const sheetId = SHEETS[entry.team];
    if(!sheetId || sheetId.indexOf("여기에") === 0)
      return json({ ok:false, error:"팀 시트 ID 미설정: " + entry.team });

    const ss = SpreadsheetApp.openById(sheetId);
    const sh = ss.getSheetByName(SHEET_NAME);
    if(!sh) return json({ ok:false, error:"탭을 찾을 수 없음: " + SHEET_NAME });

    // 중복 방지: L열에 같은 id 가 이미 있으면 이전에 만든 행 (락 밖에서 빠른 컷)
    const dupRow = findRowById_(sh, entry.id);
    if(dupRow){
      const dupNo = sh.getRange(dupRow, RECEIPTNO_COL).getValue();
      return json({ ok:true, dup:true, row:dupRow, receiptNo:(dupNo === "" ? null : Number(dupNo)) });
    }

    // [Phase 2] 번호 발급·행 기록은 락으로 직렬화 (동시 입력 시 번호 충돌 방지)
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);   // 최대 20초 대기
    let row, receiptNo;
    try{
      // 락 안에서 중복 재확인 (대기 중 같은 id 가 먼저 들어왔을 수 있음)
      const reRow = findRowById_(sh, entry.id);
      if(reRow){
        const reNo = sh.getRange(reRow, RECEIPTNO_COL).getValue();
        return json({ ok:true, dup:true, row:reRow, receiptNo:(reNo === "" ? null : Number(reNo)) });
      }

      // 마지막 데이터 행 다음 (B열=계정과목 기준)
      const maxRows = sh.getMaxRows();
      const colB = sh.getRange(FIRST_DATA_ROW, 2, maxRows - FIRST_DATA_ROW + 1, 1).getValues();
      let last = FIRST_DATA_ROW - 1;
      for(let i = 0; i < colB.length; i++){
        if(String(colB[i][0]).trim() !== "") last = FIRST_DATA_ROW + i;
      }
      row = last + 1;

      // [Phase 2] 서버 단일 발급: M열 현재 최대값 +1
      receiptNo = nextReceiptNo_(sh);

      // 입력 컬럼 기록 (E·F 원화 환산은 시트 수식이 계산)
      sh.getRange(row, 1).setValue(toDate_(entry.date));    // A 일자 (진짜 날짜값)
      sh.getRange(row, 1).setNumberFormat("yyyy-mm-dd");
      sh.getRange(row, 2).setValue(entry.account || "");     // B 계정과목
      sh.getRange(row, 3).setValue(entry.fund || "");        // C 지원금/자부담
      sh.getRange(row, 4).setValue(entry.desc || "");        // D 내역
      sh.getRange(row, 7).setValue(entry.amount);            // G 사용 금액
      sh.getRange(row, 8).setValue(entry.currency || "");    // H 사용 통화
      sh.getRange(row, 9).setValue(entry.receipt || "");     // I 영수증
      sh.getRange(row, 10).setValue(entry.note || "");       // J 비고 (회계 메모)

      // E·F 환산 수식을 12행에서 복사해 내려 채움 (수입/지출 자동 분기)
      sh.getRange(FIRST_DATA_ROW, 5, 1, 2).copyTo(sh.getRange(row, 5, 1, 2));

      // L열 id + M열 영수증번호 (M은 숫자 서식 강제 — 날짜 서식 셀이면 3→날짜로 오인 방지)
      sh.getRange(row, ID_COL).setValue(String(entry.id));
      sh.getRange(row, RECEIPTNO_COL).setValue(receiptNo).setNumberFormat("0");

      SpreadsheetApp.flush();   // 락 풀기 전에 기록 확정
    } finally {
      lock.releaseLock();
    }

    // 영수증 사진 → 드라이브 팀 폴더 → K열 링크 (서버 발급번호로 파일명, 실패해도 행 보존)
    const pr = savePhotos_(sh, row, entry, receiptNo);

    return json({ ok:true, team:entry.team, row:row, receiptNo:receiptNo, photos:pr.ok, photoFail:pr.fail });

  }catch(err){
    return json({ ok:false, error:String(err) });
  }
}

// 사진만 다시 올리기: L열로 행 찾기 (없으면 구버전 _synced 폴백)
function retryPhotos_(entry){
  try{
    const sheetId = SHEETS[entry.team];
    if(!sheetId || sheetId.indexOf("여기에") === 0)
      return json({ ok:false, error:"팀 시트 ID 미설정: " + entry.team });

    const ss = SpreadsheetApp.openById(sheetId);
    const sh = ss.getSheetByName(SHEET_NAME);
    if(!sh) return json({ ok:false, error:"탭을 찾을 수 없음: " + SHEET_NAME });

    let row = findRowById_(sh, entry.id);
    if(!row) row = findRowInSyncedLog_(ss, entry.id);   // 구행(L열 없는 옛 행) 폴백
    if(!row) return json({ ok:false, error:"원본 행을 찾을 수 없음 (먼저 동기화하세요)" });

    const pr = savePhotos_(sh, row, entry);
    return json({ ok:true, team:entry.team, row:row, photos:pr.ok, photoFail:pr.fail });

  }catch(err){
    return json({ ok:false, error:String(err) });
  }
}

// 수정: A:D + G:J 두 블록만 덮어쓰기. E·F 수식 / K 사진 / L id 는 보존.
// 행 없음 = 시트에서 수동 삭제됨 → missing:true 로 앱이 유령 카드를 정리
function updateEntry_(body){
  try{
    if(!body.id) return json({ ok:false, error:"id 없음" });
    const sheetId = SHEETS[body.team];
    if(!sheetId || sheetId.indexOf("여기에") === 0)
      return json({ ok:false, error:"팀 시트 ID 미설정: " + body.team });

    const ss = SpreadsheetApp.openById(sheetId);
    const sh = ss.getSheetByName(SHEET_NAME);
    if(!sh) return json({ ok:false, error:"탭을 찾을 수 없음: " + SHEET_NAME });

    const row = findRowById_(sh, body.id);
    if(!row) return json({ ok:true, missing:true });

    sh.getRange(row, 1, 1, 4).setValues([[
      toDate_(body.date), body.account || "", body.fund || "", body.desc || ""
    ]]);
    sh.getRange(row, 1).setNumberFormat("yyyy-mm-dd");

    sh.getRange(row, 7, 1, 4).setValues([[
      body.amount, body.currency || "", body.receipt || "", body.note || ""
    ]]);

    return json({ ok:true, team:body.team, row:row });
  }catch(err){
    return json({ ok:false, error:String(err) });
  }
}

// 삭제: K열의 드라이브 파일을 휴지통(setTrashed)에 넣은 뒤 행 삭제. 멱등.
function deleteEntry_(body){
  try{
    if(!body.id) return json({ ok:false, error:"id 없음" });
    const sheetId = SHEETS[body.team];
    if(!sheetId || sheetId.indexOf("여기에") === 0)
      return json({ ok:false, error:"팀 시트 ID 미설정: " + body.team });

    const ss = SpreadsheetApp.openById(sheetId);
    const sh = ss.getSheetByName(SHEET_NAME);
    if(!sh) return json({ ok:false, error:"탭을 찾을 수 없음: " + SHEET_NAME });

    const row = findRowById_(sh, body.id);
    if(!row) return json({ ok:true });   // 이미 없음 → 멱등 ok

    trashPhotosFromRow_(sh, row);
    sh.deleteRow(row);

    return json({ ok:true, team:body.team, row:row });
  }catch(err){
    return json({ ok:false, error:String(err) });
  }
}

// [Phase 2] 내려받기(pull): 대상 시트의 모든 데이터 행을 앱이 거울로 읽을 형태로 반환.
// + 보정: id(L) 비어 있고 A·B·C·D·G·H 가 모두 찬 '완성행'은 이때 L=id, M=영수증번호를
//   자동 부여(위→아래 순). 누가 웹에서 입력했든 pull 한 번 거치면 빠짐없이 번호가 박힌다.
//   번호 발급은 create 와 같은 nextReceiptNo_ + 같은 락 → 앱/웹 번호 안 겹침.
function pullAll_(body){
  try{
    const team = body.team;
    const sheetId = SHEETS[team];
    if(!sheetId || sheetId.indexOf("여기에") === 0)
      return json({ ok:false, error:"팀 시트 ID 미설정: " + team });

    const ss = SpreadsheetApp.openById(sheetId);
    const sh = ss.getSheetByName(SHEET_NAME);
    if(!sh) return json({ ok:false, error:"탭을 찾을 수 없음: " + SHEET_NAME });

    const tz = ss.getSpreadsheetTimeZone();

    const maxRows = sh.getMaxRows();
    const colB = sh.getRange(FIRST_DATA_ROW, 2, maxRows - FIRST_DATA_ROW + 1, 1).getValues();
    let last = FIRST_DATA_ROW - 1;
    for(let i = 0; i < colB.length; i++){
      if(String(colB[i][0]).trim() !== "") last = FIRST_DATA_ROW + i;
    }
    if(last < FIRST_DATA_ROW) return json({ ok:true, team:team, count:0, rows:[] });

    const n = last - FIRST_DATA_ROW + 1;
    const vals = sh.getRange(FIRST_DATA_ROW, 1, n, RECEIPTNO_COL).getValues();

    // [Phase 2] 보정 대상 탐지: id(L) 비어 있고 A·B·C·D·G·H 가 모두 찬 완성행(위→아래)
    const need = [];
    for(let i = 0; i < n; i++){
      const r = vals[i];
      if(!cellFilled_(r[ID_COL - 1]) && rowComplete_(r)) need.push(i);
    }

    if(need.length){
      const lock = LockService.getScriptLock();
      lock.waitLock(20000);
      try{
        // 락 안에서 L열만 신선하게 다시 읽어 '동시 pull' 중복 발급을 막는다
        const idCol = sh.getRange(FIRST_DATA_ROW, ID_COL, n, 1).getValues();
        let nextNo = nextReceiptNo_(sh);   // create 와 공유: max(M)+1
        for(let k = 0; k < need.length; k++){
          const i = need[k];
          if(cellFilled_(idCol[i][0])) continue;   // 그새 채워졌으면 건너뜀
          const newId = makeBackfillId_();
          sh.getRange(FIRST_DATA_ROW + i, ID_COL).setValue(newId);                          // L
          sh.getRange(FIRST_DATA_ROW + i, RECEIPTNO_COL).setValue(nextNo).setNumberFormat("0"); // M (숫자 서식 강제)
          vals[i][ID_COL - 1] = newId;             // 반환 JSON 에도 즉시 반영
          vals[i][RECEIPTNO_COL - 1] = nextNo;
          nextNo++;
        }
        SpreadsheetApp.flush();   // 락 풀기 전에 기록 확정
      } finally {
        lock.releaseLock();
      }
    }

    const rows = [];
    for(let i = 0; i < n; i++){
      const r = vals[i];
      const account = String(r[1] || "").trim();
      if(!account) continue;
      rows.push({
        row:       FIRST_DATA_ROW + i,
        id:        String(r[ID_COL - 1] || ""),
        receiptNo: receiptNoOf_(r[RECEIPTNO_COL - 1]),  // M (Date 오염값은 null 처리)
        date:      fmtDate_(r[0], tz),
        account:   account,
        fund:      String(r[2] || ""),
        desc:      String(r[3] || ""),
        krwIn:     num_(r[4]),
        krwOut:    num_(r[5]),
        amount:    num_(r[6]),
        currency:  String(r[7] || ""),
        receipt:   String(r[8] || ""),
        note:      String(r[9] || ""),
        photo:     String(r[10] || "")
      });
    }
    return json({ ok:true, team:team, count:rows.length, backfilled:need.length, rows:rows });
  }catch(err){
    return json({ ok:false, error:String(err) });
  }
}

// [Phase 2] 날짜 셀 → "yyyy-MM-dd" 문자열 (Date 면 시트 타임존으로 포맷, 아니면 문자열 그대로)
function fmtDate_(v, tz){
  if(v instanceof Date) return Utilities.formatDate(v, tz, "yyyy-MM-dd");
  return String(v || "");
}

// [Phase 2] 숫자 셀 → 숫자 (빈칸/수식 "" 는 0)
function num_(v){
  return (typeof v === "number") ? v : 0;
}

// [Phase 2] M열(영수증번호) 안전 변환: 빈칸/Date(과거 날짜서식 오염행) → null, 정상 숫자만 반환
function receiptNoOf_(v){
  if(v === "" || v == null) return null;
  if(v instanceof Date) return null;           // 날짜로 오염된 값은 번호로 보지 않음
  const n = Number(v);
  return (isNaN(n) || n <= 0) ? null : n;
}

// L열에서 id 가 있는 행 찾기 (없으면 0)
function findRowById_(sh, id){
  const last = sh.getLastRow();
  if(last < FIRST_DATA_ROW) return 0;
  const n = last - FIRST_DATA_ROW + 1;
  const col = sh.getRange(FIRST_DATA_ROW, ID_COL, n, 1).getValues();
  const want = String(id);
  for(let i = 0; i < col.length; i++){
    if(String(col[i][0]) === want) return FIRST_DATA_ROW + i;
  }
  return 0;
}

// 구버전 _synced 로그에서 [id, row] 찾기 — L열 없는 옛 행의 재시도 폴백 전용
function findRowInSyncedLog_(ss, id){
  const log = ss.getSheetByName("_synced");
  if(!log) return 0;
  const last = log.getLastRow();
  if(last === 0) return 0;
  const rows = log.getRange(1, 1, last, 2).getValues();
  for(let i = 0; i < rows.length; i++){
    if(String(rows[i][0]) === String(id)) return Number(rows[i][1]) || 0;
  }
  return 0;
}

// K열 링크에서 드라이브 파일 id 를 뽑아 휴지통으로 (영구삭제 X)
function trashPhotosFromRow_(sh, row){
  const kVal = String(sh.getRange(row, 11).getValue() || "");
  if(!kVal) return;
  const re = /\/d\/([A-Za-z0-9_-]{20,})/g;
  let m;
  while((m = re.exec(kVal)) !== null){
    try{ DriveApp.getFileById(m[1]).setTrashed(true); }catch(_){}
  }
}

// 사진 저장 공통 — K열에 링크 기록, {ok, fail} 반환
function savePhotos_(sh, row, entry, receiptNo){
  let okN = 0, failN = 0;
  try{
    const pics = entry.photos || [];
    if(!pics.length) return { ok:0, fail:0 };

    const folder = getTeamFolder_(entry.team);
    const acct = String(entry.account || "").replace(/[\\/:*?"<>|]/g, "-");
    const base = pad3_(receiptNo != null ? receiptNo : entry.receiptNo) + "_" + (entry.date || "nodate") + "_" + acct;
    const out = [];
    for(let i = 0; i < pics.length; i++){
      const b64  = (typeof pics[i] === "string") ? pics[i] : pics[i].data;   // 신/구버전 호환
      const want = (pics[i] && pics[i].size) ? Number(pics[i].size) : 0;
      const bytes = Utilities.base64Decode(b64);
      if(want && bytes.length !== want){
        out.push("⚠ 사진" + (i + 1) + " 전송 손상(" + bytes.length + "/" + want + ")");
        failN++;
        continue;
      }
      const nm = base + (pics.length > 1 ? "_" + (i + 1) : "") + ".jpg";
      const f = folder.createFile(Utilities.newBlob(bytes, "image/jpeg", nm));
      if(PHOTO_SHARE === "view"){
        try{ f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); }catch(_){}
      }
      out.push(f.getUrl());
      okN++;
    }
    if(out.length) sh.getRange(row, 11).setValue(out.join("\n"));  // K 사진
  }catch(perr){
    sh.getRange(row, 11).setValue("사진 업로드 오류: " + String(perr));
    failN++;
  }
  return { ok:okN, fail:failN };
}

// 연결 확인용 (브라우저에서 URL 열면 보임)
function doGet(){
  return json({ ok:true, msg:"선교 회계 수신 서버(공용) 작동 중" });
}

/* ---------------- 헬퍼 ---------------- */

// "2026-06-21" → 진짜 날짜값 (로컬 자정, 타임존 밀림 없음)
function toDate_(s){
  if(!s) return "";
  const p = String(s).split("-");
  if(p.length !== 3) return s;
  return new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]));
}

// 영수증번호 3자리 0채움 (1 → "001")
function pad3_(n){
  const s = String(n == null ? "" : n);
  return s.length >= 3 ? s : ("000" + s).slice(-3);
}

// [Phase 2] 셀이 '채워졌는지' 판정 (날짜·숫자는 채워진 것으로, 문자열은 trim 후 비교)
function cellFilled_(v){
  if(v === null || v === undefined) return false;
  if(v instanceof Date) return true;
  if(typeof v === "number") return true;
  return String(v).trim() !== "";
}

// [Phase 2] '완성행' 판정: A·B·C·D·G·H 가 모두 채워졌는가 (E·F 는 수식, I·J 는 선택값이라 제외)
function rowComplete_(r){
  return cellFilled_(r[0]) && cellFilled_(r[1]) && cellFilled_(r[2]) &&
         cellFilled_(r[3]) && cellFilled_(r[6]) && cellFilled_(r[7]);
}

// [Phase 2] 보정행용 id 생성 — 앱 id 형식과 호환(밀리초-랜덤), 끝에 'w'로 웹 보정 표시
function makeBackfillId_(){
  return String(Date.now()) + "-" + Math.random().toString(36).slice(2, 8) + "w";
}

// [Phase 2] M열(영수증번호) 현재 최대값 +1. 빈 시트면 1부터.
function nextReceiptNo_(sh){
  const last = sh.getLastRow();
  if(last < FIRST_DATA_ROW) return 1;
  const n = last - FIRST_DATA_ROW + 1;
  const col = sh.getRange(FIRST_DATA_ROW, RECEIPTNO_COL, n, 1).getValues();
  let mx = 0;
  for(let i = 0; i < col.length; i++){
    const v = Number(col[i][0]);
    if(!isNaN(v) && v > mx) mx = v;
  }
  return mx + 1;
}

function getOrCreateFolder_(parent, name){
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

function getTeamFolder_(team){
  const parent = PARENT_FOLDER_ID
    ? DriveApp.getFolderById(PARENT_FOLDER_ID)
    : getOrCreateFolder_(DriveApp.getRootFolder(), PHOTO_PARENT);
  return getOrCreateFolder_(parent, team);
}

function json(obj){
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
