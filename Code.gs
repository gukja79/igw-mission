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
  "중국1":       "1d6Wy6XnR6H45WQs8cUhrVqgD-RHswz6foMGtGqhee1s",
  "인도네시아1": "12ZlQ8bTSGSH7lQw2RlTwI-Jfg0Mblp6LsFX1cOeuMvg",
  "말레이시아1": "1RNN6EXY4f6odesLD_81TDBABWhO4-F8jyZpRGerVwFs",
  "스리랑카":    "1taiBNG4-x93FQgTricG1uoT7-qcG0Ej3AowxD4o9uy4"
};

// L열에 entry id 기록 → 수정·삭제·재시도 모두 id 로 행을 찾는다
// (행 식별을 단일 키로. 정렬 보기 탭 FILTER 가 A:K 라 L열은 안 딸려옴)
const ID_COL = 12;
const RECEIPTNO_COL = 13; // [Phase 2] M열=영수증번호(숨김). 서버 발급 max(M)+1 기준 + pull 이 읽어 표시.

// [§14] 환전내역 시트 — 통화별 가중평균 환율 자동 산출 (메인 I3/I6/I9 는 시트수식이 알아서 끌어감)
const FX_SHEET_NAME = "환전내역";  // 환전 기록 탭 (몽골만 신설·검증됨, 나머지 팀은 세팅 전까지 "탭 없음")
const FX_FIRST_ROW  = 2;          // 데이터 시작 행 (보조표 J1:M4 는 다른 열이라 안 겹침)
const FX_ID_COL     = 8;          // H열 id (메인 L열과 동형, 수정·삭제·백필 식별)
const FX_PHOTO_COL  = 7;          // G열 사진 링크 (메인 K열과 동형)
const FX_FOLDER     = "환전";      // 사진 저장 하위폴더 {팀}/환전 (영수증 사진과 분리)
// 환전내역 열: A일자 B준통화 C준금액 D받은통화 E받은금액 F메모 G사진 H id
// 보조표(서버 무관, 시트수식): J통화 K받은총량 L들어간원화 M평균환율 / J2달러 J3현지화1 J4현지화2

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

    // [Phase 3] 사진 관리 — 행은 그대로, K열 사진만 추가/단건삭제
    if(body.action === "addPhotos")   return addPhotos_(body);
    if(body.action === "deletePhoto") return deletePhoto_(body);

    // [§14] 환전내역 동기화 (별도 탭. 메인 내역서와 행시작·id열·사진열이 달라 전용 함수)
    if(body.action === "fxCreate")         return fxCreate_(body);
    if(body.action === "fxPull")           return fxPull_(body);
    if(body.action === "fxUpdate")         return fxUpdate_(body);
    if(body.action === "fxDelete")         return fxDelete_(body);
    if(body.action === "fxRenameCurrency") return fxRenameCurrency_(body);
    if(body.action === "fxAddPhotos")      return fxAddPhotos_(body);
    if(body.action === "fxDeletePhoto")    return fxDeletePhoto_(body);

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

// [Phase 3] 사진 추가 — 기존 K열 링크 보존하고 새 사진을 뒤에 append (번호 이어감)
function addPhotos_(body){
  try{
    const entry = body.entry || {};
    if(!entry.id) return json({ ok:false, error:"빈 요청(사진 추가)" });
    const pics = entry.photos || body.photos || [];
    if(!pics.length) return json({ ok:false, error:"추가할 사진이 없음" });

    const sheetId = SHEETS[entry.team];
    if(!sheetId || sheetId.indexOf("여기에") === 0)
      return json({ ok:false, error:"팀 시트 ID 미설정: " + entry.team });
    const ss = SpreadsheetApp.openById(sheetId);
    const sh = ss.getSheetByName(SHEET_NAME);
    if(!sh) return json({ ok:false, error:"탭을 찾을 수 없음: " + SHEET_NAME });
    const row = findRowById_(sh, entry.id);
    if(!row) return json({ ok:false, error:"원본 행을 찾을 수 없음 (먼저 동기화하세요)" });

    const kCell = sh.getRange(row, 11);
    const existing = String(kCell.getValue() || "").split("\n").filter(s => s.trim() !== "");
    const existingLinks = existing.filter(s => /\/d\/[A-Za-z0-9_-]{20,}/.test(s));  // ⚠ 손상줄은 번호계산서 제외

    const folder = getTeamFolder_(entry.team);
    const receiptNo = sh.getRange(row, RECEIPTNO_COL).getValue();
    const aVal = sh.getRange(row, 1).getValue();
    const dateStr = (aVal instanceof Date)
      ? Utilities.formatDate(aVal, Session.getScriptTimeZone(), "yyyy-MM-dd")
      : String(aVal || "nodate");
    const acct = String(sh.getRange(row, 2).getValue() || "").replace(/[\\/:*?"<>|]/g, "-");
    const base = pad3_(receiptNo) + "_" + dateStr + "_" + acct;

    const added = [];
    let okN = 0, failN = 0;
    for(let i = 0; i < pics.length; i++){
      const b64  = (typeof pics[i] === "string") ? pics[i] : pics[i].data;
      const want = (pics[i] && pics[i].size) ? Number(pics[i].size) : 0;
      const bytes = Utilities.base64Decode(b64);
      if(want && bytes.length !== want){
        added.push("⚠ 사진 전송 손상(" + bytes.length + "/" + want + ")");
        failN++; continue;
      }
      const n  = existingLinks.length + okN + 1;          // 절대 위치로 파일명 번호 이어감
      const nm = base + "_" + n + ".jpg";
      const f  = folder.createFile(Utilities.newBlob(bytes, "image/jpeg", nm));
      if(PHOTO_SHARE === "view"){
        try{ f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); }catch(_){}
      }
      added.push(f.getUrl());
      okN++;
    }
    const merged = existing.concat(added).join("\n");
    kCell.setValue(merged);
    return json({ ok:true, team:entry.team, row:row, photos:okN, photoFail:failN, photo:merged });
  }catch(err){ return json({ ok:false, error:String(err) }); }
}

// [Phase 3] 사진 단건 삭제 — 그 fileId만 휴지통 + K열에서 그 줄만 제거
function deletePhoto_(body){
  try{
    const entry  = body.entry || {};
    const fileId = String(body.fileId || "");
    if(!entry.id || !fileId) return json({ ok:false, error:"빈 요청(사진 삭제)" });

    const sheetId = SHEETS[entry.team];
    if(!sheetId || sheetId.indexOf("여기에") === 0)
      return json({ ok:false, error:"팀 시트 ID 미설정: " + entry.team });
    const ss = SpreadsheetApp.openById(sheetId);
    const sh = ss.getSheetByName(SHEET_NAME);
    if(!sh) return json({ ok:false, error:"탭을 찾을 수 없음: " + SHEET_NAME });
    const row = findRowById_(sh, entry.id);
    if(!row) return json({ ok:false, error:"원본 행을 찾을 수 없음 (먼저 동기화하세요)" });

    const kCell = sh.getRange(row, 11);
    const lines = String(kCell.getValue() || "").split("\n");
    const kept  = lines.filter(ln => ln.indexOf(fileId) === -1);
    try{ DriveApp.getFileById(fileId).setTrashed(true); }catch(_){}  // 실패해도 무시(이미 없을 수 있음)
    kCell.setValue(kept.join("\n"));
    return json({ ok:true, team:entry.team, row:row, photo:kept.join("\n") });
  }catch(err){ return json({ ok:false, error:String(err) }); }
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

// L열에서 id 가 있는 행 찾기 (없으면 0). firstRow·idCol 생략 시 메인 내역서 기준(12/L).
function findRowById_(sh, id, firstRow, idCol){
  firstRow = firstRow || FIRST_DATA_ROW;
  idCol    = idCol    || ID_COL;
  const last = sh.getLastRow();
  if(last < firstRow) return 0;
  const n = last - firstRow + 1;
  const col = sh.getRange(firstRow, idCol, n, 1).getValues();
  const want = String(id);
  for(let i = 0; i < col.length; i++){
    if(String(col[i][0]) === want) return firstRow + i;
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

// 사진열 링크에서 드라이브 파일 id 를 뽑아 휴지통으로 (영구삭제 X). photoCol 생략 시 메인 K(11).
function trashPhotosFromRow_(sh, row, photoCol){
  photoCol = photoCol || 11;
  const kVal = String(sh.getRange(row, photoCol).getValue() || "");
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

/* ===================================================================
   [§14] 환전내역 동기화 — 통화별 가중평균 환율 자동 산출
   메인 내역서와 다른 점: 데이터 2행부터 / id=H열 / 사진=G열 / 영수증번호 없음 / E·F 수식 없음.
   I3/6/9(메인 환율칸)·보조표 J:M(평균환율 수식)은 서버가 안 건드림 — 시트가 알아서 계산.
   =================================================================== */

// 팀 → 환전내역 시트 핸들 (없으면 throw, 각 함수 try/catch 가 잡아 json 으로 반환)
function fxSheetOf_(team){
  const sheetId = SHEETS[team];
  if(!sheetId || sheetId.indexOf("여기에") === 0) throw new Error("팀 시트 ID 미설정: " + team);
  const ss = SpreadsheetApp.openById(sheetId);
  const sh = ss.getSheetByName(FX_SHEET_NAME);
  if(!sh) throw new Error("탭을 찾을 수 없음: " + FX_SHEET_NAME + " (해당 팀 환전 시트 미세팅)");
  return { ss:ss, sh:sh };
}

// 환전 사진 폴더: {팀}/환전 (영수증 사진과 분리)
function getFxFolder_(team){
  return getOrCreateFolder_(getTeamFolder_(team), FX_FOLDER);
}

// 완성행 판정(환전): A·B·C·D·E 모두 채워짐 (F 메모·G 사진은 선택)
function fxRowComplete_(r){
  return cellFilled_(r[0]) && cellFilled_(r[1]) && cellFilled_(r[2]) &&
         cellFilled_(r[3]) && cellFilled_(r[4]);
}

// 환전 기록의 마지막 데이터 행 (B열 준통화 기준. 보조표 J:M 은 다른 열이라 안 걸림)
function fxLastRow_(sh){
  const maxRows = sh.getMaxRows();
  const colB = sh.getRange(FX_FIRST_ROW, 2, maxRows - FX_FIRST_ROW + 1, 1).getValues();
  let last = FX_FIRST_ROW - 1;
  for(let i = 0; i < colB.length; i++){
    if(String(colB[i][0]).trim() !== "") last = FX_FIRST_ROW + i;
  }
  return last;
}

// 보조표 J2:M4 스냅샷 → {usd, local1, local2} 각 {name, rate}. 빈 슬롯은 name:"" rate:null
function fxRates_(sh){
  const g = sh.getRange("J2:M4").getValues();   // 행 2,3,4 / 열 J,K,L,M
  function pick(a){
    const name = String(a[0] || "").trim();      // J 통화명
    const m    = a[3];                            // M 평균환율
    return { name:name, rate:(typeof m === "number" && !isNaN(m)) ? m : null };
  }
  return { usd:pick(g[0]), local1:pick(g[1]), local2:pick(g[2]) };
}

// [§14] 통화 코드 → 실제 통화명 변환. 앱은 코드(원화/달러/현지화1/현지화2)를 보내고
//   서버가 보조표 J셀(J2달러·J3현지화1·J4현지화2)을 읽어 시트 B·D엔 '이름'으로 기록한다.
//   (보조표 SUMIFS 가 D열을 J 라벨과 매칭하므로 B·D 는 반드시 이름이어야 함.)
//   이미 이름으로 들어온 값(코드 아님)은 그대로 통과 → 앱/서버 양쪽 형식에 안전.
function fxCcyName_(sh, code){
  const c = String(code || "").trim();
  if(c === "" || c === "원화") return "원화";
  var cell = (c === "달러") ? "J2" : (c === "현지화1") ? "J3" : (c === "현지화2") ? "J4" : "";
  if(!cell) return c;   // 이미 이름(코드 아님)이면 그대로
  var v = String(sh.getRange(cell).getValue() || "").trim();
  return v || c;        // J 라벨 비어 있으면 코드 폴백(달러는 코드==이름이라 무해)
}

// 환전 사진 저장 — G열에 링크 기록, {ok,fail}. 영수증번호 없으니 파일명=날짜+통화쌍.
function savePhotosG_(sh, row, entry){
  let okN = 0, failN = 0;
  try{
    const pics = entry.photos || [];
    if(!pics.length) return { ok:0, fail:0 };

    const folder = getFxFolder_(entry.team);
    const fromC = String(entry.fromCurrency || "").replace(/[\\/:*?"<>|]/g, "-");
    const toC   = String(entry.toCurrency   || "").replace(/[\\/:*?"<>|]/g, "-");
    const base  = (entry.date || "nodate") + "_" + fromC + "-" + toC;
    const out = [];
    for(let i = 0; i < pics.length; i++){
      const b64  = (typeof pics[i] === "string") ? pics[i] : pics[i].data;
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
    if(out.length) sh.getRange(row, FX_PHOTO_COL).setValue(out.join("\n"));   // G 사진
  }catch(perr){
    sh.getRange(row, FX_PHOTO_COL).setValue("사진 업로드 오류: " + String(perr));
    failN++;
  }
  return { ok:okN, fail:failN };
}

// 환전 1건 추가: A:F 입력 + H id + G 사진. (번호발급·E/F수식 복사 없음 → 메인보다 단순)
function fxCreate_(body){
  try{
    const entry = body.entry || {};
    if(!entry.id) return json({ ok:false, error:"빈 요청(환전)" });
    const { sh } = fxSheetOf_(entry.team);

    // 중복 컷: H열에 같은 id 있으면 이전에 만든 행 (락 밖 빠른 컷)
    const dup = findRowById_(sh, entry.id, FX_FIRST_ROW, FX_ID_COL);
    if(dup) return json({ ok:true, dup:true, row:dup });

    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    let row;
    try{
      const re = findRowById_(sh, entry.id, FX_FIRST_ROW, FX_ID_COL);   // 락 안 재확인
      if(re) return json({ ok:true, dup:true, row:re });

      row = fxLastRow_(sh) + 1;
      sh.getRange(row, 1).setValue(toDate_(entry.date));        // A 일자 (진짜 날짜값)
      sh.getRange(row, 1).setNumberFormat("yyyy-mm-dd");
      sh.getRange(row, 2).setValue(fxCcyName_(sh, entry.fromCurrency));   // B 준통화 (코드→이름)
      sh.getRange(row, 3).setValue(entry.fromAmount);           // C 준금액
      sh.getRange(row, 4).setValue(fxCcyName_(sh, entry.toCurrency));     // D 받은통화 (코드→이름)
      sh.getRange(row, 5).setValue(entry.toAmount);             // E 받은금액
      sh.getRange(row, 6).setValue(entry.memo || "");           // F 메모
      sh.getRange(row, FX_ID_COL).setValue(String(entry.id));   // H id
      SpreadsheetApp.flush();
    } finally {
      lock.releaseLock();
    }

    const pr = savePhotosG_(sh, row, entry);   // G 사진 (실패해도 행 보존)
    return json({ ok:true, team:entry.team, row:row, photos:pr.ok, photoFail:pr.fail });
  }catch(err){
    return json({ ok:false, error:String(err) });
  }
}

// 환전 내려받기: 환전 기록 목록 + 현재 평균환율(rates) + 완성행 H=id 백필
function fxPull_(body){
  try{
    const team = body.team;
    const { ss, sh } = fxSheetOf_(team);
    const tz = ss.getSpreadsheetTimeZone();
    const rates = fxRates_(sh);                       // J2:M4 스냅샷 (앱 환율 탭 표시용)

    const last = fxLastRow_(sh);
    if(last < FX_FIRST_ROW)
      return json({ ok:true, team:team, count:0, backfilled:0, rows:[], rates:rates });

    const n = last - FX_FIRST_ROW + 1;
    const vals = sh.getRange(FX_FIRST_ROW, 1, n, FX_ID_COL).getValues();   // A..H

    // 백필 대상: H(id) 비고 A·B·C·D·E 다 찬 완성행 (손입력행에 안정적 식별자 부여)
    const need = [];
    for(let i = 0; i < n; i++){
      if(!cellFilled_(vals[i][FX_ID_COL - 1]) && fxRowComplete_(vals[i])) need.push(i);
    }
    if(need.length){
      const lock = LockService.getScriptLock();
      lock.waitLock(20000);
      try{
        const idCol = sh.getRange(FX_FIRST_ROW, FX_ID_COL, n, 1).getValues();   // 락 안 신선 재확인
        for(let k = 0; k < need.length; k++){
          const i = need[k];
          if(cellFilled_(idCol[i][0])) continue;
          const newId = makeBackfillId_();
          sh.getRange(FX_FIRST_ROW + i, FX_ID_COL).setValue(newId);
          vals[i][FX_ID_COL - 1] = newId;   // 반환 JSON 에도 즉시 반영
        }
        SpreadsheetApp.flush();
      } finally {
        lock.releaseLock();
      }
    }

    const rows = [];
    for(let i = 0; i < n; i++){
      const r = vals[i];
      if(String(r[1] || "").trim() === "") continue;   // B(준통화) 빈 행 스킵
      rows.push({
        row:          FX_FIRST_ROW + i,
        id:           String(r[FX_ID_COL - 1] || ""),
        date:         fmtDate_(r[0], tz),
        fromCurrency: String(r[1] || ""),
        fromAmount:   num_(r[2]),
        toCurrency:   String(r[3] || ""),
        toAmount:     num_(r[4]),
        memo:         String(r[5] || ""),
        photo:        String(r[6] || "")    // G
      });
    }
    return json({ ok:true, team:team, count:rows.length, backfilled:need.length, rows:rows, rates:rates });
  }catch(err){
    return json({ ok:false, error:String(err) });
  }
}

// 환전 수정: A:F 덮어쓰기. G 사진·H id 보존. 행 없으면 missing:true (앱이 유령 카드 정리).
function fxUpdate_(body){
  try{
    if(!body.id) return json({ ok:false, error:"id 없음" });
    const { sh } = fxSheetOf_(body.team);
    const row = findRowById_(sh, body.id, FX_FIRST_ROW, FX_ID_COL);
    if(!row) return json({ ok:true, missing:true });

    sh.getRange(row, 1, 1, 6).setValues([[
      toDate_(body.date), fxCcyName_(sh, body.fromCurrency), body.fromAmount,
      fxCcyName_(sh, body.toCurrency), body.toAmount, body.memo || ""
    ]]);
    sh.getRange(row, 1).setNumberFormat("yyyy-mm-dd");
    return json({ ok:true, team:body.team, row:row });
  }catch(err){
    return json({ ok:false, error:String(err) });
  }
}

// 환전 삭제: G열 사진 휴지통 → 행 삭제. 멱등.
function fxDelete_(body){
  try{
    if(!body.id) return json({ ok:false, error:"id 없음" });
    const { sh } = fxSheetOf_(body.team);
    const row = findRowById_(sh, body.id, FX_FIRST_ROW, FX_ID_COL);
    if(!row) return json({ ok:true });   // 이미 없음 → 멱등 ok

    trashPhotosFromRow_(sh, row, FX_PHOTO_COL);
    sh.deleteRow(row);
    return json({ ok:true, team:body.team, row:row });
  }catch(err){
    return json({ ok:false, error:String(err) });
  }
}

// [§14-2단계] 환전 사진 추가 — G열 기존 링크 보존하고 새 사진 뒤에 append (원장 addPhotos_ 의 G열판)
function fxAddPhotos_(body){
  try{
    const entry = body.entry || {};
    if(!entry.id) return json({ ok:false, error:"빈 요청(환전 사진 추가)" });
    const pics = entry.photos || body.photos || [];
    if(!pics.length) return json({ ok:false, error:"추가할 사진이 없음" });

    const { sh } = fxSheetOf_(entry.team);
    const row = findRowById_(sh, entry.id, FX_FIRST_ROW, FX_ID_COL);
    if(!row) return json({ ok:false, error:"원본 행을 찾을 수 없음 (먼저 동기화하세요)" });

    const gCell = sh.getRange(row, FX_PHOTO_COL);
    const existing = String(gCell.getValue() || "").split("\n").filter(s => s.trim() !== "");
    const existingLinks = existing.filter(s => /\/d\/[A-Za-z0-9_-]{20,}/.test(s));   // 손상줄 제외

    const folder = getFxFolder_(entry.team);
    const aVal = sh.getRange(row, 1).getValue();
    const dateStr = (aVal instanceof Date)
      ? Utilities.formatDate(aVal, Session.getScriptTimeZone(), "yyyy-MM-dd")
      : String(aVal || "nodate");
    const fromC = String(sh.getRange(row, 2).getValue() || "").replace(/[\\/:*?"<>|]/g, "-");
    const toC   = String(sh.getRange(row, 4).getValue() || "").replace(/[\\/:*?"<>|]/g, "-");
    const base  = dateStr + "_" + fromC + "-" + toC;

    const added = [];
    let okN = 0, failN = 0;
    for(let i = 0; i < pics.length; i++){
      const b64  = (typeof pics[i] === "string") ? pics[i] : pics[i].data;
      const want = (pics[i] && pics[i].size) ? Number(pics[i].size) : 0;
      const bytes = Utilities.base64Decode(b64);
      if(want && bytes.length !== want){ added.push("⚠ 사진 전송 손상(" + bytes.length + "/" + want + ")"); failN++; continue; }
      const n  = existingLinks.length + okN + 1;
      const nm = base + "_" + n + ".jpg";
      const f  = folder.createFile(Utilities.newBlob(bytes, "image/jpeg", nm));
      if(PHOTO_SHARE === "view"){ try{ f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); }catch(_){} }
      added.push(f.getUrl());
      okN++;
    }
    const merged = existing.concat(added).join("\n");
    gCell.setValue(merged);
    return json({ ok:true, team:entry.team, row:row, photos:okN, photoFail:failN, photo:merged });
  }catch(err){ return json({ ok:false, error:String(err) }); }
}

// [§14-2단계] 환전 사진 단건 삭제 — 그 fileId만 휴지통 + G열에서 그 줄만 제거 (원장 deletePhoto_ 의 G열판)
function fxDeletePhoto_(body){
  try{
    const entry  = body.entry || {};
    const fileId = String(body.fileId || "");
    if(!entry.id || !fileId) return json({ ok:false, error:"빈 요청(환전 사진 삭제)" });

    const { sh } = fxSheetOf_(entry.team);
    const row = findRowById_(sh, entry.id, FX_FIRST_ROW, FX_ID_COL);
    if(!row) return json({ ok:false, error:"원본 행을 찾을 수 없음 (먼저 동기화하세요)" });

    const gCell = sh.getRange(row, FX_PHOTO_COL);
    const lines = String(gCell.getValue() || "").split("\n");
    const kept  = lines.filter(ln => ln.indexOf(fileId) === -1);
    const removed = lines.length - kept.length;
    try{ DriveApp.getFileById(fileId).setTrashed(true); }catch(_){}
    gCell.setValue(kept.join("\n"));
    SpreadsheetApp.flush();
    return json({ ok:true, team:entry.team, row:row, removed:removed, photo:kept.join("\n") });
  }catch(err){ return json({ ok:false, error:String(err) }); }
}

// 통화명 일괄 갱신 — 환전내역 B·D(준/받은통화) + 보조표 J 라벨을 새 이름으로.
//   slot:"local1"→J3 / "local2"→J4. B·D 와 J 를 같이 바꿔야 보조표 SUMIFS 가 안 깨짐 → 락으로 묶음.
//   메인 시트는 코드(현지화N)·셀참조라 무관 → 안 건드림.
function fxRenameCurrency_(body){
  try{
    const oldName = String(body.oldName || "").trim();
    const newName = String(body.newName || "").trim();
    if(!newName) return json({ ok:false, error:"newName 필요" });   // oldName 빈값 = 최초 설정(라벨만)
    if(oldName === newName)  return json({ ok:true, changedBD:0, changedLabel:"(동일)" });
    const labelCell = (body.slot === "local1") ? "J3" : (body.slot === "local2") ? "J4" : "";
    if(!labelCell) return json({ ok:false, error:"slot 은 local1/local2" });

    const { sh } = fxSheetOf_(body.team);

    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    let changed = 0;
    try{
      // oldName 이 있을 때만 기존 환전 기록 B·D 치환 (최초 설정이면 라벨만 기록 → 빈칸 오염 방지)
      if(oldName){
        const last = fxLastRow_(sh);
        if(last >= FX_FIRST_ROW){
          const n = last - FX_FIRST_ROW + 1;
          const bd = sh.getRange(FX_FIRST_ROW, 2, n, 3).getValues();   // B,C,D (C 는 안 건드림)
          for(let i = 0; i < n; i++){
            if(String(bd[i][0]).trim() === oldName){ sh.getRange(FX_FIRST_ROW + i, 2).setValue(newName); changed++; }  // B
            if(String(bd[i][2]).trim() === oldName){ sh.getRange(FX_FIRST_ROW + i, 4).setValue(newName); changed++; }  // D
          }
        }
      }
      sh.getRange(labelCell).setValue(newName);   // 보조표 J 라벨 (SUMIFS 매칭 키)
      SpreadsheetApp.flush();
    } finally {
      lock.releaseLock();
    }
    return json({ ok:true, team:body.team, changedBD:changed, changedLabel:labelCell });
  }catch(err){
    return json({ ok:false, error:String(err) });
  }
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
