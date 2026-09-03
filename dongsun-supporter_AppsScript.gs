/**
 * 동선 서포터즈 트래커 — Google Apps Script API (구글시트 백엔드)
 *
 * ⚠ 매우 중요 — 기존 탭 보호
 *   이 스크립트는 "입점매장_시흥_액션팀용" 스프레드시트의 기존 17개 탭(동네별·상태별·이력용)을
 *   절대 읽거나 쓰지 않습니다. 오직 아래 두 개의 신규 탭만 사용합니다.
 *     · "트래커" — 서포터즈 TA 진행 데이터 (구조화 컬럼)
 *     · "계정"   — 로그인용 이름/비번
 *   시트 이름으로만 명시 접근하므로 다른 탭은 건드리지 않습니다.
 *
 * 사용법:
 *  1) 아래 SPREADSHEET_ID / CONSULTANT_SPREADSHEET_ID 확인 (이미 채워져 있음)
 *  2) 함수 목록에서 setupTrackerSheet 실행 → "트래커" 탭 생성(빈 상태, 헤더만)
 *  3) 함수 목록에서 setupAccountSheet 실행 → "계정" 탭 생성 후 시트에서 직접 이름/비번 입력
 *  4) 배포 → 새 배포 → 유형: 웹 앱 → 실행: 나 / 액세스: 모든 사용자 → 배포
 *  5) 나온 웹앱 URL(.../exec)을 dongsun-supporter.html 최초 접속 화면에 입력
 *
 * ※ 과거 데이터 자동 이관은 하지 않습니다. 필요하면 사용자가 "트래커" 탭에 직접 복사해 넣으세요.
 */

// ── 스프레드시트 ID ───────────────────────────────────────────
// 입점매장_시흥_액션팀용 (원천DB · 서포터즈 TA)
var SPREADSHEET_ID = "1ewvEx1GdEzhVsIdbymxevSGbdusimLNcW0PQEvqQamw";
// 동선_컨설팅 DB 관리용 (전환 대상 · 컨설턴트 트래커)
var CONSULTANT_SPREADSHEET_ID = "1AtJ_qLMzsyRCuSAH-cGNCEhribVCt1qBY0a5GUzc9SI";

// ── 신규 탭 이름 (기존 탭과 절대 겹치지 않게) ──────────────────
var TRACKER_SHEET = "트래커";
var ACCOUNT_SHEET = "계정";
// 컨설턴트 시트 쪽 신규 탭 (dongsun-consultant_AppsScript.gs 와 동일해야 함)
var CONSULTANT_TRACKER_SHEET = "트래커";

var HEADERS = ["번호","담당서포터즈","가게명","점주명","연락처","업종","동네","주소",
               "방문일정","TA진행상태","TA결과","컨설팅동의여부",
               "담당컨설턴트","전환상태","전환일시","비고","수정자","수정시각",
               "매장사진","동의서"];
var PHOTO_FIELDS = ["매장사진","동의서"];
var PHOTO_FOLDER_NAME = "동선_서포터즈_사진";
var ACCOUNT_HEADERS = ["이름","비번","권한"];

var TA_STATUSES = ["대기","방문확정","부재","재방문예정"];
var AGREES = ["미접촉","컨설팅동의","컨설팅거절","보류"];
var CONVERT_STATUSES = ["","전환완료"];

// 컨설턴트 트래커(수신 측) 헤더 — 컨설턴트 백엔드와 동일하게 유지할 것
var CONSULTANT_HEADERS = ["번호","담당컨설턴트","가게명","점주명","연락처","업종","동네","주소","출처서포터즈",
                          "월납보험료","컨설팅미팅1차","컨설팅미팅2_3차","클로징확률","계약현황",
                          "비고","수정자","수정시각"];

// ── 시트 접근 (이름 지정 = 기존 탭 무간섭) ─────────────────────
function ss_(){
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  if(!ss) throw new Error("스프레드시트를 열 수 없습니다: " + SPREADSHEET_ID);
  return ss;
}
function trackerSheet_(){
  var sh = ss_().getSheetByName(TRACKER_SHEET);
  if(!sh) throw new Error('"' + TRACKER_SHEET + '" 탭이 없습니다. setupTrackerSheet를 먼저 실행하세요.');
  return sh;
}
function accountSheet_(){
  var sh = ss_().getSheetByName(ACCOUNT_SHEET);
  if(!sh) throw new Error('"' + ACCOUNT_SHEET + '" 탭이 없습니다. setupAccountSheet를 먼저 실행하세요.');
  return sh;
}

// ── 최초 1회 실행: "트래커" 탭 생성 ────────────────────────────
// 이미 있으면 헤더만 확인하고 데이터는 절대 건드리지 않습니다.
function setupTrackerSheet(){
  var ss = ss_();
  var sh = ss.getSheetByName(TRACKER_SHEET);
  if(!sh){
    sh = ss.insertSheet(TRACKER_SHEET);
    Logger.log('"' + TRACKER_SHEET + '" 탭을 새로 만들었습니다.');
  } else {
    Logger.log('"' + TRACKER_SHEET + '" 탭이 이미 있습니다 — 헤더만 확인합니다(데이터 보존).');
  }
  sh.getRange(1,1,1,HEADERS.length).setValues([HEADERS]).setFontWeight("bold")
    .setBackground("#16335B").setFontColor("#FFFFFF");
  sh.setFrozenRows(1);
  // 연락처 열은 텍스트 서식 → 앞자리 0 보존
  var phoneCol = HEADERS.indexOf("연락처")+1;
  sh.getRange(2, phoneCol, Math.max(sh.getMaxRows()-1,1), 1).setNumberFormat("@");
  applyValidations_(sh);
  autoWidth_(sh, HEADERS.length);
  Logger.log("트래커 탭 헤더 세팅 완료 (" + HEADERS.length + "열). 데이터는 비어 있는 상태로 시작합니다.");
}

function applyValidations_(sh){
  var last = Math.max(sh.getMaxRows()-1, 1);
  var mk = function(list){ return SpreadsheetApp.newDataValidation().requireValueInList(list, true).build(); };
  sh.getRange(2, HEADERS.indexOf("TA진행상태")+1,   last, 1).setDataValidation(mk(TA_STATUSES));
  sh.getRange(2, HEADERS.indexOf("컨설팅동의여부")+1, last, 1).setDataValidation(mk(AGREES));
  sh.getRange(2, HEADERS.indexOf("전환상태")+1,     last, 1).setDataValidation(mk(CONVERT_STATUSES));
}
function autoWidth_(sh, n){
  try{ sh.autoResizeColumns(1, n); }catch(e){}
}

// ── 최초 1회 실행: "계정" 탭 생성 ──────────────────────────────
function setupAccountSheet(){
  var ss = ss_();
  var sh = ss.getSheetByName(ACCOUNT_SHEET);
  if(!sh){
    sh = ss.insertSheet(ACCOUNT_SHEET);
    Logger.log('"' + ACCOUNT_SHEET + '" 탭을 새로 만들었습니다.');
  } else {
    Logger.log('"' + ACCOUNT_SHEET + '" 탭이 이미 있습니다 — 헤더만 확인합니다(계정 보존).');
  }
  sh.getRange(1,1,1,ACCOUNT_HEADERS.length).setValues([ACCOUNT_HEADERS]).setFontWeight("bold")
    .setBackground("#16335B").setFontColor("#FFFFFF");
  sh.setFrozenRows(1);
  // 안내 주석
  sh.getRange(1,1).setNote(
    "서포터즈 로그인 계정 명단입니다.\n" +
    "A열=이름(트래커 탭의 '담당서포터즈' 값과 정확히 같아야 본인 담당 필터가 동작)\n" +
    "B열=비번(단순 문자열 대조 방식)\n" +
    "C열=권한 — 비워두면 일반 계정(본인 담당 건만 조회/수정), '관리자'라고 입력하면 전체 데이터 조회·편집 + 담당자 재배정 가능\n" +
    "2행부터 한 줄에 한 명씩 추가하세요. 이 탭은 웹앱 로그인 대조에만 쓰입니다.");
  sh.getRange(1,ACCOUNT_HEADERS.length+2).setValue(
    "← 2행부터 [이름 | 비번 | 권한(관리자만 입력, 비우면 일반)]을 입력하세요. 이름은 트래커 탭의 담당서포터즈와 동일하게.");
  autoWidth_(sh, ACCOUNT_HEADERS.length+3);
  Logger.log("계정 탭 세팅 완료. 시트에서 직접 이름/비번을 입력하세요.");
}

// ── 1회성 데이터 반영: (6)/(7) 원본 탭 → "트래커" 탭 정리 ──────────
// 2026-09-02: 장곡/장현/능곡/연성/군자 장곡동(122건) + 성남(60건) 원본 탭 데이터를
// "트래커" 탭에 정식 반영. "트래커" 탭 기존 데이터(번호 없이 부실하게 붙여넣기된 182건)를
// 지우고, 아래 두 원본 탭에서 다시 정확히 매핑해서 새로 씁니다.
// 실행: Apps Script 편집기에서 이 함수(importDongsunSeed_20260902)를 선택해 1회 실행.
// 실행 후 이 함수는 지워도 되고 남겨둬도 무해합니다(다시 실행해도 매번 "트래커" 탭을 같은
// 두 원본 탭 기준으로 다시 정리할 뿐, 원본 탭(17개 레거시 탭)은 절대 건드리지 않습니다).
function importDongsunSeed_20260902(){
  var ss = ss_();
  // 탭 이름의 공백·물결표 등 미세한 문자 차이에 영향받지 않도록, 정확히 같은 이름 대신
  // 각 탭을 확실히 구분해주는 핵심 문자열(부분 일치)로 찾습니다.
  var SOURCES = [
    { match: function(n){ return n.indexOf("성남") >= 0 && n.indexOf("137") >= 0; },
      label: "(7) 성남 137_2022~26 (부분일치: '성남'+'137')", dong: "성남" },
    { match: function(n){ return n.indexOf("122_2022") >= 0 && n.indexOf("장곡동") >= 0; },
      label: "(6)장곡/장현/능곡/연성/군자 장곡동 122_2022~26 (부분일치: '122_2022'+'장곡동')", dong: "시흥" }
  ];
  // 원본 'TA 결과' 텍스트 → 트래커 'TA진행상태'(선택형 4개 옵션) 매핑
  var TA_MAP = {
    "대기": "대기",
    "부재": "부재",
    "방문 확정": "재방문예정",
    "거절": "대기",
    "재연락필요": "대기",
    "보류": "대기"
  };

  var allSheets = ss.getSheets();
  var allNames = allSheets.map(function(s){ return s.getName(); });
  Logger.log("스프레드시트의 전체 탭 목록: " + JSON.stringify(allNames));

  var out = [];
  var no = 1;
  SOURCES.forEach(function(src){
    var sh = null;
    for(var i=0;i<allSheets.length;i++){
      if(src.match(allSheets[i].getName())){ sh = allSheets[i]; break; }
    }
    if(!sh){
      throw new Error('원본 탭을 찾을 수 없습니다: ' + src.label +
        ' — 실행 로그(보기 → 실행 기록/로그)에 찍힌 전체 탭 목록을 확인하세요.');
    }
    var values = sh.getDataRange().getValues();
    if(values.length < 2) return;
    var head = values[0].map(function(h){ return String(h).trim(); });
    var idx = function(key){ return head.indexOf(key); };
    var cTA진행자=idx("TA 진행자"), cTA결과=idx("TA 결과"), cBigo=idx("비고"),
        cBangmun=idx("방문일정"), cGage=idx("가게명"), cJeomju=idx("점주명"),
        cYeonrak=idx("점주 연락처"), cUpjong=idx("업종"), cJuso=idx("지도용주소"),
        cHwalseong=idx("활성상태"), cSeseDong=idx(src.dong==="시흥" ? "장곡동" : "__none__");

    for(var r=1; r<values.length; r++){
      var row = values[r];
      var store = String(cGage>=0 ? (row[cGage]||"") : "").trim();
      if(!store) continue;

      var taResult = String(cTA결과>=0 ? (row[cTA결과]||"") : "").trim();
      var taStatus = TA_MAP[taResult] || "대기";

      var bigoParts = [];
      var srcBigo = String(cBigo>=0 ? (row[cBigo]||"") : "").trim();
      if(srcBigo) bigoParts.push(srcBigo);
      if(cSeseDong>=0){
        var sese = String(row[cSeseDong]||"").trim();
        if(sese) bigoParts.push("[세부동: " + sese + "]");
      }
      if(cHwalseong>=0 && String(row[cHwalseong]||"").trim() === "비활성"){
        bigoParts.unshift("[비활성 매장]");
      }

      var visit = cBangmun>=0 ? row[cBangmun] : null;
      var visitStr = "";
      if(visit instanceof Date) visitStr = Utilities.formatDate(visit, "Asia/Seoul", "yyyy-MM-dd");
      else if(visit) visitStr = String(visit).trim();

      var phone = cYeonrak>=0 ? row[cYeonrak] : "";
      var phoneStr = "";
      if(phone !== null && phone !== undefined && phone !== ""){
        phoneStr = (typeof phone === "number") ? String(Math.round(phone)) : String(phone).trim();
      }

      out.push({
        "번호": no++,
        "담당서포터즈": String(cTA진행자>=0 ? (row[cTA진행자]||"") : "").trim(),
        "가게명": store,
        "점주명": String(cJeomju>=0 ? (row[cJeomju]||"") : "").trim(),
        "연락처": phoneStr,
        "업종": String(cUpjong>=0 ? (row[cUpjong]||"") : "").trim(),
        "동네": src.dong,
        "주소": String(cJuso>=0 ? (row[cJuso]||"") : "").trim(),
        "방문일정": visitStr,
        "TA진행상태": taStatus,
        "TA결과": taResult,
        "컨설팅동의여부": "",
        "담당컨설턴트": "",
        "전환상태": "",
        "전환일시": "",
        "비고": bigoParts.join(" · "),
        "수정자": "일괄반영",
        "수정시각": now_()
      });
    }
  });

  var trk = trackerSheet_();
  var liveHead = trk.getRange(1,1,1,trk.getLastColumn()).getValues()[0].map(function(h){ return String(h).trim(); });
  var width = liveHead.length;

  var matrix = out.map(function(o){
    var arr = [];
    for(var c=0;c<width;c++){ arr.push(liveHead[c] in o ? o[liveHead[c]] : ""); }
    return arr;
  });

  var lastRow = trk.getLastRow();
  if(lastRow >= 2){
    trk.getRange(2, 1, lastRow-1, width).clearContent();
  }
  if(matrix.length){
    trk.getRange(2, 1, matrix.length, width).setValues(matrix);
    var phoneCol = liveHead.indexOf("연락처")+1;
    if(phoneCol > 0) trk.getRange(2, phoneCol, matrix.length, 1).setNumberFormat("@");
  }
  Logger.log("반영 완료: " + matrix.length + "건 (성남·장곡동 원본 탭 기준으로 트래커 탭을 새로 정리했습니다)");
}

// ── 1회성 데이터 정리: TA진행상태 "방문완료" → "방문확정" ──────────
// 2026-09-03: 선택 옵션 이름을 "방문완료"에서 "방문확정"으로 바꾸면서, 이미 저장되어
// 있을 수 있는 예전 값도 같이 맞춰줍니다. "트래커" 탭 TA진행상태 열만 훑어서
// 정확히 "방문완료"인 셀만 "방문확정"으로 바꿉니다. 실행 후 지워도 무해합니다.
function fixTaStatusLabel_20260903(){
  var sh = trackerSheet_();
  var lastRow = sh.getLastRow();
  if(lastRow < 2){ Logger.log("데이터 없음"); return; }
  var head = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(function(h){ return String(h).trim(); });
  var col = head.indexOf("TA진행상태");
  if(col < 0){ Logger.log('"TA진행상태" 컬럼을 찾을 수 없습니다.'); return; }
  var range = sh.getRange(2, col+1, lastRow-1, 1);
  var values = range.getValues();
  var fixed = 0;
  for(var i=0;i<values.length;i++){
    if(String(values[i][0]).trim() === "방문완료"){
      values[i][0] = "방문확정";
      fixed++;
    }
  }
  if(fixed > 0) range.setValues(values);
  Logger.log("TA진행상태 '방문완료' → '방문확정' 변경: " + fixed + "건");
}

// ── 공통 유틸 ─────────────────────────────────────────────────
function json_(obj){
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
function sheetToObjects_(sh){
  var values = sh.getDataRange().getValues();
  if(values.length < 2) return [];
  var head = values.shift();
  return values.filter(function(r){
    return r.join("") !== "";
  }).map(function(r){
    var o = {};
    for(var i=0;i<head.length;i++){ o[String(head[i]).trim()] = r[i]; }
    return o;
  });
}
function ymd_(v){
  if(v instanceof Date) return Utilities.formatDate(v, "Asia/Seoul", "yyyy-MM-dd");
  return v;
}
function now_(){ return Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm"); }

// 계정 탭 대조 (Code.js handleManagerLogin 패턴)
// 반환: {name, isAdmin} 또는 null. isAdmin은 "계정" 탭 C열(권한)이 "관리자"일 때만 true.
function auth_(name, pw){
  name = String(name||"").trim();
  pw   = String(pw||"").trim();
  if(!name || !pw) return null;
  var rows = sheetToObjects_(accountSheet_());
  var me = null;
  for(var i=0;i<rows.length;i++){
    if(String(rows[i]["이름"]||"").trim() === name && String(rows[i]["비번"]||"").trim() === pw){ me = rows[i]; break; }
  }
  if(!me) return null;
  return {name:name, isAdmin: String(me["권한"]||"").trim() === "관리자"};
}

// 트래커 전체 행 (날짜는 문자열로 정규화)
function readTracker_(){
  var sh = trackerSheet_();
  var values = sh.getDataRange().getValues();
  if(values.length < 2) return [];
  var head = values.shift();
  var rows = [];
  for(var i=0;i<values.length;i++){
    var r = values[i];
    if(String(r[0]).trim() === "" && r.join("") === "") continue;
    var o = {};
    for(var c=0;c<head.length;c++){
      var key = String(head[c]).trim();
      var v = r[c];
      if(key === "방문일정") v = ymd_(v);
      if(key === "연락처")   v = String(v==null?"":v);
      o[key] = v;
    }
    if(String(o["번호"]).trim() === "") continue;
    rows.push(o);
  }
  return rows;
}

// 컨설턴트 시트의 "계정" 탭에서 컨설턴트 이름 목록을 읽어옴 (전환 팝업용)
// 실패하면 빈 배열 → 프론트 기본값 사용
function consultantNames_(){
  try{
    var css = SpreadsheetApp.openById(CONSULTANT_SPREADSHEET_ID);
    var sh = css.getSheetByName(ACCOUNT_SHEET);
    if(!sh) return [];
    var rows = sheetToObjects_(sh);
    var out = [];
    for(var i=0;i<rows.length;i++){
      if(String(rows[i]["권한"]||"").trim() === "관리자") continue; // 관리자 계정은 전환 대상 목록에서 제외
      var n = String(rows[i]["이름"]||"").trim();
      if(n && out.indexOf(n) < 0) out.push(n);
    }
    return out;
  }catch(e){ return []; }
}

// ── doGet / doPost ────────────────────────────────────────────
function doGet(e){
  return json_({ ok:true, service:"dongsun-supporter",
                 taStatuses:TA_STATUSES, agrees:AGREES, ts:new Date().getTime(),
                 hint:"로그인은 POST {action:'login', name, pw}" });
}

function doPost(e){
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try{
    var body = JSON.parse(e.postData.contents);
    var action = String(body.action||"").trim();

    if(action === "login")   return handleLogin_(body);
    if(action === "update")  return handleUpdate_(body);
    if(action === "convert") return handleConvert_(body);
    if(action === "photo")   return handlePhoto_(body);
    return json_({ok:false, error:"알 수 없는 요청: "+action});
  } catch(err){
    return json_({ok:false, error:String(err)});
  } finally {
    lock.releaseLock();
  }
}

// action:'login' → 계정 대조 후 트래커 행 반환.
// 관리자(isAdmin)는 전체 행, 일반 계정은 본인 담당 건만 서버에서 걸러서 내려줌(네트워크로도 남의 데이터 노출 안 함).
function handleLogin_(body){
  var auth = auth_(body.name, body.pw);
  if(!auth) return json_({ok:false, error:"이름 또는 비밀번호가 올바르지 않습니다"});
  var allRows = readTracker_();
  var rows = auth.isAdmin ? allRows : allRows.filter(function(r){
    return String(r["담당서포터즈"]||"").trim() === auth.name;
  });
  return json_({
    ok:true, name:auth.name, isAdmin:auth.isAdmin,
    taStatuses:TA_STATUSES, agrees:AGREES,
    consultants:consultantNames_(),
    rows:rows, ts:new Date().getTime()
  });
}

// 트래커 탭에서 번호로 행 찾기 → {sh, head, rowIdx(1-based)}
function findRow_(no){
  var sh = trackerSheet_();
  var data = sh.getDataRange().getValues();
  var head = data[0].map(function(h){ return String(h).trim(); });
  var colNo = head.indexOf("번호");
  for(var i=1;i<data.length;i++){
    if(String(data[i][colNo]).trim() === String(no).trim()){
      return {sh:sh, head:head, row:i+1, values:data[i]};
    }
  }
  return null;
}

// action:'update' → 필드 1개 갱신 + 수정자/수정시각 자동 기록
// 관리자는 담당(소유) 여부와 전환완료 잠금을 모두 무시하고 수정 가능 + "담당서포터즈" 재배정 가능
function handleUpdate_(body){
  var auth = auth_(body.name, body.pw);
  if(!auth) return json_({ok:false, error:"인증 실패 — 다시 로그인하세요"});

  var t = findRow_(body.no);
  if(!t) return json_({ok:false, error:"행을 찾을 수 없습니다: "+body.no});

  var field = String(body.field||"").trim();
  var editable = ["방문일정","TA진행상태","TA결과","컨설팅동의여부","비고"];
  var allowed = editable.slice();
  if(auth.isAdmin) allowed = allowed.concat(["담당서포터즈"]);
  if(allowed.indexOf(field) < 0) return json_({ok:false, error:"편집할 수 없는 항목입니다: "+field});

  if(!auth.isAdmin){
    var ownerIdx = t.head.indexOf("담당서포터즈");
    if(ownerIdx >= 0 && String(t.values[ownerIdx]).trim() !== auth.name){
      return json_({ok:false, error:"본인 담당 건만 수정할 수 있습니다"});
    }
    // 전환완료된 행은 잠금 (관리자는 예외)
    var csIdx = t.head.indexOf("전환상태");
    if(csIdx >= 0 && String(t.values[csIdx]).trim() === "전환완료"){
      return json_({ok:false, error:"이미 컨설턴트로 전환된 건이라 수정할 수 없습니다"});
    }
  }

  var col = t.head.indexOf(field);
  if(col < 0) return json_({ok:false, error:"컬럼이 없습니다: "+field});
  t.sh.getRange(t.row, col+1).setValue(body.value);

  stamp_(t.sh, t.head, t.row, auth.name);
  return json_({ok:true, no:body.no, field:field, value:body.value});
}

function stamp_(sh, head, row, name){
  var uc = head.indexOf("수정자");   if(uc >= 0) sh.getRange(row, uc+1).setValue(name);
  var tc = head.indexOf("수정시각"); if(tc >= 0) sh.getRange(row, tc+1).setValue(now_());
}

// action:'convert' → 이 시트 행을 전환완료 처리 + 컨설턴트 시트 "트래커" 탭에 새 행 추가
function handleConvert_(body){
  var auth = auth_(body.name, body.pw);
  if(!auth) return json_({ok:false, error:"인증 실패 — 다시 로그인하세요"});
  if(!auth.isAdmin){
    return json_({ok:false, error:"컨설턴트 전환은 관리자만 할 수 있습니다"});
  }
  var name = auth.name;

  var consultant = String(body.consultant||"").trim();
  if(!consultant) return json_({ok:false, error:"담당 컨설턴트를 선택하세요"});

  var t = findRow_(body.no);
  if(!t) return json_({ok:false, error:"행을 찾을 수 없습니다: "+body.no});

  var g = function(k){ var i = t.head.indexOf(k); return i>=0 ? t.values[i] : ""; };
  if(String(g("전환상태")).trim() === "전환완료"){
    return json_({ok:false, error:"이미 전환된 건입니다"});
  }

  // 1) 컨설턴트 스프레드시트 "트래커" 탭에 새 행 추가
  var css = SpreadsheetApp.openById(CONSULTANT_SPREADSHEET_ID);
  var csh = css.getSheetByName(CONSULTANT_TRACKER_SHEET);
  if(!csh){
    return json_({ok:false, error:'컨설팅DB 스프레드시트에 "' + CONSULTANT_TRACKER_SHEET +
                  '" 탭이 없습니다. 컨설턴트용 스크립트의 setupTrackerSheet를 먼저 실행하세요.'});
  }
  var chead = csh.getRange(1,1,1,csh.getLastColumn()).getValues()[0].map(function(h){ return String(h).trim(); });
  if(chead.indexOf("번호") < 0){
    return json_({ok:false, error:'컨설팅DB "' + CONSULTANT_TRACKER_SHEET + '" 탭 헤더가 비어 있습니다.'});
  }
  // 다음 번호
  var lastRow = csh.getLastRow();
  var nextNo = 1;
  if(lastRow >= 2){
    var nos = csh.getRange(2, chead.indexOf("번호")+1, lastRow-1, 1).getValues();
    for(var i=0;i<nos.length;i++){
      var v = parseInt(nos[i][0], 10);
      if(!isNaN(v) && v >= nextNo) nextNo = v + 1;
    }
  }
  var map = {
    "번호": nextNo,
    "담당컨설턴트": consultant,
    "가게명": g("가게명"),
    "점주명": g("점주명"),
    "연락처": String(g("연락처")||""),
    "업종": g("업종"),
    "동네": g("동네"),
    "주소": g("주소"),
    "출처서포터즈": g("담당서포터즈") || name,
    "월납보험료": "",
    "컨설팅미팅1차": "",
    "컨설팅미팅2_3차": "",
    "클로징확률": "",
    "계약현황": "신규배정",
    "비고": "",
    "수정자": name,
    "수정시각": now_()
  };
  var newRow = chead.map(function(h){ return (h in map) ? map[h] : ""; });
  var target = lastRow + 1;
  csh.getRange(target, 1, 1, chead.length).setValues([newRow]);
  var cphone = chead.indexOf("연락처");
  if(cphone >= 0) csh.getRange(target, cphone+1).setNumberFormat("@").setValue(String(g("연락처")||""));

  // 2) 이 시트(서포터즈 트래커) 행을 전환완료로 갱신
  var setIf = function(k, v){ var i = t.head.indexOf(k); if(i>=0) t.sh.getRange(t.row, i+1).setValue(v); };
  setIf("담당컨설턴트", consultant);
  setIf("전환상태", "전환완료");
  setIf("전환일시", now_());
  setIf("컨설팅동의여부", "컨설팅동의");
  stamp_(t.sh, t.head, t.row, name);

  return json_({ok:true, no:body.no, consultant:consultant, consultantNo:nextNo});
}

// ── 사진 업로드(매장사진·동의서) — 구글드라이브에 저장 후 트래커 셀에는 URL만 기록 ──
function photoFolder_(){
  var props = PropertiesService.getScriptProperties();
  var fid = props.getProperty('PHOTO_FOLDER_ID');
  if(fid){
    try{ return DriveApp.getFolderById(fid); }catch(e){ /* 폴더가 삭제됐으면 재생성 */ }
  }
  var it = DriveApp.getFoldersByName(PHOTO_FOLDER_NAME);
  var folder = it.hasNext() ? it.next() : DriveApp.createFolder(PHOTO_FOLDER_NAME);
  props.setProperty('PHOTO_FOLDER_ID', folder.getId());
  return folder;
}

// action:'photo' → base64 이미지를 드라이브에 저장하고 링크를 트래커 셀에 기록
function handlePhoto_(body){
  var auth = auth_(body.name, body.pw);
  if(!auth) return json_({ok:false, error:"인증 실패 — 다시 로그인하세요"});
  var name = auth.name;

  var field = String(body.field||"").trim();
  if(PHOTO_FIELDS.indexOf(field) < 0) return json_({ok:false, error:"사진 항목이 아닙니다: "+field});

  var t = findRow_(body.no);
  if(!t) return json_({ok:false, error:"행을 찾을 수 없습니다: "+body.no});

  if(!auth.isAdmin){
    var ownerIdx = t.head.indexOf("담당서포터즈");
    if(ownerIdx >= 0 && String(t.values[ownerIdx]).trim() !== auth.name){
      return json_({ok:false, error:"본인 담당 건만 사진을 등록할 수 있습니다"});
    }
    var csIdx = t.head.indexOf("전환상태");
    if(csIdx >= 0 && String(t.values[csIdx]).trim() === "전환완료"){
      return json_({ok:false, error:"이미 컨설턴트로 전환된 건이라 사진을 바꿀 수 없습니다"});
    }
  }

  var b64 = String(body.data||"");
  if(!b64) return json_({ok:false, error:"사진 데이터가 없습니다"});
  if(b64.length > 8000000) return json_({ok:false, error:"사진 용량이 너무 큽니다 — 다시 촬영해보세요"});

  var mime = String(body.mime||"image/jpeg");
  var storeName = String(t.values[t.head.indexOf("가게명")]||"").trim() || "매장";
  var stamp = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyyMMdd_HHmmss");
  var fname = (String(body.no)+"_"+storeName+"_"+field+"_"+stamp+".jpg").replace(/[\\\/:*?"<>|]/g, "_");

  var blob = Utilities.newBlob(Utilities.base64Decode(b64), mime, fname);
  var folder = photoFolder_();
  var file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  var url = "https://drive.google.com/uc?export=view&id=" + file.getId();

  var col = t.head.indexOf(field);
  if(col < 0) return json_({ok:false, error:'"트래커" 탭에 "'+field+'" 컬럼이 없습니다. setupTrackerSheet를 다시 실행하세요.'});
  t.sh.getRange(t.row, col+1).setValue(url);
  stamp_(t.sh, t.head, t.row, name);

  return json_({ok:true, no:body.no, field:field, url:url});
}
