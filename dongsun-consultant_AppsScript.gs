/**
 * 동선 컨설턴트 트래커 — Google Apps Script API (구글시트 백엔드)
 *
 * ⚠ 매우 중요 — 기존 탭 보호
 *   이 스크립트는 "동선_컨설팅 DB 관리용" 스프레드시트의 기존 탭("시흥" 등 자유서식 원본)을
 *   절대 읽거나 쓰지 않습니다. 오직 아래 두 개의 신규 탭만 사용합니다.
 *     · "트래커" — 컨설팅 진행 데이터 (구조화 컬럼)
 *     · "계정"   — 로그인용 이름/비번
 *   시트 이름으로만 명시 접근하므로 다른 탭은 건드리지 않습니다.
 *
 * 사용법:
 *  1) 아래 SPREADSHEET_ID 확인 (이미 채워져 있음)
 *  2) 함수 목록에서 setupTrackerSheet 실행 → "트래커" 탭 생성(빈 상태, 헤더만)
 *  3) 함수 목록에서 setupAccountSheet 실행 → "계정" 탭 생성 후 시트에서 직접 이름/비번 입력
 *     ※ 여기에 넣은 이름이 서포터즈 트래커의 "컨설턴트 전환" 팝업 목록이 됩니다.
 *  4) 배포 → 새 배포 → 유형: 웹 앱 → 실행: 나 / 액세스: 모든 사용자 → 배포
 *  5) 나온 웹앱 URL(.../exec)을 dongsun-consultant.html 최초 접속 화면에 입력
 *
 * ※ 과거 데이터 자동 이관은 하지 않습니다. 필요하면 사용자가 "트래커" 탭에 직접 복사해 넣으세요.
 * ※ 새 행은 서포터즈 트래커의 action:'convert' 가 자동으로 추가합니다.
 */

// ── 스프레드시트 ID ───────────────────────────────────────────
// 동선_컨설팅 DB 관리용
var SPREADSHEET_ID = "1AtJ_qLMzsyRCuSAH-cGNCEhribVCt1qBY0a5GUzc9SI";

// ── 신규 탭 이름 (기존 "시흥" 등과 절대 겹치지 않게) ────────────
var TRACKER_SHEET = "트래커";
var ACCOUNT_SHEET = "계정";

// 서포터즈 백엔드(dongsun-supporter_AppsScript.gs)의 CONSULTANT_HEADERS와 동일하게 유지할 것
var HEADERS = ["번호","담당컨설턴트","가게명","점주명","연락처","업종","동네","주소","출처서포터즈",
               "월납보험료","컨설팅미팅1차","컨설팅미팅2_3차","클로징확률","계약현황",
               "비고","수정자","수정시각"];
var ACCOUNT_HEADERS = ["이름","비번","권한"];

var STATUSES = ["신규배정","상담중","청약완료","계약체결","종결·실패"];
var PROBS = ["10%","30%","50%","70%","90%","100%"];

// ── 시트 접근 (이름 지정 = 기존 탭 무간섭) ─────────────────────
var _ssCache_ = null;
function ss_(){
  if(!_ssCache_){
    _ssCache_ = SpreadsheetApp.openById(SPREADSHEET_ID);
    if(!_ssCache_) throw new Error("스프레드시트를 열 수 없습니다: " + SPREADSHEET_ID);
  }
  return _ssCache_;
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
  sh.getRange(2, HEADERS.indexOf("계약현황")+1,  last, 1).setDataValidation(mk(STATUSES));
  sh.getRange(2, HEADERS.indexOf("클로징확률")+1, last, 1).setDataValidation(mk(PROBS));
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
  sh.getRange(1,1).setNote(
    "컨설턴트 로그인 계정 명단입니다.\n" +
    "A열=이름(트래커 탭의 '담당컨설턴트' 값과 정확히 같아야 본인 담당 필터가 동작)\n" +
    "B열=비번(단순 문자열 대조 방식)\n" +
    "C열=권한 — 비워두면 일반 계정(본인 담당 건만 조회/수정), '관리자'라고 입력하면 전체 데이터 조회·편집 + 담당자 재배정 가능\n" +
    "2행부터 한 줄에 한 명씩 추가하세요.\n" +
    "※ '관리자' 계정은 서포터즈 트래커의 '컨설턴트 전환' 선택 목록에서 자동 제외됩니다.\n" +
    "※ 이 목록이 서포터즈 트래커의 '컨설턴트 전환' 선택 목록으로도 쓰입니다.");
  sh.getRange(1,ACCOUNT_HEADERS.length+2).setValue(
    "← 2행부터 [이름 | 비번 | 권한(관리자만 입력, 비우면 일반)]을 입력하세요. 예) 김광연 / 정현우 / 하명규");
  autoWidth_(sh, ACCOUNT_HEADERS.length+3);
  Logger.log("계정 탭 세팅 완료. 시트에서 직접 이름/비번을 입력하세요.");
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

function readTracker_(){
  var sh = trackerSheet_();
  var values = sh.getDataRange().getValues();
  if(values.length < 2) return [];
  var head = values.shift();
  var rows = [];
  for(var i=0;i<values.length;i++){
    var r = values[i];
    if(r.join("") === "") continue;
    var o = {};
    for(var c=0;c<head.length;c++){
      var key = String(head[c]).trim();
      var v = r[c];
      if(key === "연락처") v = String(v==null?"":v);
      o[key] = v;
    }
    if(String(o["번호"]).trim() === "") continue;
    rows.push(o);
  }
  return rows;
}

// ── doGet / doPost ────────────────────────────────────────────
function doGet(e){
  return json_({ ok:true, service:"dongsun-consultant",
                 statuses:STATUSES, probs:PROBS, ts:new Date().getTime(),
                 hint:"로그인은 POST {action:'login', name, pw}" });
}

function doPost(e){
  try{
    var body = JSON.parse(e.postData.contents);
    var action = String(body.action||"").trim();

    // 읽기 전용(login)은 락 없이 처리 — 새로고침/자동폴링이 쓰기 작업과 서로 줄서서 기다리지 않도록 함
    if(action === "login") return handleLogin_(body);

    var lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try{
      if(action === "update") return handleUpdate_(body);
      return json_({ok:false, error:"알 수 없는 요청: "+action});
    } finally {
      lock.releaseLock();
    }
  } catch(err){
    return json_({ok:false, error:String(err)});
  }
}

// action:'login' → 계정 대조 후 트래커 행 반환.
// 관리자(isAdmin)는 전체 행, 일반 계정은 본인 담당 건만 서버에서 걸러서 내려줌(네트워크로도 남의 데이터 노출 안 함).
function handleLogin_(body){
  var auth = auth_(body.name, body.pw);
  if(!auth) return json_({ok:false, error:"이름 또는 비밀번호가 올바르지 않습니다"});
  var allRows = readTracker_();
  var rows = auth.isAdmin ? allRows : allRows.filter(function(r){
    return String(r["담당컨설턴트"]||"").trim() === auth.name;
  });
  return json_({
    ok:true, name:auth.name, isAdmin:auth.isAdmin,
    statuses:STATUSES, probs:PROBS,
    rows:rows, ts:new Date().getTime()
  });
}

// 트래커 탭에서 번호로 행 찾기
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
// 관리자는 담당(소유) 여부를 무시하고 수정 가능 + "담당컨설턴트" 재배정 가능
function handleUpdate_(body){
  var auth = auth_(body.name, body.pw);
  if(!auth) return json_({ok:false, error:"인증 실패 — 다시 로그인하세요"});

  var t = findRow_(body.no);
  if(!t) return json_({ok:false, error:"행을 찾을 수 없습니다: "+body.no});

  var field = String(body.field||"").trim();
  var editable = ["월납보험료","컨설팅미팅1차","컨설팅미팅2_3차","클로징확률","계약현황","비고"];
  var allowed = editable.slice();
  if(auth.isAdmin) allowed = allowed.concat(["담당컨설턴트"]);
  if(allowed.indexOf(field) < 0) return json_({ok:false, error:"편집할 수 없는 항목입니다: "+field});

  if(!auth.isAdmin){
    var ownerIdx = t.head.indexOf("담당컨설턴트");
    if(ownerIdx >= 0 && String(t.values[ownerIdx]).trim() !== auth.name){
      return json_({ok:false, error:"본인 담당 건만 수정할 수 있습니다"});
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
