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
var SPREADSHEET_ID = "1J_otYQ_gMwVskqUHoCLdJIbkTytjF13P4t1044YN2jI";
// 서포터즈 원본 스프레드시트(읽기 전용 참조 — 위도/경도 소급 채우기용, 2026-09-07(5차) 추가)
var SUPPORTER_SPREADSHEET_ID = "1ewvEx1GdEzhVsIdbymxevSGbdusimLNcW0PQEvqQamw";

// ── 신규 탭 이름 (기존 "시흥" 등과 절대 겹치지 않게) ────────────
var TRACKER_SHEET = "트래커";
var ACCOUNT_SHEET = "계정";

// 서포터즈 백엔드(dongsun-supporter_AppsScript.gs)의 CONSULTANT_HEADERS와 동일하게 유지할 것
var HEADERS = ["번호","담당컨설턴트","가게명","점주명","연락처","업종","동네","주소","출처서포터즈",
               "월납보험료","컨설팅미팅1차","컨설팅미팅2_3차","클로징확률","계약현황",
               "비고","수정자","수정시각",
               // 2026-09-07 추가: 기존 17개 컬럼 뒤에만 추가(순서·기존 컬럼 불변) — 실제 시트에 반영하려면 setupTrackerSheet 재실행
               "청약여부","청약차수","청약금액","청약상품","종결사유",
               "AS대상","AS신청기한","AS신청상태","증빙","AS점검메모","DB지급일",
               // 2026-09-07(3차) 추가: 서포터즈의 동의서 등록 전 사전체크 항목 — 전환 시점에 1회 복사되어 옴(읽기전용), setupTrackerSheet 재실행 필요.
               // 서포터즈 쪽 PRECHECK_FIELDS와 컬럼명을 맞춰야 함(항목이 늘면 여기도 같이 추가)
               "기존설계사관계","월납보험료수준","연령대","성별","3대질환진단여부",
               // 2026-09-07(5차) 추가: 서포터즈에서 지오코딩한 좌표("가까운 순" 정렬용) — 전환 시점에 1회 복사(읽기전용), 기존 행은 backfillLatLngFromSupporter_20260907()으로 소급
               "위도","경도",
               "팀배정일"]; // 2026-09-08 추가: 영업관리자가 담당컨설턴트를 배정/재배정한 시점(비파괴, 맨 뒤) — handleUpdate_가 담당컨설턴트 변경 시 자동 기록
var ACCOUNT_HEADERS = ["이름","비번","권한","소속영업관리자"]; // 2026-09-08 추가: 영업관리자 계층 도입 — 컨설턴트 행에만 소속 영업관리자 이름을 채움(비파괴, 맨 뒤)

var STATUSES = ["신규배정","상담중","청약완료","계약체결","종결·실패"];
var PROBS = ["10%","30%","50%","70%","90%","100%"];

// 종결사유 드롭다운(프론트 CLOSE_REASONS와 동일하게 유지)
var CLOSE_REASONS = ["장기 부재","연령 기준 외","설계 불가(중대질환)","외국인","직계가족 설계사","단박 거절","보험료 0원/완납","3개월 내 보험가입","폐업","대표 아님(명의 불일치)","지인 설계사","갱신형 유지","관심 없음","가족 반대","타 GA 이관","장기 보류","기타"];
// A/S 인정 유형(프론트 AS_TYPES와 동일하게 유지) — _applyAsLogic/handleSubmitAs_에서 사용. 기존엔 정의가 누락돼 있었음(2026-09-07 수정).
var AS_TYPES = ["장기 부재","연령 기준 외","설계 불가(중대질환)","외국인","직계가족 설계사","단박 거절","보험료 0원/완납","3개월 내 보험가입","폐업","대표 아님(명의 불일치)"];
var AS_APPLY_STATUSES = ["미신청","확인중","승인","불가능"];
var AS_DEADLINE_DAYS = 9; // A/S 신청기한 = DB지급일 + 9일(지급일 포함 10일). 기존엔 정의가 누락돼 있었음(2026-09-07 수정).

// 증빙 파일 업로드(A/S 증빙) — 서포터즈 트래커의 사진 업로드 패턴과 동일(구글드라이브 저장 + URL을 "|"로 이어붙여 저장)
var PHOTO_FIELDS = ["증빙"];
var PHOTO_MAX = 5;
var PHOTO_FOLDER_NAME = "동선_컨설턴트_증빙";

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
  // 2026-09-07 추가: DB지급일·AS신청기한은 "yyyy-MM-dd" 문자열로 써넣는데, 서식이 "자동"이면
  // 구글시트가 이를 진짜 날짜 셀로 자동 변환해버려서 API가 다시 읽을 때 Date 객체(→ISO 문자열)로
  // 튀어나오는 문제가 있었음 — 텍스트 서식으로 고정해 항상 문자열로 저장·조회되게 함
  ["DB지급일","AS신청기한","팀배정일"].forEach(function(colName){
    var idx = HEADERS.indexOf(colName)+1;
    if(idx > 0) sh.getRange(2, idx, Math.max(sh.getMaxRows()-1,1), 1).setNumberFormat("@");
  });
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
    "C열=권한 — 비워두면 컨설턴트(본인 담당 건만 조회/수정), '영업관리자'라고 입력하면 본인+소속 컨설턴트(D열) 데이터를 조회하고 담당컨설턴트를 배정/재배정할 수 있음, '관리자'라고 입력하면 전체관리자(전체 데이터 조회·편집)\n" +
    "D열=소속영업관리자 — 컨설턴트 계정에만 입력(그 컨설턴트가 속한 영업관리자 이름). 전체관리자·영업관리자 행은 비워둠.\n" +
    "  예) 황태성/박지우=권한 '관리자', 정현우=권한 '영업관리자', 김광연·하명규=권한 비움(컨설턴트) + D열에 '정현우'\n" +
    "2행부터 한 줄에 한 명씩 추가하세요.\n" +
    "※ '관리자' 계정은 서포터즈 트래커의 '컨설턴트 전환' 선택 목록에서 자동 제외됩니다.\n" +
    "※ 이 목록이 서포터즈 트래커의 '컨설턴트 전환' 선택 목록으로도 쓰입니다.\n" +
    "※ 2026-09-08부터 컨설턴트 전환 시 담당컨설턴트는 기본 미배정으로 생성되고, 영업관리자가 컨설턴트 트래커에서 배정합니다.");
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
  var role = roleOf_(me["권한"]);
  return {name:name, isAdmin: role === "관리자", role: role};
}

// 2026-09-08 추가 — 영업관리자 계층: "권한" 컬럼 문자열을 3단 역할로 정규화
function roleOf_(perm){
  perm = String(perm||"").trim();
  if(perm === "관리자") return "관리자";
  if(perm === "영업관리자") return "영업관리자";
  return "컨설턴트";
}

// 2026-09-08 추가 — 영업관리자(managerName) 본인 + 소속 컨설턴트(계정 탭 "소속영업관리자"=managerName) 이름 목록
function teamMembersOf_(managerName){
  var team = [managerName];
  var rows = sheetToObjects_(accountSheet_());
  for(var i=0;i<rows.length;i++){
    if(String(rows[i]["소속영업관리자"]||"").trim() === managerName){
      var n = String(rows[i]["이름"]||"").trim();
      if(n && team.indexOf(n) < 0) team.push(n);
    }
  }
  return team;
}

// 2026-09-08 추가 — 계정 탭 전체 이름 목록(전체관리자의 "담당자 전체/개별" 필터 드롭다운용).
// 서포터즈 트래커의 supporterNames_()와 동일한 목적 — 전체관리자만 이 목록이 필요해서 handleLogin_에서 isAdmin일 때만 내려줌.
function allAccountNames_(){
  var rows = sheetToObjects_(accountSheet_());
  var out = [];
  for(var i=0;i<rows.length;i++){
    var n = String(rows[i]["이름"]||"").trim();
    if(n && out.indexOf(n) < 0) out.push(n);
  }
  return out;
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
                 statuses:STATUSES, probs:PROBS,
                 closeReasons:CLOSE_REASONS, asTypes:AS_TYPES, applyStatuses:AS_APPLY_STATUSES,
                 ts:new Date().getTime(),
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
      if(action === "update")      return handleUpdate_(body);
      if(action === "submitAs")    return handleSubmitAs_(body);
      if(action === "asDecide")    return handleAsDecide_(body);
      if(action === "photo")       return handlePhoto_(body);
      if(action === "deletePhoto") return handleDeletePhoto_(body);
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
  var team = null, rows;
  if(auth.isAdmin){
    rows = allRows;
  } else if(auth.role === "영업관리자"){
    // 2026-09-08 추가: 영업관리자는 본인 팀 담당 건 + 아직 담당컨설턴트가 비어있는(배정 대기) 건을 봄
    team = teamMembersOf_(auth.name);
    rows = allRows.filter(function(r){
      var owner = String(r["담당컨설턴트"]||"").trim();
      return !owner || team.indexOf(owner) >= 0;
    });
  } else {
    rows = allRows.filter(function(r){
      return String(r["담당컨설턴트"]||"").trim() === auth.name;
    });
  }
  var roster = auth.isAdmin ? allAccountNames_() : null; // 2026-09-08 추가: 전체관리자의 담당자 필터 드롭다운용
  return json_({
    ok:true, name:auth.name, isAdmin:auth.isAdmin, role:auth.role, team:team, roster:roster,
    statuses:STATUSES, probs:PROBS,
    closeReasons:CLOSE_REASONS, asTypes:AS_TYPES, applyStatuses:AS_APPLY_STATUSES,
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
  var editable = ["월납보험료","컨설팅미팅1차","컨설팅미팅2_3차","클로징확률","계약현황",
                  "청약여부","청약차수","청약금액","청약상품","종결사유","비고"];
  // 증빙은 action:'photo'/'deletePhoto'로만 관리(자유텍스트 update로 덮어쓰지 않음)
  var allowed = editable.slice();
  if(auth.isAdmin) allowed = allowed.concat(["담당컨설턴트","AS신청상태"]);
  else if(auth.role === "영업관리자") allowed = allowed.concat(["담당컨설턴트"]); // 2026-09-08 추가: 영업관리자는 배정/재배정만 가능(AS신청상태는 전체관리자 전용 유지)
  if(allowed.indexOf(field) < 0) return json_({ok:false, error:"편집할 수 없는 항목입니다: "+field});

  var ownerIdx = t.head.indexOf("담당컨설턴트");
  var currentOwner = ownerIdx >= 0 ? String(t.values[ownerIdx]).trim() : "";

  if(field === "담당컨설턴트" && auth.role === "영업관리자"){
    // 2026-09-08 추가: 영업관리자는 본인 팀(본인+소속 컨설턴트) 범위 안에서만 배정/재배정 가능
    var team = teamMembersOf_(auth.name);
    if(currentOwner && team.indexOf(currentOwner) < 0){
      return json_({ok:false, error:"본인 팀 담당 건만 재배정할 수 있습니다"});
    }
    var newOwner = String(body.value||"").trim();
    if(newOwner && team.indexOf(newOwner) < 0){
      return json_({ok:false, error:"본인 팀 소속 컨설턴트에게만 배정할 수 있습니다"});
    }
  } else if(!auth.isAdmin){
    if(ownerIdx >= 0 && currentOwner !== auth.name){
      return json_({ok:false, error:"본인 담당 건만 수정할 수 있습니다"});
    }
  }

  var col = t.head.indexOf(field);
  if(col < 0) return json_({ok:false, error:"컬럼이 없습니다: "+field});
  t.sh.getRange(t.row, col+1).setValue(body.value);
  t.values[col] = body.value;

  if(field === "종결사유") _applyAsLogic(t);
  if(field === "담당컨설턴트"){
    // 2026-09-08 추가: 담당컨설턴트가 배정/재배정/해제될 때마다 팀배정일 자동 기록(공란이 되면 같이 비움)
    var teamDateCol = t.head.indexOf("팀배정일");
    if(teamDateCol >= 0){
      var newOwnerVal = String(body.value||"").trim();
      var stampVal = newOwnerVal ? Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd") : "";
      t.sh.getRange(t.row, teamDateCol+1).setValue(stampVal);
      t.values[teamDateCol] = stampVal;
    }
  }

  stamp_(t.sh, t.head, t.row, auth.name);
  return json_({ok:true, no:body.no, field:field, value:body.value});
}

// 종결사유 변경 시 A/S 대상·신청기한·초기 신청상태를 서버에서 자동 계산
function _applyAsLogic(t){
  var g = function(k){ var i=t.head.indexOf(k); return i>=0 ? t.values[i] : ""; };
  var setIf = function(k, v){ var i=t.head.indexOf(k); if(i>=0){ t.sh.getRange(t.row, i+1).setValue(v); t.values[i]=v; } };
  var reason = String(g("종결사유")||"").trim();
  if(!reason){ setIf("AS대상",""); setIf("AS신청기한",""); setIf("AS신청상태",""); return; }

  var isAs = AS_TYPES.indexOf(reason) >= 0;
  setIf("AS대상", isAs ? "A/S 대상" : "영업 사유");
  if(!isAs){ setIf("AS신청기한",""); setIf("AS신청상태",""); return; }

  var payDate = g("DB지급일");
  if(payDate){
    var d = (payDate instanceof Date) ? payDate : new Date(payDate);
    if(!isNaN(d.getTime())){
      var due = new Date(d.getTime());
      due.setDate(due.getDate() + AS_DEADLINE_DAYS);
      setIf("AS신청기한", Utilities.formatDate(due, "Asia/Seoul", "yyyy-MM-dd"));
    }
  }
  var curStatus = String(g("AS신청상태")||"").trim();
  if(!curStatus) setIf("AS신청상태","미신청");
}

// action:'submitAs' → 컨설턴트가 A/S 신청 (종결사유가 A/S 인정 유형 + 증빙 입력 후)
function handleSubmitAs_(body){
  var auth = auth_(body.name, body.pw);
  if(!auth) return json_({ok:false, error:"인증 실패 — 다시 로그인하세요"});

  var t = findRow_(body.no);
  if(!t) return json_({ok:false, error:"행을 찾을 수 없습니다: "+body.no});

  if(!auth.isAdmin){
    var ownerIdx = t.head.indexOf("담당컨설턴트");
    if(ownerIdx >= 0 && String(t.values[ownerIdx]).trim() !== auth.name){
      return json_({ok:false, error:"본인 담당 건만 신청할 수 있습니다"});
    }
  }

  var g = function(k){ var i=t.head.indexOf(k); return i>=0 ? t.values[i] : ""; };
  var reason = String(g("종결사유")||"").trim();
  if(AS_TYPES.indexOf(reason) < 0) return json_({ok:false, error:"종결사유가 A/S 인정 유형이 아닙니다"});
  var proof = String(g("증빙")||"").trim();
  if(!proof) return json_({ok:false, error:"증빙 내용을 먼저 입력하세요"});
  var cur = String(g("AS신청상태")||"").trim();
  if(cur === "확인중" || cur === "승인") return json_({ok:false, error:"이미 신청됐습니다 (상태: "+cur+")"});

  var setIf = function(k, v){ var i=t.head.indexOf(k); if(i>=0) t.sh.getRange(t.row, i+1).setValue(v); };
  setIf("AS신청상태", "확인중");
  stamp_(t.sh, t.head, t.row, auth.name);
  return json_({ok:true, no:body.no, status:"확인중"});
}

// action:'asDecide' → 관리자 전용, A/S 승인·불가능 처리(+점검메모 누적)
function handleAsDecide_(body){
  var auth = auth_(body.name, body.pw);
  if(!auth) return json_({ok:false, error:"인증 실패 — 다시 로그인하세요"});
  if(!auth.isAdmin) return json_({ok:false, error:"A/S 승인·불가 처리는 관리자만 할 수 있습니다"});

  var t = findRow_(body.no);
  if(!t) return json_({ok:false, error:"행을 찾을 수 없습니다: "+body.no});

  var decision = String(body.decision||"").trim();
  if(["승인","불가능"].indexOf(decision) < 0) return json_({ok:false, error:"승인 또는 불가능만 처리할 수 있습니다"});

  var setIf = function(k, v){ var i=t.head.indexOf(k); if(i>=0) t.sh.getRange(t.row, i+1).setValue(v); };
  setIf("AS신청상태", decision);
  var memo = String(body.memo||"").trim();
  var memoLine = "["+now_()+" "+auth.name+"] "+decision+(memo?(" — "+memo):"");
  var mi = t.head.indexOf("AS점검메모");
  if(mi >= 0){
    var old = String(t.values[mi]||"").trim();
    t.sh.getRange(t.row, mi+1).setValue(old ? (old+"\n"+memoLine) : memoLine);
  }
  stamp_(t.sh, t.head, t.row, auth.name);
  return json_({ok:true, no:body.no, status:decision});
}

// ── 증빙 파일 업로드(A/S 증빙) — 구글드라이브에 저장 후 트래커 셀에는 URL만 기록(서포터즈 트래커와 동일 패턴) ──
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

// action:'photo' → base64 이미지를 드라이브에 저장하고 링크를 트래커 셀(증빙)에 기록
function handlePhoto_(body){
  var auth = auth_(body.name, body.pw);
  if(!auth) return json_({ok:false, error:"인증 실패 — 다시 로그인하세요"});
  var name = auth.name;

  var field = String(body.field||"").trim();
  if(PHOTO_FIELDS.indexOf(field) < 0) return json_({ok:false, error:"사진 항목이 아닙니다: "+field});

  var t = findRow_(body.no);
  if(!t) return json_({ok:false, error:"행을 찾을 수 없습니다: "+body.no});

  if(!auth.isAdmin){
    var ownerIdx = t.head.indexOf("담당컨설턴트");
    if(ownerIdx >= 0 && String(t.values[ownerIdx]).trim() !== auth.name){
      return json_({ok:false, error:"본인 담당 건만 증빙을 등록할 수 있습니다"});
    }
    var asIdx = t.head.indexOf("AS신청상태");
    var asSt = asIdx >= 0 ? String(t.values[asIdx]).trim() : "";
    if(asSt === "승인" || asSt === "불가능"){
      return json_({ok:false, error:"이미 처리된 A/S 건이라 증빙을 바꿀 수 없습니다"});
    }
  }

  // 드라이브 업로드 전에 저장할 컬럼이 실제로 있는지 먼저 확인 (없으면 파일만 올리고 실패하는 것을 방지)
  var col = t.head.indexOf(field);
  if(col < 0) return json_({ok:false, error:'"트래커" 탭에 "'+field+'" 컬럼이 없습니다. setupTrackerSheet를 다시 실행하세요.'});

  // 기존 저장값은 "url1|url2|..." 형태(파이프 구분)
  var existing = String(t.values[col]||"").trim();
  var urls = existing ? existing.split("|").map(function(s){ return s.trim(); }).filter(Boolean) : [];
  if(urls.length >= PHOTO_MAX){
    return json_({ok:false, error:"증빙은 최대 "+PHOTO_MAX+"장까지만 등록할 수 있습니다"});
  }

  var b64 = String(body.data||"");
  if(!b64) return json_({ok:false, error:"사진 데이터가 없습니다"});
  if(b64.length > 8000000) return json_({ok:false, error:"사진 용량이 너무 큽니다 — 다시 촬영해보세요"});

  var mime = String(body.mime||"image/jpeg");
  var storeName = String(t.values[t.head.indexOf("가게명")]||"").trim() || "매장";
  var stamp = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyyMMdd_HHmmss");
  var fname = (String(body.no)+"_"+storeName+"_"+field+"_"+stamp+"_"+(urls.length+1)+".jpg").replace(/[\\\/:*?"<>|]/g, "_");

  var blob = Utilities.newBlob(Utilities.base64Decode(b64), mime, fname);
  var folder = photoFolder_();
  var file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  var url = "https://drive.google.com/thumbnail?id=" + file.getId() + "&sz=w1600";

  urls.push(url);
  var joined = urls.join("|");
  t.sh.getRange(t.row, col+1).setValue(joined);
  stamp_(t.sh, t.head, t.row, name);

  return json_({ok:true, no:body.no, field:field, value:joined});
}

// 저장된 사진 URL에서 드라이브 파일ID 추출
function fileIdFromUrl_(url){
  var m = String(url||"").match(/[?&]id=([^&]+)/);
  return m ? m[1] : "";
}

// action:'deletePhoto' → 드라이브 파일 휴지통 이동 + 트래커 셀 값에서 제거
function handleDeletePhoto_(body){
  var auth = auth_(body.name, body.pw);
  if(!auth) return json_({ok:false, error:"인증 실패 — 다시 로그인하세요"});
  var name = auth.name;

  var field = String(body.field||"").trim();
  if(PHOTO_FIELDS.indexOf(field) < 0) return json_({ok:false, error:"사진 항목이 아닙니다: "+field});

  var t = findRow_(body.no);
  if(!t) return json_({ok:false, error:"행을 찾을 수 없습니다: "+body.no});

  if(!auth.isAdmin){
    var ownerIdx = t.head.indexOf("담당컨설턴트");
    if(ownerIdx >= 0 && String(t.values[ownerIdx]).trim() !== auth.name){
      return json_({ok:false, error:"본인 담당 건만 증빙을 삭제할 수 있습니다"});
    }
    var asIdx = t.head.indexOf("AS신청상태");
    var asSt = asIdx >= 0 ? String(t.values[asIdx]).trim() : "";
    if(asSt === "승인" || asSt === "불가능"){
      return json_({ok:false, error:"이미 처리된 A/S 건이라 증빙을 삭제할 수 없습니다"});
    }
  }

  var col = t.head.indexOf(field);
  if(col < 0) return json_({ok:false, error:'"트래커" 탭에 "'+field+'" 컬럼이 없습니다. setupTrackerSheet를 다시 실행하세요.'});

  var existing = String(t.values[col]||"").trim();
  var urls = existing ? existing.split("|").map(function(s){ return s.trim(); }).filter(Boolean) : [];

  var idx = Number(body.idx);
  if(isNaN(idx) || idx < 0 || idx >= urls.length){
    return json_({ok:false, error:"삭제할 사진을 찾을 수 없습니다"});
  }

  var removedUrl = urls[idx];
  var fid = fileIdFromUrl_(removedUrl);
  if(fid){
    try{ DriveApp.getFileById(fid).setTrashed(true); }
    catch(e){ /* 이미 삭제됐거나 접근 불가 — 셀 값은 그대로 비운다 */ }
  }

  urls.splice(idx, 1);
  var joined = urls.join("|");
  t.sh.getRange(t.row, col+1).setValue(joined);
  stamp_(t.sh, t.head, t.row, name);

  return json_({ok:true, no:body.no, field:field, value:joined});
}

// ── 일회성 마이그레이션 (2026-09-07, v2): 사용자가 서포터즈 시트에 직접 채워넣은 "컨설팅동의일시"로 DB지급일 소급 채우기 ──
// 실행 전제: 컨설턴트 트래커의 setupTrackerSheet가 이미 실행되어 "DB지급일" 컬럼이 있어야 함(완료됨).
// 서포터즈 스프레드시트("트래커" 탭)에서 가게명+연락처로 매칭해 "컨설팅동의일시"를 찾고,
// DB지급일 = 컨설팅동의일시 + 1영업일(주말 제외)로 계산해 채움. 이미 값이 있는 행은 건드리지 않음(비파괴).
function backfillDbJigeupilFromAgreeDate_20260907(){
  var SUPPORTER_SPREADSHEET_ID = "1ewvEx1GdEzhVsIdbymxevSGbdusimLNcW0PQEvqQamw";
  var addBiz = function(date, n){
    var d = new Date(date.getTime());
    var added = 0;
    while(added < n){
      d.setDate(d.getDate() + 1);
      var dow = parseInt(Utilities.formatDate(d, "Asia/Seoul", "u"), 10); // 1=월 ... 6=토, 7=일
      if(dow < 6) added++;
    }
    return d;
  };

  var sh = trackerSheet_();
  var data = sh.getDataRange().getValues();
  var head = data[0].map(function(h){ return String(h).trim(); });
  var colDbpay = head.indexOf("DB지급일");
  if(colDbpay < 0){
    Logger.log('"DB지급일" 컬럼이 없습니다. setupTrackerSheet를 먼저 실행한 뒤 다시 실행하세요.');
    return;
  }
  var colStore = head.indexOf("가게명");
  var colPhone = head.indexOf("연락처");

  var ssup = SpreadsheetApp.openById(SUPPORTER_SPREADSHEET_ID);
  var shSup = ssup.getSheetByName("트래커");
  var supData = shSup.getDataRange().getValues();
  var supHead = supData[0].map(function(h){ return String(h).trim(); });
  var sColStore = supHead.indexOf("가게명");
  var sColPhone = supHead.indexOf("연락처");
  var sColAgreedAt = supHead.indexOf("컨설팅동의일시");
  if(sColAgreedAt < 0){
    Logger.log('서포터즈 시트에 "컨설팅동의일시" 컬럼이 없습니다.');
    return;
  }

  var agreeMap = {};
  for(var i=1;i<supData.length;i++){
    var row = supData[i];
    var at = row[sColAgreedAt];
    if(!at) continue;
    var key = String(row[sColStore]||"").trim() + "|" + String(row[sColPhone]||"").trim();
    agreeMap[key] = (at instanceof Date) ? at : new Date(at);
  }

  var updated = 0, skippedHasValue = 0, skippedNoMatch = 0;
  for(var r=1;r<data.length;r++){
    var existing = String(data[r][colDbpay]||"").trim();
    if(existing){ skippedHasValue++; continue; }
    var key2 = String(data[r][colStore]||"").trim() + "|" + String(data[r][colPhone]||"").trim();
    var agreedAt = agreeMap[key2];
    if(!agreedAt){ skippedNoMatch++; continue; }
    var dbDate = addBiz(agreedAt, 1);
    sh.getRange(r+1, colDbpay+1).setValue(Utilities.formatDate(dbDate, "Asia/Seoul", "yyyy-MM-dd"));
    updated++;
  }
  Logger.log("DB지급일 소급 채우기(v2, 컨설팅동의일시 기준) 완료 — 갱신 "+updated+"건, 이미 값 있어서 건너뜀 "+skippedHasValue+"건, 매칭 실패 "+skippedNoMatch+"건.");
}

// ── 일회성 마이그레이션 (2026-09-07, v3): DB지급일/AS신청기한이 구글시트에 의해 자동으로
// 날짜(Date) 타입 셀로 바뀐 것을 "yyyy-MM-dd" 순수 텍스트로 되돌림.
// 실행 전제: 위 setupTrackerSheet를 먼저 재실행해서 두 열이 텍스트 서식(@)으로 고정돼 있어야 함
// (안 그러면 다시 써넣는 순간 시트가 또 날짜로 자동 변환해버림).
function fixDateColumnsToText_20260907(){
  var sh = trackerSheet_();
  var data = sh.getDataRange().getValues();
  var head = data[0].map(function(h){ return String(h).trim(); });
  var targets = ["DB지급일","AS신청기한"];
  var fixed = 0, already = 0;
  targets.forEach(function(colName){
    var col = head.indexOf(colName);
    if(col < 0) return;
    for(var r=1; r<data.length; r++){
      var v = data[r][col];
      if(v instanceof Date){
        var s = Utilities.formatDate(v, "Asia/Seoul", "yyyy-MM-dd");
        sh.getRange(r+1, col+1).setValue(s);
        fixed++;
      } else if(v){
        already++;
      }
    }
  });
  Logger.log("DB지급일/AS신청기한 텍스트 변환 완료 — 날짜타입→텍스트 변환 "+fixed+"건, 이미 텍스트였던 값 "+already+"건. "+
             "(먼저 setupTrackerSheet를 재실행해서 두 열이 텍스트 서식으로 고정돼 있어야 재발하지 않습니다)");
}

// ── 일회성 마이그레이션 (2026-09-07, v5): "2026.09" 원본 탭(자유서식)에 이미 기록돼 있던
// 담당설계사/월납보험료/연령대·성별/중대질환 값으로 사전체크 5항목을 소급 채움(서포터즈 PRECHECK_FIELDS와 동일한 5개 컬럼).
// 매장명 + 점주 연락처(숫자만 비교)로 매칭. 이 5개 컬럼 중 하나라도 이미 값이 있는 행은 건드리지 않음(비파괴 — 그룹 단위로 취급).
// 원본 자유서식 값이 아래 매핑표에 없는 애매한 값이면 짐작해서 채우지 않고 로그로만 남김.
// (v4에서 PREMIUM_LEVELS를 참조했는데 이 파일엔 정의돼 있지 않아 실행 시 오류가 났을 버그도 이번에 같이 고침 — 아래처럼 이 함수 안에서 직접 옵션표를 정의함)
function backfillPreCheckFromOriginalSheet_20260907(){
  var ORIGINAL_SHEET_NAME = "2026.09";
  var normPhone = function(v){
    var s = String(v||"").replace(/\D/g, "");
    if(s.indexOf("0") === 0) s = s.substring(1);
    return s;
  };
  // 원본 자유서식 값 → 서포터즈 PRECHECK_FIELDS 옵션 매핑표. 여기 없는 값은 애매하다고 보고 건너뜀(로그만 남김).
  var RELATION_MAP = { "담당자 없음":"없음", "없음":"없음", "본인":"본인", "부모":"부모", "자녀":"자녀", "배우자":"배우자" };
  var PREMIUM_OPTIONS = ["10만 원 미만","10만 원","20만 원","30만 원","40만 원","50만 원 이상","기타(잘 모르겠음/본인이 납부안함)"];
  var premiumMap = {};
  PREMIUM_OPTIONS.forEach(function(p){ premiumMap[p.replace(/\s/g,"")] = p; });
  var AGE_OPTIONS = ["10대","20대","30대","40대","50대","60대","70대 이상"];
  var GENDER_OPTIONS = ["남성","여성"];

  var sh = trackerSheet_();
  var data = sh.getDataRange().getValues();
  var head = data[0].map(function(h){ return String(h).trim(); });
  var colRel = head.indexOf("기존설계사관계");
  var colPrem = head.indexOf("월납보험료수준");
  var colAge = head.indexOf("연령대");
  var colGender = head.indexOf("성별");
  var colDis = head.indexOf("3대질환진단여부");
  var colStore = head.indexOf("가게명");
  var colPhone = head.indexOf("연락처");
  if(colRel<0 || colPrem<0 || colDis<0 || colAge<0 || colGender<0){
    Logger.log("사전체크 컬럼(기존설계사관계/월납보험료수준/연령대/성별/3대질환진단여부)이 없습니다. setupTrackerSheet를 먼저 실행하세요.");
    return;
  }

  var osh = findSheetByNameLoose_(ss_(), ORIGINAL_SHEET_NAME);
  if(!osh){ Logger.log('"' + ORIGINAL_SHEET_NAME + '" 탭을 찾을 수 없습니다.'); return; }
  var odata = osh.getDataRange().getValues();
  var ohead = odata[0].map(function(h){ return String(h).trim(); });
  var oColName = ohead.indexOf("매장명");
  var oColPhone = ohead.indexOf("점주 연락처");
  var oColRel = ohead.indexOf("담당설계사");
  var oColPrem = ohead.indexOf("월납보험료");
  var oColAgeGender = ohead.indexOf("연령대 / 성별");
  var oColDis = ohead.indexOf("중대질환");
  if(oColName<0 || oColPhone<0){
    Logger.log('"' + ORIGINAL_SHEET_NAME + '" 탭에서 "매장명"/"점주 연락처" 컬럼을 찾을 수 없습니다.');
    return;
  }

  var srcMap = {};
  for(var i=1;i<odata.length;i++){
    var orow = odata[i];
    var name = String(orow[oColName]||"").trim();
    if(!name || name==="ex)") continue; // 예시행 제외
    var key = name + "|" + normPhone(orow[oColPhone]);
    var ageGenderRaw = oColAgeGender>=0 ? String(orow[oColAgeGender]||"").trim() : ""; // 예: "30대 / 남성"
    var ageGenderParts = ageGenderRaw.split("/").map(function(s){ return s.trim(); });
    srcMap[key] = {
      relation: String(orow[oColRel]||"").trim(),
      premium: String(orow[oColPrem]||"").trim(),
      age: ageGenderParts[0] || "",
      gender: ageGenderParts[1] || "",
      disease: String(orow[oColDis]||"").trim()
    };
  }

  var updated=0, skippedHasValue=0, skippedNoMatch=0, ambiguous=[];
  for(var r=1;r<data.length;r++){
    var hasAny = String(data[r][colRel]||"").trim() || String(data[r][colPrem]||"").trim() ||
                 String(data[r][colAge]||"").trim() || String(data[r][colGender]||"").trim() ||
                 String(data[r][colDis]||"").trim();
    if(hasAny){ skippedHasValue++; continue; }
    var storeName = String(data[r][colStore]||"").trim();
    var key2 = storeName + "|" + normPhone(data[r][colPhone]);
    var src = srcMap[key2];
    if(!src){ skippedNoMatch++; continue; }

    var relVal = RELATION_MAP[src.relation];
    var premVal = premiumMap[src.premium.replace(/\s/g,"")];
    var ageVal = AGE_OPTIONS.indexOf(src.age) >= 0 ? src.age : "";
    var genderVal = GENDER_OPTIONS.indexOf(src.gender) >= 0 ? src.gender : "";
    var disVal = "";
    if(src.disease === "없음") disVal = "아니오";
    else if(src.disease === "있음") disVal = "예: 상세 미입력(원본 시트에 세부 진단명 없음)";

    if(!relVal || !premVal || !ageVal || !genderVal || !disVal){
      ambiguous.push(storeName + " (담당설계사='"+src.relation+"', 월납보험료='"+src.premium+"', 연령대/성별='"+(src.age+"/"+src.gender)+"', 중대질환='"+src.disease+"')");
      continue;
    }

    sh.getRange(r+1, colRel+1).setValue(relVal);
    sh.getRange(r+1, colPrem+1).setValue(premVal);
    sh.getRange(r+1, colAge+1).setValue(ageVal);
    sh.getRange(r+1, colGender+1).setValue(genderVal);
    sh.getRange(r+1, colDis+1).setValue(disVal);
    updated++;
  }
  Logger.log("사전체크 5항목 소급 채우기 완료 — 갱신 "+updated+"건, 이미 값 있어서 건너뜀 "+skippedHasValue+"건, 매칭 실패 "+skippedNoMatch+"건, 매핑 애매해서 건너뜀 "+ambiguous.length+"건" +
             (ambiguous.length ? (": " + ambiguous.join(" / ")) : "") + ".");
}

// ── 동기화 (2026-09-07): "2026.09" 탭(자유서식 공유시트, 팀이 계속 병행해서 쓰는 원본)의
// 1차/2차/3차 컨설팅 현황·일정을 컨설턴트 "트래커"의 컨설팅미팅1차/컨설팅미팅2_3차 컬럼으로 반영.
// (서포터즈 쪽 syncFromSharedSheet_20260907과 같은 "공유시트가 원본" 동기화 작업의 컨설턴트 쪽 부분 —
// 자세한 배경/결정사항은 컨설턴트_트래커_설계안_v1.md 11절, 서포터즈_트래커_개발이력.md 9절 참고)
//   "1차 현황" + "1차 컨설팅 일정" → 컨설팅미팅1차 (예: "확정 · 9/7(월) 14시")
//   "2차 현황" + "2차 컨설팅 일정", "3차 현황" + "3차 컨설팅 일정" → 컨설팅미팅2_3차
//     (2차·3차 각각 "현황 · 일정" 형태로 합친 뒤, 둘 다 있으면 " / "로 이어붙임 — 예: "2차: 확정 · 9/10(목) / 3차: 대기")
// 이 두 컬럼은 컨설턴트가 트래커 화면에서 직접 자유롭게 수정하는 편집 가능 필드라서(자동배정 필드인
// 담당서포터즈 등과 성격이 다름), 여기서는 값이 달라도 덮어쓰지 않고 **트래커 쪽이 비어있을 때만** 채움
// (비파괴). "🎉 청약" 컬럼은 자유서식이라 청약여부/청약차수/청약금액/청약상품으로 안전하게 자동 파싱하기
// 어려워 사용자 확인 후 이번 동기화 대상에서 제외함 — 필요시 트래커 UI에서 컨설턴트가 직접 입력.
function syncMeetingScheduleFromSharedSheet_20260907(){
  var ORIGINAL_SHEET_NAME = "2026.09";
  var normPhone = function(v){
    var s = String(v||"").replace(/\D/g, "");
    if(s.indexOf("0") === 0) s = s.substring(1);
    return s;
  };
  var combine = function(status, sched){
    status = String(status||"").trim();
    sched = String(sched||"").trim();
    if(status && sched) return status + " · " + sched;
    return status || sched || "";
  };

  var osh = findSheetByNameLoose_(ss_(), ORIGINAL_SHEET_NAME);
  if(!osh){ Logger.log('"' + ORIGINAL_SHEET_NAME + '" 탭을 찾을 수 없습니다.'); return; }
  var odata = osh.getDataRange().getValues();
  var ohead = odata[0].map(function(h){ return String(h).trim(); });
  var oColName = ohead.indexOf("매장명");
  var oColPhone = ohead.indexOf("점주 연락처");
  var oCol1Status = ohead.indexOf("1차 현황");
  var oCol1Sched = ohead.indexOf("1차 컨설팅 일정");
  var oCol2Status = ohead.indexOf("2차 현황");
  var oCol2Sched = ohead.indexOf("2차 컨설팅 일정");
  var oCol3Status = ohead.indexOf("3차 현황");
  var oCol3Sched = ohead.indexOf("3차 컨설팅 일정");
  if(oColName<0 || oColPhone<0){
    Logger.log('"' + ORIGINAL_SHEET_NAME + '" 탭에서 "매장명"/"점주 연락처" 컬럼을 찾을 수 없습니다.');
    return;
  }

  var srcMap = {};
  for(var i=1;i<odata.length;i++){
    var orow = odata[i];
    var name = String(orow[oColName]||"").trim();
    if(!name || name==="ex)") continue;
    var key = name + "|" + normPhone(orow[oColPhone]);
    var part1 = combine(oCol1Status>=0?orow[oCol1Status]:"", oCol1Sched>=0?orow[oCol1Sched]:"");
    var part2 = combine(oCol2Status>=0?orow[oCol2Status]:"", oCol2Sched>=0?orow[oCol2Sched]:"");
    var part3 = combine(oCol3Status>=0?orow[oCol3Status]:"", oCol3Sched>=0?orow[oCol3Sched]:"");
    var parts23 = [];
    if(part2) parts23.push("2차: " + part2);
    if(part3) parts23.push("3차: " + part3);
    srcMap[key] = { meeting1: part1, meeting23: parts23.join(" / ") };
  }

  var sh = trackerSheet_();
  var data = sh.getDataRange().getValues();
  var head = data[0].map(function(h){ return String(h).trim(); });
  var colStore = head.indexOf("가게명");
  var colPhone = head.indexOf("연락처");
  var colM1 = head.indexOf("컨설팅미팅1차");
  var colM23 = head.indexOf("컨설팅미팅2_3차");
  if(colStore<0 || colPhone<0 || colM1<0 || colM23<0){
    Logger.log("트래커 탭 필수 컬럼(가게명/연락처/컨설팅미팅1차/컨설팅미팅2_3차)이 없습니다.");
    return;
  }

  var filled1=0, filled23=0, skippedNoMatch=0;
  for(var r=1;r<data.length;r++){
    var storeName = String(data[r][colStore]||"").trim();
    if(!storeName) continue;
    var key2 = storeName + "|" + normPhone(data[r][colPhone]);
    var src = srcMap[key2];
    if(!src){ skippedNoMatch++; continue; }
    var rowNum = r+1;

    if(src.meeting1 && !String(data[r][colM1]||"").trim()){
      sh.getRange(rowNum, colM1+1).setValue(src.meeting1);
      filled1++;
    }
    if(src.meeting23 && !String(data[r][colM23]||"").trim()){
      sh.getRange(rowNum, colM23+1).setValue(src.meeting23);
      filled23++;
    }
  }

  Logger.log(
    "컨설팅미팅 일정 동기화 완료 — 컨설팅미팅1차 신규채움 " + filled1 + "건, " +
    "컨설팅미팅2_3차 신규채움 " + filled23 + "건, " +
    "매칭 실패 " + skippedNoMatch + "건. (이미 값이 있던 행은 비파괴 원칙에 따라 건드리지 않음)"
  );
}

// 2026-09-08 추가 — syncMeetingScheduleFromSharedSheet_20260907()를 1시간마다 자동 실행되게 하는
// 트리거 설치/제거. 최초 1회만 setupHourlyTrigger_20260908()을 Apps Script 편집기에서 실행하면
// 됨(재실행해도 중복 생성 안 되고 안전). 자동 실행 결과는 "실행 기록"에서 Logger.log 메시지로 확인 가능.
function setupHourlyTrigger_20260908(){
  var FN = "syncMeetingScheduleFromSharedSheet_20260907";
  ScriptApp.getProjectTriggers().forEach(function(t){
    if(t.getHandlerFunction() === FN) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger(FN).timeBased().everyHours(1).create();
  Logger.log(FN + " — 1시간마다 자동 실행 트리거 설치 완료.");
}
function removeHourlyTrigger_20260908(){
  var FN = "syncMeetingScheduleFromSharedSheet_20260907";
  var removed = 0;
  ScriptApp.getProjectTriggers().forEach(function(t){
    if(t.getHandlerFunction() === FN){ ScriptApp.deleteTrigger(t); removed++; }
  });
  Logger.log(FN + " — 트리거 " + removed + "개 제거 완료.");
}

// 유니코드 정규화 차이(NFC/NFD)로 인해 getSheetByName이 육안상 동일한 이름의 탭을
// 못 찾는 문제를 방지하기 위한 느슨한 탭 찾기(정규화+trim 후 비교). 특히 한글이 섞인
// 탭 이름(다른 사람이 다른 환경에서 만든 "공유시트" 등)에서 이런 불일치가 생길 수 있음
// (2026-09-07, syncFromSharedSheet_20260907이 "탭을 찾을 수 없습니다" 오류를 낸 것을 보고 추가).
// 일회성 마이그레이션(2026-09-07, 5차): 서포터즈 트래커에 있는 위도/경도를 가게명+연락처로 매칭해
// 이 트래커의 비어있는 위도/경도 칸에만 채움(비파괴). 앞으로 새로 전환되는 건은 handleConvert_가 자동 복사하므로
// 이 함수는 이미 전환완료된 기존 행에 대해서만 필요.
function backfillLatLngFromSupporter_20260907(){
  var normPhone = function(v){
    var s = String(v||"").replace(/\D/g, "");
    if(s.indexOf("0") === 0) s = s.substring(1);
    return s;
  };
  var sss = SpreadsheetApp.openById(SUPPORTER_SPREADSHEET_ID);
  var ssh = sss.getSheetByName("트래커");
  if(!ssh){ Logger.log('서포터즈 스프레드시트에서 "트래커" 탭을 찾을 수 없습니다.'); return; }
  var sdata = ssh.getDataRange().getValues();
  var shead = sdata[0].map(function(h){ return String(h).trim(); });
  var sColStore = shead.indexOf("가게명"), sColPhone = shead.indexOf("연락처"),
      sColLat = shead.indexOf("위도"), sColLng = shead.indexOf("경도");
  if(sColStore<0 || sColPhone<0 || sColLat<0 || sColLng<0){
    Logger.log("서포터즈 트래커에서 가게명/연락처/위도/경도 컬럼을 찾을 수 없습니다.");
    return;
  }
  var srcMap = {};
  for(var i=1;i<sdata.length;i++){
    var srow = sdata[i];
    var name = String(srow[sColStore]||"").trim();
    if(!name) continue;
    srcMap[name + "|" + normPhone(srow[sColPhone])] = { lat: srow[sColLat], lng: srow[sColLng] };
  }

  var sh = trackerSheet_();
  var data = sh.getDataRange().getValues();
  var head = data[0].map(function(h){ return String(h).trim(); });
  var colStore = head.indexOf("가게명"), colPhone = head.indexOf("연락처"),
      colLat = head.indexOf("위도"), colLng = head.indexOf("경도");
  if(colStore<0 || colPhone<0 || colLat<0 || colLng<0){
    Logger.log("컨설턴트 트래커에서 가게명/연락처/위도/경도 컬럼을 찾을 수 없습니다. setupTrackerSheet를 재실행하세요.");
    return;
  }

  var updated=0, skippedHasValue=0, skippedNoMatch=0;
  for(var r=1;r<data.length;r++){
    var storeName = String(data[r][colStore]||"").trim();
    if(!storeName) continue;
    if(String(data[r][colLat]||"").trim() && String(data[r][colLng]||"").trim()){ skippedHasValue++; continue; }
    var src = srcMap[storeName + "|" + normPhone(data[r][colPhone])];
    if(!src || src.lat==="" || src.lng===""){ skippedNoMatch++; continue; }
    sh.getRange(r+1, colLat+1).setValue(src.lat);
    sh.getRange(r+1, colLng+1).setValue(src.lng);
    updated++;
  }
  Logger.log(
    "서포터즈 좌표 소급 채우기 완료 — 갱신 " + updated + "건, 이미 값 있어서 건너뜀 " + skippedHasValue + "건, " +
    "매칭 실패/좌표없음 " + skippedNoMatch + "건."
  );
}

function findSheetByNameLoose_(ss, name){
  var target = String(name).normalize("NFC").trim();
  var sheets = ss.getSheets();
  for(var i=0;i<sheets.length;i++){
    var actual = sheets[i].getName().normalize("NFC").trim();
    if(actual === target) return sheets[i];
  }
  return null;
}
function stamp_(sh, head, row, name){
  var uc = head.indexOf("수정자");   if(uc >= 0) sh.getRange(row, uc+1).setValue(name);
  var tc = head.indexOf("수정시각"); if(tc >= 0) sh.getRange(row, tc+1).setValue(now_());
}
