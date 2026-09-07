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
 *  6) 함수 목록에서 setupStagingSheet 실행 → "신규유입" 탭 생성 (신규 매장 리스트 반입용)
 *
 * ※ 과거 데이터 자동 이관은 하지 않습니다. 필요하면 사용자가 "트래커" 탭에 직접 복사해 넣으세요.
 *
 * ※ 신규 매장 계속 추가하기 — "신규유입" 탭 + 원클릭 메뉴:
 *   새 매장 리스트가 생기면 "신규유입" 탭에 담당서포터즈까지 채워서 붙여넣고,
 *   시트 상단 메뉴 "🎯 동선 관리 → 신규매장 트래커에 반영"을 클릭하면 됩니다.
 *   담당서포터즈가 "계정" 탭에 없는 이름이면 팝업으로 직접 입력하거나 보류할 수 있고,
 *   이미 "트래커"에 있는 매장(가게명+연락처 동일)은 자동으로 건너뜁니다.
 *   처리된 행은 "신규유입" 탭에 반영완료/보류/중복-건너뜀으로 표시되어 남습니다.
 */

// ── 스프레드시트 ID ───────────────────────────────────────────
// 입점매장_시흥_액션팀용 (원천DB · 서포터즈 TA)
var SPREADSHEET_ID = "1ewvEx1GdEzhVsIdbymxevSGbdusimLNcW0PQEvqQamw";
// 동선_컨설팅 DB 관리용 (전환 대상 · 컨설턴트 트래커)
var CONSULTANT_SPREADSHEET_ID = "1J_otYQ_gMwVskqUHoCLdJIbkTytjF13P4t1044YN2jI";

// ── 신규 탭 이름 (기존 탭과 절대 겹치지 않게) ──────────────────
var TRACKER_SHEET = "트래커";
var ACCOUNT_SHEET = "계정";
// 컨설턴트 시트 쪽 신규 탭 (dongsun-consultant_AppsScript.gs 와 동일해야 함)
var CONSULTANT_TRACKER_SHEET = "트래커";

var HEADERS = ["번호","담당서포터즈","가게명","점주명","연락처","업종","동네","주소",
               "방문일정","TA진행상태","TA결과","컨설팅동의여부",
               "담당컨설턴트","전환상태","전환일시","비고","수정자","수정시각",
               "매장사진","동의서",
               "컨설팅동의일시", // 2026-09-07 추가: 컨설팅동의여부가 "컨설팅동의"로 바뀐 시점(DB지급일 계산 기준 — 전환일시와는 별개)
               "기존설계사관계","월납보험료수준","연령대","성별","3대질환진단여부", // 2026-09-07(3차) 추가: 동의서 등록 전 사전체크 항목(PRECHECK_FIELDS 참고) — action:'precheck'로 저장, handleConvert_가 컨설턴트 트래커로 1회 복사. 항목을 늘릴 때는 여기 컬럼 추가 + 아래 PRECHECK_FIELDS에 정의만 추가하면 됨(비파괴, 항상 뒤에 추가)
               "위도","경도"]; // 2026-09-07(5차) 추가: 주소 지오코딩 결과 좌표("가까운 순" 정렬 기능용) — handleAddStore_/importStagingToTracker에서 신규 등록 시 자동 채움, 기존 행은 backfillLatLngFromAddress_20260907()으로 소급. handleConvert_가 컨설턴트 트래커로 1회 복사.
var PHOTO_FIELDS = ["매장사진","동의서"];
var PHOTO_MAX = 5; // 사진 항목당 최대 등록 장수 — 셀에 URL을 "|"로 이어붙여 저장
var PHOTO_FOLDER_NAME = "동선_서포터즈_사진";
var ACCOUNT_HEADERS = ["이름","비번","권한"];

var TA_STATUSES = ["대기","방문확정","부재","재접촉필요","거절","보류"];
var AGREES = ["미접촉","컨설팅동의","컨설팅거절","보류"];
// 2026-09-07(3차) 추가: 동의서 등록 전 사전체크 팝업(action:'precheck') 항목 정의 — 프론트가 로그인 응답(precheckFields)으로
// 받아서 팝업을 통째로 동적 렌더링함. 항목을 추가하려면 위 HEADERS에 컬럼명을 추가하고 아래 배열에 정의 하나만 더 넣으면
// 팝업 UI·검증·시트 저장·컨설턴트 트래커 복사(handleConvert_)까지 전부 자동으로 반영됨(하드코딩 반복 없앰).
// type:"select" = 단일선택. type:"select_detail" = 단일선택 + options 중 detailOn 값을 고르면 상세텍스트 입력이 추가로 열리고
// 최종값은 "선택값: 상세내용" 형태로 한 컬럼에 결합 저장됨.
var RELATION_TYPES = ["본인","부모","자녀","배우자","없음"];
var PREMIUM_LEVELS = ["10만 원 미만","10만 원","20만 원","30만 원","40만 원","50만 원 이상","기타(잘 모르겠음/본인이 납부안함)"];
var AGE_GROUPS = ["10대","20대","30대","40대","50대","60대","70대 이상"];
var GENDERS = ["남성","여성"];
var PRECHECK_FIELDS = [
  { key:"기존설계사관계", question:"기존 보험설계사 관계", type:"select", options:RELATION_TYPES },
  { key:"월납보험료수준", question:"현재 납입 중이신 월평균 보험료 수준은 어떻게 되시나요?", type:"select", options:PREMIUM_LEVELS },
  { key:"연령대", question:"연령대", type:"select", options:AGE_GROUPS },
  { key:"성별", question:"성별", type:"select", options:GENDERS },
  { key:"3대질환진단여부", question:"주요 3대 질환(암, 뇌혈관, 심장질환) 진단을 받으신 적 있으신가요?", type:"select_detail", options:["아니오","예"], detailOn:"예" }
];
var CONVERT_STATUSES = ["","전환완료"];

// 컨설턴트 트래커(수신 측) 헤더 — 컨설턴트 백엔드와 동일하게 유지할 것
var CONSULTANT_HEADERS = ["번호","담당컨설턴트","가게명","점주명","연락처","업종","동네","주소","출처서포터즈",
                          "월납보험료","컨설팅미팅1차","컨설팅미팅2_3차","클로징확률","계약현황",
                          "비고","수정자","수정시각"];

// ── 시트 접근 (이름 지정 = 기존 탭 무간섭) ─────────────────────
var _ssCache_ = null;
function ss_(){
  if(!_ssCache_){
    _ssCache_ = SpreadsheetApp.openById(SPREADSHEET_ID);
    if(!_ssCache_) throw new Error("스프레드시트를 열 수 없습니다: " + SPREADSHEET_ID);
  }
  return _ssCache_;
}
// 유니코드 정규화 차이(NFC/NFD)로 인해 getSheetByName이 육안상 동일한 이름의 탭을
// 못 찾는 문제를 방지하기 위한 느슨한 탭 찾기(정규화+trim 후 비교). 특히 한글이 섞인
// 탭 이름(다른 사람이 다른 환경에서 만든 "공유시트" 등)에서 이런 불일치가 생길 수 있음
// (2026-09-07, syncFromSharedSheet_20260907이 "탭을 찾을 수 없습니다" 오류를 낸 것을 보고 추가).
function findSheetByNameLoose_(ss, name){
  var target = String(name).normalize("NFC").trim();
  var sheets = ss.getSheets();
  for(var i=0;i<sheets.length;i++){
    var actual = sheets[i].getName().normalize("NFC").trim();
    if(actual === target) return sheets[i];
  }
  return null;
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
  sh.getRange(2, HEADERS.indexOf("TA결과")+1,   last, 1).setDataValidation(mk(TA_STATUSES));
  sh.getRange(2, HEADERS.indexOf("컨설팅동의여부")+1, last, 1).setDataValidation(mk(AGREES));
  sh.getRange(2, HEADERS.indexOf("전환상태")+1,     last, 1).setDataValidation(mk(CONVERT_STATUSES));
}
function autoWidth_(sh, n){
  try{ sh.autoResizeColumns(1, n); }catch(e){}
}

// 주소 → {lat,lng} 지오코딩 (Apps Script 내장 Maps 서비스 — 별도 API 키/과금 설정 없이 기본 무료 쿼터 내에서 동작.
// 실패하거나(주소 형식이 이상하거나 매칭 안 됨) 쿼터 초과 시 null 반환 — 호출부에서 null이면 위도/경도를 비워둠(에러로 막지 않음).
// 2026-09-07(5차) "가까운 순" 정렬 기능 추가 시 도입.
function geocodeAddress_(address){
  var addr = String(address||"").trim();
  if(!addr) return null;
  try{
    var geocoder = Maps.newGeocoder().setRegion("kr");
    var resp = geocoder.geocode(addr);
    if(resp && resp.status === "OK" && resp.results && resp.results.length){
      var loc = resp.results[0].geometry.location;
      return { lat: loc.lat, lng: loc.lng };
    }
  }catch(e){
    Logger.log("지오코딩 실패: \"" + addr + "\" — " + e);
  }
  return null;
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

// ── 1회성 라벨 변경: TA진행상태 "재방문예정" → "재접촉필요" (2026-09-04 옵션 개편) ──
// 옵션 개편(재방문예정→재접촉필요로 이름 변경 + 거절 옵션 추가)으로 인해, 이미 저장된
// "재방문예정" 값들이 새 드롭다운 목록에 없는 값이 되는 것을 방지하기 위한 1회성 정리 함수.
function fixTaStatusLabel_20260904(){
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
    if(String(values[i][0]).trim() === "재방문예정"){
      values[i][0] = "재접촉필요";
      fixed++;
    }
  }
  if(fixed > 0) range.setValues(values);
  Logger.log("TA진행상태 '재방문예정' → '재접촉필요' 변경: " + fixed + "건");
}

// ── 1회성 통합: "TA진행상태" 칼럼을 없애고 "TA결과"가 그 역할(5개 고정값 드롭다운)을
// 대신하도록 통합 (2026-09-05). 기존 TA진행상태 값을 TA결과 칸으로 옮기고, 만약 TA결과에
// 이미 다른 메모(자유 텍스트)가 들어있었다면 데이터 손실 없이 "비고" 칸으로 옮겨 보존한 뒤
// TA결과를 상태값으로 덮어씁니다. "TA진행상태" 칼럼 자체는 지우지 않습니다(열 순서가 밀리면
// 다른 칼럼과 어긋날 위험이 있어 그대로 두되, 앞으로 웹앱은 이 칼럼을 더 이상 읽거나 쓰지 않음).
function mergeTaIntoResult_20260905(){
  var sh = trackerSheet_();
  var lastRow = sh.getLastRow();
  if(lastRow < 2){ Logger.log("데이터 없음"); return; }
  var head = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(function(h){ return String(h).trim(); });
  var colStatus = head.indexOf("TA진행상태");
  var colResult = head.indexOf("TA결과");
  var colNote = head.indexOf("비고");
  if(colStatus < 0 || colResult < 0){ Logger.log("TA진행상태/TA결과 컬럼을 찾을 수 없습니다."); return; }
  var range = sh.getRange(2, 1, lastRow-1, head.length);
  var values = range.getValues();
  var moved=0, kept=0, noted=0;
  for(var i=0;i<values.length;i++){
    var status = String(values[i][colStatus]||"").trim();
    var result = String(values[i][colResult]||"").trim();
    if(!status) continue; // 상태값 자체가 없던 행은 손대지 않음
    if(!result){
      values[i][colResult] = status;
      moved++;
    } else if(TA_STATUSES.indexOf(result) >= 0){
      kept++; // 이미 정상 상태값과 일치 — 그대로 둠
    } else {
      if(colNote >= 0){
        var oldNote = String(values[i][colNote]||"").trim();
        values[i][colNote] = (oldNote ? oldNote + " / " : "") + "[구 TA결과] " + result;
      }
      values[i][colResult] = status;
      noted++;
    }
  }
  range.setValues(values);
  Logger.log("TA진행상태 → TA결과 통합: 신규반영 " + moved + "건 · 이미일치 " + kept + "건 · 메모는 비고로 보존 후 덮어씀 " + noted + "건");
}

// ── 동기화 (2026-09-07): "공유시트"("(6)장곡장현능곡연성군자 장곡동 122_2022~26" 탭, 팀이 병행해서
// 계속 수동으로 쓰고 있는 자유서식 원본)를 "트래커" 탭으로 반영. 공유시트가 원본(authoritative)이라는
// 전제 하에, 아래 5개 필드는 값이 다르면 공유시트 값으로 덮어씀(비고만 예외 — 비파괴):
//   "TA 진행"   → 담당서포터즈 (담당자 배정 필드라 값이 바뀌면 매번 로그에 남김)
//   "TA 결과"   → TA결과       (아래 매핑표로 정규화. 매핑에 없는 값은 짐작하지 않고 로그만 남기고 건너뜀)
//   "방문일정"  → 방문일정     (같은 스프레드시트 내 Date 값이라 시간대 변환 없이 그대로 복사)
//   "동의 여부" → 컨설팅동의여부 (매핑표로 정규화. 새로 "컨설팅동의"가 되는데 컨설팅동의일시가 비어있으면
//                                지금 시각으로 채움 — handleUpdate_의 기존 동작과 동일한 규칙)
//   "비고"     → 비고          (트래커 쪽 비고가 비어있을 때만 채움 — 기존 메모를 덮어쓰지 않음)
// 매칭 키: 가게명 + 점주 연락처(숫자만, 맨 앞 0 제거). 이후 시간 간격을 두고 재실행해도 안전하도록
// 이미 같은 값이면 건드리지 않고, 실제로 값이 바뀐 건수만 집계함.
function syncFromSharedSheet_20260907(){
  var SHARED_SHEET_NAME = "(6)장곡장현능곡연성군자 장곡동 122_2022~26";
  var normPhone = function(v){
    var s = String(v||"").replace(/\D/g, "");
    if(s.indexOf("0") === 0) s = s.substring(1);
    return s;
  };
  var TA_RESULT_MAP = { "대기":"대기", "방문 확정":"방문확정", "부재":"부재", "재연락필요":"재접촉필요", "거절":"거절", "보류":"보류" };
  var AGREE_MAP = { "동의":"컨설팅동의", "거절":"컨설팅거절", "보류":"보류" };

  var osh = findSheetByNameLoose_(ss_(), SHARED_SHEET_NAME);
  if(!osh){ Logger.log('"' + SHARED_SHEET_NAME + '" 탭을 찾을 수 없습니다.'); return; }
  var odata = osh.getDataRange().getValues();
  var ohead = odata[0].map(function(h){ return String(h).trim(); });
  var oColOwner = ohead.indexOf("TA 진행");
  var oColResult = ohead.indexOf("TA 결과");
  var oColNote = ohead.indexOf("비고");
  var oColVisit = ohead.indexOf("방문일정");
  var oColAgree = ohead.indexOf("동의 여부");
  var oColName = ohead.indexOf("가게명");
  var oColPhone = ohead.indexOf("점주 연락처");
  if(oColName<0 || oColPhone<0){
    Logger.log('"' + SHARED_SHEET_NAME + '" 탭에서 "가게명"/"점주 연락처" 컬럼을 찾을 수 없습니다.');
    return;
  }

  var srcMap = {};
  for(var i=1;i<odata.length;i++){
    var orow = odata[i];
    var name = String(orow[oColName]||"").trim();
    if(!name) continue;
    var key = name + "|" + normPhone(orow[oColPhone]);
    srcMap[key] = {
      owner: oColOwner>=0 ? String(orow[oColOwner]||"").trim() : "",
      result: oColResult>=0 ? String(orow[oColResult]||"").trim() : "",
      note: oColNote>=0 ? String(orow[oColNote]||"").trim() : "",
      visit: oColVisit>=0 ? orow[oColVisit] : "",
      agree: oColAgree>=0 ? String(orow[oColAgree]||"").trim() : ""
    };
  }

  var sh = trackerSheet_();
  var data = sh.getDataRange().getValues();
  var head = data[0].map(function(h){ return String(h).trim(); });
  var colStore = head.indexOf("가게명");
  var colPhone = head.indexOf("연락처");
  var colOwner = head.indexOf("담당서포터즈");
  var colResult = head.indexOf("TA결과");
  var colVisit = head.indexOf("방문일정");
  var colAgree = head.indexOf("컨설팅동의여부");
  var colAgreedAt = head.indexOf("컨설팅동의일시");
  var colNote = head.indexOf("비고");
  if(colStore<0 || colPhone<0 || colOwner<0 || colResult<0 || colVisit<0 || colAgree<0 || colNote<0){
    Logger.log("트래커 탭 필수 컬럼(가게명/연락처/담당서포터즈/TA결과/방문일정/컨설팅동의여부/비고)이 없습니다.");
    return;
  }

  var ownerChanges = [], resultUpdated=0, visitUpdated=0, agreeUpdated=0, noteFilled=0;
  var skippedNoMatch=0, ambiguousResult=[], ambiguousAgree=[];

  for(var r=1;r<data.length;r++){
    var storeName = String(data[r][colStore]||"").trim();
    if(!storeName) continue;
    var key2 = storeName + "|" + normPhone(data[r][colPhone]);
    var src = srcMap[key2];
    if(!src){ skippedNoMatch++; continue; }
    var rowNum = r+1;

    // 담당서포터즈 ← TA 진행 (값이 있고 다르면 덮어쓰고 로그)
    if(src.owner){
      var curOwner = String(data[r][colOwner]||"").trim();
      if(curOwner !== src.owner){
        sh.getRange(rowNum, colOwner+1).setValue(src.owner);
        ownerChanges.push(storeName + ": '" + curOwner + "' → '" + src.owner + "'");
      }
    }

    // TA결과 ← TA 결과 (매핑표로 정규화)
    if(src.result){
      var mappedResult = TA_RESULT_MAP[src.result];
      if(mappedResult){
        var curResult = String(data[r][colResult]||"").trim();
        if(curResult !== mappedResult){
          sh.getRange(rowNum, colResult+1).setValue(mappedResult);
          resultUpdated++;
        }
      } else {
        ambiguousResult.push(storeName + " (TA 결과='" + src.result + "')");
      }
    }

    // 방문일정 ← 방문일정 (같은 스프레드시트 내 Date라 그대로 복사)
    if(src.visit){
      var curVisit = data[r][colVisit];
      var curVisitMs = (curVisit instanceof Date) ? curVisit.getTime() : null;
      var srcVisitMs = (src.visit instanceof Date) ? src.visit.getTime() : null;
      if(srcVisitMs !== null && srcVisitMs !== curVisitMs){
        sh.getRange(rowNum, colVisit+1).setValue(src.visit);
        visitUpdated++;
      }
    }

    // 컨설팅동의여부 ← 동의 여부 (매핑표로 정규화 + 신규 동의 시 컨설팅동의일시 채움)
    if(src.agree){
      var mappedAgree = AGREE_MAP[src.agree];
      if(mappedAgree){
        var curAgree = String(data[r][colAgree]||"").trim();
        if(curAgree !== mappedAgree){
          sh.getRange(rowNum, colAgree+1).setValue(mappedAgree);
          agreeUpdated++;
          if(mappedAgree === "컨설팅동의" && colAgreedAt>=0 && !String(data[r][colAgreedAt]||"").trim()){
            sh.getRange(rowNum, colAgreedAt+1).setValue(now_());
          }
        }
      } else {
        ambiguousAgree.push(storeName + " (동의 여부='" + src.agree + "')");
      }
    }

    // 비고 ← 비고 (트래커 쪽이 비어있을 때만 채움 — 비파괴)
    if(src.note && !String(data[r][colNote]||"").trim()){
      sh.getRange(rowNum, colNote+1).setValue(src.note);
      noteFilled++;
    }
  }

  Logger.log(
    "공유시트→트래커 동기화 완료 — " +
    "담당서포터즈 변경 " + ownerChanges.length + "건" + (ownerChanges.length ? (" [" + ownerChanges.join(" / ") + "]") : "") + ", " +
    "TA결과 갱신 " + resultUpdated + "건, " +
    "방문일정 갱신 " + visitUpdated + "건, " +
    "컨설팅동의여부 갱신 " + agreeUpdated + "건, " +
    "비고 신규채움 " + noteFilled + "건, " +
    "매칭 실패 " + skippedNoMatch + "건, " +
    "TA결과 매핑 애매(건너뜀) " + ambiguousResult.length + "건" + (ambiguousResult.length ? (": " + ambiguousResult.join(" / ")) : "") + ", " +
    "동의여부 매핑 애매(건너뜀) " + ambiguousAgree.length + "건" + (ambiguousAgree.length ? (": " + ambiguousAgree.join(" / ")) : "") + "."
  );
}

// ── 공통 유틸 ───────────────────────────────────────────────
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
function ymd_(v){ // 방문일정 표시용 — 이제 시간까지 포함
  if(v instanceof Date) return Utilities.formatDate(v, "Asia/Seoul", "yyyy-MM-dd'T'HH:mm");
  return v;
}
// 프론트에서 온 "yyyy-MM-ddTHH:mm" 문자열을 한국시간(UTC+9) 기준 실제 Date로 변환.
// 스크립트 프로젝트의 타임존 설정과 무관하게 항상 정확한 시각이 저장되도록 UTC로 직접 계산.
function parseKstDateTime_(s){
  var m = String(s||"").match(/^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})/);
  if(!m) return null;
  var utcMs = Date.UTC(Number(m[1]), Number(m[2])-1, Number(m[3]), Number(m[4]), Number(m[5])) - 9*3600*1000;
  return new Date(utcMs);
}
function now_(){ return Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm"); }

// 영업일 계산(주말만 제외, 공휴일은 고려하지 않음) — DB지급일 = 생성(전환) 시점 + n영업일
function addBusinessDays_(date, n){
  var d = new Date(date.getTime());
  var added = 0;
  while(added < n){
    d.setDate(d.getDate() + 1);
    var dow = parseInt(Utilities.formatDate(d, "Asia/Seoul", "u"), 10); // 1=월 ... 6=토, 7=일
    if(dow < 6) added++;
  }
  return d;
}

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
    var cache = CacheService.getScriptCache();
    var cached = cache.get("consultantNames_v1");
    if(cached) return JSON.parse(cached);
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
    cache.put("consultantNames_v1", JSON.stringify(out), 300); // 5분 캐시 — 매 새로고침마다 다른 스프레드시트를 여는 비용 제거
    return out;
  }catch(e){ return []; }
}

// ── doGet / doPost ────────────────────────────────────────────
function doGet(e){
  return json_({ ok:true, service:"dongsun-supporter",
                 taStatuses:TA_STATUSES, agrees:AGREES,
                 precheckFields:PRECHECK_FIELDS, ts:new Date().getTime(),
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
      if(action === "precheck")    return handlePreCheck_(body);
      if(action === "convert")     return handleConvert_(body);
      if(action === "photo")       return handlePhoto_(body);
      if(action === "deletePhoto") return handleDeletePhoto_(body);
      if(action === "addStore")    return handleAddStore_(body);
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
    return String(r["담당서포터즈"]||"").trim() === auth.name;
  });
  return json_({
    ok:true, name:auth.name, isAdmin:auth.isAdmin,
    taStatuses:TA_STATUSES, agrees:AGREES,
    precheckFields:PRECHECK_FIELDS,
    consultants:consultantNames_(),
    supporters:supporterNames_(),
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

  var val = body.value;
  if(field === "방문일정"){
    var dt = parseKstDateTime_(val);
    if(dt) val = dt; // 날짜+시간을 실제 Date로 저장(시트에서도 날짜로 인식되도록)
  }
  t.sh.getRange(t.row, col+1).setValue(val);

  // 컨설팅동의여부가 "컨설팅동의"로 바뀌는 시점을 별도 컬럼에 기록 — DB지급일 계산 기준(전환일시와는 별개 이벤트)
  if(field === "컨설팅동의여부" && String(val).trim() === "컨설팅동의"){
    var agreeColIdx = t.head.indexOf("컨설팅동의일시");
    if(agreeColIdx >= 0) t.sh.getRange(t.row, agreeColIdx+1).setValue(now_());
  }

  stamp_(t.sh, t.head, t.row, auth.name);
  return json_({ok:true, no:body.no, field:field, value:body.value});
}

// action:'precheck' → 동의서 등록 전 사전체크 항목(PRECHECK_FIELDS 정의 전체)을 한 번에 저장
// (2026-09-07(3차) 일반화) 프론트가 { answers: { "기존설계사관계":"없음", "3대질환진단여부":"예", "3대질환진단여부Detail":"위암", ... } }
// 형태로 보내면, PRECHECK_FIELDS를 순회하며 검증·저장함 — 항목이 늘어나도 이 함수는 수정할 필요 없음.
// 동의서 사진을 올리기 전에 프론트에서 이 항목들이 비어있으면 팝업으로 먼저 받아서 이 액션으로 저장하고,
// 저장 성공 후에만 사진 선택창을 엶. handleConvert_가 컨설턴트 전환 시점에 이 값들을 컨설턴트 트래커로 1회 복사함.
function handlePreCheck_(body){
  var auth = auth_(body.name, body.pw);
  if(!auth) return json_({ok:false, error:"인증 실패 — 다시 로그인하세요"});

  var t = findRow_(body.no);
  if(!t) return json_({ok:false, error:"행을 찾을 수 없습니다: "+body.no});

  if(!auth.isAdmin){
    var ownerIdx = t.head.indexOf("담당서포터즈");
    if(ownerIdx >= 0 && String(t.values[ownerIdx]).trim() !== auth.name){
      return json_({ok:false, error:"본인 담당 건만 입력할 수 있습니다"});
    }
    var csIdx = t.head.indexOf("전환상태");
    if(csIdx >= 0 && String(t.values[csIdx]).trim() === "전환완료"){
      return json_({ok:false, error:"이미 컨설턴트로 전환된 건이라 수정할 수 없습니다"});
    }
  }

  var answers = body.answers || {};
  var result = {};
  for(var i=0;i<PRECHECK_FIELDS.length;i++){
    var f = PRECHECK_FIELDS[i];
    var raw = String(answers[f.key]||"").trim();
    if(f.options.indexOf(raw) < 0) return json_({ok:false, error:(f.question||f.key)+"을(를) 선택하세요"});
    var val = raw;
    if(f.type === "select_detail" && raw === f.detailOn){
      var detail = String(answers[f.key+"Detail"]||"").trim();
      val = raw + ": " + (detail || "상세 미입력");
    }
    result[f.key] = val;
  }

  var setIf = function(k, v){ var i2 = t.head.indexOf(k); if(i2>=0) t.sh.getRange(t.row, i2+1).setValue(v); };
  for(var j=0;j<PRECHECK_FIELDS.length;j++){ setIf(PRECHECK_FIELDS[j].key, result[PRECHECK_FIELDS[j].key]); }
  stamp_(t.sh, t.head, t.row, auth.name);

  return json_({ok:true, no:body.no, values:result});
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

  // DB지급일 계산 기준 시점 확보: "컨설팅동의" 시점이 원칙. 별도 동의 이벤트 없이(전환 버튼으로) 바로
  // 전환하는 경우엔 컨설팅동의일시가 비어있으므로, 지금을 동의 시점으로 기록해서 기준으로 삼음.
  var agreeColIdx2 = t.head.indexOf("컨설팅동의일시");
  var agreedRaw = agreeColIdx2 >= 0 ? t.values[agreeColIdx2] : "";
  if(!agreedRaw){
    agreedRaw = now_();
    if(agreeColIdx2 >= 0) t.sh.getRange(t.row, agreeColIdx2+1).setValue(agreedRaw);
  }
  var agreedDate = (agreedRaw instanceof Date) ? agreedRaw : new Date(agreedRaw);

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
    "위도": g("위도"),
    "경도": g("경도"),
    "출처서포터즈": g("담당서포터즈") || name,
    "DB지급일": Utilities.formatDate(addBusinessDays_(agreedDate, 1), "Asia/Seoul", "yyyy-MM-dd"), // 컨설팅동의 시점 + 1영업일(주말 제외) — 컨설턴트 쪽 A/S 신청기한 계산 기준
    "월납보험료": "",
    "컨설팅미팅1차": "",
    "컨설팅미팅2_3차": "",
    "클로징확률": "",
    "계약현황": "신규배정",
    "비고": "",
    "수정자": name,
    "수정시각": now_()
  };
  // 2026-09-07(3차): 사전체크 항목(PRECHECK_FIELDS)은 목록이 늘어날 수 있어 여기서 순회로 복사 —
  // 항목 추가 시 이 코드는 손댈 필요 없이 위 PRECHECK_FIELDS 배열에 정의만 추가하면 됨
  PRECHECK_FIELDS.forEach(function(f){ map[f.key] = g(f.key); });
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

  // 드라이브 업로드 전에 저장할 컬럼이 실제로 있는지 먼저 확인 (없으면 파일만 올리고 실패하는 것을 방지)
  var col = t.head.indexOf(field);
  if(col < 0) return json_({ok:false, error:'"트래커" 탭에 "'+field+'" 컬럼이 없습니다. setupTrackerSheet를 다시 실행하세요.'});

  // 기존 저장값은 "url1|url2|..." 형태(파이프 구분) — 기존 단일 URL 셀도 1개짜리 배열로 그대로 해석됨
  var existing = String(t.values[col]||"").trim();
  var urls = existing ? existing.split("|").map(function(s){ return s.trim(); }).filter(Boolean) : [];
  if(urls.length >= PHOTO_MAX){
    return json_({ok:false, error:"사진은 최대 "+PHOTO_MAX+"장까지만 등록할 수 있습니다"});
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
  // uc?export=view 방식은 구글이 핫링크(<img>) 렌더링을 자주 막아 미리보기가 깨짐 → thumbnail 엔드포인트로 변경(안정적으로 이미지 바이트 반환)
  var url = "https://drive.google.com/thumbnail?id=" + file.getId() + "&sz=w1600";

  urls.push(url);
  var joined = urls.join("|");
  t.sh.getRange(t.row, col+1).setValue(joined);
  stamp_(t.sh, t.head, t.row, name);

  return json_({ok:true, no:body.no, field:field, value:joined});
}

// 저장된 사진 URL(uc?export=view&id=... 또는 thumbnail?id=...) 어느 형식이든 파일ID 추출
function fileIdFromUrl_(url){
  var m = String(url||"").match(/[?&]id=([^&]+)/);
  return m ? m[1] : "";
}

// action:'deletePhoto' → 드라이브 파일 휴지통 이동 + 트래커 셀 값 비우기
function handleDeletePhoto_(body){
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
      return json_({ok:false, error:"본인 담당 건만 사진을 삭제할 수 있습니다"});
    }
    var csIdx = t.head.indexOf("전환상태");
    if(csIdx >= 0 && String(t.values[csIdx]).trim() === "전환완료"){
      return json_({ok:false, error:"이미 컨설턴트로 전환된 건이라 사진을 삭제할 수 없습니다"});
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

// action:'addStore' → 서포터즈가 현장 방문 중 새로 발견한 매장을 트래커에 즉시 등록.
// 일반 계정은 담당서포터즈가 항상 본인으로 고정(클라이언트가 뭘 보내든 무시). 관리자는 담당자 지정/재배정 가능.
function handleAddStore_(body){
  var auth = auth_(body.name, body.pw);
  if(!auth) return json_({ok:false, error:"인증 실패 — 다시 로그인하세요"});

  var storeName = String(body.storeName||"").trim();
  if(!storeName) return json_({ok:false, error:"가게명을 입력하세요"});

  var owner = auth.isAdmin ? (String(body.owner||"").trim() || auth.name) : auth.name;
  var phone = String(body.phone||"").trim();

  var sh = trackerSheet_();
  var data = sh.getDataRange().getValues();
  var head = data[0].map(function(h){ return String(h).trim(); });
  var colNo = head.indexOf("번호"), colStore = head.indexOf("가게명"), colPhone = head.indexOf("연락처");
  var colOwnerHdr = head.indexOf("담당서포터즈");

  var maxNo = 0;
  for(var i=1;i<data.length;i++){
    var r = data[i];
    if(String(r[colNo]).trim() !== "") maxNo = Math.max(maxNo, Number(r[colNo])||0);
    if(phone && String(r[colStore]||"").trim() === storeName && String(r[colPhone]||"").trim() === phone){
      return json_({ok:false, error:"이미 등록된 매장입니다 (담당: " + String(r[colOwnerHdr]||"").trim() + ")"});
    }
  }

  var newRow = new Array(head.length).fill("");
  var set = function(k, v){ var c = head.indexOf(k); if(c>=0) newRow[c]=v; };
  var no = maxNo + 1;
  set("번호", no);
  set("담당서포터즈", owner);
  set("가게명", storeName);
  set("점주명", String(body.ownerName||"").trim());
  set("연락처", phone);
  set("업종", String(body.biz||"").trim());
  set("동네", String(body.town||"").trim());
  var addStoreAddr_ = String(body.addr||"").trim();
  set("주소", addStoreAddr_);
  var addStoreGeo_ = geocodeAddress_(addStoreAddr_);
  if(addStoreGeo_){ set("위도", addStoreGeo_.lat); set("경도", addStoreGeo_.lng); }
  set("비고", String(body.note||"").trim());
  set("TA결과", TA_STATUSES[0] || "대기");
  set("수정자", auth.name);
  set("수정시각", now_());

  sh.getRange(sh.getLastRow()+1, 1, 1, head.length).setValues([newRow]);

  var obj = {};
  for(var c=0;c<head.length;c++){ obj[head[c]] = newRow[c]; }
  return json_({ok:true, row:obj});
}

// ── 일회성 마이그레이션(2026-09-07, 5차): 기존 행 중 위도/경도가 비어있는데 주소가 있는 행을
// 지오코딩해서 소급 채움("가까운 순" 정렬 기능이 기존 데이터에도 적용되도록). 이미 값이 있으면 건드리지 않음(비파괴).
function backfillLatLngFromAddress_20260907(){
  var sh = trackerSheet_();
  var data = sh.getDataRange().getValues();
  var head = data[0].map(function(h){ return String(h).trim(); });
  var colAddr = head.indexOf("주소"), colLat = head.indexOf("위도"), colLng = head.indexOf("경도"), colStore = head.indexOf("가게명");
  if(colAddr<0 || colLat<0 || colLng<0){
    Logger.log("주소/위도/경도 컬럼을 찾을 수 없습니다. setupTrackerSheet를 재실행하세요.");
    return;
  }
  var updated=0, skippedHasValue=0, skippedNoAddr=0, failed=[];
  for(var i=1;i<data.length;i++){
    var row = data[i];
    var store = String(row[colStore]||"").trim();
    if(!store) continue;
    if(String(row[colLat]||"").trim() && String(row[colLng]||"").trim()){ skippedHasValue++; continue; }
    var addr = String(row[colAddr]||"").trim();
    if(!addr){ skippedNoAddr++; continue; }
    var geo = geocodeAddress_(addr);
    if(geo){
      sh.getRange(i+1, colLat+1).setValue(geo.lat);
      sh.getRange(i+1, colLng+1).setValue(geo.lng);
      updated++;
    } else {
      failed.push(store + "(" + addr + ")");
    }
    Utilities.sleep(200); // 지오코딩 연속 호출 과부하 방지
  }
  Logger.log(
    "주소→좌표 소급 채우기 완료 — 갱신 " + updated + "건, 이미 값 있어서 건너뜀 " + skippedHasValue + "건, " +
    "주소 없어서 건너뜀 " + skippedNoAddr + "건, 지오코딩 실패 " + failed.length + "건" +
    (failed.length ? (": " + failed.join(" / ")) : "") + "."
  );
}


// ── "신규유입" 스테이징 탭 → "트래커" 탭 원클릭 반영 ────────────────
var STAGING_SHEET = "신규유입";
var STAGING_HEADERS = ["처리상태","담당서포터즈","가게명","점주명","연락처","업종","동네","주소","비고"];
var STAGING_STATUSES = ["","보류","반영완료","중복-건너뜀"];

function stagingSheet_(){
  var sh = ss_().getSheetByName(STAGING_SHEET);
  if(!sh) throw new Error('"' + STAGING_SHEET + '" 탭이 없습니다. setupStagingSheet를 먼저 실행하세요.');
  return sh;
}

// ── 최초 1회 실행: "신규유입" 탭 생성 ──────────────────────────
function setupStagingSheet(){
  var ss = ss_();
  var sh = ss.getSheetByName(STAGING_SHEET);
  if(!sh){
    sh = ss.insertSheet(STAGING_SHEET);
    Logger.log('"' + STAGING_SHEET + '" 탭을 새로 만들었습니다.');
  } else {
    Logger.log('"' + STAGING_SHEET + '" 탭이 이미 있습니다 — 헤더만 확인합니다(데이터 보존).');
  }
  sh.getRange(1,1,1,STAGING_HEADERS.length).setValues([STAGING_HEADERS]).setFontWeight("bold")
    .setBackground("#16335B").setFontColor("#FFFFFF");
  sh.setFrozenRows(1);
  var last = Math.max(sh.getMaxRows()-1, 1);
  var mk = SpreadsheetApp.newDataValidation().requireValueInList(STAGING_STATUSES, true).build();
  sh.getRange(2, STAGING_HEADERS.indexOf("처리상태")+1, last, 1).setDataValidation(mk);
  var phoneCol = STAGING_HEADERS.indexOf("연락처")+1;
  sh.getRange(2, phoneCol, last, 1).setNumberFormat("@");
  sh.getRange(1, STAGING_HEADERS.length+2).setValue(
    "← 새 매장 리스트를 2행부터 붙여넣으세요. 담당서포터즈는 아는 만큼만 채워도 됩니다(비어있거나 " +
    "계정에 없는 이름은 반영 시 팝업으로 물어봅니다). 처리상태는 자동으로 채워지니 직접 입력하지 마세요. " +
    "다 채웠으면 시트 메뉴 '🎯 동선 관리 → 신규매장 트래커에 반영'을 누르세요.");
  autoWidth_(sh, STAGING_HEADERS.length+3);
  Logger.log("신규유입 탭 세팅 완료.");
}

// 계정 탭에 등록된 서포터즈 이름 목록(관리자 제외) — 담당서포터즈 유효성 검사용
function supporterNames_(){
  var rows = sheetToObjects_(accountSheet_());
  var out = [];
  for(var i=0;i<rows.length;i++){
    if(String(rows[i]["권한"]||"").trim() === "관리자") continue;
    var n = String(rows[i]["이름"]||"").trim();
    if(n && out.indexOf(n) < 0) out.push(n);
  }
  return out;
}

// 시트를 열면 커스텀 메뉴를 자동으로 추가 (Apps Script 편집기 없이 시트에서 바로 실행)
function onOpen(){
  SpreadsheetApp.getUi()
    .createMenu('🎯 동선 관리')
    .addItem('신규매장 트래커에 반영', 'importStagingToTracker')
    .addToUi();
}

// "신규유입" 탭의 미처리 행을 "트래커" 탭에 반영. 시트 메뉴로 실행(팝업 사용 위해 UI 컨텍스트 필요).
// 담당서포터즈가 미등록 이름이면 행마다 팝업으로 물어봄 → 입력하면 그 이름으로 반영, 비워두면 "보류".
// 가게명+연락처가 이미 트래커에 있으면 "중복-건너뜀"으로 표시하고 넘어감.
function importStagingToTracker(){
  var ui = SpreadsheetApp.getUi();
  var stSh = stagingSheet_();
  var trSh = trackerSheet_();

  var stData = stSh.getDataRange().getValues();
  if(stData.length < 2){ ui.alert("신규유입 탭에 데이터가 없습니다."); return; }
  var stHead = stData[0].map(function(h){ return String(h).trim(); });
  var colStatus = stHead.indexOf("처리상태");
  var colOwner  = stHead.indexOf("담당서포터즈");
  var colStore  = stHead.indexOf("가게명");
  var colPhone  = stHead.indexOf("연락처");
  if(colStatus<0 || colOwner<0 || colStore<0){
    ui.alert('"신규유입" 탭 헤더가 올바르지 않습니다. setupStagingSheet를 다시 실행하세요.');
    return;
  }

  // 트래커 기존 행 로드 (중복 체크 + 번호 이어쓰기용)
  var trData = trSh.getDataRange().getValues();
  var trHead = trData[0].map(function(h){ return String(h).trim(); });
  var tNo = trHead.indexOf("번호"), tStore = trHead.indexOf("가게명"), tPhone = trHead.indexOf("연락처");
  var maxNo = 0;
  var existing = {}; // "가게명|연락처" → true
  for(var i=1;i<trData.length;i++){
    var r = trData[i];
    if(String(r[tNo]).trim() !== "") maxNo = Math.max(maxNo, Number(r[tNo])||0);
    var key = String(r[tStore]||"").trim() + "|" + String(r[tPhone]||"").trim();
    if(key !== "|") existing[key] = true;
  }

  var known = supporterNames_();
  var appendRows = [];
  var added=0, held=0, dup=0;

  for(var row=1; row<stData.length; row++){
    var r = stData[row];
    var storeName = String(r[colStore]||"").trim();
    var status = String(r[colStatus]||"").trim();
    if(!storeName) continue; // 빈 행
    if(status === "반영완료" || status === "중복-건너뜀") continue; // 이미 처리됨

    var owner = String(r[colOwner]||"").trim();
    if(!owner || known.indexOf(owner) < 0){
      var resp = ui.prompt(
        '담당서포터즈 확인 필요',
        '"' + storeName + '" 행의 담당서포터즈 "' + (owner||"(비어있음)") + '"이(가) 계정에 등록되어 있지 않습니다.\n' +
        '정확한 이름을 입력하고 확인을 누르거나, 비워둔 채 확인/취소를 누르면 이 행은 보류 처리됩니다.',
        ui.ButtonSet.OK_CANCEL);
      var typed = resp.getResponseText ? String(resp.getResponseText()).trim() : "";
      if(resp.getSelectedButton() !== ui.Button.OK || !typed){
        stSh.getRange(row+1, colStatus+1).setValue("보류");
        held++;
        continue;
      }
      owner = typed;
    }

    var key = storeName + "|" + String(r[colPhone]||"").trim();
    if(existing[key]){
      stSh.getRange(row+1, colStatus+1).setValue("중복-건너뜀");
      dup++;
      continue;
    }

    maxNo++;
    var newRow = new Array(trHead.length).fill("");
    var set = function(k, v){ var c = trHead.indexOf(k); if(c>=0) newRow[c]=v; };
    set("번호", maxNo);
    set("담당서포터즈", owner);
    set("가게명", storeName);
    set("점주명", r[stHead.indexOf("점주명")]);
    set("연락처", r[colPhone]);
    set("업종", r[stHead.indexOf("업종")]);
    set("동네", r[stHead.indexOf("동네")]);
    var stagingAddr_ = r[stHead.indexOf("주소")];
    set("주소", stagingAddr_);
    var stagingGeo_ = geocodeAddress_(stagingAddr_);
    if(stagingGeo_){ set("위도", stagingGeo_.lat); set("경도", stagingGeo_.lng); Utilities.sleep(150); }
    set("비고", r[stHead.indexOf("비고")]);
    set("TA결과", TA_STATUSES[0] || "대기");
    set("수정자", "일괄추가");
    set("수정시각", now_());
    appendRows.push(newRow);
    existing[key] = true;
    stSh.getRange(row+1, colStatus+1).setValue("반영완료");
    added++;
  }

  if(appendRows.length){
    trSh.getRange(trSh.getLastRow()+1, 1, appendRows.length, trHead.length).setValues(appendRows);
  }
  ui.alert("반영 완료: " + added + "건 추가 / " + held + "건 보류 / " + dup + "건 중복 건너뜀");
}