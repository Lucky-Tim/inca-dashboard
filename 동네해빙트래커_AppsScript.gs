/**
 * 동네해빙 트래커 — Google Apps Script API (구글시트 백엔드)
 * 사용법:
 *  1) 새 구글시트 만들기 → 확장 프로그램 → Apps Script
 *  2) 이 코드 전체를 붙여넣고 저장
 *  3) 상단 함수 목록에서 initSheet 선택 → 실행(최초 1회, 권한 승인) → 50곳 자동 생성
 *  4) 배포 → 새 배포 → 유형: 웹 앱 → 실행: 나 / 액세스: 모든 사용자 → 배포
 *  5) 나온 웹앱 URL(.../exec)을 트래커 페이지에 연결
 */

// ★ 독립형(standalone) 스크립트면 아래에 구글시트 ID를 넣으세요.
//   시트 URL https://docs.google.com/spreadsheets/d/★여기ID★/edit 의 가운데 부분.
//   구글시트 안에서 "확장 프로그램→Apps Script"로 만든 경우엔 비워둬도 됩니다.
var SPREADSHEET_ID = "";
var SHEET_NAME = "매장";
// 시트 원본 컬럼(동·주소·입점일·점주명·인입경로·요금제) + 운영 컬럼(등급·상태·방문일·계약금액·담당·메모) 통합
var HEADERS = ["번호","매장명","업종","동","주소","입점일","점주명","연락처","인입경로","요금제","등급","상태","방문일","계약금액","담당","메모","수정시각"];
var STATUSES = ["미착수","방문예정","부재","유선거절","상담접수","전환","계약체결","이탈"];

// 출처: 구글시트 "업셀링 26.06"(2026-01-01~05-31 입점 배곧동 매장 50곳)
// 열 순서: [매장명, 업종, 동, 주소, 입점일, 점주명, 연락처, 인입경로, 요금제, 등급(초안)]
// 등급 초안 규칙: 건강/뷰티=A(월납 기대 큼) · 생활/비활성=C · 음식점·카페=B → 현장서 조정
function seedData(){
  return [
    ["건강한공간","건강/뷰티","배곧동","경기 시흥시 서울대학로264번길 44-4 (배곧동, 에스프라자) 4층","2026-01-06","이대규 외 1명","01062974254","시루 가입","스타터","A"],
    ["강짚탄 배곧본점","음식점","배곧동","경기 시흥시 서울대학로264번길 7 A-107호A-108호 퍼플동(강짚탄)","2026-01-23","홍승희","01082535707","상인회","스타터","B"],
    ["스퀘어마블스 배곧케밥","음식점","배곧동","경기 시흥시 서울대학로264번길 7 에이동-149호","2026-01-23","이동현","01094521119","상인회","스타터","B"],
    ["카페테리아","카페/제과","배곧동","경기 시흥시 서울대학로 172-20 (배곧동, 한라비발디캠퍼스3차)","2026-01-27","조대한","01029762500","아파트 내 카페","비활성","C"],
    ["숨뷰티 배곧점","건강/뷰티","배곧동","경기 시흥시 배곧동 218 아브뉴프랑 그린동 405호","2026-01-28","심지윤","01041109257","상인회","스타터","A"],
    ["배곧 밸런스바레 번지스튜디오","건강/뷰티","배곧동","경기 시흥시 서울대학로264번길 7 아브뉴프랑 퍼플B동 331호 밸런스 스튜디오","2026-01-28","황운식","01098448253","상인회","스타터","A"],
    ["베베드피노 배곧신도시점","생활","배곧동","경기 시흥시 서울대학로278번길 26 116호","2026-02-12","한무비","01040439161","상인회","스타터","C"],
    ["신의주찹쌀순대 배곧점","음식점","배곧동","경기 시흥시 배곧동 200-2 111호 신의주","2026-03-18","박희철","01090000410","시루 가입","스타터","B"],
    ["플로에브","생활","배곧동","경기 시흥시 배곧4로 32-19 sk월드프라자 102호","2026-04-07","송지현","01064121587","시루 가입","스타터","C"],
    ["금홍","음식점","배곧동","경기 시흥시 배곧4로 32-17 로자벨202호 금홍","2026-04-07","이병철","01092487155","시루 가입","스타터","B"],
    ["프레이즈","카페/제과","배곧동","경기 시흥시 서울대학로278번길 19-13 상가A동 158호 프레이즈","2026-04-10","정욱진","01077778339","시루 가입","스타터","B"],
    ["엘리","카페/제과","배곧동","경기 시흥시 서울대학로278번길 8 1층 B동 116호(시흥배곧아브뉴프랑 센트럴 레드)","2026-04-13","유저DD35","01085582203","시루 가입","스타터","B"],
    ["메가톤피자","음식점","배곧동","경기 시흥시 서울대학로264번길 12 상가 B동 157호 메가톤피자","2026-04-13","조민식","01055687529","시루 가입","스타터","B"],
    ["일품양평해장국시흥배곧","음식점","배곧동","경기 시흥시 배곧4로 22 124호","2026-04-14","조래일","01022026013","시루 가입","스타터","B"],
    ["감자바우옹심이","음식점","배곧동","경기 시흥시 배곧4로 22 112호","2026-04-14","김민애","01036911679","카페 추천","스타터","B"],
    ["백년상회","음식점","배곧동","경기 시흥시 서울대학로264번길 26-18 120,121호 백년상회","2026-04-14","윤송현","01091304113","시루 가입","스타터","B"],
    ["발레더블유학원","건강/뷰티","배곧동","경기 시흥시 배곧3로 86 505호","2026-04-17","곽예진","01088630651","시루 가입","스타터","A"],
    ["코어랩필라테스","건강/뷰티","배곧동","경기 시흥시 배곧5로 66-11 4층 코어랩필라테스","2026-04-21","백세리","01053853056","시루 가입","스타터","A"],
    ["대판오징어 배곧점","음식점","배곧동","경기 시흥시 서울대학로278번길 19-8 17 18호","2026-04-27",".","01098727214","방문영업","스타터","B"],
    ["본죽앤비빔밥시흥배곧점","음식점","배곧동","경기 시흥시 서울대학로278번길 19-8 115호","2026-04-27","계화","01027746263","방문영업","스타터","B"],
    ["냥카츠","음식점","배곧동","경기 시흥시 배곧3로 80 1층 냥카츠","2026-04-28","박찬호","01054443054","방문영업","스타터","B"],
    ["청년피자배곧신도시점","음식점","배곧동","경기 시흥시 서울대학로278번길 19-14 119호 청년피자","2026-04-28","Kunhyok 짱","01031493687","방문영업","스타터","B"],
    ["광안천지식당","음식점","배곧동","경기 시흥시 서울대학로278번길 19-8 110호 광안천지식당","2026-04-28","mine","01050464484","방문영업","스타터","B"],
    ["깨떡이네","음식점","배곧동","경기 시흥시 서울대학로278번길 25-24 113호깨떡이네","2026-04-28","박예지","01053139428","방문영업","스타터","B"],
    ["더종로빈대떡","음식점","배곧동","경기 시흥시 배곧3로 80 123호","2026-04-28","이용만","01083241372","방문영업","스타터","B"],
    ["보령식당","음식점","배곧동","경기 시흥시 서울대학로278번길 25-24 102호","2026-04-28","임수민","01093983921","방문영업","스타터","B"],
    ["족발야시장배곧신도시점","음식점","배곧동","경기 시흥시 서울대학로278번길 25-32 107호","2026-04-28","박정운","01094781426","방문영업","스타터","B"],
    ["나답게","생활","배곧동","경기 시흥시 서울대학로278번길 34 126호","2026-04-28","은 경","01058585584","방문영업","스타터","C"],
    ["록갈비 배곧신도시점","음식점","배곧동","경기 시흥시 서울대학로278번길 26 135,135호 록갈비","2026-04-28","유저8A3B","01083313510","방문영업","스타터","B"],
    ["올데이커피365","카페/제과","배곧동","경기 시흥시 서울대학로278번길 26 111호","2026-04-28","양태준","01091859133","방문영업","스타터","B"],
    ["노모어피자배곧신도시점","음식점","배곧동","경기 시흥시 서울대학로278번길 26 1층 103호","2026-04-28","강우영","01084259406","방문영업","스타터","B"],
    ["모모라멘","음식점","배곧동","경기 시흥시 서울대학로264번길 25 112호","2026-04-28","모모라멘","01095504895","방문영업","스타터","B"],
    ["야키토리 동작 배곧점","음식점","배곧동","경기 시흥시 서울대학로278번길 26 133호","2026-04-28","정회민","01093471931","방문영업","스타터","B"],
    ["사이공 본가 쌀국수배곧점","음식점","배곧동","경기 시흥시 서울대학로264번길 25 110호","2026-04-29","유저KG7E","01035106779","방문영업","스타터","B"],
    ["마루상","음식점","배곧동","경기 시흥시 서울대학로264번길 25 블루동 136호","2026-04-29","김 호","01099956778","방문영업","스타터","B"],
    ["플러스82 배곧점","음식점","배곧동","경기 시흥시 배곧동 233 2차 비동 124호","2026-04-29","강서연","01034595474","방문영업","스타터","B"],
    ["24카페폴","카페/제과","배곧동","경기 시흥시 서울대학로278번길 34 옐로우동 1층 124,125호","2026-04-29",".","01082032564","방문영업","스타터","B"],
    ["가유카페","카페/제과","배곧동","경기 시흥시 서울대학로264번길 50 A동 113호 가유카페","2026-04-29","장경진","01098490116","방문영업","스타터","B"],
    ["치킨왕김닭구 배곧2호점","음식점","배곧동","경기 시흥시 서울대학로264번길 50 에이동 123호,124호","2026-04-29","김닭구","01067480607","방문영업","스타터","B"],
    ["소소반점","음식점","배곧동","경기 시흥시 서울대학로278번길 19-8 109호","2026-04-29","정선희","01077075022","방문영업","스타터","B"],
    ["명객중화요리","음식점","배곧동","경기 시흥시 서울대학로278번길 21 101호","2026-04-29","임경만","01084368118","방문영업","스타터","B"],
    ["고향옥얼큰순대국","음식점","배곧동","경기 시흥시 서울대학로278번길 21 106호","2026-04-29","위혜진","01022967264","방문영업","스타터","B"],
    ["샤오마라","음식점","배곧동","경기 시흥시 서울대학로278번길 61 1층S8호","2026-04-29","김정호","01040670910","방문영업","스타터","B"],
    ["CARA 카라","생활","배곧동","경기 시흥시 서울대학로278번길 61 1층 T-5호 1층 카라(CARA) 배곧동, 서영베니스스퀘어","2026-04-29","김소림","01037088643","방문영업","스타터","C"],
    ["코지먼트배곧","생활","배곧동","경기 시흥시 서울대학로278번길 61 a-27","2026-04-29","강해","01046869791","방문영업","스타터","C"],
    ["피어싱갤러리","생활","배곧동","경기 시흥시 서울대학로278번길 61 베니스스퀘어 C28호 1층","2026-04-29","김일남","01022036302","방문영업","스타터","C"],
    ["구구바베큐","음식점","배곧동","경기 시흥시 서울대학로278번길 70 B동 103호","2026-04-30","이순화","01055154583","방문영업","스타터","B"],
    ["감성커피 시흥배곧점","카페/제과","배곧동","경기 시흥시 서울대학로278번길 61 s-24호","2026-04-30","신형술","01026106307","방문영업","스타터","B"],
    ["죽이야기 배곧점","음식점","배곧동","경기 시흥시 배곧4로 95 지음프라자 106호 죽이야기 배곧점","2026-05-11","박진주","01083454601","방문영업","스타터","B"],
    ["부대찌개대사관배곧서울대학로점","음식점","배곧동","경기 시흥시 서울대학로 59-69 1층123호","2026-05-22","장연정","01071992258","시루 가입","스타터","B"]
  ];
}

function getSheet_(){
  var ss = SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  if(!ss){ throw new Error("스프레드시트를 찾을 수 없습니다. 상단 SPREADSHEET_ID에 구글시트 ID를 넣거나, 구글시트에서 확장 프로그램→Apps Script로 열어 실행하세요."); }
  var sh = ss.getSheetByName(SHEET_NAME);
  if(!sh){ sh = ss.insertSheet(SHEET_NAME); }
  return sh;
}

function initSheet(){
  var sh = getSheet_();
  sh.clear();
  sh.getRange(1,1,1,HEADERS.length).setValues([HEADERS]).setFontWeight("bold");
  // seed 열: [매장명,업종,동,주소,입점일,점주명,연락처,인입경로,요금제,등급]
  var rows = seedData().map(function(r,i){
    return [i+1, r[0], r[1], r[2], r[3], r[4], r[5], r[6], r[7], r[8], r[9], "미착수", "", "", "", "", ""];
  });
  // 연락처 열을 텍스트 서식으로 지정 → 앞자리 0 보존
  var phoneCol = HEADERS.indexOf("연락처")+1;
  sh.getRange(2, phoneCol, rows.length, 1).setNumberFormat("@");
  sh.getRange(2,1,rows.length,HEADERS.length).setValues(rows);
  sh.setFrozenRows(1);
  // 상태·등급 드롭다운
  var stCol = HEADERS.indexOf("상태")+1, grCol = HEADERS.indexOf("등급")+1;
  sh.getRange(2,stCol,rows.length,1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(STATUSES, true).build());
  sh.getRange(2,grCol,rows.length,1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(["A","B","C"], true).build());
  Logger.log("초기화 완료: " + rows.length + "곳");
}

function readAll_(){
  var sh = getSheet_();
  var values = sh.getDataRange().getValues();
  var head = values.shift();
  var rows = values.filter(function(r){ return r[0] !== "" && r[0] !== null; }).map(function(r){
    var o = {};
    for(var i=0;i<head.length;i++){ o[head[i]] = r[i]; }
    return o;
  });
  return rows;
}

function json_(obj){
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e){
  return json_({ ok:true, statuses:STATUSES, rows:readAll_(), ts:new Date().getTime() });
}

function doPost(e){
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try{
    var body = JSON.parse(e.postData.contents);
    var sh = getSheet_();
    var data = sh.getDataRange().getValues();
    var head = data[0];
    if(body.action === "reset"){ initSheet(); return json_({ok:true, reset:true}); }
    // find row by 번호
    var colNo = head.indexOf("번호");
    var target = -1;
    for(var i=1;i<data.length;i++){ if(String(data[i][colNo]) === String(body.no)){ target = i+1; break; } }
    if(target < 0) return json_({ok:false, err:"row not found: "+body.no});
    // update one field
    var col = head.indexOf(body.field);
    if(col < 0) return json_({ok:false, err:"field not found: "+body.field});
    sh.getRange(target, col+1).setValue(body.value);
    // stamp 수정시각
    var tcol = head.indexOf("수정시각");
    if(tcol>=0) sh.getRange(target, tcol+1).setValue(Utilities.formatDate(new Date(),"Asia/Seoul","MM-dd HH:mm"));
    return json_({ok:true, no:body.no, field:body.field, value:body.value});
  } catch(err){
    return json_({ok:false, err:String(err)});
  } finally {
    lock.releaseLock();
  }
}


/* =========================================================================
 * 일일 자동 보고 (구글 서버 트리거 → 노션) — PC 꺼져 있어도 실행됨
 * 세팅:
 *  1) https://www.notion.so/my-integrations 에서 "새 통합"(내부용) 생성 → 토큰(ntn_.../secret_...) 복사
 *  2) 노션 "동네해빙 일일 전환율 보고" 페이지 → 우측상단 ••• → 연결(Connections) → 방금 만든 통합 추가
 *  3) 아래 NOTION_TOKEN 에 토큰 붙여넣기 → 저장
 *  4) 함수목록에서 createDailyTrigger 실행(최초 1회, 권한 승인) → 매일 08시 트리거 생성
 *  5) (테스트) dailyReport 한 번 실행해 노션에 오늘 보고가 붙는지 확인
 * ========================================================================= */
var NOTION_TOKEN = "";  // ★ 노션 통합 토큰 붙여넣기
var NOTION_PAGE_ID = "390d244ade3c81008d92daa2dde0b5fd";  // 일일 보고 페이지 ID

function computeStats_(){
  var sh = getSheet_();
  var values = sh.getDataRange().getValues();
  var head = values.shift(); // 헤더
  var gi = head.indexOf("등급"), si = head.indexOf("상태");
  var rows = values.filter(function(r){ return r[0] !== "" && r[0] !== null; });
  function cnt(s){ var c=0; rows.forEach(function(r){ if(r[si]===s) c++; }); return c; }
  var n=rows.length, waiting=cnt("미착수"), buja=cnt("부재");
  var eff=n-waiting-buja, conv=cnt("전환")+cnt("계약체결");
  var rate=eff>0 ? Math.round(conv/eff*1000)/10 : 0;
  var aTot=0, aConv=0;
  rows.forEach(function(r){
    if(r[gi]==="A"){ aTot++; if(r[si]==="상담접수"||r[si]==="전환"||r[si]==="계약체결") aConv++; }
  });
  return {n:n,waiting:waiting,buja:buja,reject:cnt("유선거절"),sangdam:cnt("상담접수"),
          conv:conv,gyeyak:cnt("계약체결"),eff:eff,rate:rate,aTot:aTot,aConv:aConv};
}

function dailyReport(){
  if(!NOTION_TOKEN){ throw new Error("NOTION_TOKEN을 채워주세요 (노션 통합 토큰)."); }
  var s = computeStats_();
  var today = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd");
  function bullet(txt, bold){
    return {object:"block", type:"bulleted_list_item",
      bulleted_list_item:{ rich_text:[{ type:"text", text:{content:txt}, annotations:{bold: !!bold} }] }};
  }
  var children = [
    {object:"block", type:"heading_2", heading_2:{ rich_text:[{ type:"text", text:{content:"📅 "+today} }] }},
    bullet("전체 "+s.n+" / 미착수 "+s.waiting),
    bullet("부재 "+s.buja+" · 유선거절 "+s.reject),
    bullet("상담접수(동의) "+s.sangdam+"곳"),
    bullet("전환(컨설팅DB) "+s.conv+"곳 / 계약 "+s.gyeyak+"곳"),
    bullet("전환율 "+s.rate+"% (유효접촉 "+s.eff+") · A급 진척 "+s.aConv+"/"+s.aTot, true),
    {object:"block", type:"divider", divider:{}}
  ];
  var res = UrlFetchApp.fetch("https://api.notion.com/v1/blocks/"+NOTION_PAGE_ID+"/children", {
    method:"patch", contentType:"application/json",
    headers:{ "Authorization":"Bearer "+NOTION_TOKEN, "Notion-Version":"2022-06-28" },
    payload: JSON.stringify({ children: children }),
    muteHttpExceptions: true
  });
  Logger.log(res.getResponseCode()+" "+res.getContentText());
  if(res.getResponseCode() >= 300) throw new Error("노션 기록 실패: "+res.getContentText());
}

function createDailyTrigger(){
  ScriptApp.getProjectTriggers().forEach(function(t){
    if(t.getHandlerFunction()==="dailyReport") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("dailyReport").timeBased().atHour(8).everyDays(1).inTimezone("Asia/Seoul").create();
  Logger.log("매일 08시(KST) dailyReport 트리거 생성 완료");
}
