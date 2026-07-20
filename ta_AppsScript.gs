/**
 * 동네해빙 운영OS — TA·개척 트래커 Apps Script 백엔드
 * 원본 '업셀링 DB' 시트에서 매장을 가져와(import) 운영OS 워크플로 컬럼을 붙인다.
 * 구조: Google Sheet + Apps Script + HTML (React/서버 없음). 컬럼은 미래 단계까지 미리 확장.
 *
 * 사용법:
 *  1) 새 구글시트 → 확장 프로그램 → Apps Script → 이 코드 전체 붙여넣기 → 저장
 *  2) importFromSource 실행(최초 1회, 권한 승인) → 원본 125곳 가져옴
 *  3) 배포 → 새 배포 → 웹 앱 → 실행:나 / 액세스:모든 사용자 → /exec URL 복사
 *  4) ta.html 의 API_URL 에 붙여넣고 git push
 *  (원본 갱신 시 importMergeFromSource 실행 → 진행상태 유지, 신규매장만 추가)
 */

var SOURCE_ID = "1ewvEx1GdEzhVsIdbymxevSGbdusimLNcW0PQEvqQamw"; // 원본 업셀링 DB 파일ID
var SHEET_NAME = "TA개척";

// 운영OS 스키마 (미래 단계 컬럼 미리 포함 — 좌표/예약/방문결과)
var HEADERS = ["번호","동네","업종","등급","가게명","점주명","연락처","주소","지도주소","위도","경도",
  "담당","TA상태","시도횟수","다음접촉일","예약일","예약시간","예약확정","방문일시","방문결과",
  "컨설턴트배정","메모","등록일","수정시각"];

// 상태 파이프라인 + 분기(부재/보류/거절/불가)
var STATES = ["미배정","담당배정","방문예정","방문완료","예약확정","컨설팅완료","계약진행","계약완료","사후관리",
  "부재","보류","거절","불가"];

function getSheet_(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if(!sh) sh = ss.insertSheet(SHEET_NAME);
  return sh;
}

function shortDong_(s){ s=String(s||"").trim(); var p=s.split(/\s+/); return p[p.length-1]||s; }
function normTel_(t){ t=String(t||"").replace(/[^0-9]/g,""); if(/^\d{9,10}$/.test(t)) t="0"+t; return t; }

// 업종 → 등급 초안 (고객단가/보험가치 기준): 건강뷰티=A, 음식/카페=B, 생활/여가=C
function gradeOf_(u){
  u=String(u||"");
  if(/건강|뷰티|필라|PT|병원|학원/.test(u)) return "A";
  if(/음식|카페|제과|식당|주점/.test(u)) return "B";
  return "C";
}

// 원본 비고/TA여부 → 초기 상태 추정 (운영OS 상태로 매핑)
function guessState_(bigo, ta, agent){
  var b=String(bigo||"")+" "+String(ta||"");
  if(/확정|만남|방문예정/.test(b)) return "예약확정";
  if(/폐업|결번|휴점/.test(b)) return "불가";
  if(/거절/.test(b)) return "거절";
  if(/부재/.test(b)) return "부재";
  if(/보류|재통화|재콜|확인|다시\s*연락|다음\s*주|내일/.test(b)) return "보류";
  if(/완료/.test(String(ta||""))) return "방문완료";
  return String(agent||"").trim() ? "담당배정" : "미배정";
}
function countAttempts_(ta){
  var m=String(ta||"").match(/\d{6}|\d{1,2}\/\d{1,2}/g);
  if(m) return m.length;
  return String(ta||"").trim()? 1 : 0;
}

function readSource_(){
  var src = SpreadsheetApp.openById(SOURCE_ID);
  var sheets = src.getSheets();
  for(var s=0;s<sheets.length;s++){
    var vals = sheets[s].getDataRange().getValues();
    for(var i=0;i<Math.min(vals.length,12);i++){
      if(vals[i].indexOf("가게명")>=0){ return { head: vals[i], vals: vals, hi: i }; }
    }
  }
  throw new Error("원본에서 '가게명' 헤더를 찾지 못했습니다. SOURCE_ID를 확인하세요.");
}

function buildRow_(head, row, n){
  function v(name){ var c=head.indexOf(name); return c>=0? row[c] : ""; }
  var st = guessState_(v("비고"), v("TA 여부"), v("TA 진행자"));
  var visit = (st==="예약확정") ? (String(v("일정")||"")+" "+String(v("비고")||"")).trim() : "";
  return [ n, shortDong_(v("동네")), v("업종"), gradeOf_(v("업종")), String(v("가게명")).trim(), v("점주명"),
    normTel_(v("점주 연락처")), v("주소"), v("지도용주소")||v("주소"), "", "",
    v("TA 진행자"), st, countAttempts_(v("TA 여부")), "", "", "", "", visit, "",
    "", String(v("비고")||""), v("등록일"), "" ];
}

function importFromSource(){
  var src = readSource_();
  var out = [HEADERS]; var n=0;
  for(var r=src.hi+1;r<src.vals.length;r++){
    var name = String(src.vals[r][src.head.indexOf("가게명")]||"").trim();
    if(!name || name==="가게명") continue;
    n++; out.push(buildRow_(src.head, src.vals[r], n));
  }
  var sh=getSheet_(); sh.clear();
  sh.getRange(1,1,out.length,HEADERS.length).setValues(out);
  finalize_(sh, out.length);
  Logger.log("가져오기 완료: "+(out.length-1)+"곳");
}

function importMergeFromSource(){
  var sh=getSheet_();
  var cur=sh.getDataRange().getValues();
  if(cur.length<2 || cur[0].indexOf("가게명")<0){ importFromSource(); return; }
  var chead=cur[0];
  var iName=chead.indexOf("가게명"), iAddr=chead.indexOf("주소"), iNo=chead.indexOf("번호");
  var seen={}, maxNo=0;
  for(var i=1;i<cur.length;i++){
    seen[(cur[i][iName]||"")+"|"+(cur[i][iAddr]||"")]=true;
    maxNo=Math.max(maxNo, Number(cur[i][iNo])||0);
  }
  var src=readSource_(); var add=[]; var n=maxNo;
  for(var r=src.hi+1;r<src.vals.length;r++){
    var nm=String(src.vals[r][src.head.indexOf("가게명")]||"").trim(); if(!nm||nm==="가게명") continue;
    var ad=src.vals[r][src.head.indexOf("주소")]||"";
    if(seen[nm+"|"+ad]) continue;
    n++; add.push(buildRow_(src.head, src.vals[r], n));
  }
  if(add.length) sh.getRange(sh.getLastRow()+1,1,add.length,HEADERS.length).setValues(add);
  finalize_(sh, sh.getLastRow());
  Logger.log("신규 추가: "+add.length+"곳");
}

function finalize_(sh, lastRow){
  sh.setFrozenRows(1);
  sh.getRange(1,1,1,HEADERS.length).setFontWeight("bold");
  var rows = lastRow-1; if(rows<1) return;
  sh.getRange(2, HEADERS.indexOf("연락처")+1, rows, 1).setNumberFormat("@");
  sh.getRange(2, HEADERS.indexOf("TA상태")+1, rows, 1)
    .setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(STATES,true).build());
  sh.getRange(2, HEADERS.indexOf("등급")+1, rows, 1)
    .setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(["A","B","C"],true).build());
}

function readAll_(){
  var sh=getSheet_(); var values=sh.getDataRange().getValues();
  var head=values.shift();
  return values.filter(function(r){ return r[0]!=="" && r[0]!==null; }).map(function(r){
    var o={}; for(var i=0;i<head.length;i++){ o[head[i]]=r[i]; } return o;
  });
}
function json_(o){ return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function doGet(e){ return json_({ ok:true, states:STATES, rows:readAll_(), ts:new Date().getTime() }); }

function doPost(e){
  var lock=LockService.getScriptLock(); lock.waitLock(20000);
  try{
    var body=JSON.parse(e.postData.contents);
    var sh=getSheet_(); var data=sh.getDataRange().getValues(); var head=data[0];
    if(body.action==="import"){ importMergeFromSource(); return json_({ok:true, merged:true}); }
    var colNo=head.indexOf("번호"), target=-1;
    for(var i=1;i<data.length;i++){ if(String(data[i][colNo])===String(body.no)){ target=i+1; break; } }
    if(target<0) return json_({ok:false, err:"row not found: "+body.no});
    // 여러 필드 동시 저장 지원 (방문결과 팝업 등) — body.fields = {필드:값}
    var updates = body.fields || {}; if(body.field) updates[body.field]=body.value;
    for(var f in updates){
      var col=head.indexOf(f); if(col<0) continue;
      sh.getRange(target,col+1).setValue(updates[f]);
      if(f==="TA상태" && (updates[f]==="부재"||updates[f]==="보류") && body.bump){
        var ac=head.indexOf("시도횟수");
        if(ac>=0){ var cur=Number(data[target-1][ac])||0; sh.getRange(target,ac+1).setValue(cur+1); }
      }
    }
    var tc=head.indexOf("수정시각");
    if(tc>=0) sh.getRange(target,tc+1).setValue(Utilities.formatDate(new Date(),"Asia/Seoul","MM-dd HH:mm"));
    return json_({ok:true});
  }catch(err){ return json_({ok:false, err:String(err)}); }
  finally{ lock.releaseLock(); }
}
