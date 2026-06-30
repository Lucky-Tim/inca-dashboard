// ═══════════════════════════════════════════════════════════
// 인카금융 운영 플랫폼 — Google Apps Script v6 (사전집계 캐시)
// ─────────────────────────────────────────────────────────
// 설정 방법:
//   1. https://script.google.com 접속 → 기존 프로젝트 열기
//   2. 이 코드 전체 붙여넣기 (Ctrl+A → Ctrl+V)
//   3. 저장(Ctrl+S)
//   4. 배포 → 기존 배포 관리 → ✏️ 편집 → "새 버전" → 배포
// ═══════════════════════════════════════════════════════════

const SYNC_KEY    = 'INCA2026';
const SS_PROP     = 'SPREADSHEET_ID';
const CACHE_PROP  = 'CACHE_FILE_ID';   // Drive JSON 캐시 파일 ID
const CACHE_NAME  = 'inca_cache.json'; // 캐시 파일명

// ─── 데이터 소스 레지스트리 ──────────────────────────────
const DATA_SOURCES = [
  { key: 'employees',   sheetName: '직원관리',  parser: parseEmpRows  },
  { key: 'performance', sheetName: '건별실적',  parser: parsePerfRows },
];

// ─── 스프레드시트 자동 생성/연결 ────────────────────────
function getSpreadsheet() {
  const props = PropertiesService.getScriptProperties();
  let ssId = props.getProperty(SS_PROP);
  if (ssId) {
    try { return SpreadsheetApp.openById(ssId); } catch(e) {}
  }
  const ss = SpreadsheetApp.create('인카_운영데이터');
  props.setProperty(SS_PROP, ss.getId());
  return ss;
}

// ─── Drive JSON 캐시 읽기 ────────────────────────────────
function readCache() {
  try {
    const props  = PropertiesService.getScriptProperties();
    const fileId = props.getProperty(CACHE_PROP);
    if (!fileId) return null;
    const file    = DriveApp.getFileById(fileId);
    const content = file.getBlob().getDataAsString();
    return JSON.parse(content);
  } catch(e) { return null; }
}

// ─── Drive JSON 캐시 쓰기 ────────────────────────────────
function writeCache(data) {
  try {
    const props   = PropertiesService.getScriptProperties();
    const content = JSON.stringify(data);
    let fileId    = props.getProperty(CACHE_PROP);
    if (fileId) {
      try {
        // 기존 파일 덮어쓰기
        DriveApp.getFileById(fileId).setContent(content);
        return;
      } catch(e) { /* 파일 없으면 새로 생성 */ }
    }
    // 새 파일 생성
    const file = DriveApp.createFile(CACHE_NAME, content, 'application/json');
    props.setProperty(CACHE_PROP, file.getId());
  } catch(e) {
    Logger.log('캐시 쓰기 실패: ' + e.message);
  }
}

// ─── GET: 캐시 우선 반환 → 없으면 Sheets에서 읽기 ────────
function doGet(e) {
  try {
    // 1) Drive 캐시 확인
    const cached = readCache();
    if (cached && cached.ok && cached.employees?.length) {
      // slim=1: performance 원본 제외 (뷰어 전용 — 로딩 최적화)
      if (e && e.parameter && e.parameter.slim === '1') {
        const slim = JSON.parse(JSON.stringify(cached));
        delete slim.performance;
        return json(slim);
      }
      return json(cached);
    }

    // 2) 캐시 없으면 Sheets에서 직접 읽고 캐시 갱신
    return readSheetsAndCache();
  } catch(err) {
    return json({ ok: false, error: err.message });
  }
}

// ─── Sheets 전체 읽기 + 캐시 저장 ───────────────────────
function readSheetsAndCache() {
  const ss     = getSpreadsheet();
  const result = { ok: true, lastUpdated: getMeta(ss, 'lastUpdated') };

  DATA_SOURCES.forEach(src => {
    const sheet = ss.getSheetByName(src.sheetName);
    result[src.key] = sheet ? src.parser(sheetToObjects(sheet)) : [];
  });

  if (!result.employees.length && !result.performance.length) {
    return json({ ok: false, error: '데이터 없음 — Sheets에 직원관리·건별실적 탭을 확인하세요' });
  }

  // 사전집계 생성
  result.aggregated = buildAggregated(result.performance, result.employees);
  const topResult   = buildTop10(result.performance, result.employees);
  result.top10      = topResult.top10;
  result.topByBranch = topResult.topByBranch;

  // 캐시 저장 (다음 GET부터 빠르게 응답)
  writeCache(result);
  return json(result);
}

// ─── POST: 데이터 저장 + 캐시 즉시 갱신 ─────────────────
// ※ 건별실적은 월별 교체 — 보낸 달만 덮어쓰고 나머지 달은 유지
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    if (payload.key !== SYNC_KEY) return json({ ok: false, error: '인증 실패' });

    const ss = getSpreadsheet();

    // ── 직원관리 저장 (전체 교체, 월 개념 없음) ──────────────
    if (payload.employees?.length) {
      const sheet = getSheet(ss, '직원관리');
      sheet.clearContents();
      const rows = [['id','name','type','b1','b2','b3','joinDate']];
      payload.employees.forEach(emp =>
        rows.push([emp.id||'', emp.name||'', emp.type||'', emp.b1||'', emp.b2||'', emp.b3||'', emp.joinDate||''])
      );
      sheet.getRange(1, 1, rows.length, 7).setValues(rows);
    }

    // ── 건별실적 저장 (월별 교체) ────────────────────────────
    let allPerf  = [];
    let newMonths = new Set();

    if (payload.performance?.length) {
      const sheet = getSheet(ss, '건별실적');
      newMonths   = new Set(payload.performance.map(r => String(r.month || '').trim()));

      // 기존 데이터 읽기 → 새 월과 겹치는 행만 제거, 나머지 달은 보존
      const existing = sheetToObjects(sheet);
      const kept     = existing.filter(r => !newMonths.has(String(r.month || '').trim()));

      // 보존된 기존 데이터 + 신규 데이터 합치기
      allPerf = [...kept, ...payload.performance];

      // 전체 다시 쓰기 (500행 청크)
      const headers = [['SUNAB_PK','증권번호','month','agentId','premium','credit','category','company','payMethod']];
      sheet.clearContents();
      sheet.getRange(1, 1, 1, 9).setValues(headers);
      const CHUNK = 500;
      for (let i = 0; i < allPerf.length; i += CHUNK) {
        const chunk = allPerf.slice(i, i + CHUNK).map(r => [
          r.SUNAB_PK||'', r.증권번호||'', r.month||'', r.agentId||'',
          r.premium||0, r.credit||0, r.category||'', r.company||'', r.payMethod||''
        ]);
        sheet.getRange(i + 2, 1, chunk.length, 9).setValues(chunk);
      }
    } else {
      // performance 미전송 시 기존 Sheets 데이터 그대로 사용
      const sheet = ss.getSheetByName('건별실적');
      allPerf = sheet ? parsePerfRows(sheetToObjects(sheet)) : [];
    }

    const now = new Date().toISOString();
    setMeta(ss, 'lastUpdated', now);
    setMeta(ss, 'empCount',    payload.employees?.length || 0);
    setMeta(ss, 'perfCount',   allPerf.length);

    // 사전집계는 전체(merged) 기준으로 생성
    const employees  = payload.employees || [];
    const perfForAgg = allPerf.map(r => ({...r, status:'실제데이터'}));
    const aggData    = buildAggregated(perfForAgg, employees);
    const topResult  = buildTop10(perfForAgg, employees);

    // 캐시 즉시 갱신
    writeCache({
      ok:           true,
      lastUpdated:  now,
      employees,
      performance:  perfForAgg,
      aggregated:   aggData,
      top10:        topResult.top10,
      topByBranch:  topResult.topByBranch,
    });

    return json({
      ok:        true,
      empCount:  employees.length,
      perfCount: allPerf.length,
      updatedMonths: [...newMonths],
      aggCount:  aggData.length,
    });
  } catch(err) {
    return json({ ok: false, error: err.message });
  }
}

// ─── 사전집계 생성 (월×본부×조직3×납입방법) ──────────────
// 뷰어가 수백 rows로 KPI를 즉시 집계할 수 있도록
function buildAggregated(performance, employees) {
  // employee id → { branch(b2), branch3(b3) } 맵
  const empMap = {};
  (employees || []).forEach(e => {
    empMap[String(e.id)] = {
      branch:  (e.b2 || '').trim(),
      branch3: (e.b3 || '').trim(),
    };
  });

  const map = {};
  (performance || []).forEach(r => {
    const month   = String(r.month  || '').trim();
    const emp     = empMap[String(r.agentId)] || {};
    const branch  = (r.branch2 || emp.branch  || '').trim();
    const branch3 = (r.branch3 || emp.branch3 || '').trim();
    const pay     = (r.payMethod || '').trim();
    const key     = month + '|' + branch + '|' + branch3 + '|' + pay;

    if (!map[key]) {
      map[key] = {
        month, branch, branch3, payMethod: pay,
        premium: 0, credit: 0, count: 0,
        activeSet: {}   // agentId → true (Set 대신 객체로 직렬화 회피)
      };
    }
    const v = map[key];
    const pre = Number(r.premium) || 0;
    const cre = Number(r.credit)  || 0;
    v.premium += pre;
    v.credit  += cre;
    v.count   += 1;
    if (pre > 0 && r.agentId) {
      v.activeSet[String(r.agentId)] = true;
    }
  });

  return Object.values(map).map(v => ({
    month:        v.month,
    branch:       v.branch,
    branch3:      v.branch3,
    payMethod:    v.payMethod,
    premium:      v.premium,
    credit:       v.credit,
    count:        v.count,
    activeAgents: Object.keys(v.activeSet).length,
  }));
}

// ─── TOP10 생성 (당월 제외 직전 6개월 완전마감 기준) ──────
// 개인별 6개월 월평균 실적 상위 10명, 필터 무관
function buildTop10(performance, employees) {
  // 현재 월 (YYYYMM)
  const now = new Date();
  const currM = String(now.getFullYear()) + String(now.getMonth() + 1).padStart(2, '0');

  // 모든 유니크 월 목록 → 당월 제외 → 최근 6개월
  const allMonths = [];
  const monthSet  = {};
  (performance || []).forEach(r => {
    const m = String(r.month || '').trim();
    if (m && !monthSet[m]) { monthSet[m] = true; allMonths.push(m); }
  });
  allMonths.sort();

  const closedMonths  = allMonths.filter(m => m < currM);
  const targetMonths  = new Set(closedMonths.slice(-6));
  const n             = targetMonths.size;

  if (n === 0) return [];

  // employee map: id → { name, branch(b2), branch3(b3) }
  const empMap = {};
  (employees || []).forEach(e => {
    empMap[String(e.id)] = {
      name:    (e.name || '').trim(),
      branch:  (e.b2   || '').trim(),
      branch3: (e.b3   || '').trim(),
    };
  });

  // agentId별 해당 기간 총 premium 집계 (납입방법 무관, 전 본부)
  const agentData = {};
  (performance || []).forEach(r => {
    const m = String(r.month || '').trim();
    if (!targetMonths.has(m)) return;
    const id  = String(r.agentId || '');
    const pre = Number(r.premium) || 0;
    if (!id || pre <= 0) return;
    if (!agentData[id]) agentData[id] = 0;
    agentData[id] += pre;
  });

  // 개인별 avgPremium (6개월 총합 / n)
  const list = Object.entries(agentData).map(([id, total]) => {
    const emp = empMap[id] || { name: id, branch: '' };
    return {
      agentId:      id,
      name:         emp.name,
      branch:       emp.branch,
      branch3:      emp.branch3,
      totalPremium: total,
      avgPremium:   Math.round(total / n),
      months:       n,
    };
  });

  list.sort((a, b) => b.avgPremium - a.avgPremium);

  // 전체 TOP10
  const top10 = list.slice(0, 10).map((v, i) => ({ rank: i + 1, ...v }));

  // 소속사업단(branch)별 TOP10 — 이미 전체 정렬된 list에서 분기
  const byBranch = {};
  list.forEach(agent => {
    const b = agent.branch || '미분류';
    if (!byBranch[b]) byBranch[b] = [];
    byBranch[b].push(agent);
  });
  const topByBranch = {};
  Object.keys(byBranch).forEach(b => {
    topByBranch[b] = byBranch[b].slice(0, 10).map((v, i) => ({ rank: i + 1, ...v }));
  });

  return { top10, topByBranch };
}

// ─── 파서 함수 ───────────────────────────────────────────
function parseEmpRows(rows) {
  return rows.map(e => ({ ...e, _source: 'real' }));
}

function parsePerfRows(rows) {
  return rows.map(r => ({
    ...r,
    premium: Number(r.premium) || 0,
    credit:  Number(r.credit)  || 0,
    status:  '실제데이터',
  }));
}

// ─── 유틸 ────────────────────────────────────────────────
function getSheet(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function sheetToObjects(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0].map(h => String(h).trim());
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { if (h) obj[h] = row[i]; });
    return obj;
  });
}

function getMeta(ss, key) {
  try {
    const sheet = ss.getSheetByName('메타');
    if (!sheet) return null;
    const row = sheet.getDataRange().getValues().find(r => r[0] === key);
    return row ? row[1] : null;
  } catch(e) { return null; }
}

function setMeta(ss, key, value) {
  const sheet = getSheet(ss, '메타');
  const rows  = sheet.getDataRange().getValues();
  for (let i = 0; i < rows.length; i++) {
    if (rows[i][0] === key) { sheet.getRange(i+1, 2).setValue(value); return; }
  }
  sheet.appendRow([key, value]);
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─── 캐시 수동 갱신 (최초 1회 또는 Sheets 직접 수정 후) ──
// GAS 편집기에서 이 함수를 실행하면 캐시가 새로 만들어짐
function refreshCache() {
  const result = readSheetsAndCache();
  Logger.log('캐시 갱신 완료: ' + result.getContent().slice(0, 100));
}

// ─── WARMUP (콜드 스타트 방지) ───────────────────────────
function warmup() {
  Logger.log('warmup ok: ' + new Date().toISOString());
}

function setupWarmupTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  const already = triggers.some(t => t.getHandlerFunction() === 'warmup');
  if (already) { Logger.log('이미 트리거 등록됨'); return; }
  ScriptApp.newTrigger('warmup')
    .timeBased()
    .everyMinutes(10)
    .create();
  Logger.log('warmup 트리거 등록 완료 (10분 간격)');
}
