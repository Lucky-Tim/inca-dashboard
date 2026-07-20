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
const ORG_OVR_PROP = 'ORG_OVR_FILE_ID';    // Drive JSON 조직지정 마스터 파일 ID
const ORG_OVR_NAME = '_org_overrides.json'; // 조직지정 마스터 파일명 (캐시와 같은 폴더)

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

// ─── 조직 지정 마스터(오버라이드) — Drive JSON 별도 파일 ──────
// 사번 → [사업단,본부](구버전 배열) 또는 {b2,b3,name}(신버전 객체). Script Properties는
// 9KB/키 제한이 있어 사원 수백 명 규모에 부적합 → Drive 파일로 별도 관리(캐시 파일과 같은 위치).
function readOrgOverrides() {
  try {
    const props  = PropertiesService.getScriptProperties();
    const fileId = props.getProperty(ORG_OVR_PROP);
    if (!fileId) return {};
    const file = DriveApp.getFileById(fileId);
    const obj  = JSON.parse(file.getBlob().getDataAsString());
    return (obj && typeof obj === 'object') ? obj : {};
  } catch(e) { return {}; }
}

function writeOrgOverrides(data) {
  try {
    const props   = PropertiesService.getScriptProperties();
    const content = JSON.stringify(data);
    let fileId    = props.getProperty(ORG_OVR_PROP);
    if (fileId) {
      try { DriveApp.getFileById(fileId).setContent(content); return; }
      catch(e) { /* 파일이 삭제된 경우 새로 생성 */ }
    }
    const file = DriveApp.createFile(ORG_OVR_NAME, content, 'application/json');
    props.setProperty(ORG_OVR_PROP, file.getId());
  } catch(e) {
    Logger.log('조직 오버라이드 저장 실패: ' + e.message);
  }
}

// 오버라이드 값 정규화: [사업단,본부] 배열(구버전) 또는 {b2,b3,name} 객체(신버전) 둘 다 지원
function _ovrParts(v) {
  if (!v) return null;
  if (Array.isArray(v)) return { b2: v[0]||'', b3: v[1]||'', name: '' };
  return { b2: v.b2||'', b3: v.b3||'', name: v.name||'' };
}

// 직원관리 명단에 오버라이드 적용 (b2/b3만 교체 — 성명은 직원관리 원본 유지)
function applyOrgOverridesToEmployees(employees, overrides) {
  if (!overrides || !Object.keys(overrides).length) return employees || [];
  return (employees || []).map(e => {
    const p = _ovrParts(overrides[String(e.id)]);
    if (!p) return e;
    const out = Object.assign({}, e);
    if (p.b2) out.b2 = p.b2;
    if (p.b3) out.b3 = p.b3;
    return out;
  });
}

// 집계 전용: 직원관리에 없는(실적만 있는) 사번도 오버라이드가 있으면 가상 직원 레코드로 추가
// → buildAggregated/buildTop10/buildTop10ByMonth의 empMap 조회 시 조직이 반영되도록 함
// (실제 result.employees에는 추가하지 않음 — 뷰어의 재직인원 집계를 왜곡하지 않기 위함)
function addOverrideStubs(employees, overrides) {
  const list = (employees || []).slice();
  if (!overrides) return list;
  const haveIds = new Set(list.map(e => String(e.id)));
  Object.keys(overrides).forEach(id => {
    if (haveIds.has(String(id))) return;
    const p = _ovrParts(overrides[id]);
    if (!p || (!p.b2 && !p.b3)) return;
    list.push({ id, name: p.name || '미매칭', type: '미정', b1:'', b2: p.b2||'', b3: p.b3||'', joinDate:'' });
  });
  return list;
}

// ─── GET: 캐시 우선 반환 → 없으면 Sheets에서 읽기 ────────
function doGet(e) {
  try {
    // 0) 강제 캐시 재생성(시트 기준) — 재배포 후 빠른 갱신용. 48k 재업로드 불필요.
    if (e && e.parameter && e.parameter.rebuild === '1' && e.parameter.key === SYNC_KEY) {
      return readSheetsAndCache();
    }
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

  // 미매칭 사원 조직 지정 마스터 적용 (b2/b3 교체) — 재업로드 없이도 대시보드에서 지정한 값이 반영됨
  const orgOverrides = readOrgOverrides();
  result.employees   = applyOrgOverridesToEmployees(result.employees, orgOverrides);
  const empForAgg    = addOverrideStubs(result.employees, orgOverrides); // 실적-only 사원도 집계에 포함

  // 사전집계 생성
  result.aggregated = buildAggregated(result.performance, empForAgg);
  const topResult   = buildTop10(result.performance, empForAgg);
  result.top10        = topResult.top10;
  result.topByBranch  = topResult.topByBranch;
  result.topByBranch3 = topResult.topByBranch3;
  // 뷰어 납입연/월 필터 연동용 — 월별 TOP10 스냅샷(전체/사업단별/본부별)
  result.top10ByMonth = buildTop10ByMonth(result.performance, empForAgg);

  // 월별 마감/가마감 상태
  result.monthStatus = buildMonthStatus(result.performance);

  // 조직 지정 마스터도 함께 내려줌 — 대시보드가 localStorage와 병합해 오프라인에서도 활용
  result.orgOverrides = orgOverrides;

  // 캐시 저장 (다음 GET부터 빠르게 응답)
  writeCache(result);
  return json(result);
}

// ─── POST: 데이터 저장 + 캐시 즉시 갱신 ─────────────────
// ※ 건별실적은 월별 교체 — 보낸 달만 덮어쓰고 나머지 달은 유지
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    // ── 관리자 로그인 (사번+비번 → 소속 데이터만 반환) ──
    if (payload.action === 'login') return handleManagerLogin(payload);
    // ── 미매칭 사원 조직 지정 저장 (대시보드 팝업 → 서버 마스터에 병합) ──
    if (payload.action === 'saveOrgOverrides') return handleSaveOrgOverrides(payload);
    if (payload.key !== SYNC_KEY) return json({ ok: false, error: '인증 실패' });

    const ss = getSpreadsheet();
    const _postSrc = (payload.source||'').toString().trim();  // '마감' | '가마감'

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
      const headers = [['SUNAB_PK','증권번호','month','agentId','premium','credit','category','company','payMethod','source','day']];
      sheet.clearContents();
      sheet.getRange(1, 1, 1, 11).setValues(headers);
      const CHUNK = 500;
      for (let i = 0; i < allPerf.length; i += CHUNK) {
        const chunk = allPerf.slice(i, i + CHUNK).map(r => [
          r.SUNAB_PK||'', r.증권번호||'', r.month||'', r.agentId||'',
          r.premium||0, r.credit||0, r.category||'', r.company||'', r.payMethod||'',
          (r.source||_postSrc||'마감'), (r.day||0)
        ]);
        sheet.getRange(i + 2, 1, chunk.length, 11).setValues(chunk);
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
    const employeesRaw = payload.employees || [];
    const perfForAgg    = allPerf.map(r => ({...r, status:'실제데이터'}));

    // 미매칭 사원 조직 지정 마스터 적용 (재업로드 시에도 지정값이 살아남도록)
    const orgOverrides = readOrgOverrides();
    const employees     = applyOrgOverridesToEmployees(employeesRaw, orgOverrides);
    const empForAgg      = addOverrideStubs(employees, orgOverrides);

    const aggData    = buildAggregated(perfForAgg, empForAgg);
    const topResult  = buildTop10(perfForAgg, empForAgg);
    const top10ByMonthData = buildTop10ByMonth(perfForAgg, empForAgg);

    // 캐시 즉시 갱신
    writeCache({
      ok:           true,
      lastUpdated:  now,
      employees,
      performance:  perfForAgg,
      aggregated:   aggData,
      top10:         topResult.top10,
      topByBranch:   topResult.topByBranch,
      topByBranch3:  topResult.topByBranch3,
      top10ByMonth:  top10ByMonthData,
      monthStatus:   buildMonthStatus(perfForAgg),
      orgOverrides:  orgOverrides,
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

// ─── 미매칭 사원 조직 지정 저장 (대시보드 팝업 → 서버 마스터 병합) ──
// payload: { action:'saveOrgOverrides', key, overrides:{ id:[사업단,본부] 또는 {b2,b3,name}, ... } }
// 값이 null이면 해당 사번 지정을 삭제. 저장 즉시 캐시를 재계산해 재업로드 없이 반영한다.
function handleSaveOrgOverrides(payload) {
  try {
    if (payload.key !== SYNC_KEY) return json({ ok: false, error: '인증 실패' });
    const incoming = payload.overrides || {};
    const master   = readOrgOverrides();
    Object.keys(incoming).forEach(id => {
      const v = incoming[id];
      if (v === null) delete master[id];
      else master[id] = v;
    });
    writeOrgOverrides(master);

    // 저장 즉시 시트 기준으로 캐시/사전집계 재계산 (재배포·재업로드 불필요)
    try { readSheetsAndCache(); } catch(e) { /* 캐시 재계산 실패해도 마스터 저장 자체는 성공 처리 */ }

    return json({ ok: true, count: Object.keys(master).length });
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

  // ── 규정 가동: 설계사별 월 인정실적(마감)/영수금(가마감) 10만원 이상 ──
  const _agm = {}, _ACT_TH = 100000;
  (performance || []).forEach(r => {
    const m = String(r.month || '').trim(); const id = String(r.agentId || ''); if (!m || !id) return;
    const k = id + '|' + m; if (!_agm[k]) _agm[k] = { v:0 };
    _agm[k].v += ((r.source||'').toString().trim() === '가마감') ? (Number(r.premium)||0) : (Number(r.credit)||0);
  });
  function _gActive(id, m){ const a = _agm[id + '|' + m]; return a ? a.v >= _ACT_TH : false; }

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
        provPrem: 0, provCred: 0,   // 가마감(잠정) 부분 — 마감/잠정 분해용
        activeSet: {}   // agentId → true (Set 대신 객체로 직렬화 회피)
      };
    }
    const v = map[key];
    const pre = Number(r.premium) || 0;
    const cre = Number(r.credit)  || 0;
    v.premium += pre;
    v.credit  += cre;
    v.count   += 1;
    if ((r.source||'').toString().trim() === '가마감') { v.provPrem += pre; v.provCred += cre; }
    if (r.agentId && _gActive(String(r.agentId), month)) {
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
    provPrem:     v.provPrem,
    provCred:     v.provCred,
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
  const latM = allMonths.length ? allMonths[allMonths.length - 1] : '';

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

  // agentId별: 마감6개월 총 premium/credit + 당월(latM) premium
  // ⚠ 월납 기준 (일시납·연납 대형건이 순위를 왜곡하는 것 방지 — 뷰어 KPI·조직표와 기준 통일)
  const agentData = {};
  (performance || []).forEach(r => {
    if ((r.payMethod || '').trim() !== '월납') return;
    const m = String(r.month || '').trim();
    const id  = String(r.agentId || '');
    if (!id) return;
    const pre = Number(r.premium) || 0;
    const cre = Number(r.credit)  || 0;
    if (!agentData[id]) agentData[id] = { totalPrem: 0, latPrem: 0, credit: 0 };
    if (targetMonths.has(m)) { agentData[id].totalPrem += pre; agentData[id].credit += cre; }
    if (m === latM)          { agentData[id].latPrem  += pre; }
  });

  // 개인별: avgPremium(6개월 월평균 보험료) + latPremium(당월 보험료) + avgCredit(월평균 인정)
  const list = Object.entries(agentData).map(([id, d]) => {
    const emp = empMap[id] || { name: id, branch: '', branch3: '' };
    return {
      agentId:      id,
      name:         emp.name,
      branch:       emp.branch,
      branch3:      emp.branch3,
      totalPremium: d.totalPrem,
      avgPremium:   Math.round(d.totalPrem / n),
      latPremium:   d.latPrem,
      avgCredit:    Math.round(d.credit / n),
      months:       n,
    };
  });

  // 당월 실적순 정렬
  list.sort((a, b) => b.latPremium - a.latPremium);

  // 전체 TOP10
  const top10 = list.slice(0, 10).map((v, i) => ({ rank: i + 1, ...v }));

  // 소속사업단(branch/b2)별 TOP10
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

  // 소속본부(branch3/b3)별 TOP10
  const byBranch3 = {};
  list.forEach(agent => {
    const b3 = agent.branch3 || '직할';
    if (!byBranch3[b3]) byBranch3[b3] = [];
    byBranch3[b3].push(agent);
  });
  const topByBranch3 = {};
  Object.keys(byBranch3).forEach(b3 => {
    topByBranch3[b3] = byBranch3[b3].slice(0, 10).map((v, i) => ({ rank: i + 1, ...v }));
  });

  return { top10, topByBranch, topByBranch3 };
}

// ─── TOP10 월별 스냅샷 생성 (뷰어 납입연/월 필터 연동용) ──
// 각 실적월(YYYYMM)마다 그 달 실적(latPremium) 기준 TOP10 / 사업단별 TOP10 / 본부별 TOP10
// 반환: { 'YYYYMM': { all:[...10], byBranch:{사업단:[...10]}, byBranch3:{본부:[...10]} }, ... }
function buildTop10ByMonth(performance, employees) {
  // employee map: id → { name, branch(b2), branch3(b3) }
  const empMap = {};
  (employees || []).forEach(e => {
    empMap[String(e.id)] = {
      name:    (e.name || '').trim(),
      branch:  (e.b2   || '').trim(),
      branch3: (e.b3   || '').trim(),
    };
  });

  // month → agentId → { premium, credit } 그 달 합계 (월납 기준 — buildTop10과 동일)
  const monthAgentMap = {};
  (performance || []).forEach(r => {
    if ((r.payMethod || '').trim() !== '월납') return;
    const m  = String(r.month   || '').trim();
    const id = String(r.agentId || '');
    if (!m || !id) return;
    if (!monthAgentMap[m]) monthAgentMap[m] = {};
    if (!monthAgentMap[m][id]) monthAgentMap[m][id] = { premium: 0, credit: 0 };
    monthAgentMap[m][id].premium += Number(r.premium) || 0;
    monthAgentMap[m][id].credit  += Number(r.credit)  || 0;
  });

  const top10ByMonth = {};
  Object.keys(monthAgentMap).forEach(m => {
    const list = Object.entries(monthAgentMap[m]).map(([id, d]) => {
      const emp = empMap[id] || { name: id, branch: '', branch3: '' };
      return {
        agentId:    id,
        name:       emp.name,
        branch:     emp.branch,
        branch3:    emp.branch3,
        latPremium: d.premium,   // 해당 월 월보험료
        avgCredit:  d.credit,    // 해당 월 인정실적 (단월 — "월평균" 아님, 뷰어에서 라벨 구분)
      };
    });

    // 그 달 실적순 정렬
    list.sort((a, b) => b.latPremium - a.latPremium);

    const all = list.slice(0, 10).map((v, i) => ({ rank: i + 1, ...v }));

    const byBranch = {};
    list.forEach(a => { const b = a.branch || '미분류'; (byBranch[b] = byBranch[b] || []).push(a); });
    const branchTop = {};
    Object.keys(byBranch).forEach(b => {
      branchTop[b] = byBranch[b].slice(0, 10).map((v, i) => ({ rank: i + 1, ...v }));
    });

    const byBranch3 = {};
    list.forEach(a => { const b3 = a.branch3 || '직할'; (byBranch3[b3] = byBranch3[b3] || []).push(a); });
    const branch3Top = {};
    Object.keys(byBranch3).forEach(b3 => {
      branch3Top[b3] = byBranch3[b3].slice(0, 10).map((v, i) => ({ rank: i + 1, ...v }));
    });

    top10ByMonth[m] = { all, byBranch: branchTop, byBranch3: branch3Top };
  });

  return top10ByMonth;
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
    source:  (r.source||'').toString().trim() || '마감',  // 마감=건별실적 / 가마감=자동실적
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

// ─── 월별 마감/가마감 상태 (가마감=자동실적 1건이라도 있으면 가마감) ──
function buildMonthStatus(performance) {
  const prov = {}, fin = {};
  (performance || []).forEach(r => {
    const m = String(r.month||'').trim(); if (!m) return;
    if ((r.source||'마감').toString().trim() === '가마감') prov[m] = true; else fin[m] = true;
  });
  const st = {};
  new Set([...Object.keys(prov), ...Object.keys(fin)]).forEach(m => { st[m] = (prov[m] && !fin[m]) ? '가마감' : '마감'; });
  return st;
}

// ─── 관리자 로그인: 사번+비번 검증 후 소속 데이터만 반환 ───
// '관리자' 시트 헤더: 사번 | 비번 | 레벨(사업단|본부) | 사업단 | 본부 | 성명
function handleManagerLogin(payload) {
  const emp = String(payload.emp || '').trim();
  const pw  = String(payload.pw  || '').trim();
  if (!emp || !pw) return json({ ok: false, error: '사원번호와 비밀번호를 입력하세요' });

  const ss    = getSpreadsheet();
  const sheet = ss.getSheetByName('관리자');
  if (!sheet) return json({ ok: false, error: '관리자 명단이 설정되지 않았습니다 — 운영자에게 문의하세요' });

  const rows = sheetToObjects(sheet);
  const me   = rows.find(r =>
    String(r['사번'] || '').trim() === emp &&
    String(r['비번'] || '').trim() === pw);
  if (!me) return json({ ok: false, error: '사원번호 또는 비밀번호가 올바르지 않습니다' });

  const level = String(me['레벨']  || '').trim();   // '사업단' | '본부'
  const sadan = String(me['사업단'] || '').trim();
  const bonbu = String(me['본부']  || '').trim();
  const name  = String(me['성명']  || '').trim() || emp;

  // 전체 캐시 확보 (없으면 재생성)
  let cached = readCache();
  if (!(cached && cached.ok && cached.employees)) { readSheetsAndCache(); cached = readCache(); }
  if (!cached) return json({ ok: false, error: '데이터 캐시 없음 — 운영자에게 문의하세요' });

  // 소속 범위 판정: 사업단 레벨=사업단 전체 / 본부 레벨=해당 본부만
  const inScope = (b2, b3) => (level === '사업단')
    ? ((b2 || '').trim() === sadan)
    : ((b3 || '').trim() === bonbu);

  const employees  = (cached.employees  || []).filter(e => inScope(e.b2, e.b3));
  const aggregated = (cached.aggregated || []).filter(r => inScope(r.branch, r.branch3));
  const top10      = (level === '사업단')
    ? ((cached.topByBranch  || {})[sadan] || [])
    : ((cached.topByBranch3 || {})[bonbu] || []);

  return json({
    ok:          true,
    scope:       { level, 사업단: sadan, 본부: bonbu, name },
    lastUpdated: cached.lastUpdated,
    employees,
    aggregated,
    monthStatus: cached.monthStatus || {},
    top10,
    topByBranch:  {},
    topByBranch3: {},
  });
}

// ─── (1회 실행용) '관리자' 시트 생성 + 헤더/예시행 ───────────
// Apps Script 편집기에서 이 함수를 직접 실행하면 탭이 만들어집니다.
function setupManagerSheet() {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName('관리자');
  if (!sheet) sheet = ss.insertSheet('관리자');
  sheet.clearContents();
  const rows = [
    ['사번', '비번', '레벨', '사업단', '본부', '성명'],
    ['999001', 'inca1234', '사업단', '해빙총괄사업단', '', '(예시)사업단장'],
    ['999002', 'inca1234', '본부',   '해빙총괄사업단', '미라클본부', '(예시)본부장'],
  ];
  sheet.getRange(1, 1, rows.length, 6).setValues(rows);
  Logger.log('관리자 시트 생성 완료 — 예시행을 실제 관리자로 교체하세요');
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
