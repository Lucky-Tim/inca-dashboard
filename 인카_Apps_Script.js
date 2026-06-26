// ═══════════════════════════════════════════════════════════
// 인카금융 운영 플랫폼 — Google Apps Script v4 (레지스트리 패턴)
// ─────────────────────────────────────────────────────────
// 설정 방법:
//   1. https://script.google.com 접속 → 기존 프로젝트 열기
//   2. 이 코드 전체 붙여넣기 (Ctrl+A → Ctrl+V)
//   3. 저장(Ctrl+S)
//   4. 배포 → 기존 배포 관리 → ✏️ 편집 → "새 버전" → 배포
//
// 데이터 업데이트 방법:
//   Google Sheets (인카_운영데이터) 탭을 직접 수정하거나
//   xlsx 파일을 Drive에서 Sheets로 가져오기(Import) 후 덮어쓰기
//
// 확장 방법:
//   새 데이터 타입 추가 → DATA_SOURCES 배열에 한 줄 추가
//                       + 하단 파서 함수 작성
// ═══════════════════════════════════════════════════════════

const SYNC_KEY = 'INCA2026';
const SS_PROP  = 'SPREADSHEET_ID';

// ─── 데이터 소스 레지스트리 ──────────────────────────────
// 새 데이터 추가 시 여기에 한 줄만 등록
const DATA_SOURCES = [
  { key: 'employees',   sheetName: '직원관리',  parser: parseEmpRows  },
  { key: 'performance', sheetName: '건별실적',  parser: parsePerfRows },
  // 향후 확장 예시 (시트 추가 후 주석 해제):
  // { key: 'collection', sheetName: '수금율', parser: parseCollRows },
  // { key: 'retention',  sheetName: '유지율', parser: parseRetRows  },
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

// ─── GET: 레지스트리 순회 → JSON 반환 ───────────────────
function doGet(e) {
  try {
    const ss     = getSpreadsheet();
    const result = { ok: true, lastUpdated: getMeta(ss, 'lastUpdated') };

    DATA_SOURCES.forEach(src => {
      const sheet = ss.getSheetByName(src.sheetName);
      result[src.key] = sheet ? src.parser(sheetToObjects(sheet)) : [];
    });

    // 최소 데이터 확인 (직원 + 실적)
    if (!result.employees.length && !result.performance.length) {
      return json({ ok: false, error: '데이터 없음 — Sheets에 직원관리·건별실적 탭을 확인하세요' });
    }

    return json(result);
  } catch(err) {
    return json({ ok: false, error: err.message });
  }
}

// ─── POST: 외부에서 데이터 밀어넣기 (선택적 사용) ───────
// Sheets 직접 편집이 주 방식. 이 엔드포인트는 보조 수단.
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    if (payload.key !== SYNC_KEY) return json({ ok: false, error: '인증 실패' });

    const ss = getSpreadsheet();

    // 직원관리 저장
    if (payload.employees?.length) {
      const sheet = getSheet(ss, '직원관리');
      sheet.clearContents();
      const rows = [['id','name','type','b1','b2','b3','joinDate']];
      payload.employees.forEach(e =>
        rows.push([e.id||'', e.name||'', e.type||'', e.b1||'', e.b2||'', e.b3||'', e.joinDate||''])
      );
      sheet.getRange(1, 1, rows.length, 7).setValues(rows);
    }

    // 건별실적 저장 (500행 청크)
    if (payload.performance?.length) {
      const sheet   = getSheet(ss, '건별실적');
      const headers = [['SUNAB_PK','증권번호','month','agentId','premium','credit','category','company','payMethod']];
      sheet.clearContents();
      sheet.getRange(1, 1, 1, 9).setValues(headers);
      const CHUNK = 500;
      for (let i = 0; i < payload.performance.length; i += CHUNK) {
        const chunk = payload.performance.slice(i, i + CHUNK).map(r => [
          r.SUNAB_PK||'', r.증권번호||'', r.month||'', r.agentId||'',
          r.premium||0, r.credit||0, r.category||'', r.company||'', r.payMethod||''
        ]);
        sheet.getRange(i + 2, 1, chunk.length, 9).setValues(chunk);
      }
    }

    // 향후 확장: payload에 collection / retention 등 추가 시 여기서 처리

    setMeta(ss, 'lastUpdated', new Date().toISOString());
    setMeta(ss, 'empCount',    payload.employees?.length  || 0);
    setMeta(ss, 'perfCount',   payload.performance?.length || 0);

    return json({
      ok:        true,
      empCount:  payload.employees?.length,
      perfCount: payload.performance?.length,
    });
  } catch(err) {
    return json({ ok: false, error: err.message });
  }
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

// 향후 파서 함수 (시트 추가 시 작성):
// function parseCollRows(rows) { return rows; }  // 수금율
// function parseRetRows(rows)  { return rows; }  // 유지율

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
