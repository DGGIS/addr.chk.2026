/* ============================================================
   기초번호판 보고 시스템 - 공통 데이터 레이어 (app.js)
   모든 페이지(index/report/stats/upload)가 이 파일을 공유합니다.

   데이터 모델
   -----------
   주차 스냅샷(week snapshot) 1개 = {
     baseDate, uploadedAt, fileName,
     regions: [...66개 시군구],
     stats:   [...점검결과통계],
     vendors: [...업체별 배정현황]
   }

   저장 위치 2단:
   1) "커밋된" 데이터  : /data/index.json + /data/weekly/{baseDate}.json
      -> 저장소(GitHub)에 실제로 올라가 모든 사람/기기에서 공유됨 (정적 fetch)
   2) "로컬" 데이터    : localStorage
      -> 이 브라우저에서 방금 업로드했지만 아직 /data 에 커밋하지 않은 주차.
         업로드 페이지에서 JSON을 내려받아 저장소에 커밋하면 (1)로 승격됩니다.
   ============================================================ */

const LS_LOCAL_WEEKS = 'kaba_local_weeks';      // ["2026-06-21", "2026-06-28", ...]
const LS_WEEK_PREFIX  = 'kaba_week_';            // kaba_week_2026-06-28 -> snapshot JSON
const LS_NOTES        = 'kaba_report_notes';

const num = v => (typeof v === 'number' && isFinite(v)) ? v : (parseFloat(v) || 0);
const fmt = v => Math.round(num(v)).toLocaleString('ko-KR');
const pct = v => (num(v) * 100).toFixed(1) + '%';
const pctNum = v => num(v) * 100;
const $ = (s, el = document) => el.querySelector(s);
const $$ = (s, el = document) => Array.from(el.querySelectorAll(s));

/* ---------------- xlsx -> snapshot ---------------- */
const Parser = (() => {
  function toAOA(ws) { return XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true }); }

  function findHeaderRow(aoa, mustHave) {
    for (let i = 0; i < aoa.length; i++) {
      const row = (aoa[i] || []).map(c => c == null ? '' : String(c).trim());
      if (mustHave.every(k => row.includes(k))) return i;
    }
    return -1;
  }

  function extractBaseDate(aoa) {
    for (const row of aoa.slice(0, 4)) {
      const txt = (row || []).map(c => c == null ? '' : String(c)).join(' ');
      const m = txt.match(/기준일자\s*:?\s*([0-9\-\.\/]{8,10})/);
      if (m) return m[1];
    }
    return '';
  }

  function parseRegions(wb) {
    if (!wb.SheetNames.includes('시군구별_진도현황')) return { regions: [], baseDate: '' };
    const aoa = toAOA(wb.Sheets['시군구별_진도현황']);
    const baseDate = extractBaseDate(aoa);
    const hIdx = findHeaderRow(aoa, ['No', '시도', '시군구']);
    if (hIdx < 0) return { regions: [], baseDate };
    const headerRow = aoa[hIdx].map(c => c == null ? '' : String(c).trim());
    const idIdx = headerRow.indexOf('No');
    const regions = [];
    for (let r = hIdx + 1; r < aoa.length; r++) {
      const row = aoa[r]; if (!row) continue;
      if (typeof row[idIdx] !== 'number') continue;
      const o = {}; headerRow.forEach((h, i) => { if (h) o[h] = row[i]; });
      regions.push({
        no: o['No'], sido: o['시도'] || '', sigungu: o['시군구'] || '', vendor: o['담당업체'] || '',
        target: num(o['전체대상건수']), prevWeek: num(o['전주누적건수']), thisWeek: num(o['금주점검건수']),
        cumulative: num(o['누적점검건수']), progress: num(o['진도율(%)']), remain: num(o['잔여건수']),
        normal: num(o['정상']), damage: num(o['훼손']), lost: num(o['망실']), note: o['비고'] || ''
      });
    }
    return { regions, baseDate };
  }

  function parseStats(wb) {
    if (!wb.SheetNames.includes('점검결과통계')) return [];
    const aoa = toAOA(wb.Sheets['점검결과통계']);
    const hIdx = findHeaderRow(aoa, ['No', '시도', '시군구']);
    if (hIdx < 0) return [];
    const headerRow = aoa[hIdx].map(c => c == null ? '' : String(c).trim());
    const idIdx = headerRow.indexOf('No');
    const rows = [];
    for (let r = hIdx + 1; r < aoa.length; r++) {
      const row = aoa[r]; if (!row) continue;
      if (typeof row[idIdx] !== 'number') continue;
      const o = {}; headerRow.forEach((h, i) => { if (h) o[h] = row[i]; });
      rows.push({
        sido: o['시도'] || '', sigungu: o['시군구'] || '', inspectRate: num(o['점검율(%)']), damageRate: num(o['훼손율(%)']),
        total: num(o['점검계']), normal: num(o['정상']), damage: num(o['훼손']), lost: num(o['망실']),
        fixedRepair: num(o['기조치(보수)']), fixedReplace: num(o['기조치(교체)']), needAction: num(o['조치필요']),
        planBudget: num(o['교체예정(예산부족)']), planRoad: num(o['교체예정(도로공사연계)']), actionRate: num(o['조치율(%)'])
      });
    }
    return rows;
  }

  function parseVendors(wb) {
    if (!wb.SheetNames.includes('업체별_배정현황')) return [];
    const aoa = toAOA(wb.Sheets['업체별_배정현황']);
    const hIdx = findHeaderRow(aoa, ['No', '업체명']);
    if (hIdx < 0) return [];
    const headerRow = aoa[hIdx].map(c => c == null ? '' : String(c).trim());
    const idIdx = headerRow.indexOf('No');
    const rows = [];
    for (let r = hIdx + 1; r < aoa.length; r++) {
      const row = aoa[r]; if (!row) continue;
      if (typeof row[idIdx] !== 'number') continue;
      const o = {}; headerRow.forEach((h, i) => { if (h) o[h] = row[i]; });
      const name = (o['업체명'] || '').toString().trim();
      const area = (o['담당 시군구'] || '').toString().trim();
      if (!name || name.startsWith('[입력')) continue;
      rows.push({
        name, area, count: o['배정 시군구 수'], target: num(o['배정 대상건수']),
        cumulative: num(o['누적 점검건수']), progress: num(o['진도율']), note: o['비고'] || ''
      });
    }
    return rows;
  }

  function parse(arrayBuffer, fileName) {
    const wb = XLSX.read(arrayBuffer, { type: 'array', cellDates: false });
    const { regions, baseDate } = parseRegions(wb);
    if (regions.length === 0) return null;
    return {
      baseDate: baseDate || new Date().toISOString().slice(0, 10),
      uploadedAt: new Date().toISOString(),
      fileName,
      regions,
      stats: parseStats(wb),
      vendors: parseVendors(wb)
    };
  }

  return { parse };
})();

/* ---------------- local storage ---------------- */
const LocalStore = {
  listWeeks() {
    try { return JSON.parse(localStorage.getItem(LS_LOCAL_WEEKS) || '[]'); } catch (e) { return []; }
  },
  saveWeek(snapshot) {
    const weeks = new Set(this.listWeeks());
    weeks.add(snapshot.baseDate);
    localStorage.setItem(LS_LOCAL_WEEKS, JSON.stringify([...weeks].sort()));
    localStorage.setItem(LS_WEEK_PREFIX + snapshot.baseDate, JSON.stringify(snapshot));
  },
  getWeek(date) {
    try { return JSON.parse(localStorage.getItem(LS_WEEK_PREFIX + date) || 'null'); } catch (e) { return null; }
  },
  deleteWeek(date) {
    const weeks = this.listWeeks().filter(d => d !== date);
    localStorage.setItem(LS_LOCAL_WEEKS, JSON.stringify(weeks));
    localStorage.removeItem(LS_WEEK_PREFIX + date);
  }
};

/* ---------------- committed (data/) + local 통합 레이어 ---------------- */
const DataLayer = (() => {
  let committedIndex = null; // ["2026-06-14", ...] or null if not fetched yet
  const committedCache = {};

  async function fetchCommittedIndex() {
    if (committedIndex !== null) return committedIndex;
    try {
      const res = await fetch('data/index.json', { cache: 'no-store' });
      if (!res.ok) throw new Error('no index');
      const json = await res.json();
      committedIndex = Array.isArray(json.weeks) ? json.weeks : [];
    } catch (e) {
      committedIndex = [];
    }
    return committedIndex;
  }

  async function fetchCommittedWeek(date) {
    if (committedCache[date]) return committedCache[date];
    try {
      const res = await fetch(`data/weekly/${date}.json`, { cache: 'no-store' });
      if (!res.ok) throw new Error('not found');
      const json = await res.json();
      committedCache[date] = json;
      return json;
    } catch (e) {
      return null;
    }
  }

  async function listAllWeeks() {
    const committed = await fetchCommittedIndex();
    const local = LocalStore.listWeeks();
    const map = new Map();
    committed.forEach(d => map.set(d, 'committed'));
    local.forEach(d => map.set(d, map.has(d) ? 'both' : 'local'));
    return [...map.entries()].map(([date, source]) => ({ date, source })).sort((a, b) => a.date.localeCompare(b.date));
  }

  async function getWeek(date) {
    const local = LocalStore.getWeek(date);
    if (local) return local; // 로컬이 있으면 최신으로 간주 (방금 업로드분 우선)
    return await fetchCommittedWeek(date);
  }

  async function getLatestWeek() {
    const weeks = await listAllWeeks();
    if (weeks.length === 0) return null;
    return await getWeek(weeks[weeks.length - 1].date);
  }

  return { listAllWeeks, getWeek, getLatestWeek };
})();

/* ---------------- aggregation ---------------- */
function totals(regions) {
  const t = { target: 0, prevWeek: 0, thisWeek: 0, cumulative: 0, remain: 0, normal: 0, damage: 0, lost: 0, count: regions.length };
  regions.forEach(r => {
    t.target += r.target; t.prevWeek += r.prevWeek; t.thisWeek += r.thisWeek; t.cumulative += r.cumulative;
    t.remain += r.remain; t.normal += r.normal; t.damage += r.damage; t.lost += r.lost;
  });
  t.progress = t.target ? t.cumulative / t.target : 0;
  return t;
}
function bySido(regions) {
  const map = {};
  regions.forEach(r => {
    const k = r.sido || '기타';
    if (!map[k]) map[k] = { sido: k, count: 0, target: 0, cumulative: 0, thisWeek: 0, normal: 0, damage: 0, lost: 0 };
    const m = map[k];
    m.count++; m.target += r.target; m.cumulative += r.cumulative; m.thisWeek += r.thisWeek;
    m.normal += r.normal; m.damage += r.damage; m.lost += r.lost;
  });
  return Object.values(map).map(m => ({ ...m, progress: m.target ? m.cumulative / m.target : 0 })).sort((a, b) => b.target - a.target);
}
function byVendorFromRegions(regions) {
  const map = {};
  regions.forEach(r => {
    const k = r.vendor || '미배정';
    if (!map[k]) map[k] = { name: k, count: 0, target: 0, cumulative: 0 };
    const m = map[k];
    m.count++; m.target += r.target; m.cumulative += r.cumulative;
  });
  return Object.values(map).map(m => ({ ...m, progress: m.target ? m.cumulative / m.target : 0 })).sort((a, b) => b.target - a.target);
}

/* ---------------- shared UI builders ---------------- */
function statusBadge(r) {
  if (r.progress >= 1) return '<span class="badge ok">완료</span>';
  if (r.thisWeek > 0) return '<span class="badge ok">진행중</span>';
  if (r.cumulative === 0) return '<span class="badge danger">미착수</span>';
  return '<span class="badge warn">지연</span>';
}
function kpiCard(label, value, cls = '') {
  return `<div class="kpi ${cls}"><div class="label">${label}</div><div class="value">${value}</div></div>`;
}
function barRow(name, progressFraction, valueText) {
  return `<div class="bar-row">
    <div class="nm">${name}</div>
    <div class="track"><div class="fill" style="width:${Math.min(pctNum(progressFraction), 100).toFixed(1)}%"></div></div>
    <div class="pct">${valueText != null ? valueText : pct(progressFraction)}</div>
  </div>`;
}
function downloadJSON(obj, filename) {
  const blob = new Blob([JSON.stringify(obj, null, 2)], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a'); a.href = url; a.download = filename;
  document.body.appendChild(a); a.click(); a.remove();
  URL.revokeObjectURL(url);
}

/* ---------------- nav bar (shared across pages) ---------------- */
function renderNav(active) {
  const items = [
    { id: 'view', href: 'index.html', label: '요약', icon: '<rect x="3" y="3" width="7" height="9"/><rect x="14" y="3" width="7" height="5"/><rect x="14" y="12" width="7" height="9"/><rect x="3" y="16" width="7" height="5"/>' },
    { id: 'report', href: 'report.html', label: '보고', icon: '<rect x="3" y="4" width="18" height="18" rx="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/>' },
    { id: 'stats', href: 'stats.html', label: '통계', icon: '<circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 3"/>' },
    { id: 'upload', href: 'upload.html', label: '업로드', icon: '<path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/>' },
  ];
  const html = items.map(it => `
    <a href="${it.href}" class="${it.id === active ? 'active' : ''}">
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">${it.icon}</svg>
      <span>${it.label}</span>
    </a>`).join('');
  const el = document.getElementById('tabbar');
  if (el) el.innerHTML = html;
}

/* ---------------- week selector (shared) ---------------- */
async function buildWeekSelector(selectEl, onChange, presetDate) {
  const weeks = await DataLayer.listAllWeeks();
  if (weeks.length === 0) {
    selectEl.innerHTML = `<option>데이터 없음</option>`;
    selectEl.disabled = true;
    return { weeks, current: null };
  }
  selectEl.disabled = false;
  selectEl.innerHTML = weeks.map(w => `<option value="${w.date}">${w.date}${w.source === 'local' ? ' (로컬)' : ''}</option>`).join('');
  const def = presetDate && weeks.find(w => w.date === presetDate) ? presetDate : weeks[weeks.length - 1].date;
  selectEl.value = def;
  selectEl.addEventListener('change', () => onChange(selectEl.value));
  return { weeks, current: def };
}
