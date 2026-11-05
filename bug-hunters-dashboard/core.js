/* Bug Hunters dashboard – pure data logic (no DOM). Works in browser (window.BH) and Node. */
(function (root) {
  const QA = ['Passed', 'Opportunity', 'Failed', 'Critical'];
  const MONTHS = ['january','february','march','april','may','june','july','august','september','october','november','december'];
  const MIN_REVIEWED = 5;   // minimum QA-reviewed cases to show an error %
  // "QA Fixed Comment" starting with NA / N/A means the QA finding does not apply.
  const isNA = c => /^\s*N\/?A\b/i.test(String(c == null ? '' : c));

  const norm = s => String(s == null ? '' : s).trim().toLowerCase().replace(/\s+/g, ' ');
  const cap = s => s ? s[0].toUpperCase() + s.slice(1).toLowerCase() : s;

  function normQA(v) {
    const s = norm(v);
    if (!s) return -1;
    if (s === 'passed' || s === 'pass') return 0;
    if (s === 'opportunity') return 1;
    if (s === 'failed' || s === 'fail') return 2;
    if (s === 'critical') return 3;
    return 4; // e.g. "In progress": kept in the data, excluded from QA math
  }

  function parseDate(v) {
    if (v instanceof Date && !isNaN(v)) return v;
    if (typeof v === 'number' && v > 20000 && v < 80000) return new Date(Math.round((v - 25569) * 86400000));
    if (typeof v === 'string') { const d = new Date(v); if (!isNaN(d) && /\d/.test(v)) return d; }
    return null;
  }

  // Find a column by header name, ignoring case and spaces.
  function pick(obj, names) {
    for (const k of Object.keys(obj)) if (names.includes(norm(k))) return obj[k];
    return undefined;
  }

  /** raw: array of objects (one per spreadsheet row, keyed by header). Returns {model} or {error}. */
  function build(raw) {
    if (!raw || !raw.length) return { error: 'The sheet has no rows.' };
    const first = raw[0];
    const have = k => Object.keys(first).some(h => norm(h) === k);
    const missing = ['owner', 'hours', 'type'].filter(k => !have(k));
    if (!have('month') && !have('date')) missing.push('month (or date)');
    if (missing.length) return { error: 'Missing column(s): ' + missing.join(', ').toUpperCase() };

    const owners = [], types = [], monthInfo = {}, rows = [];
    const commentKey = Object.keys(first).find(h => norm(h).includes('fixed') || norm(h).includes('comment'));
    raw.forEach(r => {
      const owner = String(pick(r, ['owner']) || '').trim();
      const hours = Number(pick(r, ['hours']));
      if (!owner || !isFinite(hours)) return;
      // unify "DDC - Site Services" / "DDC – Site Services" and similar hyphen variants
      const type = String(pick(r, ['type']) || '(none)').trim().replace(/\s*[-–—]\s*/g, ' – ');
      const date = parseDate(pick(r, ['date']));
      let mname = String(pick(r, ['month']) || '').trim();
      if (!mname && date) mname = cap(MONTHS[date.getMonth()]);
      if (!mname) return;
      mname = cap(mname);
      const idRaw = pick(r, ['id']);
      const id = idRaw == null || idRaw === '' ? '' : String(idRaw).trim();
      if (!owners.includes(owner)) owners.push(owner);
      if (!types.includes(type)) types.push(type);
      const mi = monthInfo[mname] || (monthInfo[mname] = { name: mname, min: null, max: null });
      if (date) { if (!mi.min || date < mi.min) mi.min = date; if (!mi.max || date > mi.max) mi.max = date; }
      const q = normQA(pick(r, ['qa status']));
      const na = q >= 1 && q <= 3 && commentKey && isNA(r[commentKey]) ? 1 : 0;
      rows.push({ o: owners.indexOf(owner), id, t: types.indexOf(type), h: hours, q, na, mname });
    });
    if (!rows.length) return { error: 'No valid rows found (need OWNER, HOURS, TYPE and Month/DATE).' };

    const months = Object.values(monthInfo).sort((a, b) => {
      if (a.min && b.min) return a.min - b.min;
      return MONTHS.indexOf(a.name.toLowerCase()) - MONTHS.indexOf(b.name.toLowerCase());
    });
    // A month is "partial" if it is the latest one and its last date is well before month end.
    const last = months[months.length - 1];
    months.forEach(m => m.partial = false);
    if (last && last.max) {
      const dim = new Date(last.max.getFullYear(), last.max.getMonth() + 1, 0).getDate();
      last.partial = last.max.getDate() < dim - 3;
    }
    rows.forEach(r => { r.m = months.findIndex(m => m.name === r.mname); delete r.mname; });
    return { model: { owners, types, rows, months: months.map(m => ({ name: m.name, partial: m.partial, last: m.max ? m.max.toISOString().slice(0, 10) : null })) } };
  }

  /** Owners with very few rows overall (e.g. someone who left) are treated as outside the team. */
  function teamOwners(model, includeAll) {
    const c = {};
    model.rows.forEach(r => c[r.o] = (c[r.o] || 0) + 1);
    const idx = model.owners.map((_, i) => i);
    return {
      team: idx.filter(i => includeAll || (c[i] || 0) >= 10),
      extra: idx.filter(i => !includeAll && (c[i] || 0) < 10).map(i => ({ owner: model.owners[i], rows: c[i] || 0 }))
    };
  }

  function stats(rs) {
    const n = rs.length, h = rs.reduce((a, r) => a + r.h, 0);
    const c = [0, 0, 0, 0], nac = [0, 0, 0, 0];   // c = valid findings, nac = "NA" findings (do not apply)
    rs.forEach(r => { if (r.q >= 0 && r.q < 4) (r.na ? nac : c)[r.q]++; });
    const rev = c[0] + c[1] + c[2] + c[3];          // QA-reviewed cases that apply
    const na = nac[0] + nac[1] + nac[2] + nac[3];
    return { n, h, avg: n ? h / n : 0, c, nac, na, rev, errors: c[2] + c[3],
      err: rev >= MIN_REVIEWED ? (c[2] + c[3]) / rev : null, errRaw: rev ? (c[2] + c[3]) / rev : null, cov: n ? (rev + na) / n : 0 };
  }

  /**
   * Cases for the "longest case" metric. Counting stays per row; only the effort (hours) of a case is the
   * sum of every row that shares its ID in that month (DDC row + SS Support / SS RW rows, any owner).
   * Rows without an ID are their own case. If person >= 0, only cases where that person has a row.
   */
  function cases(model, month, person) {
    const map = new Map();
    model.rows.forEach((r, i) => {
      if (month != null && r.m !== month) return;
      const key = r.id ? r.m + '|' + r.id : 'row' + i;
      let c = map.get(key);
      if (!c) map.set(key, c = { id: r.id, h: 0, rows: [] });
      c.h += r.h; c.rows.push(r);
    });
    let list = [...map.values()];
    if (person != null && person >= 0) list = list.filter(c => c.rows.some(r => r.o === person));
    return list.sort((a, b) => b.h - a.h);
  }

  const api = { QA, MIN_REVIEWED, build, teamOwners, stats, cases, normQA, parseDate };
  if (typeof module !== 'undefined' && module.exports) module.exports = api; else root.BH = api;
})(typeof window !== 'undefined' ? window : globalThis);
