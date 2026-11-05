/* Bug Hunters dashboard – UI. Needs core.js (window.BH) and SheetJS (window.XLSX). */
(function () {
  const $ = id => document.getElementById(id);
  const { QA, stats, cases, build, teamOwners } = BH;
  const STORE = 'bh-model-v2';
  const QCOL = ['var(--good)', 'var(--warn)', 'var(--serious)', 'var(--crit)', 'var(--muted)'];
  const mcol = i => `var(--m${(i % 6) + 1})`;
  const NAME = e => e.split('@')[0].split('.').map(s => s ? s[0].toUpperCase() + s.slice(1) : s).join(' ');
  const FIRST = e => NAME(e).split(' ')[0];

  let model = null, team = [], extra = [];
  let M = 0, P = -1, shown = new Set();

  // ---------- formatting
  const nf = (d) => x => x.toLocaleString('en-US', { minimumFractionDigits: d, maximumFractionDigits: d });
  const f0 = nf(0), f1 = nf(1), f2 = nf(2);
  const pc = x => x == null ? '—' : (x * 100).toFixed(1) + '%';
  function delta(cur, prev) {
    if (prev == null || cur == null || prev === 0) return '';
    const d = (cur - prev) / prev;
    return `<span>${d > 0 ? '▲' : d < 0 ? '▼' : '■'} ${d > 0 ? '+' : ''}${(d * 100).toFixed(0)}%</span>`;
  }
  function dpp(cur, prev) { // percentage-point change of a rate; more errors = worse
    if (cur == null || prev == null) return '';
    const d = (cur - prev) * 100;
    return `<span class="${d > 0 ? 'up' : d < 0 ? 'dn' : ''}">${d > 0 ? '▲' : d < 0 ? '▼' : '■'} ${d > 0 ? '+' : ''}${d.toFixed(1)} pp</span>`;
  }

  // ---------- helpers
  const rowsOf = (m, p) => model.rows.filter(r => team.includes(r.o) && (m == null || r.m === m) && (p == null || p < 0 || r.o === p));
  const teamModel = () => ({ ...model, rows: model.rows.filter(r => team.includes(r.o)) });
  const mname = i => model.months[i].name + (model.months[i].partial ? ' (partial)' : '');
  const tip = $('tip');
  function showTip(e, html) { tip.innerHTML = html; tip.style.opacity = 1; const w = tip.offsetWidth; let x = e.clientX + 12; if (x + w > innerWidth - 8) x = e.clientX - w - 12; tip.style.left = x + 'px'; tip.style.top = e.clientY + 12 + 'px'; }
  const hideTip = () => tip.style.opacity = 0;
  function bindTips(el) { el.querySelectorAll('[data-tt]').forEach(n => { n.addEventListener('mousemove', e => showTip(e, decodeURIComponent(n.dataset.tt))); n.addEventListener('mouseleave', hideTip); }); }
  const enc = s => encodeURIComponent(s);

  // ---------- loading
  function setModel(m, status) {
    model = m;
    const lastFull = [...model.months.keys()].filter(i => !model.months[i].partial);
    M = lastFull.length ? lastFull[lastFull.length - 1] : model.months.length - 1;
    shown = new Set(lastFull.length ? lastFull : model.months.keys());
    P = -1;
    $('status').textContent = status;
    $('clear').hidden = false;
    $('err').hidden = true;
    $('app').hidden = false;
    render();
  }
  function showError(msg) { $('err').textContent = msg; $('err').hidden = false; }

  function handleFile(file) {
    if (!file) return;
    if (typeof XLSX === 'undefined') return showError('Could not load the spreadsheet reader (SheetJS). Check your connection and reload.');
    const rd = new FileReader();
    rd.onerror = () => showError('Could not read the file.');
    rd.onload = ev => {
      try {
        const wb = XLSX.read(new Uint8Array(ev.target.result), { type: 'array', cellDates: true });
        let name = wb.SheetNames.find(n => n.trim().toLowerCase() === 'datos');
        if (!name) name = wb.SheetNames.find(n => { const r = XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1 })[0] || []; return r.some(h => String(h).trim().toLowerCase() === 'owner'); });
        if (!name) return showError('No sheet with an OWNER column was found. Expected a sheet named "Datos".');
        const raw = XLSX.utils.sheet_to_json(wb.Sheets[name], { defval: null });
        const { model: m, error } = build(raw);
        if (error) return showError(error);
        try { localStorage.setItem(STORE, JSON.stringify(m)); } catch (e) { /* storage unavailable: fine */ }
        setModel(m, `${file.name} · sheet "${name}" · ${m.rows.length} rows`);
      } catch (e) { showError('Could not parse the file: ' + e.message); }
    };
    rd.readAsArrayBuffer(file);
  }
  $('file').addEventListener('change', e => { handleFile(e.target.files[0]); e.target.value = ''; });
  const drop = $('drop');
  ['dragover', 'dragenter'].forEach(t => drop.addEventListener(t, e => { e.preventDefault(); drop.classList.add('over'); }));
  ['dragleave', 'drop'].forEach(t => drop.addEventListener(t, e => { e.preventDefault(); drop.classList.remove('over'); }));
  drop.addEventListener('drop', e => handleFile(e.dataTransfer.files[0]));
  $('clear').addEventListener('click', () => {
    try { localStorage.removeItem(STORE); } catch (e) {}
    model = null; $('app').hidden = true; $('clear').hidden = true; $('status').textContent = ''; $('sub').textContent = 'Upload your Excel to get started.';
  });

  // ---------- controls
  function controls() {
    const seg = $('mseg'); seg.innerHTML = '';
    model.months.forEach((m, i) => { const b = document.createElement('button'); b.textContent = mname(i); b.setAttribute('aria-pressed', i === M); b.onclick = () => { M = i; render(); }; seg.appendChild(b); });
    const ps = $('psel');
    ps.innerHTML = '<option value="-1">Whole team</option>' + team.map(o => `<option value="${o}">${NAME(model.owners[o])}</option>`).join('');
    ps.value = P; ps.onchange = () => { P = +ps.value; render(); };
    const ch = $('chips'); ch.innerHTML = '';
    model.months.forEach((m, i) => {
      const b = document.createElement('button'); b.className = 'chip'; b.setAttribute('aria-pressed', shown.has(i));
      b.innerHTML = `<i style="background:${mcol(i)}"></i>${mname(i)}`;
      b.onclick = () => { shown.has(i) ? shown.delete(i) : shown.add(i); render(); };
      ch.appendChild(b);
    });
    $('lmonths').innerHTML = '';
  }

  // ---------- charts
  function grouped(el, metric, fmt, o = {}) {
    const mons = [...shown].sort((a, b) => a - b);
    if (!mons.length) { el.innerHTML = '<p class="s">Select at least one month above.</p>'; return; }
    const people = P < 0 ? team : [P];
    const W = 460, H = 250, L = 40, R = 6, T = 16, B = 30;
    const vals = people.map(p => mons.map(m => metric(rowsOf(m, p))));
    let mx = Math.max(...vals.flat().filter(v => v != null), 0) * 1.14 || 1;
    const iw = W - L - R, ih = H - T - B, gw = iw / people.length, bw = Math.min(22, (gw - 8) / mons.length - 2);
    let s = `<svg viewBox="0 0 ${W} ${H}" role="img" aria-label="${o.label}">`;
    for (let k = 0; k <= 4; k++) { const y = T + ih - ih * k / 4; s += `<line x1="${L}" x2="${W - R}" y1="${y}" y2="${y}" stroke="${k ? 'var(--grid)' : 'var(--axis)'}"/><text x="${L - 6}" y="${y + 4}" text-anchor="end" font-size="11" fill="var(--muted)">${o.axis ? o.axis(mx * k / 4) : f0(mx * k / 4)}</text>`; }
    people.forEach((p, i) => {
      const cx = L + gw * i + gw / 2, x0 = cx - mons.length * (bw + 2) / 2;
      mons.forEach((m, j) => {
        const v = vals[i][j], x = x0 + j * (bw + 2);
        if (v == null) { s += `<text x="${x + bw / 2}" y="${T + ih - 4}" text-anchor="middle" font-size="10" fill="var(--muted)">—</text>`; return; }
        const h = Math.max(1, ih * v / mx), y = T + ih - h, st = stats(rowsOf(m, p));
        const tt = `<b>${NAME(model.owners[p])}</b> · ${mname(m)}<br>${o.tt ? o.tt(st) : fmt(v)}`;
        s += `<path d="M${x},${T + ih}V${y + 4}Q${x},${y} ${x + 4},${y}H${x + bw - 4}Q${x + bw},${y} ${x + bw},${y + 4}V${T + ih}Z" fill="${mcol(m)}" data-tt="${enc(tt)}"/>`;
        if (m === M) s += `<text x="${x + bw / 2}" y="${y - 4}" text-anchor="middle" font-size="10.5" font-weight="600" fill="var(--ink)">${fmt(v)}</text>`;
      });
      s += `<text x="${cx}" y="${H - 10}" text-anchor="middle" font-size="12" fill="var(--ink)">${FIRST(model.owners[p])}</text>`;
    });
    el.innerHTML = s + '</svg>'; bindTips(el);
  }

  function statusBars(el) {
    const people = P < 0 ? team : [P], W = 560, rowH = 42, L = 80, R = 56;
    let s = `<svg viewBox="0 0 ${W} ${people.length * rowH + 6}" role="img" aria-label="QA status by person">`;
    people.forEach((p, i) => {
      const st = stats(rowsOf(M, p)), y = i * rowH + 4, iw = W - L - R, tot = st.rev + st.na;
      s += `<text x="${L - 8}" y="${y + 17}" text-anchor="end" font-size="12" fill="var(--ink)">${FIRST(model.owners[p])}</text>`;
      if (!tot) { s += `<text x="${L}" y="${y + 17}" font-size="12" fill="var(--muted)">no QA-reviewed cases</text>`; return; }
      let x = L;
      [...st.c, st.na].forEach((c, q) => {
        if (!c) return; const w = iw * c / tot;
        const tt = `<b>${NAME(model.owners[p])}</b> · ${mname(M)}<br>${q === 4 ? 'N/A (not applicable)' : QA[q]}: ${c} of ${tot} reviewed (${pc(c / tot)})` + (q === 4 ? `<br>Failed ${st.nac[2]} · Critical ${st.nac[3]} · Opportunity ${st.nac[1]}` : '');
        s += `<rect x="${x}" y="${y}" width="${Math.max(w - 2, 1)}" height="26" rx="3" fill="${QCOL[q]}" data-tt="${enc(tt)}"/>`;
        if (w > 26) s += `<text x="${x + (w - 2) / 2}" y="${y + 17}" text-anchor="middle" font-size="11" font-weight="600" fill="${q === 1 ? '#0b0b0b' : '#fff'}" pointer-events="none">${c}</text>`;
        x += w;
      });
      s += `<text x="${W - R + 8}" y="${y + 17}" font-size="12" fill="var(--ink2)">n=${tot}</text>`;
    });
    el.innerHTML = s + '</svg>'; bindTips(el);
    $('l3').innerHTML = [...QA, 'N/A (not applicable)'].map((q, i) => `<span><i style="background:${QCOL[i]}"></i>${q}</span>`).join('');
  }

  function typeBars(el, mode) {
    const count = mode === 'count';
    const by = {}, pb = {};
    rowsOf(M, P).forEach(r => { (by[r.t] = by[r.t] || { h: 0, n: 0 }); by[r.t].h += r.h; by[r.t].n++; });
    if (M > 0) rowsOf(M - 1, P).forEach(r => { pb[r.t] = (pb[r.t] || 0) + (count ? 1 : r.h); });
    const val = v => count ? v.n : v.h;
    const list = Object.entries(by).sort((a, b) => val(b[1]) - val(a[1]));
    if (!list.length) { el.innerHTML = '<p class="s">No data.</p>'; return; }
    const W = 560, rowH = 30, L = 150, R = 100, mx = val(list[0][1]);
    let s = `<svg viewBox="0 0 ${W} ${list.length * rowH + 4}" role="img" aria-label="${count ? 'Cases' : 'Hours'} by case type">`;
    list.forEach(([t, v], i) => {
      const y = i * rowH + 2, x = val(v), w = Math.max(3, (W - L - R) * x / mx);
      const tt = `<b>${model.types[t]}</b> · ${mname(M)}<br>${v.n} cases · ${f1(v.h)} h · ${f2(v.h / v.n)} h/case` + (pb[t] != null ? `<br>Previous month: ${count ? pb[t] + ' cases' : f1(pb[t]) + ' h'}` : '');
      s += `<text x="${L - 8}" y="${y + 16}" text-anchor="end" font-size="12" fill="var(--ink)">${model.types[t]}</text><path d="M${L},${y + 2}H${L + w - 4}Q${L + w},${y + 2} ${L + w},${y + 6}V${y + 16}Q${L + w},${y + 20} ${L + w - 4},${y + 20}H${L}Z" fill="${count ? 'var(--p3)' : 'var(--m3)'}" data-tt="${enc(tt)}"/><text x="${L + w + 6}" y="${y + 16}" font-size="12" fill="var(--ink)">${count ? f0(x) : f1(x) + ' h'}</text>`;
      if (pb[t]) { const d = (x - pb[t]) / pb[t]; s += `<text x="${W}" y="${y + 16}" text-anchor="end" font-size="11" fill="var(--ink2)">${d > 0 ? '▲ +' : d < 0 ? '▼ ' : '■ '}${(d * 100).toFixed(0)}%</text>`; }
    });
    el.innerHTML = s + '</svg>'; bindTips(el);
  }

  // Team-wide donut of QA statuses (ignores the person filter)
  function donut(el) {
    const st = stats(rowsOf(M, -1));
    const parts = st.c, names = QA, tot = parts.reduce((a, b) => a + b, 0);
    if (!tot) { el.innerHTML = '<p class="s">No QA-reviewed cases this month.</p>'; return; }
    const cx = 100, cy = 100, R = 92, r = 58;
    let a0 = -Math.PI / 2, s = `<svg viewBox="0 0 200 200" role="img" aria-label="QA status breakdown" style="max-width:230px">`;
    parts.forEach((c, i) => {
      if (!c) return;
      const tt = `<b>${names[i]}</b> · ${mname(M)}<br>${c} of ${tot} (${pc(c / tot)})`;
      if (c === tot) { s += `<circle cx="${cx}" cy="${cy}" r="${(R + r) / 2}" fill="none" stroke="${QCOL[i]}" stroke-width="${R - r}" data-tt="${enc(tt)}"/>`; return; }
      const a1 = a0 + 2 * Math.PI * c / tot, lg = a1 - a0 > Math.PI ? 1 : 0, p = (rad, a) => [cx + rad * Math.cos(a), cy + rad * Math.sin(a)];
      const [x0, y0] = p(R, a0), [x1, y1] = p(R, a1), [x2, y2] = p(r, a1), [x3, y3] = p(r, a0);
      s += `<path d="M${x0},${y0}A${R},${R} 0 ${lg} 1 ${x1},${y1}L${x2},${y2}A${r},${r} 0 ${lg} 0 ${x3},${y3}Z" fill="${QCOL[i]}" stroke="var(--surface)" stroke-width="2" data-tt="${enc(tt)}"/>`;
      a0 = a1;
    });
    s += `<text x="${cx}" y="${cy + 4}" text-anchor="middle" font-size="26" font-weight="600" fill="var(--ink)">${tot}</text><text x="${cx}" y="${cy + 22}" text-anchor="middle" font-size="11" fill="var(--ink2)">QA findings</text></svg>`;
    const legend = parts.map((c, i) => `<div style="display:flex;gap:8px;align-items:center;margin:3px 0"><i style="width:10px;height:10px;border-radius:2px;background:${QCOL[i]};display:inline-block"></i><span style="min-width:130px">${names[i]}</span><b>${c}</b><span style="color:var(--ink2)">${pc(c / tot)}</span></div>`).join('');
    el.innerHTML = `<div style="display:flex;flex-wrap:wrap;gap:16px;align-items:center"><div style="flex:0 1 230px">${s}</div><div style="font-size:13px">${legend}</div></div>`;
    bindTips(el);
  }

  // Team-wide line over the selected months. opts: metric(m) -> number|null, fmt, axis, tt(m), pp (rate in pp)
  function line(el, opts) {
    const mons = [...shown].sort((a, b) => a - b);
    if (!mons.length) { el.innerHTML = '<p class="s">Select at least one month above.</p>'; return; }
    const vals = mons.map(m => opts.metric(m));
    const W = 460, H = 250, L = 44, R = 22, T = 20, B = 30, iw = W - L - R, ih = H - T - B;
    const mx = Math.max(...vals.filter(v => v != null), 0) * 1.15 || 1;
    const X = i => L + (mons.length === 1 ? iw / 2 : iw * i / (mons.length - 1)), Y = v => T + ih - ih * v / mx;
    let s = `<svg viewBox="0 0 ${W} ${H}" role="img" aria-label="${opts.label}">`;
    for (let k = 0; k <= 4; k++) { const y = T + ih - ih * k / 4; s += `<line x1="${L}" x2="${W - R}" y1="${y}" y2="${y}" stroke="${k ? 'var(--grid)' : 'var(--axis)'}"/><text x="${L - 6}" y="${y + 4}" text-anchor="end" font-size="11" fill="var(--muted)">${opts.axis(mx * k / 4)}</text>`; }
    const pts = mons.map((m, i) => vals[i] == null ? null : X(i) + ',' + Y(vals[i])).filter(Boolean);
    if (pts.length > 1) s += `<polyline fill="none" stroke="var(--p1)" stroke-width="2" stroke-linejoin="round" points="${pts.join(' ')}"/>`;
    mons.forEach((m, i) => {
      const lab = `<text x="${X(i)}" y="${H - 10}" text-anchor="middle" font-size="12" fill="var(--ink)">${model.months[m].name.slice(0, 3)}${model.months[m].partial ? '*' : ''}</text>`;
      if (vals[i] == null) { s += lab + `<text x="${X(i)}" y="${T + ih - 6}" text-anchor="middle" font-size="10" fill="var(--muted)">—</text>`; return; }
      const pv = i && vals[i - 1] != null ? vals[i - 1] : null;
      let ch = '';
      if (pv != null && pv !== 0) { const d = opts.pp ? (vals[i] - pv) * 100 : (vals[i] - pv) / pv * 100; ch = `<br>${d > 0 ? '▲ +' : d < 0 ? '▼ ' : '■ '}${d.toFixed(opts.pp ? 1 : 0)}${opts.pp ? ' pp' : '%'} vs ${model.months[mons[i - 1]].name}`; }
      const tt = `<b>${mname(m)}</b><br>${opts.tt(m)}${ch}`;
      s += `<circle cx="${X(i)}" cy="${Y(vals[i])}" r="5" fill="var(--p1)" stroke="var(--surface)" stroke-width="2"/><text x="${X(i)}" y="${Y(vals[i]) - 11}" text-anchor="middle" font-size="12" font-weight="600" fill="var(--ink)">${opts.fmt(vals[i])}</text>${lab}<circle cx="${X(i)}" cy="${Y(vals[i])}" r="16" fill="transparent" data-tt="${enc(tt)}"/>`;
    });
    el.innerHTML = s + '</svg>'; bindTips(el);
  }

  function topCases(el) {
    const list = cases(teamModel(), M, P).slice(0, 10);
    el.innerHTML = '<table><tr><th>ID</th><th>Hours (sum of ID)</th><th>Rows</th><th>People</th><th>Types</th></tr>' + list.map(c => {
      const ppl = [...new Set(c.rows.map(r => FIRST(model.owners[r.o])))].join(', ');
      const tys = [...new Set(c.rows.map(r => model.types[r.t]))].join(' + ');
      return `<tr><td>${c.id || '(no ID)'}</td><td>${f2(c.h)}</td><td>${c.rows.length}</td><td style="text-align:left">${ppl}</td><td style="text-align:left">${tys}</td></tr>`;
    }).join('') + '</table>';
  }

  // ---------- Summary (English)
  function summary() {
    const c = stats(rowsOf(M, P)), pv = M > 0 ? stats(rowsOf(M - 1, P)) : null;
    const vs = pv ? `vs ${model.months[M - 1].name}: ` : '';
    const top = cases(teamModel(), M, P)[0];
    $('h-sum').textContent = `Summary — ${mname(M)} · ${P < 0 ? 'whole team' : NAME(model.owners[P])}`;
    const k = [
      ['Completed cases', f0(c.n), pv ? vs + delta(c.n, pv.n) : ''],
      ['Total hours', f1(c.h), pv ? vs + delta(c.h, pv.h) : ''],
      ['Avg hours per case', f2(c.avg), pv ? vs + delta(c.avg, pv.avg) : ''],
      ['Longest case (by ID)', top ? f2(top.h) + ' h' : '—', top ? `ID ${top.id || 'n/a'} · ${top.rows.length} row(s)` : ''],
      ['Error % (Failed + Critical)', pc(c.err), pv ? vs + dpp(c.err, pv.err) : ''],
      ['Passed %', pc(c.rev ? c.c[0] / c.rev : null), `${c.c[0]} of ${c.rev} reviewed`],
      ['QA coverage', pc(c.cov), `${c.rev + c.na} of ${c.n} cases reviewed`],
      ['Not applicable (N/A)', f0(c.na), `${c.na} of ${c.na + c.c[1] + c.c[2] + c.c[3]} QA findings (Failed, Critical, Opportunity)`]
    ];
    $('kpis').innerHTML = k.map(([t, v, d]) => `<div class="card kpi"><div class="t">${t}</div><div class="v">${v}</div><div class="d">${d}</div></div>`).join('');

    $('s7').textContent = `${mname(M)}${pv ? ', with change vs ' + model.months[M - 1].name : ''}. Error % = (Failed + Critical) ÷ QA-reviewed cases, excluding findings marked N/A; shown with ${BH.MIN_REVIEWED}+ reviewed. Longest case sums all rows sharing an ID.`;
    const body = team.filter(p => P < 0 || p === P).map((p, i) => {
      const s = stats(rowsOf(M, p)), q = M > 0 ? stats(rowsOf(M - 1, p)) : null;
      const t = cases(teamModel(), M, p)[0];
      return `<tr><td><span class="dot" style="background:var(--p${(team.indexOf(p) % 8) + 1})"></span>${NAME(model.owners[p])}</td>
        <td>${s.n}<small>${q ? delta(s.n, q.n) : ''}</small></td>
        <td>${f1(s.h)}<small>${q ? delta(s.h, q.h) : ''}</small></td>
        <td>${f2(s.avg)}<small>${q ? delta(s.avg, q.avg) : ''}</small></td>
        <td>${t ? f2(t.h) : '—'}<small>${t ? 'ID ' + (t.id || 'n/a') : ''}</small></td>
        <td>${s.rev}<small>${pc(s.cov)} of cases</small></td>
        <td>${s.c[0]}</td><td>${s.c[1]}</td><td>${s.c[2]}</td><td>${s.c[3]}</td>
        <td>${pc(s.err)}<small>${q ? dpp(s.err, q.err) : ''}</small></td></tr>`;
    }).join('');
    $('c7').innerHTML = `<table><tr><th>Person</th><th>Cases</th><th>Hours</th><th>Avg h/case</th><th>Longest case (h)</th><th>QA reviewed</th><th>Passed</th><th>Opportunity</th><th>Failed</th><th>Critical</th><th>Error %</th></tr>${body}</table>`;
  }

  function errorsTable() {
    const people = team.filter(p => P < 0 || p === P);
    const sum = people.map(p => stats(rowsOf(M, p)));
    // Only Failed and Critical are errors. Opportunity is listed but never counted as an error.
    const raw = s => ({ f: s.c[2] + s.nac[2], c: s.c[3] + s.nac[3], o: s.c[1] + s.nac[1], na: s.nac[2] + s.nac[3], naO: s.nac[1] });
    const row = (label, r, dot, style) => { const total = r.f + r.c + r.o; return `<tr${style || ''}><td>${dot}${label}</td><td>${r.f}</td><td>${r.c}</td><td>${r.o}</td><td><b>${total}</b></td><td>${r.na}${r.naO ? `<small>+${r.naO} on Opportunity</small>` : ''}</td><td><b>${total - r.o - r.na}</b></td></tr>`; };
    const tot = { f: 0, c: 0, o: 0, na: 0, naO: 0 };
    sum.forEach(s => { const r = raw(s); for (const k in tot) tot[k] += r[k]; });
    $('s8').textContent = `${mname(M)}. Only Failed and Critical count as errors. Errors that count = Total − Opportunity − N/A, where N/A are the Failed/Critical findings whose "QA Fixed Comment" starts with NA (they do not apply).`;
    $('c8').innerHTML = `<table><tr><th>Person</th><th>Failed</th><th>Critical</th><th>Opportunity</th><th>Total</th><th>N/A (Failed + Critical)</th><th>Errors that count</th></tr>` +
      people.map((p, i) => row(NAME(model.owners[p]), raw(sum[i]), `<span class="dot" style="background:var(--p${(team.indexOf(p) % 8) + 1})"></span>`)).join('') +
      (people.length > 1 ? row('Team total', tot, '', ' style="font-weight:600"') : '') + '</table>';
  }

  // ---------- render
  function render() {
    if (!model) return;
    const t = teamOwners(model, false); team = t.team; extra = t.extra;
    controls();
    $('sub').textContent = `${mname(M)} · ${P < 0 ? 'whole team' : NAME(model.owners[P])} · ${team.length} people`;
    const notes = [];
    if (model.months[M].partial) notes.push(`${model.months[M].name} is a partial month: data up to ${model.months[M].last}. Do not compare it with full months without adjusting.`);
    if (extra.length) notes.push(`Excluded ${extra.length} person(s) with fewer than 10 rows: ${extra.map(e => e.owner + ' (' + e.rows + ')').join(', ')}.`);
    $('note').hidden = !notes.length; $('note').textContent = notes.join(' ');

    grouped($('c1'), rs => rs.length, f0, { label: 'Cases by person and month' });
    grouped($('c2'), rs => stats(rs).h, f1, { label: 'Hours by person and month', tt: s => `${f1(s.h)} h · ${s.n} cases` });
    grouped($('c4'), rs => stats(rs).err, pc, { label: 'Error %', axis: v => (v * 100).toFixed(0) + '%', tt: s => `${pc(s.err)} · ${s.c[2]} Failed + ${s.c[3]} Critical of ${s.rev} reviewed` });
    $('lmonths').innerHTML = [...shown].sort((a, b) => a - b).map(i => `<span><i style="background:${mcol(i)}"></i>${mname(i)}</span>`).join('');

    $('s3').textContent = `${mname(M)} · QA-reviewed cases; bar width = share of each person's reviewed cases.`;
    statusBars($('c3'));
    $('s5').textContent = `${mname(M)} · hours by case type${M > 0 ? ' (change vs ' + model.months[M - 1].name + ')' : ''}.`;
    typeBars($('c5'), 'hours');
    $('s11').textContent = `${mname(M)} · number of cases (rows) by case type${M > 0 ? ' (change vs ' + model.months[M - 1].name + ')' : ''}.`;
    typeBars($('c11'), 'count');
    $('s10').textContent = `${mname(M)} · whole team (not filtered by person). QA findings that apply (N/A already removed).`;
    donut($('c10'));
    line($('c9'), { label: 'Total cases by month', metric: m => rowsOf(m, -1).length, fmt: f0, axis: f0, tt: m => `${rowsOf(m, -1).length} cases` });
    line($('c12'), { label: 'Team error % by month', pp: true, metric: m => stats(rowsOf(m, -1)).err, fmt: pc, axis: v => (v * 100).toFixed(0) + '%',
      tt: m => { const st = stats(rowsOf(m, -1)); return `${pc(st.err)} · ${st.errors} errors (Failed + Critical) of ${st.rev} QA-reviewed (N/A excluded)`; } });
    $('s6').textContent = `${mname(M)} · a case's hours add up every row sharing the same ID (e.g. DDC Site Services + SS Support). Case counts stay per row.`;
    topCases($('c6'));
    summary();
    errorsTable();
    $('foot').textContent = `${model.rows.length} rows loaded. Each row counts as one case.`;
  }

  // restore last upload, if the browser kept it
  try {
    const saved = localStorage.getItem(STORE);
    if (saved) setModel(JSON.parse(saved), 'Data restored from the last upload in this browser');
  } catch (e) { /* ignore */ }
})();
