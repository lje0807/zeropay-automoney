(function(){
  'use strict';
  const HOLIDAYS = Recon.HOLIDAYS;
  const LS = 'recon.holidays.v1';
  let custom = [], removed = [];
  try { const s = JSON.parse(localStorage.getItem(LS) || '{}'); custom = s.custom || []; removed = s.removed || []; } catch (e) {}
  const saveLS = () => { try { localStorage.setItem(LS, JSON.stringify({ custom, removed })); } catch (e) {} };

  const $ = id => document.getElementById(id);
  const fmt = n => n.toLocaleString('ko-KR');
  const esc = s => String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const state = { bankFile: null, voucherFile: null, bank: null, voucher: null, res: null };

  // 기준표
  $('rulesBody').innerHTML = Recon.ITEMS.map(it => {
    const bankName = { p_buy: '입금 · 개인 구매금', c_buy: '입금 · 법인 구매금', p_ref: '출금 · 개인 환불금', c_ref: '출금 · 법인 환불금', qr_col: '입금 · QR추심/기타', qr_set: '출금 · QR정산/기타', pg: '출금 · PG정산' }[it.key];
    return `<tr><td>${it.name}</td><td>${bankName}</td><td>${it.src}</td><td><span class="code">${it.lag ? 'D-1' : 'D'}</span></td></tr>`;
  }).join('');

  function holidayList() {
    const s = $('start').value, e = $('end').value;
    const base = Object.keys(HOLIDAYS).filter(d => !removed.includes(d));
    const all = [...new Set([...base, ...custom])].sort();
    return { all, shown: all.filter(d => (!s || d >= s) && (!e || d <= e)) };
  }
  function renderChips() {
    const { shown } = holidayList();
    $('chips').innerHTML = shown.length ? shown.map(d => {
      const name = HOLIDAYS[d] || '직접 추가';
      const w = Recon.weekday(d);
      return `<span class="chip${HOLIDAYS[d] ? '' : ' custom'}" title="${esc(name)}">${Recon.short(d)}(${w}) ${esc(name)}<button type="button" data-d="${d}" aria-label="${d} 휴일 제외">×</button></span>`;
    }).join('') : '<span class="hint" style="margin:0">선택한 기간에 등록된 휴일이 없어요.</span>';
  }
  $('chips').addEventListener('click', e => {
    const d = e.target.closest('button')?.dataset.d; if (!d) return;
    if (custom.includes(d)) custom = custom.filter(x => x !== d); else removed.push(d);
    saveLS(); renderChips(); if (state.res) run();
  });
  $('addBtn').addEventListener('click', () => {
    const d = $('addDate').value; if (!d) return;
    removed = removed.filter(x => x !== d);
    if (!HOLIDAYS[d] && !custom.includes(d)) custom.push(d);
    $('addDate').value = ''; saveLS(); renderChips(); if (state.res) run();
  });
  ['start', 'end'].forEach(id => $(id).addEventListener('change', renderChips));

  async function readWb(file) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(await file.arrayBuffer());
    return wb;
  }
  function setDrop(kind, cls, text) {
    const el = $(kind === 'bank' ? 'dropBank' : 'dropVoucher');
    el.classList.remove('done', 'err'); if (cls) el.classList.add(cls);
    $(kind === 'bank' ? 'nameBank' : 'nameVoucher').textContent = text;
  }
  async function takeFile(kind, file) {
    if (!file) return;
    setDrop(kind, '', '읽는 중…');
    try {
      const wb = await readWb(file);
      if (kind === 'bank') {
        const b = Recon.parseBank(wb); const ds = Object.keys(b.data).sort();
        if (!ds.length) throw new Error('날짜 행을 찾지 못했어요.');
        state.bank = b; state.bankFile = file;
        $('start').value = ds[0]; $('end').value = ds[ds.length - 1];
        setDrop(kind, 'done', `${file.name} · ${Recon.short(ds[0])}~${Recon.short(ds[ds.length - 1])} ${ds.length}일`);
      } else {
        const v = Recon.parseVoucher(wb); const ds = Object.keys(v.data).sort();
        if (!ds.length) throw new Error('날짜 행을 찾지 못했어요.');
        state.voucher = v; state.voucherFile = file;
        setDrop(kind, 'done', `${file.name} · ${Recon.short(ds[0])}~${Recon.short(ds[ds.length - 1])} ${ds.length}일`);
      }
    } catch (err) {
      if (kind === 'bank') state.bank = null; else state.voucher = null;
      setDrop(kind, 'err', `${file.name} · ${err.message || '파일을 읽지 못했어요.'}`);
    }
    renderChips();
    $('run').disabled = !(state.bank && state.voucher);
  }
  [['bank', 'fileBank', 'dropBank'], ['voucher', 'fileVoucher', 'dropVoucher']].forEach(([k, inp, drop]) => {
    $(inp).addEventListener('change', e => takeFile(k, e.target.files[0]));
    const d = $(drop);
    d.addEventListener('dragover', e => { e.preventDefault(); d.classList.add('over'); });
    d.addEventListener('dragleave', () => d.classList.remove('over'));
    d.addEventListener('drop', e => { e.preventDefault(); d.classList.remove('over'); takeFile(k, e.dataTransfer.files[0]); });
  });

  function run() {
    $('msg').className = 'msg'; $('msg').textContent = '';
    const s = $('start').value, e = $('end').value;
    if (s && e && s > e) { $('msg').className = 'msg err'; $('msg').textContent = '시작일이 종료일보다 늦어요.'; return; }
    try {
      state.res = Recon.reconcile(state.bank, state.voucher, { start: s, end: e, holidays: holidayList().all });
      render(state.res);
    } catch (err) { $('msg').className = 'msg err'; $('msg').textContent = err.message; }
  }
  $('form').addEventListener('submit', e => { e.preventDefault(); run(); });

  function render(res) {
    const I = Recon.ITEMS;
    const bad = res.mismatches.filter(m => !m.open);
    const pill = bad.length ? `<span class="pill bad">불일치 영업일 ${bad.length}일</span>` : `<span class="pill ok">전 영업일 일치</span>`;
    const warnPill = res.warnings.length ? ` <span class="pill warn">확인 필요 ${res.warnings.length}건</span>` : '';
    const sumRows = I.map(it => { const t = res.totals[it.key]; return `<tr><td>${it.name} <span class="code">${it.lag ? 'D-1' : 'D'}</span></td><td class="num">${fmt(t.b)}</td><td class="num">${fmt(t.v)}</td><td class="num${t.diff ? ' neg' : ''}">${fmt(t.diff)}</td></tr>`; }).join('');
    const misRows = bad.length ? bad.map(m => `<tr class="badrow"><td>${Recon.short(m.date)} (${Recon.weekday(m.date)})</td><td class="lbl">${m.label}</td><td style="text-align:left">${I.filter(it => m.diff[it.key]).map(it => `${it.name} <span class="num neg" style="font-family:var(--mono)">${fmt(m.diff[it.key])}</span>`).join('<br>')}</td></tr>`).join('') : `<tr><td colspan="3" class="empty">불일치 영업일이 없어요.</td></tr>`;
    const bankOnly = res.bankOnly.filter(o => o.total).map(o => `<tr><td>${o.name}</td><td class="num">${fmt(o.total)}</td><td class="lbl">${o.days.map(x => `${Recon.short(x.date)} ${fmt(x.amt)}`).join(', ')}</td></tr>`).join('') || `<tr><td colspan="3" class="empty">해당 기간에 금액이 없어요.</td></tr>`;
    const daily = res.rows.map(r => {
      if (!r.closes) return `<tr class="hol"><td>${Recon.short(r.date)} (${Recon.weekday(r.date)})</td><td class="lbl">→ ${Recon.short(r.block[r.block.length - 1])} 합산</td>${I.map(() => '<td></td>').join('')}<td></td></tr>`;
      const tag = r.open ? '<span class="tag ref">참고</span>' : r.ok ? '<span class="tag ok">일치</span>' : '<span class="tag bad">불일치</span>';
      return `<tr class="${r.ok ? '' : 'badrow'}"><td>${Recon.short(r.date)} (${Recon.weekday(r.date)})</td><td class="lbl">${r.label}</td>${I.map(it => `<td class="num${r.diff[it.key] ? ' neg' : ''}">${r.diff[it.key] ? fmt(r.diff[it.key]) : '0'}</td>`).join('')}<td style="text-align:center">${tag}</td></tr>`;
    }).join('');
    $('result').innerHTML = `
      <div class="card">
        <div class="status">
          <div>${pill}${warnPill}<div class="period" style="margin-top:6px">통장일 ${res.start} ~ ${res.end} · ${res.dates.length}일</div></div>
          <button type="button" class="btn dl" id="dlBtn">대조 엑셀 다운로드</button>
        </div>
        <p class="msg" id="dlMsg" role="status"></p>
        ${res.warnings.length ? `<ul class="warns" style="margin-top:12px">${res.warnings.map(w => `<li>${esc(w)}</li>`).join('')}</ul>` : ''}
      </div>
      <div class="row2">
        <div class="card"><h3>항목별 합계</h3><div class="tw"><table><thead><tr><th>항목</th><th>통장</th><th>상품권</th><th>차액</th></tr></thead><tbody>${sumRows}</tbody></table></div></div>
        <div class="card"><h3>불일치 영업일</h3><div class="tw"><table><thead><tr><th>통장일</th><th style="text-align:left">합산기간</th><th style="text-align:left">항목 · 차액</th></tr></thead><tbody>${misRows}</tbody></table></div></div>
      </div>
      <div class="card"><h3>일별 대조 (통장 − 상품권)</h3><div class="tw daily"><table><thead><tr><th>통장일</th><th style="text-align:left">합산기간</th>${I.map(it => `<th>${it.name}</th>`).join('')}<th style="text-align:center">결과</th></tr></thead><tbody>${daily}</tbody></table></div></div>
      <div class="card"><h3>통장 단독 항목 (대사 제외)</h3><div class="tw"><table><thead><tr><th>항목</th><th>합계</th><th style="text-align:left">일자별</th></tr></thead><tbody>${bankOnly}</tbody></table></div></div>`;
    $('dlBtn').addEventListener('click', download);
  }

  async function download() {
    const btn = $('dlBtn'), m = $('dlMsg');
    btn.disabled = true; const label = btn.textContent; btn.innerHTML = '<span class="spin"></span> 만드는 중';
    m.className = 'msg'; m.textContent = '';
    try {
      const res = state.res;
      const wb = await Recon.buildWorkbook(ExcelJS, res, { bankName: state.bankFile.name, voucherName: state.voucherFile.name, createdAt: new Date().toLocaleString('ko-KR') });
      const buf = await wb.xlsx.writeBuffer();
      const filename = Recon.fileName(res);
      const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a'); a.href = url; a.download = filename;
      document.body.appendChild(a); a.click(); a.remove();
      setTimeout(() => URL.revokeObjectURL(url), 10000);
      m.className = 'msg ok'; m.textContent = `${filename} 파일을 내려받았어요.`;
    } catch (err) {
      m.className = 'msg err';
      m.textContent = `엑셀을 만들지 못했어요. ${err.message || ''}`;
    } finally { btn.disabled = false; btn.textContent = label; }
  }

  renderChips();
})();
