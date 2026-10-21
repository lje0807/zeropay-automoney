// ===== 자금대사 엔진 (브라우저/Node 공용) =====
(function (root) {
  'use strict';

  const norm = s => String(s == null ? '' : (typeof s === 'object' && s.richText) ? s.richText.map(t => t.text).join('') : s)
    .replace(/[\s/·]/g, '');

  const pad = n => String(n).padStart(2, '0');
  const keyOf = (y, m, d) => `${y}-${pad(m)}-${pad(d)}`;
  function parseDate(v) {
    if (v == null) return null;
    if (v instanceof Date) return keyOf(v.getUTCFullYear(), v.getUTCMonth() + 1, v.getUTCDate());
    if (typeof v === 'number' && v > 30000 && v < 80000) { // 엑셀 일련번호
      const d = new Date(Math.round((v - 25569) * 86400000));
      return keyOf(d.getUTCFullYear(), d.getUTCMonth() + 1, d.getUTCDate());
    }
    const m = String(v).trim().match(/^(\d{4})[.\-/년\s]+(\d{1,2})[.\-/월\s]+(\d{1,2})/);
    return m ? keyOf(+m[1], +m[2], +m[3]) : null;
  }
  const toDate = k => { const [y, m, d] = k.split('-').map(Number); return new Date(Date.UTC(y, m - 1, d)); };
  const addDays = (k, n) => { const d = toDate(k); d.setUTCDate(d.getUTCDate() + n); return keyOf(d.getUTCFullYear(), d.getUTCMonth() + 1, d.getUTCDate()); };
  const WD = ['일', '월', '화', '수', '목', '금', '토'];
  const weekday = k => WD[toDate(k).getUTCDay()];
  const short = k => k.slice(5).replace('-', '.');
  const num = v => {
    if (v == null || v === '') return 0;
    if (typeof v === 'object') { if ('result' in v) return num(v.result); return 0; }
    const n = Number(String(v).replace(/,/g, ''));
    return isFinite(n) ? n : 0;
  };

  // 시트를 2차원 배열로 (병합셀은 ExcelJS가 마스터 값을 돌려줌)
  function grid(ws) {
    const rows = [];
    ws.eachRow({ includeEmpty: true }, (row, r) => {
      const arr = [];
      for (let c = 1; c <= ws.columnCount; c++) arr.push(row.getCell(c).value);
      rows[r - 1] = arr;
    });
    for (let i = 0; i < rows.length; i++) if (!rows[i]) rows[i] = [];
    return rows;
  }
  const ffill = arr => { let last = ''; return arr.map(v => { const n = norm(v); if (n) last = n; return last; }); };

  // ---------- 통장(발행처) ----------
  const BANK_FIELDS = [
    { key: 'p_buy', g: '입금', l: '개인구매금' },
    { key: 'c_buy', g: '입금', l: '법인구매금' },
    { key: 'disc', g: '입금', l: '할인지원금' },
    { key: 'point', g: '입금', l: '적립금' },
    { key: 'auto', g: '입금', l: '자동생성대금' },
    { key: 'qr_col', g: '입금', l: 'QR추심기타' },
    { key: 'trust_in', g: '입금', l: '신탁이체이자' },
    { key: 'in_total', g: '입금', l: '입금총액' },
    { key: 'p_ref', g: '출금', l: '개인환불금' },
    { key: 'c_ref', g: '출금', l: '법인환불금' },
    { key: 'pg', g: '출금', l: 'PG정산' },
    { key: 'qr_set', g: '출금', l: 'QR정산기타' },
    { key: 'trust_out', g: '출금', l: '신탁이체이자' },
    { key: 'out_total', g: '출금', l: '출금총액' },
    { key: 'balance', g: '잔액', l: '' },
  ];
  function parseBank(wb) {
    for (const ws of wb.worksheets) {
      const g = grid(ws);
      const h = g.findIndex(r => norm(r[0]) === '일자');
      if (h < 0) continue;
      const groups = ffill(g[h]);
      // 라벨: 헤더 아래 날짜가 나오기 전까지의 행을 이어붙임
      let first = h + 1; const labelRows = [];
      while (first < g.length && !parseDate(g[first][0]) && !/잔액/.test(norm(g[first][0])) && labelRows.length < 3) { labelRows.push(g[first]); first++; }
      // 세로 병합셀은 아래 행에도 같은 값이 반복되므로 중복은 건너뜀
      const labels = g[h].map((_, c) => labelRows.map(r => norm(r[c])).filter((v, i, a) => v && v !== a[i - 1]).join(''));
      const col = {};
      for (const f of BANK_FIELDS) {
        const idx = groups.findIndex((gr, c) => c > 0 && gr === f.g && (f.l === '' || labels[c] === f.l));
        col[f.key] = idx;
      }
      const missing = BANK_FIELDS.filter(f => col[f.key] < 0 && f.key !== 'balance').map(f => `${f.g}·${f.l}`);
      const data = {}; let openBalance = null; let account = '';
      for (let r = 0; r < h; r++) { const t = g[r].map(norm).find(Boolean); if (t && /계좌/.test(t)) account = String(g[r].find(v => v != null)); }
      for (let r = first; r < g.length; r++) {
        const row = g[r];
        if (/전월잔액/.test(norm(row[0]))) { openBalance = col.balance >= 0 ? num(row[col.balance]) : null; continue; }
        const d = parseDate(row[0]);
        if (!d) continue;
        const rec = {};
        for (const f of BANK_FIELDS) rec[f.key] = col[f.key] >= 0 ? num(row[col[f.key]]) : 0;
        data[d] = rec;
      }
      return { data, openBalance, account, missing, sheet: ws.name };
    }
    throw new Error('통장(발행처) 파일에서 "일자" 머리글을 찾지 못했어요. 은행 거래 집계 엑셀이 맞는지 확인해 주세요.');
  }

  // ---------- 상품권 ----------
  const V_FIELDS = [
    { key: 'p_buy', g: /^개인구매\(구매일\)/, l: '고객구매금' },
    { key: 'c_buy', g: /^법인구매\(입금완료일\)/, l: '법인구매금' },
    { key: 'p_ref', g: /^개인환불\(신청일\)/, l: '고객금액' },
    { key: 'c_ref', g: /^법인환불\(관리자승인일\)/, l: '고객금액' },
    { key: 'qr_set', g: /^QR결제/, l: '정산' },
    { key: 'qr_col', g: /^QR결제/, l: '추심' },
    { key: 'pg', g: /^온라인결제/, l: '정산' },
  ];
  function parseVoucher(wb) {
    for (const ws of wb.worksheets) {
      const g = grid(ws);
      const h = g.findIndex(r => norm(r[0]) === '날짜');
      if (h < 0) continue;
      const groups = ffill(g[h]);
      const labels = (g[h + 1] || []).map(norm);
      const col = {};
      for (const f of V_FIELDS) col[f.key] = groups.findIndex((gr, c) => c > 0 && f.g.test(gr) && labels[c] === f.l);
      const missing = V_FIELDS.filter(f => col[f.key] < 0).map(f => `${f.g.source.replace(/[\^\\]/g, '')}·${f.l}`);
      let title = '';
      for (let r = 0; r < h; r++) { const v = g[r].find(x => x != null); if (v) { title = String(v); break; } }
      const data = {};
      for (let r = h + 2; r < g.length; r++) {
        const d = parseDate(g[r][0]); if (!d) continue;
        const rec = {};
        for (const f of V_FIELDS) rec[f.key] = col[f.key] >= 0 ? num(g[r][col[f.key]]) : 0;
        data[d] = rec;
      }
      return { data, title, missing, sheet: ws.name };
    }
    throw new Error('상품권 파일에서 "날짜" 머리글을 찾지 못했어요. 일별 상품권 잔액집계 엑셀이 맞는지 확인해 주세요.');
  }

  // ---------- 대사 ----------
  const ITEMS = [
    { key: 'p_buy', name: '개인구매금', lag: 1, src: '개인구매(구매일) 고객구매금' },
    { key: 'c_buy', name: '법인구매금', lag: 1, src: '법인구매(입금완료일) 법인구매금' },
    { key: 'p_ref', name: '개인환불금', lag: 1, src: '개인환불(신청일) 고객금액' },
    { key: 'c_ref', name: '법인환불금', lag: 1, src: '법인환불(관리자승인일) 고객금액' },
    { key: 'qr_col', name: 'QR추심', lag: 0, src: 'QR 결제 추심' },
    { key: 'qr_set', name: 'QR정산', lag: 0, src: 'QR 결제 정산' },
    { key: 'pg', name: 'PG정산(온라인)', lag: 0, src: '온라인 결제 정산' },
  ];

  function reconcile(bank, voucher, opts) {
    const holidays = new Set(opts.holidays || []);
    const isBiz = k => { const w = toDate(k).getUTCDay(); return w !== 0 && w !== 6 && !holidays.has(k); };
    const allDates = Object.keys(bank.data).sort();
    const start = opts.start || allDates[0], end = opts.end || allDates[allDates.length - 1];
    const dates = allDates.filter(d => d >= start && d <= end);
    if (!dates.length) throw new Error('선택한 기간에 통장 데이터가 없어요.');
    const warnings = [];
    if (bank.missing.length) warnings.push('통장 파일에서 찾지 못한 열: ' + bank.missing.join(', ') + ' (0원으로 처리)');
    if (voucher.missing.length) warnings.push('상품권 파일에서 찾지 못한 열: ' + voucher.missing.join(', ') + ' (0원으로 처리)');

    // 상품권 금액을 통장일 기준으로 정렬
    const missingV = new Set();
    const vAligned = {};
    for (const d of dates) {
      vAligned[d] = {};
      for (const it of ITEMS) {
        const src = addDays(d, -it.lag);
        const rec = voucher.data[src];
        if (!rec) missingV.add(src);
        vAligned[d][it.key] = rec ? rec[it.key] : 0;
      }
    }
    if (missingV.size) warnings.push('상품권 파일에 없는 거래일: ' + [...missingV].sort().map(short).join(', ') + ' (0원으로 처리했어요. D-1 항목 때문에 시작일 전날 데이터도 필요해요)');

    // 영업일 묶음: 휴일은 다음 영업일에 합산
    const rows = []; let pending = [];
    for (const d of dates) {
      const biz = isBiz(d);
      pending.push(d);
      if (biz) {
        const block = pending; pending = [];
        for (const x of block) rows.push({ date: x, biz: isBiz(x), block });
      }
    }
    if (pending.length) {
      for (const x of pending) rows.push({ date: x, biz: false, block: pending, open: true });
      warnings.push(`종료일(${short(end)})이 휴일이라 ${short(pending[0])}~${short(end)} 금액은 다음 영업일과 합산되지 않았어요. 차액은 참고용이에요.`);
    }
    if (dates[0] === start && !isBiz(addDays(start, -1)) && bank.data[addDays(start, -1)] === undefined)
      warnings.push(`시작일 전날(${short(addDays(start, -1))})이 휴일이에요. 그 앞 휴일분이 ${short(start)}에 합산돼 있으면 차액이 생길 수 있어요.`);

    for (const r of rows) {
      r.closes = r.biz || (r.open && r.date === r.block[r.block.length - 1]);
      r.label = r.block.length > 1 ? `${short(r.block[0])}~${short(r.block[r.block.length - 1])}` : short(r.date);
      if (r.closes) {
        r.diff = {}; r.bSum = {}; r.vSum = {};
        for (const it of ITEMS) {
          const b = r.block.reduce((s, x) => s + bank.data[x][it.key], 0);
          const v = r.block.reduce((s, x) => s + vAligned[x][it.key], 0);
          r.bSum[it.key] = b; r.vSum[it.key] = v; r.diff[it.key] = b - v;
        }
        // QR 상계: 추심분이 가맹점 정산 지급액에서 빠지면 QR추심·QR정산 차액이 같은 금액으로 생긴다. 이 경우는 일치로 본다
        const qrOffset = r.diff.qr_col !== 0 && r.diff.qr_col === r.diff.qr_set
          && ITEMS.every(it => it.key === 'qr_col' || it.key === 'qr_set' || r.diff[it.key] === 0);
        r.offset = qrOffset ? r.diff.qr_col : 0;
        r.ok = qrOffset || ITEMS.every(it => r.diff[it.key] === 0);
      }
    }

    const totals = {};
    for (const it of ITEMS) {
      const b = dates.reduce((s, d) => s + bank.data[d][it.key], 0);
      const v = dates.reduce((s, d) => s + vAligned[d][it.key], 0);
      totals[it.key] = { b, v, diff: b - v };
    }
    const bankOnly = ['disc', 'point', 'auto', 'trust_in', 'trust_out'].map(k => ({
      key: k, name: { disc: '할인지원금', point: '적립금', auto: '자동생성대금', trust_in: '신탁이체/이자(입금)', trust_out: '신탁이체/이자(출금)' }[k],
      total: dates.reduce((s, d) => s + bank.data[d][k], 0),
      days: dates.filter(d => bank.data[d][k]).map(d => ({ date: d, amt: bank.data[d][k] })),
    }));
    const mismatches = rows.filter(r => r.closes && !r.ok);
    const offsets = rows.filter(r => r.closes && r.offset);
    const usedHolidays = dates.filter(d => !isBiz(d));
    return { start: dates[0], end: dates[dates.length - 1], dates, rows, vAligned, totals, bankOnly, mismatches, offsets, warnings, usedHolidays, bank, voucher };
  }

  // ---------- 엑셀 생성 ----------
  async function buildWorkbook(ExcelJS, res, meta) {
    const wb = new ExcelJS.Workbook();
    wb.creator = '상품권 자금대사';
    wb.created = new Date();
    const FMT = '#,##0;[Red]-#,##0;0';
    const thin = { style: 'thin', color: { argb: 'FFBFC7D1' } };
    const border = { top: thin, left: thin, bottom: thin, right: thin };
    const fill = argb => ({ type: 'pattern', pattern: 'solid', fgColor: { argb } });
    const HEAD = fill('FF1F3A5F'), SUB = fill('FFDCE4EE'), HOL = fill('FFF3F4F6'), BAD = fill('FFFDE2E1'), GOOD = fill('FFE6F4EA'), OFFSET = fill('FFFDF3DC');
    const head = (cell, sub) => { cell.font = { bold: true, color: { argb: sub ? 'FF1F2937' : 'FFFFFFFF' } }; cell.fill = sub ? SUB : HEAD; cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true }; cell.border = border; };
    const colL = n => { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
    const B = res.bank.data;
    const firstRow = 4;
    const rowOf = {}; res.dates.forEach((d, i) => { rowOf[d] = firstRow + i; });

    // 1. 통장
    const s1 = wb.addWorksheet('1. 통장', { views: [{ state: 'frozen', xSplit: 2, ySplit: 3 }] });
    const bCols = [
      ['입금', '개인 구매금', 'p_buy'], ['입금', '법인 구매금', 'c_buy'], ['입금', '할인지원금', 'disc'], ['입금', '적립금', 'point'],
      ['입금', '자동생성대금', 'auto'], ['입금', 'QR추심/기타', 'qr_col'], ['입금', '신탁이체/이자', 'trust_in'], ['입금', '입금총액', 'in_total'],
      ['출금', '개인 환불금', 'p_ref'], ['출금', '법인 환불금', 'c_ref'], ['출금', 'PG정산', 'pg'], ['출금', 'QR정산/기타', 'qr_set'],
      ['출금', '신탁이체/이자', 'trust_out'], ['출금', '출금총액', 'out_total'], ['잔액', '잔액', 'balance'],
    ];
    s1.getCell('A1').value = `통장(발행처) 기준 ${res.bank.account ? '· ' + res.bank.account : ''}`; s1.getCell('A1').font = { bold: true, size: 12 };
    s1.mergeCells('A2:A3'); s1.getCell('A2').value = '통장일'; head(s1.getCell('A2'));
    s1.mergeCells('B2:B3'); s1.getCell('B2').value = '요일'; head(s1.getCell('B2'));
    s1.mergeCells('C2:J2'); s1.getCell('C2').value = '입금'; head(s1.getCell('C2'));
    s1.mergeCells('K2:P2'); s1.getCell('K2').value = '출금'; head(s1.getCell('K2'));
    s1.mergeCells('Q2:Q3'); s1.getCell('Q2').value = '잔액'; head(s1.getCell('Q2'));
    bCols.forEach((c, i) => { if (c[2] !== 'balance') { const cell = s1.getCell(3, 3 + i); cell.value = c[1]; head(cell, true); } });
    const bColIdx = {}; bCols.forEach((c, i) => { bColIdx[c[2]] = 3 + i; });
    res.dates.forEach(d => {
      const r = rowOf[d]; const row = s1.getRow(r);
      row.getCell(1).value = toDate(d); row.getCell(1).numFmt = 'yyyy-mm-dd';
      row.getCell(2).value = weekday(d);
      bCols.forEach((c, i) => { const cell = row.getCell(3 + i); cell.value = B[d][c[2]]; cell.numFmt = FMT; });
      const isHol = res.usedHolidays.includes(d);
      for (let c = 1; c <= 17; c++) { row.getCell(c).border = border; if (isHol) row.getCell(c).fill = HOL; }
    });
    s1.getColumn(1).width = 12; s1.getColumn(2).width = 6; for (let c = 3; c <= 17; c++) s1.getColumn(c).width = 15;

    // 2. 상품권 (통장일 기준 정렬)
    const s2 = wb.addWorksheet('2. 상품권', { views: [{ state: 'frozen', xSplit: 2, ySplit: 3 }] });
    s2.getCell('A1').value = `상품권 거래 · 통장일 기준 정렬 (구매·환불은 D-1, QR·온라인은 D) ${res.voucher.title ? '· ' + res.voucher.title : ''}`; s2.getCell('A1').font = { bold: true, size: 12 };
    s2.mergeCells('A2:A3'); s2.getCell('A2').value = '통장일'; head(s2.getCell('A2'));
    s2.mergeCells('B2:B3'); s2.getCell('B2').value = '요일'; head(s2.getCell('B2'));
    ITEMS.forEach((it, i) => {
      const c1 = s2.getCell(2, 3 + i); c1.value = it.name; head(c1);
      const c2 = s2.getCell(3, 3 + i); c2.value = `${it.lag ? 'D-1' : 'D'} · ${it.src}`; head(c2, true);
    });
    const tdc = 3 + ITEMS.length;
    s2.mergeCells(2, tdc, 3, tdc); s2.getCell(2, tdc).value = 'D-1 거래일'; head(s2.getCell(2, tdc));
    res.dates.forEach(d => {
      const r = rowOf[d]; const row = s2.getRow(r);
      row.getCell(1).value = toDate(d); row.getCell(1).numFmt = 'yyyy-mm-dd';
      row.getCell(2).value = weekday(d);
      ITEMS.forEach((it, i) => { const c = row.getCell(3 + i); c.value = res.vAligned[d][it.key]; c.numFmt = FMT; });
      const pc = row.getCell(tdc); pc.value = toDate(addDays(d, -1)); pc.numFmt = 'yyyy-mm-dd';
      const isHol = res.usedHolidays.includes(d);
      for (let c = 1; c <= tdc; c++) { row.getCell(c).border = border; if (isHol) row.getCell(c).fill = HOL; }
    });
    s2.getColumn(1).width = 12; s2.getColumn(2).width = 6; for (let c = 3; c < tdc; c++) s2.getColumn(c).width = 20; s2.getColumn(tdc).width = 12;
    s2.getRow(3).height = 32;

    // 3. 대조 (통장 − 상품권, 휴일은 다음 영업일에 합산)
    const s3 = wb.addWorksheet('3. 대조', { views: [{ state: 'frozen', xSplit: 3, ySplit: 3 }] });
    s3.getCell('A1').value = '대조 (통장 − 상품권) · 휴일 금액은 다음 영업일 행에 합산'; s3.getCell('A1').font = { bold: true, size: 12 };
    ['통장일', '요일', '합산기간'].forEach((t, i) => { s3.mergeCells(2, 1 + i, 3, 1 + i); s3.getCell(2, 1 + i).value = t; head(s3.getCell(2, 1 + i)); });
    s3.mergeCells(2, 4, 2, 3 + ITEMS.length); s3.getCell(2, 4).value = '차액'; head(s3.getCell(2, 4));
    ITEMS.forEach((it, i) => { const c = s3.getCell(3, 4 + i); c.value = it.name; head(c, true); });
    const stc = 4 + ITEMS.length;
    s3.mergeCells(2, stc, 3, stc); s3.getCell(2, stc).value = '결과'; head(s3.getCell(2, stc));
    const bankKeyCol = { p_buy: bColIdx.p_buy, c_buy: bColIdx.c_buy, p_ref: bColIdx.p_ref, c_ref: bColIdx.c_ref, qr_col: bColIdx.qr_col, qr_set: bColIdx.qr_set, pg: bColIdx.pg };
    res.rows.forEach(r => {
      const rr = rowOf[r.date]; const row = s3.getRow(rr);
      row.getCell(1).value = toDate(r.date); row.getCell(1).numFmt = 'yyyy-mm-dd';
      row.getCell(2).value = weekday(r.date);
      if (r.closes) {
        row.getCell(3).value = r.label;
        const r1 = rowOf[r.block[0]], r2 = rowOf[r.block[r.block.length - 1]];
        ITEMS.forEach((it, i) => {
          const bc = colL(bankKeyCol[it.key]), vc = colL(3 + i);
          const f = r1 === r2 ? `'1. 통장'!${bc}${r1}-'2. 상품권'!${vc}${r1}` : `SUM('1. 통장'!${bc}${r1}:${bc}${r2})-SUM('2. 상품권'!${vc}${r1}:${vc}${r2})`;
          const c = row.getCell(4 + i); c.value = { formula: f, result: r.diff[it.key] }; c.numFmt = FMT;
          if (r.diff[it.key] !== 0) c.fill = r.offset ? OFFSET : BAD;
        });
        const st = row.getCell(stc);
        st.value = r.open ? '참고(합산 미완료)' : r.offset ? '일치(QR 상계)' : r.ok ? '일치' : '불일치';
        st.fill = r.open ? HOL : r.ok ? GOOD : BAD; st.alignment = { horizontal: 'center' };
      } else {
        const close = r.block[r.block.length - 1];
        row.getCell(3).value = `→ ${short(close)} 합산`;
        row.getCell(3).font = { color: { argb: 'FF6B7280' } };
        for (let c = 1; c <= stc; c++) row.getCell(c).fill = HOL;
      }
      for (let c = 1; c <= stc; c++) row.getCell(c).border = border;
    });
    // 합계 행
    const lastR = firstRow + res.dates.length - 1, sumR = lastR + 2;
    s3.getCell(sumR, 1).value = '기간 합계'; s3.getCell(sumR, 1).font = { bold: true };
    s3.mergeCells(sumR, 1, sumR, 3);
    ITEMS.forEach((it, i) => {
      const L = colL(4 + i); const c = s3.getCell(sumR, 4 + i);
      c.value = { formula: `SUM(${L}${firstRow}:${L}${lastR})`, result: res.rows.filter(r => r.closes).reduce((s, r) => s + r.diff[it.key], 0) };
      c.numFmt = FMT; c.font = { bold: true }; c.border = border;
    });
    s3.getColumn(1).width = 12; s3.getColumn(2).width = 6; s3.getColumn(3).width = 15;
    for (let c = 4; c < stc; c++) s3.getColumn(c).width = 15; s3.getColumn(stc).width = 16;

    // 4. 요약
    const s4 = wb.addWorksheet('4. 요약');
    s4.getColumn(1).width = 26; [2, 3, 4].forEach(c => { s4.getColumn(c).width = 18; }); s4.getColumn(5).width = 40;
    let r = 1;
    const put = (vals, opt = {}) => { vals.forEach((v, i) => { const c = s4.getCell(r, 1 + i); c.value = v; if (typeof v === 'number') c.numFmt = FMT; if (opt.head) head(c, opt.sub); else if (opt.border !== false && v !== undefined) c.border = border; if (opt.bold) c.font = { bold: true }; }); r++; };
    s4.getCell(r, 1).value = '상품권 자금대사 요약'; s4.getCell(r, 1).font = { bold: true, size: 14 }; r += 2;
    put(['대사 기간(통장일)', `${res.start} ~ ${res.end}`], {});
    put(['통장 파일', meta.bankName || ''], {});
    put(['상품권 파일', meta.voucherName || ''], {});
    put(['생성 시각', meta.createdAt || ''], {});
    put(['불일치 영업일', res.mismatches.filter(m => !m.open).length ? `${res.mismatches.filter(m => !m.open).length}일` : '없음'], {});
    const offsetSum = res.offsets.reduce((s, o) => s + o.offset, 0);
    put(['QR 상계(일치 처리)', res.offsets.length ? `${res.offsets.length}일 · ${offsetSum.toLocaleString('ko-KR')}원` : '없음'], {});
    r++;
    put(['항목', '통장', '상품권', '차액', '기준'], { head: true });
    ITEMS.forEach(it => {
      const t = res.totals[it.key];
      // QR 차액이 상계 합계와 같으면 설명되는 차액이므로 노란색으로 구분
      const explained = (it.key === 'qr_col' || it.key === 'qr_set') && t.diff === offsetSum;
      put([it.name, t.b, t.v, t.diff, `${it.lag ? 'D-1' : 'D'} · ${it.src}${explained && t.diff ? ' (차액은 QR 상계분)' : ''}`]);
      const c = s4.getCell(r - 1, 4); if (t.diff) c.fill = explained ? OFFSET : BAD;
    });
    r++;
    put(['통장 단독 항목 (대사 제외)', '합계', '입금·출금일', '', ''], { head: true });
    res.bankOnly.forEach(o => put([o.name, o.total, o.days.length ? o.days.map(x => `${short(x.date)} ${x.amt.toLocaleString('ko-KR')}`).join(', ') : '-']));
    r++;
    put(['불일치 내역', '항목', '차액', '합산기간', ''], { head: true });
    if (!res.mismatches.length) put(['없음']);
    res.mismatches.forEach(m => ITEMS.forEach(it => { if (m.diff[it.key]) put([`${m.date} (${weekday(m.date)})${m.open ? ' · 참고' : ''}`, it.name, m.diff[it.key], m.label]); }));
    r++;
    put(['QR 상계 내역 (일치 처리)', 'QR추심·QR정산 차액', '', '합산기간', ''], { head: true });
    if (!res.offsets.length) put(['없음']);
    res.offsets.forEach(o => put([`${o.date} (${weekday(o.date)})${o.open ? ' · 참고' : ''}`, o.offset, undefined, o.label]));
    r++;
    put(['휴일로 처리한 날', res.usedHolidays.length ? res.usedHolidays.map(d => `${short(d)}(${weekday(d)})`).join(', ') : '없음'], {});
    if (res.warnings.length) { r++; put(['확인 필요'], { head: true }); res.warnings.forEach(w => put([w])); }

    wb.worksheets.forEach(ws => { ws.properties.defaultRowHeight = 18; });
    // 요약 시트를 맨 앞에서 열리게
    wb.views = [{ activeTab: 3 }];
    return wb;
  }

  // ---------- 결과 파일명: 상품권 대조_{지역}_{종료일 YYMMDD}.xlsx ----------
  const PROV = { '서울특별시': '서울', '부산광역시': '부산', '대구광역시': '대구', '인천광역시': '인천', '광주광역시': '광주', '대전광역시': '대전', '울산광역시': '울산', '세종특별자치시': '세종', '경기도': '경기', '강원특별자치도': '강원', '강원도': '강원', '충청북도': '충북', '충청남도': '충남', '전북특별자치도': '전북', '전라북도': '전북', '전라남도': '전남', '경상북도': '경북', '경상남도': '경남', '제주특별자치도': '제주' };
  function region(title) {
    const parts = String(title || '').split(/\s+/);
    if (!PROV[parts[0]]) return '';
    // 광역시 제목은 바로 '일별 …'이 오므로 시·군·구로 끝나는 단어만 지역으로 본다
    const city = /(시|군|구)$/.test(parts[1] || '') ? parts[1].replace(/(시|군|구)$/, '') : '';
    return [PROV[parts[0]], city].filter(Boolean).join(' ');
  }
  function fileName(res) {
    const rg = region(res.voucher.title);
    return `상품권 대조_${rg ? rg + '_' : ''}${res.end.slice(2).replace(/-/g, '')}.xlsx`;
  }

  // ---------- 휴일 (주말 외, 은행 휴무 기준. 매년 말에 다음 해 추가) ----------
  const HOLIDAYS = {
    '2025-01-01':'신정','2025-01-27':'임시공휴일','2025-01-28':'설날','2025-01-29':'설날','2025-01-30':'설날','2025-03-03':'대체(삼일절)',
    '2025-05-01':'근로자의 날','2025-05-05':'어린이날·부처님오신날','2025-05-06':'대체공휴일','2025-06-03':'대통령선거','2025-06-06':'현충일',
    '2025-08-15':'광복절','2025-10-03':'개천절','2025-10-06':'추석','2025-10-07':'추석','2025-10-08':'대체(추석)','2025-10-09':'한글날','2025-12-25':'성탄절',
    '2026-01-01':'신정','2026-02-16':'설날','2026-02-17':'설날','2026-02-18':'설날','2026-03-02':'대체(삼일절)','2026-05-01':'노동절',
    '2026-05-05':'어린이날','2026-05-25':'대체(부처님오신날)','2026-06-03':'지방선거','2026-07-17':'제헌절','2026-08-17':'대체(광복절)',
    '2026-09-24':'추석','2026-09-25':'추석','2026-10-05':'대체(개천절)','2026-10-09':'한글날','2026-12-25':'성탄절',
    '2027-01-01':'신정','2027-02-08':'설날','2027-02-09':'대체(설날)','2027-03-01':'삼일절','2027-05-03':'대체(노동절)','2027-05-05':'어린이날',
    '2027-05-13':'부처님오신날','2027-07-19':'대체(제헌절)','2027-08-16':'대체(광복절)','2027-09-14':'추석','2027-09-15':'추석','2027-09-16':'추석',
    '2027-10-04':'대체(개천절)','2027-10-11':'대체(한글날)','2027-12-27':'대체(성탄절)'
  };

  const api = { parseBank, parseVoucher, reconcile, buildWorkbook, fileName, region, ITEMS, HOLIDAYS, parseDate, addDays, weekday, short };
  if (typeof module !== 'undefined' && module.exports) module.exports = api; else root.Recon = api;
})(typeof window !== 'undefined' ? window : globalThis);
