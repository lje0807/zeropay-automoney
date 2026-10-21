// 엔진 회귀 테스트 (CLAUDE.md '2026년 9월 데이터 검증 결과' 기준)
// 실행: npm test  ·  샘플 위치: ../samples/{bank,voucher}.xlsx (RECON_SAMPLES 환경변수로 변경 가능)
'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const ExcelJS = require('exceljs');
const R = require('../js/engine.js');

const SAMPLES = process.env.RECON_SAMPLES || path.join(__dirname, '..', '..', 'samples');
const BANK = path.join(SAMPLES, 'bank.xlsx'), VOUCHER = path.join(SAMPLES, 'voucher.xlsx');
const haveSamples = fs.existsSync(BANK) && fs.existsSync(VOUCHER);
const skip = haveSamples ? false : `샘플 파일 없음: ${SAMPLES}`;

const load = async p => { const wb = new ExcelJS.Workbook(); await wb.xlsx.readFile(p); return wb; };
let cached;
async function sept() {
  if (!cached) {
    const bank = R.parseBank(await load(BANK));
    const voucher = R.parseVoucher(await load(VOUCHER));
    const res = R.reconcile(bank, voucher, { start: '2026-09-01', end: '2026-09-30', holidays: Object.keys(R.HOLIDAYS) });
    cached = { bank, voucher, res };
  }
  return cached;
}

// ---------- 샘플 없이 도는 단위 테스트 ----------
test('parseDate: 문자열·Date·엑셀 일련번호', () => {
  assert.equal(R.parseDate('2026.09.01'), '2026-09-01');
  assert.equal(R.parseDate('2026-9-1'), '2026-09-01');
  assert.equal(R.parseDate(new Date(Date.UTC(2026, 8, 30))), '2026-09-30');
  assert.equal(R.parseDate(46295), '2026-09-30');
  assert.equal(R.parseDate('전월 잔액'), null);
});

test('region / fileName: 제목에서 지역명 추출', () => {
  assert.equal(R.region('경상남도 함양군 일별 상품권 잔액집계'), '경남 함양');
  assert.equal(R.region('서울특별시 일별 상품권 잔액집계'), '서울');
  assert.equal(R.region('전북특별자치도 익산시 일별 상품권 잔액집계'), '전북 익산');
  assert.equal(R.region('알 수 없는 제목'), '');
  assert.equal(R.fileName({ end: '2026-09-30', voucher: { title: '경상남도 함양군 일별 상품권 잔액집계' } }), '상품권 대조_경남 함양_260930.xlsx');
  assert.equal(R.fileName({ end: '2026-09-30', voucher: { title: '' } }), '상품권 대조_260930.xlsx');
});

test('reconcile: 휴일은 다음 영업일에 합산, 종료일이 휴일이면 참고', () => {
  const z = { p_buy: 0, c_buy: 0, p_ref: 0, c_ref: 0, qr_col: 0, qr_set: 0, pg: 0, disc: 0, point: 0, auto: 0, trust_in: 0, trust_out: 0 };
  // 2026-09-04(금) ~ 09-07(월) ~ 09-12(토). 통장은 토·월에 나눠 들어오고, 상품권(D-1)은 금·토·일 발생
  const bank = { missing: [], data: {
    '2026-09-04': { ...z }, '2026-09-05': { ...z, p_buy: 30 }, '2026-09-06': { ...z }, '2026-09-07': { ...z, p_buy: 70 },
    '2026-09-12': { ...z, p_buy: 5 } } };
  const voucher = { missing: [], data: {
    '2026-09-04': { ...z, p_buy: 40 }, '2026-09-05': { ...z, p_buy: 35 }, '2026-09-06': { ...z, p_buy: 25 },
    '2026-09-11': { ...z, p_buy: 5 } } };
  const res = R.reconcile(bank, voucher, { start: '2026-09-04', end: '2026-09-12', holidays: [] });
  const mon = res.rows.find(r => r.date === '2026-09-07');
  assert.equal(mon.label, '09.05~09.07');
  assert.equal(mon.diff.p_buy, 0);
  assert.ok(mon.ok);
  assert.ok(!res.rows.find(r => r.date === '2026-09-05').closes, '휴일 행은 따로 비교하지 않음');
  const sat = res.rows.find(r => r.date === '2026-09-12');
  assert.ok(sat.open && sat.closes);
  assert.ok(res.warnings.some(w => w.includes('종료일(09.12)이 휴일')));
  assert.ok(res.warnings.some(w => w.includes('상품권 파일에 없는 거래일: 09.03')), 'D-1 데이터가 빠진 날은 경고');
});

// ---------- 2026년 9월 샘플 회귀 테스트 ----------
test('9월 샘플: 열을 모두 찾는다', { skip }, async () => {
  const { bank, voucher } = await sept();
  assert.deepEqual(bank.missing, []);
  assert.deepEqual(voucher.missing, []);
  assert.equal(Object.keys(bank.data).filter(d => d.startsWith('2026-09')).length, 30);
  assert.match(voucher.title, /^경상남도 함양군/);
});

test('9월 샘플: 불일치 영업일은 9/14, 9/16, 9/23, 9/28 (QR추심·QR정산 상계)', { skip }, async () => {
  const { res } = await sept();
  const expected = { '2026-09-14': -250000, '2026-09-16': -15000, '2026-09-23': -2000, '2026-09-28': -7000 };
  assert.deepEqual(res.mismatches.map(m => m.date), Object.keys(expected));
  for (const m of res.mismatches) {
    assert.equal(m.open, undefined);
    for (const it of R.ITEMS) {
      const want = it.key === 'qr_col' || it.key === 'qr_set' ? expected[m.date] : 0;
      assert.equal(m.diff[it.key], want, `${m.date} ${it.name}`);
    }
  }
  assert.equal(res.rows.find(r => r.date === '2026-09-28').label, '09.24~09.28', '추석·주말 묶음');
  assert.deepEqual(res.warnings, []);
});

test('9월 샘플: 항목별 합계', { skip }, async () => {
  const { res } = await sept();
  for (const k of ['p_buy', 'c_buy', 'p_ref', 'c_ref', 'pg']) assert.equal(res.totals[k].diff, 0, k);
  assert.deepEqual(res.totals.qr_col, { b: 1296000, v: 1570000, diff: -274000 });
  assert.deepEqual(res.totals.qr_set, { b: 2535813961, v: 2536087961, diff: -274000 });
  const disc = res.bankOnly.find(o => o.key === 'disc');
  assert.deepEqual(disc.days, [{ date: '2026-09-17', amt: 227977000 }]);
  assert.equal(R.fileName(res), '상품권 대조_경남 함양_260930.xlsx');
});

test('9월 샘플: 대조 시트 수식이 통장·상품권 시트의 올바른 열을 참조한다', { skip }, async () => {
  const { res } = await sept();
  const built = await R.buildWorkbook(ExcelJS, res, { bankName: 'bank.xlsx', voucherName: 'voucher.xlsx', createdAt: 'test' });
  // 실제 파일로 저장했다가 다시 읽어서 확인
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(await built.xlsx.writeBuffer());
  assert.deepEqual(wb.worksheets.map(ws => ws.name), ['1. 통장', '2. 상품권', '3. 대조', '4. 요약']);
  const s1 = wb.getWorksheet('1. 통장'), s2 = wb.getWorksheet('2. 상품권'), s3 = wb.getWorksheet('3. 대조');
  const BANK_LABEL = { p_buy: '개인 구매금', c_buy: '법인 구매금', p_ref: '개인 환불금', c_ref: '법인 환불금', qr_col: 'QR추심/기타', qr_set: 'QR정산/기타', pg: 'PG정산' };
  const val = v => (v && typeof v === 'object' && 'result' in v) ? v.result : (v || 0);
  const sumRange = (ws, col, r1, r2) => { let s = 0; for (let r = r1; r <= r2; r++) s += val(ws.getCell(`${col}${r}`).value); return s; };
  const ref = /^(?:SUM\()?'1\. 통장'!([A-Z]+)(\d+)(?::([A-Z]+)(\d+)\))?-(?:SUM\()?'2\. 상품권'!([A-Z]+)(\d+)(?::([A-Z]+)(\d+)\))?$/;

  let checked = 0;
  R.ITEMS.forEach((it, i) => {
    const vCol = String.fromCharCode(67 + i); // C부터 ITEMS 순서
    assert.equal(s3.getCell(3, 4 + i).value, it.name);
    assert.equal(s2.getCell(`${vCol}2`).value, it.name);
    for (let r = 4; r < 4 + res.dates.length; r++) {
      const cell = s3.getCell(r, 4 + i).value;
      if (!cell || !cell.formula) continue;
      const m = cell.formula.match(ref);
      assert.ok(m, `수식 형식: ${cell.formula}`);
      const [, bc, br1, , br2, vc, vr1, , vr2] = m;
      assert.equal(s1.getCell(`${bc}3`).value, BANK_LABEL[it.key], `${it.name} 통장 열 (${cell.formula})`);
      assert.equal(vc, vCol, `${it.name} 상품권 열 (${cell.formula})`);
      assert.equal(br1, vr1); assert.equal(br2, vr2);
      const calc = sumRange(s1, bc, +br1, +(br2 || br1)) - sumRange(s2, vc, +vr1, +(vr2 || vr1));
      // 파일에는 <v>0</v>으로 저장되지만 ExcelJS가 다시 읽을 때 0 결과를 버리므로 undefined는 0으로 본다
      assert.equal(calc, cell.result ?? 0, `${r}행 ${it.name}: 수식 계산값과 저장값`);
      checked++;
    }
  });
  assert.ok(checked > 0);
  // 9/14(월)은 9/12~9/14 묶음 → 15~17행
  assert.equal(s3.getCell('H17').value.formula, "SUM('1. 통장'!H15:H17)-SUM('2. 상품권'!G15:G17)"); // QR추심
  assert.equal(s3.getCell('I17').value.formula, "SUM('1. 통장'!N15:N17)-SUM('2. 상품권'!H15:H17)"); // QR정산
  assert.equal(s3.getCell('I17').value.result, -250000);
  assert.equal(s3.getCell('C15').value, '→ 09.14 합산');
});
