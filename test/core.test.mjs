import { test } from 'node:test';
import assert from 'node:assert/strict';
import { HyperFormula } from 'hyperformula';
import { shiftFormula, parseRange, colName, colIndex, rangeText } from '../lib/a1.mjs';
import { parseInput, display, textRaw, editText } from '../lib/input.mjs';
import { Y, initBook, setRaw, needTab, addTab, tabList } from '../lib/book.mjs';
import { Engine, HF_CONFIG } from '../lib/engine.mjs';

test('addresses', () => {
  assert.equal(colName(0), 'A'); assert.equal(colName(25), 'Z'); assert.equal(colName(26), 'AA'); assert.equal(colName(701), 'ZZ');
  assert.equal(colIndex('AA'), 26);
  assert.deepEqual(parseRange('b3'), { tab: null, r1: 2, c1: 1, r2: 2, c2: 1 });
  assert.deepEqual(parseRange("'Q3 pipeline'!$A$1:C10"), { tab: 'Q3 pipeline', r1: 0, c1: 0, r2: 9, c2: 2 });
  assert.equal(parseRange('A:C').r2, Infinity);
  assert.equal(rangeText({ r1: 0, c1: 0, r2: 9, c2: 2 }, 'Q3 pipeline'), "'Q3 pipeline'!A1:C10");
  assert.throws(() => parseRange('hello'));
});

test('shifting a formula matches HyperFormula copy and paste', () => {
  const cases = ['=A1+B2', '=$A$1+A$1+$A1', '=SUM(A1:B3)*2', '=Data!B2', "='My data'!C3+1", '=LOG10(A2)', '="A1"&B1', '=SUM(A:A)', '=SUM(2:3)', '=IF(A1>0,"yes "&B1,C1)', '=a1*2', '=VLOOKUP(A2,Data!$A$1:$C$9,2,FALSE)', '=2.5*A1'];
  for (const f of cases) {
    for (const [dr, dc] of [[3, 0], [0, 2], [5, 1]]) {
      const hf = HyperFormula.buildFromSheets({ S: [], Data: [], 'My data': [] }, HF_CONFIG);
      hf.setCellContents({ sheet: 0, row: 2, col: 2 }, [[f]]);
      hf.copy({ start: { sheet: 0, row: 2, col: 2 }, end: { sheet: 0, row: 2, col: 2 } });
      hf.paste({ sheet: 0, row: 2 + dr, col: 2 + dc });
      const want = hf.getCellSerialized({ sheet: 0, row: 2 + dr, col: 2 + dc });
      assert.equal(shiftFormula(f, dr, dc).toUpperCase().replace(/\s/g, ''), want.toUpperCase().replace(/\s/g, ''), `${f} by ${dr},${dc}`);
      hf.destroy();
    }
  }
  assert.equal(shiftFormula('=A1', -1, 0), '=#REF!');
  // HyperFormula's own paste leaves A1 alone after a number in E notation; Excel and Google Sheets move it.
  assert.equal(shiftFormula('=1E5+A1', 3, 0), '=1E5+A4');
});

test('typing', () => {
  assert.deepEqual(parseInput('42'), { raw: 42 });
  assert.deepEqual(parseInput('$1,200'), { raw: 1200, nf: '$#,##0' });
  assert.deepEqual(parseInput('$1,200.50'), { raw: 1200.5, nf: '$#,##0.00' });
  assert.deepEqual(parseInput('12%'), { raw: 0.12, nf: '0%' });
  assert.deepEqual(parseInput('2026-10-07'), { raw: 46302, nf: 'yyyy-mm-dd' });
  assert.deepEqual(parseInput('10/7/2026'), { raw: 46302, nf: 'm/d/yyyy' });
  assert.deepEqual(parseInput('=SUM(A1:A2)'), { raw: '=SUM(A1:A2)' });
  assert.deepEqual(parseInput('hello'), { raw: 'hello' });
  assert.deepEqual(parseInput('true'), { raw: true });
  assert.deepEqual(parseInput('007', { nf: '@' }), { raw: "'007" });
  assert.equal(textRaw('007'), "'007");
  assert.equal(textRaw('Acme'), 'Acme');
  assert.equal(display(1234.5, '$#,##0.00'), '$1,234.50');
  assert.equal(display(46302, 'yyyy-mm-dd'), '2026-10-07');
  assert.equal(display(0.125, '0.0%'), '12.5%');
  assert.equal(display(1 / 3), '0.333333333');
  assert.equal(editText(46302, 'yyyy-mm-dd'), '2026-10-07');
  assert.equal(editText(0.12, '0%'), '12%');
});

test('the engine follows the document, including someone else\'s changes', () => {
  const doc = new Y.Doc();
  initBook(doc, { tabs: [{ name: 'Deals' }, { name: 'Summary' }] });
  const deals = needTab(doc, 'Deals').id, sum = needTab(doc, 'Summary').id;
  const e = new Engine(doc);
  doc.transact(() => { setRaw(doc, deals, 0, 0, 100); setRaw(doc, deals, 1, 0, 250); setRaw(doc, sum, 0, 0, '=SUM(Deals!A1:A10)'); setRaw(doc, sum, 0, 1, '=AVERAGEIFS(Deals!A1:A2,Deals!A1:A2,">50")'); setRaw(doc, sum, 0, 2, '=CONCAT("a",Deals!A1)'); setRaw(doc, sum, 0, 3, '=RANK(Deals!A2,Deals!A1:A2)'); setRaw(doc, sum, 0, 4, '2026-10-07'); });
  assert.equal(e.value(sum, 0, 0), 350);
  assert.equal(e.value(sum, 0, 1), 175);
  assert.equal(e.value(sum, 0, 2), 'a100');
  assert.equal(e.value(sum, 0, 3), 1);
  assert.equal(e.value(sum, 0, 4), '2026-10-07', 'text that looks like a date stays text');
  // A second copy (another person) gets the change as a Yjs update.
  const doc2 = new Y.Doc();
  Y.applyUpdate(doc2, Y.encodeStateAsUpdate(doc));
  const e2 = new Engine(doc2);
  doc2.on('update', (u) => Y.applyUpdate(doc, u));
  doc2.transact(() => setRaw(doc2, deals, 2, 0, 50));
  assert.equal(e.value(sum, 0, 0), 400);
  // Rename a tab: formulas follow (the server writes the renamed formulas into the document).
  const before = tabList(doc).length;
  addTab(doc, { name: 'Extra' });
  assert.equal(tabList(doc).length, before + 1);
  assert.equal(e.value(sum, 0, 0), 400);
  // Insert two rows at the top of Deals: references move.
  e.structural((hf) => hf.addRows(e.sheet(deals), [0, 2]));
  assert.equal(doc.getMap('cells').get(sum).get('0,0'), '=SUM(Deals!A3:A12)');
  assert.equal(doc.getMap('cells').get(deals).get('2,0'), 100);
  assert.equal(e2.value(sum, 0, 0), 400);
  e.destroy(); e2.destroy();
});
