const test = require('node:test');
const assert = require('node:assert');
const G = require('../genkou.js');

const show = (rows) => rows.map((r) => r.map((c) => c.text + c.hang).join('|'));

test('全角は1マス1文字、20マスで改行', () => {
  const rows = G.toRows(G.toCells('あ'.repeat(25)));
  assert.equal(rows.length, 2);
  assert.equal(rows[0].length, 20);
  assert.equal(rows[1].length, 5);
});

test('半角は1マスに2文字', () => {
  const cells = G.toCells('ROE10%向上');
  assert.deepEqual(cells.map((c) => c.text), ['RO', 'E1', '0%', '向', '上']);
});

test('半角が奇数個なら最後は1文字マス', () => {
  assert.deepEqual(G.toCells('IT化').map((c) => c.text), ['IT', '化']);
  assert.deepEqual(G.toCells('A社').map((c) => c.text), ['A', '社']);
});

test('行頭に来る句読点は前行の最終マスへ', () => {
  const rows = G.toRows(G.toCells('あ'.repeat(20) + '。い'));
  assert.equal(rows.length, 2);
  assert.equal(rows[0][19].text, 'あ');
  assert.equal(rows[0][19].hang, '。');
  assert.deepEqual(show(rows)[1], 'い');
});

test('ちょうど20マスで句読点が終わるならぶら下げない', () => {
  const rows = G.toRows(G.toCells('あ'.repeat(19) + '。'));
  assert.equal(rows.length, 1);
  assert.equal(rows[0][19].text, '。');
});

test('閉じ括弧はオプションで前行へ', () => {
  const s = 'あ'.repeat(20) + '」。';
  assert.equal(G.toRows(G.toCells(s))[1][0].text, '」');
  const rows = G.toRows(G.toCells(s), { brackets: true });
  assert.equal(rows[0][19].hang, '」。');
  assert.equal(rows.length, 1);
});

test('見出しと字数の解析', () => {
  const blocks = G.layout('# 第1問 １００字\nあいう\nえお\n\n#第2問(60字)\nか');
  assert.equal(blocks.length, 2);
  assert.equal(blocks[0].limit, 100);
  assert.equal(blocks[0].rows.length, 5);
  assert.equal(blocks[0].count, 5);
  assert.equal(blocks[1].limit, 60);
  assert.equal(blocks[1].rows.length, 3);
});

test('字数超過', () => {
  const [b] = G.layout('# 20字\n' + 'あ'.repeat(21));
  assert.equal(b.over, true);
  assert.equal(b.rows.length, 2);
  const [c] = G.layout('# 20字\n' + 'あ'.repeat(20) + '。');
  assert.equal(c.over, false); // 句読点ぶら下げで20マスに収まる
});

test('docx は zip として生成される', () => {
  const bytes = G.buildDocx(G.layout('# 第1問 40字\nA社の強みは、高い技術力である。'), { title: 't' });
  assert.equal(bytes[0], 0x50);
  assert.equal(bytes[1], 0x4b);
});

test('pdf はヘッダと xref を持ち、オフセットが正しい', () => {
  let txt = '';
  for (let i = 1; i <= 8; i++) txt += `# 第${i}問 120字\n` + 'テスト。'.repeat(20) + '\n';
  const s = new TextDecoder().decode(G.buildPdf(G.layout(txt), { title: 't' }));
  assert.ok(s.startsWith('%PDF-1.4'));
  const xref = parseInt(s.match(/startxref\n(\d+)/)[1], 10);
  assert.ok(s.slice(xref).startsWith('xref'));
  const offs = [...s.slice(xref).matchAll(/^(\d{10}) 00000 n $/gm)].map((m) => parseInt(m[1], 10));
  offs.forEach((o, i) => assert.ok(s.slice(o).startsWith(`${i + 1} 0 obj`)));
  assert.ok(/\/Count [2-9]/.test(s)); // 複数ページ
});
