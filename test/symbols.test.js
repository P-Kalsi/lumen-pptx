const test = require('node:test');
const assert = require('node:assert/strict');
const { fixSymbolText } = require('../src/symbols.js');
const run = (rPr, text) => `<a:r><a:rPr lang="en-US">${rPr}</a:rPr><a:t>${text}</a:t></a:r>`;

test('maps private-use Symbol code points in runs with a Symbol sym font', () => {
  const xml = run('<a:latin typeface="+mn-lt"/><a:sym typeface="Symbol" charset="2"/>', 'a b  c ');
  assert.equal(fixSymbolText(xml), run('<a:latin typeface="+mn-lt"/><a:sym typeface="Symbol" charset="2"/>', 'a← b ∧ c γ'));
});
test('maps ASCII when the run font itself is Symbol, keeping entities', () => {
  const rPr = '<a:latin typeface="Symbol"/>';
  assert.equal(fixSymbolText(run(rPr, 'ab&amp;')), run(rPr, 'αβ&amp;'));
});
test('leaves non-Symbol runs untouched', () => {
  const xml = run('<a:latin typeface="Arial"/>', 'g ') + run('<a:sym typeface="Wingdings"/>', 'Symbol');
  assert.equal(fixSymbolText(xml), xml);
});
