// Legacy decks type ←, ∧, γ, etc. in the Windows "Symbol" font. PowerPoint stores those as either
// private-use code points U+F020–U+F0FF (with <a:sym typeface="Symbol"/>) or plain ASCII in a run
// whose <a:latin> font is Symbol ("g" shows as γ). Browsers have no such mapping and draw tofu (□),
// so rewrite those characters to their real Unicode equivalents before rendering.

// Symbol font byte 0x20–0xFF → Unicode (Adobe Symbol encoding; 0x7F–0x9F and 0xF0 unused).
const TABLE = {
  0x22: '∀', 0x24: '∃', 0x27: '∋', 0x2A: '∗', 0x2D: '−', 0x40: '≅',
  0x41: 'Α', 0x42: 'Β', 0x43: 'Χ', 0x44: 'Δ', 0x45: 'Ε', 0x46: 'Φ', 0x47: 'Γ', 0x48: 'Η', 0x49: 'Ι',
  0x4A: 'ϑ', 0x4B: 'Κ', 0x4C: 'Λ', 0x4D: 'Μ', 0x4E: 'Ν', 0x4F: 'Ο', 0x50: 'Π', 0x51: 'Θ', 0x52: 'Ρ',
  0x53: 'Σ', 0x54: 'Τ', 0x55: 'Υ', 0x56: 'ς', 0x57: 'Ω', 0x58: 'Ξ', 0x59: 'Ψ', 0x5A: 'Ζ',
  0x5C: '∴', 0x5E: '⊥', 0x60: '‾',
  0x61: 'α', 0x62: 'β', 0x63: 'χ', 0x64: 'δ', 0x65: 'ε', 0x66: 'φ', 0x67: 'γ', 0x68: 'η', 0x69: 'ι',
  0x6A: 'ϕ', 0x6B: 'κ', 0x6C: 'λ', 0x6D: 'μ', 0x6E: 'ν', 0x6F: 'ο', 0x70: 'π', 0x71: 'θ', 0x72: 'ρ',
  0x73: 'σ', 0x74: 'τ', 0x75: 'υ', 0x76: 'ϖ', 0x77: 'ω', 0x78: 'ξ', 0x79: 'ψ', 0x7A: 'ζ', 0x7E: '∼',
  0xA0: '€', 0xA1: 'ϒ', 0xA2: '′', 0xA3: '≤', 0xA4: '⁄', 0xA5: '∞', 0xA6: 'ƒ', 0xA7: '♣', 0xA8: '♦',
  0xA9: '♥', 0xAA: '♠', 0xAB: '↔', 0xAC: '←', 0xAD: '↑', 0xAE: '→', 0xAF: '↓',
  0xB0: '°', 0xB1: '±', 0xB2: '″', 0xB3: '≥', 0xB4: '×', 0xB5: '∝', 0xB6: '∂', 0xB7: '•', 0xB8: '÷',
  0xB9: '≠', 0xBA: '≡', 0xBB: '≈', 0xBC: '…', 0xBD: '⏐', 0xBE: '⎯', 0xBF: '↵',
  0xC0: 'ℵ', 0xC1: 'ℑ', 0xC2: 'ℜ', 0xC3: '℘', 0xC4: '⊗', 0xC5: '⊕', 0xC6: '∅', 0xC7: '∩', 0xC8: '∪',
  0xC9: '⊃', 0xCA: '⊇', 0xCB: '⊄', 0xCC: '⊂', 0xCD: '⊆', 0xCE: '∈', 0xCF: '∉',
  0xD0: '∠', 0xD1: '∇', 0xD2: '®', 0xD3: '©', 0xD4: '™', 0xD5: '∏', 0xD6: '√', 0xD7: '⋅', 0xD8: '¬',
  0xD9: '∧', 0xDA: '∨', 0xDB: '⇔', 0xDC: '⇐', 0xDD: '⇑', 0xDE: '⇒', 0xDF: '⇓',
  0xE0: '◊', 0xE1: '〈', 0xE2: '®', 0xE3: '©', 0xE4: '™', 0xE5: '∑', 0xE6: '⎛', 0xE7: '⎜', 0xE8: '⎝',
  0xE9: '⎡', 0xEA: '⎢', 0xEB: '⎣', 0xEC: '⎧', 0xED: '⎨', 0xEE: '⎩', 0xEF: '⎪',
  0xF1: '〉', 0xF2: '∫', 0xF3: '⌠', 0xF4: '⎮', 0xF5: '⌡', 0xF6: '⎞', 0xF7: '⎟', 0xF8: '⎠',
  0xF9: '⎤', 0xFA: '⎥', 0xFB: '⎦', 0xFC: '⎫', 0xFD: '⎬', 0xFE: '⎭'
};
function mapCode(code) {
  if (code < 0x20 || code > 0xFF) return null;
  return TABLE[code] || (code < 0x7F ? String.fromCharCode(code) : null);
}
const isSymbol = (rPr, tag) => new RegExp(`<a:${tag}\\b[^>]*typeface="Symbol"`).test(rPr);

// Rewrite one run's text. `wholeRun`: the run's own font is Symbol, so ASCII maps too (g → γ);
// otherwise only private-use U+F0xx characters (which PowerPoint draws with the sym font) map.
function mapText(text, wholeRun) {
  // Keep XML entities (&amp;, &#x…;) intact; they are never Symbol characters.
  return text.replace(/&[^;]+;|[\s\S]/g, ch => {
    if (ch.length > 1) return ch;
    const code = ch.charCodeAt(0);
    if (code >= 0xF020 && code <= 0xF0FF) return mapCode(code - 0xF000) ?? ch;
    return wholeRun ? mapCode(code) ?? ch : ch;
  });
}

// Returns the XML with Symbol-font runs converted to Unicode (unchanged if none).
function fixSymbolText(xml) {
  if (!xml.includes('Symbol')) return xml;
  return xml.replace(/<a:(r|fld)\b([^>]*)>([\s\S]*?)<\/a:\1>/g, (run, tag, attrs, body) => {
    const rPr = (body.match(/<a:rPr\b[^>]*\/>|<a:rPr\b[^>]*>[\s\S]*?<\/a:rPr>/) || [''])[0];
    const whole = isSymbol(rPr, 'latin'), sym = whole || isSymbol(rPr, 'sym');
    if (!sym) return run;
    return `<a:${tag}${attrs}>${body.replace(/(<a:t>)([^<]*)(<\/a:t>)/g, (_, open, text, close) => open + mapText(text, whole) + close)}</a:${tag}>`;
  });
}

const PARTS = /^ppt\/(slides|slideLayouts|slideMasters|notesSlides)\/[^/]+\.xml$/;
// Applies fixSymbolText to every slide-like part of a JSZip archive. Returns true if any changed.
async function fixSymbolFonts(zip) {
  let changed = false;
  await Promise.all(Object.keys(zip.files).filter(name => PARTS.test(name)).map(async name => {
    const xml = await zip.file(name).async('string');
    const fixed = fixSymbolText(xml);
    if (fixed !== xml) { zip.file(name, fixed); changed = true; }
  }));
  return changed;
}
module.exports = { fixSymbolText, fixSymbolFonts };
