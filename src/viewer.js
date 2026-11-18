import { PptxViewer, RECOMMENDED_ZIP_LIMITS } from '@aiden0z/pptx-renderer';
import JSZip from 'jszip';
import { renderToSvg } from 'wmf-emf-renderer';
import { fixSymbolFonts } from './symbols.js';
const api = acquireVsCodeApi();
const $ = id => document.getElementById(id);
let viewer, thumbnails = [], thumbObserver, notes = [], matches = [], matchIndex = -1, wmfBySlide = [];
let loading = Promise.resolve(), revision = 0, currentPage = 0, resizeTimer;
let operations = Promise.resolve();
const state = api.getState() || {};
const savedZoom = state.zoom || 'fit';
// A pinch can save an arbitrary percentage the dropdown has no option for; add it so it restores.
if (!isNaN(Number(savedZoom)) && ![...$('zoom').options].some(o => o.value === String(savedZoom))) {
  $('zoom').add(new Option(`${savedZoom}%`, String(savedZoom)));
}
$('zoom').value = savedZoom;
function status(text, error = false) {
  $('status').textContent = text; $('status').hidden = !text;
  $('status').classList.toggle('error', error);
}
function run(fn) {
  operations = operations.catch(() => {}).then(fn).catch(error => status(error.message, true));
  return operations;
}
function save() { api.setState({ page: currentPage, zoom: $('zoom').value, sidebar: !$('sidebar').hidden }); }
function pageChanged(index) {
  currentPage = index; $('page').value = index + 1;
  $('previous').disabled = index === 0;
  $('next').disabled = !viewer || index >= viewer.slideCount - 1;
  $('sidebar').querySelectorAll('.thumbnail').forEach((el, i) => el.classList.toggle('active', i === index));
  $('notes-text').textContent = notes[index] || 'No speaker notes for this slide.';
  save();
}
async function go(index) {
  if (!viewer) return;
  index = Math.max(0, Math.min(viewer.slideCount - 1, index));
  await viewer.goToSlide(index, { behavior: 'instant', block: 'start' });
  pageChanged(index);
}
// The renderer stays at a fixed 1:1 render (fitMode 'none', zoom 100) so a slide element is
// exactly slideWidth px. ALL user zoom — fit width, fit slide, percentages, pinch — is a
// persistent CSS `zoom` on the (vector) slides. Because the renderer is never in 'contain' mode,
// its handleContainerResize() early-returns: it never re-fits or re-renders when a scrollbar
// toggles at fit width, so there is no white flash and the base scale (100%) is constant.
// effective% == cssZoom * 100.
let cssZoom = 1, zoomSaveTimer, customOption;
function effectiveZoom() { return cssZoom * 100; }
// The percentage that fits a slide to the viewport width ('fit') or fully ('page').
function fitPercent(mode) {
  const width = ($('viewport').clientWidth - 40) / viewer.slideWidth * 100;
  return mode === 'page' ? Math.min(width, ($('viewport').clientHeight - 40) / viewer.slideHeight * 100) : width;
}
// A fit/page/percentage choice: apply the CSS zoom but keep the dropdown showing the mode so it
// re-fits on resize. Pinch/buttons/keys go through userZoom instead, which switches to a number.
function zoom() {
  if (!viewer) return;
  const selected = $('zoom').value;
  const percent = selected === 'fit' || selected === 'page' ? fitPercent(selected) : Number(selected);
  applyCssZoom(Math.max(25, Math.min(400, percent)) / 100, null);
  save();
}
function applyCssZoom(target, event) {
  const slides = $('slides'), viewport = $('viewport');
  const vp = viewport.getBoundingClientRect(), r = slides.getBoundingClientRect();
  const px = event ? event.clientX : vp.left + viewport.clientWidth / 2;
  const py = event ? event.clientY : vp.top + viewport.clientHeight / 2;
  // Anchor the point under the cursor (viewport centre for keyboard/buttons) so zooming does
  // not drift. L0/T0 are the slides' fixed position in the scroll content; cx/cy the unzoomed
  // content coordinates beneath the anchor.
  const L0 = viewport.scrollLeft + (r.left - vp.left), T0 = viewport.scrollTop + (r.top - vp.top);
  const cx = (px - r.left) / cssZoom, cy = (py - r.top) / cssZoom;
  cssZoom = target;
  slides.style.zoom = cssZoom === 1 ? '' : String(cssZoom);
  viewport.scrollLeft = Math.max(0, Math.min(L0 + cx * cssZoom - (px - vp.left), viewport.scrollWidth - viewport.clientWidth));
  viewport.scrollTop = Math.max(0, Math.min(T0 + cy * cssZoom - (py - vp.top), viewport.scrollHeight - viewport.clientHeight));
}
// Reflect the live percentage in the toolbar using one reusable option so the dropdown is not
// flooded with intermediate values.
function showZoom(percent) {
  percent = Math.round(percent);
  if ([...$('zoom').options].some(o => o !== customOption && o.value === String(percent))) { $('zoom').value = String(percent); return; }
  if (!customOption) { customOption = new Option('', ''); $('zoom').add(customOption); }
  customOption.text = `${percent}%`; customOption.value = String(percent); $('zoom').value = String(percent);
}
function scheduleZoomSave() { clearTimeout(zoomSaveTimer); zoomSaveTimer = setTimeout(save, 300); }
function userZoom(nextPercent, event) {
  if (!viewer) return;
  const clamped = Math.max(25, Math.min(400, nextPercent));
  applyCssZoom(clamped / 100, event);
  showZoom(clamped); scheduleZoomSave();
}
function pinchZoom(event) { userZoom(effectiveZoom() * Math.exp(-event.deltaY / 200), event); }
function setupThumbnails() {
  thumbObserver?.disconnect(); thumbnails.forEach(h => h.dispose()); thumbnails = [];
  $('sidebar').replaceChildren();
  thumbObserver = new IntersectionObserver(entries => {
    for (const entry of entries) {
      const index = Number(entry.target.dataset.index);
      if (entry.isIntersecting && !thumbnails[index]) {
        const handle = viewer.renderThumbnailToContainer(index, entry.target.querySelector('.thumb-slide'), { width: 160 });
        if (handle) { thumbnails[index] = handle; handle.ready.catch(error => status(error.message, true)); }
      } else if (!entry.isIntersecting && thumbnails[index]) {
        thumbnails[index].dispose(); delete thumbnails[index];
      }
    }
  }, { root: $('sidebar'), rootMargin: '200px' });
  for (let i = 0; i < viewer.slideCount; i++) {
    const button = document.createElement('button'); button.className = 'thumbnail'; button.dataset.index = i;
    button.setAttribute('aria-label', `Go to slide ${i + 1}`);
    const thumb = document.createElement('div'); thumb.className = 'thumb-slide';
    thumb.style.height = `${160 * viewer.slideHeight / viewer.slideWidth}px`;
    const label = document.createElement('span'); label.className = 'thumb-number'; label.textContent = i + 1;
    button.append(thumb, label); button.onclick = () => run(() => go(i)); $('sidebar').append(button); thumbObserver.observe(button);
  }
}
async function extractNotes(zip, presentation) {
  const parser = new DOMParser();
  return Promise.all(presentation.slides.map(async slide => {
    const relFile = slide.slidePath.replace(/\/([^/]+)$/, '/_rels/$1.rels');
    const relEntry = zip.file(relFile); if (!relEntry) return '';
    const rels = parser.parseFromString(await relEntry.async('string'), 'application/xml');
    const rel = [...rels.getElementsByTagName('Relationship')].find(r => r.getAttribute('Type')?.endsWith('/notesSlide') && r.getAttribute('TargetMode') !== 'External');
    if (!rel) return '';
    const name = new URL(rel.getAttribute('Target'), 'https://pptx.local/' + slide.slidePath).pathname.slice(1);
    const entry = zip.file(name); if (!entry) return '';
    const xml = await entry.async('string'); if (xml.length > 2 * 1024 * 1024) return 'Notes too large to display.';
    const doc = parser.parseFromString(xml, 'application/xml');
    const ns = 'http://schemas.openxmlformats.org/presentationml/2006/main';
    const a = 'http://schemas.openxmlformats.org/drawingml/2006/main';
    return [...doc.getElementsByTagNameNS(ns, 'sp')].filter(sp => sp.getElementsByTagNameNS(ns, 'ph')[0]?.getAttribute('type') === 'body')
      .flatMap(sp => [...sp.getElementsByTagNameNS(a, 'p')].map(p => [...p.getElementsByTagNameNS(a, 't')].map(t => t.textContent).join(''))).join('\n');
  }));
}
// The renderer decodes EMF but not WMF (it shows an "Unsupported format: WMF" placeholder).
// Collect each slide's WMF image bytes in document order so they can be drawn in later.
async function extractWmf(zip, presentation) {
  const parser = new DOMParser();
  return Promise.all(presentation.slides.map(async slide => {
    const slideEntry = zip.file(slide.slidePath); if (!slideEntry) return [];
    const relFile = slide.slidePath.replace(/\/([^/]+)$/, '/_rels/$1.rels');
    const relEntry = zip.file(relFile); if (!relEntry) return [];
    const rels = parser.parseFromString(await relEntry.async('string'), 'application/xml');
    const targets = {};
    for (const r of rels.getElementsByTagName('Relationship')) targets[r.getAttribute('Id')] = r.getAttribute('Target');
    const xml = await slideEntry.async('string');
    const out = [];
    for (const m of xml.matchAll(/r:embed="([^"]+)"/g)) {
      const target = targets[m[1]];
      if (!target || !/\.wmf$/i.test(target)) continue;
      const path = new URL(target, 'https://pptx.local/' + slide.slidePath).pathname.slice(1);
      const entry = zip.file(path);
      out.push(entry ? await entry.async('uint8array') : null);
    }
    return out;
  }));
}
// Replace each "Unsupported format: WMF" placeholder in a rendered slide with the decoded image.
function fillWmf(index, element) {
  const items = wmfBySlide[index]; if (!items || !items.length) return;
  const placeholders = [...element.querySelectorAll('div')]
    .filter(d => d.childElementCount === 0 && /^Unsupported format:\s*WMF/i.test((d.textContent || '').trim()));
  placeholders.forEach((textDiv, k) => {
    const bytes = items[k], box = textDiv.parentElement?.parentElement;
    if (!bytes || !box) return;
    try {
      const svg = renderToSvg(bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength));
      const wrap = document.createElement('div');
      wrap.style.cssText = 'width:100%;height:100%;';
      wrap.innerHTML = svg;
      const node = wrap.querySelector('svg');
      if (!node) return;
      node.removeAttribute('width'); node.removeAttribute('height');
      node.style.cssText = 'width:100%;height:100%;display:block;';
      box.replaceChildren(wrap);
    } catch { /* keep the placeholder if decoding fails */ }
  });
}
async function open(data, request) {
  await operations;
  if (request !== revision) return;
  status('Reading presentation…');
  thumbObserver?.disconnect(); thumbnails.forEach(h => h.dispose()); thumbnails = [];
  viewer?.destroy(); viewer = null; $('sidebar').replaceChildren(); notes = []; matches = []; wmfBySlide = [];
  let bytes = Uint8Array.from(atob(data), char => char.charCodeAt(0));
  const zip = await JSZip.loadAsync(bytes);
  // Symbol-font characters would render as tofu (□); rewrite them to Unicode first.
  if (await fixSymbolFonts(zip)) bytes = await zip.generateAsync({ type: 'uint8array', compression: 'STORE' });
  let failures = 0;
  viewer = new PptxViewer($('slides'), {
    scrollContainer: $('viewport'), fitMode: 'none', zoomPercent: 100, externalMedia: 'block', pdfjs: false,
    zipLimits: RECOMMENDED_ZIP_LIMITS,
    onSlideChange: pageChanged,
    onSlideRendered: fillWmf,
    onSlideError: () => { failures++; status('Some slide content could not be rendered.', true); },
    onNodeError: () => { failures++; status('Some slide content could not be rendered.', true); }
  });
  await viewer.open(bytes, { renderMode: 'list', listOptions: { windowed: true, initialSlides: 3, batchSize: 3 } });
  if (request !== revision) return;
  if (!viewer.slideCount) throw new Error('This file contains no slides.');
  notes = await extractNotes(zip, viewer.presentationData);
  wmfBySlide = await extractWmf(zip, viewer.presentationData);
  // Slides mounted before wmfBySlide was ready (the initial batch) still show placeholders — fill them now.
  $('slides').querySelectorAll('[data-slide-index]').forEach(el => fillWmf(Number(el.dataset.slideIndex), el));
  $('total').textContent = `/ ${viewer.slideCount}`; $('page').max = viewer.slideCount;
  currentPage = Math.min(state.page || 0, viewer.slideCount - 1);
  setupThumbnails();
  zoom(); await go(currentPage);
  if (!failures) status('');
  document.body.dataset.ready = 'true';
}
window.addEventListener('message', ({ data }) => {
  if (data.type === 'document') {
    const request = ++revision;
    loading = loading.catch(() => {}).then(() => open(data.data, request)).catch(error => status(error.message || String(error), true));
  } else if (data.type === 'error') status(data.text, true);
});
$('previous').onclick = () => run(() => go(currentPage - 1));
$('next').onclick = () => run(() => go(currentPage + 1));
$('page').onchange = () => { const target = (Number.parseInt($('page').value, 10) || 1) - 1; run(() => go(target)); };
$('zoom').onchange = zoom;
$('zoom-in').onclick = () => userZoom(effectiveZoom() + 25);
$('zoom-out').onclick = () => userZoom(effectiveZoom() - 25);
$('sidebar-toggle').onclick = () => {
  $('sidebar').hidden = !$('sidebar').hidden;
  $('sidebar-toggle').setAttribute('aria-expanded', String(!$('sidebar').hidden));
  save(); zoom();
};
$('notes-toggle').onclick = () => {
  $('notes').hidden = !$('notes').hidden; $('notes-toggle').setAttribute('aria-expanded', String(!$('notes').hidden)); zoom();
};
function find(show) { $('findbar').hidden = !show; if (show) { $('query').focus(); $('query').select(); } }
$('find-toggle').onclick = () => find($('findbar').hidden);
$('find-close').onclick = () => { find(false); clearSearch(); };
// The renderer only exposes a match's whole text-box bounds, so its highlight covers the entire
// shape. Instead highlight the matched word itself: since slide text is real HTML, find the
// occurrence in the slide DOM and paint just that range with the CSS Custom Highlight API.
const canHighlight = typeof Highlight !== 'undefined' && !!(window.CSS && CSS.highlights);
function clearSearch() {
  if (canHighlight) CSS.highlights.delete('pptx-search'); else viewer?.clearSearchHighlights();
}
async function slideElement(index) {
  for (let i = 0; i < 20; i++) {
    const el = $('slides').querySelector(`[data-slide-index="${index}"]`);
    if (el && el.textContent) return el;
    await new Promise(r => setTimeout(r, 50));
  }
  return $('slides').querySelector(`[data-slide-index="${index}"]`);
}
// A DOM Range over the `occurrence`-th case-insensitive hit of `query` inside `slideEl`,
// spanning text nodes as needed (the word may be split across runs).
function rangeForMatch(slideEl, query, occurrence) {
  const walker = document.createTreeWalker(slideEl, NodeFilter.SHOW_TEXT);
  const nodes = []; let full = '';
  for (let n = walker.nextNode(); n; n = walker.nextNode()) { nodes.push({ node: n, start: full.length }); full += n.nodeValue; }
  const needle = query.toLowerCase(); if (!needle) return null;
  let idx = full.toLowerCase().indexOf(needle), count = 0;
  while (idx !== -1 && count < occurrence) { idx = full.toLowerCase().indexOf(needle, idx + needle.length); count++; }
  if (idx === -1) return null;
  const locate = pos => nodes.find(e => pos <= e.start + e.node.nodeValue.length);
  const a = locate(idx), b = locate(idx + needle.length); if (!a || !b) return null;
  const range = document.createRange();
  range.setStart(a.node, idx - a.start); range.setEnd(b.node, idx + needle.length - b.start);
  return range;
}
async function showMatch(delta) {
  if (!matches.length || !viewer) return;
  matchIndex = (matchIndex + delta + matches.length) % matches.length;
  const match = matches[matchIndex];
  pageChanged(match.slideIndex);
  $('matches').textContent = `${matchIndex + 1} of ${matches.length}`;
  if (!canHighlight) {
    viewer.clearSearchHighlights();
    await viewer.highlightSearchResult(match, { scrollIntoView: { behavior: 'instant', block: 'center' } });
    return;
  }
  await viewer.goToSlide(match.slideIndex, { behavior: 'instant', block: 'start' });
  const slideEl = await slideElement(match.slideIndex);
  // This match's position among matches on the same slide picks which occurrence to paint.
  const occurrence = matches.slice(0, matchIndex).filter(m => m.slideIndex === match.slideIndex).length;
  const range = slideEl && rangeForMatch(slideEl, $('query').value, occurrence);
  const highlight = new Highlight(); if (range) highlight.add(range);
  CSS.highlights.set('pptx-search', highlight);
  range?.startContainer.parentElement?.scrollIntoView({ block: 'center', behavior: 'instant' });
}
$('query').oninput = () => run(async () => {
  if (!viewer) return;
  clearSearch(); matches = viewer.searchText($('query').value); matchIndex = -1;
  $('matches').textContent = matches.length ? `${matches.length} matches` : 'No matches';
  if (matches.length) await showMatch(1);
});
$('find-next').onclick = () => run(() => showMatch(1)); $('find-prev').onclick = () => run(() => showMatch(-1));
$('query').onkeydown = event => { if (event.key === 'Enter') { event.preventDefault(); run(() => showMatch(event.shiftKey ? -1 : 1)); } };
$('reload').onclick = () => { status('Reading presentation…'); api.postMessage({ type: 'reload' }); };
// Pinch-to-zoom on a trackpad arrives as wheel events with ctrlKey set; Cmd/Ctrl + scroll zooms too.
$('viewport').addEventListener('wheel', event => {
  if (!(event.ctrlKey || event.metaKey) || !viewer) return;
  event.preventDefault();
  pinchZoom(event);
}, { passive: false });
document.addEventListener('keydown', event => {
  if ((event.metaKey || event.ctrlKey) && event.key.toLowerCase() === 'f') { event.preventDefault(); find(true); return; }
  if (event.metaKey || event.ctrlKey) {
    if (event.key === '=' || event.key === '+') { event.preventDefault(); userZoom(effectiveZoom() + 10); return; }
    if (event.key === '-' || event.key === '_') { event.preventDefault(); userZoom(effectiveZoom() - 10); return; }
    if (event.key === '0') { event.preventDefault(); $('zoom').value = 'fit'; zoom(); return; }
  }
  if (event.key === 'Escape') { find(false); clearSearch(); }
  if (['INPUT', 'SELECT', 'BUTTON'].includes(event.target.tagName)) return;
  if (['ArrowRight', 'ArrowLeft', 'PageDown', 'PageUp', 'Home', 'End'].includes(event.key)) {
    event.preventDefault(); const target = event.key === 'Home' ? 0 : event.key === 'End' ? viewer?.slideCount - 1
      : currentPage + (['ArrowRight', 'PageDown'].includes(event.key) ? 1 : -1);
    run(() => go(target));
  }
});
document.addEventListener('click', event => {
  const anchor = event.target.closest('a[href]');
  if (anchor && /^(https?:|mailto:)/i.test(anchor.getAttribute('href'))) {
    event.preventDefault(); event.stopPropagation(); api.postMessage({ type: 'link', url: anchor.href });
  }
}, true);
new ResizeObserver(() => {
  clearTimeout(resizeTimer); resizeTimer = setTimeout(() => {
    if (viewer && ['fit', 'page'].includes($('zoom').value)) zoom(); // re-fit via CSS, no re-render
  }, 200);
}).observe($('viewport'));
if (state.sidebar) { $('sidebar').hidden = false; $('sidebar-toggle').setAttribute('aria-expanded', 'true'); }
api.postMessage({ type: 'ready' });
