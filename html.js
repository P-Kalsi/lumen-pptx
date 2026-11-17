function viewerHtml({ cspSource, nonce, asset }) {
  return `<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8">
  <meta name="viewport" content="width=device-width,initial-scale=1">
  <meta http-equiv="Content-Security-Policy" content="default-src 'none'; script-src 'nonce-${nonce}'; style-src ${cspSource} 'unsafe-inline'; img-src ${cspSource} blob: data:; font-src ${cspSource} blob: data:; media-src blob: data:; connect-src 'none';">
  <link rel="stylesheet" href="${asset('viewer.css')}"><title>PPTX Preview</title></head>
  <body><header role="toolbar" aria-label="Presentation controls">
  <button id="sidebar-toggle" title="Toggle thumbnails" aria-label="Toggle thumbnails" aria-expanded="false">☰</button>
  <button id="find-toggle" title="Find text (⌘F)" aria-label="Find text">⌕</button><span class="separator"></span>
  <button id="previous" title="Previous slide" aria-label="Previous slide">↑</button>
  <button id="next" title="Next slide" aria-label="Next slide">↓</button>
  <input id="page" type="number" min="1" value="1" aria-label="Slide number"><span id="total">/ —</span>
  <span class="separator"></span><button id="zoom-out" aria-label="Zoom out" title="Zoom out">−</button>
  <button id="zoom-in" aria-label="Zoom in" title="Zoom in">+</button>
  <select id="zoom" aria-label="Zoom"><option value="fit">Fit width</option><option value="page">Fit slide</option><option value="50">50%</option><option value="75">75%</option><option value="100">100%</option><option value="125">125%</option><option value="150">150%</option><option value="200">200%</option><option value="300">300%</option><option value="400">400%</option></select>
  <span class="spacer"></span><button id="notes-toggle" title="Speaker notes" aria-label="Speaker notes" aria-expanded="false">Notes</button>
  <button id="reload" title="Reload presentation" aria-label="Reload presentation">↻</button>
  </header>
  <div id="findbar" hidden><input id="query" type="search" placeholder="Find in presentation…" aria-label="Find in presentation"><button id="find-prev" aria-label="Previous match">↑</button><button id="find-next" aria-label="Next match">↓</button><span id="matches" role="status"></span><button id="find-close" aria-label="Close search">×</button></div>
  <div id="status" role="status">Reading presentation…</div>
  <div id="workspace"><aside id="sidebar" aria-label="Slide thumbnails" hidden></aside>
  <main id="viewport" tabindex="0" aria-label="Presentation slides"><div id="slides"></div></main>
  <aside id="notes" aria-label="Speaker notes" hidden><h2>Speaker notes</h2><div id="notes-text"></div></aside></div>
  <script nonce="${nonce}" src="${asset('viewer.js')}"></script></body></html>`;
}
module.exports = { viewerHtml };
