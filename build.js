const esbuild = require('esbuild');
const fs = require('fs');
const path = require('path');
// Layout fixes for @aiden0z/pptx-renderer 1.4.0, applied to its source at bundle time. Each target
// must match exactly once, so the build fails (rather than silently regressing) if the renderer changes.
const RENDERER_PATCHES = [
  // PowerPoint styles text in a plain (non-placeholder) shape from the presentation's
  // defaultTextStyle; the master's otherStyle does not apply. The renderer layers otherStyle on
  // top anyway, so old decks (default 28pt, otherStyle 18pt) render small, cramped text.
  // `d` is the shape's placeholder info, undefined for plain shapes. Table cells DO use otherStyle
  // (e.g. 18pt rather than the 28pt default), so the table renderer tags its options `tableCell`.
  ['const V = s === "title" ? n.master.textStyles.titleStyle : s === "body" ? n.master.textStyles.bodyStyle : n.master.textStyles.otherStyle;',
   'const V = s === "title" ? n.master.textStyles.titleStyle : s === "body" ? n.master.textStyles.bodyStyle : d || (a == null ? void 0 : a.tableCell) ? n.master.textStyles.otherStyle : void 0;'],
  ['        const v = {\n          defaultLineHeight: "1",\n          trimOuterParagraphSpacing: !0,',
   '        const v = {\n          tableCell: !0,\n          defaultLineHeight: "1",\n          trimOuterParagraphSpacing: !0,'],
  // Default tab stops: PowerPoint measures them from the text box's inner edge, but CSS tab-size
  // measures from the paragraph's content edge, i.e. after marL, so every column lands marL too far
  // right (and can wrap). Route all left-aligned tabbed paragraphs through the renderer's measured
  // tab layout, which already measures from the box edge and falls back to defTabSz multiples.
  ['gt = !!((F = G.tabStops) != null && F.length) && G.rtl !== !0', 'gt = G.rtl !== !0'],
  ['gt && ft.length > 0 && G.tabStops && zf(\n      P,\n      ft,\n      G.tabStops,',
   'gt && ft.length > 0 && zf(\n      P,\n      ft,\n      G.tabStops ?? [],'],
  // A tab that starts exactly on a default stop (e.g. the second of "\t\t") measures as 191.999px
  // under zoom rounding, so floor() picks the same stop and the tab collapses to zero width.
  ['position: (Math.floor(e / n) + 1) * n,\n    align: "l"', 'position: (Math.floor((e + 0.5) / n) + 1) * n,\n    align: "l"'],
  // A placeholder can carry its own fill (e.g. a picture of typeset math) with no prstGeom, since
  // geometry is inherited and defaults to a rectangle. With no geometry the renderer builds no path,
  // so the fill was never painted and the slide body came out blank. Default such shapes to "rect".
  ['    nodeType: "shape",\n    presetGeometry: r,', '    nodeType: "shape",\n    presetGeometry: r ?? (!s && u ? "rect" : void 0),'],
  // PowerPoint grows a table row to fit its text, so a table can extend past its frame's stored
  // height. The renderer clipped the table to the frame, cutting off the bottom rows' text.
  ['a.style.width = `${e.size.w}px`, a.style.height = `${e.size.h}px`, a.style.overflow = "hidden";\n  const r = [];\n  e.rotation !== 0 && r.push(`rotate(${e.rotation}deg)`)',
   'a.style.width = `${e.size.w}px`, a.style.height = `${e.size.h}px`, a.style.overflow = "visible";\n  const r = [];\n  e.rotation !== 0 && r.push(`rotate(${e.rotation}deg)`)']
];
const rendererFix = { name: 'renderer-layout-fixes', setup(build) {
  build.onLoad({ filter: /aiden0z-pptx-renderer\.es\.js$/ }, async args => {
    let source = await fs.promises.readFile(args.path, 'utf8');
    for (const [from, to] of RENDERER_PATCHES) {
      if (source.split(from).length !== 2) throw new Error(`pptx-renderer changed; re-check this patch in build.js:\n${from}`);
      source = source.replace(from, () => to);
    }
    return { contents: source, loader: 'js' };
  });
} };
(async () => {
 await esbuild.build({ entryPoints:['src/viewer.js'], plugins:[rendererFix], bundle:true, outfile:'media/viewer.js', platform:'browser', format:'iife', target:'chrome130', minify:true, define:{'process.env.NODE_ENV':'"production"'}, logLevel:'info' });
 fs.mkdirSync('licenses', {recursive:true});
 for (const name of ['@aiden0z/pptx-renderer','wmf-emf-renderer','jszip','echarts','zrender','mtx-decompressor','pako','lie','immediate','setimmediate','readable-stream','safe-buffer','string_decoder','util-deprecate','inherits','isarray','core-util-is','process-nextick-args','tslib']) {
   const root=path.join('node_modules',name); if(!fs.existsSync(root)) continue;
   for (const file of fs.readdirSync(root).filter(f=>/^(license|copying|notice|third_party_notices)/i.test(f))) {
     if(fs.statSync(path.join(root,file)).isFile()) fs.copyFileSync(path.join(root,file),path.join('licenses',name.replace(/[@/]/g,'_')+'-'+file));
   }
 }
 fs.cpSync('node_modules/@aiden0z/pptx-renderer/licenses','licenses/renderer',{recursive:true});
 fs.cpSync('node_modules/echarts/licenses','licenses/echarts',{recursive:true});
})();
