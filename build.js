const esbuild = require('esbuild');
const fs = require('fs');
const path = require('path');
(async () => {
 await esbuild.build({ entryPoints:['src/viewer.js'], bundle:true, outfile:'media/viewer.js', platform:'browser', format:'iife', target:'chrome130', minify:true, define:{'process.env.NODE_ENV':'"production"'}, logLevel:'info' });
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
