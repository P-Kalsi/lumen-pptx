const {test}=require('node:test');
const assert=require('node:assert/strict');
const fs=require('node:fs/promises');
const path=require('node:path');
const vm=require('node:vm');
const {viewerHtml}=require('../html');

test('PPTX bytes are delivered in memory; the editor never requests conversion or disk writes',async()=>{
 const bytes=Buffer.from('PK\x03\x04original pptx bytes');
 const uri={path:'/slides/lecture.pptx',scheme:'file'};
 let provider,onMessage,onDispose,onChange;
 const sent=[];
 const disposable={dispose(){}};
 const vscode={Uri:{joinPath:(base,...parts)=>({path:path.posix.join(base.path,...parts)})},
  RelativePattern:class {constructor(base,pattern){this.base=base;this.pattern=pattern;}},
  workspace:{fs:{stat:async()=>({size:bytes.length}),readFile:async input=>{assert.equal(input,uri);return bytes;}},
   createFileSystemWatcher:()=>({dispose(){},onDidChange:f=>{onChange=f;return disposable;},onDidCreate:()=>disposable,onDidDelete:()=>disposable})},
  window:{registerCustomEditorProvider:(_,p)=>{provider=p;return disposable;}},
  env:{openExternal:()=>{throw new Error('Unexpected external request');}}
 };
 const module={exports:{}};
 vm.runInNewContext(await fs.readFile(path.join(__dirname,'../extension.js'),'utf8'),{
  module,Buffer,setTimeout,clearTimeout,
  require:name=>{if(name==='vscode')return vscode;if(name==='./html')return {viewerHtml};if(name==='node:crypto')return require(name);throw new Error(`Unexpected dependency: ${name}`);}
 });
 module.exports.activate({subscriptions:[],extensionUri:{path:'/extension'}});
 const panel={webview:{cspSource:'test:',asWebviewUri:uri=>uri.path,postMessage:async m=>sent.push(m),onDidReceiveMessage:f=>{onMessage=f;return disposable;}},onDidDispose:f=>{onDispose=f;return disposable;}};
 await provider.resolveCustomEditor({uri},panel);
 onMessage({type:'ready'});await new Promise(r=>setTimeout(r,10));
 assert.equal(sent[0].type,'document');assert.deepEqual(Buffer.from(sent[0].data,'base64'),bytes);
 assert.equal(panel.webview.options.localResourceRoots.length,1);
 assert.equal(panel.webview.options.localResourceRoots[0].path,'/extension/media');
 assert.ok(!panel.webview.html.includes('slides.pdf'));
 onDispose();onMessage({type:'ready'});await new Promise(r=>setTimeout(r,10));assert.equal(sent.length,1);
});
