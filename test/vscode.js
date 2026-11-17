const {chromium}=require('@playwright/test');
const {spawn}=require('node:child_process');
const fs=require('node:fs/promises');
const path=require('node:path');
const assert=require('node:assert/strict');
(async()=>{
 const root=path.resolve(__dirname,'..');
 const temp=await fs.mkdtemp('/private/tmp/vscode-pptx-integration-');
 const workspace=path.join(temp,'workspace');await fs.mkdir(path.join(workspace,'.vscode'),{recursive:true});
 await fs.writeFile(path.join(workspace,'.vscode/settings.json'),JSON.stringify({'workbench.editorAssociations':{'*.pptx':'vscode-pptx.preview'}}));
 const deck=path.join(workspace,'Course Overview.pptx');await fs.symlink(path.resolve(root,'../01_CourseOverview_WhatIsAI_post-2.pptx'),deck);
 await fs.mkdir(path.join(temp,'user/User'),{recursive:true});
 await fs.writeFile(path.join(temp,'user/User/settings.json'),JSON.stringify({'security.workspace.trust.enabled':false,'workbench.startupEditor':'none','window.restoreWindows':'none','telemetry.telemetryLevel':'off','update.mode':'none'}));
 let app, child;
 try{
  child=spawn('/Applications/Visual Studio Code.app/Contents/MacOS/Electron', [`--user-data-dir=${temp}/user`,`--extensions-dir=${temp}/extensions`,`--extensionDevelopmentPath=${root}`,'--remote-debugging-port=9238','--skip-welcome','--skip-release-notes','--disable-workspace-trust',workspace,deck], {stdio:['ignore','pipe','pipe'], env:{...process.env, ELECTRON_RUN_AS_NODE:undefined}});
  child.stderr.on('data',d=>{if(String(d).includes('ERROR'))console.log(String(d).slice(0,500))});
  for(let i=0;i<60;i++){try{app=await chromium.connectOverCDP('http://127.0.0.1:9238',{timeout:3000});break;}catch{await new Promise(r=>setTimeout(r,1000));}}
  if(!app)throw new Error('VS Code debug connection unavailable');
  let window;
  for(let i=0;i<30;i++){window=app.contexts()[0].pages().find(p=>p.url().includes('workbench'));if(window)break;await new Promise(r=>setTimeout(r,1000));}
  if(!window)throw new Error('VS Code workbench not found');
  await window.waitForLoadState('domcontentloaded');
  await window.locator('.monaco-workbench').waitFor({state:'visible'});
  await window.bringToFront();
  // The workspace settings.json associates *.pptx with vscode-pptx.preview, so opening the
  // deck loads the custom editor directly — no command palette navigation required. If the deck
  // did not open as the active editor, reopen it explicitly through the file explorer.
  let frame;
  for(let attempt=0;attempt<60;attempt++){
   frame=window.frames().find(f=>f.url().includes('vscode-webview')&&f.url().includes('index.html')===false);
   for(const candidate of window.frames()){
    if(await candidate.locator('#total').count().catch(()=>0)){frame=candidate;break;}
   }
   if(frame&&await frame.locator('#total').count().catch(()=>0))break;
   // Nudge the deck open from the explorer if the association did not activate on launch.
   if(attempt===5||attempt===15){
    await window.getByText('Course Overview.pptx',{exact:true}).first().dblclick().catch(()=>{});
   }
   await new Promise(r=>setTimeout(r,1000));
  }
  if(!frame||!await frame.locator('#total').count()){
   await window.screenshot({path:'/private/tmp/vscode-pptx-integration-debug.png'});
   console.log('Windows:',app.contexts()[0].pages().map(w=>w.url()));
   console.log('Frames:',window.frames().map(f=>f.url()));
   console.log((await window.locator('body').innerText()).slice(-5000));
   throw new Error('PPTX webview did not appear');
  }
  await frame.waitForFunction(()=>document.body.dataset.ready==='true',null,{timeout:90000});
  assert.equal(await frame.locator('#total').textContent(),'/ 55');
  await frame.locator('#page').fill('5');await frame.locator('#page').dispatchEvent('change');
  await frame.waitForFunction(()=>document.querySelector('#slides').textContent.includes('Course Prerequisites'));
  await frame.locator('#sidebar-toggle').click();
  await frame.waitForFunction(()=>document.querySelector('.thumb-slide')?.children.length>0);
  await window.screenshot({path:'/private/tmp/vscode-pptx-in-vscode.png'});
  assert.equal(await frame.locator('#status.error').count(),0);
  console.log('Actual VS Code integration passed: PPTX webview loads 55 slides, navigation and thumbnails work.');
 }catch(error){if(app){const w=app.contexts()[0].pages()[0];if(w){await w.screenshot({path:'/private/tmp/vscode-pptx-integration-debug.png'});console.log((await w.locator('body').innerText()).slice(-4000));}}throw error;}finally{await app?.close();child?.kill('SIGTERM');await new Promise(r=>setTimeout(r,1000));await fs.rm(temp,{recursive:true,force:true});}
})().catch(e=>{console.error(e);process.exitCode=1});
