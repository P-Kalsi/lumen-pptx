const { chromium } = require('@playwright/test');
const fs = require('node:fs/promises');
const path = require('node:path');
const http = require('node:http');
const assert = require('node:assert/strict');
const JSZip = require('jszip');
const { viewerHtml } = require('../html');
(async () => {
  const root = path.resolve(__dirname,'..'); let origin;
  const server = http.createServer(async (req,res) => {
    try {
      if(req.url === '/') {res.setHeader('Content-Type','text/html');res.end(viewerHtml({cspSource:origin,nonce:'test',asset:name=>origin+'/media/'+name}));return;}
      const file=path.join(root,new URL(req.url,origin).pathname);
      if(!file.startsWith(path.join(root,'media')+path.sep)){res.writeHead(404).end();return;}
      res.setHeader('Content-Type',file.endsWith('.js')?'text/javascript':'text/css');res.end(await fs.readFile(file));
    }catch{res.writeHead(404).end();}
  });
  await new Promise(r=>server.listen(0,'127.0.0.1',r));origin=`http://127.0.0.1:${server.address().port}`;
  let browser;
  try {
    browser=await chromium.launch({executablePath:'/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',headless:true});
    const page=await browser.newPage({viewport:{width:1440,height:1100}});
    await page.addInitScript(()=>{let s;window.acquireVsCodeApi=()=>({getState:()=>s,setState:v=>s=v,postMessage:()=>{}})});
    const errors=[];page.on('pageerror',e=>errors.push(e.message));
    const deckDir=path.resolve(process.env.PPTX_DIR||path.join(root,'..'));
    const decks=(await fs.readdir(deckDir)).filter(n=>n.endsWith('.pptx')).sort();
    for(const name of decks){
      await page.goto(origin);
      const bytes=await fs.readFile(path.join(deckDir,name));const zip=await JSZip.loadAsync(bytes);
      const xml=await zip.file('ppt/presentation.xml').async('string');const expected=(xml.match(/<p:sldId\s/g)||[]).length;
      await page.evaluate(data=>window.postMessage({type:'document',data},'*'),bytes.toString('base64'));
      await page.waitForFunction(()=>document.body.dataset.ready==='true',null,{timeout:90000});
      assert.equal(await page.locator('#total').textContent(),`/ ${expected}`);
      assert.equal(await page.locator('#status.error').count(),0,`${name}: render error`);
      if(name.startsWith('05_')){
        // Slide 25 has WMF equation images the renderer cannot decode; we fill them from the zip.
        await page.locator('#page').fill('25');await page.locator('#page').dispatchEvent('change');
        await page.waitForTimeout(300);
        const wmf=await page.evaluate(()=>{let ph=0,svg=0;document.querySelectorAll('#slides [data-slide-index="24"] div').forEach(d=>{if(d.childElementCount===0&&/Unsupported format:\s*WMF/i.test(d.textContent||''))ph++;});document.querySelectorAll('#slides [data-slide-index="24"] svg').forEach(()=>svg++);return {ph,svg};});
        assert.equal(wmf.ph,0,'WMF placeholders must be replaced');
        assert.ok(wmf.svg>0,'WMF slide should contain decoded SVG');
      }
      if(name.startsWith('01_')){
        await page.locator('#page').fill('5');await page.locator('#page').dispatchEvent('change');
        await page.waitForFunction(()=>document.querySelector('#slides').textContent.includes('Course Prerequisites'));
        await page.screenshot({path:'/private/tmp/vscode-pptx-slide5.png'});
        await page.locator('#sidebar-toggle').click();
        await page.waitForFunction(()=>document.querySelector('.thumb-slide')?.children.length>0);
        await page.locator('#find-toggle').click();await page.locator('#query').fill('Prerequisite');
        await page.waitForFunction(()=>document.querySelector('#matches').textContent.includes('of'));
        // The highlight must cover the matched word only, not the whole text box/shape.
        await page.waitForFunction(()=>{const h=CSS.highlights.get('pptx-search');return h&&[...h].length>0;},{timeout:10000});
        const highlighted=await page.evaluate(()=>{const r=[...CSS.highlights.get('pptx-search')][0];return r?r.toString():null;});
        assert.equal((highlighted||'').toLowerCase(),'prerequisite','search highlights the word, not the shape');
        await page.locator('#find-close').click();
        await page.locator('#zoom').selectOption('150');
        // Trackpad pinch-zoom: a ctrlKey wheel event must change the live scale.
        const before=await page.evaluate(()=>document.querySelector('#slides [data-slide-index]').getBoundingClientRect().width);
        await page.locator('#viewport').dispatchEvent('wheel',{deltaY:-120,ctrlKey:true});
        await page.waitForFunction(w=>document.querySelector('#slides [data-slide-index]').getBoundingClientRect().width>w+2,before,{timeout:10000});
        // Pinch must zoom live via persistent CSS zoom, never a renderer re-render.
        assert.ok(parseFloat(await page.evaluate(()=>document.querySelector('#slides').style.zoom||'0'))>1,'pinch applies CSS zoom');
        // Repeated zoom-out must keep shrinking — guards the feedback-loop runaway/inversion.
        const zoomedIn=await page.evaluate(()=>document.querySelector('#slides [data-slide-index]').getBoundingClientRect().width);
        for(let i=0;i<6;i++){await page.locator('#viewport').dispatchEvent('wheel',{deltaY:80,ctrlKey:true});await page.waitForTimeout(20);}
        const zoomedOut=await page.evaluate(()=>document.querySelector('#slides [data-slide-index]').getBoundingClientRect().width);
        assert.ok(zoomedOut<zoomedIn,`zoom-out must shrink, not invert (${zoomedIn} -> ${zoomedOut})`);
        await page.locator('#zoom').selectOption('150');
        await page.locator('#page').fill('9999');await page.locator('#page').dispatchEvent('change');
        await page.waitForFunction(()=>document.querySelector('#next').disabled);
      }
      // Visit the final slide as well, exercising windowed mounting across the deck.
      await page.locator('#page').fill(String(expected));await page.locator('#page').dispatchEvent('change');
      await page.waitForFunction(()=>document.querySelector('#next').disabled);
      console.log(`${name}: ${expected} slides loaded; first/last slide navigation passed`);
    }
    assert.deepEqual(errors,[]);console.log('All browser checks passed.');
  }finally{await browser?.close();await new Promise(r=>server.close(r));}
})().catch(e=>{console.error(e);process.exitCode=1});
