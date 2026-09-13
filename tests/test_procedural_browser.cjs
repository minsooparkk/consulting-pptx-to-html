/* Optional real-time browser regression. Requires Playwright and local Chromium.
   node tests/test_procedural_browser.cjs generated-procedures.html
   Set PPT_HTML_BROWSER to use an installed Chrome/Edge executable.
   Does not seek/finish animations, clear them, or disable them for normal checks. */
'use strict';
const fs=require('node:fs'),path=require('node:path'),assert=require('node:assert/strict');
const {pathToFileURL}=require('node:url');
const {chromium}=require('playwright');

(async()=>{
  assert.ok(process.argv[2],'Provide HTML generated with --procedure-manifest');
  const file=path.resolve(process.argv[2]),original=fs.readFileSync(file);
  const browser=await chromium.launch({headless:true,...(process.env.PPT_HTML_BROWSER?{executablePath:process.env.PPT_HTML_BROWSER}:{})});
  const report={ok:false,slides:[],checks:[],pageErrors:[],externalRequests:[]};
  try{
    const context=await browser.newContext({viewport:{width:1440,height:900},offline:true,reducedMotion:'no-preference'});
    context.on('request',r=>{if(/^https?:/.test(r.url()))report.externalRequests.push(r.url());});
    const page=await context.newPage();page.on('pageerror',e=>report.pageErrors.push(String(e)));
    await page.goto(pathToFileURL(file).href);await page.evaluate(()=>PPTPlayer.ready);
    const specs=await page.evaluate(()=>JSON.parse(document.getElementById('ppt-manifest').textContent).procedural_reveal);
    assert.ok(specs?.slides.length,'Expected explicit procedure groups');
    const state=async()=>page.evaluate(()=>{
      const s=PPTAutoMotion.getState(),plan=PPTAutoMotion.describeGroups().find(x=>x.slide===s.slide);
      return {allVisible:plan.groups.every(g=>g.objects.every(id=>Number(getComputedStyle(document.getElementById(id)).opacity)>.999)),
        pending:document.querySelectorAll('.dyn-procedure-pending').length,
        timers:s.pendingTimers,enter:document.querySelectorAll('.dyn-auto-enter').length};
    });
    for(const spec of specs.slides){
      const observed=await page.evaluate(async spec=>{
        // Leave the current slide first, so re-entry always starts a fresh run.
        PPTPlayer.go(spec.slide===1?2:1,false);PPTPlayer.go(spec.slide,false);
        const cfg=JSON.parse(document.getElementById('ppt-manifest').textContent).procedural_reveal;
        const groups=[...spec.groups,...(spec.conclusionIds.length?[{ids:spec.conclusionIds}]:[])];
        const start=performance.now(),onsets=groups.map(()=>null),fullyVisible=groups.map(()=>false);
        const end=cfg.initial_delay_ms+(groups.length-1)*cfg.interval_ms+cfg.duration_ms+350;
        let cumulative=true,staticVisible=true,final=[];
        while(performance.now()-start<end){
          final=groups.map(g=>g.ids.map(id=>Number(getComputedStyle(document.getElementById(id)).opacity)));
          final.forEach((values,i)=>{
            if(values.every(v=>v>0)&&onsets[i]===null)onsets[i]=performance.now()-start;
            if(fullyVisible[i]&&!values.every(v=>v>.999))cumulative=false;
            if(values.every(v=>v>.999))fullyVisible[i]=true;
          });
          if(!spec.staticIds.every(id=>Number(getComputedStyle(document.getElementById(id)).opacity)>.999))staticVisible=false;
          await new Promise(resolve=>setTimeout(resolve,25));
        }
        return {slide:spec.slide,onsets,cumulative,staticVisible,finalAllVisible:final.every(v=>v.every(x=>x>.999)),state:PPTAutoMotion.getState()};
      },spec);
      assert.ok(observed.onsets.every(x=>x!==null),'Every stage must appear before completion');
      const intervals=observed.onsets.slice(1).map((v,i)=>v-observed.onsets[i]);
      assert.ok(intervals.every(ms=>Math.abs(ms-specs.interval_ms)<Math.max(100,specs.interval_ms*.25)),`Observed intervals: ${intervals}`);
      assert.ok(observed.cumulative&&observed.staticVisible&&observed.finalAllVisible,'Prior groups and static furniture remain visible');
      assert.deepEqual(await state(),{allVisible:true,pending:0,timers:0,enter:0});
      report.slides.push({...observed,intervals});
    }
    report.checks.push('real-time onset intervals, cumulative groups, static furniture and completion cleanup');
    const first=specs.slides[0],duration=specs.initial_delay_ms+first.groups.length*specs.interval_ms+specs.duration_ms+350;
    const reenter=async()=>page.evaluate(n=>{PPTPlayer.go(n===1?2:1,false);PPTPlayer.go(n,false);},first.slide);
    // Fault injection: CSS completion cannot be required to restore source content.
    const paused=await page.addStyleTag({content:'.dyn-auto-enter{animation-play-state:paused!important}'});
    await reenter();await page.waitForTimeout(duration);
    assert.deepEqual(await state(),{allVisible:true,pending:0,timers:0,enter:0});await paused.evaluate(el=>el.remove());
    report.checks.push('paused CSS still reaches all visible content');
    await reenter();const saved=await page.evaluate(()=>PPTPlayer.serialize());
    assert.ok(await page.evaluate(s=>new DOMParser().parseFromString(s,'text/html').querySelectorAll('.dyn-procedure-pending,.dyn-auto-enter,[data-auto-motion]').length===0,saved));
    const reopened=await context.newPage();await reopened.setContent(saved);await reopened.evaluate(()=>PPTPlayer.ready);await reopened.waitForTimeout(duration);
    assert.ok(await reopened.evaluate(()=>{
      const p=PPTAutoMotion.describeGroups().find(x=>x.slide===PPTPlayer.getIndex());
      return p.groups.every(g=>g.objects.every(id=>Number(getComputedStyle(document.getElementById(id)).opacity)>.999))&&PPTAutoMotion.getState().pendingTimers===0;
    }));await reopened.close();report.checks.push('mid-run save strips hidden state and reopens cumulatively');
    await reenter();await page.evaluate(()=>PPTPlayer.setEditing(true));assert.ok((await state()).allVisible);await page.evaluate(()=>PPTPlayer.setEditing(false));
    await reenter();await page.evaluate(()=>dispatchEvent(new Event('beforeprint')));assert.ok((await state()).allVisible);await page.evaluate(()=>dispatchEvent(new Event('afterprint')));
    await page.emulateMedia({reducedMotion:'reduce'});await reenter();assert.ok((await state()).allVisible);
    assert.ok((await page.evaluate(()=>PPTPlayer.audit())).ok);
    report.checks.push('editing, print, reduced motion and layout audit');
    assert.equal(report.pageErrors.length,0);assert.equal(report.externalRequests.length,0);assert.deepEqual(fs.readFileSync(file),original);
    report.ok=true;
  }finally{await browser.close();console.log(JSON.stringify(report));}
})().catch(error=>{console.error(error);process.exitCode=1;});
