/* Optional DOM regression: node tests/test_player.cjs [generated.html ...]
   Requires jsdom on the caller's module path. No browser or network is used. */
const {JSDOM}=require('jsdom');
const fs=require('fs'),path=require('path'),assert=require('node:assert/strict');
const flush=()=>new Promise(resolve=>setImmediate(resolve));
async function create(raw,{reduced=false,pendingFonts=false}={}){
  const dom=new JSDOM(raw,{runScripts:'outside-only',url:'https://presentation.test/'});
  const w=dom.window;
  w.matchMedia=()=>({matches:reduced,addEventListener(){}});
  // Geometry is intentionally unavailable: this tests state, not visual layout.
  w.Range.prototype.getClientRects=()=>[];
  let releaseFonts;
  if(pendingFonts)Object.defineProperty(w.document,'fonts',{value:{ready:new Promise(resolve=>releaseFonts=resolve)}});
  for(const script of w.document.scripts)if(!script.type||script.type==='text/javascript')w.eval(script.textContent);
  if(!pendingFonts){await w.PPTPlayer.ready;await flush();}
  return {dom,releaseFonts};
}
async function verify(file){
  const raw=fs.readFileSync(file,'utf8');
  const {dom}=await create(raw),w=dom.window,d=w.document,p=w.PPTPlayer;
  const slides=[...d.querySelectorAll('.ppt-slide')];
  const runs=[...d.querySelectorAll('.ppt-run')].map(el=>el.textContent);
  const key=k=>d.body.dispatchEvent(new w.KeyboardEvent('keydown',{key:k,bubbles:true,cancelable:true}));
  const checks=[];
  for(let n=1;n<=slides.length;n++){
    p.go(n);await flush();
    const markers=[...d.querySelectorAll('[data-auto-motion]')];
    if(slides[n-1].querySelector('[data-object-id*="_slide_"]'))assert.ok(markers.length);
    assert.equal(w.PPTAutoMotion.getState().slide,n);
    assert.ok(markers.every(el=>el.closest('.ppt-slide')===slides[n-1]));
    assert.ok(markers.every(el=>parseFloat(el.style.getPropertyValue('--dyn-delay'))<=500));
    assert.ok(markers.every(el=>parseFloat(el.style.getPropertyValue('--dyn-duration'))<=420));
    assert.ok(markers.every(el=>!/_master_|_layout_/.test(el.dataset.objectId||'')));
    assert.equal(d.querySelectorAll('.dyn-hidden,[inert],.dyn-focus-ring,.dyn-controls,[data-dynamic-trigger]').length,0);
  }
  checks.push('all slides: automatic entrance scope and <=920ms timing; no item/disclosure controls');
  p.go(1);key('ArrowRight');assert.equal(p.getIndex(),2);key(' ');assert.equal(p.getIndex(),3);
  key('ArrowLeft');assert.equal(p.getIndex(),2);d.getElementById('next').click();assert.equal(p.getIndex(),3);
  key('Home');assert.equal(p.getIndex(),1);key('End');assert.equal(p.getIndex(),slides.length);
  p.go(2);p.go(3);p.go(2);assert.equal(w.PPTAutoMotion.getState().slide,2);
  assert.ok([...d.querySelectorAll('[data-auto-motion]')].every(el=>el.closest('.ppt-slide')===slides[1]));
  const stage=d.getElementById('stage'), start=new w.Event('touchstart'),end=new w.Event('touchend');
  Object.defineProperty(start,'touches',{value:[{clientX:200,clientY:100}]});Object.defineProperty(end,'changedTouches',{value:[{clientX:80,clientY:100}]});
  stage.dispatchEvent(start);stage.dispatchEvent(end);assert.equal(p.getIndex(),3);
  const firstObject=slides[2].querySelector('.ppt-object');firstObject?.click();assert.equal(p.getIndex(),3);
  checks.push('navigation keys, buttons, return/replay, rapid changes and swipe remain one page per action');
  p.setEditing(true);assert.equal(d.querySelectorAll('[data-auto-motion]').length,0);
  p.go(2);assert.equal(d.querySelectorAll('[data-auto-motion]').length,0);p.setEditing(false);assert.equal(d.querySelectorAll('[data-auto-motion]').length,0);
  p.go(3);w.dispatchEvent(new w.Event('beforeprint'));assert.equal(d.querySelectorAll('[data-auto-motion]').length,0);
  p.go(2);assert.equal(d.querySelectorAll('[data-auto-motion]').length,0);w.dispatchEvent(new w.Event('afterprint'));
  const reduced=await create(raw,{reduced:true});reduced.dom.window.PPTPlayer.go(2);assert.equal(reduced.dom.window.document.querySelectorAll('[data-auto-motion]').length,0);
  checks.push('edit, print and reduced-motion suppress animation and leave content accessible');
  p.go(3);const saved=p.serialize(), parsed=new JSDOM(saved).window.document;
  assert.equal(parsed.querySelectorAll('.dyn-auto-enter,[data-auto-motion]').length,0);
  assert.deepEqual([...parsed.querySelectorAll('.ppt-run')].map(el=>el.textContent),runs);
  const reopened=await create(saved);assert.equal(reopened.dom.window.PPTAutoMotion.getState().slide,3);
  checks.push('serialization strips transient state, retains exact text and reopens on current slide');
  const pending=await create(raw,{pendingFonts:true}),pw=pending.dom.window;
  const audit=pw.PPTPlayer.audit();assert.equal(pw.PPTPlayer.getState().auditing,true);
  assert.equal(pw.document.querySelectorAll('[data-auto-motion]').length,0);
  pending.releaseFonts();await audit;await flush();assert.equal(pw.PPTPlayer.getState().auditing,false);
  assert.equal(pw.document.querySelectorAll('[data-auto-motion]').length,0);
  checks.push('audit suspends animation before font readiness and restores player state');
  const groups=w.PPTAutoMotion.describeGroups();
  dom.window.close();reduced.dom.window.close();reopened.dom.window.close();pending.dom.window.close();
  return {file:path.basename(file),ok:true,slides:slides.length,textRuns:runs.length,checks,groupCounts:groups.map(item=>item.groups.length),visualLayout:'NOT_TESTED'};
}
(async()=>{
  const files=process.argv.slice(2);
  if(!files.length)files.push(path.resolve(__dirname,'../examples/sample.html'));
  for(const file of files)console.log(JSON.stringify(await verify(file)));
})().catch(error=>{console.error(error);process.exitCode=1;});
