/* Dependency-free automatic entrances, with explicit cumulative procedure groups. */
(() => {
  'use strict';
  const player = window.PPTPlayer;
  const manifest = JSON.parse(document.getElementById('ppt-manifest').textContent);
  if (!player || manifest.motion?.mode !== 'auto') return;
  const deck = document.getElementById('deck');
  const reduced = window.matchMedia ? window.matchMedia('(prefers-reduced-motion: reduce)') : {matches:false};
  const duration = Math.max(0, Math.min(420, Number(manifest.motion.duration_ms) || 420));
  const maxDelay = Math.max(0, Math.min(500, Number(manifest.motion.max_delay_ms) || 500));
  const procedural = manifest.procedural_reveal || {};
  const proceduralBySlide = new Map((procedural.slides || []).map(spec => [Number(spec.slide),spec]));
  const procedureInterval = Number(procedural.interval_ms) || 300;
  const procedureInitialDelay = Number(procedural.initial_delay_ms) || 150;
  const epsilon = Math.max(2, Number(manifest.width) / 320);
  const sourceBox = element => {
    const x = parseFloat(element.style.left) || 0, y = parseFloat(element.style.top) || 0;
    const w = parseFloat(element.style.width) || 0, h = parseFloat(element.style.height) || 0;
    return {x,y,w,h,right:x+w,bottom:y+h};
  };
  function bounds(items) {
    const x = Math.min(...items.map(item=>item.box.x)), y = Math.min(...items.map(item=>item.box.y));
    const right = Math.max(...items.map(item=>item.box.right)), bottom = Math.max(...items.map(item=>item.box.bottom));
    return {x,y,right,bottom,w:right-x,h:bottom-y};
  }
  function group(items, kind) { return {elements:items.flatMap(item=>item.elements),box:bounds(items),kind}; }
  function arrowDirection(item) {
    const polygon = item.elements[0].querySelector('svg polygon');
    if (!polygon) return null;
    const numbers = (polygon.getAttribute('points') || '').trim().split(/[\s,]+/).map(Number);
    if (numbers.length < 6 || numbers.length % 2 || numbers.some(value=>!Number.isFinite(value))) return null;
    const xs = numbers.filter((_,index)=>index%2===0), left = Math.min(...xs), right = Math.max(...xs);
    const atLeft = xs.filter(x=>Math.abs(x-left)<.001).length, atRight = xs.filter(x=>Math.abs(x-right)<.001).length;
    if (atRight === 1 && atLeft > 1) return 'right';
    if (atLeft === 1 && atRight > 1) return 'left';
    return null;
  }
  // Overlap along one source axis identifies a row or column, without DOM wrapping.
  function bands(items, axis, gap = epsilon) {
    const end = axis === 'x' ? 'right' : 'bottom';
    const sorted = [...items].sort((a,b)=>a.box[axis]-b.box[axis]);
    const groups = [];
    for (const item of sorted) {
      const last = groups[groups.length-1];
      if (last && item.box[axis] <= last.end + gap) {
        last.items.push(item); last.end = Math.max(last.end,item.box[end]);
      } else groups.push({items:[item],end:item.box[end]});
    }
    return groups.map(entry=>entry.items);
  }
  function contentGroups(items, arrows) {
    if (!items.length) return arrows.length ? [group(arrows,'together')] : [];
    const area = bounds(items);
    const wide = items.filter(item=>item.box.w >= area.w*.78);
    const narrow = items.filter(item=>!wide.includes(item));
    const columns = bands(narrow,'x');
    const boxes = columns.map(bounds);
    const comparable = boxes.length >= 2 && boxes.length <= 6 &&
      Math.max(...boxes.map(box=>box.w)) <= Math.min(...boxes.map(box=>box.w))*1.5 &&
      Math.min(...boxes.map(box=>box.bottom)) - Math.max(...boxes.map(box=>box.y)) >=
        Math.min(...boxes.map(box=>box.h))*.5;
    if (comparable) {
      const result = columns.map(column=>group(column,'column'));
      // Pair only a recognizable arrow tip between aligned source columns.
      for (const arrow of arrows) {
        const direction = arrowDirection(arrow);
        const choices = result.slice(1).map((column,i)=>({column,i})).filter(({column,i})=>
          arrow.box.x >= result[i].box.right-epsilon && arrow.box.right <= column.box.x+epsilon &&
          arrow.box.y >= column.box.y-epsilon && arrow.box.bottom <= column.box.bottom+epsilon);
        if (!direction || choices.length !== 1) return [group([...items,...arrows],'together')];
        const target = direction === 'right' ? choices[0].column : result[choices[0].i];
        target.elements.push(...arrow.elements);
      }
      // Wide notes or conclusion bands follow their source vertical position.
      const extras = bands(wide,'y',epsilon*2).map(row=>group(row,'row'));
      if (extras.some(extra=>extra.box.y < Math.min(...boxes.map(box=>box.bottom)) &&
          extra.box.bottom > Math.max(...boxes.map(box=>box.y)))) return [group([...items,...arrows],'together')];
      return [...result,...extras].sort((a,b)=>a.box.y-b.box.y || a.box.x-b.box.x);
    }
    if (arrows.length) return [group([...items,...arrows],'together')];
    return bands(items,'y',epsilon*2).map(row=>group(row,'row'));
  }
  function makePlan(slide) {
    const spec = proceduralBySlide.get(Number(slide.id.replace('slide-','')));
    if (spec) {
      const entries = [...spec.groups, ...(spec.conclusionIds.length ? [{label:'결론',ids:spec.conclusionIds,conclusion:true}] : [])];
      const plan = entries.map(entry => ({
        elements:entry.ids.map(id=>document.getElementById(id)),
        kind:entry.conclusion?'procedure-conclusion':'procedure-step',label:entry.label
      }));
      if (plan.every(entry=>entry.elements.length && entry.elements.every(element=>element && slide.contains(element)))) return plan;
      console.warn('Procedural mapping unavailable for slide '+spec.slide+'; keeping all content visible.');
      return [];
    }
    // Converter IDs carry source origin; keep inherited master/layout furniture static.
    const objects = [...slide.children].filter(element=>element.classList.contains('ppt-object') &&
      (element.dataset.objectId || '').includes('_slide_'));
    const tables = objects.filter(element=>element.querySelector('.ppt-table'));
    const named = objects.filter(element=>/^body-/i.test(element.dataset.name || '') && !tables.includes(element));
    const connectors = objects.filter(element=>/^flow-arrow-/i.test(element.dataset.name || ''));
    const item = element=>({elements:[element],box:sourceBox(element)});
    const entries = contentGroups(named.map(item),connectors.map(item));
    for (const object of tables) {
      const table = object.querySelector('.ppt-table'), box = sourceBox(object);
      const rows = [...table.rows];
      if ([...table.querySelectorAll('td,th')].some(cell=>cell.rowSpan > 1)) {
        entries.push({elements:[object],box,kind:'merged-table'}); continue;
      }
      let y = box.y;
      for (const row of rows) {
        const h = parseFloat(row.style.height) || box.h/Math.max(1,rows.length);
        entries.push({elements:[row],box:{...box,y,h,bottom:y+h},kind:'table-row'}); y += h;
      }
    }
    // Unknown templates get one simultaneous content entrance. Do not invent an order.
    if (!entries.length) {
      const content = objects.length ? objects : [...slide.querySelectorAll('.ppt-text')];
      if (content.length) entries.push({elements:content,box:{x:0,y:0},kind:'together'});
    }
    // Interleaved table/card layouts have no reliable reading order: settle together.
    if (tables.length && named.some(element=>tables.some(table=>{
      const a=sourceBox(element), b=sourceBox(table); return a.y<b.bottom && a.bottom>b.y;
    }))) return [{elements:[...new Set(entries.flatMap(entry=>entry.elements))],kind:'together'}];
    return entries.sort((a,b)=>a.box.y-b.box.y || a.box.x-b.box.x);
  }
  const records = [...deck.querySelectorAll('.ppt-slide')].map(slide=>{
    const groups=makePlan(slide);
    return {slide,groups,procedural:groups.some(entry=>entry.kind==='procedure-step')};
  });
  let active = player.getIndex()-1, state = player.getState(), printing = false, ready = false;
  const timers = new Set();
  let generation = 0, sequence = null;
  function release(element) {
    element.classList.remove('dyn-procedure-pending','dyn-auto-enter');
    element.removeAttribute('data-auto-motion');
    element.style.removeProperty('--dyn-delay');
    element.style.removeProperty('--dyn-duration');
  }
  function clear() {
    generation++;
    timers.forEach(id=>clearTimeout(id)); timers.clear(); sequence=null;
    deck.querySelectorAll('.dyn-procedure-pending,.dyn-auto-enter,[data-auto-motion]').forEach(release);
  }
  function schedule(callback,delay,run) {
    const id=setTimeout(()=>{
      timers.delete(id);
      if (run!==generation) return;
      const current=player.getState();
      if (current.index-1!==active || current.editing || current.auditing || printing || reduced.matches) { clear(); return; }
      try { callback(); } catch(error) { clear(); console.error('Entrance recovered with all content visible.',error); }
    },delay);
    timers.add(id);
  }
  function playProcedure(record) {
    const run=++generation;
    const total=procedureInitialDelay+(record.groups.length-1)*procedureInterval+duration;
    sequence={startedAt:Date.now(),totalMs:total,shownGroups:0,totalGroups:record.groups.length};
    // Register the final cleanup before hiding anything. Visibility does not depend on CSS completing.
    schedule(clear,total+80,run);
    record.groups.forEach((entry,i)=>{
      entry.elements.forEach(element=>{
        element.setAttribute('data-auto-motion','');
        element.classList.add('dyn-procedure-pending');
      });
      schedule(()=>{
        entry.elements.forEach(element=>{
          element.style.setProperty('--dyn-delay','0ms');
          element.style.setProperty('--dyn-duration',duration+'ms');
          element.classList.add('dyn-auto-enter');
          element.classList.remove('dyn-procedure-pending');
        });
        if (sequence) sequence.shownGroups=i+1;
        // Each revealed group becomes plain, permanently visible source content.
        schedule(()=>entry.elements.forEach(release),duration+20,run);
      },procedureInitialDelay+i*procedureInterval,run);
    });
  }
  function play() {
    if (!ready || state.editing || state.auditing || printing || reduced.matches) return;
    const record = records[active];
    if (!record) return;
    if (record.procedural) { playProcedure(record); return; }
    const interval = Math.min(100,maxDelay/Math.max(1,record.groups.length-1));
    record.groups.forEach((entry,i)=>entry.elements.forEach(element=>{
      element.style.setProperty('--dyn-delay',Math.round(i*interval)+'ms');
      element.style.setProperty('--dyn-duration',duration+'ms');
      element.setAttribute('data-auto-motion',''); element.classList.add('dyn-auto-enter');
    }));
  }
  function sync() {
    const next = player.getState(), changed = next.index-1 !== active;
    const modeChanged = next.editing !== state.editing || next.auditing !== state.auditing;
    state = next;
    if (!changed && !modeChanged) return;
    clear(); active=next.index-1;
    if (changed) play();
  }
  document.addEventListener('pptplayer:change',sync);
  if (reduced.addEventListener) reduced.addEventListener('change',clear);
  window.addEventListener('beforeprint',()=>{printing=true;clear();});
  window.addEventListener('afterprint',()=>{printing=false;});
  document.addEventListener('visibilitychange',()=>{if(sequence) clear();});
  window.addEventListener('pageshow',event=>{if(event.persisted) clear();});
  window.addEventListener('focus',()=>{if(sequence && Date.now()-sequence.startedAt>=sequence.totalMs) clear();});
  window.PPTAutoMotion = {
    getState:()=>({slide:active+1,editing:state.editing,auditing:state.auditing,reducedMotion:reduced.matches,printing,
      scheduler:'cumulative-timers',pendingTimers:timers.size,sequence:sequence?{...sequence}:null,
      durationMs:duration,maxDelayMs:records[active]?.procedural?procedureInitialDelay+(records[active].groups.length-1)*procedureInterval:maxDelay,groups:records[active]?.groups.length || 0,
      procedural:Boolean(records[active]?.procedural),
      intervalMs:records[active]?.procedural?procedureInterval:null,
      initialDelayMs:records[active]?.procedural?procedureInitialDelay:0}),
    describeGroups:()=>records.map((record,index)=>({slide:index+1,procedural:record.procedural,groups:record.groups.map(entry=>({
      kind:entry.kind,label:entry.label || '',objects:entry.elements.map(element=>element.dataset.objectId || element.closest('.ppt-object')?.dataset.objectId || ''),
      names:entry.elements.map(element=>element.dataset.name || ''),size:entry.elements.length
    }))})),clear
  };
  // Wait for font layout, and re-read state because an audit can start immediately.
  player.ready.then(()=>{ready=true;state=player.getState();active=state.index-1;play();});
})();
