/* Original, dependency-free PPTX HTML player. Source layout is never reflowed. */
(() => {
  'use strict';
  const manifest = JSON.parse(document.getElementById('ppt-manifest').textContent);
  const stage = document.getElementById('stage');
  const deck = document.getElementById('deck');
  const slides = [...deck.querySelectorAll('.ppt-slide')];
  const controls = document.querySelector('.ppt-controls');
  const previous = document.getElementById('prev');
  const nextButton = document.getElementById('next');
  const jump = document.getElementById('jump');
  const editButton = document.getElementById('edit');
  const notesButton = document.getElementById('notes-toggle');
  const notesPanel = document.getElementById('notes');
  const warning = document.getElementById('player-warning');
  const width = Number(manifest.width), height = Number(manifest.height);
  if (!slides.length || !(width > 0 && height > 0)) {
    warning.hidden = false; warning.textContent = '유효한 슬라이드가 없습니다.';
    return;
  }
  let index = 1, editing = false, touch = null, auditing = false;
  const notes = Array.isArray(manifest.notes) ? manifest.notes : [];
  const sourceErrors = (manifest.issues || []).filter(x => x.severity === 'error');
  if (sourceErrors.length) {
    warning.hidden = false;
    warning.textContent = `미완료 미리보기: 지원되지 않은 변환 항목 ${sourceErrors.length}건. audit.json을 확인하세요.`;
  }
  document.getElementById('total').textContent = String(slides.length);
  jump.max = String(slides.length);
  notesButton.hidden = !notes.some(Boolean);
  deck.style.width = width + 'px'; deck.style.height = height + 'px';
  const printStyle = document.getElementById('ppt-print-size') || document.createElement('style');
  printStyle.id = 'ppt-print-size';
  printStyle.textContent = `@page{size:${width}px ${height}px;margin:0}`;
  if (!printStyle.parentNode) document.head.append(printStyle);

  function fit() {
    if (auditing) return;
    const bounds = stage.getBoundingClientRect();
    const scale = Math.max(0.01, Math.min(bounds.width / width, bounds.height / height));
    deck.style.transform = `translate(-50%,-50%) scale(${scale})`;
  }
  function valid(n) {
    n = Number(n);
    return Number.isFinite(n) ? Math.max(1, Math.min(slides.length, Math.floor(n))) : 1;
  }
  function setVisible(n) {
    slides.forEach((s, i) => {
      s.classList.toggle('is-active', i === n - 1);
      s.setAttribute('aria-hidden', i === n - 1 ? 'false' : 'true');
    });
  }
  function go(n, updateHash = true) {
    index = valid(n); setVisible(index);
    document.getElementById('current').textContent = String(index);
    jump.value = String(index);
    previous.disabled = index === 1; nextButton.disabled = index === slides.length;
    notesPanel.textContent = '이 파일에 포함된 노트입니다. 수신자도 읽을 수 있습니다.\n\n' + (notes[index - 1] || '(이 슬라이드의 노트 없음)');
    if (updateHash && location.hash !== '#' + index) {
      try { history.replaceState(null, '', '#' + index); }
      catch (_) { location.hash = String(index); }
    }
    fit();
    document.dispatchEvent(new CustomEvent('pptplayer:change'));
    return index;
  }
  function hashIndex() {
    return /^#\d+$/.test(location.hash) ? valid(location.hash.slice(1)) : valid(document.documentElement.dataset.startSlide || 1);
  }
  function setEditing(value) {
    editing = Boolean(value);
    document.body.classList.toggle('is-editing', editing);
    deck.querySelectorAll('.ppt-run[data-editable]').forEach(run => {
      if (editing) { run.setAttribute('contenteditable', 'true'); run.setAttribute('spellcheck', 'false'); }
      else { run.removeAttribute('contenteditable'); run.removeAttribute('spellcheck'); }
    });
    editButton.setAttribute('aria-pressed', String(editing));
    editButton.textContent = editing ? '편집 종료' : '편집';
    editButton.title = 'E: 편집 전환. Ctrl/Cmd+S: 수정한 HTML 저장. PPTX에는 역반영되지 않습니다.';
    document.dispatchEvent(new CustomEvent('pptplayer:change'));
    return editing;
  }
  function serialize() {
    const clone = document.documentElement.cloneNode(true);
    clone.dataset.startSlide = String(index);
    clone.querySelector('body').classList.remove('is-editing', 'is-auditing');
    clone.querySelectorAll('.dyn-auto-enter,[data-auto-motion]').forEach(el => {
      el.classList.remove('dyn-auto-enter');
      el.style.removeProperty('--dyn-delay');
      el.style.removeProperty('--dyn-duration');
      el.removeAttribute('data-auto-motion');
    });
    clone.querySelectorAll('[contenteditable]').forEach(el => { el.removeAttribute('contenteditable'); el.removeAttribute('spellcheck'); });
    const edit = clone.querySelector('#edit');
    edit.setAttribute('aria-pressed', 'false'); edit.textContent = '편집';
    clone.querySelectorAll('[data-audit-result]').forEach(el => el.remove());
    if (!sourceErrors.length) { const alert = clone.querySelector('#player-warning'); alert.hidden = true; alert.textContent = ''; }
    clone.querySelector('#notes').hidden = true;
    return '<!doctype html>\n' + clone.outerHTML;
  }
  function save() {
    const data = serialize();
    const url = URL.createObjectURL(new Blob([data], {type:'text/html;charset=utf-8'}));
    const link = document.createElement('a');
    link.href = url;
    link.download = (document.title || 'presentation').replace(/[<>:"/\\|?*\x00-\x1f]/g, '_') + '_edited.html';
    document.body.append(link); link.click(); link.remove();
    setTimeout(() => URL.revokeObjectURL(url), 3000);
  }
  async function fullscreen() {
    try {
      if (document.fullscreenElement) await document.exitFullscreen();
      else if (document.documentElement.requestFullscreen) await document.documentElement.requestFullscreen();
    } catch (_) {
      warning.hidden = false;
      warning.textContent = '브라우저가 전체 화면을 허용하지 않았습니다. 전체 화면 버튼을 직접 눌러주세요.';
    }
    fit();
  }
  function toggleNotes() {
    if (notesButton.hidden) return;
    notesPanel.hidden = !notesPanel.hidden;
    notesButton.setAttribute('aria-expanded', String(!notesPanel.hidden));
    fit();
  }
  function editableTarget(target) {
    return target instanceof Element && (target.closest('[contenteditable="true"]') || target.closest('input,textarea,select'));
  }
  function insertPlain(event) {
    if (!editing || !event.target.closest('[contenteditable="true"]')) return;
    event.preventDefault();
    const data = (event.clipboardData || event.dataTransfer).getData('text/plain');
    document.execCommand('insertText', false, data);
  }
  deck.addEventListener('paste', insertPlain);
  deck.addEventListener('drop', insertPlain);
  deck.addEventListener('beforeinput', event => {
    if (editing && event.inputType === 'insertParagraph' && event.target.closest('[contenteditable="true"]')) {
      event.preventDefault(); document.execCommand('insertLineBreak');
    }
  });
  previous.addEventListener('click', () => go(index - 1));
  nextButton.addEventListener('click', () => go(index + 1));
  jump.addEventListener('change', () => go(jump.value));
  jump.addEventListener('keydown', event => { if (event.key === 'Enter') { go(jump.value); jump.blur(); stage.focus(); } });
  editButton.addEventListener('click', () => setEditing(!editing));
  notesButton.addEventListener('click', toggleNotes);
  document.getElementById('fullscreen').addEventListener('click', fullscreen);
  window.addEventListener('hashchange', () => go(hashIndex(), false));
  window.addEventListener('resize', fit);
  document.addEventListener('fullscreenchange', fit);
  if (window.ResizeObserver) { const observer = new ResizeObserver(fit); observer.observe(stage); observer.observe(controls); }
  document.addEventListener('keydown', event => {
    const key = event.key.toLowerCase();
    if ((event.ctrlKey || event.metaKey) && key === 's') { event.preventDefault(); save(); return; }
    if (event.ctrlKey || event.metaKey || event.altKey || editableTarget(event.target)) return;
    // Space must activate a focused button once through its native click event.
    if (key === ' ' && event.target instanceof Element && event.target.closest('button,[role="button"],a[href]')) return;
    if (['arrowright','arrowdown','pagedown',' '].includes(key)) { event.preventDefault(); go(index + 1); }
    else if (['arrowleft','arrowup','pageup'].includes(key)) { event.preventDefault(); go(index - 1); }
    else if (key === 'home') { event.preventDefault(); go(1); }
    else if (key === 'end') { event.preventDefault(); go(slides.length); }
    else if (key === 'f') { event.preventDefault(); fullscreen(); }
    else if (key === 'e') { event.preventDefault(); setEditing(!editing); }
    else if (key === 'n') { event.preventDefault(); toggleNotes(); }
  });
  stage.addEventListener('touchstart', event => {
    if (!editing && event.touches.length === 1) touch = {x:event.touches[0].clientX, y:event.touches[0].clientY};
  }, {passive:true});
  stage.addEventListener('touchend', event => {
    if (!touch || editing) return;
    const dx = event.changedTouches[0].clientX - touch.x, dy = event.changedTouches[0].clientY - touch.y;
    touch = null;
    if (Math.abs(dx) > 55 && Math.abs(dx) > Math.abs(dy) * 1.5) go(index + (dx < 0 ? 1 : -1));
  }, {passive:true});

  async function audit() {
    if (auditing) throw new Error('Audit already running');
    auditing = true;
    document.body.classList.add('is-auditing');
    document.dispatchEvent(new CustomEvent('pptplayer:change'));
    const tolerance = 2.5, results = [], originalIndex = index;
    try {
      await ready;
      for (let i = 0; i < slides.length; i++) {
        setVisible(i + 1);
        if (document.fonts) await document.fonts.ready;
        const slide = slides[i], box = slide.getBoundingClientRect(), scale = box.width / width;
        const overflows = [], outOfBounds = [], manualReview = [];
        const objects = [...slide.querySelectorAll('.ppt-object')];
        for (const object of objects) {
          const r = object.getBoundingClientRect();
          const excess = Math.max(box.left-r.left, r.right-box.right, box.top-r.top, r.bottom-box.bottom) / scale;
          if (excess > tolerance) outOfBounds.push({id:object.dataset.objectId, px:Math.round(excess*10)/10});
        }
        const frames = [...slide.querySelectorAll('.ppt-text')];
        frames.forEach((frame, fi) => {
          const object = frame.closest('.ppt-object');
          const id = (object ? object.dataset.objectId : 'frame') + ':text-' + fi;
          let parent = frame.parentElement, rotated = false;
          while (parent && parent !== slide) {
            if (parent.classList.contains('ppt-object') && parent.style.transform && parent.style.transform !== 'none') rotated = true;
            parent = parent.parentElement;
          }
          if (rotated) { manualReview.push({id, reason:'rotated-text'}); return; }
          const r = frame.getBoundingClientRect(), style = getComputedStyle(frame);
          const left = r.left + parseFloat(style.paddingLeft || 0)*scale;
          const right = r.right - parseFloat(style.paddingRight || 0)*scale;
          const top = r.top + parseFloat(style.paddingTop || 0)*scale;
          const bottom = r.bottom - parseFloat(style.paddingBottom || 0)*scale;
          const walker = document.createTreeWalker(frame, NodeFilter.SHOW_TEXT);
          let node, x = 0, y = 0;
          while ((node = walker.nextNode())) {
            if (!node.nodeValue.trim()) continue;
            const range = document.createRange(); range.selectNodeContents(node);
            for (const ink of range.getClientRects()) {
              x = Math.max(x, (left-ink.left)/scale, (ink.right-right)/scale);
              y = Math.max(y, (top-ink.top)/scale, (ink.bottom-bottom)/scale);
            }
          }
          if (x > tolerance || y > tolerance) overflows.push({id,x:Math.round(x*10)/10,y:Math.round(y*10)/10,text:frame.textContent.trim().slice(0,80)});
        });
        results.push({index:i+1,overflows,outOfBounds,manualReview});
      }
    } finally {
      setVisible(originalIndex); auditing = false;
      document.body.classList.remove('is-auditing');
      document.dispatchEvent(new CustomEvent('pptplayer:change'));
      fit();
    }
    return {ok:!sourceErrors.length && results.every(s=>!s.overflows.length&&!s.outOfBounds.length),
      slides:results,sourceErrors:sourceErrors.length,tolerancePx:tolerance,
      fontsRequested:manifest.fonts || [],fontAvailability:'NOT_VERIFIED',
      powerpointVisualComparison:'NOT_PERFORMED'};
  }
  const ready = (async () => {
    go(hashIndex(), false); setEditing(false);
    if (document.fonts) await document.fonts.ready;
    fit(); return true;
  })();
  window.PPTPlayer = {go,next:()=>go(index+1),prev:()=>go(index-1),getIndex:()=>index,
    getState:()=>({index,editing,auditing}),setEditing,save,serialize,audit,ready};
})();
