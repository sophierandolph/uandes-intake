/* One-question-per-screen survey engine. Each page (index.html = Part 1, part-2.html later)
   sets window.SURVEY before loading this file. Answers are POSTed as JSON to a Google Apps Script
   web app that appends one row to a Google Sheet in Sophie's Rice account. Nothing goes in the URL. */
(function () {
  const S = window.SURVEY;
  const main = document.querySelector('main');
  const bar = document.querySelector('.progress > div');
  const count = document.querySelector('.step-count');
  const DRAFT_KEY = 'uandes-intake-' + S.part;
  const SQ = ['var(--cyan)', 'var(--green)', 'var(--orange)', 'var(--red)', 'var(--mid-blue)'];

  let answers = {};
  try { answers = JSON.parse(localStorage.getItem(DRAFT_KEY) || '{}'); } catch (e) { answers = {}; }
  try {
    if (!answers.full_name && localStorage.getItem('uandes-intake-name')) answers.full_name = localStorage.getItem('uandes-intake-name');
    if (!answers.call_me && localStorage.getItem('uandes-intake-callme')) answers.call_me = localStorage.getItem('uandes-intake-callme');
  } catch (e) {}
  const save = () => { try { localStorage.setItem(DRAFT_KEY, JSON.stringify(answers)); } catch (e) {} };

  if (!S.endpoint) {
    const b = document.createElement('div');
    b.className = 'preview-banner';
    b.textContent = 'Preview mode: answers are not saved anywhere';
    document.body.prepend(b);
  }

  // Questions marked skipIfKnown are hidden when we already have the answer (e.g. the name from Part 1).
  const qs = S.questions.filter(q => !(q.skipIfKnown && answers[q.id]));
  let i = -1; // -1 = welcome, qs.length = submit/thanks

  const el = (tag, attrs = {}, ...kids) => {
    const n = document.createElement(tag);
    for (const [k, v] of Object.entries(attrs)) {
      if (k === 'class') n.className = v; else if (k === 'html') n.innerHTML = v; else n.setAttribute(k, v);
    }
    kids.flat().forEach(k => k != null && n.append(k));
    return n;
  };

  function firstName() {
    const call = (answers.call_me || '').trim();
    const raw = call || (answers.full_name || '').trim();
    const w = raw.split(/[\s(,]+/)[0];
    return w ? w.charAt(0).toUpperCase() + w.slice(1) : '';
  }

  function kicker(text, colorIdx) {
    return el('div', { class: 'kicker' }, el('span', { class: 'sq', style: 'background:' + SQ[colorIdx % SQ.length] }), text);
  }

  function setProgress() {
    const pct = i < 0 ? 0 : Math.min(100, Math.round((i / qs.length) * 100));
    bar.style.width = pct + '%';
    count.textContent = i >= 0 && i < qs.length ? (i + 1) + ' / ' + qs.length : '';
  }

  function render() {
    setProgress();
    main.innerHTML = '';
    window.scrollTo(0, 0);
    if (i < 0) return renderWelcome();
    if (i >= qs.length) return renderThanks();
    renderQuestion(qs[i]);
  }

  function renderWelcome() {
    const w = S.welcome;
    const kids = [kicker(w.kicker, 0), el('h1', {}, w.title), el('p', { class: 'lede' }, w.lede)];
    if (S.videoUrl) kids.push(el('div', { class: 'video' }, el('iframe', { src: S.videoUrl, title: 'Hello from Sophie and Hesam', allow: 'encrypted-media; picture-in-picture', allowfullscreen: '' })));
    kids.push(el('ul', { class: 'facts' }, w.facts.map(f => el('li', {}, f))));
    const start = el('button', { class: 'primary', type: 'button' }, Object.keys(answers).length ? 'Continue' : "Let's go");
    start.onclick = () => { i = 0; render(); };
    main.append(el('div', { class: 'screen' }, kids), el('nav', { class: 'actions' }, start));
  }

  function renderQuestion(q) {
    const err = el('div', { class: 'error', role: 'alert' });
    const title = typeof q.title === 'function' ? q.title({ name: firstName() }) : q.title;
    // Keep hyphenated words (Wi-Fi, T-shirt) on one line in big headings.
    const titleText = title.replace(/(\w)-(\w)/g, '$1\u2060-\u2060$2');
    const pre = q.greet && firstName() ? kicker((q.greetText || 'Nice to meet you') + ', ' + firstName() + '!', i) : kicker(q.section || S.welcome.kicker, i);
    const h = el('h2', { id: 'q-title' }, titleText, q.required ? el('span', { class: 'req', 'aria-hidden': 'true' }, '\u00a0*') : null);
    const kids = [pre, h];
    if (q.help) kids.push(el('p', { class: 'help', html: q.help }));
    const isLast = i === qs.length - 1;
    const input = buildInput(q);
    kids.push(input.node, err);

    const back = el('button', { class: 'ghost', type: 'button' }, 'Back');
    back.onclick = () => { input.commit(); i -= 1; render(); };
    const next = el('button', { class: 'primary', type: 'button' }, isLast ? 'Send' : (q.required || input.hasValue() ? 'Next' : 'Skip'));
    // Single-tap answers move on by themselves, except on the last question (Send stays a deliberate tap).
    input.autoNext = isLast ? null : () => setTimeout(() => next.click(), 250);
    input.onChange = () => { if (!isLast) next.textContent = q.required || input.hasValue() ? 'Next' : 'Skip'; err.textContent = ''; };
    next.onclick = () => {
      input.commit();
      if (q.required && !input.hasValue()) { err.textContent = 'This one is required.'; return; }
      if (isLast) return submit(next, err);
      i += 1; render();
    };
    main.append(el('div', { class: 'screen' }, kids), el('nav', { class: 'actions' }, back, next));
    const focusable = main.querySelector('input[type=text],input[type=tel],input[type=url],input[type=number],textarea');
    if (focusable) focusable.focus({ preventScroll: true });
    main.querySelectorAll('input[type=text],input[type=tel],input[type=url],input[type=number]').forEach(n => {
      n.addEventListener('keydown', e => { if (e.key === 'Enter') { e.preventDefault(); next.click(); } });
    });
  }

  function buildInput(q) {
    const api = { onChange: () => {} };
    if (q.type === 'text' || q.type === 'tel' || q.type === 'url' || q.type === 'number' || q.type === 'long') {
      const n = q.type === 'long'
        ? el('textarea', { 'aria-labelledby': 'q-title', placeholder: q.placeholder || '' })
        : el('input', { type: q.type, 'aria-labelledby': 'q-title', placeholder: q.placeholder || '', inputmode: q.type === 'number' ? 'numeric' : (q.type === 'tel' ? 'tel' : 'text'), autocomplete: q.autocomplete || 'off' });
      n.value = answers[q.id] || '';
      n.addEventListener('input', () => api.onChange());
      api.node = n;
      api.commit = () => { const v = n.value.trim(); if (v) answers[q.id] = v; else delete answers[q.id]; save(); };
      api.hasValue = () => n.value.trim().length > 0;
      return api;
    }
    if (q.type === 'scale') {
      let val = answers[q.id] != null ? String(answers[q.id]) : '';
      const grid = el('div', { class: 'scale', role: 'radiogroup', 'aria-labelledby': 'q-title' });
      const btns = [];
      for (let n = 0; n <= 10; n++) {
        const b = el('button', { type: 'button', class: 'scale-btn', role: 'radio', 'aria-checked': 'false' }, String(n));
        b.onclick = () => { val = String(n); sync(); api.onChange(); if (api.autoNext) api.autoNext(); };
        btns.push(b);
      }
      function sync() { btns.forEach((b, n) => { const on = String(n) === val; b.classList.toggle('on', on); b.setAttribute('aria-checked', on); }); }
      grid.append(...btns); sync();
      const wrap = el('div', {}, grid);
      if (q.ends) wrap.append(el('div', { class: 'scale-ends' }, el('span', {}, q.ends[0]), el('span', {}, q.ends[1])));
      api.node = wrap;
      api.commit = () => { if (val !== '') answers[q.id] = val; else delete answers[q.id]; save(); };
      api.hasValue = () => val !== '';
      return api;
    }
    // single / multi choice
    const multi = q.type === 'multi';
    let picked = new Set([].concat(answers[q.id] || []).filter(v => q.options.includes(v) || v === 'Other'));
    let otherText = answers[q.id + '_other'] || '';
    const wrap = el('div', { class: 'choices' + (q.grid ? ' grid' : ''), role: multi ? 'group' : 'radiogroup', 'aria-labelledby': 'q-title' });
    const otherBox = q.other ? el('input', { type: 'text', class: 'other-input', placeholder: 'Tell us more', 'aria-label': 'Other, please describe' }) : null;
    if (otherBox) { otherBox.value = otherText; otherBox.addEventListener('input', () => { otherText = otherBox.value; api.onChange(); }); }
    const opts = q.other ? q.options.concat('Other') : q.options;
    const labels = opts.map(opt => {
      const inp = el('input', { type: multi ? 'checkbox' : 'radio', name: q.id, value: opt });
      const lab = el('label', { class: 'choice' }, inp, el('span', { class: 'box' }), el('span', {}, opt));
      inp.checked = picked.has(opt);
      inp.addEventListener('change', () => {
        if (!multi) picked = new Set([opt]);
        else if (inp.checked) {
          if (q.exclusive && q.exclusive.includes(opt)) picked = new Set([opt]);
          else { q.exclusive && q.exclusive.forEach(x => picked.delete(x)); picked.add(opt); }
          if (q.max && picked.size > q.max) { picked.delete(opt); }
        } else picked.delete(opt);
        sync(); api.onChange();
        if (!multi && opt !== 'Other' && api.autoNext) api.autoNext();
      });
      return lab;
    });
    function sync() {
      labels.forEach(l => { const inp = l.querySelector('input'); inp.checked = picked.has(inp.value); l.classList.toggle('on', inp.checked); });
      if (otherBox) otherBox.style.display = picked.has('Other') ? '' : 'none';
    }
    wrap.append(...labels);
    if (otherBox) wrap.append(otherBox);
    sync();
    api.node = wrap;
    api.commit = () => {
      const arr = opts.filter(o => picked.has(o));
      if (arr.length) answers[q.id] = multi ? arr : arr[0]; else delete answers[q.id];
      if (picked.has('Other') && otherText.trim()) answers[q.id + '_other'] = otherText.trim(); else delete answers[q.id + '_other'];
      save();
    };
    api.hasValue = () => picked.size > 0;
    return api;
  }

  async function submit(btn, err) {
    btn.disabled = true; btn.textContent = 'Sending...';
    const row = { part: S.part, submitted_at: new Date().toISOString() };
    S.questions.filter(q => !qs.includes(q)).forEach(q => { row[q.id] = answers[q.id] || ''; }); // carried over, still sent for matching
    qs.forEach(q => {
      const v = answers[q.id];
      row[q.id] = Array.isArray(v) ? v.join('; ') : (v || '');
      if (q.other) row[q.id + '_other'] = answers[q.id + '_other'] || '';
    });
    try {
      if (S.endpoint) {
        // text/plain avoids a CORS preflight, which Apps Script web apps do not answer.
        await fetch(S.endpoint, { method: 'POST', mode: 'no-cors', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify(row) });
      } else {
        console.log('[preview] would send:', row);
      }
      try { localStorage.removeItem(DRAFT_KEY); if (answers.full_name) localStorage.setItem('uandes-intake-name', answers.full_name); if (answers.call_me) localStorage.setItem('uandes-intake-callme', answers.call_me); } catch (e) {}
      i = qs.length; render();
    } catch (e) {
      btn.disabled = false; btn.textContent = 'Send';
      err.textContent = "That didn't send. Check your connection and tap Send again. Your answers are still here.";
    }
  }

  function renderThanks() {
    bar.style.width = '100%';
    const t = S.thanks;
    const sq = el('div', { class: 'squares' }, SQ.slice(0, 4).map(c => el('span', { style: 'background:' + c })));
    const media = t.gif ? el('img', { class: 'thanks-gif', src: t.gif, alt: t.gifAlt || '' }) : null;
    if (media) media.onerror = () => media.remove(); // file not added yet: show nothing
    main.append(el('div', { class: 'screen' }, sq, el('h1', {}, typeof t.title === 'function' ? t.title(firstName()) : t.title), media, el('p', { class: 'lede' }, t.body)));
    if (t.next) {
      const go = el('a', { class: 'button primary', href: t.next.href }, t.next.label);
      main.append(el('nav', { class: 'actions' }, go));
      if (t.next.note) main.append(el('p', { class: 'help', style: 'margin-top:10px;text-align:center' }, t.next.note));
    }
    confetti();
  }

  // A one-time shower of Lilie squares. Skipped for people who turn off motion.
  function confetti() {
    if (window.matchMedia && matchMedia('(prefers-reduced-motion: reduce)').matches) return;
    const box = el('div', { class: 'confetti', 'aria-hidden': 'true' });
    const colors = ['#0037A6', '#17B9E2', '#39B54A', '#F15A29', '#EE243E', '#FFF200', '#0E6BA5'];
    for (let n = 0; n < 70; n++) {
      const size = 6 + Math.round(Math.random() * 8);
      box.append(el('span', { style: [
        'left:' + (Math.random() * 100) + '%',
        'width:' + size + 'px', 'height:' + size + 'px',
        'background:' + colors[n % colors.length],
        'animation-delay:' + (Math.random() * 0.9).toFixed(2) + 's',
        'animation-duration:' + (2.2 + Math.random() * 1.6).toFixed(2) + 's',
        '--drift:' + Math.round(Math.random() * 160 - 80) + 'px',
        '--spin:' + Math.round(Math.random() * 720 - 360) + 'deg'
      ].join(';') }));
    }
    document.body.append(box);
    setTimeout(() => box.remove(), 5000);
  }

  render();
})();
