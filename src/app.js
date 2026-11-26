/* AnCR ready reckoner: the HTML tool's interface. */
(function () {
  'use strict';
  var A = ANCR;
  var STORE_KEY = 'ancr-ready-reckoner-v1';
  var XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

  var state = A.emptyState();
  var active = 0;            // index of the key judgement on show, or 'summary'
  var storageOk = true;
  var unsaved = false;       // changes since the last Excel save, used only when storage is unavailable
  var saveTimer = null;

  var $ = function (id) { return document.getElementById(id); };
  var tabsEl = $('tabs'), panelEl = $('panel'), statusEl = $('status');

  /* ---------- small DOM helper ---------- */
  function h(tag, attrs, kids) {
    var el = document.createElement(tag), k;
    if (attrs) for (k in attrs) {
      if (!Object.prototype.hasOwnProperty.call(attrs, k) || attrs[k] === null || attrs[k] === undefined || attrs[k] === false) continue;
      if (k === 'class') el.className = attrs[k];
      else if (k === 'text') el.textContent = attrs[k];
      else el.setAttribute(k, attrs[k] === true ? '' : attrs[k]);
    }
    (kids || []).forEach(function (c) {
      if (c === null || c === undefined || c === false) return;
      el.appendChild(typeof c === 'string' ? document.createTextNode(c) : c);
    });
    return el;
  }
  function chip(band, extra) {
    return h('span', { class: 'chip ' + (band ? band.toLowerCase() : 'none') + (extra ? ' ' + extra : ''), text: band || 'Not set' });
  }

  /* ---------- storage ---------- */
  function setStatus(text, warn) {
    statusEl.textContent = text;
    statusEl.className = 'status' + (warn ? ' warn' : '');
  }
  function clock() {
    var d = new Date(), p = function (n) { return (n < 10 ? '0' : '') + n; };
    return p(d.getHours()) + ':' + p(d.getMinutes());
  }
  function persist() {
    saveTimer = null;
    try {
      if (A.stateIsEmpty(state)) window.localStorage.removeItem(STORE_KEY);
      else window.localStorage.setItem(STORE_KEY, JSON.stringify(state));
      storageOk = true;
      setStatus(A.stateIsEmpty(state) ? '' : 'Kept in this browser at ' + clock() + '. Save as Excel to keep a file.');
    } catch (e) {
      storageOk = false;
      setStatus('This browser cannot keep your work. Save as Excel before closing.', true);
    }
  }
  function changed() {
    unsaved = true;
    if (saveTimer) clearTimeout(saveTimer);
    saveTimer = setTimeout(persist, 400);
  }
  function restore() {
    try {
      var raw = window.localStorage.getItem(STORE_KEY);
      if (raw) {
        state = A.normState(JSON.parse(raw));
        setStatus('Restored the work kept in this browser.');
      }
    } catch (e) {
      storageOk = false;
    }
  }
  window.addEventListener('beforeunload', function (e) {
    if (saveTimer) { clearTimeout(saveTimer); persist(); }
    if (!storageOk && unsaved && !A.stateIsEmpty(state)) { e.preventDefault(); e.returnValue = ''; }
  });

  /* ---------- assessment details ---------- */
  var META = ['title', 'analyst', 'date', 'marking'];
  function fillMeta() {
    META.forEach(function (k) { $('meta-' + k).value = state.meta[k]; });
    showMarking();
  }
  function showMarking() {
    var m = state.meta.marking.trim();
    ['marking-top', 'marking-bottom'].forEach(function (id) {
      $(id).textContent = m;
      $(id).hidden = !m;
    });
  }
  META.forEach(function (k) {
    $('meta-' + k).addEventListener('input', function () {
      state.meta[k] = this.value;
      if (k === 'marking') showMarking();
      changed();
    });
  });

  /* ---------- tabs ---------- */
  function renderTabs() {
    tabsEl.textContent = '';
    state.kjs.forEach(function (kj, i) {
      var b = A.band(kj.scores.overall);
      tabsEl.appendChild(h('button', {
        type: 'button', class: 'tab', role: 'tab', 'data-tab': String(i),
        'aria-selected': active === i ? 'true' : 'false', tabindex: active === i ? '0' : '-1'
      }, ['KJ ' + (i + 1), b ? h('span', { class: 'dot ' + b.toLowerCase(), text: b }) : null]));
    });
    if (state.kjs.length < A.MAX_KJ) {
      tabsEl.appendChild(h('button', { type: 'button', class: 'tab add', 'data-tab': 'add', tabindex: '-1', text: '+ Add key judgement' }));
    }
    tabsEl.appendChild(h('button', {
      type: 'button', class: 'tab', role: 'tab', 'data-tab': 'summary',
      'aria-selected': active === 'summary' ? 'true' : 'false', tabindex: active === 'summary' ? '0' : '-1', text: 'Summary'
    }));
  }
  function go(tab, focus) {
    active = tab;
    renderTabs();
    renderPanel();
    if (focus) {
      var sel = tabsEl.querySelector('[aria-selected=true]');
      if (sel) sel.focus();
    }
  }
  tabsEl.addEventListener('click', function (e) {
    var t = e.target.closest ? e.target.closest('[data-tab]') : null;
    if (!t) return;
    var v = t.getAttribute('data-tab');
    if (v === 'add') {
      state.kjs.push(A.emptyKJ());
      changed();
      go(state.kjs.length - 1, true);
    } else {
      go(v === 'summary' ? 'summary' : Number(v), true);
    }
  });
  tabsEl.addEventListener('keydown', function (e) {
    if (e.key !== 'ArrowRight' && e.key !== 'ArrowLeft') return;
    var order = state.kjs.map(function (k, i) { return i; }).concat(['summary']);
    var at = order.indexOf(active);
    at = (at + (e.key === 'ArrowRight' ? 1 : order.length - 1)) % order.length;
    e.preventDefault();
    go(order[at], true);
  });

  /* ---------- key judgement panel ---------- */
  function sliderRow(kj, key, item, refLabel, isRoll) {
    var v = kj.scores[key], unset = v === null;
    var range = h('input', {
      type: 'range', min: '0', max: '100', step: '10', value: unset ? '50' : String(v),
      class: unset ? 'unset' : null, 'data-score': key,
      'aria-label': item.name + ' score, 0 to 100 in steps of 10'
    });
    range.style.setProperty('--v', unset ? 50 : v);
    var slider = h('div', { class: 'slider' }, [
      isRoll ? h('div', { class: 'zones', 'aria-hidden': 'true' }, [h('span', { text: 'Low' }), h('span', { text: 'Moderate' }), h('span', { text: 'High' })]) : null,
      range,
      h('div', { class: 'ticks', 'aria-hidden': 'true' }, A.STEPS.map(function (s) { return h('span', { text: String(s) }); }))
    ]);
    return h('div', { class: 'row' + (isRoll ? ' roll' : ''), 'data-row': key }, [
      h('div', { class: 'row-head' }, [
        refLabel ? h('span', { class: 'ref', text: refLabel }) : null,
        h('strong', { text: item.name }),
        h('p', { class: 'prompt', text: item.prompt })
      ]),
      h('div', { class: 'scale' }, [
        h('p', { class: 'anchor low' }, [h('b', { text: 'At 0' }), item.low]),
        slider,
        h('p', { class: 'anchor high' }, [h('b', { text: 'At 100' }), item.high])
      ]),
      h('div', { class: 'readout' }, [
        h('output', { class: unset ? 'unset' : null, 'data-out': key, text: unset ? 'Not set' : String(v) }),
        isRoll ? (function () { var c = chip(A.band(v), 'js-band'); c.hidden = unset; return c; })() : null,
        h('button', { type: 'button', class: 'btn small', 'data-clear': key, hidden: unset, text: 'Clear', 'aria-label': 'Clear ' + item.name + ' score' })
      ]),
      h('label', { class: 'notes' }, [
        h('span', { text: isRoll ? 'Notes: why this rating?' : 'Notes' }),
        h('textarea', { rows: '2', maxlength: '8000', 'data-note': key, text: kj.notes[key] })
      ])
    ]);
  }

  function renderKJ(i) {
    var kj = state.kjs[i], frag = document.createDocumentFragment();
    frag.appendChild(h('div', { class: 'panel-head' }, [
      h('h2', { text: 'Key judgement ' + (i + 1) }),
      state.kjs.length > 1 ? h('button', { type: 'button', class: 'btn small', id: 'btn-remove', text: 'Remove this key judgement' }) : null
    ]));
    frag.appendChild(h('label', { class: 'field' }, [
      h('span', { text: 'Key judgement' }),
      h('textarea', { rows: '3', maxlength: '8000', 'data-field': 'statement', placeholder: 'State the judgement, including its probability term.', text: kj.statement })
    ]));

    A.CRITERIA.forEach(function (c) {
      var body = h('div', { class: 'card-body' });
      c.factors.forEach(function (f) { body.appendChild(sliderRow(kj, f.id, f, f.id, false)); });
      body.appendChild(sliderRow(kj, c.id, c.rollup, '', true));
      frag.appendChild(h('section', { class: 'card' }, [h('h3', { text: 'Criterion ' + c.num + ': ' + c.name }), body]));
    });

    var mirrors = h('div', { class: 'mirrors' });
    A.CRITERIA.forEach(function (c) {
      mirrors.appendChild(h('div', { class: 'm-name', text: c.name }));
      mirrors.appendChild(h('div', { class: 'scale' }, [h('span'), h('div', { class: 'mini', 'data-mini': c.id }, [h('i')]), h('span')]));
      mirrors.appendChild(h('div', { class: 'm-chip', 'data-mchip': c.id }));
    });
    frag.appendChild(h('section', { class: 'card' }, [
      h('h3', { text: 'Overall analytical confidence' }),
      h('div', { class: 'card-body' }, [
        mirrors,
        h('p', { class: 'mirrors-note', text: 'Your three criterion ratings, carried down for reference.' }),
        sliderRow(kj, 'overall', A.OVERALL, '', true)
      ])
    ]));

    frag.appendChild(h('section', { class: 'card' }, [
      h('h3', { text: 'Ingredients for the AnCR statement' }),
      h('div', { class: 'card-body' }, [
        h('dl', { class: 'ing' }, [
          h('dt', { text: 'Overall confidence' }), h('dd', { 'data-ing': 'overall' }),
          h('dt', { text: 'Strongest factors' }), h('dd', { 'data-ing': 'strong' }),
          h('dt', { text: 'Weakest factors' }), h('dd', { 'data-ing': 'weak' }),
          h('dt', { text: 'Writing frame' }), h('dd', null, [
            h('p', { class: 'frame', 'data-ing': 'frame' }),
            h('p', { class: 'frame-actions' }, [h('button', { type: 'button', class: 'btn small', id: 'btn-use-frame', text: 'Copy frame into my statement' })])
          ])
        ]),
        h('label', { class: 'field ancr' }, [
          h('span', { text: 'Your AnCR statement' }),
          h('textarea', { rows: '6', maxlength: '8000', 'data-field': 'ancr', text: kj.ancr })
        ])
      ])
    ]));

    panelEl.textContent = '';
    panelEl.appendChild(frag);
    refreshDerived();
  }

  /* Everything that depends on the scores but is not itself a control. */
  function refreshDerived() {
    if (active === 'summary') return;
    var kj = state.kjs[active];
    A.CRITERIA.forEach(function (c) {
      var v = kj.scores[c.id];
      var mini = panelEl.querySelector('[data-mini="' + c.id + '"]');
      mini.className = 'mini' + (v === null ? ' unset' : '');
      mini.firstChild.style.left = (v === null ? 50 : v) + '%';
      var holder = panelEl.querySelector('[data-mchip="' + c.id + '"]');
      holder.textContent = '';
      holder.appendChild(chip(A.band(v)));
    });
    function put(name, text, fallback) {
      var el = panelEl.querySelector('[data-ing="' + name + '"]');
      el.textContent = text || fallback;
      if (name !== 'frame') el.className = text ? '' : 'empty';
    }
    var ob = A.band(kj.scores.overall), dd = panelEl.querySelector('[data-ing="overall"]');
    dd.textContent = '';
    dd.appendChild(chip(ob));
    put('strong', A.strongest(kj), 'Score the factors to see which are strongest.');
    put('weak', A.weakest(kj), 'Score the factors to see which are weakest.');
    put('frame', A.frame(kj), '');
  }

  function setScore(key, v) {
    var kj = state.kjs[active];
    kj.scores[key] = v;
    var row = panelEl.querySelector('[data-row="' + key + '"]');
    var range = row.querySelector('input[type=range]'), out = row.querySelector('output');
    range.className = v === null ? 'unset' : '';
    if (v === null) range.value = '50';
    range.style.setProperty('--v', v === null ? 50 : v);
    range.setAttribute('aria-valuetext', v === null ? 'Not set' : v + (row.querySelector('.js-band') ? ', ' + A.band(v) : ''));
    out.textContent = v === null ? 'Not set' : String(v);
    out.className = v === null ? 'unset' : '';
    row.querySelector('[data-clear]').hidden = v === null;
    var band = row.querySelector('.js-band');
    if (band) {
      var b = A.band(v);
      band.className = 'chip ' + (b ? b.toLowerCase() : 'none') + ' js-band';
      band.textContent = b || 'Not set';
      band.hidden = !b;
    }
    refreshDerived();
    if (key === 'overall') renderTabs();
    changed();
  }

  panelEl.addEventListener('input', function (e) {
    var t = e.target, kj = state.kjs[active];
    if (active === 'summary' || !kj) return;
    if (t.hasAttribute('data-score')) setScore(t.getAttribute('data-score'), A.normScore(t.value));
    else if (t.hasAttribute('data-note')) { kj.notes[t.getAttribute('data-note')] = t.value; changed(); }
    else if (t.hasAttribute('data-field')) { kj[t.getAttribute('data-field')] = t.value; changed(); }
  });
  /* An untouched slider is "not set" rather than 50: a click or Space/Enter on it records its position. */
  panelEl.addEventListener('click', function (e) {
    var t = e.target;
    if (t.hasAttribute && t.hasAttribute('data-score') && t.classList.contains('unset')) {
      setScore(t.getAttribute('data-score'), A.normScore(t.value));
    } else if (t.hasAttribute && t.hasAttribute('data-clear')) {
      var key = t.getAttribute('data-clear');
      setScore(key, null);
      panelEl.querySelector('[data-score="' + key + '"]').focus();
    } else if (t.id === 'btn-remove') {
      var kj = state.kjs[active];
      if (!A.kjIsEmpty(kj) && !window.confirm('Remove key judgement ' + (active + 1) + ' and everything entered for it? Later key judgements will be renumbered.')) return;
      state.kjs.splice(active, 1);
      changed();
      go(Math.min(active, state.kjs.length - 1), true);
    } else if (t.id === 'btn-use-frame') {
      var k = state.kjs[active], box = panelEl.querySelector('[data-field="ancr"]');
      if (k.ancr.trim() && !window.confirm('Replace the statement you have written with the writing frame?')) return;
      k.ancr = A.frame(k);
      box.value = k.ancr;
      box.focus();
      changed();
    } else if (t.hasAttribute && t.hasAttribute('data-goto')) {
      go(Number(t.getAttribute('data-goto')), true);
    }
  });
  panelEl.addEventListener('keydown', function (e) {
    var t = e.target;
    if ((e.key === ' ' || e.key === 'Enter') && t.hasAttribute && t.hasAttribute('data-score') && t.classList.contains('unset')) {
      e.preventDefault();
      setScore(t.getAttribute('data-score'), A.normScore(t.value));
    }
  });

  /* ---------- summary panel ---------- */
  function renderSummary() {
    var body = h('tbody');
    state.kjs.forEach(function (kj, i) {
      var cells = [
        h('td', null, [h('button', { type: 'button', class: 'kj-link', 'data-goto': String(i), text: 'KJ ' + (i + 1) })]),
        h('td', { class: 'pre' + (kj.statement ? '' : ' muted'), text: kj.statement || 'No key judgement entered yet.' })
      ];
      ['c1', 'c2', 'c3', 'overall'].forEach(function (k) { cells.push(h('td', { class: 'c' }, [chip(A.band(kj.scores[k]))])); });
      cells.push(h('td', { class: 'pre' + (kj.ancr ? '' : ' muted'), text: kj.ancr || 'Not written yet.' }));
      body.appendChild(h('tr', null, cells));
    });
    var heads = ['KJ', 'Key judgement', 'Source base', 'Analysis', 'Complexity and volatility', 'Overall confidence', 'AnCR statement'];
    panelEl.textContent = '';
    panelEl.appendChild(h('div', { class: 'panel-head' }, [h('h2', { text: 'Summary' })]));
    panelEl.appendChild(h('div', { class: 'table-wrap' }, [
      h('table', { class: 'summary' }, [
        h('thead', null, [h('tr', null, heads.map(function (t) { return h('th', { scope: 'col', text: t }); }))]),
        body
      ])
    ]));
  }
  function renderPanel() {
    if (active !== 'summary' && active >= state.kjs.length) active = state.kjs.length - 1;
    if (active === 'summary') renderSummary(); else renderKJ(active);
  }

  /* ---------- save and open ---------- */
  function fileName() {
    var d = new Date(), p = function (n) { return (n < 10 ? '0' : '') + n; };
    var stamp = '' + d.getFullYear() + p(d.getMonth() + 1) + p(d.getDate());
    var title = state.meta.title.replace(/[\\/:*?"<>|\u0000-\u001F]+/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 80);
    return stamp + ' AnCR ready reckoner' + (title ? ' - ' + title : '') + '.xlsx';
  }
  function download(blob, name) {
    var url = URL.createObjectURL(blob), a = h('a', { href: url, download: name });
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(function () { URL.revokeObjectURL(url); }, 30000);
    unsaved = false;
    setStatus('Saved ' + name + ' to your downloads folder.');
  }
  function saveXlsx() {
    var blob, name = fileName();
    try {
      blob = new Blob([A.buildWorkbook(state)], { type: XLSX_MIME });
    } catch (e) {
      setStatus('The workbook could not be created: ' + e.message, true);
      return;
    }
    if (window.showSaveFilePicker) {
      window.showSaveFilePicker({ suggestedName: name, types: [{ description: 'Excel workbook', accept: (function () { var o = {}; o[XLSX_MIME] = ['.xlsx']; return o; })() }] })
        .then(function (handle) {
          return handle.createWritable().then(function (w) {
            return w.write(blob).then(function () { return w.close(); });
          }).then(function () {
            unsaved = false;
            setStatus('Saved ' + handle.name + '.');
          });
        })
        .catch(function (err) {
          if (err && err.name === 'AbortError') return;
          download(blob, name);
        });
    } else {
      download(blob, name);
    }
  }
  function openXlsx(file) {
    var reader = new FileReader();
    reader.onerror = function () { setStatus('That file could not be read.', true); };
    reader.onload = function () {
      var next;
      try {
        next = A.readWorkbook(new Uint8Array(reader.result));
      } catch (e) {
        setStatus(e.message || 'That file could not be opened.', true);
        return;
      }
      if (!A.stateIsEmpty(state) && !window.confirm('Opening "' + file.name + '" replaces the work currently on screen. Continue?')) return;
      state = A.normState(next);
      fillMeta();
      if (saveTimer) { clearTimeout(saveTimer); saveTimer = null; }
      persist();
      go(0, false);
      if (storageOk) setStatus('Opened ' + file.name + '.');
      unsaved = false;
    };
    reader.readAsArrayBuffer(file);
  }

  $('btn-save').addEventListener('click', saveXlsx);
  $('btn-open').addEventListener('click', function () { $('file').value = ''; $('file').click(); });
  $('file').addEventListener('change', function () { if (this.files && this.files[0]) openXlsx(this.files[0]); });
  $('btn-clear').addEventListener('click', function () {
    if (!A.stateIsEmpty(state) && !window.confirm('Start again? This removes everything on screen and the copy kept in this browser. Workbooks you have already saved are not affected.')) return;
    state = A.emptyState();
    if (saveTimer) { clearTimeout(saveTimer); saveTimer = null; }
    try { window.localStorage.removeItem(STORE_KEY); } catch (e) { /* nothing to remove */ }
    unsaved = false;
    fillMeta();
    go(0, false);
    setStatus('Cleared.');
  });

  /* ---------- start ---------- */
  restore();
  fillMeta();
  renderTabs();
  renderPanel();

  /* Exposed for testing from the browser console. */
  window.ANCR_APP = { getState: function () { return state; }, build: function () { return A.buildWorkbook(state); } };
})();
