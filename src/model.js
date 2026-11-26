/* AnCR ready reckoner: shared model.
 * Criteria, factors, wording and the small amount of logic that both the
 * HTML tool and the Excel workbook rely on. No DOM access in this file. */
var ANCR = (function () {
  'use strict';

  var MAX_KJ = 10;
  var STEPS = [0, 10, 20, 30, 40, 50, 60, 70, 80, 90, 100];

  var TITLE = 'Analytical Confidence Rating (AnCR) ready reckoner';
  var SUBTITLE = 'A structured aid for judging analytical confidence. Every rating is the analyst’s judgement, not a calculation.';

  var CRITERIA = [
    {
      id: 'c1', num: 1, name: 'Source base',
      factors: [
        { id: '1a', name: 'Coverage',
          prompt: 'How much of the available information has been examined?',
          low: 'Coverage is limited. There are significant gaps and the judgement rests heavily on assumption.',
          high: 'Coverage is comprehensive, with no significant gaps.' },
        { id: '1b', name: 'Consistency',
          prompt: 'Is the information broadly in agreement, or is there substantial conflicting evidence?',
          low: 'Sources disagree significantly, or source bias is encountered frequently.',
          high: 'Sources are fully consistent with one another.' },
        { id: '1c', name: 'Source reliability',
          prompt: 'How far are the sources trusted, and how valid is the information they provide?',
          low: 'Denial or deception is evident, information is second- or third-hand, or sources are repeating one another.',
          high: 'Sources are reliable and have good access and insight.' },
        { id: '1d', name: 'Background knowledge',
          prompt: 'Is the analyst a subject-matter expert, or new to the topic?',
          low: 'The analyst is new to the topic or is not a specialist.',
          high: 'The analyst is an expert with a strong grasp of the context and background.' }
      ],
      rollup: {
        name: 'Source base overall',
        prompt: 'Taken together, how strong is the source base? Set this yourself; it is not an average.',
        low: 'The source base is weak. There may be significant unknown unknowns.',
        high: 'The source base is strong.'
      },
      frameSubject: 'The source base supports ',
      frameHint: '[coverage, consistency, source reliability and background knowledge]'
    },
    {
      id: 'c2', num: 2, name: 'Analysis',
      factors: [
        { id: '2a', name: 'Argumentation',
          prompt: 'Is there a clearly mapped argument?',
          low: 'The argument is weak, assumptions are unstated, and structured analysis is unclear or absent.',
          high: 'The argument is clear and mapped, and bias has been mitigated.' },
        { id: '2b', name: 'Assumptions',
          prompt: 'Have the load-bearing assumptions been identified?',
          low: 'Assumptions have not been identified or examined.',
          high: 'An assumptions check is complete and the load-bearing assumptions have been explored.' },
        { id: '2c', name: 'Challenge',
          prompt: 'Have the arguments and judgements been subject to internal or external challenge?',
          low: 'No alternative hypotheses have been examined. The forecast is single-outcome.',
          high: 'Alternatives have been examined thoroughly and the judgements have been challenged.' }
      ],
      rollup: {
        name: 'Analysis overall',
        prompt: 'Taken together, how strong is the analysis? Set this yourself; it is not an average.',
        low: 'The analysis is very weak.',
        high: 'The analysis is very strong.'
      },
      frameSubject: 'The analysis supports ',
      frameHint: '[strength of argument, load-bearing assumptions and challenge applied]'
    },
    {
      id: 'c3', num: 3, name: 'Complexity and volatility',
      factors: [
        { id: '3a', name: 'Complexity',
          prompt: 'How many inputs, drivers and actors does the problem involve?',
          low: 'Highly complex, with many interacting parts.',
          high: 'Straightforward, with few drivers or actors.' },
        { id: '3b', name: 'Volatility',
          prompt: 'How quickly is the situation changing? Do the actors behave consistently, or are abrupt and unexpected shifts frequent?',
          low: 'The situation changes rapidly and with little warning.',
          high: 'The situation is very stable.' }
      ],
      rollup: {
        name: 'Complexity and volatility overall',
        prompt: 'Taken together, how far does the nature of the subject support confidence? A higher score means a more stable, simpler subject. Set this yourself; it is not an average.',
        low: 'A highly volatile and complex subject.',
        high: 'A stable, linear and straightforward subject.'
      },
      frameSubject: 'The complexity and volatility of the subject support ',
      frameHint: '[how complex and how fast-moving the subject is]'
    }
  ];

  var OVERALL = {
    id: 'overall',
    name: 'Overall analytical confidence',
    prompt: 'Taken together, what is your overall confidence in this key judgement? Set this yourself; it is not an average.',
    low: 'Low confidence.',
    high: 'High confidence.'
  };

  var HOW_TO = 'How to use this workbook. ' +
    '1. Enter each key judgement on its own sheet (KJ 1 to KJ ' + MAX_KJ + '). Leave unused sheets blank. ' +
    '2. For each factor, choose a score from the list (0 to 100 in steps of 10) between the two anchor statements, and record your reasoning in the notes. ' +
    '3. Set the overall score for each criterion, and the overall confidence, yourself. They are judgements, not averages. ' +
    'A score of 0 to 30 reads as Low, 40 to 60 as Moderate and 70 to 100 as High. ' +
    '4. Use the ingredients and the writing frame at the foot of each sheet to write the AnCR statement. ' +
    'Cells outlined in blue are for you to complete. The rest of each sheet is protected, without a password, to keep the layout intact (Review > Unprotect Sheet).';

  var KJ_HOW_TO = 'Choose each score from the list: 0 to 100 in steps of 10, where 0 matches the lower anchor and 100 the upper anchor. ' +
    'Overall scores are your judgement, not an average: 0 to 30 is Low, 40 to 60 is Moderate, 70 to 100 is High. Cells outlined in blue are for you to complete.';

  /* Every scored item, in sheet order. */
  var FACTORS = [];
  var KEYS = [];
  CRITERIA.forEach(function (c) {
    c.factors.forEach(function (f) { FACTORS.push(f); KEYS.push(f.id); });
    KEYS.push(c.id);
  });
  KEYS.push('overall');

  function band(v) {
    if (v === null || v === undefined || v === '') return '';
    return v <= 30 ? 'Low' : (v <= 60 ? 'Moderate' : 'High');
  }

  function normScore(v) {
    if (v === null || v === undefined || v === '') return null;
    var n = Number(v);
    if (!isFinite(n)) return null;
    n = Math.round(n / 10) * 10;
    return Math.max(0, Math.min(100, n));
  }

  function emptyKJ() {
    var kj = { statement: '', scores: {}, notes: {}, ancr: '' };
    KEYS.forEach(function (k) { kj.scores[k] = null; kj.notes[k] = ''; });
    return kj;
  }

  function emptyState() {
    return { meta: { title: '', analyst: '', date: '', marking: '' }, kjs: [emptyKJ()] };
  }

  function kjIsEmpty(kj) {
    if (kj.statement || kj.ancr) return false;
    for (var i = 0; i < KEYS.length; i++) {
      if (kj.scores[KEYS[i]] !== null || kj.notes[KEYS[i]]) return false;
    }
    return true;
  }

  function stateIsEmpty(s) {
    var m = s.meta;
    if (m.title || m.analyst || m.date || m.marking) return false;
    return s.kjs.every(kjIsEmpty);
  }

  /* Tidy anything loaded from storage or a file into a well-formed state. */
  function normState(raw) {
    var s = emptyState();
    if (!raw || typeof raw !== 'object') return s;
    var m = raw.meta || {};
    ['title', 'analyst', 'date', 'marking'].forEach(function (k) {
      if (typeof m[k] === 'string') s.meta[k] = m[k];
    });
    var kjs = Array.isArray(raw.kjs) ? raw.kjs.slice(0, MAX_KJ) : [];
    s.kjs = kjs.map(function (r) {
      var kj = emptyKJ();
      if (!r || typeof r !== 'object') return kj;
      if (typeof r.statement === 'string') kj.statement = r.statement;
      if (typeof r.ancr === 'string') kj.ancr = r.ancr;
      KEYS.forEach(function (k) {
        if (r.scores) kj.scores[k] = normScore(r.scores[k]);
        if (r.notes && typeof r.notes[k] === 'string') kj.notes[k] = r.notes[k];
      });
      return kj;
    });
    if (!s.kjs.length) s.kjs.push(emptyKJ());
    return s;
  }

  /* Factors sharing the highest (or lowest) score, e.g. "Coverage (80); Challenge (80)". */
  function extremes(kj, pickMax) {
    var vals = [];
    FACTORS.forEach(function (f) { if (kj.scores[f.id] !== null) vals.push(kj.scores[f.id]); });
    if (!vals.length) return '';
    var target = pickMax ? Math.max.apply(null, vals) : Math.min.apply(null, vals);
    var out = [];
    FACTORS.forEach(function (f) {
      if (kj.scores[f.id] === target) out.push(f.name + ' (' + target + ')');
    });
    return out.join('; ');
  }
  function strongest(kj) { return extremes(kj, true); }
  function weakest(kj) { return extremes(kj, false); }

  var FRAME_TAIL = 'Confidence would increase if [further collection, corroboration or challenge]. ' +
    'It would decrease if [the assumption or indicator that would undermine the judgement].';

  function frame(kj) {
    var o = band(kj.scores.overall);
    var t = 'We have ' + (o ? o.toUpperCase() : '[overall rating]') + ' confidence in this judgement. ';
    CRITERIA.forEach(function (c) {
      var b = band(kj.scores[c.id]);
      t += c.frameSubject + (b ? b.toUpperCase() : '[rating]') + ' confidence because ' + c.frameHint + '. ';
    });
    return t + FRAME_TAIL;
  }

  return {
    MAX_KJ: MAX_KJ, STEPS: STEPS, TITLE: TITLE, SUBTITLE: SUBTITLE,
    CRITERIA: CRITERIA, OVERALL: OVERALL, FACTORS: FACTORS, KEYS: KEYS,
    HOW_TO: HOW_TO, KJ_HOW_TO: KJ_HOW_TO, FRAME_TAIL: FRAME_TAIL,
    band: band, normScore: normScore, emptyKJ: emptyKJ, emptyState: emptyState,
    kjIsEmpty: kjIsEmpty, stateIsEmpty: stateIsEmpty, normState: normState,
    strongest: strongest, weakest: weakest, frame: frame
  };
})();
