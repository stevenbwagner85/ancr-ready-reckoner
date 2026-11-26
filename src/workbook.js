/* AnCR ready reckoner: builds and reads the Excel workbook.
 * buildWorkbook(state) returns the .xlsx bytes; with an empty state it is the
 * blank template. readWorkbook(bytes) recovers the state from a workbook that
 * was produced here, including after it has been edited and re-saved in Excel.
 * The workbook uses no macros and no ActiveX: scores are list dropdowns and the
 * "slider" is drawn with conditional formatting. */
(function (A) {
  'use strict';

  /* ---------- layout ---------- */
  var COL = { ref: 1, factor: 2, low: 3, score: 4, track0: 5, high: 16, band: 17, notes: 18 };
  var LAST_COL = 18;
  var KJ_WIDTHS = [6, 34, 30, 8, 3.4, 3.4, 3.4, 3.4, 3.4, 3.4, 3.4, 3.4, 3.4, 3.4, 3.4, 30, 12, 50];
  var SUM_WIDTHS = [26, 58, 14, 14, 16, 16, 70];
  var ROW = {
    banner: 1, title: 2, statement: 3, howto: 4, head: 5,
    overallHead: 21, mirror0: 22, overall: 25,
    ingHead: 27, ingOverall: 28, ingStrong: 29, ingWeak: 30, ingFrame: 31, ancr: 32, bannerEnd: 34
  };
  /* key -> row on a KJ sheet */
  var KEY_ROW = {};
  var CRIT_HEAD_ROW = {};
  (function () {
    var r = 6;
    A.CRITERIA.forEach(function (c) {
      CRIT_HEAD_ROW[c.id] = r++;
      c.factors.forEach(function (f) { KEY_ROW[f.id] = r++; });
      KEY_ROW[c.id] = r++;
    });
    KEY_ROW.overall = ROW.overall;
  })();
  var SUM = { banner: 1, title: 2, sub: 3, meta0: 4, howto: 9, head: 11, first: 12, bannerEnd: 12 + A.MAX_KJ + 1 };
  var META_KEYS = ['title', 'analyst', 'date', 'marking'];
  var META_LABELS = ['Assessment title', 'Analyst or team', 'Date', 'Protective marking (optional)'];

  /* ---------- styles (indexes into cellXfs) ---------- */
  var S = {
    base: 0, banner: 1, title: 2, sub: 3, label: 4, input: 5, head: 6, headScale: 7, crit: 8,
    ref: 9, text: 10, score: 11, track: 12, trackZone: 13, band: 14, rollText: 15, rollRef: 16,
    mirror: 17, frame: 18, sumCell: 19, kjTag: 20, howto: 21, inputDate: 22, rollBlank: 23, sumText: 24
  };
  var NAVY = 'FF002A53', WHITE = 'FFFFFFFF', PALE = 'FFE3EAF6', SKY = 'FF51C1E1', LINK = 'FF0089CF',
      GREY = 'FFD9D9D9', TRACK = 'FFEDEFF2', CHARCOAL = 'FF4A4A49';

  function stylesXml() {
    function font(sz, color, bold) {
      return '<font>' + (bold ? '<b/>' : '') + '<sz val="' + sz + '"/><color rgb="' + color + '"/><name val="Arial"/><family val="2"/></font>';
    }
    function solid(rgb) {
      return '<fill><patternFill patternType="solid"><fgColor rgb="' + rgb + '"/><bgColor indexed="64"/></patternFill></fill>';
    }
    function side(tag, style, rgb) {
      return style ? '<' + tag + ' style="' + style + '"><color rgb="' + rgb + '"/></' + tag + '>' : '<' + tag + '/>';
    }
    function border(l, r, t, b) {
      return '<border>' + side('left', l[0], l[1]) + side('right', r[0], r[1]) + side('top', t[0], t[1]) + side('bottom', b[0], b[1]) + '<diagonal/></border>';
    }
    var g = ['thin', GREY], bl = ['thin', LINK], w = ['thin', WHITE], nv = ['medium', NAVY], none = [null, null];
    var fonts = [
      font(10, NAVY, false),      // 0
      font(10, NAVY, true),       // 1
      font(16, NAVY, true),       // 2
      font(10, WHITE, true),      // 3
      font(7, WHITE, false),      // 4
      font(11, NAVY, true),       // 5
      font(12, NAVY, true),       // 6
      font(9, CHARCOAL, false),   // 7
      font(11, WHITE, true)       // 8
    ];
    var fills = [
      '<fill><patternFill patternType="none"/></fill>',
      '<fill><patternFill patternType="gray125"/></fill>',
      solid(NAVY),   // 2
      solid(PALE),   // 3
      solid(TRACK),  // 4
      solid(WHITE)   // 5
    ];
    var borders = [
      border(none, none, none, none), // 0
      border(g, g, g, g),             // 1
      border(bl, bl, bl, bl),         // 2 input
      border(w, w, w, w),             // 3 track
      border(nv, w, w, w)             // 4 track, start of a zone
    ];
    function xf(fontId, fillId, borderId, h, v, wrap, unlocked, numFmt) {
      var a = '<xf numFmtId="' + (numFmt || 0) + '" fontId="' + fontId + '" fillId="' + fillId + '" borderId="' + borderId + '" xfId="0"' +
        (numFmt ? ' applyNumberFormat="1"' : '') + ' applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"' + (unlocked ? ' applyProtection="1"' : '') + '>';
      a += '<alignment' + (h ? ' horizontal="' + h + '"' : '') + ' vertical="' + (v || 'top') + '"' + (wrap ? ' wrapText="1"' : '') + '/>';
      if (unlocked) a += '<protection locked="0"/>';
      return a + '</xf>';
    }
    var xfs = [
      xf(0, 0, 0, null, 'top', false),            // 0 base
      xf(5, 0, 0, 'center', 'center', false),     // 1 banner
      xf(2, 0, 0, null, 'center', false),         // 2 title
      xf(7, 0, 0, null, 'top', true),             // 3 sub
      xf(1, 0, 0, null, 'top', true),             // 4 label
      xf(0, 5, 2, null, 'top', true, true),       // 5 input
      xf(3, 2, 1, 'center', 'center', true),      // 6 head
      xf(4, 2, 0, 'center', 'bottom', false),     // 7 headScale
      xf(8, 2, 0, null, 'center', false),         // 8 criterion heading
      xf(1, 0, 1, 'center', 'top', false),        // 9 ref
      xf(0, 0, 1, null, 'top', true),             // 10 text
      xf(6, 5, 2, 'center', 'center', false, true), // 11 score
      xf(0, 4, 3, null, 'center', false),         // 12 track
      xf(0, 4, 4, null, 'center', false),         // 13 track zone start
      xf(1, 0, 1, 'center', 'center', false),     // 14 band
      xf(0, 3, 1, null, 'top', true),             // 15 roll-up text
      xf(1, 3, 1, 'center', 'top', false),        // 16 roll-up ref
      xf(6, 0, 1, 'center', 'center', false),     // 17 mirrored score
      xf(0, 3, 1, null, 'top', true),             // 18 frame
      xf(0, 0, 1, 'center', 'top', true),         // 19 summary cell
      xf(5, 0, 0, 'right', 'center', false),      // 20 KJ tag
      xf(0, 3, 0, null, 'top', true),             // 21 how-to panel
      xf(0, 5, 2, null, 'top', true, true, 49),   // 22 date input (text)
      xf(0, 3, 1, null, 'top', false),            // 23 roll-up blank
      xf(0, 0, 1, null, 'top', true)              // 24 summary text
    ];
    function dxf(fillRgb, fontRgb, bold) {
      return '<dxf>' + (fontRgb ? '<font>' + (bold ? '<b/>' : '') + '<color rgb="' + fontRgb + '"/></font>' : '') +
        '<fill><patternFill><bgColor rgb="' + fillRgb + '"/></patternFill></fill></dxf>';
    }
    var dxfs = [
      dxf(NAVY, null),           // 0 slider position
      dxf(SKY, null),            // 1 filled part of the track
      dxf(NAVY, WHITE, true),    // 2 High
      dxf(SKY, NAVY, true),      // 3 Moderate
      dxf(PALE, NAVY, true)      // 4 Low
    ];
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      '<fonts count="' + fonts.length + '">' + fonts.join('') + '</fonts>' +
      '<fills count="' + fills.length + '">' + fills.join('') + '</fills>' +
      '<borders count="' + borders.length + '">' + borders.join('') + '</borders>' +
      '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
      '<cellXfs count="' + xfs.length + '">' + xfs.join('') + '</cellXfs>' +
      '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
      '<dxfs count="' + dxfs.length + '">' + dxfs.join('') + '</dxfs>' +
      '<tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>' +
      '</styleSheet>';
  }

  /* ---------- XML helpers ---------- */
  function esc(s) {
    return String(s)
      .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]/g, '')
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }
  function unesc(s) {
    return s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'")
      .replace(/&#x([0-9a-fA-F]+);/g, function (m, h) { return String.fromCharCode(parseInt(h, 16)); })
      .replace(/&#(\d+);/g, function (m, d) { return String.fromCharCode(parseInt(d, 10)); })
      .replace(/&amp;/g, '&');
  }
  function colName(n) {
    var s = '';
    while (n > 0) { s = String.fromCharCode(65 + (n - 1) % 26) + s; n = Math.floor((n - 1) / 26); }
    return s;
  }
  function ref(c, r) { return colName(c) + r; }
  function clip(s) { s = String(s == null ? '' : s).replace(/\r\n?/g, '\n'); return s.length > 32000 ? s.slice(0, 32000) : s; }

  /* A sheet under construction: cells keyed by row then column. */
  function Sheet() { this.rows = {}; this.heights = {}; this.merges = []; }
  Sheet.prototype.set = function (c, r, style, value) {
    (this.rows[r] = this.rows[r] || {})[c] = { s: style, v: value };
  };
  Sheet.prototype.fill = function (c1, c2, r, style) {
    for (var c = c1; c <= c2; c++) if (!(this.rows[r] && this.rows[r][c])) this.set(c, r, style, null);
  };
  Sheet.prototype.merge = function (c1, r1, c2, r2, style, value) {
    for (var r = r1; r <= r2; r++) for (var c = c1; c <= c2; c++) this.set(c, r, style, null);
    this.set(c1, r1, style, value);
    this.merges.push(ref(c1, r1) + ':' + ref(c2, r2));
  };
  Sheet.prototype.dataXml = function () {
    var self = this, out = ['<sheetData>'];
    Object.keys(this.rows).map(Number).sort(function (a, b) { return a - b; }).forEach(function (r) {
      var h = self.heights[r];
      out.push('<row r="' + r + '"' + (h ? ' ht="' + h + '" customHeight="1"' : '') + '>');
      var row = self.rows[r];
      Object.keys(row).map(Number).sort(function (a, b) { return a - b; }).forEach(function (c) {
        out.push(cellXml(ref(c, r), row[c].s, row[c].v));
      });
      out.push('</row>');
    });
    out.push('</sheetData>');
    return out.join('');
  };

  var RUN_BOLD = '<rPr><b/><sz val="10"/><color rgb="' + NAVY + '"/><rFont val="Arial"/><family val="2"/></rPr>';
  var RUN_PLAIN = '<rPr><sz val="10"/><color rgb="' + NAVY + '"/><rFont val="Arial"/><family val="2"/></rPr>';

  function cellXml(r, s, v) {
    var a = '<c r="' + r + '" s="' + s + '"';
    if (v === null || v === undefined || v === '') return a + '/>';
    if (typeof v === 'number') return a + '><v>' + v + '</v></c>';
    if (typeof v === 'string') return a + ' t="inlineStr"><is><t xml:space="preserve">' + esc(v) + '</t></is></c>';
    if (v.rich) {
      return a + ' t="inlineStr"><is>' + v.rich.map(function (run) {
        return '<r>' + (run[0] ? RUN_BOLD : RUN_PLAIN) + '<t xml:space="preserve">' + esc(run[1]) + '</t></r>';
      }).join('') + '</is></c>';
    }
    if (v.f) {
      if (typeof v.v === 'number') return a + '><f>' + esc(v.f) + '</f><v>' + v.v + '</v></c>';
      return a + ' t="str"><f>' + esc(v.f) + '</f><v>' + esc(v.v == null ? '' : v.v) + '</v></c>';
    }
    return a + '/>';
  }

  /* Rough row height (points) for wrapped Arial 10 text in a column of the given width. */
  function lines(text, width) {
    if (!text) return 1;
    var per = Math.max(4, Math.floor(width * 1.08)), n = 0;
    String(text).split('\n').forEach(function (p) { n += Math.max(1, Math.ceil(p.length / per)); });
    return n;
  }
  function height(nLines, min) { return Math.min(409, Math.max(min || 15, Math.round(nLines * 12.9 + 5))); }
  function spanWidth(widths, c1, c2) { var w = 0; for (var c = c1; c <= c2; c++) w += widths[c - 1]; return w; }

  /* ---------- formulas ---------- */
  var STEP_LIST = '"' + A.STEPS.join(',') + '"';
  function bandF(cell) {
    return 'IF(' + cell + '="","",IF(' + cell + '<=30,"Low",IF(' + cell + '<=60,"Moderate","High")))';
  }
  var FACTOR_RANGE = (function () {
    return A.CRITERIA.map(function (c) {
      return ref(COL.score, KEY_ROW[c.factors[0].id]) + ':' + ref(COL.score, KEY_ROW[c.factors[c.factors.length - 1].id]);
    }).join(',');
  })();
  function extremesF(fn) {
    var parts = A.FACTORS.map(function (f) {
      var cell = ref(COL.score, KEY_ROW[f.id]);
      return 'IF(AND(' + cell + '<>"",' + cell + '=' + fn + '(' + FACTOR_RANGE + ')),"' + f.name + ' ("&' + cell + '&"); ","")';
    }).join('&');
    return 'IF(COUNT(' + FACTOR_RANGE + ')=0,"",LEFT(' + parts + ',LEN(' + parts + ')-2))';
  }
  function frameF() {
    var bandCell = function (key) { return ref(COL.band, KEY_ROW[key]); };
    var f = '"We have "&IF(' + bandCell('overall') + '="","[overall rating]",UPPER(' + bandCell('overall') + '))&" confidence in this judgement. "';
    A.CRITERIA.forEach(function (c) {
      f += '&"' + c.frameSubject + '"&IF(' + bandCell(c.id) + '="","[rating]",UPPER(' + bandCell(c.id) + '))&" confidence because ' + c.frameHint + '. "';
    });
    return f + '&"' + A.FRAME_TAIL + '"';
  }

  /* ---------- sheets ---------- */
  function sheetXml(opts) {
    var x = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>' +
      '<sheetViews><sheetView showGridLines="0"' + (opts.active ? ' tabSelected="1"' : '') + ' zoomScale="90" zoomScaleNormal="90" workbookViewId="0">' +
      (opts.freezeRow ? '<pane ySplit="' + opts.freezeRow + '" topLeftCell="A' + (opts.freezeRow + 1) + '" activePane="bottomLeft" state="frozen"/>' +
        '<selection pane="bottomLeft" activeCell="' + opts.activeCell + '" sqref="' + opts.activeCell + '"/>'
        : '<selection activeCell="' + opts.activeCell + '" sqref="' + opts.activeCell + '"/>') +
      '</sheetView></sheetViews>' +
      '<sheetFormatPr defaultRowHeight="15"/>' +
      '<cols>' + opts.widths.map(function (w, i) {
        return '<col min="' + (i + 1) + '" max="' + (i + 1) + '" width="' + w + '" customWidth="1"/>';
      }).join('') + '</cols>' +
      opts.sheet.dataXml() +
      '<sheetProtection sheet="1" objects="1" scenarios="1" formatColumns="0" formatRows="0"/>' +
      '<mergeCells count="' + opts.sheet.merges.length + '">' + opts.sheet.merges.map(function (m) { return '<mergeCell ref="' + m + '"/>'; }).join('') + '</mergeCells>' +
      (opts.cf || '') + (opts.dv || '') +
      '<printOptions horizontalCentered="1"/>' +
      '<pageMargins left="0.4" right="0.4" top="0.5" bottom="0.6" header="0.3" footer="0.3"/>' +
      '<pageSetup paperSize="9" orientation="landscape" fitToHeight="0"/>' +
      '<headerFooter><oddFooter>&amp;L&amp;"Arial,Regular"&amp;8&amp;A&amp;R&amp;"Arial,Regular"&amp;8Page &amp;P of &amp;N</oddFooter></headerFooter>' +
      (opts.drawing ? '<drawing r:id="rId1"/>' : '') +
      '</worksheet>';
    return x;
  }

  function bandCf(sqref, p) {
    function rule(text, dxf, pr) {
      return '<cfRule type="cellIs" dxfId="' + dxf + '" priority="' + pr + '" operator="equal"><formula>"' + text + '"</formula></cfRule>';
    }
    return '<conditionalFormatting sqref="' + sqref + '">' + rule('High', 2, p) + rule('Moderate', 3, p + 1) + rule('Low', 4, p + 2) + '</conditionalFormatting>';
  }

  function kjSheet(n, kj, isActive) {
    var sh = new Sheet(), W = KJ_WIDTHS, r, c;
    var bannerF = 'IF(Summary!$B$7="","",Summary!$B$7)';

    sh.merge(1, ROW.banner, LAST_COL, ROW.banner, S.banner, { f: bannerF, v: kj.__marking });
    sh.heights[ROW.banner] = 20;
    sh.merge(1, ROW.title, COL.high, ROW.title, S.title, A.TITLE);
    sh.merge(COL.band, ROW.title, LAST_COL, ROW.title, S.kjTag, 'Key judgement ' + n);
    sh.heights[ROW.title] = 28;
    sh.merge(1, ROW.statement, 2, ROW.statement, S.label, 'Key judgement ' + n);
    sh.merge(3, ROW.statement, LAST_COL, ROW.statement, S.input, clip(kj.statement));
    sh.heights[ROW.statement] = height(lines(kj.statement, spanWidth(W, 3, LAST_COL)), 44);
    sh.merge(1, ROW.howto, LAST_COL, ROW.howto, S.sub, A.KJ_HOW_TO);
    sh.heights[ROW.howto] = 28;

    // column headings
    sh.set(COL.ref, ROW.head, S.head, 'Ref');
    sh.set(COL.factor, ROW.head, S.head, 'Factor');
    sh.set(COL.low, ROW.head, S.head, 'Lower anchor (0)');
    sh.set(COL.score, ROW.head, S.head, 'Score');
    A.STEPS.forEach(function (v, i) { sh.set(COL.track0 + i, ROW.head, S.headScale, v); });
    sh.set(COL.high, ROW.head, S.head, 'Upper anchor (100)');
    sh.set(COL.band, ROW.head, S.head, 'Rating');
    sh.set(COL.notes, ROW.head, S.head, 'Analyst’s notes');
    sh.heights[ROW.head] = 24;

    function track(row, zoned) {
      for (var i = 0; i < 11; i++) sh.set(COL.track0 + i, row, (zoned && (i === 4 || i === 7)) ? S.trackZone : S.track, null);
    }
    function scoreRow(row, key, item, label, isRoll) {
      var note = clip(kj.notes[key]);
      sh.set(COL.ref, row, isRoll ? S.rollRef : S.ref, label);
      sh.set(COL.factor, row, isRoll ? S.rollText : S.text, { rich: [[true, item.name + '\n'], [false, item.prompt]] });
      sh.set(COL.low, row, isRoll ? S.rollText : S.text, item.low);
      sh.set(COL.score, row, S.score, kj.scores[key]);
      track(row, isRoll);
      sh.set(COL.high, row, isRoll ? S.rollText : S.text, item.high);
      if (isRoll) sh.set(COL.band, row, S.band, { f: bandF(ref(COL.score, row)), v: A.band(kj.scores[key]) });
      else sh.set(COL.band, row, S.text, null);
      sh.set(COL.notes, row, S.input, note);
      var n1 = Math.max(lines(item.name, W[1]) + lines(item.prompt, W[1]), lines(item.low, W[2]), lines(item.high, W[15]), lines(note, W[17]));
      sh.heights[row] = height(n1, 44);
    }

    A.CRITERIA.forEach(function (crit) {
      r = CRIT_HEAD_ROW[crit.id];
      sh.merge(1, r, LAST_COL, r, S.crit, 'Criterion ' + crit.num + ': ' + crit.name);
      sh.heights[r] = 22;
      crit.factors.forEach(function (f) { scoreRow(KEY_ROW[f.id], f.id, f, f.id, false); });
      scoreRow(KEY_ROW[crit.id], crit.id, crit.rollup, String(crit.num), true);
    });

    // overall block
    sh.merge(1, ROW.overallHead, LAST_COL, ROW.overallHead, S.crit, 'Overall analytical confidence');
    sh.heights[ROW.overallHead] = 22;
    A.CRITERIA.forEach(function (crit, i) {
      r = ROW.mirror0 + i;
      var src = ref(COL.score, KEY_ROW[crit.id]);
      sh.set(COL.ref, r, S.ref, String(crit.num));
      sh.merge(COL.factor, r, COL.low, r, S.text, crit.name + ' (carried down from above)');
      sh.set(COL.score, r, S.mirror, { f: 'IF(' + src + '="","",' + src + ')', v: kj.scores[crit.id] === null ? '' : kj.scores[crit.id] });
      track(r, true);
      sh.set(COL.high, r, S.text, null);
      sh.set(COL.band, r, S.band, { f: bandF(ref(COL.score, r)), v: A.band(kj.scores[crit.id]) });
      sh.set(COL.notes, r, S.text, null);
      sh.heights[r] = 22;
    });
    scoreRow(ROW.overall, 'overall', A.OVERALL, '', true);

    // ingredients
    sh.merge(1, ROW.ingHead, LAST_COL, ROW.ingHead, S.crit, 'Ingredients for the AnCR statement');
    sh.heights[ROW.ingHead] = 22;
    function ing(row, label, style, value, minH, text) {
      sh.merge(1, row, 2, row, S.label, label);
      sh.merge(3, row, LAST_COL, row, style, value);
      sh.heights[row] = height(lines(text, spanWidth(W, 3, LAST_COL)), minH);
    }
    var ob = A.band(kj.scores.overall);
    ing(ROW.ingOverall, 'Overall confidence', S.frame, { f: 'IF(' + ref(COL.band, ROW.overall) + '="","",' + ref(COL.band, ROW.overall) + ')', v: ob }, 18, ob);
    ing(ROW.ingStrong, 'Strongest factors', S.frame, { f: extremesF('MAX'), v: A.strongest(kj) }, 18, '');
    ing(ROW.ingWeak, 'Weakest factors', S.frame, { f: extremesF('MIN'), v: A.weakest(kj) }, 18, '');
    ing(ROW.ingFrame, 'Writing frame', S.frame, { f: frameF(), v: A.frame(kj) }, 58, A.frame(kj));
    ing(ROW.ancr, 'Your AnCR statement', S.input, clip(kj.ancr), 90, kj.ancr);
    sh.merge(1, ROW.bannerEnd, LAST_COL, ROW.bannerEnd, S.banner, { f: bannerF, v: kj.__marking });
    sh.heights[ROW.bannerEnd] = 20;

    // conditional formatting: slider track and rating colours
    var t0 = colName(COL.track0), t1 = colName(COL.track0 + 10), sc = '$' + colName(COL.score);
    var first = KEY_ROW['1a'];
    var cf = '<conditionalFormatting sqref="' + t0 + first + ':' + t1 + ROW.overall + '">' +
      '<cfRule type="expression" dxfId="0" priority="1" stopIfTrue="1"><formula>' + esc('AND(ISNUMBER(' + sc + first + '),(COLUMN()-' + COL.track0 + ')*10=' + sc + first + ')') + '</formula></cfRule>' +
      '<cfRule type="expression" dxfId="1" priority="2"><formula>' + esc('AND(ISNUMBER(' + sc + first + '),(COLUMN()-' + COL.track0 + ')*10<' + sc + first + ')') + '</formula></cfRule>' +
      '</conditionalFormatting>' +
      bandCf(colName(COL.band) + first + ':' + colName(COL.band) + ROW.overall, 3);

    var scoreCells = A.KEYS.map(function (k) { return ref(COL.score, KEY_ROW[k]); }).join(' ');
    var dv = '<dataValidations count="1"><dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" ' +
      'errorTitle="Score" error="Choose a score from 0 to 100 in steps of 10, or leave the cell blank." ' +
      'promptTitle="Score" prompt="0 matches the lower anchor, 100 the upper anchor. Steps of 10." sqref="' + scoreCells + '">' +
      '<formula1>' + esc(STEP_LIST) + '</formula1></dataValidation></dataValidations>';

    return sheetXml({ sheet: sh, widths: W, freezeRow: ROW.head, activeCell: 'D7', active: isActive, cf: cf, dv: dv });
  }

  function summarySheet(state) {
    var sh = new Sheet(), W = SUM_WIDTHS, m = state.meta, i, r;
    var bannerF = 'IF($B$7="","",$B$7)';
    sh.merge(1, SUM.banner, 7, SUM.banner, S.banner, { f: bannerF, v: m.marking });
    sh.heights[SUM.banner] = 20;
    sh.set(1, SUM.title, S.base, null);
    sh.merge(2, SUM.title, 7, SUM.title, S.title, A.TITLE);
    sh.heights[SUM.title] = 66;
    sh.set(1, SUM.sub, S.base, null);
    sh.merge(2, SUM.sub, 7, SUM.sub, S.sub, A.SUBTITLE);
    sh.heights[SUM.sub] = 18;
    META_KEYS.forEach(function (k, j) {
      r = SUM.meta0 + j;
      sh.set(1, r, S.label, META_LABELS[j]);
      sh.merge(2, r, 7, r, k === 'date' ? S.inputDate : S.input, clip(m[k]));
      sh.heights[r] = height(lines(m[k], spanWidth(W, 2, 7)), 18);
    });
    sh.merge(1, SUM.howto, 7, SUM.howto, S.howto, A.HOW_TO);
    sh.heights[SUM.howto] = height(lines(A.HOW_TO, spanWidth(W, 1, 7)), 44);

    var heads = ['KJ', 'Key judgement', 'Source base', 'Analysis', 'Complexity and volatility', 'Overall confidence', 'AnCR statement'];
    heads.forEach(function (h, j) { sh.set(j + 1, SUM.head, S.head, h); });
    sh.heights[SUM.head] = 30;

    for (i = 1; i <= A.MAX_KJ; i++) {
      r = SUM.first + i - 1;
      var kj = state.kjs[i - 1] || A.emptyKJ();
      var q = "'KJ " + i + "'!", has = !!kj.statement;
      var st = q + '$' + colName(3) + '$' + ROW.statement, an = q + '$' + colName(3) + '$' + ROW.ancr;
      sh.set(1, r, S.ref, 'KJ ' + i);
      sh.set(2, r, S.sumText, { f: 'IF(' + st + '="","",' + st + ')', v: clip(kj.statement) });
      ['c1', 'c2', 'c3', 'overall'].forEach(function (k, j) {
        var cell = q + '$' + colName(COL.band) + '$' + KEY_ROW[k];
        sh.set(3 + j, r, S.sumCell, { f: 'IF($B' + r + '="","",' + cell + ')', v: has ? A.band(kj.scores[k]) : '' });
      });
      sh.set(7, r, S.sumText, { f: 'IF($B' + r + '="","",IF(' + an + '="","",' + an + '))', v: has ? clip(kj.ancr) : '' });
      sh.heights[r] = height(Math.max(lines(kj.statement, W[1]), has ? lines(kj.ancr, W[6]) : 1), 30);
    }
    sh.merge(1, SUM.bannerEnd, 7, SUM.bannerEnd, S.banner, { f: bannerF, v: m.marking });
    sh.heights[SUM.bannerEnd] = 20;

    var cf = bandCf('C' + SUM.first + ':F' + (SUM.first + A.MAX_KJ - 1), 1);
    return sheetXml({ sheet: sh, widths: W, freezeRow: 0, activeCell: 'B4', active: true, cf: cf, drawing: true });
  }

  /* ---------- package ---------- */
  var NS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  var XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

  function buildWorkbook(rawState) {
    var state = A.normState(rawState), files = [], i;
    function add(name, text) { files.push({ name: name, data: typeof text === 'string' ? A.utf8Encode(text) : text }); }
    var nSheets = A.MAX_KJ + 1;

    var ct = XML_HEAD + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Default Extension="png" ContentType="image/png"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
      '<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>' +
      '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
      '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>';
    for (i = 1; i <= nSheets; i++) ct += '<Override PartName="/xl/worksheets/sheet' + i + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>';
    add('[Content_Types].xml', ct + '</Types>');

    add('_rels/.rels', XML_HEAD + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="' + NS_REL + '/officeDocument" Target="xl/workbook.xml"/>' +
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' +
      '<Relationship Id="rId3" Type="' + NS_REL + '/extended-properties" Target="docProps/app.xml"/>' +
      '</Relationships>');

    add('docProps/core.xml', XML_HEAD + '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ' +
      'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
      '<dc:title>' + esc(state.meta.title || A.TITLE) + '</dc:title><dc:creator>' + esc(state.meta.analyst) + '</dc:creator>' +
      '</cp:coreProperties>');
    add('docProps/app.xml', XML_HEAD + '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">' +
      '<Application>AnCR ready reckoner</Application></Properties>');

    var sheets = '<sheet name="Summary" sheetId="1" r:id="rId1"/>', names = '<definedName name="_xlnm.Print_Titles" localSheetId="0">Summary!$1:$1</definedName>';
    var rels = '<Relationship Id="rId1" Type="' + NS_REL + '/worksheet" Target="worksheets/sheet1.xml"/>';
    for (i = 1; i <= A.MAX_KJ; i++) {
      sheets += '<sheet name="KJ ' + i + '" sheetId="' + (i + 1) + '" r:id="rId' + (i + 1) + '"/>';
      names += '<definedName name="_xlnm.Print_Titles" localSheetId="' + i + '">\'KJ ' + i + '\'!$1:$1</definedName>';
      rels += '<Relationship Id="rId' + (i + 1) + '" Type="' + NS_REL + '/worksheet" Target="worksheets/sheet' + (i + 1) + '.xml"/>';
    }
    rels += '<Relationship Id="rId' + (nSheets + 1) + '" Type="' + NS_REL + '/styles" Target="styles.xml"/>';
    add('xl/workbook.xml', XML_HEAD + '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="' + NS_REL + '">' +
      '<workbookPr/><bookViews><workbookView xWindow="0" yWindow="0" windowWidth="28800" windowHeight="15600" tabRatio="800"/></bookViews>' +
      '<sheets>' + sheets + '</sheets><definedNames>' + names + '</definedNames><calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>');
    add('xl/_rels/workbook.xml.rels', XML_HEAD + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + rels + '</Relationships>');
    add('xl/styles.xml', stylesXml());

    add('xl/worksheets/sheet1.xml', summarySheet(state));
    add('xl/worksheets/_rels/sheet1.xml.rels', XML_HEAD + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="' + NS_REL + '/drawing" Target="../drawings/drawing1.xml"/></Relationships>');
    add('xl/drawings/drawing1.xml', XML_HEAD + '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" ' +
      'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="' + NS_REL + '">' +
      '<xdr:oneCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>90000</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>60000</xdr:rowOff></xdr:from>' +
      '<xdr:ext cx="720000" cy="720000"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="2" name="Brunel logo" descr="Brunel University of London logo"/>' +
      '<xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>' +
      '<xdr:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>' +
      '<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="720000" cy="720000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>' +
      '</xdr:pic><xdr:clientData/></xdr:oneCellAnchor></xdr:wsDr>');
    add('xl/drawings/_rels/drawing1.xml.rels', XML_HEAD + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="' + NS_REL + '/image" Target="../media/image1.png"/></Relationships>');
    add('xl/media/image1.png', A.base64Decode(A.ASSETS.logoNavy));

    for (i = 1; i <= A.MAX_KJ; i++) {
      var kj = state.kjs[i - 1] || A.emptyKJ();
      kj.__marking = state.meta.marking;
      add('xl/worksheets/sheet' + (i + 1) + '.xml', kjSheet(i, kj, false));
      delete kj.__marking;
    }
    return A.zipWrite(files);
  }

  /* ---------- reading ---------- */
  function sharedStrings(xml) {
    var out = [], re = /<si\b[^>]*?(?:\/>|>([\s\S]*?)<\/si>)/g, m;
    while ((m = re.exec(xml))) out.push(textRuns(m[1] || ''));
    return out;
  }
  function textRuns(xml) {
    // concatenate <t> runs, ignoring phonetic text
    xml = xml.replace(/<rPh\b[\s\S]*?<\/rPh>/g, '');
    var s = '', re = /<t\b[^>]*?(?:\/>|>([\s\S]*?)<\/t>)/g, m;
    while ((m = re.exec(xml))) s += unesc(m[1] || '');
    return s.replace(/_x000D_/g, '').replace(/\r\n?/g, '\n');
  }
  function sheetCells(xml, strings) {
    var cells = {}, re = /<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g, m;
    while ((m = re.exec(xml))) {
      var attrs = m[1], body = m[2] || '';
      var r = /\br="([A-Z]+[0-9]+)"/.exec(attrs), t = /\bt="([^"]*)"/.exec(attrs);
      if (!r) continue;
      var type = t ? t[1] : 'n', v = /<v\b[^>]*>([\s\S]*?)<\/v>/.exec(body), val = '';
      if (type === 'inlineStr') val = textRuns(body);
      else if (v) {
        if (type === 's') val = strings[parseInt(v[1], 10)] || '';
        else if (type === 'n') val = Number(v[1]);
        else if (type === 'b') val = v[1] === '1' ? 'TRUE' : 'FALSE';
        else val = unesc(v[1]).replace(/\r\n?/g, '\n');
      }
      cells[r[1]] = val;
    }
    return cells;
  }
  function asText(v) { return v === undefined || v === null ? '' : String(v); }

  function readWorkbook(bytes) {
    var files = A.zipRead(bytes);
    function text(name) { return files[name] ? A.utf8Decode(files[name]) : ''; }
    var wb = text('xl/workbook.xml'), rels = text('xl/_rels/workbook.xml.rels');
    if (!wb) throw new Error('This is not an Excel workbook (.xlsx).');
    var target = {}, re = /<Relationship\b([^>]*)\/?>/g, m, id, tg;
    while ((m = re.exec(rels))) {
      id = /\bId="([^"]*)"/.exec(m[1]); tg = /\bTarget="([^"]*)"/.exec(m[1]);
      if (id && tg) target[id[1]] = tg[1].replace(/^\/?(xl\/)?/, 'xl/');
    }
    var path = {};
    re = /<sheet\b([^>]*)\/?>/g;
    while ((m = re.exec(wb))) {
      var nm = /\bname="([^"]*)"/.exec(m[1]), rid = /\br:id="([^"]*)"/.exec(m[1]);
      if (nm && rid && target[rid[1]]) path[unesc(nm[1])] = target[rid[1]];
    }
    if (!path.Summary || !path['KJ 1']) {
      throw new Error('This workbook was not made with the AnCR ready reckoner: it has no "Summary" and "KJ 1" sheets.');
    }
    var strings = sharedStrings(text('xl/sharedStrings.xml'));
    var state = A.emptyState(), cells = sheetCells(text(path.Summary), strings);
    META_KEYS.forEach(function (k, j) { state.meta[k] = asText(cells[ref(2, SUM.meta0 + j)]); });
    state.kjs = [];
    var lastUsed = 0;
    for (var i = 1; i <= A.MAX_KJ; i++) {
      var kj = A.emptyKJ(), p = path['KJ ' + i];
      if (p) {
        cells = sheetCells(text(p), strings);
        kj.statement = asText(cells[ref(3, ROW.statement)]);
        kj.ancr = asText(cells[ref(3, ROW.ancr)]);
        A.KEYS.forEach(function (k) {
          var v = cells[ref(COL.score, KEY_ROW[k])];
          kj.scores[k] = (typeof v === 'number' || (typeof v === 'string' && /^\s*\d+\s*$/.test(v))) ? A.normScore(v) : null;
          kj.notes[k] = asText(cells[ref(COL.notes, KEY_ROW[k])]);
        });
      }
      state.kjs.push(kj);
      if (!A.kjIsEmpty(kj)) lastUsed = i;
    }
    state.kjs = state.kjs.slice(0, Math.max(1, lastUsed));
    return state;
  }

  A.buildWorkbook = buildWorkbook;
  A.readWorkbook = readWorkbook;
  A.LAYOUT = { COL: COL, ROW: ROW, KEY_ROW: KEY_ROW, SUM: SUM };
})(ANCR);
