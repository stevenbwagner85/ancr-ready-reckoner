/* Run by tests/run.py with model.js, zip.js and workbook.js loaded first.
 * `DEFLATED_HEX` holds a workbook re-compressed with zlib, as Excel would save it. */
function assert(c, m) { if (!c) throw new Error('FAIL: ' + m); }
function eq(a, b, m) { assert(JSON.stringify(a) === JSON.stringify(b), m + '\n  got  ' + JSON.stringify(a) + '\n  want ' + JSON.stringify(b)); }

var s = ANCR.emptyState();
s.meta = { title: 'Test assessment & "quotes" <tags>', analyst: 'Test team', date: '26 November 2026', marking: 'OFFICIAL' };
s.kjs = [ANCR.emptyKJ(), ANCR.emptyKJ(), ANCR.emptyKJ()];
s.kjs[0].statement = 'It is HIGHLY LIKELY that X will happen.\nSecond line with unicode: café – £100 😀';
s.kjs[0].scores = { '1a': 0, '1b': 50, '1c': 70, '1d': 40, c1: 30, '2a': 80, '2b': 80, '2c': 50, c2: 60, '3a': 10, '3b': 30, c3: 20, overall: 100 };
s.kjs[0].notes['1a'] = 'Gaps in coverage of <ports> & shipping';
s.kjs[0].notes.overall = 'Line one\nLine two';
s.kjs[0].ancr = 'We have HIGH confidence...';
s.kjs[2].statement = 'Third judgement';
s.kjs[2].scores['2c'] = 90;
s = ANCR.normState(s);

eq(ANCR.band(0), 'Low', 'band 0'); eq(ANCR.band(30), 'Low', 'band 30'); eq(ANCR.band(40), 'Moderate', 'band 40');
eq(ANCR.band(60), 'Moderate', 'band 60'); eq(ANCR.band(70), 'High', 'band 70'); eq(ANCR.band(null), '', 'band null');
eq(ANCR.strongest(s.kjs[0]), 'Argumentation (80); Assumptions (80)', 'strongest');
eq(ANCR.weakest(s.kjs[0]), 'Coverage (0)', 'weakest');
eq(ANCR.strongest(ANCR.emptyKJ()), '', 'strongest of nothing');

var bytes = ANCR.buildWorkbook(s);
eq(ANCR.readWorkbook(bytes), s, 'round trip (stored zip)');

var deflated = new Uint8Array(DEFLATED_HEX.length / 2);
for (var i = 0; i < deflated.length; i++) deflated[i] = parseInt(DEFLATED_HEX.substr(i * 2, 2), 16);
if (deflated.length) eq(ANCR.readWorkbook(deflated), s, 'round trip (deflated zip)');

var blank = ANCR.readWorkbook(ANCR.buildWorkbook(ANCR.emptyState()));
eq(blank, ANCR.emptyState(), 'blank template reads back empty');

var threw = false;
try { ANCR.readWorkbook(new Uint8Array([1, 2, 3, 4])); } catch (e) { threw = true; }
assert(threw, 'rejects a non-zip file');

if (typeof EMIT !== 'undefined' && EMIT) {
  var hex = [];
  for (i = 0; i < bytes.length; i++) hex.push((bytes[i] < 16 ? '0' : '') + bytes[i].toString(16));
  print(hex.join(''));
} else {
  print('ok');
}
