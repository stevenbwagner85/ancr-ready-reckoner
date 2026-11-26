#!/usr/bin/env python3
"""Tests for the workbook generator and reader. Usage: python3 tests/run.py

1. Builds a filled-in workbook, checks every XML part is well formed.
2. Re-compresses it with zlib (as Excel does on save) and checks the JavaScript
   reader, including its own inflate, recovers exactly the same state.
3. Does the same with tests/fixtures/saved-by-excel.xlsx, the test workbook
   after it was opened and saved in desktop Excel.
"""
import base64, io, os, shutil, subprocess, sys, zipfile
import xml.dom.minidom

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
JSC = '/System/Library/Frameworks/JavaScriptCore.framework/Versions/Current/Helpers/jsc'


def read(*p):
    with open(os.path.join(ROOT, *p), encoding='utf-8') as f:
        return f.read()


def run(source):
    tmp = os.path.join(ROOT, '.test-tmp.js')
    node = shutil.which('node')
    with open(tmp, 'w', encoding='utf-8') as f:
        f.write(('var print = console.log;\n' if node else '') + source)
    try:
        r = subprocess.run([node or JSC, tmp], capture_output=True, text=True)
    finally:
        os.remove(tmp)
    if r.returncode != 0 or 'FAIL' in r.stdout + r.stderr or 'Exception' in r.stdout + r.stderr:
        sys.exit('JavaScript test failed:\n' + r.stdout[-3000:] + r.stderr[-3000:])
    return r.stdout.strip()


def main():
    with open(os.path.join(ROOT, 'assets', 'brunel-logo-navy-block.png'), 'rb') as f:
        assets = 'ANCR.ASSETS = {logoNavy: "%s"};\n' % base64.b64encode(f.read()).decode()
    core = '\n'.join([read('src', 'model.js'), assets, read('src', 'zip.js'), read('src', 'workbook.js')])
    test = read('tests', 'roundtrip.js')

    stored = bytes.fromhex(run(core + '\nvar DEFLATED_HEX = "", EMIT = true;\n' + test))
    zin = zipfile.ZipFile(io.BytesIO(stored))
    assert zin.testzip() is None, 'zip CRC check failed'
    for name in zin.namelist():
        if name.endswith('.xml') or name.endswith('.rels'):
            xml.dom.minidom.parseString(zin.read(name))
    print('parts well formed:', len(zin.namelist()))

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED, compresslevel=9) as zout:
        for name in zin.namelist():
            zout.writestr(name, zin.read(name))
    out = run(core + '\nvar DEFLATED_HEX = "%s", EMIT = false;\n' % buf.getvalue().hex() + test)
    print('round trip through deflate:', out)

    # A copy of the same workbook after a real save in Excel (shared strings, Excel's own XML).
    fixture = os.path.join(ROOT, 'tests', 'fixtures', 'saved-by-excel.xlsx')
    if os.path.exists(fixture):
        with open(fixture, 'rb') as f:
            out = run(core + '\nvar DEFLATED_HEX = "%s", EMIT = false;\n' % f.read().hex() + test)
        print('round trip through Excel:', out)

    if len(sys.argv) > 1:
        with open(sys.argv[1], 'wb') as f:
            f.write(stored)
        print('wrote', sys.argv[1])


if __name__ == '__main__':
    main()
