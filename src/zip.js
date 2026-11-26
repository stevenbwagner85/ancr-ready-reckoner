/* AnCR ready reckoner: minimal ZIP reader and writer, plus UTF-8 and base64
 * helpers. Dependency-free so the tool works from a local drive, offline.
 * Writing stores entries uncompressed (Excel accepts this); reading handles
 * both stored and deflated entries, as produced when Excel re-saves a file. */
(function (A) {
  'use strict';

  /* ---------- UTF-8 ---------- */
  function utf8Encode(str) {
    var out = [], i, c, c2;
    for (i = 0; i < str.length; i++) {
      c = str.charCodeAt(i);
      if (c >= 0xD800 && c <= 0xDBFF && i + 1 < str.length) {
        c2 = str.charCodeAt(i + 1);
        if (c2 >= 0xDC00 && c2 <= 0xDFFF) { c = 0x10000 + ((c - 0xD800) << 10) + (c2 - 0xDC00); i++; }
      }
      if (c < 0x80) out.push(c);
      else if (c < 0x800) out.push(0xC0 | (c >> 6), 0x80 | (c & 63));
      else if (c < 0x10000) out.push(0xE0 | (c >> 12), 0x80 | ((c >> 6) & 63), 0x80 | (c & 63));
      else out.push(0xF0 | (c >> 18), 0x80 | ((c >> 12) & 63), 0x80 | ((c >> 6) & 63), 0x80 | (c & 63));
    }
    return new Uint8Array(out);
  }

  function utf8Decode(bytes) {
    var parts = [], chunk = [], i = 0, c, cp;
    if (bytes.length >= 3 && bytes[0] === 0xEF && bytes[1] === 0xBB && bytes[2] === 0xBF) i = 3;
    while (i < bytes.length) {
      c = bytes[i++];
      if (c < 0x80) cp = c;
      else if (c < 0xE0) cp = ((c & 31) << 6) | (bytes[i++] & 63);
      else if (c < 0xF0) { cp = ((c & 15) << 12) | ((bytes[i] & 63) << 6) | (bytes[i + 1] & 63); i += 2; }
      else { cp = ((c & 7) << 18) | ((bytes[i] & 63) << 12) | ((bytes[i + 1] & 63) << 6) | (bytes[i + 2] & 63); i += 3; }
      if (cp >= 0x10000) { cp -= 0x10000; chunk.push(0xD800 + (cp >> 10), 0xDC00 + (cp & 1023)); }
      else chunk.push(cp);
      if (chunk.length > 8000) { parts.push(String.fromCharCode.apply(null, chunk)); chunk = []; }
    }
    parts.push(String.fromCharCode.apply(null, chunk));
    return parts.join('');
  }

  /* ---------- base64 ---------- */
  function base64Decode(b64) {
    var chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/';
    var map = {}, i, out = [], buf = 0, bits = 0, v;
    for (i = 0; i < 64; i++) map[chars.charAt(i)] = i;
    for (i = 0; i < b64.length; i++) {
      v = map[b64.charAt(i)];
      if (v === undefined) continue;
      buf = (buf << 6) | v; bits += 6;
      if (bits >= 8) { bits -= 8; out.push((buf >> bits) & 255); }
    }
    return new Uint8Array(out);
  }

  /* ---------- CRC-32 ---------- */
  var crcTable = null;
  function crc32(bytes) {
    var i, j, c;
    if (!crcTable) {
      crcTable = [];
      for (i = 0; i < 256; i++) {
        c = i;
        for (j = 0; j < 8; j++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
        crcTable[i] = c >>> 0;
      }
    }
    c = 0xFFFFFFFF;
    for (i = 0; i < bytes.length; i++) c = crcTable[(c ^ bytes[i]) & 255] ^ (c >>> 8);
    return (c ^ 0xFFFFFFFF) >>> 0;
  }

  /* ---------- ZIP writer (stored) ---------- */
  function zipWrite(files) {
    // files: [{name: string, data: Uint8Array}]
    var chunks = [], central = [], offset = 0;
    function u16(a, v) { a.push(v & 255, (v >> 8) & 255); }
    function u32(a, v) { a.push(v & 255, (v >>> 8) & 255, (v >>> 16) & 255, (v >>> 24) & 255); }
    var dosTime = 0, dosDate = ((2026 - 1980) << 9) | (1 << 5) | 1;
    files.forEach(function (f) {
      var name = utf8Encode(f.name), crc = crc32(f.data), h = [], c = [];
      u32(h, 0x04034b50); u16(h, 20); u16(h, 0x0800); u16(h, 0);
      u16(h, dosTime); u16(h, dosDate); u32(h, crc);
      u32(h, f.data.length); u32(h, f.data.length); u16(h, name.length); u16(h, 0);
      chunks.push(new Uint8Array(h), name, f.data);
      u32(c, 0x02014b50); u16(c, 20); u16(c, 20); u16(c, 0x0800); u16(c, 0);
      u16(c, dosTime); u16(c, dosDate); u32(c, crc);
      u32(c, f.data.length); u32(c, f.data.length); u16(c, name.length);
      u16(c, 0); u16(c, 0); u16(c, 0); u16(c, 0); u32(c, 0); u32(c, offset);
      central.push(new Uint8Array(c), name);
      offset += h.length + name.length + f.data.length;
    });
    var cdSize = 0;
    central.forEach(function (x) { cdSize += x.length; });
    var e = [];
    u32(e, 0x06054b50); u16(e, 0); u16(e, 0); u16(e, files.length); u16(e, files.length);
    u32(e, cdSize); u32(e, offset); u16(e, 0);
    var all = chunks.concat(central, [new Uint8Array(e)]);
    var out = new Uint8Array(offset + cdSize + e.length), p = 0;
    all.forEach(function (x) { out.set(x, p); p += x.length; });
    return out;
  }

  /* ---------- Inflate (RFC 1951) ---------- */
  var LEN_BASE = [3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 15, 17, 19, 23, 27, 31, 35, 43, 51, 59, 67, 83, 99, 115, 131, 163, 195, 227, 258];
  var LEN_EXTRA = [0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 5, 0];
  var DIST_BASE = [1, 2, 3, 4, 5, 7, 9, 13, 17, 25, 33, 49, 65, 97, 129, 193, 257, 385, 513, 769, 1025, 1537, 2049, 3073, 4097, 6145, 8193, 12289, 16385, 24577];
  var DIST_EXTRA = [0, 0, 0, 0, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 6, 6, 7, 7, 8, 8, 9, 9, 10, 10, 11, 11, 12, 12, 13, 13];
  var CL_ORDER = [16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15];

  function huffman(lengths, n) {
    var count = new Uint16Array(16), offs = new Uint16Array(16), symbol = new Uint16Array(n), i;
    for (i = 0; i < n; i++) count[lengths[i]]++;
    count[0] = 0;
    for (i = 1; i < 16; i++) offs[i] = offs[i - 1] + count[i - 1];
    for (i = 0; i < n; i++) if (lengths[i]) symbol[offs[lengths[i]]++] = i;
    return { count: count, symbol: symbol };
  }

  function inflate(src, expected) {
    var pos = 0, bitBuf = 0, bitCnt = 0;
    var out = new Uint8Array(expected && expected > 0 ? expected : Math.max(1024, src.length * 4)), op = 0;

    function need(n) {
      if (op + n > out.length) {
        var bigger = new Uint8Array(Math.max(out.length * 2, op + n));
        bigger.set(out); out = bigger;
      }
    }
    function bits(n) {
      var v = bitBuf;
      while (bitCnt < n) {
        if (pos >= src.length) throw new Error('Unexpected end of compressed data');
        v |= src[pos++] << bitCnt; bitCnt += 8;
      }
      bitBuf = v >>> n; bitCnt -= n;
      return v & ((1 << n) - 1);
    }
    function decode(h) {
      var code = 0, first = 0, index = 0, len, c;
      for (len = 1; len <= 15; len++) {
        code |= bits(1);
        c = h.count[len];
        if (code - c < first) return h.symbol[index + (code - first)];
        index += c; first += c; first <<= 1; code <<= 1;
      }
      throw new Error('Invalid compressed data');
    }
    function codes(lh, dh) {
      var sym, len, dist, i;
      for (;;) {
        sym = decode(lh);
        if (sym < 256) { need(1); out[op++] = sym; }
        else if (sym === 256) return;
        else {
          sym -= 257;
          if (sym >= 29) throw new Error('Invalid compressed data');
          len = LEN_BASE[sym] + bits(LEN_EXTRA[sym]);
          sym = decode(dh);
          if (sym >= 30) throw new Error('Invalid compressed data');
          dist = DIST_BASE[sym] + bits(DIST_EXTRA[sym]);
          if (dist > op) throw new Error('Invalid compressed data');
          need(len);
          for (i = 0; i < len; i++) { out[op] = out[op - dist]; op++; }
        }
      }
    }

    var fixedL = null, fixedD = null, last, type, i;
    do {
      last = bits(1); type = bits(2);
      if (type === 0) {
        bitBuf = 0; bitCnt = 0;
        var len = src[pos] | (src[pos + 1] << 8); pos += 4;
        need(len);
        out.set(src.subarray(pos, pos + len), op); op += len; pos += len;
      } else if (type === 1) {
        if (!fixedL) {
          var l = new Uint8Array(288);
          for (i = 0; i < 144; i++) l[i] = 8;
          for (; i < 256; i++) l[i] = 9;
          for (; i < 280; i++) l[i] = 7;
          for (; i < 288; i++) l[i] = 8;
          fixedL = huffman(l, 288);
          var d = new Uint8Array(30);
          for (i = 0; i < 30; i++) d[i] = 5;
          fixedD = huffman(d, 30);
        }
        codes(fixedL, fixedD);
      } else if (type === 2) {
        var nlen = bits(5) + 257, ndist = bits(5) + 1, ncode = bits(4) + 4;
        var lens = new Uint8Array(320), idx = 0, sym, rep, prev;
        for (i = 0; i < ncode; i++) lens[CL_ORDER[i]] = bits(3);
        var clh = huffman(lens.subarray(0, 19), 19);
        var ll = new Uint8Array(320);
        while (idx < nlen + ndist) {
          sym = decode(clh);
          if (sym < 16) ll[idx++] = sym;
          else {
            prev = 0;
            if (sym === 16) { if (idx === 0) throw new Error('Invalid compressed data'); prev = ll[idx - 1]; rep = 3 + bits(2); }
            else if (sym === 17) rep = 3 + bits(3);
            else rep = 11 + bits(7);
            if (idx + rep > nlen + ndist) throw new Error('Invalid compressed data');
            while (rep--) ll[idx++] = prev;
          }
        }
        codes(huffman(ll.subarray(0, nlen), nlen), huffman(ll.subarray(nlen, nlen + ndist), ndist));
      } else {
        throw new Error('Invalid compressed data');
      }
    } while (!last);
    return out.subarray(0, op);
  }

  /* ---------- ZIP reader ---------- */
  function zipRead(buf) {
    // returns { 'path/name': Uint8Array }
    function u16(p) { return buf[p] | (buf[p + 1] << 8); }
    function u32(p) { return (buf[p] | (buf[p + 1] << 8) | (buf[p + 2] << 16) | (buf[p + 3] << 24)) >>> 0; }
    var e = -1, i;
    for (i = buf.length - 22; i >= 0 && i >= buf.length - 66000; i--) {
      if (u32(i) === 0x06054b50) { e = i; break; }
    }
    if (e < 0) throw new Error('This is not an Excel workbook (.xlsx).');
    var n = u16(e + 10), p = u32(e + 16), files = {};
    for (i = 0; i < n; i++) {
      if (u32(p) !== 0x02014b50) throw new Error('The file appears to be damaged.');
      var flags = u16(p + 8), method = u16(p + 10), csize = u32(p + 20), usize = u32(p + 24);
      var nlen = u16(p + 28), xlen = u16(p + 30), clen = u16(p + 32), lho = u32(p + 42);
      var name = utf8Decode(buf.subarray(p + 46, p + 46 + nlen));
      if (flags & 1) throw new Error('This workbook is password-protected and cannot be opened here.');
      var dataStart = lho + 30 + u16(lho + 26) + u16(lho + 28);
      var raw = buf.subarray(dataStart, dataStart + csize);
      if (method === 0) files[name] = raw;
      else if (method === 8) files[name] = inflate(raw, usize);
      else throw new Error('Unsupported compression in this file.');
      p += 46 + nlen + xlen + clen;
    }
    return files;
  }

  A.utf8Encode = utf8Encode;
  A.utf8Decode = utf8Decode;
  A.base64Decode = base64Decode;
  A.crc32 = crc32;
  A.zipWrite = zipWrite;
  A.zipRead = zipRead;
  A.inflate = inflate;
})(ANCR);
