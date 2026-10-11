/* ============================================================
 * PPT Studio — 零依赖 ZIP 打包器
 * .pptx 本质就是一个 ZIP（OOXML 包）。浏览器里没有 zlib，
 * 所以这里自己实现 CRC32 + ZIP 结构；压缩优先用原生
 * CompressionStream('deflate-raw')，不可用时自动降级为
 * 「存储（store）」模式 —— 任何解压器都能读，PowerPoint 也接受。
 * ============================================================ */
(function (G) {
  'use strict';

  var PPT = (G.PPT = G.PPT || {});

  /* ---------- CRC32 ---------- */
  var CRC_TABLE = (function () {
    var t = new Uint32Array(256);
    for (var n = 0; n < 256; n++) {
      var c = n;
      for (var k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
      t[n] = c >>> 0;
    }
    return t;
  })();

  function crc32(bytes) {
    var c = 0xFFFFFFFF;
    for (var i = 0; i < bytes.length; i++) c = CRC_TABLE[(c ^ bytes[i]) & 0xFF] ^ (c >>> 8);
    return (c ^ 0xFFFFFFFF) >>> 0;
  }

  /* ---------- 文本 → UTF-8 字节 ---------- */
  function utf8(str) {
    if (typeof TextEncoder !== 'undefined') return new TextEncoder().encode(str);
    var out = [], i, c;
    for (i = 0; i < str.length; i++) {
      c = str.charCodeAt(i);
      if (c < 0x80) out.push(c);
      else if (c < 0x800) out.push(0xC0 | (c >> 6), 0x80 | (c & 63));
      else if (c >= 0xD800 && c <= 0xDBFF) {
        var c2 = str.charCodeAt(++i);
        var cp = 0x10000 + ((c - 0xD800) << 10) + (c2 - 0xDC00);
        out.push(0xF0 | (cp >> 18), 0x80 | ((cp >> 12) & 63), 0x80 | ((cp >> 6) & 63), 0x80 | (cp & 63));
      } else out.push(0xE0 | (c >> 12), 0x80 | ((c >> 6) & 63), 0x80 | (c & 63));
    }
    return new Uint8Array(out);
  }

  /* ---------- deflate-raw（可选） ---------- */
  function deflateRaw(bytes) {
    // 1) 浏览器原生（Chrome/Edge 103+）
    if (typeof CompressionStream === 'function') {
      try {
        var cs = new CompressionStream('deflate-raw');
        var w = cs.writable.getWriter();
        w.write(bytes); w.close();
        return new Response(cs.readable).arrayBuffer().then(function (buf) { return new Uint8Array(buf); });
      } catch (e) { /* 继续降级 */ }
    }
    // 2) Node（构建期的自动化测试走这里）
    try {
      if (typeof require === 'function') {
        var zlib = require('zlib');
        return Promise.resolve(new Uint8Array(zlib.deflateRawSync(Buffer.from(bytes))));
      }
    } catch (e2) { /* 继续降级 */ }
    // 3) 存储模式
    return Promise.resolve(null);
  }

  function dosTime(d) {
    return {
      time: ((d.getHours() & 31) << 11) | ((d.getMinutes() & 63) << 5) | ((d.getSeconds() / 2) & 31),
      date: (((d.getFullYear() - 1980) & 127) << 9) | (((d.getMonth() + 1) & 15) << 5) | (d.getDate() & 31)
    };
  }

  /**
   * 打包为 ZIP。
   * @param {Array<{name:string, data:string|Uint8Array}>} entries
   * @returns {Promise<Uint8Array>}
   */
  function buildZip(entries) {
    var dt = dosTime(new Date());
    var parts = [];        // 待拼接的字节块
    var central = [];      // 中央目录记录
    var offset = 0;

    function u16(v) { return [v & 0xFF, (v >> 8) & 0xFF]; }
    function u32(v) { return [v & 0xFF, (v >>> 8) & 0xFF, (v >>> 16) & 0xFF, (v >>> 24) & 0xFF]; }

    return entries.reduce(function (chain, e) {
      return chain.then(function () {
        var raw = (typeof e.data === 'string') ? utf8(e.data) : e.data;
        var nameBytes = utf8(e.name);
        var crc = crc32(raw);
        return deflateRaw(raw).then(function (comp) {
          var useDeflate = !!comp && comp.length < raw.length;
          var body = useDeflate ? comp : raw;
          var method = useDeflate ? 8 : 0;
          var flags = 0x0800; // UTF-8 文件名

          var local = [].concat(
            u32(0x04034B50), u16(20), u16(flags), u16(method), u16(dt.time), u16(dt.date),
            u32(crc), u32(body.length), u32(raw.length), u16(nameBytes.length), u16(0)
          );
          parts.push(new Uint8Array(local), nameBytes, body);

          central.push({ nameBytes: nameBytes, method: method, flags: flags, crc: crc,
                         compSize: body.length, rawSize: raw.length, offset: offset });
          offset += local.length + nameBytes.length + body.length;
        });
      });
    }, Promise.resolve()).then(function () {
      var cdStart = offset;
      central.forEach(function (c) {
        var rec = [].concat(
          u32(0x02014B50), u16(20), u16(20), u16(c.flags), u16(c.method), u16(dt.time), u16(dt.date),
          u32(c.crc), u32(c.compSize), u32(c.rawSize), u16(c.nameBytes.length),
          u16(0), u16(0), u16(0), u16(0), u32(0), u32(c.offset)
        );
        parts.push(new Uint8Array(rec), c.nameBytes);
        offset += rec.length + c.nameBytes.length;
      });
      var cdSize = offset - cdStart;
      var eocd = [].concat(
        u32(0x06054B50), u16(0), u16(0), u16(central.length), u16(central.length),
        u32(cdSize), u32(cdStart), u16(0)
      );
      parts.push(new Uint8Array(eocd));

      var total = parts.reduce(function (n, p) { return n + p.length; }, 0);
      var out = new Uint8Array(total);
      var pos = 0;
      parts.forEach(function (p) { out.set(p, pos); pos += p.length; });
      return out;
    });
  }

  PPT.zip = { crc32: crc32, utf8: utf8, buildZip: buildZip };

  if (typeof module !== 'undefined' && module.exports) module.exports = PPT;
})(typeof window !== 'undefined' ? window : globalThis);
