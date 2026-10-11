/* ============================================================
 * PPT Studio — 布局引擎
 * 把「幻灯片数据」编译成中立的绘制指令（ops），单位统一为英寸。
 *   - pptx.js  把 ops 翻译成 OOXML
 *   - preview.js 把 ops 翻译成 HTML
 * 两边共用同一套排版结果，所以网页预览 = 导出效果（WYSIWYG）。
 * ============================================================ */
(function (G) {
  'use strict';

  var PPT = (G.PPT = G.PPT || {});

  var W = 13.333, H = 7.5;          // 16:9 画布（英寸）
  var MX = 0.85;                    // 左右页边距
  var CW = W - MX * 2;              // 正文宽度 11.633
  var BODY_TOP = 2.02;              // 正文起始 y
  var BODY_BOTTOM = 6.84;           // 正文结束 y（下面留给页脚）

  /* ---------- 文字宽度估算（单位：em） ---------- */
  function em(str) {
    str = String(str == null ? '' : str);
    var w = 0, i, c;
    for (i = 0; i < str.length; i++) {
      c = str.charCodeAt(i);
      if (c >= 0x1100 && (c <= 0x115F || (c >= 0x2E80 && c <= 0xA4CF) ||
          (c >= 0xAC00 && c <= 0xD7A3) || (c >= 0xF900 && c <= 0xFAFF) ||
          (c >= 0xFE30 && c <= 0xFE4F) || (c >= 0xFF00 && c <= 0xFF60) ||
          (c >= 0xFFE0 && c <= 0xFFE6))) {
        w += 1.0;                                   // 全角
      } else if (c < 128) {
        if (c >= 65 && c <= 90) w += 0.66;          // 大写
        else if (c === 32) w += 0.28;
        else if (c >= 97 && c <= 122) w += 0.54;
        else if (c >= 48 && c <= 57) w += 0.55;
        else w += 0.42;
      } else {
        w += 0.88;                                  // 其它宽字符
      }
    }
    return w;
  }

  /** 估算若干条目在给定宽度/字号下占用的高度（英寸）。 */
  function measure(items, widthIn, sz, lh) {
    lh = lh || 1.3;
    var total = 0;
    for (var i = 0; i < items.length; i++) {
      var it = items[i];
      var s = sz * (it.mul || 1);
      var avail = Math.max(0.5, widthIn - (it.indent || 0));
      var lines = Math.max(1, Math.ceil(em(it.text) * s / 72 / avail));
      total += lines * s * lh / 72 + (it.gap || 0);
    }
    return total;
  }

  /** 在给定盒子里挑一个能放下的字号（pt）。 */
  function fitSize(items, widthIn, heightIn, maxSz, minSz, lh) {
    lh = lh || 1.3;
    for (var sz = maxSz; sz >= minSz; sz--) {
      if (measure(items, widthIn, sz, lh) <= heightIn) return sz;
    }
    return minSz;
  }

  /* ---------- 绘制指令构造器 ---------- */
  function rect(x, y, w, h, fill, o) {
    o = o || {};
    return { type: o.round ? 'round' : 'rect', x: x, y: y, w: w, h: h, fill: fill,
             line: o.line, lineW: o.lineW, radius: o.radius, alpha: o.alpha };
  }
  function ellipse(x, y, w, h, fill, o) {
    o = o || {};
    return { type: 'ellipse', x: x, y: y, w: w, h: h, fill: fill, alpha: o.alpha, line: o.line, lineW: o.lineW };
  }
  function text(x, y, w, h, paras, o) {
    o = o || {};
    return { type: 'text', x: x, y: y, w: w, h: h, paras: paras, anchor: o.anchor || 't',
             align: o.align || 'l', wrap: o.wrap !== false, ins: o.ins || 0, shrink: o.shrink };
  }
  function para(runs, o) {
    o = o || {};
    if (typeof runs === 'string') runs = [{ t: runs }];
    return { runs: runs, align: o.align, marL: o.marL || 0, indent: o.indent || 0,
             bullet: o.bullet || null, line: o.line || 0, before: o.before || 0, after: o.after || 0 };
  }
  function run(t, o) {
    o = o || {};
    return { t: t, sz: o.sz, b: o.b, i: o.i, color: o.color, spc: o.spc, alpha: o.alpha, font: o.font };
  }
  function bullet(ch, color, mul) { return { char: ch, color: color, size: mul || 1 }; }

  /* ---------- 页眉 / 页脚 ---------- */
  function titleSize(t, maxSz) {
    var n = em(t);
    if (n <= 20) return maxSz;
    if (n <= 26) return maxSz - 3;
    if (n <= 34) return maxSz - 6;
    if (n <= 44) return maxSz - 9;
    return maxSz - 12;
  }

  function header(ops, th, kicker, title) {
    if (kicker) {
      ops.push(text(MX, 0.56, CW - 2, 0.28,
        [para([run(kicker, { sz: 11, b: true, color: th.accent, spc: 160 })])]));
    }
    var tsz = titleSize(title, 30);
    ops.push(text(MX, 0.84, CW - 0.6, 0.8,
      [para([run(title, { sz: tsz, b: true, color: th.text })])], { anchor: 'b' }));
    ops.push(rect(MX, 1.70, 0.72, 0.055, th.accent));
  }

  function footer(ops, th, deck, page) {
    ops.push(rect(MX, 6.92, CW, 0.008, th.border));
    var left = deck.title || '';
    if (left.length > 26) left = left.slice(0, 25) + '…';
    ops.push(text(MX, 6.99, CW - 1.2, 0.26,
      [para([run(left, { sz: 9, color: th.muted })])]));
    ops.push(text(MX + CW - 1.2, 6.99, 1.2, 0.26,
      [para([run(String(page), { sz: 9, color: th.muted })])], { align: 'r' }));
  }

  /* ---------- 工具：把 bullets 归一化 ---------- */
  function normBullets(list) {
    var out = [];
    (list || []).forEach(function (b) {
      if (b == null) return;
      if (typeof b === 'string') { if (b.trim()) out.push({ t: b.trim(), s: [] }); return; }
      var main = b.t || b.text || b.title || '';
      var subs = b.s || b.sub || b.children || [];
      if (typeof subs === 'string') subs = [subs];
      if (main || subs.length) out.push({ t: String(main), s: subs.map(String) });
    });
    return out;
  }

  /* ============================================================
   * 各版式
   * ============================================================ */

  function slideCover(deck, th) {
    var ops = [];
    ops.push(ellipse(9.9, -1.9, 5.6, 5.6, { c: th.accent, a: 0.10 }));
    ops.push(ellipse(11.5, 1.1, 2.4, 2.4, { c: th.accent2, a: 0.13 }));
    ops.push(ellipse(-1.4, 5.0, 3.4, 3.4, { c: th.accent2, a: 0.07 }));

    ops.push(rect(MX, 2.30, 0.07, 1.52, th.accent));
    var tsz = fitSize([{ text: deck.title || '' }], 9.5, 1.7, 46, 26, 1.18);
    ops.push(text(1.20, 2.24, 9.7, 1.7,
      [para([run(deck.title || '', { sz: tsz, b: true, color: th.text })])], { anchor: 'b' }));
    if (deck.subtitle) {
      ops.push(text(1.22, 4.06, 9.6, 0.95,
        [para([run(deck.subtitle, { sz: 17, color: th.muted })], { line: 1.35 })]));
    }
    var meta = [deck.author, deck.org, deck.date].filter(Boolean).join('   ·   ');
    if (meta) {
      ops.push(rect(1.22, 5.72, 11.2, 0.008, th.border));
      ops.push(text(1.22, 5.86, 11.2, 0.34, [para([run(meta, { sz: 12, color: th.muted })])]));
    }
    return ops;
  }

  function slideSection(s, th, idx) {
    var ops = [];
    ops.push(rect(0, 0, 0.26, H, th.accent));
    ops.push(ellipse(10.3, 4.4, 4.6, 4.6, { c: th.accent, a: 0.07 }));
    ops.push(text(MX + 0.05, 1.05, 6.0, 2.0,
      [para([run(s.index || pad2(idx), { sz: 96, b: true, color: th.accent, alpha: 0.22 })])]));
    ops.push(rect(MX + 0.07, 3.18, 0.9, 0.07, th.accent));
    var tsz = titleSize(s.title || '', 40);
    ops.push(text(MX + 0.05, 3.42, 10.6, 1.2,
      [para([run(s.title || '', { sz: tsz, b: true, color: th.text })])]));
    if (s.subtitle) {
      ops.push(text(MX + 0.05, 4.66, 10.6, 1.0,
        [para([run(s.subtitle, { sz: 16, color: th.muted })], { line: 1.4 })]));
    }
    return ops;
  }

  function slideBullets(s, th, deck, page) {
    var ops = [];
    header(ops, th, s.kicker, s.title || '');
    var items = normBullets(s.bullets);
    var L = [];
    items.forEach(function (b, i) {
      L.push({ text: b.t, indent: 0.34, mul: 1, gap: i === 0 ? 0 : 0.10 });
      b.s.forEach(function (sub, j) {
        L.push({ text: sub, indent: 0.62, mul: 0.82, gap: j === 0 ? 0.045 : 0.03 });
      });
    });
    if (!L.length) L.push({ text: '（本页暂无内容）', indent: 0, mul: 1, gap: 0 });
    var band = BODY_BOTTOM - BODY_TOP;
    var sz = fitSize(L, CW, band, 21, 11, 1.3);

    var paras = [];
    items.forEach(function (b, i) {
      paras.push(para([run(b.t, { sz: sz, color: th.text, b: false })], {
        bullet: bullet('▪', th.accent), marL: 0.34, indent: -0.34, line: 1.3, before: i === 0 ? 0 : 7
      }));
      b.s.forEach(function (sub, j) {
        paras.push(para([run(sub, { sz: Math.max(11, Math.round(sz * 0.82)), color: th.muted })], {
          bullet: bullet('–', th.muted), marL: 0.68, indent: -0.30, line: 1.28, before: 3
        }));
      });
    });
    var hgt = Math.min(band, measure(L, CW, sz, 1.3));
    ops.push(text(MX, BODY_TOP + Math.max(0, (band - hgt) / 2) - 0.10, CW, hgt + 0.20, paras));
    footer(ops, th, deck, page);
    return ops;
  }

  function slideAgenda(s, th, deck, page) {
    var ops = [];
    header(ops, th, s.kicker || 'CONTENTS', s.title || '目录');
    var items = (s.items && s.items.length ? s.items : (s.bullets || [])).map(function (it) {
      return typeof it === 'string' ? { title: it } : it;
    });
    if (!items.length) items = [{ title: '（请输入大纲）' }];
    var n = items.length;
    var startY = BODY_TOP + 0.18;
    var availH = BODY_BOTTOM - startY - 0.12;
    var rowH = Math.min(0.90, availH / n);
    var tsz = fitSize(items.map(function (it) { return { text: it.title || it.t || '', indent: 0.95 }; }),
      CW - 0.95, availH, 19, 12, 1.25);
    var y = startY;
    items.forEach(function (it, i) {
      var label = it.title || it.t || '';
      ops.push(text(MX, y, 0.78, rowH,
        [para([run(pad2(i + 1), { sz: Math.min(24, Math.round(rowH * 30)), b: true, color: th.accent })])],
        { anchor: 'ctr' }));
      ops.push(text(MX + 0.95, y, CW - 0.95, rowH,
        [para([run(label, { sz: tsz, color: th.text })], { line: 1.25 })], { anchor: 'ctr' }));
      if (i < n - 1) ops.push(rect(MX, y + rowH - 0.01, CW, 0.007, th.border));
      y += rowH;
    });
    footer(ops, th, deck, page);
    return ops;
  }

  function slideTwoCol(s, th, deck, page) {
    var ops = [];
    header(ops, th, s.kicker, s.title || '');
    var colW = 5.33, rx = MX + colW + 0.62;
    var band = BODY_BOTTOM - BODY_TOP, HEAD = 0.62;

    // 先把两栏都解析出来，再统一字号（左右字号不一致会很明显）
    var cols = [['left', MX], ['right', rx]].map(function (pair) {
      var side = s[pair[0]];
      if (!side) return null;
      var items = normBullets(side.bullets);
      var L = [];
      items.forEach(function (b, i) {
        L.push({ text: b.t, indent: 0.26, mul: 1, gap: i === 0 ? 0 : 0.09 });
        b.s.forEach(function (sub, j) {
          L.push({ text: sub, indent: 0.5, mul: 0.85, gap: j === 0 ? 0.04 : 0.03 });
        });
      });
      return { x: pair[1], side: side, items: items, L: L, top: HEAD, hasHead: !!side.title };
    }).filter(Boolean);

    var sizeCap = cols.map(function (c) {
      return fitSize(c.L, colW, band - (c.hasHead ? HEAD : 0), 15, 10, 1.28);
    });
    var sz = sizeCap.length ? Math.min.apply(null, sizeCap) : 14;

    var tallest = Math.max.apply(null, cols.map(function (c) { return measure(c.L, colW, sz, 1.28); }).concat([0.4]));
    var anyHead = cols.some(function (c) { return c.hasHead; });
    var groupTop = BODY_TOP + Math.max(0, (band - (tallest + (anyHead ? HEAD : 0))) / 2);
    var groupBottom = groupTop + tallest + (anyHead ? HEAD : 0);
    ops.push(rect(MX + colW + 0.30, groupTop, 0.008, Math.max(0.8, groupBottom - groupTop), th.border));

    cols.forEach(function (c) {
      var top = groupTop;
      if (c.hasHead) {
        ops.push(text(c.x, top, colW, 0.42,
          [para([run(c.side.title, { sz: 17, b: true, color: th.accent })])]));
        ops.push(rect(c.x, top + 0.44, 0.42, 0.045, th.accent));
        top += HEAD;
      }
      var paras = [];
      c.items.forEach(function (b, i) {
        paras.push(para([run(b.t, { sz: sz, color: th.text })], {
          bullet: bullet('▪', th.accent), marL: 0.26, indent: -0.26, line: 1.28, before: i === 0 ? 0 : 6
        }));
        b.s.forEach(function (sub) {
          paras.push(para([run(sub, { sz: Math.max(10, Math.round(sz * 0.85)), color: th.muted })], {
            bullet: bullet('–', th.muted), marL: 0.5, indent: -0.24, line: 1.26, before: 3
          }));
        });
      });
      if (paras.length) ops.push(text(c.x, top, colW, Math.max(0.5, tallest), paras));
    });
    footer(ops, th, deck, page);
    return ops;
  }

  function slideStats(s, th, deck, page) {
    var ops = [];
    header(ops, th, s.kicker, s.title || '');
    var stats = (s.stats || []).slice(0, 8);
    if (!stats.length) stats = [{ value: '—', label: '请输入指标' }];
    var rows = stats.length <= 4 ? 1 : 2;
    var per = Math.ceil(stats.length / rows);
    var gap = 0.30;
    var cw = (CW - gap * (per - 1)) / per;
    var ch = rows === 1 ? 2.60 : 2.05;
    var top = rows === 1 ? 2.35 : 2.15;
    stats.forEach(function (st, i) {
      var r = Math.floor(i / per), c = i % per;
      var x = MX + c * (cw + gap), y = top + r * (ch + 0.30);
      ops.push(rect(x, y, cw, ch, th.surface, { round: true, radius: 0.12, line: th.border, lineW: 0.75 }));
      ops.push(rect(x + 0.30, y + 0.34, 0.52, 0.055, i % 2 ? th.accent2 : th.accent));
      var vsz = fitSize([{ text: String(st.value == null ? '' : st.value) }], cw - 0.6, 0.95, 40, 20, 1.1);
      ops.push(text(x + 0.30, y + 0.55, cw - 0.6, 0.95,
        [para([run(String(st.value == null ? '' : st.value), { sz: vsz, b: true, color: i % 2 ? th.accent2 : th.accent })])],
        { anchor: 'ctr' }));
      ops.push(text(x + 0.30, y + 1.52, cw - 0.6, 0.42,
        [para([run(st.label || '', { sz: 14, color: th.text })], { line: 1.2 })], { anchor: 't' }));
      if (st.sub) {
        ops.push(text(x + 0.30, y + 1.94, cw - 0.6, 0.5,
          [para([run(st.sub, { sz: 10.5, color: th.muted })], { line: 1.2 })]));
      }
    });
    footer(ops, th, deck, page);
    return ops;
  }

  function slideTable(s, th, deck, page) {
    var ops = [];
    header(ops, th, s.kicker, s.title || '');
    var t = s.table || {};
    var headers = (t.headers || []).map(String);
    var rows = (t.rows || []).map(function (r) { return r.map(String); });
    var nCol = Math.max(headers.length, rows.reduce(function (m, r) { return Math.max(m, r.length); }, 0), 1);
    // 补齐列
    function padRow(r) { var o = r.slice(0, nCol); while (o.length < nCol) o.push(''); return o; }
    if (headers.length) headers = padRow(headers);

    // 列宽：按内容宽度加权
    var weights = [];
    for (var c = 0; c < nCol; c++) {
      var mx = 2;
      if (headers[c]) mx = Math.max(mx, em(headers[c]) * 1.1);
      rows.forEach(function (r) { mx = Math.max(mx, em(r[c] || '')); });
      weights.push(Math.min(mx, 26));
    }
    var sum = weights.reduce(function (a, b) { return a + b; }, 0);
    var colW = weights.map(function (w) { return Math.max(0.95, CW * w / sum); });
    var scale = CW / colW.reduce(function (a, b) { return a + b; }, 0);
    colW = colW.map(function (w) { return w * scale; });

    var nRow = rows.length + (headers.length ? 1 : 0);
    var availH = BODY_BOTTOM - BODY_TOP;
    var rowH = Math.max(0.34, Math.min(0.86, availH / Math.max(1, nRow)));
    var fs = nRow > 10 ? 9.5 : nRow > 8 ? 10.5 : nRow > 6 ? 11.5 : 12.5;

    var trs = [];
    if (headers.length) trs.push({ cells: headers, head: true, h: rowH });
    rows.forEach(function (r) { trs.push({ cells: padRow(r), head: false, h: rowH }); });

    ops.push({
      type: 'table', x: MX, y: BODY_TOP + 0.05, w: CW, h: rowH * Math.max(1, nRow),
      colW: colW, rows: trs,
      fontSize: fs, theme: th, headFill: th.accent, headText: 'FFFFFF',
      bodyFill: th.surface, altFill: th.bg, border: th.border, text: th.text
    });
    footer(ops, th, deck, page);
    return ops;
  }

  function slideTimeline(s, th, deck, page) {
    var ops = [];
    header(ops, th, s.kicker, s.title || '');
    var items = (s.items || []).slice(0, 5);
    if (!items.length) items = [{ when: '阶段一', what: '描述' }];
    var x0 = MX + 0.25, x1 = W - MX - 0.25;
    var n = items.length, span = x1 - x0, iw = span / n, colW = Math.max(1.2, iw - 0.20);

    var caps = items.map(function (it) {
      return fitSize([{ text: it.what || it.title || '' }], colW, 1.6, 13.5, 10, 1.35);
    });
    var dsz = caps.length ? Math.min.apply(null, caps) : 12.5;
    var descH = Math.max.apply(null, items.map(function (it) {
      return measure([{ text: it.what || it.title || '' }], colW, dsz, 1.35);
    }).concat([0.4]));

    var periodH = 0.44, dotH = 0.26, gapD = 0.20;
    var blockH = periodH + 0.16 + dotH + gapD + descH;
    var band = BODY_BOTTOM - BODY_TOP;
    var top = BODY_TOP + Math.max(0, (band - blockH) / 2);
    var cy = top + periodH + 0.16, descY = cy + dotH + gapD;

    ops.push(rect(x0, cy + 0.12, span, 0.022, th.border));
    items.forEach(function (it, i) {
      var cx = x0 + iw * i + iw / 2;
      ops.push(text(cx - iw / 2 + 0.10, top, colW, periodH,
        [para([run(it.when || '', { sz: 15, b: true, color: th.accent })], { line: 1.2, align: 'c' })],
        { anchor: 'b', align: 'c' }));
      ops.push(ellipse(cx - 0.13, cy, dotH, dotH, i % 2 ? th.accent2 : th.accent));
      ops.push(ellipse(cx - 0.05, cy + 0.08, 0.10, 0.10, th.bg));
      ops.push(text(cx - iw / 2 + 0.10, descY, colW, descH + 0.20,
        [para([run(it.what || it.title || '', { sz: dsz, color: th.text })], { line: 1.35, align: 'c' })],
        { align: 'c' }));
    });
    footer(ops, th, deck, page);
    return ops;
  }

  function slideQuote(s, th, deck, page) {
    var ops = [];
    var txt = s.text || s.title || '';
    var sz = fitSize([{ text: txt }], CW - 0.9, 3.0, 28, 15, 1.45);
    var hgt = measure([{ text: txt }], CW - 0.9, sz, 1.45);
    var band = BODY_BOTTOM - BODY_TOP;
    var textTop = BODY_TOP + Math.max(0.1, (band - (hgt + 0.95 + (s.author ? 0.42 : 0))) / 2);
    var lineY = textTop + hgt + 0.34;
    ops.push(text(MX, textTop - 0.72, 3.0, 1.6,
      [para([run('\u201C', { sz: 120, b: true, color: th.accent, alpha: 0.20 })])]));
    ops.push(text(MX + 0.75, textTop, CW - 0.9, hgt + 0.40,
      [para([run(txt, { sz: sz, color: th.text })], { line: 1.45 })]));
    ops.push(rect(MX + 0.75, lineY, 0.55, 0.055, th.accent));
    if (s.author) {
      ops.push(text(MX + 0.75, lineY + 0.22, CW - 0.9, 0.40,
        [para([run(s.author, { sz: 14, color: th.muted })])]));
    }
    footer(ops, th, deck, page);
    return ops;
  }

  function slideClosing(deck, th) {
    var ops = [];
    ops.push(ellipse(4.4, -2.6, 6.2, 6.2, { c: th.accent, a: 0.09 }));
    ops.push(ellipse(9.4, 4.0, 4.6, 4.6, { c: th.accent2, a: 0.11 }));
    var s = deck.closing || {};
    var tsz = fitSize([{ text: s.title || '谢谢观看' }], 10.4, 1.3, 46, 28, 1.15);
    ops.push(text(1.46, 2.55, 10.4, 1.3,
      [para([run(s.title || '谢谢观看', { sz: tsz, b: true, color: th.text })], { align: 'c' })],
      { anchor: 'b', align: 'c' }));
    ops.push(rect((W - 0.8) / 2, 4.12, 0.8, 0.06, th.accent));
    if (s.subtitle) {
      ops.push(text(1.46, 4.42, 10.4, 0.8,
        [para([run(s.subtitle, { sz: 16, color: th.muted })], { align: 'c', line: 1.35 })], { align: 'c' }));
    }
    var contact = [deck.author, deck.org, deck.date].filter(Boolean).join('   ·   ');
    if (contact) {
      ops.push(text(1.46, 5.62, 10.4, 0.4,
        [para([run(contact, { sz: 12, color: th.muted })], { align: 'c' })], { align: 'c' }));
    }
    return ops;
  }

  function pad2(n) { return (n < 10 ? '0' : '') + n; }

  /* ============================================================
   * 主入口
   * ============================================================ */
  function build(deck) {
    var th = PPT.getTheme(deck.theme);
    var slides = deck.slides || [];
    var out = [];
    var page = 0;
    slides.forEach(function (s, i) {
      var type = s.layout || s.type || 'bullets';
      var ops, bg = { fill: th.bg }, numbered = true;
      page++;
      if (type === 'cover') { ops = slideCover(deck, th); bg = { grad: [th.bg, th.bgAlt], angle: 45 }; numbered = false; }
      else if (type === 'section') { ops = slideSection(s, th, i + 1); bg = { grad: [th.bg, th.bgAlt], angle: 45 }; numbered = false; }
      else if (type === 'agenda' || type === 'toc') ops = slideAgenda(s, th, deck, page);
      else if (type === 'twoCol' || type === 'two-col') ops = slideTwoCol(s, th, deck, page);
      else if (type === 'stats' || type === 'kpi') ops = slideStats(s, th, deck, page);
      else if (type === 'table') ops = slideTable(s, th, deck, page);
      else if (type === 'timeline') ops = slideTimeline(s, th, deck, page);
      else if (type === 'quote') ops = slideQuote(s, th, deck, page);
      else if (type === 'closing' || type === 'end') { ops = slideClosing(deck, th); bg = { grad: [th.bg, th.bgAlt], angle: 45 }; numbered = false; }
      else ops = slideBullets(s, th, deck, page);

      out.push({
        layout: type, bg: bg, ops: ops, notes: s.notes || '',
        title: s.title || (type === 'cover' ? deck.title : ''),
        numbered: numbered
      });
    });
    return { w: W, h: H, theme: th, slides: out };
  }

  PPT.layout = {
    W: W, H: H, MX: MX, CW: CW,
    BODY_TOP: BODY_TOP, BODY_BOTTOM: BODY_BOTTOM,
    build: build, em: em, fitSize: fitSize, measure: measure, normBullets: normBullets
  };

  if (typeof module !== 'undefined' && module.exports) module.exports = PPT;
})(typeof window !== 'undefined' ? window : globalThis);
