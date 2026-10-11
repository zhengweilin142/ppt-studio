/* ============================================================
 * PPT Studio — 网页预览渲染器
 * 复用 layout.js 的绘制指令，在浏览器里 1:1 还原幻灯片。
 *   1 英寸 = 96 px，1 pt = 96/72 px，因此预览尺寸与 PPT 完全同构。
 * ============================================================ */
(function (G) {
  'use strict';

  var PPT = (G.PPT = G.PPT || {});
  var U = 96;                     // 每英寸像素
  var SLIDE_W = 1280, SLIDE_H = 720;

  function px(inch) { return (inch * U) + 'px'; }
  function pt(v) { return (v * U / 72) + 'px'; }
  function hex(c) {
    var s = String(c == null ? '000000' : c).replace(/^#/, '');
    if (s.length === 3) s = s[0] + s[0] + s[1] + s[1] + s[2] + s[2];
    return '#' + s;
  }
  function rgba(c, a) {
    var h = hex(c).slice(1);
    var r = parseInt(h.slice(0, 2), 16), g = parseInt(h.slice(2, 4), 16), b = parseInt(h.slice(4, 6), 16);
    return 'rgba(' + r + ',' + g + ',' + b + ',' + (a == null ? 1 : a) + ')';
  }
  function colorOf(c, extraAlpha) {
    if (c && typeof c === 'object') return rgba(c.c, c.a == null ? 1 : c.a);
    return rgba(c, extraAlpha == null ? 1 : extraAlpha);
  }
  function div(cls) { var d = document.createElement('div'); if (cls) d.className = cls; return d; }
  function fontFamily(th) {
    return '"' + th.latin + '","' + th.ea + '","PingFang SC","Hiragino Sans GB","Source Han Sans SC",sans-serif';
  }

  function applyBox(el, op) {
    el.style.position = 'absolute';
    el.style.left = px(op.x);
    el.style.top = px(op.y);
    el.style.width = px(op.w);
    el.style.height = px(op.h);
  }

  function shapeEl(op) {
    var el = div('op op-shape');
    applyBox(el, op);
    el.style.background = op.fill == null ? 'transparent' : colorOf(op.fill, op.alpha);
    if (op.type === 'ellipse') el.style.borderRadius = '50%';
    if (op.type === 'round') el.style.borderRadius = px(op.radius == null ? 0.1 : op.radius);
    if (op.line) el.style.border = (op.lineW || 1) * U / 72 + 'px solid ' + colorOf(op.line);
    return el;
  }

  function textEl(op, th) {
    var box = div('op op-text');
    applyBox(box, op);
    box.style.display = 'flex';
    box.style.flexDirection = 'column';
    box.style.justifyContent = op.anchor === 'ctr' ? 'center' : (op.anchor === 'b' ? 'flex-end' : 'flex-start');
    box.style.whiteSpace = op.wrap === false ? 'nowrap' : 'pre-wrap';
    box.style.alignItems = 'stretch';
    box.style.overflow = 'visible';

    (op.paras || []).forEach(function (p) {
      var el = document.createElement('p');
      el.style.margin = '0';
      el.style.padding = '0';
      el.style.textAlign = ({ l: 'left', c: 'center', r: 'right', j: 'justify' }[p.align]) || 'left';
      if (p.line) el.style.lineHeight = String(p.line);
      if (p.marL) el.style.paddingLeft = px(p.marL);
      if (p.indent) el.style.textIndent = px(p.indent);
      if (p.before) el.style.marginTop = pt(p.before);
      if (p.after) el.style.marginBottom = pt(p.after);
      if (p.bullet) {
        el.className = 'bullet';
        el.setAttribute('data-b', p.bullet.char);
        el.style.setProperty('--bw', px(Math.abs(p.indent) || 0.3));
        el.style.setProperty('--buc', colorOf(p.bullet.color));
        el.style.setProperty('--bsz', String(p.bullet.size || 1));
      }
      var runs = (p.runs || []);
      if (!runs.length) {
        el.appendChild(document.createTextNode('\u00a0'));
      }
      runs.forEach(function (r) {
        var s = document.createElement('span');
        s.textContent = r.t;
        if (r.sz) s.style.fontSize = pt(r.sz);
        if (r.b) s.style.fontWeight = '700';
        if (r.i) s.style.fontStyle = 'italic';
        if (r.color) s.style.color = colorOf(r.color, r.alpha);
        if (r.spc) s.style.letterSpacing = pt(r.spc / 100);
        el.appendChild(s);
      });
      box.appendChild(el);
    });
    return box;
  }

  function tableEl(op) {
    var box = div('op op-table');
    applyBox(box, op);
    var total = op.colW.reduce(function (a, b) { return a + b; }, 0) || 1;
    var tbl = document.createElement('table');
    tbl.style.width = '100%';
    tbl.style.tableLayout = 'fixed';
    tbl.style.borderCollapse = 'collapse';
    var cg = document.createElement('colgroup');
    op.colW.forEach(function (w) {
      var c = document.createElement('col');
      c.style.width = (w / total * 100) + '%';
      cg.appendChild(c);
    });
    tbl.appendChild(cg);
    op.rows.forEach(function (r, ri) {
      var tr = document.createElement('tr');
      tr.style.height = px(r.h);
      r.cells.forEach(function (txt) {
        var td = document.createElement('td');
        td.textContent = txt;
        td.style.padding = '0 ' + px(0.10);
        td.style.verticalAlign = 'middle';
        td.style.fontSize = pt(op.fontSize);
        td.style.color = colorOf(r.head ? op.headText : op.text);
        if (r.head) td.style.fontWeight = '700';
        td.style.background = colorOf(r.head ? op.headFill : (ri % 2 === 0 ? op.altFill : op.bodyFill));
        td.style.border = '1px solid ' + colorOf(op.border);
        td.style.overflow = 'hidden';
        tr.appendChild(td);
      });
      tbl.appendChild(tr);
    });
    box.appendChild(tbl);
    return box;
  }

  /** 生成一张 1280x720 的幻灯片 DOM。 */
  function slideCanvas(slide, theme) {
    var th = theme || PPT.getTheme(slide.themeId);
    var root = div('slide-canvas');
    root.style.width = SLIDE_W + 'px';
    root.style.height = SLIDE_H + 'px';
    root.style.position = 'relative';
    root.style.overflow = 'hidden';
    root.style.fontFamily = fontFamily(th);
    root.style.color = hex(th.text);
    var bg = slide.bg || {};
    if (bg.grad) {
      root.style.background = 'linear-gradient(' + ((bg.angle || 0) + 90) + 'deg,' +
        hex(bg.grad[0]) + ',' + hex(bg.grad[1]) + ')';
    } else {
      root.style.background = hex(bg.fill == null ? th.bg : bg.fill);
    }
    slide.ops.forEach(function (op) {
      if (op.type === 'text') root.appendChild(textEl(op, th));
      else if (op.type === 'table') root.appendChild(tableEl(op));
      else root.appendChild(shapeEl(op));
    });
    return root;
  }

  /** 把幻灯片挂载到容器，按容器宽度等比缩放。 */
  function mount(container, slide, theme, width) {
    var w = width || 640;
    container.innerHTML = '';
    container.classList.add('slide-frame');
    container.style.width = w + 'px';
    container.style.height = Math.round(w * SLIDE_H / SLIDE_W) + 'px';
    container.style.position = 'relative';
    container.style.overflow = 'hidden';
    var inner = slideCanvas(slide, theme);
    inner.style.position = 'absolute';
    inner.style.left = '0';
    inner.style.top = '0';
    inner.style.transformOrigin = 'top left';
    inner.style.transform = 'scale(' + (w / SLIDE_W) + ')';
    container.appendChild(inner);
    return container;
  }

  PPT.preview = {
    slideCanvas: slideCanvas,
    mount: mount,
    rgba: rgba,
    hex: hex,
    SLIDE_W: SLIDE_W,
    SLIDE_H: SLIDE_H
  };

  if (typeof module !== 'undefined' && module.exports) module.exports = PPT;
})(typeof window !== 'undefined' ? window : globalThis);
