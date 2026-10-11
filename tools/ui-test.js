/* ============================================================
 * 界面无头测试
 * 浏览器在沙箱里起不来，所以这里实现一个最小 DOM 垫片，
 * 按 index.html 的真实结构把整个前端跑一遍：
 *   启动 → 渲染预览 → 切主题 → 重新生成 → 编辑 → 删除 → 放映 → 导出
 * 导出走的是和「导出 .pptx」按钮完全相同的代码路径。
 * ============================================================ */
'use strict';
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const root = path.join(__dirname, '..');

let fail = 0, pass = 0;
function ok(cond, msg) {
  if (cond) { pass++; console.log('  ✓ ' + msg); }
  else { fail++; console.log('  ✗ ' + msg); }
}
function section(t) { console.log('\n== ' + t + ' =='); }

/* ---------------- 极简 HTML 解析 ---------------- */
const VOID = new Set(['meta', 'link', 'br', 'hr', 'img', 'input', 'source', 'col']);
function decode(s) {
  return String(s).replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"')
    .replace(/&#10;/g, '\n').replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&');
}
function parseHTML(html, makeEl) {
  html = html.replace(/<!--[\s\S]*?-->/g, '').replace(/<!DOCTYPE[^>]*>/i, '');
  const rootEl = makeEl('#document');
  const stack = [rootEl];
  const re = /<\/?([a-zA-Z][\w-]*)((?:\s+[^>]*?)?)\/?>/g;
  let last = 0, m;
  function text(t) {
    if (!t.trim()) return;
    const top = stack[stack.length - 1];
    top.appendChild(makeEl('#text', decode(t.trim())));
  }
  while ((m = re.exec(html))) {
    text(html.slice(last, m.index));
    last = re.lastIndex;
    const tag = m[1].toLowerCase();
    const closing = m[0][1] === '/';
    if (closing) {
      for (let i = stack.length - 1; i > 0; i--) {
        if (stack[i].tag === tag) { stack.length = i; break; }
      }
      continue;
    }
    const el = makeEl(tag);
    const attrs = m[2] || '';
    const ar = /([\w:-]+)(?:\s*=\s*"([^"]*)"|\s*=\s*'([^']*)')?/g;
    let a;
    while ((a = ar.exec(attrs))) {
      const name = a[1].toLowerCase();
      const val = a[2] != null ? a[2] : (a[3] != null ? a[3] : '');
      el.setAttribute(name, decode(val));
    }
    stack[stack.length - 1].appendChild(el);
    if (!VOID.has(tag)) stack.push(el);
  }
  text(html.slice(last));
  return rootEl;
}

/* ---------------- DOM 垫片 ---------------- */
function createDOM(html) {
  let uid = 0;
  function makeEl(tag, textNode) {
    const isText = tag === '#text';
    const el = {
      __id: ++uid, tagName: isText ? '#text' : tag.toUpperCase(), nodeType: isText ? 3 : 1,
      childNodes: [], parentNode: null, attributes: {}, _listeners: {}, _rawHTML: null,
      tag: isText ? '#text' : tag.toLowerCase()
    };
    if (isText) el.textContent = textNode || '';
    Object.defineProperty(el, 'children', { get: () => el.childNodes.filter(c => c.nodeType === 1) });
    Object.defineProperty(el, 'textContent', {
      get() { return el._text != null ? el._text : el.childNodes.map(c => c.textContent).join(''); },
      set(v) { el._text = String(v); el.childNodes = []; }
    });
    Object.defineProperty(el, 'innerHTML', {
      get() { return el._rawHTML != null ? el._rawHTML : ''; },
      set(v) { el.childNodes = []; el._rawHTML = String(v); el._text = String(v).replace(/<[^>]*>/g, ''); }
    });
    Object.defineProperty(el, 'className', {
      get() { return el._class || ''; },
      set(v) { el._class = String(v); }
    });
    el.classList = {
      add(...cs) { cs.forEach(c => { if (!el.classList.contains(c)) el._class = (el._class ? el._class + ' ' : '') + c; }); },
      remove(...cs) { cs.forEach(c => { el._class = (el._class || '').split(/\s+/).filter(x => x && x !== c).join(' '); }); },
      contains(c) { return (el._class || '').split(/\s+/).indexOf(c) >= 0; },
      toggle(c, force) { const has = el.classList.contains(c); const want = force === undefined ? !has : !!force; if (want) el.classList.add(c); else el.classList.remove(c); return want; }
    };
    el.style = new Proxy({ setProperty(k, v) { this[k] = v; }, getPropertyValue(k) { return this[k] || ''; } }, {
      set(t, k, v) { t[k] = v; return true; }
    });
    el.dataset = new Proxy({}, {
      get(t, k) { return el.attributes['data-' + String(k).replace(/[A-Z]/g, s => '-' + s.toLowerCase())]; },
      set(t, k, v) { el.attributes['data-' + k] = v; return true; }
    });
    el.appendChild = c => { c.parentNode = el; el.childNodes.push(c); el._text = null; return c; };
    el.removeChild = c => { const i = el.childNodes.indexOf(c); if (i >= 0) el.childNodes.splice(i, 1); return c; };
    el.remove = () => { if (el.parentNode) el.parentNode.removeChild(el); };
    el.setAttribute = (k, v) => {
      el.attributes[k] = v;
      if (k === 'class') el._class = v;
      if (k === 'id') { el.id = v; if (!doc.__byId[v]) doc.__byId[v] = el; }
      if (k === 'hidden') el.hidden = true;
      if (k.startsWith('data-')) { /* dataset proxy reads live */ }
    };
    el.getAttribute = k => (el.attributes[k] === undefined ? null : el.attributes[k]);
    el.hasAttribute = k => el.attributes[k] !== undefined;
    el.addEventListener = (t, fn) => { (el._listeners[t] = el._listeners[t] || []).push(fn); };
    el.removeEventListener = () => {};
    el.dispatch = (t, ev) => {
      const e = Object.assign({ type: t, target: el, preventDefault() {}, stopPropagation() {} }, ev || {});
      (el._listeners[t] || []).forEach(fn => fn.call(el, e));
      const h = el['on' + t];
      if (typeof h === 'function') h.call(el, e);
    };
    el.click = () => el.dispatch('click');   // 统一由 dispatch 触发，避免回调执行两次
    el.focus = () => {};
    el.select = () => {};
    el.querySelectorAll = sel => queryAll(el, sel);
    el.querySelector = sel => queryAll(el, sel)[0] || null;
    if (isText) { el.appendChild = () => {}; }
    return el;
  }

  function matchesSimple(el, sel) {
    sel = sel.trim();
    if (sel.indexOf(',') >= 0) return sel.split(',').some(s => matchesSimple(el, s));
    if (sel.startsWith('#')) return el.id === sel.slice(1);
    if (sel.startsWith('.')) return el.classList.contains(sel.slice(1));
    if (sel.startsWith('[') && sel.endsWith(']')) {
      const inner = sel.slice(1, -1);
      const eq = inner.indexOf('=');
      if (eq < 0) return el.hasAttribute(inner);
      return el.getAttribute(inner.slice(0, eq)) === inner.slice(eq + 1).replace(/^["']|["']$/g, '');
    }
    return el.tag === sel.toLowerCase();
  }
  /** 支持后代选择器：#a .b .c */
  function matches(el, sel) {
    if (el.nodeType !== 1) return false;
    const parts = String(sel).trim().split(/\s+/);
    if (!matchesSimple(el, parts[parts.length - 1])) return false;
    let node = el.parentNode;
    for (let i = parts.length - 2; i >= 0; i--) {
      while (node && node.nodeType === 1 && !matchesSimple(node, parts[i])) node = node.parentNode;
      if (!node || node.nodeType !== 1) return false;
      node = node.parentNode;
    }
    return true;
  }
  function walk(node, fn) {
    node.childNodes.forEach(c => { if (c.nodeType === 1) { fn(c); walk(c, fn); } });
  }
  function queryAll(scope, sel) {
    const out = [];
    walk(scope, el => { if (matches(el, sel)) out.push(el); });
    return out;
  }

  const doc = makeEl('#document');
  doc.__byId = {};
  const body = makeEl('body');
  doc.appendChild(body);
  const domRoot = parseHTML(html, makeEl);
  domRoot.childNodes.slice().forEach(c => { doc.appendChild(c); });
  walk(doc, el => { if (el.id) doc.__byId[el.id] = el; });
  doc.body = body;
  // body 里的内容其实在解析树中，这里把真实 body 找出来
  const realBody = queryAll(doc, 'body')[0];
  if (realBody) { doc.body = realBody; doc.childNodes = doc.childNodes.filter(c => c !== realBody); doc.appendChild(realBody); }

  doc.querySelector = sel => queryAll(doc, sel)[0] || null;
  doc.querySelectorAll = sel => queryAll(doc, sel);
  doc.getElementById = id => doc.__byId[id] || null;
  doc.createElement = tag => makeEl(tag);
  doc.createTextNode = t => makeEl('#text', t);
  doc.readyState = 'complete';
  doc.execCommand = () => true;
  doc.addEventListener = (t, fn) => { (doc._listeners[t] = doc._listeners[t] || []).push(fn); };
  doc.dispatch = (t, ev) => (doc._listeners[t] || []).forEach(fn => fn(ev || {}));
  return doc;
}

/* ---------------- 全局环境 ---------------- */
const html = fs.readFileSync(path.join(root, 'index.html'), 'utf8');
const document = createDOM(html);
const storage = {};
const captured = { blobs: [], lastBytes: null };

const win = {
  document,
  innerWidth: 1680, innerHeight: 1000,
  addEventListener: () => {},
  setTimeout, clearTimeout, setInterval, clearInterval,
  console, Promise, JSON, Math, Date, Object, Array, String, Number, isFinite, parseInt, parseFloat, Error,
  localStorage: {
    getItem: k => (k in storage ? storage[k] : null),
    setItem: (k, v) => { storage[k] = String(v); },
    removeItem: k => { delete storage[k]; }
  },
  Blob: function (parts, opts) { this.parts = parts; this.type = opts && opts.type; },
  URL: {
    createObjectURL(b) { captured.blobs.push(b); return 'blob:stub-' + captured.blobs.length; },
    revokeObjectURL() {}
  },
  fetch: () => Promise.reject(new Error('无网络（测试环境不请求接口）')),
  TextEncoder,
  CompressionStream: undefined
};
win.window = win;

global.window = win;
global.document = document;
global.localStorage = win.localStorage;
global.Blob = win.Blob;
global.URL = win.URL;
global.fetch = win.fetch;
global.TextEncoder = TextEncoder;

// 拦截下载：把 Blob 内容取出来
const origCreate = win.URL.createObjectURL;
win.URL.createObjectURL = function (blob) {
  captured.lastBlob = blob;
  return origCreate.call(win.URL, blob);
};

/* ---------------- 加载前端 ---------------- */
section('加载前端脚本（按 index.html 顺序）');
const scripts = ['themes.js', 'zip.js', 'layout.js', 'pptx.js', 'engine.js', 'ai.js', 'preview.js', 'app.js'];
let loaded = 0;
scripts.forEach(f => {
  const code = fs.readFileSync(path.join(root, 'js', f), 'utf8');
  try {
    vm.runInThisContext(code, { filename: f });
    loaded++;
  } catch (e) {
    fail++;
    console.log('  ✗ ' + f + ' 执行失败: ' + e.message);
    console.log((e.stack || '').split('\n').slice(0, 8).join('\n'));
  }
});
ok(loaded === scripts.length, `全部 ${scripts.length} 个脚本执行成功（实际 ${loaded}）`);
const PPT = win.PPT;
ok(!!PPT && !!PPT.pptx && !!PPT.preview && !!PPT.engine, 'PPT 命名空间已挂载');

/* ---------------- 启动后的界面状态 ---------------- */
section('启动后的界面状态');
const $ = s => document.querySelector(s);
const thumbs = () => document.querySelectorAll('.thumb');
ok(thumbs().length > 0, `预览区渲染出缩略图（${thumbs().length} 张）`);
ok(document.querySelectorAll('.slide-canvas').length === thumbs().length, '每张缩略图都挂了幻灯片画布');
ok(document.querySelectorAll('#themePicker .theme-dot').length === 6, '主题选择器渲染出 6 个色块');
ok(!!$('#panelRight').querySelector('.insp-head'), '右侧检查器已渲染');
ok(document.querySelectorAll('#inScenario option').length === 7, '场景下拉有 7 项（自动 + 6 场景）');
ok(document.querySelectorAll('#stageBody .op').length > 20, '幻灯片绘制指令已生成 DOM（' + document.querySelectorAll('#stageBody .op').length + ' 个图元）');
ok(document.querySelectorAll('#stageBody .bullet').length > 0, '要点项目符号已渲染');

/* ---------------- 交互 ---------------- */
section('交互：切换主题');
const dot = document.querySelectorAll('#themePicker .theme-dot')[1];
dot.click();
ok(win.PPTApp.state.deck.theme === 'midnight', '点击色块后 deck.theme 变为 midnight');
ok(thumbs().length > 0, '切换主题后预览仍然正常');

section('交互：生成演示文稿');
$('#inTopic').value = '产品发布评审';
$('#inTopic').dispatch('input');
$('#btnGen').click();
const n1 = win.PPTApp.state.deck.slides.length;
ok(n1 >= 8, `点击生成后得到 ${n1} 页`);
ok(win.PPTApp.state.deck.title === '产品发布评审', '主题文字实时生效');
ok(win.PPTApp.state.deck.theme === 'midnight', '重新生成后仍保留所选主题（回归）');

section('交互：单页视图与翻页');
document.querySelectorAll('#viewSeg button')[1].click();
ok(document.querySelectorAll('#stageBody .single').length === 1, '切换到单页视图');
const before = win.PPTApp.state.current;
document.dispatch('keydown', { key: 'ArrowRight', target: { tagName: 'BODY' } });
ok(win.PPTApp.state.current === (before + 1) % n1, '→ 键翻到下一页');
document.querySelectorAll('#viewSeg button')[0].click();

section('交互：编辑幻灯片');
win.PPTApp.state.current = 3;
win.PPTApp.rebuild();
const insp = $('#panelRight');
const tas = insp.querySelectorAll('textarea');
ok(insp.querySelectorAll('input,textarea,select').length >= 3, '检查器渲染出编辑控件');
const beforeOps = document.querySelectorAll('#stageBody .op').length;
const last = tas[tas.length - 1];
last.value = '这是备注文字';
last.dispatch('input');
win.PPTApp.state.current = 4;
win.PPTApp.rebuild();
ok(win.PPTApp.state.deck.slides[3].notes === '这是备注文字', '编辑备注写入 deck');

section('交互：幻灯片增删与排序');
const n2 = win.PPTApp.state.deck.slides.length;
win.PPTApp.state.current = 2;
win.PPTApp.rebuild();
const btns = $('#panelRight').querySelectorAll('.btn');
console.log('  · 检查器按钮：' + btns.map(b => '[' + b.textContent + ']').join(' '));
btns[0].click();   // 上移
ok(win.PPTApp.state.current === 1, '上移后焦点跟随（current=' + win.PPTApp.state.current + '）');
const btns2 = $('#panelRight').querySelectorAll('.btn');
btns2[3].click();   // 删除
ok(win.PPTApp.state.deck.slides.length === n2 - 1, `删除后页数 ${n2} → ${win.PPTApp.state.deck.slides.length}`);

section('交互：放映模式');
$('#btnPresent').click();
ok(!$('#present').hidden, '放映层已打开');
ok(document.querySelectorAll('#presentStage .slide-canvas').length === 1, '放映层渲染了当前页');
document.dispatch('keydown', { key: 'Escape' });
ok($('#present').hidden, 'Esc 退出放映');

/* ---------------- 导出 ---------------- */
section('导出 .pptx（走按钮的真实代码路径）');
const n3 = win.PPTApp.state.deck.slides.length;
$('#btnExport').click();

function blobToBuffer(blob) {
  const p = blob.parts[0];
  if (p instanceof Uint8Array) return Buffer.from(p);
  if (typeof p === 'string') return Buffer.from(p, 'utf8');
  if (p && p.buffer) return Buffer.from(p.buffer);
  return null;
}

setTimeout(() => {
  const blob = captured.lastBlob;
  ok(!!blob, '点击导出后产生了 Blob');
  if (blob) {
    const buf = blobToBuffer(blob);
    ok(!!buf && buf.length > 8000, `导出字节数合理（${buf ? buf.length : 0}）`);
    ok(blob.type.indexOf('presentationml') > 0, 'MIME 类型是 pptx');
    if (buf) {
      const out = path.join(root, 'dist', 'ui-export.pptx');
      fs.writeFileSync(out, buf);
      console.log('  → 已写出 ' + out + '（' + n3 + ' 页）');
    }
  }
  console.log(`\n通过 ${pass} 项，失败 ${fail} 项`);
  process.exit(fail ? 1 : 0);
}, 400);
