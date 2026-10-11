/* ============================================================
 * PPT Studio — 应用控制器
 * 状态 → 布局 → 预览 → 导出，所有交互都在这里收口。
 * ============================================================ */
(function (G) {
  'use strict';

  var PPT = G.PPT;
  var LS_KEY = 'ppt-studio-v1';

  var LAYOUTS = [
    ['cover', '封面'], ['agenda', '目录'], ['section', '章节页'], ['bullets', '要点页'],
    ['twoCol', '两栏对比'], ['stats', '数据卡片'], ['table', '表格'], ['timeline', '时间轴'],
    ['quote', '金句'], ['closing', '结尾页']
  ];
  var LAY_MAP = {};
  LAYOUTS.forEach(function (p) { LAY_MAP[p[0]] = p[1]; });

  var ZOOMS = [200, 260, 320, 400, 520];

  var state = {
    deck: null, plan: null, current: 0, view: 'grid', zoom: 2, presentAt: 0,
    settings: { baseUrl: 'https://api.deepseek.com/v1', apiKey: '', model: 'deepseek-chat' },
    input: { topic: '', subtitle: '', author: '', org: '', date: '', outline: '', scenario: 'auto', slideCount: 6 }
  };

  function $(s) { return document.querySelector(s); }
  function el(tag, cls, txt) {
    var n = document.createElement(tag);
    if (cls) n.className = cls;
    if (txt != null) n.textContent = txt;
    return n;
  }
  function pad2(n) { return (n < 10 ? '0' : '') + n; }

  var toastTimer = null;
  function toast(msg, kind) {
    var t = $('#toast');
    t.textContent = msg;
    t.className = 'toast on' + (kind ? ' ' + kind : '');
    clearTimeout(toastTimer);
    toastTimer = setTimeout(function () { t.className = 'toast'; }, kind === 'err' ? 7000 : 3200);
  }
  function tip(msg, kind) {
    var t = $('#tip');
    t.textContent = msg;
    t.className = 'tip' + (kind ? ' ' + kind : '');
  }

  function safeName(s) {
    return String(s || '演示文稿').replace(/[\\/:*?"<>|\r\n\t]/g, '_').slice(0, 60) || '演示文稿';
  }
  function download(bytes, name, mime) {
    var blob = new Blob([bytes], { type: mime || 'application/octet-stream' });
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url; a.download = name;
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    setTimeout(function () { URL.revokeObjectURL(url); }, 5000);
  }

  /* ============================================================
   * 持久化
   * ============================================================ */
  var saveTimer = null;
  function save() {
    clearTimeout(saveTimer);
    saveTimer = setTimeout(function () {
      try {
        localStorage.setItem(LS_KEY, JSON.stringify({
          deck: state.deck, input: state.input, settings: state.settings, view: state.view, current: state.current
        }));
      } catch (e) { /* 隐私模式下可能不可用，忽略 */ }
    }, 400);
  }
  function load() {
    try {
      var raw = localStorage.getItem(LS_KEY);
      if (!raw) return false;
      var o = JSON.parse(raw);
      if (o.settings) state.settings = Object.assign(state.settings, o.settings);
      if (o.input) state.input = Object.assign(state.input, o.input);
      if (o.view) state.view = o.view;
      if (o.current) state.current = o.current;
      if (o.deck && o.deck.slides && o.deck.slides.length) { state.deck = o.deck; return true; }
    } catch (e) { /* 数据损坏则重新开始 */ }
    return false;
  }

  /* ============================================================
   * 渲染
   * ============================================================ */
  function rebuild(opts) {
    opts = opts || {};
    state.plan = PPT.layout.build(state.deck);
    if (state.current >= state.deck.slides.length) state.current = Math.max(0, state.deck.slides.length - 1);
    renderStage();
    renderMeta();
    if (!opts.keepInspector) renderInspector();
    syncFormFromState();
    save();
  }

  var stageTimer = null;
  function scheduleStage() {
    clearTimeout(stageTimer);
    stageTimer = setTimeout(function () {
      state.plan = PPT.layout.build(state.deck);
      renderStage();
      renderMeta();
      save();
    }, 130);
  }

  function renderMeta() {
    var th = PPT.getTheme(state.deck.theme);
    $('#stageMeta').textContent = '共 ' + state.deck.slides.length + ' 页 · 主题：' + th.name + ' · 16:9 宽屏';
  }

  function renderStage() {
    var host = $('#stageBody');
    var th = state.plan.theme;
    host.innerHTML = '';
    if (state.view === 'single') {
      var wrap = el('div', 'single');
      var avail = Math.max(360, Math.min(1180, host.clientWidth - 40));
      var frame = el('div');
      wrap.appendChild(frame);
      var nav = el('div', 'single-nav');
      var prev = el('button', null, '‹ 上一页');
      var info = el('span', null, (state.current + 1) + ' / ' + state.deck.slides.length + '　' + (LAY_MAP[state.deck.slides[state.current].layout] || ''));
      var next = el('button', null, '下一页 ›');
      prev.onclick = function () { go(-1); };
      next.onclick = function () { go(1); };
      nav.appendChild(prev); nav.appendChild(info); nav.appendChild(next);
      wrap.appendChild(nav);
      host.appendChild(wrap);
      PPT.preview.mount(frame, state.plan.slides[state.current], th, avail);
      return;
    }

    // 网格
    var w = ZOOMS[state.zoom];
    var grid = el('div', 'grid');
    grid.style.gridTemplateColumns = 'repeat(auto-fill, ' + w + 'px)';
    grid.style.gap = '30px 18px';
    state.plan.slides.forEach(function (sl, i) {
      var thumb = el('div', 'thumb' + (i === state.current ? ' on' : ''));
      thumb.appendChild(el('div', 'idx', pad2(i + 1)));
      thumb.appendChild(el('div', 'lay', LAY_MAP[sl.layout] || sl.layout));
      var frame = el('div');
      thumb.appendChild(frame);
      var ops = el('div', 'ops');
      [['up', '↑', '上移'], ['down', '↓', '下移'], ['dup', '⧉', '复制'], ['del', '✕', '删除']].forEach(function (o) {
        var b = el('button', null, o[1]);
        b.title = o[2];
        b.onclick = function (ev) { ev.stopPropagation(); slideOp(o[0], i); };
        ops.appendChild(b);
      });
      thumb.appendChild(ops);
      thumb.onclick = function () {
        if (state.current === i) {
          state.view = 'single'; syncViewSeg(); renderStage();
        } else {
          state.current = i; renderStage(); renderInspector(); save();
        }
      };
      grid.appendChild(thumb);
      PPT.preview.mount(frame, sl, th, w);
    });
    host.appendChild(grid);
  }

  function syncViewSeg() {
    Array.prototype.forEach.call($('#viewSeg').children, function (b) {
      b.classList.toggle('on', b.dataset.view === state.view);
    });
  }

  function go(d) {
    var n = state.deck.slides.length;
    state.current = (state.current + d + n) % n;
    renderStage(); renderInspector(); save();
  }

  function slideOp(act, i) {
    var s = state.deck.slides;
    if (act === 'up' && i > 0) { var a = s[i - 1]; s[i - 1] = s[i]; s[i] = a; state.current = i - 1; }
    else if (act === 'down' && i < s.length - 1) { var b = s[i + 1]; s[i + 1] = s[i]; s[i] = b; state.current = i + 1; }
    else if (act === 'dup') { s.splice(i + 1, 0, JSON.parse(JSON.stringify(s[i]))); state.current = i + 1; }
    else if (act === 'del') {
      if (s.length <= 1) { toast('至少要保留一页'); return; }
      s.splice(i, 1); state.current = Math.min(i, s.length - 1);
    }
    rebuild({ keepInspector: false });
  }

  /* ============================================================
   * 右侧检查器
   * ============================================================ */
  function mkInput(labelTxt, value, oninput, placeholder, tag) {
    var f = el('div', 'field');
    f.appendChild(el('label', null, labelTxt));
    var n = document.createElement(tag || 'input');
    if (tag === 'textarea') { n.rows = 5; n.spellcheck = false; }
    n.value = value == null ? '' : value;
    if (placeholder) n.placeholder = placeholder;
    n.addEventListener('input', function () { oninput(n.value); });
    f.appendChild(n);
    return f;
  }

  function bulletsToText(bullets) {
    var out = [];
    PPT.layout.normBullets(bullets).forEach(function (b) {
      out.push(b.t);
      b.s.forEach(function (s) { out.push('  ' + s); });
    });
    return out.join('\n');
  }
  function textToBullets(text) {
    var out = [];
    String(text).split('\n').forEach(function (line) {
      if (!line.trim()) return;
      var isSub = /^\s{2,}|^\t/.test(line);
      var t = line.trim().replace(/^[-*•·]\s*/, '');
      if (isSub && out.length) {
        var last = out[out.length - 1];
        if (typeof last === 'string') { last = { t: last, s: [] }; out[out.length - 1] = last; }
        last.s.push(t);
      } else out.push(t);
    });
    return out;
  }
  function statsToText(stats) {
    return (stats || []).map(function (s) {
      return [s.value, s.label, s.sub].filter(function (x) { return x != null && x !== ''; }).join(' | ');
    }).join('\n');
  }
  function textToStats(t) {
    return String(t).split('\n').filter(function (l) { return l.trim(); }).map(function (l) {
      var p = l.split('|').map(function (x) { return x.trim(); });
      return { value: p[0] || '', label: p[1] || '', sub: p.slice(2).join(' ').trim() };
    }).slice(0, 8);
  }
  function tableToText(t) {
    t = t || {};
    if (!t.headers || !t.headers.length) return '';
    var rows = [t.headers].concat(t.rows || []);
    var out = ['| ' + t.headers.join(' | ') + ' |', '|' + t.headers.map(function () { return ' --- '; }).join('|') + '|'];
    (t.rows || []).forEach(function (r) { out.push('| ' + r.join(' | ') + ' |'); });
    return out.join('\n');
  }
  function textToTable(txt) {
    var rows = String(txt).split('\n').map(function (l) { return l.trim(); })
      .filter(function (l) { return /^\|.*\|$/.test(l); })
      .map(function (l) { return l.replace(/^\||\|$/g, '').split('|').map(function (c) { return c.trim(); }); })
      .filter(function (r) { return !r.every(function (c) { return /^:?-{2,}:?$/.test(c) || !c; }); });
    return { headers: rows[0] || [], rows: rows.slice(1) };
  }
  function timelineToText(items) {
    return (items || []).map(function (i) { return (i.when || '') + ' | ' + (i.what || ''); }).join('\n');
  }
  function textToTimeline(txt) {
    return String(txt).split('\n').filter(function (l) { return l.trim(); }).map(function (l) {
      var p = l.split(/[|:：]/);
      return { when: (p[0] || '').trim(), what: p.slice(1).join(' ').trim() };
    }).slice(0, 5);
  }
  function linesToItems(txt) {
    return String(txt).split('\n').filter(function (l) { return l.trim(); }).map(function (l) { return { title: l.trim() }; });
  }

  function renderInspector() {
    var host = $('#panelRight');
    host.innerHTML = '';
    var s = state.deck.slides[state.current];
    if (!s) { host.appendChild(el('div', 'empty', '还没有幻灯片')); return; }
    var wrap = el('div', 'insp');

    var head = el('div', 'insp-head');
    head.appendChild(el('h3', null, '第 ' + (state.current + 1) + ' 页 · 编辑'));
    head.appendChild(el('span', 'tag', LAY_MAP[s.layout] || s.layout));
    wrap.appendChild(head);

    // 版式切换
    var laySec = el('div', 'insp-sec');
    laySec.appendChild(el('label', null, '版式'));
    var sel = document.createElement('select');
    LAYOUTS.forEach(function (p) {
      var o = document.createElement('option');
      o.value = p[0]; o.textContent = p[1];
      if (p[0] === s.layout) o.selected = true;
      sel.appendChild(o);
    });
    sel.onchange = function () {
      s.layout = sel.value;
      if (s.layout === 'closing') s.title = s.title || '';
      rebuild();
    };
    laySec.appendChild(sel);
    wrap.appendChild(laySec);

    var body = el('div', 'insp-sec');
    var L = s.layout;

    if (L === 'cover') {
      body.appendChild(el('p', 'hint', '封面文字取自左侧的「演示主题 / 副标题 / 汇报人 / 机构 / 日期」，修改那几项即可实时更新。'));
    } else if (L === 'closing') {
      state.deck.closing = state.deck.closing || {};
      body.appendChild(mkInput('大标题', state.deck.closing.title, function (v) { state.deck.closing.title = v; scheduleStage(); }));
      body.appendChild(mkInput('副标题', state.deck.closing.subtitle, function (v) { state.deck.closing.subtitle = v; scheduleStage(); }));
    } else if (L === 'section') {
      body.appendChild(mkInput('章节序号', s.index, function (v) { s.index = v; scheduleStage(); }, '01'));
      body.appendChild(mkInput('章节标题', s.title, function (v) { s.title = v; scheduleStage(); }));
      body.appendChild(mkInput('章节说明', s.subtitle, function (v) { s.subtitle = v; scheduleStage(); }, '选填'));
    } else if (L === 'agenda') {
      body.appendChild(mkInput('页标题', s.title, function (v) { s.title = v; scheduleStage(); }));
      body.appendChild(mkInput('目录条目（每行一条）', (s.items || []).map(function (i) { return i.title || i; }).join('\n'),
        function (v) { s.items = linesToItems(v); scheduleStage(); }, '', 'textarea'));
    } else if (L === 'bullets') {
      body.appendChild(mkInput('小标签', s.kicker, function (v) { s.kicker = v; scheduleStage(); }, '选填，如 MARKET'));
      body.appendChild(mkInput('页标题', s.title, function (v) { s.title = v; scheduleStage(); }));
      body.appendChild(mkInput('要点（每行一条；行首缩进两格 = 二级要点）', bulletsToText(s.bullets),
        function (v) { s.bullets = textToBullets(v); scheduleStage(); }, '', 'textarea'));
    } else if (L === 'twoCol') {
      s.left = s.left || { title: '', bullets: [] };
      s.right = s.right || { title: '', bullets: [] };
      body.appendChild(mkInput('页标题', s.title, function (v) { s.title = v; scheduleStage(); }));
      body.appendChild(mkInput('左栏标题', s.left.title, function (v) { s.left.title = v; scheduleStage(); }));
      body.appendChild(mkInput('左栏要点', bulletsToText(s.left.bullets), function (v) { s.left.bullets = textToBullets(v); scheduleStage(); }, '', 'textarea'));
      body.appendChild(mkInput('右栏标题', s.right.title, function (v) { s.right.title = v; scheduleStage(); }));
      body.appendChild(mkInput('右栏要点', bulletsToText(s.right.bullets), function (v) { s.right.bullets = textToBullets(v); scheduleStage(); }, '', 'textarea'));
    } else if (L === 'stats') {
      body.appendChild(mkInput('页标题', s.title, function (v) { s.title = v; scheduleStage(); }));
      body.appendChild(mkInput('指标（每行：数值 | 名称 | 说明）', statsToText(s.stats),
        function (v) { s.stats = textToStats(v); scheduleStage(); }, '如：+38% | 效率提升 | 季度环比', 'textarea'));
    } else if (L === 'table') {
      body.appendChild(mkInput('页标题', s.title, function (v) { s.title = v; scheduleStage(); }));
      body.appendChild(mkInput('表格（Markdown 格式）', tableToText(s.table),
        function (v) { s.table = textToTable(v); scheduleStage(); }, '', 'textarea'));
    } else if (L === 'timeline') {
      body.appendChild(mkInput('页标题', s.title, function (v) { s.title = v; scheduleStage(); }));
      body.appendChild(mkInput('节点（每行：时间 | 描述）', timelineToText(s.items),
        function (v) { s.items = textToTimeline(v); scheduleStage(); }, '如：2025 Q3 | 完成技术选型', 'textarea'));
    } else if (L === 'quote') {
      body.appendChild(mkInput('金句内容', s.text, function (v) { s.text = v; scheduleStage(); }, '', 'textarea'));
      body.appendChild(mkInput('出处', s.author, function (v) { s.author = v; scheduleStage(); }, '选填'));
    }
    wrap.appendChild(body);

    // 备注
    var noteSec = el('div', 'insp-sec');
    noteSec.appendChild(mkInput('演讲者备注（会写入 .pptx 的备注页）', s.notes,
      function (v) { s.notes = v; save(); }, '选填', 'textarea'));
    wrap.appendChild(noteSec);

    // 幻灯片操作
    var opSec = el('div', 'insp-sec');
    opSec.appendChild(el('label', null, '幻灯片操作'));
    var r1 = el('div', 'btnrow');
    [['↑ 上移', 'up'], ['↓ 下移', 'down'], ['⧉ 复制', 'dup'], ['✕ 删除', 'del']].forEach(function (p) {
      var b = el('button', 'btn ghost sm', p[0]);
      b.onclick = function () { slideOp(p[1], state.current); };
      r1.appendChild(b);
    });
    opSec.appendChild(r1);
    var r2 = el('div', 'btnrow');
    var addSel = document.createElement('select');
    LAYOUTS.forEach(function (p) {
      var o = document.createElement('option');
      o.value = p[0]; o.textContent = '插入：' + p[1];
      addSel.appendChild(o);
    });
    addSel.value = 'bullets';
    var addBtn = el('button', 'btn primary sm', '＋ 新增页');
    addBtn.onclick = function () {
      state.deck.slides.splice(state.current + 1, 0, blankSlide(addSel.value));
      state.current++;
      rebuild();
    };
    r2.appendChild(addSel); r2.appendChild(addBtn);
    opSec.appendChild(r2);
    wrap.appendChild(opSec);

    host.appendChild(wrap);
  }

  function blankSlide(layout) {
    var t = state.deck.title || '';
    switch (layout) {
      case 'cover': return { layout: 'cover' };
      case 'closing': return { layout: 'closing' };
      case 'section': return { layout: 'section', index: pad2(state.deck.slides.length), title: '新章节', subtitle: '' };
      case 'agenda': return { layout: 'agenda', title: '目录', items: [{ title: '第一部分' }, { title: '第二部分' }] };
      case 'twoCol': return { layout: 'twoCol', title: '对比分析', left: { title: '方案 A', bullets: ['要点一', '要点二'] }, right: { title: '方案 B', bullets: ['要点一', '要点二'] } };
      case 'stats': return { layout: 'stats', title: '关键数据', stats: [{ value: '+38%', label: '指标名称', sub: '口径说明' }, { value: '1.2万', label: '指标名称', sub: '' }] };
      case 'table': return { layout: 'table', title: '数据一览', table: { headers: ['列一', '列二'], rows: [['内容', '内容']] } };
      case 'timeline': return { layout: 'timeline', title: '时间线', items: [{ when: '第一阶段', what: '描述' }, { when: '第二阶段', what: '描述' }, { when: '第三阶段', what: '描述' }] };
      case 'quote': return { layout: 'quote', text: '在这里写一句有力量的话。', author: '' };
      default: return { layout: 'bullets', title: '新页面', bullets: ['要点一', '要点二', '要点三'] };
    }
  }

  /* ============================================================
   * 表单同步
   * ============================================================ */
  function syncFormFromState() {
    $('#inTopic').value = state.input.topic || '';
    $('#inSubtitle').value = state.input.subtitle || '';
    $('#inAuthor').value = state.input.author || '';
    $('#inOrg').value = state.input.org || '';
    $('#inDate').value = state.input.date || '';
    $('#inOutline').value = state.input.outline || '';
    $('#inScenario').value = state.input.scenario || 'auto';
    $('#inCount').value = state.input.slideCount || 6;
  }
  function syncInputFromForm() {
    state.input.topic = $('#inTopic').value.trim();
    state.input.subtitle = $('#inSubtitle').value.trim();
    state.input.author = $('#inAuthor').value.trim();
    state.input.org = $('#inOrg').value.trim();
    state.input.date = $('#inDate').value.trim();
    state.input.outline = $('#inOutline').value;
    state.input.scenario = $('#inScenario').value;
    state.input.slideCount = parseInt($('#inCount').value, 10) || 6;
  }

  function generateLocal(silent) {
    syncInputFromForm();
    // 主题不在表单里，但重新生成时必须保留用户当前选的配色
    state.input.theme = (state.deck && state.deck.theme) || 'ocean';
    state.deck = PPT.engine.generate(state.input);
    state.current = 0;
    rebuild();
    if (!silent) {
      toast('已生成 ' + state.deck.slides.length + ' 页，可在右侧逐页微调', 'ok');
    }
  }

  /* ============================================================
   * 主题选择
   * ============================================================ */
  function renderThemes() {
    var host = $('#themePicker');
    host.innerHTML = '';
    PPT.themeList.forEach(function (t) {
      var b = el('button', 'theme-dot' + (state.deck.theme === t.id ? ' on' : ''));
      b.style.background = '#' + t.bg;
      var i = el('i');
      i.style.background = '#' + t.accent;
      b.appendChild(i);
      b.title = t.name;
      b.onclick = function () {
        state.deck.theme = t.id;
        renderThemes();
        rebuild({ keepInspector: true });
        toast('已切换到「' + t.name + '」');
      };
      host.appendChild(b);
    });
  }

  /* ============================================================
   * 模态框
   * ============================================================ */
  function openModal(sel) {
    var m = $(sel);
    m.hidden = false;
  }
  function closeModals() {
    Array.prototype.forEach.call(document.querySelectorAll('.modal'), function (m) { m.hidden = true; });
  }

  function openSettings() {
    $('#setBase').value = state.settings.baseUrl || '';
    $('#setKey').value = state.settings.apiKey || '';
    $('#setModel').value = state.settings.model || '';
    $('#setStatus').textContent = '';
    openModal('#modalSettings');
  }
  function saveSettings() {
    state.settings.baseUrl = $('#setBase').value.trim() || 'https://api.deepseek.com/v1';
    state.settings.apiKey = $('#setKey').value.trim();
    state.settings.model = $('#setModel').value.trim() || 'deepseek-chat';
    save();
    closeModals();
    toast('AI 设置已保存到本机', 'ok');
  }

  /* ============================================================
   * 放映
   * ============================================================ */
  function renderPresent() {
    var host = $('#presentStage');
    var host2 = host;
    var w = Math.min(window.innerWidth - 80, Math.round((window.innerHeight - 150) * 16 / 9));
    w = Math.max(480, w);
    PPT.preview.mount(host2, state.plan.slides[state.presentAt], state.plan.theme, w);
    $('#pCount').textContent = (state.presentAt + 1) + ' / ' + state.deck.slides.length;
  }
  function openPresent() {
    state.presentAt = state.current;
    $('#present').hidden = false;
    renderPresent();
  }
  function closePresent() { $('#present').hidden = true; }
  function presentGo(d) {
    var n = state.deck.slides.length;
    state.presentAt = (state.presentAt + d + n) % n;
    renderPresent();
  }

  /* ============================================================
   * 事件绑定
   * ============================================================ */
  function bind() {
    // 主题信息（实时改写 deck，不重新生成幻灯片）
    $('#inTopic').addEventListener('input', function () {
      state.input.topic = this.value;
      state.deck.title = this.value || '未命名演示';
      scheduleStage(); save();
    });
    $('#inSubtitle').addEventListener('input', function () {
      state.input.subtitle = this.value; state.deck.subtitle = this.value; scheduleStage(); save();
    });
    $('#inAuthor').addEventListener('input', function () {
      state.input.author = this.value; state.deck.author = this.value; scheduleStage(); save();
    });
    $('#inOrg').addEventListener('input', function () {
      state.input.org = this.value; state.deck.org = this.value; scheduleStage(); save();
    });
    $('#inDate').addEventListener('input', function () {
      state.input.date = this.value; state.deck.date = this.value; scheduleStage(); save();
    });
    $('#inOutline').addEventListener('input', function () { state.input.outline = this.value; save(); });
    $('#inScenario').addEventListener('change', function () { state.input.scenario = this.value; save(); });
    $('#inCount').addEventListener('change', function () { state.input.slideCount = parseInt(this.value, 10) || 6; save(); });

    $('#btnGen').onclick = function () { generateLocal(); };

    $('#btnAI').onclick = function () {
      syncInputFromForm();
      if (!state.input.topic && !state.input.outline) { toast('请先填写演示主题', 'err'); $('#inTopic').focus(); return; }
      if (!state.settings.apiKey) { openSettings(); toast('请先填写 API Key；不填也可以用「生成演示文稿」离线生成', 'err'); return; }
      var btn = this;
      btn.disabled = true;
      var old = btn.textContent;
      btn.textContent = 'AI 生成中…';
      tip('正在请求大模型生成大纲…', 'warn');
      var scn = state.input.scenario !== 'auto' && PPT.engine.scenarios[state.input.scenario]
        ? PPT.engine.scenarios[state.input.scenario].name : '';
      PPT.ai.outline({
        topic: state.input.topic, outline: state.input.outline, scenarioName: scn,
        slideCount: state.input.slideCount, settings: state.settings
      }).then(function (md) {
        state.input.outline = md;
        $('#inOutline').value = md;
        generateLocal(true);
        tip('AI 大纲已写入左侧文本框，可直接修改后再点「生成演示文稿」。', 'ok');
        toast('AI 大纲已生成，共 ' + state.deck.slides.length + ' 页', 'ok');
      }).catch(function (e) {
        tip(e.message, 'err');
        toast('AI 生成失败，详见左下角提示', 'err');
      }).then(function () {
        btn.disabled = false;
        btn.textContent = old;
      });
    };

    $('#btnDemo').onclick = function () {
      var d = PPT.engine.generate({
        topic: 'AI 驱动的企业知识管理平台',
        subtitle: '2025 产品战略与落地路线',
        author: '张伟', org: '产品战略部', date: '',
        outline: PPT.engine.demoOutline, theme: state.deck ? state.deck.theme : 'ocean'
      });
      state.deck = d;
      state.input = {
        topic: d.title, subtitle: d.subtitle, author: d.author, org: d.org, date: d.date,
        outline: PPT.engine.demoOutline, scenario: 'auto', slideCount: 6
      };
      state.current = 0;
      rebuild();
      toast('已载入示例大纲', 'ok');
    };

    $('#btnClear').onclick = function () {
      state.input = { topic: '', subtitle: '', author: '', org: '', date: '', outline: '', scenario: 'auto', slideCount: 6 };
      syncFormFromState();
      $('#inTopic').focus();
      toast('已清空输入');
    };

    $('#btnHelp').onclick = function () { openModal('#modalHelp'); };
    $('#btnSettings').onclick = openSettings;
    $('#btnSaveSettings').onclick = saveSettings;

    $('#btnTest').onclick = function () {
      var st = $('#setStatus');
      var btn = this;
      btn.disabled = true;
      st.textContent = '正在测试…';
      PPT.ai.test({
        baseUrl: $('#setBase').value.trim(), apiKey: $('#setKey').value.trim(), model: $('#setModel').value.trim()
      }).then(function () {
        st.textContent = '✅ 连接成功，可以使用 AI 生成。';
      }).catch(function (e) {
        st.textContent = '❌ ' + e.message;
      }).then(function () { btn.disabled = false; });
    };

    $('#btnJson').onclick = function () {
      $('#jsonArea').value = JSON.stringify(state.deck, null, 2);
      $('#jsonStatus').textContent = '';
      openModal('#modalJson');
    };
    $('#btnJsonApply').onclick = function () {
      try {
        var o = JSON.parse($('#jsonArea').value);
        if (!o.slides || !o.slides.length) throw new Error('缺少 slides 数组');
        state.deck = o;
        state.current = 0;
        closeModals();
        rebuild();
        toast('已应用 JSON', 'ok');
      } catch (e) {
        $('#jsonStatus').textContent = '❌ JSON 解析失败：' + e.message;
      }
    };
    $('#btnJsonCopy').onclick = function () {
      var a = $('#jsonArea');
      a.select();
      try { document.execCommand('copy'); $('#jsonStatus').textContent = '✅ 已复制到剪贴板'; }
      catch (e) { $('#jsonStatus').textContent = '请手动复制（Ctrl+C）'; }
    };

    $('#btnExport').onclick = function () {
      var btn = this;
      btn.disabled = true;
      var old = btn.textContent;
      btn.textContent = '正在打包…';
      PPT.pptx.build(state.deck).then(function (bytes) {
        download(bytes, safeName(state.deck.title) + '.pptx',
          'application/vnd.openxmlformats-officedocument.presentationml.presentation');
        toast('已导出 ' + (bytes.length / 1024).toFixed(0) + ' KB 的 .pptx 文件', 'ok');
      }).catch(function (e) {
        console.error(e);
        toast('导出失败：' + e.message, 'err');
      }).then(function () {
        btn.disabled = false;
        btn.textContent = old;
      });
    };

    $('#btnPresent').onclick = openPresent;
    $('#pPrev').onclick = function () { presentGo(-1); };
    $('#pNext').onclick = function () { presentGo(1); };
    $('#pExit').onclick = closePresent;

    Array.prototype.forEach.call($('#viewSeg').children, function (b) {
      b.onclick = function () {
        state.view = b.dataset.view;
        syncViewSeg();
        renderStage();
        save();
      };
    });

    $('#zoomIn').onclick = function () { state.zoom = Math.min(ZOOMS.length - 1, state.zoom + 1); updateZoom(); };
    $('#zoomOut').onclick = function () { state.zoom = Math.max(0, state.zoom - 1); updateZoom(); };

    Array.prototype.forEach.call(document.querySelectorAll('.modal'), function (m) {
      m.addEventListener('mousedown', function (e) { if (e.target === m) closeModals(); });
      Array.prototype.forEach.call(m.querySelectorAll('[data-close]'), function (b) { b.onclick = closeModals; });
    });

    document.addEventListener('keydown', function (e) {
      if (!$('#present').hidden) {
        if (e.key === 'ArrowRight' || e.key === ' ' || e.key === 'PageDown') { e.preventDefault(); presentGo(1); }
        else if (e.key === 'ArrowLeft' || e.key === 'PageUp') { e.preventDefault(); presentGo(-1); }
        else if (e.key === 'Escape') closePresent();
        return;
      }
      if (e.key === 'Escape') { closeModals(); return; }
      var tag = (e.target.tagName || '').toLowerCase();
      if (tag === 'input' || tag === 'textarea' || tag === 'select') return;
      if (e.key === 'ArrowRight') go(1);
      else if (e.key === 'ArrowLeft') go(-1);
    });

    window.addEventListener('resize', function () {
      if (!$('#present').hidden) renderPresent();
    });
  }

  function updateZoom() {
    $('#zoomVal').textContent = Math.round(ZOOMS[state.zoom] / 1280 * 100) + '%';
    renderStage();
    save();
  }

  /* ============================================================
   * 启动
   * ============================================================ */
  function boot() {
    // 场景下拉
    var sc = $('#inScenario');
    var o0 = document.createElement('option');
    o0.value = 'auto'; o0.textContent = '自动判断';
    sc.appendChild(o0);
    Object.keys(PPT.engine.scenarios).forEach(function (k) {
      var o = document.createElement('option');
      o.value = k; o.textContent = PPT.engine.scenarios[k].name;
      sc.appendChild(o);
    });

    var restored = load();
    if (!restored) {
      state.deck = PPT.engine.generate({
        topic: 'AI 驱动的企业知识管理平台',
        subtitle: '2025 产品战略与落地路线',
        author: '张伟', org: '产品战略部',
        outline: PPT.engine.demoOutline, theme: 'ocean'
      });
      state.input = {
        topic: 'AI 驱动的企业知识管理平台', subtitle: '2025 产品战略与落地路线',
        author: '张伟', org: '产品战略部', date: '',
        outline: PPT.engine.demoOutline, scenario: 'auto', slideCount: 6
      };
    }
    if (!state.deck.date) state.deck.date = state.input.date || '';
    if (!state.deck.date) state.deck.date = PPT.engine.today();

    bind();
    syncViewSeg();
    updateZoomLabel();
    renderThemes();
    rebuild();

    if (!state.settings.apiKey) {
      tip('开箱即用：不填任何 API Key，本地模板引擎也能生成并导出 .pptx。想让内容更聪明，可点右上角「AI 设置」。');
    }
  }
  function updateZoomLabel() {
    $('#zoomVal').textContent = Math.round(ZOOMS[state.zoom] / 1280 * 100) + '%';
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', boot);
  else boot();

  G.PPTApp = { state: state, rebuild: rebuild, toast: toast, generateLocal: generateLocal };
})(typeof window !== 'undefined' ? window : globalThis);
