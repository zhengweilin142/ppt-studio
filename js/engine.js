/* ============================================================
 * PPT Studio — 内容引擎（规则驱动，完全离线）
 * 两种输入都能生成结构完整的演示文稿：
 *   1) 有素材：解析 Markdown 大纲 / 要点列表 / 连续段落，自动分页
 *   2) 只有主题：按场景模板搭出可直接填写的骨架
 * ============================================================ */
(function (G) {
  'use strict';

  var PPT = (G.PPT = G.PPT || {});

  /* ---------------- 小工具 ---------------- */
  function pad2(n) { return (n < 10 ? '0' : '') + n; }

  function isCJK(s) {
    var cjk = (String(s).match(/[\u4e00-\u9fff\u3040-\u30ff]/g) || []).length;
    var all = String(s).replace(/\s/g, '').length || 1;
    return cjk / all > 0.15;
  }

  /** 从一句话里抽一个像标题的短句 */
  function deriveTitle(line, max) {
    max = max || 18;
    var t = String(line).replace(/^[-*+•\s]+/, '').replace(/^\d+[.、)]\s*/, '').trim();
    var m = t.match(/^(.{2,}?)[，。；：:、,;!?！？(（]/);
    if (m && m[1].length <= max) return m[1];
    if (t.length <= max) return t;
    return t.slice(0, max) + '…';
  }

  function today() {
    var d = new Date();
    return d.getFullYear() + '-' + pad2(d.getMonth() + 1) + '-' + pad2(d.getDate());
  }

  var NUM_RE = /[+\-¥$]?\s?\d[\d,.]*\s?(%|％|万|亿|k|K|w|倍|人|天|小时|分钟|元|美元)?/;

  /* ============================================================
   * 1. 大纲解析
   * 支持：Markdown 标题(#/##)、要点(- * 1.)、引用(>)、表格(|a|b|)、
   *      指标(k: v)、时间轴(时间 - 描述)、@指令强制版式
   * ============================================================ */
  function parseOutline(text) {
    var rawLines = String(text).replace(/\r/g, '').split('\n');
    var blocks = [];        // 顶层块
    var cur = null;         // 当前块
    var pendingLayout = null;

    function flush() { if (cur) { blocks.push(cur); cur = null; } }
    function ensure(type) {
      if (!cur || cur.type !== type) {
        flush();
        cur = { type: type, lines: [], layout: pendingLayout, level: 0, title: '' };
        pendingLayout = null;
      }
      return cur;
    }

    for (var i = 0; i < rawLines.length; i++) {
      var line = rawLines[i];
      var t = line.trim();

      if (!t) { flush(); continue; }
      if (/^[-=_*]{3,}$/.test(t)) { flush(); continue; }

      // @指令
      var dir = t.match(/^@(stats|kpi|table|timeline|quote|two-?col|section|bullets|agenda|cover|closing)\b\s*(.*)$/i);
      if (dir) {
        flush();
        pendingLayout = dir[1].toLowerCase().replace('two-col', 'twoCol').replace('kpi', 'stats');
        if (dir[2]) { cur = { type: 'dir', lines: [dir[2]], layout: pendingLayout, level: 0, title: dir[2] }; }
        continue;
      }

      // 标题
      var h = t.match(/^(#{1,6})\s+(.*)$/);
      if (h) {
        flush();
        cur = { type: 'head', lines: [h[2].trim()], layout: pendingLayout, level: h[1].length, title: h[2].trim() };
        pendingLayout = null;
        continue;
      }

      // 表格
      if (/^\|.*\|$/.test(t)) {
        var tb = ensure('table');
        tb.lines.push(t);
        continue;
      }

      // 引用
      if (/^>\s?/.test(t)) {
        var q = ensure('quote');
        q.lines.push(t.replace(/^>\s?/, ''));
        continue;
      }

      // 要点
      var b = line.match(/^(\s*)([-*+•]|\d+[.、)])\s+(.*)$/);
      if (b) {
        var ind = Math.floor(b[1].replace(/\t/g, '  ').length / 2);
        var bb = ensure('bullet');
        bb.level = Math.min(bb.level, 0);
        bb.lines.push({ text: b[3].trim(), indent: ind });
        continue;
      }

      // 普通段落行
      var p = ensure('para');
      p.lines.push(t);
    }
    flush();

    /* ---- 块 → 幻灯片 ---- */
    var deckTitle = '';
    var slides = [];        // {title, layout, bullets|stats|table|items|text|author, level}
    var sections = [];      // {title, slides:[]}

    function lastSlide() { return slides[slides.length - 1]; }
    function pushToSection(sl) {
      if (!sections.length) sections.push({ title: '', slides: [] });
      sections[sections.length - 1].slides.push(sl);
    }

    blocks.forEach(function (blk, bi) {
      if (blk.type === 'head') {
        if (!deckTitle && blk.level === 1 && !slides.length && sections.length <= 1) {
          deckTitle = blk.title;                 // 首个 H1 = 文稿标题
          return;
        }
        if (blk.level === 1) { sections.push({ title: blk.title, slides: [] }); return; }
        // H2+ = 内容页
        var sl = { title: blk.title, layout: blk.layout || 'bullets', bullets: [], _head: true };
        slides.push(sl); pushToSection(sl);
        return;
      }

      if (blk.type === 'table') {
        var rows = blk.lines.map(function (l) {
          return l.replace(/^\||\|$/g, '').split('|').map(function (c) { return c.trim(); });
        }).filter(function (r) {
          return !r.every(function (c) { return /^:?-{2,}:?$/.test(c) || !c; });
        });
        var sl2 = { layout: 'table', title: '数据一览', table: { headers: rows[0] || [], rows: rows.slice(1) } };
        slides.push(sl2); pushToSection(sl2);
        return;
      }

      if (blk.type === 'quote') {
        var qt = blk.lines.join(' ');
        var sl3 = { layout: 'quote', title: '观点', text: qt.replace(/^[“"']|[”"']$/g, ''), author: '' };
        var au = qt.match(/[—–-]{1,2}\s*([^—–\-]{2,20})$/);
        if (au) { sl3.author = au[1].trim(); sl3.text = qt.replace(/[—–-]{1,2}\s*[^—–\-]{2,20}$/, '').trim(); }
        slides.push(sl3); pushToSection(sl3);
        return;
      }

      if (blk.type === 'dir') {
        var dt = blk.lines[0];
        var sl4 = { layout: blk.layout || 'bullets', title: dt, bullets: [] };
        slides.push(sl4); pushToSection(sl4);
        return;
      }

      if (blk.type === 'bullet') {
        var items = blk.lines.map(function (l) { return { text: l.text, indent: l.indent }; });
        // 指标页？
        var stats = extractStats(items);
        if (blk.layout === 'stats' || (!blk.layout && stats && stats.length >= 2)) {
          var s5 = { layout: 'stats', title: '关键指标', stats: stats };
          slides.push(s5); pushToSection(s5);
          return;
        }
        // 时间轴？
        var tl = extractTimeline(items);
        if (blk.layout === 'timeline' || (!blk.layout && tl && tl.length >= 3)) {
          var s6 = { layout: 'timeline', title: '时间线', items: tl };
          slides.push(s6); pushToSection(s6);
          return;
        }
        var target = blk.layout ? { layout: blk.layout, title: '要点', bullets: [] } : null;
        if (!target) {
          // 只有紧跟在标题/指令之后的要点块才并入上一页，块与块之间必须分页
          var prevBlk = bi > 0 ? blocks[bi - 1] : null;
          var canAppend = prevBlk && ((prevBlk.type === 'head' && prevBlk.level >= 2) || prevBlk.type === 'dir');
          var prev = lastSlide();
          if (prev && canAppend && prev.layout === 'bullets') target = prev;
          else {
            target = { layout: 'bullets', title: deriveTitle(items[0].text, 20), bullets: [], _auto: true };
            slides.push(target); pushToSection(target);
            items = items.slice(1);
          }
        }
        items.forEach(function (it) {
          if (it.indent > 0 && target.bullets.length) {
            var last = target.bullets[target.bullets.length - 1];
            if (typeof last === 'string') { last = { t: last, s: [] }; target.bullets[target.bullets.length - 1] = last; }
            last.s = last.s || []; last.s.push(it.text);
          } else {
            target.bullets.push(it.text);
          }
        });
        return;
      }

      if (blk.type === 'para') {
        var lines = blk.lines;
        // 一段 = 一页；首句作标题
        var title = deriveTitle(lines[0], 20);
        var rest = lines.slice(1);
        var bodyLines = (lines[0].length > title.length + 2) ? [lines[0]].concat(rest) : rest;
        if (!bodyLines.length) bodyLines = [lines[0]];
        var s7 = { layout: 'bullets', title: title, bullets: splitClauses(bodyLines) };
        slides.push(s7); pushToSection(s7);
        return;
      }
    });

    /* H1 下若无内容页，把它本身当内容页 */
    sections = sections.filter(function (s) { return s.slides.length || s.title; });
    slides.forEach(function (s) { delete s._open; delete s._head; delete s._auto; });

    return { title: deckTitle, sections: sections, slides: slides };
  }

  /** 把长句拆成要点（不用 lookbehind，兼容 Safari） */
  function splitClauses(lines) {
    var out = [];
    lines.forEach(function (l) {
      if (l.length <= 34) { out.push(l); return; }
      var parts = l.split(/([。！？!?；;])/);
      var buf = '';
      for (var i = 0; i < parts.length; i++) {
        buf += parts[i];
        if (i % 2 === 1) { if (buf.trim()) out.push(buf.trim()); buf = ''; }
      }
      if (buf.trim()) out.push(buf.trim());
    });
    return out.slice(0, 8);
  }

  function extractStats(items) {
    var out = [];
    items.forEach(function (it) {
      var m = it.text.match(/^(.{1,20}?)\s*[:：]\s*(.+)$/);
      if (!m) return;
      var v = m[2].trim();
      if (!NUM_RE.test(v) || v.length > 18) return;
      out.push({ value: v, label: m[1].trim(), sub: '' });
    });
    return out.length >= 2 ? out.slice(0, 4) : null;
  }

  function extractTimeline(items) {
    var out = [];
    items.forEach(function (it) {
      var m = it.text.match(/^((?:\d{4}(?:[Qq][1-4]|[-/.]\d{1,2})?)|(?:[Qq][1-4])|(?:第?[一二三四五六七八九十\d]{1,3}[阶段期月周]))\s*[:：\-—|]\s*(.+)$/);
      if (!m) return;
      out.push({ when: m[1], what: m[2].trim() });
    });
    return out.length >= 3 ? out.slice(0, 5) : null;
  }

  /* ============================================================
   * 2. 场景模板（只有主题时使用）
   * ============================================================ */
  var SCENARIOS = {
    business: {
      name: '商业方案 / 项目汇报',
      sections: [
        { title: '项目背景', bullets: ['当前业务面临的核心问题是什么', '问题带来的量化影响（成本 / 效率 / 风险）', '为什么是现在：外部趋势与内部时机'] },
        { title: '目标与范围', bullets: ['本期要达成的 3 个具体目标', '明确的边界：本期不做什么', '衡量成功的标准与验收口径'] },
        { title: '方案概述', bullets: ['整体解决思路（一句话讲清）', '方案的三个关键组成部分', '与现有做法/竞品的核心差异'] },
        { title: '实施计划', bullets: ['阶段划分与关键里程碑', '所需资源：人力、预算、外部支持', '责任分工与协作机制'] },
        { title: '预期收益', bullets: ['效率提升：可量化的时间/成本节省', '业务价值：对收入或体验的正向影响', '长期价值：可沉淀的资产与能力'] },
        { title: '风险与应对', bullets: ['主要风险及发生概率', '对应的缓解措施与责任人', '最坏情况的兜底方案'] }
      ],
      closing: { title: '谢谢观看', subtitle: '欢迎讨论与指正' }
    },
    product: {
      name: '产品发布 / 介绍',
      sections: [
        { title: '用户痛点', bullets: ['目标用户是谁、在什么场景下遇到问题', '现有方案为什么不够好', '用户为此付出的代价'] },
        { title: '产品定位', bullets: ['一句话价值主张', '与竞品的差异化定位', '不服务哪些人群（保持聚焦）'] },
        { title: '核心能力', bullets: ['能力一：解决什么问题、怎么实现', '能力二：解决什么问题、怎么实现', '能力三：解决什么问题、怎么实现'] },
        { title: '体验亮点', bullets: ['上手成本：多久能完成第一次成功体验', '关键链路的体验优化点', '真实的用户反馈或数据'] },
        { title: '商业化与增长', bullets: ['定价与套餐设计', '获客渠道与转化路径', '增长目标与关键指标'] },
        { title: '下一步', bullets: ['近期迭代计划', '需要的支持与资源', '关键时间节点'] }
      ],
      closing: { title: '谢谢观看', subtitle: '期待与你一起把它做得更好' }
    },
    report: {
      name: '工作汇报 / 复盘',
      sections: [
        { title: '本期概览', bullets: ['一句话总结本期的整体表现', '3 个最关键的成果', '与目标的差距'] },
        { title: '关键进展', bullets: ['进展一：结果与数据', '进展二：结果与数据', '进展三：结果与数据'] },
        { title: '问题与复盘', bullets: ['遇到的主要问题', '根因分析（人 / 流程 / 资源）', '已经采取的纠偏动作'] },
        { title: '数据表现', bullets: ['核心指标的完成情况', '与上期/目标的对比', '异常波动的解释'] },
        { title: '下期计划', bullets: ['下期最重要的 3 件事', '需要的支持与协作', '关键时间节点'] }
      ],
      closing: { title: '汇报完毕', subtitle: '请各位领导批评指正' }
    },
    training: {
      name: '培训 / 课程',
      sections: [
        { title: '课程目标', bullets: ['学完这门课你能做到什么', '适合的听众与前置要求', '课程结构与时间安排'] },
        { title: '核心概念', bullets: ['概念一：定义与直觉解释', '概念二：定义与直觉解释', '常见误解与澄清'] },
        { title: '方法与实践', bullets: ['方法框架的整体步骤', '每一步的关键动作与要点', '一个完整的示范案例'] },
        { title: '动手练习', bullets: ['练习任务描述', '评价标准', '常见错误与提示'] },
        { title: '总结与延伸', bullets: ['本课要点回顾', '推荐的学习资料', '课后练习与答疑方式'] }
      ],
      closing: { title: '课程结束', subtitle: '欢迎随时提问交流' }
    },
    research: {
      name: '调研 / 研究综述',
      sections: [
        { title: '研究问题', bullets: ['我们想回答的核心问题', '为什么这个问题重要', '研究范围与边界'] },
        { title: '方法与数据', bullets: ['采用的研究方法', '数据来源与样本情况', '方法的局限性'] },
        { title: '主要发现', bullets: ['发现一：结论 + 支撑数据', '发现二：结论 + 支撑数据', '发现三：结论 + 支撑数据'] },
        { title: '讨论', bullets: ['发现意味着什么', '与既有研究/认知的一致与冲突', '可能的替代解释'] },
        { title: '结论与建议', bullets: ['核心结论', '对实践的建议', '后续可深入的方向'] }
      ],
      closing: { title: '谢谢观看', subtitle: '欢迎批评与讨论' }
    },
    pitch: {
      name: '融资路演 / BP',
      sections: [
        { title: '市场机会', bullets: ['市场规模与增长趋势', '未被满足的需求', '为什么现在是最好的时机'] },
        { title: '解决方案', bullets: ['我们提供什么', '它如何解决问题', '技术或模式上的壁垒'] },
        { title: '商业模式', bullets: ['怎么赚钱：定价与毛利', '单位经济模型（LTV / CAC）', '可复制性与规模化路径'] },
        { title: '进展与数据', bullets: ['关键经营数据', '标杆客户与案例', '增长速度与留存'] },
        { title: '团队', bullets: ['核心成员与互补性', '相关行业经验', '顾问与外部资源'] },
        { title: '融资计划', bullets: ['本轮融资金额与出让比例', '资金用途拆解', '未来 18 个月的里程碑'] }
      ],
      closing: { title: '谢谢观看', subtitle: '期待与您共同成长' }
    }
  };

  var SCENARIOS_EN = {
    business: { name: 'Business proposal',
      sections: [
        { title: 'Context', bullets: ['The core problem we face today', 'Its measurable impact on cost, speed and risk', 'Why now: external trends and internal timing'] },
        { title: 'Goals & Scope', bullets: ['Three concrete goals for this cycle', 'What is explicitly out of scope', 'How success will be measured'] },
        { title: 'Proposed Approach', bullets: ['The idea in one sentence', 'Three key components', 'How it differs from today’s approach'] },
        { title: 'Execution Plan', bullets: ['Phases and milestones', 'Resources required', 'Ownership and collaboration model'] },
        { title: 'Expected Impact', bullets: ['Efficiency gains in time and cost', 'Business value on revenue or experience', 'Long-term assets we build'] },
        { title: 'Risks & Mitigation', bullets: ['Key risks and likelihood', 'Mitigation and owners', 'Fallback plan'] }
      ], closing: { title: 'Thank you', subtitle: 'Questions and feedback welcome' } },
    product: { name: 'Product launch',
      sections: [
        { title: 'User Pain', bullets: ['Who the users are and when they struggle', 'Why today’s options fall short', 'What it costs them'] },
        { title: 'Positioning', bullets: ['Value proposition in one line', 'How we differ from alternatives', 'Who we deliberately do not serve'] },
        { title: 'Core Capabilities', bullets: ['Capability one: problem and how it works', 'Capability two: problem and how it works', 'Capability three: problem and how it works'] },
        { title: 'Experience Highlights', bullets: ['Time to first success', 'Key flow improvements', 'What users tell us'] },
        { title: 'Growth & Monetisation', bullets: ['Pricing and packaging', 'Acquisition channels', 'Targets and key metrics'] },
        { title: 'Next Steps', bullets: ['Near-term roadmap', 'Support needed', 'Key dates'] }
      ], closing: { title: 'Thank you', subtitle: 'Let’s build it together' } },
    report: { name: 'Status report',
      sections: [
        { title: 'Overview', bullets: ['One-line summary of the period', 'Three most important outcomes', 'Gap against target'] },
        { title: 'Key Progress', bullets: ['Workstream one: result and data', 'Workstream two: result and data', 'Workstream three: result and data'] },
        { title: 'Issues & Lessons', bullets: ['Main issues encountered', 'Root cause analysis', 'Corrective actions taken'] },
        { title: 'Metrics', bullets: ['Performance of core metrics', 'Comparison with last period', 'Explaining anomalies'] },
        { title: 'Next Period', bullets: ['Top three priorities', 'Support required', 'Key dates'] }
      ], closing: { title: 'End of report', subtitle: 'Feedback welcome' } },
    training: { name: 'Training course',
      sections: [
        { title: 'Learning Goals', bullets: ['What you will be able to do', 'Who this is for', 'Structure and timing'] },
        { title: 'Core Concepts', bullets: ['Concept one: definition and intuition', 'Concept two: definition and intuition', 'Common misconceptions'] },
        { title: 'Method & Practice', bullets: ['The framework step by step', 'Key actions at each step', 'A worked example'] },
        { title: 'Hands-on Exercise', bullets: ['The task', 'How it will be assessed', 'Common mistakes and hints'] },
        { title: 'Wrap-up', bullets: ['Key takeaways', 'Further reading', 'How to get help'] }
      ], closing: { title: 'End of course', subtitle: 'Happy to keep the conversation going' } },
    research: { name: 'Research review',
      sections: [
        { title: 'Research Question', bullets: ['The question we set out to answer', 'Why it matters', 'Scope and boundaries'] },
        { title: 'Method & Data', bullets: ['Approach used', 'Data sources and sample', 'Limitations'] },
        { title: 'Findings', bullets: ['Finding one with supporting data', 'Finding two with supporting data', 'Finding three with supporting data'] },
        { title: 'Discussion', bullets: ['What the findings imply', 'Agreement and tension with prior work', 'Alternative explanations'] },
        { title: 'Conclusion', bullets: ['Core conclusions', 'Practical recommendations', 'Directions for further work'] }
      ], closing: { title: 'Thank you', subtitle: 'Comments welcome' } },
    pitch: { name: 'Investor pitch',
      sections: [
        { title: 'Market Opportunity', bullets: ['Market size and growth', 'Unmet need', 'Why now'] },
        { title: 'Solution', bullets: ['What we offer', 'How it solves the problem', 'Defensibility'] },
        { title: 'Business Model', bullets: ['Pricing and gross margin', 'Unit economics (LTV / CAC)', 'Scalability'] },
        { title: 'Traction', bullets: ['Key operating metrics', 'Reference customers', 'Growth and retention'] },
        { title: 'Team', bullets: ['Core team and complementarity', 'Relevant experience', 'Advisors'] },
        { title: 'The Ask', bullets: ['Round size and dilution', 'Use of funds', '18-month milestones'] }
      ], closing: { title: 'Thank you', subtitle: 'Excited to build this with you' } }
  };

  function pickScenario(topic) {
    var t = String(topic || '');
    if (/融资|路演|BP|商业计划|投资人|pitch/i.test(t)) return 'pitch';
    if (/发布|产品|上线|功能|launch/i.test(t)) return 'product';
    if (/培训|课程|教学|分享|入门|教程|training|course/i.test(t)) return 'training';
    if (/调研|研究|综述|分析报告|论文|research|survey/i.test(t)) return 'research';
    if (/汇报|总结|复盘|周报|月报|季度|述职|report|review/i.test(t)) return 'report';
    return 'business';
  }

  /* ============================================================
   * 3. 组装
   * ============================================================ */
  function generate(input) {
    input = input || {};
    var raw = (input.outline || '').trim();
    var parsed = raw ? parseOutline(raw) : null;
    var en = !isCJK(input.topic || '') && !!input.topic;
    var topic = (input.topic || '').trim() || (parsed && parsed.title) || (en ? 'Untitled Presentation' : '未命名演示');

    var deck = {
      title: topic,
      subtitle: input.subtitle || (en ? 'Generated with PPT Studio' : '由 PPT Studio 自动生成'),
      author: input.author || '',
      org: input.org || '',
      date: input.date || today(),
      theme: input.theme || 'ocean',
      closing: { title: en ? 'Thank you' : '谢谢观看', subtitle: en ? 'Questions and feedback welcome' : '欢迎讨论与指正' }
    };

    var slides = [{ layout: 'cover' }];

    var hasParsed = parsed && parsed.slides && parsed.slides.length >= 1;
    if (hasParsed) {
      var secs = parsed.sections && parsed.sections.length > 1
        ? parsed.sections.filter(function (s) { return s.slides.length; })
        : [];
      if (secs.length >= 2) {
        slides.push({ layout: 'agenda', kicker: en ? 'CONTENTS' : '目录', title: en ? 'Agenda' : '目录',
          items: secs.map(function (s) { return { title: s.title || deriveTitle(s.slides[0].title || '', 18) }; }) });
        var n = 0;
        secs.forEach(function (s) {
          var title = s.title || deriveTitle((s.slides[0] && s.slides[0].title) || '', 18);
          if (s.title) slides.push({ layout: 'section', index: pad2(++n), title: title, subtitle: '' });
          s.slides.forEach(function (sl) { if (sl.layout !== 'cover' && sl.layout !== 'closing') slides.push(sl); });
        });
      } else {
        // 扁平大纲（只有 ## 没有 #）：用各内容页标题组成目录，不插章节页
        var titled = parsed.slides.filter(function (sl) {
          var lay = sl.layout || 'bullets';
          return sl.title && ['bullets', 'twoCol', 'stats', 'table', 'timeline'].indexOf(lay) >= 0;
        });
        var uniq = [];
        titled.forEach(function (sl) { if (uniq.indexOf(sl.title) < 0) uniq.push(sl.title); });
        if (uniq.length >= 3) {
          slides.push({
            layout: 'agenda', kicker: en ? 'CONTENTS' : '目录', title: en ? 'Agenda' : '目录',
            items: uniq.slice(0, 6).map(function (t) { return { title: t }; })
          });
        }
        parsed.slides.forEach(function (sl) { slides.push(sl); });
      }
      // 若解析出的首页标题与主题重复，去掉重复感
      if (!deck.subtitle) deck.subtitle = '';
    } else {
      var rules = en ? SCENARIOS_EN : SCENARIOS;
      var key = SCENARIOS[input.scenario] ? input.scenario : pickScenario(topic);
      var sc = rules[key] || rules.business;
      if (sc.closing) deck.closing = sc.closing;
      var limit = Math.max(3, Math.min(9, parseInt(input.slideCount, 10) || sc.sections.length));
      var chosen = sc.sections.slice(0, limit);
      slides.push({ layout: 'agenda', kicker: en ? 'CONTENTS' : '目录', title: en ? 'Agenda' : '目录',
        items: chosen.map(function (s) { return { title: s.title }; }) });
      chosen.forEach(function (sec, i) {
        slides.push({ layout: 'section', index: pad2(i + 1), title: sec.title, subtitle: '' });
        slides.push({ layout: 'bullets', kicker: topic, title: sec.title, bullets: sec.bullets.slice(),
          notes: en ? 'Fill in the specifics for this section.' : '请针对本节补充具体内容与数据。' });
      });
    }

    if (input.closing !== false) slides.push({ layout: 'closing' });
    deck.slides = slides;
    return deck;
  }

  var DEMO = [
    '# 企业知识管理平台立项汇报',
    '',
    '## 项目背景',
    '- 知识散落在 7 个以上系统中，检索平均耗时 18 分钟',
    '- 核心经验高度依赖个人，人员流动即造成资产流失',
    '  - 46% 的员工表示难以找到所需资料',
    '- 合规与审计要求持续收紧',
    '',
    '## 目标与范围',
    '- 本期目标：检索耗时下降 60%，问答准确率 ≥ 85%',
    '- 不做什么：暂不覆盖外部客户与合作伙伴',
    '',
    '@stats',
    '- 检索耗时: -62%',
    '- 知识复用率: +38%',
    '- 年度工时节省: 420万元',
    '',
    '@table',
    '| 阶段 | 时间 | 关键交付 | 预算 |',
    '| --- | --- | --- | --- |',
    '| P0 立项 | 2025 Q3 | 技术选型与试点场景确认 | 80 万 |',
    '| P1 试点 | 2025 Q4 | 知识库接入与问答上线 | 160 万 |',
    '| P2 推广 | 2026 Q1 | 全公司开放 | 240 万 |',
    '',
    '@timeline',
    '- 2025 Q3: 完成技术选型，锁定两个试点场景',
    '- 2025 Q4: 接入三个核心知识源',
    '- 2026 Q1: 全员开放并打通权限体系',
    '',
    '> 知识的价值不在于被存储，而在于被需要的那一刻可以被立刻找到。—— 项目组'
  ].join('\n');

  PPT.engine = {
    generate: generate,
    parseOutline: parseOutline,
    pickScenario: pickScenario,
    scenarios: SCENARIOS,
    demoOutline: DEMO,
    today: today,
    isCJK: isCJK
  };

  if (typeof module !== 'undefined' && module.exports) module.exports = PPT;
})(typeof window !== 'undefined' ? window : globalThis);
