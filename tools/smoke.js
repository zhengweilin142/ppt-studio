/* 冒烟测试：内容引擎各分支 + 全主题导出，纯 Node 运行 */
'use strict';
const path = require('path');
const fs = require('fs');
const js = p => path.join(__dirname, '..', 'js', p);
require(js('themes.js'));
const PPT = require(js('layout.js'));
require(js('zip.js'));
require(js('pptx.js'));
require(js('engine.js'));

let fail = 0, pass = 0;
function ok(cond, msg) {
  if (cond) { pass++; } else { fail++; console.log('  [FAIL] ' + msg); }
}
function section(t) { console.log('\n== ' + t + ' =='); }

/* ---------- 1. 大纲解析 ---------- */
section('大纲解析');
const cases = {
  'Markdown 大纲': PPT.engine.demoOutline,
  '纯要点列表': '- 第一点内容\n- 第二点内容\n- 第三点内容\n\n- 另起一块的第一点\n- 另起一块的第二点',
  '纯段落': '这是第一段的内容，讲了一件事情的来龙去脉。\n它还有第二句话。\n\n这是第二段，讲另一件事。',
  'H1+H2 结构': '# 总标题\n## 第一节\n- 要点 A\n- 要点 B\n## 第二节\n- 要点 C\n\n# 第二部分\n## 第三节\n- 要点 D',
  '指令与表格': '@stats\n- 收入: +38%\n- 成本: -12%\n\n@table\n| 名称 | 值 |\n| --- | --- |\n| A | 1 |\n\n@timeline\n- 2025 Q1: 起步\n- 2025 Q2: 扩张\n- 2025 Q3: 收敛\n\n> 一句话观点 —— 某位智者',
  '空输入': ''
};
Object.keys(cases).forEach(name => {
  let r;
  try { r = PPT.engine.parseOutline(cases[name]); }
  catch (e) { ok(false, `${name} 抛异常: ${e.message}`); return; }
  ok(Array.isArray(r.slides), name + ' 返回结构正确');
  console.log(`  ${name}: ${r.slides.length} 页, 标题="${r.title}", 章节=${r.sections.length}`);
});

/* ---------- 2. 场景生成 ---------- */
section('场景模板生成');
const scenarios = Object.keys(PPT.engine.scenarios);
scenarios.forEach(sc => {
  let deck;
  try { deck = PPT.engine.generate({ topic: '测试主题', scenario: sc, slideCount: 6 }); }
  catch (e) { ok(false, `${sc} 生成抛异常: ${e.message}`); return; }
  ok(deck.slides.length >= 8, `${sc} 页数合理 (${deck.slides.length})`);
  ok(deck.slides[0].layout === 'cover', sc + ' 首页是封面');
  ok(deck.slides[deck.slides.length - 1].layout === 'closing', sc + ' 末页是结尾');
  let plan;
  try { plan = PPT.layout.build(deck); } catch (e) { ok(false, `${sc} 布局抛异常: ${e.message}`); return; }
  ok(plan.slides.length === deck.slides.length, sc + ' 布局页数一致');
  console.log(`  ${sc}: ${deck.slides.length} 页, 版式 [${deck.slides.map(s => s.layout).join(',')}]`);
});

/* ---------- 3. 英文主题 ---------- */
section('英文主题');
const en = PPT.engine.generate({ topic: 'Quarterly Business Review', scenario: 'report', slideCount: 5 });
ok(en.slides.length >= 8, '英文场景页数合理');
ok(en.slides.some(s => (s.bullets || []).some(b => typeof b === 'string' && /[A-Za-z]/.test(b))), '英文要点已生成');
console.log(`  英文: ${en.slides.length} 页，标题="${en.title}"`);

/* ---------- 4. 全主题布局 + 边界检查 ---------- */
section('全主题布局与坐标边界');
const W = PPT.layout.W, H = PPT.layout.H;
PPT.themeList.forEach(t => {
  const deck = PPT.engine.generate({ topic: '全主题检查', outline: PPT.engine.demoOutline, theme: t.id });
  const plan = PPT.layout.build(deck);
  let bad = [], count = 0;
  plan.slides.forEach((sl, si) => {
    sl.ops.forEach(op => {
      count++;
      ['x', 'y', 'w', 'h'].forEach(k => {
        if (typeof op[k] !== 'number' || !isFinite(op[k])) bad.push(`slide${si + 1}.${op.type}.${k}=${op[k]}`);
      });
      if (op.type === 'text') {
        if (op.x < -0.01 || op.x + op.w > W + 0.01) bad.push(`slide${si + 1} 文本横向越界 x=${op.x} w=${op.w}`);
        if (op.y + op.h > H + 0.6) bad.push(`slide${si + 1} 文本纵向越界 y=${op.y} h=${op.h}`);
      }
      if (op.type === 'text') {
        op.paras.forEach(p => (p.runs || []).forEach(r => {
          if (r.color != null && typeof r.color === 'string' && !/^[0-9A-Fa-f]{6}$/.test(r.color)) {
            bad.push(`slide${si + 1} 颜色非法: ${r.color}`);
          }
        }));
      }
    });
  });
  ok(bad.length === 0, `${t.id} 绘制指令合法` + (bad.length ? ' → ' + bad.slice(0, 3).join('; ') : ''));
  console.log(`  ${t.name}: ${plan.slides.length} 页 / ${count} 个绘制指令`);
});

/* ---------- 5. 各主题导出 pptx ---------- */
section('各主题导出 .pptx');
const outDir = path.join(__dirname, '..', 'dist', 'themes');
fs.mkdirSync(outDir, { recursive: true });
const jobs = PPT.themeList.map(t => {
  const deck = PPT.engine.generate({ topic: 'AI 驱动的企业知识管理平台', subtitle: '全主题导出检查', author: '测试', outline: PPT.engine.demoOutline, theme: t.id });
  deck.date = '2025-06-18';
  return PPT.pptx.build(deck).then(bytes => {
    ok(bytes.length > 8000, `${t.id} 字节数合理 (${bytes.length})`);
    fs.writeFileSync(path.join(outDir, 'theme-' + t.id + '.pptx'), Buffer.from(bytes));
    console.log(`  ${t.name}: ${(bytes.length / 1024).toFixed(1)} KB`);
  }).catch(e => ok(false, `${t.id} 导出失败: ${e.message}`));
});

Promise.all(jobs).then(() => {
  console.log(`\n通过 ${pass} 项，失败 ${fail} 项`);
  process.exit(fail ? 1 : 0);
});
