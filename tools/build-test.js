/* 构建期测试：Node 下生成一份覆盖全部版式的 .pptx 供校验 */
'use strict';
const path = require('path');
const fs = require('fs');

const js = p => path.join(__dirname, '..', 'js', p);
require(js('themes.js'));
const PPT = require(js('layout.js'));
require(js('zip.js'));
require(js('pptx.js'));

const deck = {
  title: 'AI 驱动的企业知识管理平台',
  subtitle: '2025 产品战略与落地路线',
  author: '张伟',
  org: '产品战略部',
  date: '2025-06-18',
  theme: process.argv[2] || 'ocean',
  closing: { title: '谢谢观看', subtitle: '期待与您共同推进下一步' },
  slides: [
    { layout: 'cover', notes: '开场：30 秒讲清楚我们为什么做这件事。' },
    { layout: 'agenda', kicker: 'CONTENTS', title: '目录', items: [
      { title: '市场与机会' }, { title: '产品方案' }, { title: '技术架构' },
      { title: '落地路线图' }, { title: '投入与收益' }, { title: '风险与对策' }] },
    { layout: 'section', index: '01', title: '市场与机会', subtitle: '企业知识资产正在快速膨胀，而检索与复用效率仍然低下' },
    { layout: 'bullets', kicker: 'MARKET', title: '行业现状与核心痛点', bullets: [
      { t: '知识碎片化：文档散落在 7 个以上系统中，检索平均耗时 18 分钟', s: ['46% 的员工表示难以找到需要的资料', '重复撰写造成约 12% 的工时浪费'] },
      { t: '经验难沉淀：核心知识高度依赖个人，人员流动即造成资产流失' },
      { t: '合规压力上升：数据分级、权限与审计要求日趋严格' },
      { t: '现有工具割裂：搜索、问答、协作三套系统彼此不通' }],
      notes: '强调 18 分钟这个数字来自我们 2024 年对 32 家客户的访谈。' },
    { layout: 'twoCol', kicker: 'SOLUTION', title: '方案对比：外采 vs 自建', left: {
        title: '外采成熟产品', bullets: [
          { t: '上线快，2 周可用' }, { t: '按席位付费，前期成本低' },
          { t: '定制能力弱，难接入内部权限体系' }] }, right: {
        title: '基于大模型自建', bullets: [
          { t: '与内部权限、数据分级天然打通' },
          { t: '可针对业务语料持续微调' },
          { t: '需要 3 个月建设周期与专职团队' }] },
      notes: '结论：核心场景自建，长尾场景外采。' },
    { layout: 'stats', kicker: 'IMPACT', title: '预期业务收益', stats: [
      { value: '-62%', label: '检索耗时', sub: '18 分钟 → 6.8 分钟' },
      { value: '+38%', label: '知识复用率', sub: '季度环比' },
      { value: '¥420万', label: '年度工时节省', sub: '按 3000 人规模测算' },
      { value: '4.2分', label: '目标满意度', sub: '5 分制，季度调研' }] },
    { layout: 'table', kicker: 'ROADMAP', title: '关键里程碑与资源投入', table: {
        headers: ['阶段', '时间', '关键交付', '人力投入', '预算'],
        rows: [
          ['P0 立项', '2025 Q3', '技术选型与试点场景确认', '4 人', '80 万'],
          ['P1 试点', '2025 Q4', '知识库接入 + 问答上线', '8 人', '160 万'],
          ['P2 推广', '2026 Q1', '全公司开放，权限体系打通', '12 人', '240 万'],
          ['P3 深化', '2026 Q2', '垂直场景智能体与效果度量', '10 人', '200 万']] } },
    { layout: 'timeline', kicker: 'PLAN', title: '四步走落地路线', items: [
      { when: '2025 Q3', what: '完成技术选型，锁定 2 个高频试点场景' },
      { when: '2025 Q4', what: '接入 3 个核心知识源，问答准确率 ≥ 85%' },
      { when: '2026 Q1', what: '全员开放，打通权限与数据分级' },
      { when: '2026 Q2', what: '沉淀 10 个业务智能体，形成度量闭环' }] },
    { layout: 'quote', text: '知识的价值不在于被存储，而在于被需要的那一刻可以被立刻找到。', author: '—— 项目组共识' },
    { layout: 'bullets', kicker: 'RISK', title: '风险与应对', bullets: [
      { t: '数据安全风险', s: ['应对：私有化部署 + 全链路审计日志'] },
      { t: '模型效果不达预期', s: ['应对：建立 200 条标准评测集，按周回归'] },
      { t: '员工使用意愿不足', s: ['应对：嵌入现有工作流，而非新增入口'] }] },
    { layout: 'closing' }
  ]
};

PPT.pptx.build(deck).then(bytes => {
  const out = path.join(__dirname, '..', 'dist', 'test-' + deck.theme + '.pptx');
  fs.mkdirSync(path.dirname(out), { recursive: true });
  fs.writeFileSync(out, Buffer.from(bytes));
  const parts = PPT.pptx.partsFor(deck).parts;
  console.log('OK  ' + out);
  console.log('size  ' + (bytes.length / 1024).toFixed(1) + ' KB, parts ' + parts.length);
  console.log(parts.map(p => '   - ' + p.name + '  (' + (typeof p.data === 'string' ? p.data.length : p.data.length) + ')').join('\n'));
}).catch(e => { console.error('FAIL', e); process.exit(1); });
