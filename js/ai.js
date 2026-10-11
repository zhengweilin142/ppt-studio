/* ============================================================
 * PPT Studio — 可选的大模型接入（OpenAI 兼容接口）
 * 设计取舍：AI 只负责产出「Markdown 大纲」，版式仍由本地引擎渲染。
 * 好处是模型输出永远不会破坏排版，而且生成的大纲用户能直接改。
 * ============================================================ */
(function (G) {
  'use strict';

  var PPT = (G.PPT = G.PPT || {});

  var SYS = [
    '你是一位资深的演示文稿（PPT）内容架构师，擅长把零散想法整理成逻辑严密、有说服力的汇报大纲。',
    '请根据用户给出的主题，输出一份可以直接生成幻灯片的 Markdown 大纲。',
    '',
    '严格遵守以下格式规则：',
    '1. 只输出大纲正文。不要任何开场白、解释、总结或 ``` 代码块围栏。',
    '2. 第一行用 `# 主标题` 写整个文稿的标题，只出现一次。',
    '3. 用 `## 小节标题` 表示每一页内容页，每个小节 3-5 条要点；要点以 `- ` 开头，二级要点缩进两个空格后同样以 `- ` 开头。',
    '4. 需要突出关键数据时，单独占一行写 `@stats`，随后每行写成 `- 指标名: 数值`（2-4 条）。',
    '5. 需要表格时，单独占一行写 `@table`，随后写标准 Markdown 表格（含表头与分隔行）。',
    '6. 需要时间线时，单独占一行写 `@timeline`，随后每行写成 `- 2025 Q3: 描述`（3-5 条）。',
    '7. 需要强调观点时，写成 `> 一句话 —— 出处`。',
    '8. 语言与用户主题保持一致；若主题是中文，全文用中文。',
    '9. 内容要具体：给出数字、判断和可执行的动作，避免「提升效率」「加强管理」这类空话。缺少真实数据时用 xx% 之类的占位符，且不要编造具体机构名或人名。'
  ].join('\n');

  function prompt(opts) {
    var L = [];
    L.push('主题：' + (opts.topic || '未指定'));
    if (opts.scenarioName) L.push('场景类型：' + opts.scenarioName);
    if (opts.audience) L.push('听众：' + opts.audience);
    if (opts.slideCount) L.push('期望页数：约 ' + opts.slideCount + ' 页内容页（不含封面与结尾）');
    if (opts.extra) L.push('补充要求：' + opts.extra);
    if (opts.outline && opts.outline.trim()) {
      L.push('');
      L.push('以下是用户提供的原始素材，请以此为主要依据，不要偏离：');
      L.push('"""');
      L.push(opts.outline.trim().slice(0, 6000));
      L.push('"""');
    }
    L.push('');
    L.push('请直接输出 Markdown 大纲。');
    return L.join('\n');
  }

  function friendlyError(status, text) {
    var t = String(text || '').slice(0, 300);
    if (status === 401) return 'API Key 无效或已过期（401）。请检查设置中的 Key。';
    if (status === 403) return '该 Key 没有访问此模型的权限（403）。';
    if (status === 404) return '接口地址或模型名不正确（404）。请检查 Base URL 是否以 /v1 结尾、模型名是否存在。';
    if (status === 429) return '触发限流或余额不足（429）。请稍后重试或检查账户余额。';
    if (status >= 500) return '服务端错误（' + status + '）。请稍后重试。';
    return '请求失败（' + status + '）：' + t;
  }

  function call(messages, settings, dryRun) {
    var s = settings || {};
    var key = (s.apiKey || '').trim();
    if (!key) return Promise.reject(new Error('尚未填写 API Key。点击右上角「AI 设置」填入后即可使用；不填也可以直接用本地模板引擎生成。'));
    var base = (s.baseUrl || 'https://api.deepseek.com/v1').trim().replace(/\/+$/, '');
    if (!/^https?:\/\//i.test(base)) return Promise.reject(new Error('Base URL 必须以 http:// 或 https:// 开头。'));

    var body = {
      model: s.model || 'deepseek-chat',
      temperature: dryRun ? 0 : 0.6,
      max_tokens: dryRun ? 8 : 3000,
      stream: false,
      messages: messages
    };

    return fetch(base + '/chat/completions', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + key },
      body: JSON.stringify(body)
    }).then(function (res) {
      if (!res.ok) {
        return res.text().then(function (t) { throw new Error(friendlyError(res.status, t)); });
      }
      return res.json();
    }).then(function (j) {
      var msg = j && j.choices && j.choices[0] && j.choices[0].message;
      var content = msg && (msg.content || msg.reasoning_content);
      if (!content) throw new Error('接口返回为空，可能是模型名不受支持。');
      return content;
    }).catch(function (e) {
      if (e && e.name === 'TypeError') {
        // fetch 本身失败：网络不可达，或被浏览器 CORS 拦截
        throw new Error('无法连接到 ' + base + '。可能是网络不通，或该接口未开放浏览器跨域（CORS）。\n' +
          '解决办法：换用支持 CORS 的接口地址，或把本站放到本地服务器上访问。\n' +
          '你也可以不使用 AI —— 本地模板引擎（「生成演示文稿」按钮）完全离线可用。');
      }
      throw e;
    });
  }

  function clean(out) {
    var t = String(out || '').trim();
    t = t.replace(/^```[a-zA-Z]*\s*\n?/, '').replace(/\n?```\s*$/, '');
    return t.trim();
  }

  /** 根据主题生成 Markdown 大纲文本 */
  function outline(opts) {
    return call([
      { role: 'system', content: SYS },
      { role: 'user', content: prompt(opts) }
    ], opts.settings).then(clean);
  }

  /** 连通性自检 */
  function test(settings) {
    return call([{ role: 'user', content: 'ping' }], settings, true).then(function () { return true; });
  }

  PPT.ai = { outline: outline, test: test, prompt: prompt, system: SYS };

  if (typeof module !== 'undefined' && module.exports) module.exports = PPT;
})(typeof window !== 'undefined' ? window : globalThis);
