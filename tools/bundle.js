/* 把 index.html + css + js 打包成单文件版本，方便直接分发/双击打开 */
'use strict';
const fs = require('fs');
const path = require('path');
const root = path.join(__dirname, '..');

let html = fs.readFileSync(path.join(root, 'index.html'), 'utf8');
const used = [];

html = html.replace(/<link rel="stylesheet" href="([^"]+)">/g, (m, href) => {
  used.push(href);
  return '<style>\n' + fs.readFileSync(path.join(root, href), 'utf8').trim() + '\n</style>';
});

html = html.replace(/<script src="([^"]+)"><\/script>/g, (m, src) => {
  used.push(src);
  const code = fs.readFileSync(path.join(root, src), 'utf8');
  if (code.includes('</script')) throw new Error('内联脚本包含 </script，需要转义：' + src);
  return '<script>\n' + code.trim() + '\n</script>';
});

html = html.replace('<title>', '<!-- 单文件版本：无需任何外部依赖，双击即可使用 -->\n<title>');

const out = path.join(root, 'dist', 'PPT-Studio.html');
fs.mkdirSync(path.dirname(out), { recursive: true });
fs.writeFileSync(out, html, 'utf8');
console.log('已打包 ' + used.length + ' 个资源 → ' + out);
console.log('大小 ' + (Buffer.byteLength(html, 'utf8') / 1024).toFixed(1) + ' KB');
