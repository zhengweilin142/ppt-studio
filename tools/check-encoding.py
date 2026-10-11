"""仓库文本文件编码体检：确保没有 BOM、没有二次编码乱码、都是干净的 UTF-8。

背景：在 Windows PowerShell 5.1（活动代码页 936）里用
    (Get-Content -Raw f) -replace ... | Set-Content -Encoding utf8
做文本替换，会把 UTF-8 字节先按 GBK 解码、再按 UTF-8 写回，
产生「文件是合法 UTF-8，但字符全是错的」这种最难发现的损坏。
这个脚本用来防止它再次发生。

用法: python check-encoding.py
"""
import os
import subprocess
import sys

sys.stdout.reconfigure(encoding='utf-8')
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
GIT = os.path.join(os.environ.get('LOCALAPPDATA', ''), r'Programs\MinGit\cmd\git.exe')
if not os.path.exists(GIT):
    GIT = 'git'

BINARY_EXT = {'.png', '.jpg', '.jpeg', '.gif', '.ico', '.pptx', '.zip', '.woff', '.woff2'}

# 每个文件里应当出现的正确中文（用于识破二次编码乱码）
EXPECT = {
    'README.md': '自动生成演示文稿',
    '.gitignore': '构建产物',
    '.gitattributes': '换行符',
    'index.html': '演示主题',
    'css/app.css': '界面样式',
    'js/themes.js': '深海商务',
    'js/layout.js': '布局引擎',
    'js/zip.js': '零依赖 ZIP',
    'js/pptx.js': 'OOXML',
    'js/engine.js': '内容引擎',
    'js/ai.js': '大模型',
    'js/preview.js': '预览渲染',
    'js/app.js': '应用控制器',
    'tools/smoke.js': '冒烟测试',
    'tools/ui-test.js': '界面无头测试',
    'tools/bundle.js': '单文件',
    'tools/validate.py': '结构校验',
    'tools/check-encoding.py': '编码体检',
}

# 二次编码的典型特征字符（UTF-8 被当成 GBK 读出来的结果）
MOJIBAKE = '鑱姹夌煭锛鈥鍜鐨鏂囩鐢熸垚'


def tracked_files():
    out = subprocess.run([GIT, '-C', ROOT, 'ls-files'], capture_output=True)
    return out.stdout.decode('utf-8', 'replace').split()


def main():
    files = tracked_files()
    if not files:
        print('没有取到文件列表，确认这里是 git 仓库')
        return 1

    problems = []
    print(f"{'文件':<26}{'BOM':<7}{'UTF-8':<8}中文校验")
    print('-' * 60)
    for f in files:
        path = os.path.join(ROOT, f)
        if not os.path.exists(path):
            continue
        ext = os.path.splitext(f)[1].lower()
        if ext in BINARY_EXT:
            continue
        raw = open(path, 'rb').read()
        bom = raw.startswith(b'\xef\xbb\xbf')
        if bom:
            problems.append((f, '带 UTF-8 BOM'))
        try:
            txt = raw.decode('utf-8')
            dec = 'ok'
        except UnicodeDecodeError as e:
            problems.append((f, f'不是合法 UTF-8: {e}'))
            print(f'{f:<26}{str(bom):<7}{"FAIL":<8}-')
            continue
        hit = [c for c in MOJIBAKE if c in txt]
        if hit:
            problems.append((f, '疑似二次编码乱码: ' + ''.join(hit[:6])))
        mark = '-'
        if f in EXPECT:
            mark = 'ok' if EXPECT[f] in txt else 'FAIL'
            if mark == 'FAIL':
                problems.append((f, f'找不到预期中文「{EXPECT[f]}」'))
        print(f'{f:<26}{str(bom):<7}{dec:<8}{mark}')

    print()
    if problems:
        print(f'发现 {len(problems)} 个问题：')
        for f, why in problems:
            print(f'  - {f}: {why}')
        print('\n修复提示：不要用 PowerShell 的 Get-Content/Set-Content 做文本往返，')
        print('改用带显式编码的 .NET 方法，例如：')
        print("  [IO.File]::WriteAllText($p, [IO.File]::ReadAllText($p, [Text.UTF8Encoding]::new($false)), [Text.UTF8Encoding]::new($false))")
        return 1
    print('全部文本文件编码正常 ✅')
    return 0


if __name__ == '__main__':
    sys.exit(main())
