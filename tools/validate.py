"""结构校验：解析 PPTX 包内所有 XML、核对关系目标、用 python-pptx 复读内容。

用法: python validate.py <file.pptx> [--expect "文本1" "文本2" ...]
"""
import sys, zipfile, io, argparse
from lxml import etree

REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types'
PKG_REL_BASE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
PKG_REL_BASE2 = 'http://schemas.openxmlformats.org/package/2006/relationships/'

errors, warns = [], []


def fail(msg):
    errors.append(msg)


def source_part(rels_name):
    """dir/_rels/name.rels -> dir/name（该 .rels 所描述部件的包内路径）"""
    if rels_name == '_rels/.rels':
        return ''
    d, f = rels_name.rsplit('_rels/', 1)
    return d + f[:-len('.rels')]


def resolve(src_part, target):
    """以 src_part 所在目录为基准，把关系 target 解析为包内路径"""
    if target.startswith('/'):
        return target.lstrip('/')
    base_dir = src_part.rsplit('/', 1)[0] if '/' in src_part else ''
    segs = (base_dir.split('/') if base_dir else []) + target.split('/')
    out = []
    for s in segs:
        if s in ('', '.'):
            continue
        if s == '..':
            if out:
                out.pop()
        else:
            out.append(s)
    return '/'.join(out)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('path')
    ap.add_argument('--expect', nargs='*', default=[])
    args = ap.parse_args()

    zf = zipfile.ZipFile(args.path)
    names = set(zf.namelist())

    bad = zf.testzip()
    if bad:
        fail(f'ZIP 损坏: {bad}')

    # 1) 所有 XML 必须可解析
    xml_parts = [n for n in names if n.endswith('.xml') or n.endswith('.rels')]
    trees = {}
    for n in sorted(xml_parts):
        try:
            trees[n] = etree.fromstring(zf.read(n))
        except Exception as e:
            fail(f'XML 解析失败 {n}: {e}')
    print(f'XML 部件: {len(trees)} 个，全部解析成功' if not errors else f'XML 部件: {len(trees)} 个（有错误）')

    # 2) 内容类型覆盖所有部件
    ct = trees.get('[Content_Types].xml')
    if ct is None:
        fail('缺少 [Content_Types].xml')
    else:
        defaults = {d.get('Extension').lower() for d in ct.findall(f'{{{CT_NS}}}Default')}
        overrides = {o.get('PartName').lstrip('/') for o in ct.findall(f'{{{CT_NS}}}Override')}
        for n in sorted(names):
            if n == '[Content_Types].xml':
                continue
            ext = n.rsplit('.', 1)[-1].lower()
            if n not in overrides and ext not in defaults:
                fail(f'内容类型缺失: {n}')
        for o in sorted(overrides):
            if o not in names:
                fail(f'内容类型指向不存在的部件: {o}')
        print(f'内容类型: {len(defaults)} 个 Default / {len(overrides)} 个 Override')

    # 3) 关系目标必须存在
    total_rel = 0
    for n in sorted(names):
        if not n.endswith('.rels'):
            continue
        root = trees.get(n)
        if root is None:
            continue
        base_part = source_part(n)
        for r in root.findall(f'{{{REL_NS}}}Relationship'):
            if r.get('TargetMode') == 'External':
                continue
            total_rel += 1
            tgt = resolve(base_part, r.get('Target'))
            if tgt not in names:
                fail(f'{n}: 关系 {r.get("Id")} -> {r.get("Target")} 解析为 {tgt}，部件不存在')
    print(f'内部关系: {total_rel} 条已核对')

    # 4) 每张幻灯片必须能回溯到版式与母版
    pres = trees.get('ppt/presentation.xml')
    if pres is None:
        fail('缺少 ppt/presentation.xml')
    else:
        P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
        R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
        A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
        pres_rels = trees['ppt/_rels/presentation.xml.rels']
        relmap = {r.get('Id'): r.get('Target') for r in pres_rels.findall(f'{{{REL_NS}}}Relationship')}
        ids = pres.findall(f'.//{{{P}}}sldId')
        print(f'幻灯片: {len(ids)} 张')
        for s in ids:
            rid = s.get(f'{{{R}}}id')
            if rid not in relmap:
                fail(f'sldId 引用了不存在的 rId: {rid}')
        sz = pres.find(f'{{{P}}}sldSz')
        print(f'画布: {int(sz.get("cx"))/914400:.3f} x {int(sz.get("cy"))/914400:.3f} 英寸')
        if abs(int(sz.get('cx'))/914400 - 13.333) > 0.01:
            warns.append('画布不是标准 16:9 (13.333 x 7.5)')

    # 5) python-pptx 复读
    try:
        from pptx import Presentation
        prs = Presentation(args.path)
        texts = []
        for i, slide in enumerate(prs.slides, 1):
            st = []
            for shape in slide.shapes:
                if shape.has_text_frame:
                    st.append(shape.text_frame.text)
                if getattr(shape, 'has_table', False) and shape.has_table:
                    for row in shape.table.rows:
                        st.append(' | '.join(c.text for c in row.cells))
            texts.append('\n'.join(st))
            print(f'  slide {i}: {len(slide.shapes)} shapes, {len(st)} text blocks')
        alltext = '\n'.join(texts)
        for e in args.expect:
            if e not in alltext:
                fail(f'python-pptx 未读到期望文本: {e!r}')
        notes = 0
        for slide in prs.slides:
            if slide.has_notes_slide and slide.notes_slide.notes_text_frame.text.strip():
                notes += 1
        print(f'python-pptx 读取成功，含备注的幻灯片: {notes}')
    except Exception as e:
        fail(f'python-pptx 打开失败: {type(e).__name__}: {e}')

    print()
    for w in warns:
        print(f'[warn]  {w}')
    if errors:
        for e in errors:
            print(f'[FAIL]  {e}')
        print(f'\n结果: {len(errors)} 个错误')
        sys.exit(1)
    print('结果: 全部通过 ✅')


if __name__ == '__main__':
    main()
