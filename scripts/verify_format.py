# -*- coding: utf-8 -*-
"""jiaoan-skill 通用格式校验脚本
用法: python verify_format.py <教案路径> [--texts 新文本A|新文本B] [--noleak 旧文本A|旧文本B]
校验: 表格结构(表0=9行/板书4列第4列空)/行距固定22磅/字体无五号残留/无封面批注/可选新文本与残留检查
"""
import sys, os
import docx
from docx.enum.text import WD_LINE_SPACING
sys.stdout.reconfigure(encoding='utf-8', errors='replace')

def main():
    args = sys.argv[1:]
    if not args:
        print('用法: python verify_format.py <教案docx路径> [--texts 新文本1|新文本2] [--noleak 旧文本1|旧文本2]')
        sys.exit(1)
    path = args[0]
    texts = []
    noleak = []
    i = 1
    while i < len(args):
        if args[i] == '--texts':
            texts = args[i+1].split('|')
            i += 2
        elif args[i] == '--noleak':
            noleak = args[i+1].split('|')
            i += 2
        else:
            i += 1

    ok = True
    def check(cond, msg):
        nonlocal ok
        print(('✅' if cond else '❌'), msg)
        if not cond:
            ok = False

    d = docx.Document(path)
    print(f'== 校验: {path} ==')

    # 1. 表格结构
    tbs = d.tables
    check(len(tbs) >= 3, f'表格数≥3（实际{len(tbs)}）')
    tb0 = tbs[0]
    check(len(tb0.rows) == 9, f'表0行数=9行教案表（实际{len(tb0.rows)}）')
    # 板书表：最后一表为板书，列数=4且第4列空
    wb = tbs[-1]
    cols4 = len(wb.rows[0].cells)
    check(cols4 == 4, f'板书表列数=4（实际{cols4}）')
    if cols4 == 4:
        c3 = wb.rows[-1].cells[3].text.strip() if len(wb.rows) > 1 else ''
        check(c3 == '', '板书第4列为空（框架列留空）')

    # 2. 行距：固定值22磅（允许图片段单倍）
    exact = 0; single = 0; other = 0; total = 0
    for p in d.paragraphs:
        pf = p.paragraph_format
        total += 1
        if pf.line_spacing_rule == WD_LINE_SPACING.EXACTLY:
            exact += 1
        elif pf.line_spacing_rule == WD_LINE_SPACING.SINGLE:
            single += 1
        else:
            other += 1
    print(f'   段落: 总数{total} 固定行距{exact} 单倍{single}(图片段) 其他{other}')
    check(other == 0, '无未知行距段落')
    check(exact >= total - single - 2, '固定行距覆盖正文（除图片段）')

    # 3. 字体：检测五号10.5pt残留（正文/表格不应出现；页眉页脚除外）
    bad_fonts = []
    for p in d.paragraphs:
        for r in p.runs:
            sz = r.font.size
            if sz is not None and sz.pt == 10.5:
                bad_fonts.append(r.text[:15])
    for tb in tbs:
        for row in tb.rows:
            for c in row.cells:
                for p in c.paragraphs:
                    for r in p.runs:
                        sz = r.font.size
                        if sz is not None and sz.pt == 10.5:
                            bad_fonts.append(r.text[:15])
    check(len(bad_fonts) == 0, f'无五号10.5pt残留（若列出请人工确认非页眉页脚）: {bad_fonts[:5]}')

    # 4. 无封面/批注
    import zipfile
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
    check('word/comments.xml' not in names, '无批注comments.xml')
    # 首段不应是封面（封面通常含"教案""课程名称"大标题行；此处检查首个非空段不是"教案"单字）
    first_text = ''
    for p in d.paragraphs:
        if p.text.strip():
            first_text = p.text.strip()
            break
    check('教案' != first_text and '齐齐哈尔工程学院' not in first_text, f'首段非封面（首段: {first_text[:20]}）')

    # 5. 可选新/旧文本检查
    def all_text(doc):
        parts = [p.text for p in doc.paragraphs]
        for tb in doc.tables:
            for row in tb.rows:
                for c in row.cells:
                    parts.append(c.text)
        return '\n'.join(parts)
    if texts or noleak:
        full = all_text(d)
        for s in texts:
            check(s in full, f'新文本存在: {s[:30]}')
        for s in noleak:
            check(s not in full, f'旧文本无残留: {s[:30]}')

    print('== 校验', '全部通过 ✅' if ok else '存在未通过项 ❌', '==')
    sys.exit(0 if ok else 1)

if __name__ == '__main__':
    main()
