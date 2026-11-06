#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""deck 产出校验器。改完 build_deck.py 后必须跑一遍，全绿才能提交。

用法：
    python3 check_deck.py                       # 只校验自身（页码/几何/换行编码）
    python3 check_deck.py 参考.pptx              # 额外与参考文件逐页比对文案
    python3 check_deck.py 参考.pptx --strict     # 文案差异也算失败

退出码：0 = 通过；1 = 有问题（详见输出）。
"""
import sys
import zipfile
from pptx import Presentation

EMU = 12700.0
DECK = "SSW AI 业务_v3.pptx"
FOOTER_LINE_Y = 379.0   # 页脚横线的 y 坐标（pt）
PAGE_H = 405.0


def vlen(s):
    """粗略视觉宽度：中文 1.0，ASCII 0.55，空格 0.32。"""
    w = 0.0
    for ch in s:
        o = ord(ch)
        w += 0.32 if ch == ' ' else (0.55 if o < 0x2E80 else 1.0)
    return w


def page_texts(slide):
    """该页所有非页脚/非页码文本，按 (top, left) 排序后返回列表。"""
    items = []
    for sh in slide.shapes:
        if not sh.has_text_frame:
            continue
        t = sh.text_frame.text.strip()
        if not t:
            continue
        if t.isdigit() and len(t) <= 2:      # 页码
            continue
        if 'SSW AI 业务 ·' in t:              # 页脚文案
            continue
        items.append((round(sh.top / EMU, 1), round(sh.left / EMU, 1),
                      t.replace('\n', '').replace('\x0b', '')))
    items.sort()
    return [t for _, _, t in items]


def check_pageno(prs):
    """页脚右下角页码必须等于物理页序；附录为 A1/A2。"""
    bad = []
    total = len(prs.slides._sldIdLst)
    for i, s in enumerate(prs.slides, 1):
        foot = None
        for sh in s.shapes:
            if not sh.has_text_frame:
                continue
            t = sh.text_frame.text.strip()
            # 页码框是页脚带里「宽 >= 100 且文字 <= 4 字符」的那个：
            # 正文页 x=570 w=114；分隔页 x=36 w=648（右对齐）；案例导航点 w=20，据此排除
            if (sh.top / EMU > 378 and sh.width / EMU >= 100
                    and 1 <= len(t) <= 4 and not t.startswith('SSW')):
                foot = t
        # 按页眉标签判断：标签以「附录一/二」开头 → A1/A2，其余按物理页序。
        # 不依赖附录在第几页，挪页也不会失效。
        tag = ''
        for sh in s.shapes:
            if not sh.has_text_frame:
                continue
            t = sh.text_frame.text.strip()
            if sh.top / EMU < 40 and t.startswith('附录'):
                tag = t
                break
        if tag.startswith('附录一'):
            want = 'A1'
        elif tag.startswith('附录二'):
            want = 'A2'
        else:
            want = '%02d' % i
        if foot != want:
            bad.append((i, foot, want))
    return bad


def check_overflow(prs):
    """估算每个文本框是否框内溢出 / 越页脚线。"""
    bad = []
    for i, s in enumerate(prs.slides, 1):
        for sh in s.shapes:
            if not sh.has_text_frame or not sh.text_frame.text.strip():
                continue
            top = sh.top / EMU
            if top >= FOOTER_LINE_Y:
                continue
            w, h = sh.width / EMU, sh.height / EMU
            total = 0.0
            for p in sh.text_frame.paragraphs:
                txt = ''.join(r.text for r in p.runs)
                if not txt:
                    continue
                sz = next((r.font.size.pt for r in p.runs if r.font.size), 10)
                ls = 1.2
                q = p._p.find('.//{http://schemas.openxmlformats.org/drawingml/2006/main}lnSpc')
                if q is not None:
                    z = q.find('{http://schemas.openxmlformats.org/drawingml/2006/main}spcPct')
                    if z is not None:
                        ls = int(z.get('val')) / 100000.0
                cpl = max(w - 4, 20) / (sz * 0.98)
                est = vlen(txt) / cpl
                total += max(1, int(est) + (1 if est % 1 > 0.02 else 0)) * sz * ls
                sa = p._p.find('.//{http://schemas.openxmlformats.org/drawingml/2006/main}spcAft')
                if sa is not None:
                    z2 = sa.find('{http://schemas.openxmlformats.org/drawingml/2006/main}spcPts')
                    if z2 is not None:
                        total += int(z2.get('val')) / 100.0
            label = sh.text_frame.text[:24].replace('\n', '/')
            if total > h + 6:
                bad.append((i, '框内溢出', round(top), round(h), round(total), label))
            if top + total > FOOTER_LINE_Y:
                bad.append((i, '越页脚线', round(top), round(top + total), label))
    return bad


def check_newlines(path):
    """run 的 <a:t> 里不能出现裸 LF 或 \x0b —— PowerPoint 对 run 内裸换行处理不可靠。
    换行必须写成 <a:br/> 或拆成独立段落。"""
    bad = []
    z = zipfile.ZipFile(path)
    for n in z.namelist():
        if not (n.startswith('ppt/slides/slide') and n.endswith('.xml')):
            continue
        x = z.read(n).decode('utf-8')
        import re
        for m in re.finditer(r'<a:t>([^<]*)</a:t>', x):
            if '\n' in m.group(1) or '\x0b' in m.group(1):
                bad.append((n, m.group(1)[:30]))
    return bad


def check_against(prs, ref_path):
    """与参考 pptx 逐页比对文案集合（顺序无关、忽略换行编码）。"""
    ref = Presentation(ref_path)
    if len(ref.slides._sldIdLst) != len(prs.slides._sldIdLst):
        print('  ⚠ 页数不同：参考 %d 页 / 当前 %d 页'
              % (len(ref.slides._sldIdLst), len(prs.slides._sldIdLst)))
    diffs = []
    for i in range(min(len(ref.slides._sldIdLst), len(prs.slides._sldIdLst))):
        a = sorted(page_texts(prs.slides[i]))
        b = sorted(page_texts(ref.slides[i]))
        if a != b:
            only_now = [x for x in a if x not in b]
            only_ref = [x for x in b if x not in a]
            diffs.append((i + 1, only_now[:3], only_ref[:3]))
    return diffs


def main():
    ref = None
    strict = '--strict' in sys.argv
    for a in sys.argv[1:]:
        if not a.startswith('--'):
            ref = a
    prs = Presentation(DECK)
    total = len(prs.slides._sldIdLst)
    print('校验对象：%s（%d 页）\n' % (DECK, total))
    fail = 0

    bad = check_pageno(prs)
    if bad:
        fail += 1
        print('❌ 页码与物理页序不符 %d 处：' % len(bad))
        for x in bad[:15]:
            print('   第 %s 页显示 %r，应为 %r' % x)
    else:
        print('✅ 页码：%d 页全部与物理页序一致' % total)

    bad = check_overflow(prs)
    if bad:
        fail += 1
        print('❌ 版面越界 %d 处：' % len(bad))
        for x in bad[:15]:
            print('   %s' % (x,))
    else:
        print('✅ 版面：0 处溢出 / 0 处越页脚线')

    bad = check_newlines(DECK)
    if bad:
        fail += 1
        print('❌ run 内裸换行 %d 处（必须改用 <a:br/>）：' % len(bad))
        for x in bad[:10]:
            print('   %s  %r' % x)
    else:
        print('✅ 换行编码：0 处裸 LF / 软换行')

    if ref:
        diffs = check_against(prs, ref)
        if diffs:
            if strict:
                fail += 1
            print('%s 文案与参考不同 %d 页：' % ('❌' if strict else '⚠ ', len(diffs)))
            for pg, only_now, only_ref in diffs[:10]:
                print('   第 %d 页' % pg)
                for t in only_now:
                    print('     仅当前：%s' % t[:80])
                for t in only_ref:
                    print('     仅参考：%s' % t[:80])
        else:
            print('✅ 文案：与参考逐页等价（忽略顺序与换行编码）')

    print()
    if fail:
        print('结论：不通过（%d 项）。修完重新跑。' % fail)
        return 1
    print('结论：全部通过。')
    return 0


if __name__ == '__main__':
    sys.exit(main())
