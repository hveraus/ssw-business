# -*- coding: utf-8 -*-
"""Build "SSW AI 业务" company-intro / case-study deck v2, replicating the
original deck's visual system (720x405pt, Microsoft YaHei, #CC4141 accent)."""
import copy
from pathlib import Path
from pptx import Presentation
from pptx.util import Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

DECK_DIR = Path(__file__).resolve().parent
OUT = DECK_DIR / "SSW AI 业务_v3.pptx"
ASSETS = str(DECK_DIR / "assets")

W, H = 720.0, 405.0
RED = "CC4141"; DARK = "333333"; MID = "797979"; LIGHT = "AAAAAA"
CARD = "F7F7F8"; BORDER = "E3E3E6"; WHITE = "FFFFFF"; TINT = "FBEFEF"
DOT = "B5B9BE"
FONT = "Microsoft YaHei"
FOOTER_TEXT = "SSW AI 业务 · 企业 AI 转型"

ALIGN = {"l": PP_ALIGN.LEFT, "c": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT, "j": PP_ALIGN.JUSTIFY}
ANCH = {"t": MSO_ANCHOR.TOP, "m": MSO_ANCHOR.MIDDLE, "b": MSO_ANCHOR.BOTTOM}


# ---------------------------------------------------------------- primitives
def vlen(t):
    n = 0.0
    for ch in t:
        o = ord(ch)
        if ch == " ":
            n += 0.3
        elif o < 0x2000:
            n += 0.55
        else:
            n += 1.0
    return n


LINK_QUEUE = []          # [(shape, 目标页码)]，构建完成后统一解析


def rect(s, l, t, w, h, fill=None, line=None, lw=0.75):
    sh = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Pt(l), Pt(t), Pt(w), Pt(h))
    sh.shadow.inherit = False
    if fill:
        sh.fill.solid(); sh.fill.fore_color.rgb = RGBColor.from_string(fill)
    else:
        sh.fill.background()
    if line:
        sh.line.color.rgb = RGBColor.from_string(line); sh.line.width = Pt(lw)
    else:
        sh.line.fill.background()
    sh.text_frame.text = ""
    return sh


def _rpr(run, size, bold, color, font=FONT):
    f = run.font
    f.size = Pt(size); f.bold = bold
    f.color.rgb = RGBColor.from_string(color)
    f.name = font
    rPr = run._r.get_or_add_rPr()
    rPr.set("lang", "zh-CN")
    for tag in ("a:latin", "a:ea", "a:cs"):
        el = rPr.find(qn(tag))
        if el is None:
            el = rPr.makeelement(qn(tag), {}); rPr.append(el)
        el.set("typeface", font)


def txt(s, l, t, w, h, paras, anchor="t", align="l", ls=1.12):
    """paras: list of dicts {runs:[(text,size,bold,color)] or text=..., size, bold, color, align, sa}"""
    tb = s.shapes.add_textbox(Pt(l), Pt(t), Pt(w), Pt(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = ANCH[anchor]
    for i, p in enumerate(paras):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.alignment = ALIGN[p.get("align", align)]
        para.line_spacing = p.get("ls", ls)
        para.space_after = Pt(p.get("sa", 0))
        if "runs" in p:
            for rt, rsz, rb, rc in p["runs"]:
                for k, part in enumerate(rt.split("\n")):
                    if k:
                        r = para.add_run(); r.text = part; _rpr(r, rsz, rb, rc)
                        r._r.addprevious(para._p.makeelement(qn("a:br"), {}))
                    elif part:
                        r = para.add_run(); r.text = part; _rpr(r, rsz, rb, rc)
        else:
            for k, part in enumerate(p["text"].split("\n")):
                if k:
                    r = para.add_run(); r.text = part
                    _rpr(r, p.get("size", 9), p.get("bold", False), p.get("color", DARK))
                    r._r.addprevious(para._p.makeelement(qn("a:br"), {}))
                elif part:
                    r = para.add_run(); r.text = part
                    _rpr(r, p.get("size", 9), p.get("bold", False), p.get("color", DARK))
    return tb


def P(text, size=9, bold=False, color=DARK, align=None, sa=0, ls=1.12):
    d = {"text": text, "size": size, "bold": bold, "color": color, "sa": sa, "ls": ls}
    if align: d["align"] = align
    return d


def R(runs, size=9, align=None, sa=0, ls=1.12):
    d = {"runs": [(t, size, b, c) for (t, b, c) in runs], "sa": sa, "ls": ls}
    if align: d["align"] = align
    return d


# ---------------------------------------------------------------- page chrome
def case_chrome(s, cur=None):
    """案例段标识：左侧贯穿红条（案例首页更粗）+ 页脚 01–05 案例导航点。cur=当前案例序号(1-5)。"""
    rect(s, 0, 4.3, 10.44 if cur else 5.4, 371.5, fill=RED)
    if cur:
        rect(s, 0, 0, W, 6.84, fill=RED)
        for i in range(5):
            txt(s, 313 + i * 28, 383, 20, 14,
                [P("%02d" % (i + 1), 7, i == cur - 1, RED if i == cur - 1 else DOT)],
                anchor="m")


def base(s, bg=WHITE, topbar=True, footer=True, page=None):
    rect(s, 0, 0, W, H, fill=bg, line=bg)
    if topbar:
        rect(s, 0, 0, W, 4, fill=RED)
    if footer:
        rect(s, 36, 379, 648, 1, fill=BORDER)
        txt(s, 36, 383, 500, 14, [P(FOOTER_TEXT, 7.5, False, LIGHT)])
        if page:
            txt(s, 570, 383, 114, 14, [P(page, 7.5, False, LIGHT, align="c")], align="c")


def header(s, tag, title, sub=None, tsize=None):
    w = max(78, 26 + 9.6 * vlen(tag))
    rect(s, 36, 20, w, 20, fill=RED)
    txt(s, 36, 20, w, 20, [P(tag, 8, True, WHITE, align="c")], anchor="m", align="c")
    tsize = tsize or fit_size(title)
    txt(s, 36, 44, 648, 32, [P(title, tsize, True, DARK)], anchor="m")
    rect(s, 36, 89, 40, 3, fill=RED)
    if sub:
        txt(s, 88, 84, 596, 15, [P(sub, 10, False, MID)], anchor="m")


def fit_size(t, maxw=644, sizes=(27, 24, 22, 20, 18, 16)):
    for sz in sizes:
        if vlen(t) * sz <= maxw:
            return sz
    return 16


# ---------------------------------------------------------------- components
def stat_cards(s, stats, x0=36, y=114, w=156, h=64, gap=8, nsize=17, lsize=8):
    for i, (num, lab) in enumerate(stats):
        x = x0 + i * (w + gap)
        rect(s, x, y, w, h, fill=CARD, line=BORDER)
        txt(s, x + 10, y + 8, w - 20, 22, [P(num, nsize, True, RED)], anchor="m")
        txt(s, x + 10, y + 32, w - 20, 26, [P(lab, lsize, False, DARK, ls=1.1)], anchor="t")


def card_head(s, l, t, w, label, color=RED, size=10):
    rect(s, l, t, 18, 2.5, fill=RED)
    txt(s, l, t + 6, w, 18, [P(label, size, True, color)])


def bullet_list(s, l, t, w, items, gap_y=24, size=8.5, color=DARK, dash=True, mark_w=14):
    y = t
    for it in items:
        if dash:
            rect(s, l, y + 5, 6, 2, fill=RED)
        txt(s, l + mark_w, y, w - mark_w, gap_y - 2, [P(it, size, False, color, ls=1.08)])
        y += gap_y



def three_cols(s, heads, cols, y=116, h=222, hl_first=True, card3_fill=CARD,
               count_rows=None, row_labels=None, sub_groups=None):
    xs = [36, 257, 478]
    for i in range(3):
        fill = card3_fill if i == 2 else WHITE
        rect(s, xs[i], y, 206, h, fill=fill, line=BORDER)
        if hl_first and i == 2:
            rect(s, xs[i], y, 206, 3, fill=RED)
        card_head(s, xs[i] + 14, y + 12, 178, heads[i], RED if i != 0 else MID, 9.5)
    return xs


# ================================================================ page builders
def p01(s):
    base(s, bg=CARD, topbar=False, footer=False)
    rect(s, 0, 0, 10, H, fill=RED)
    txt(s, 65, 76, 300, 24, [P("SSW AI 业务", 10, False, MID)])
    rect(s, 65, 112, 79, 3, fill=RED)
    txt(s, 62, 150, 410, 46, [P("企业 AI 转型专家", 34, True, DARK, ls=1.18)])
    txt(s, 62, 214, 410, 40,
        [P("将 AI 接入企业数据与业务流程，交付可持续运行的系统", 11, False, MID, ls=1.3)])
    rect(s, 475, 133, 1, 196, fill=BORDER)
    nav = [("01", "时代背景与公司", "P03–P09", 3), ("02", "客户案例", "P10–P29", 10),
           ("03", "AI 中台设施", "P30–P33", 30), ("04", "我们提供哪些服务", "P34–P41", 34),
           ("05", "AI 平台落地", "P42–P50", 42), ("06", "如何落地", "P51–P54", 51),
           ("07", "合作模式", "P55–P58", 55), ("08", "数据安全", "P59", 59)]
    for i, (num, name, rng, tgt) in enumerate(nav):
        yy = 132 + i * 24
        _a = txt(s, 491, yy, 140, 22,
                 [R([(num + "　", False, LIGHT), (name, True, DARK)], 10.5)], anchor="m")
        _b = txt(s, 631, yy, 60, 22, [P(rng, 8.5, False, LIGHT, align="r")],
                 anchor="m", align="r")
        LINK_QUEUE.extend([(_a, tgt), (_b, tgt)])
    txt(s, 36, 383, 648, 14, [P("01", 7.5, False, LIGHT)], align="r")


def p02(s):
    base(s, page="06")
    header(s, "公司介绍", "关于 SSW：30+ 年企业软件交付经验",
           "30+ 年做企业软件，现在把这些年攒下的工程做法用在 AI 上")
    rect(s, 36, 116, 330, 202, fill=CARD, line=BORDER)
    txt(s, 50, 128, 302, 24, [P("智鹏瑞尔软件（杭州）有限公司", 12, True, DARK)], anchor="m")
    txt(s, 50, 158, 302, 148, [
        P("智鹏瑞尔软件（杭州）有限公司是 SSW 在中国的交付主体，总部在澳大利亚悉尼。",
          8.5, False, DARK, sa=6, ls=1.22),
        P("微软云合作伙伴，做数据与 AI、数字化创新和系统建设。",
          8.5, False, DARK, sa=6, ls=1.22),
        P("我们做事的顺序是先找到问题的最优解，再动手写代码，最后交一份能长期维护的软件。",
          8.5, False, DARK, ls=1.22),
    ])
    stats = [("30+", "年企业软件经验"), ("1,400+", "服务企业数量"),
             ("15", "覆盖国家与地区"), ("100+", "专业 IT 顾问")]
    for i, (num, lab) in enumerate(stats):
        cx = 382 + (i % 2) * 159
        cy = 116 + (i // 2) * 106
        rect(s, cx, cy, 143, 96, fill=CARD, line=BORDER)
        txt(s, cx + 8, cy + 12, 127, 44, [P(num, 24, True, RED)], anchor="m", align="c")
        txt(s, cx + 8, cy + 58, 127, 30, [P(lab, 8.5, False, DARK, ls=1.15)], align="c")
    txt(s, 36, 308, 648, 16,
        [P("以上资历数字统计截至 2026 年。", 8, False, MID)])
    txt(s, 36, 330, 648, 18,
        [P("SSW TV　英文社区，YouTube 4 万粉丝，1000 多条视频", 8.5, False, MID)])
    txt(s, 36, 350, 648, 18,
        [P("FireUG　中文社区，B 站、小红书 6 万粉丝", 8.5, False, MID)])


def p03(s):
    base(s, page="07")
    header(s, "公司介绍", "SSW 全球办公室",
           "杭州做交付，悉尼做技术把关，全程同一家公司，没有转包")
    s.shapes.add_picture(ASSETS + "/worldmap_red.png", Pt(36), Pt(122), width=Pt(376))
    offices = [("悉尼（总部）", "新南威尔士 · 澳大利亚"), ("布里斯班", "昆士兰 · 澳大利亚"),
               ("墨尔本", "维多利亚 · 澳大利亚"), ("纽卡斯尔", "新南威尔士 · 澳大利亚"),
               ("杭州", "浙江省 · 中国"), ("斯特拉斯堡", "阿尔萨斯 · 法国")]
    for i, (city, region) in enumerate(offices):
        y = 116 + i * 34
        rect(s, 432, y, 252, 30, fill=CARD, line=BORDER)
        txt(s, 444, y, 228, 30, [R([(city + "　", True, DARK), (region, False, MID)], 9)],
            anchor="m")
    rect(s, 432, 322, 252, 40, fill=TINT, line=BORDER)
    txt(s, 444, 322, 228, 40, [P("6 个办公室 · 3 国 · 3 大洲", 11, True, RED)], anchor="m")


def p04(s):
    base(s, page="08")
    header(s, "公司介绍", "交付组织与驻场方式",
           "同一批人从方案做到上线；人员变动时有交接、有复审，不掉链子")
    stats = [("30+ 年", "企业软件交付经验，从桌面软件、网站、云一路做到 AI"),
             ("100+", "专业顾问，分布澳大利亚 4 城市 + 中国 + 欧洲"),
             ("6 个", "办公室 · 3 国 · 3 大洲，交付由杭州团队执行"),
             ("0转包", "所有交付都由自己的团队完成，中间不经第三方")]
    for i, (num, lab) in enumerate(stats):
        cx = 36 + (i % 2) * 331
        cy = 114 + (i // 2) * 72
        rect(s, cx, cy, 317, 64, fill=CARD, line=BORDER)
        txt(s, cx + 13, cy, 94, 64, [P(num, 22, True, RED)], anchor="m")
        txt(s, cx + 112, cy, 194, 64, [P(lab, 9, False, DARK, ls=1.15)], anchor="m")
    rect(s, 36, 262, 648, 100, fill=WHITE, line=BORDER)
    rect(s, 36, 262, 648, 3, fill=RED)
    txt(s, 50, 272, 320, 22, [P("三个角色，一起对结果负责", 12, True, DARK)], anchor="m")
    txt(s, 50, 296, 620, 18,
        [P("项目负责人常驻你的现场，悉尼总部做技术评审与把关，杭州团队负责实施——三张签名在同一个交付单上。", 8.5, False, MID)])
    for i, item in enumerate(["项目负责人：常驻现场、定优先级", "技术评审：悉尼把关方案与架构", "交付执行：杭州实施、边做边改"]):
        x = 50 + i * 213
        rect(s, x, 330, 22, 3, fill=RED)
        txt(s, x, 338, 200, 18, [P(item, 9, False, DARK)])
    txt(s, 36, 366, 648, 12,
        [P("核心团队从方案到上线全程连续；人员变动时做完整交接，并由悉尼总部复审。", 8, False, LIGHT)])


def p05(s):
    base(s, page="09")
    header(s, "公司介绍", "我们的解决方案",
           "想清楚要做什么，把系统建起来，再让它一直有人用")
    items = [("AI 转型解决方案", "把 AI 接进现有业务流程：查资料、写报告、做分析、答客户"),
             ("AI 模型与用量", "代采主流模型，比官方直购便宜（见 P40）"),
             ("企业系统与架构", "架构可维护、文档齐全、能持续升级"),
             ("云与基础设施", "上云、自动发布、自动备份，少出事故"),
             ("客户与销售体系（CRM）", "从找客户到签合同的流程自动化"),
             ("数据与经营决策", "把散在各处的数据变成管理层能看的报表")]
    for i, (t, b) in enumerate(items):
        cx = 36 + (i % 3) * 221
        cy = 116 + (i // 3) * 120
        rect(s, cx, cy, 206, 110, fill=CARD, line=BORDER)
        rect(s, cx, cy, 206, 3, fill=RED)
        txt(s, cx + 14, cy + 14, 178, 22, [P(t, 11, True, DARK)], anchor="m")
        txt(s, cx + 14, cy + 40, 178, 62, [P(b, 8.5, False, MID, ls=1.25)])
    txt(s, 36, 358, 648, 16,
        [P("上面这些能力都有真实项目可查。架构可维护、文档齐全、能持续升级，不用冷门技术锁死你。",
           8.5, False, MID, align="c")], align="c")


def p06(s):
    base(s, page="11")
    header(s, "成功案例", "案例简介",
           "5 个行业，5 个项目")
    cards = [
        ("精细化工", "杭州某化工公司", "历史配方和试验记录难查", "2 个月", "建成可查配方的知识库"),
        ("外贸制造", "宁波某外贸工艺品厂家", "产品资料分散，提案靠手工拼", "80% 以上", "提案制作耗时降幅"),
        ("证券 / 资管", "北京某证券机构", "尽调报告手工写，还要逐条核出处", "80%", "尽调报告生成耗时降幅"),
        ("食品供应链", "某供应链企业", "客户靠人工拉群，跟单靠人盯", "1 万个客户群", "由 4 名运营人员管理"),
        ("高校 · 科研", "某高校与科研机构", "研究原型难以投入日常使用", "科研平台上线", "我们接手校方原型\n完成开发和部署"),
    ]
    for i, (tag, name, pain, value, label) in enumerate(cards):
        x = 36 + i * 131
        rect(s, x, 122, 124, 200, fill=CARD, line=BORDER)
        rect(s, x, 122, 124, 3, fill=RED)
        pw = min(104, max(56, 10 + 7.2 * vlen(tag)))
        rect(s, x + 10, 134, pw, 18, fill=RED)
        txt(s, x + 10, 134, pw, 18, [P(tag, 7, True, WHITE, align="c")], anchor="m", align="c")
        txt(s, x + 10, 162, 104, 40, [P(name, 10, True, DARK, ls=1.2)])
        txt(s, x + 10, 208, 104, 52, [P(pain, 8, False, MID, ls=1.2)])
        rect(s, x + 10, 255, 104, 1, fill=BORDER)
        txt(s, x + 10, 263, 104, 25, [P(value, 15, True, RED)], anchor="m")
        txt(s, x + 10, 293, 104, 23, [P(label, 8.5, True, DARK)])
    txt(s, 36, 332, 648, 18,
        [P("每个案例都按「痛点 → 我们做了什么 → 结果」展开。", 8.5, False, MID)])


# ---- case 1 ----------------------------------------------------------------
def p07(s):
    base(s, page="15")
    case_badge(s, 2)
    rows = [("行业", "外贸制造"), ("项目类型", "系统 AI 升级"),
            ("规模", "7 个模块全面升级"), ("交付方式", "2→4 月 · 2 周一个版本")]
    # manual (case_hero kept for reference)
    header(s, "案例二 · 宁波某外贸工艺品厂家",
           "用 AI 重做产品库与提案工作流",
           "宁波某外贸工艺品厂家 · 产品库与提案辅助")
    rect(s, 36, 116, 400, 230, fill=CARD, line=BORDER)
    card_head(s, 50, 132, 372, "项目简介", RED, 9.5)
    txt(s, 50, 152, 372, 42, [P("一家做外贸工艺品生意的工厂，产品型号多、认证杂，资料散在服务器和邮箱里。我们把产品资料建成一个能搜的库，再让销售一句话就能出带图提案和多语言文案。", 8.5, False, DARK, ls=1.25)])
    card_head(s, 50, 200, 372, "要解决的问题", RED, 9.5)
    txt(s, 50, 220, 372, 46, [P(
        "老销售的提案本事装在自己脑子里，人一走本事就没了；美国产品团队天天在做 PPT，"
        "做完就散，进不了公司资料库；文档、认证、专利散在服务器和邮箱里，只能搜文件名，"
        "找不到、找不准；每份外贸提案都要人工东拼西凑，耗时很长。", 8, False, DARK, ls=1.2)])
    card_head(s, 50, 272, 372, "我们做了什么", RED, 9.5)
    txt(s, 50, 292, 372, 46, [P(
        "先把产品资料建库：散落的 PPT、产品手册、认证和专利自动拆开整理，不用人工录入；"
        "\n把产品、认证和专利关联起来，销售一次就能查到答案和出处，也能用图片搜索同款样品。\n"
        "最后直接出带图提案和多语言文案。", 8, False, DARK, ls=1.2)])
    rect(s, 444, 116, 240, 148, fill=WHITE, line=BORDER)
    txt(s, 458, 126, 212, 16, [P("项目一览", 9, True, RED)])
    y = 148
    for k, v in rows:
        txt(s, 458, y, 212, 11, [P(k, 7, False, LIGHT)])
        txt(s, 458, y + 11, 212, 17, [P(v, 8.5, True, DARK, ls=1.05)])
        y += 29
    rect(s, 444, 272, 240, 74, fill=TINT, line=BORDER)
    txt(s, 456, 280, 216, 40, [P("80%+", 28, True, RED)], anchor="m")
    txt(s, 456, 322, 216, 22, [P("外贸提案时间缩短\nAI 一键生成带图提案 + 多语言文案", 8, False, DARK, ls=1.15)])


def p08(s):
    base(s, page="16")
    header(s, "案例二 · 宁波某外贸工艺品厂家", "AI 原生的产品库建设",
           "从设计上就是给 AI 用的，不是先建好再去适配")
    xs = [36, 257, 478]
    for i, (h, col) in enumerate([("客户", None),
                                  ("一期已完成 · 二期已交付", None),
                                  ("核心痛点（客户原话）", None)]):
        rect(s, xs[i], 116, 206, 222, fill=CARD if i == 2 else WHITE, line=BORDER)
        if i == 2:
            rect(s, xs[i], 116, 206, 3, fill=RED)
        card_head(s, xs[i] + 14, 128, 178, h, RED if i != 0 else MID, 9.5)
    # col 1
    rows = [("客户名称", "外贸厂家"), ("行业", "外贸制造"), ("项目类型", "系统 AI 升级"),
            ("系统形态", "网站 + 内部管理后台"), ("AI 能力", "查资料 · 写提案 · 配图")]
    y = 152
    for k, v in rows:
        txt(s, 50, y, 178, 11, [P(k, 7, False, LIGHT)])
        txt(s, 50, y + 11, 178, 17, [P(v, 8.5, True, DARK, ls=1.05)])
        y += 37
    # col 2
    txt(s, 271, 152, 178, 16, [P("一期：按机器能读的标准建库", 8, True, MID)])
    bullet_list(s, 271, 170, 178, ["产品信息管理中心", "样品间管理 + 四级权限体系",
                                   "门户检索 · 与物料 / 质量 / 协作系统连通"], gap_y=22, size=8.5)
    txt(s, 271, 240, 178, 16, [P("二期已交付（AI 全面升级）", 8.5, True, RED)])
    bullet_list(s, 271, 258, 178, ["产品资料自动提取 · 会说话的搜索",
                                   "传张图就能搜到同款样品",
                                   "带图提案 + 多语言文案，直接出稿"], gap_y=22, size=8.5)
    # col 3
    bullet_list(s, 492, 152, 178, [
        "提案本事在销售个人手里，传不下去",
        "美国产品团队持续出 PPT，进不了库",
        "搜索只能按关键词找",
        "急用时找不到，找到的也不准",
        "中美两地资料对不上",
        "一份提案靠人工拼装"], gap_y=30, size=8.5)
    txt(s, 36, 346, 648, 30,
        [P("产品库从设计上就是给 AI 用的：一期把字段、关系、权限理清楚，机器能直接读；"
           "二期加 AI 时不用重做数据层，直接在上面接助手。顺序对了，钱不用花两遍。", 9, True, RED)])




def p10(s):
    base(s, page="18")
    header(s, "案例二 · 宁波某外贸工艺品厂家", "多语言提案智能体",
           "会聊天，能记事，也能自己去公司系统里找东西")
    s.shapes.add_picture(ASSETS + "/agent-ui.png", Pt(36), Pt(116), width=Pt(384))
    txt(s, 444, 116, 240, 16, [P("核心能力", 9, True, RED)])
    items = [("说中文说英文", "都行"),
             ("多个助手分工协作", "（一个查产品、一个写文案、一个出图）"),
             ("每个人可以按自己的习惯", "调成顺手的样子"),
             ("直连产品库、样品间", "，不用来回切系统"),
             ("每个助手有自己的「记忆和性格」", "，越用越顺手"),
             ("会自己总结、自己积累", "")]
    y = 140
    for red, plain in items:
        rect(s, 444, y + 5, 6, 2, fill=RED)
        runs = [(red, True, RED)] + ([(plain, False, DARK)] if plain else [])
        txt(s, 458, y, 226, 30, [R(runs, 8.5, ls=1.15)])
        y += 32


# ---- case 2 ----------------------------------------------------------------
def p11(s):
    base(s, page="19")
    case_badge(s, 3)
    header(s, "案例三 · 北京某证券机构", "金融行业智能体：给 AI 装上合规护栏",
           "北京某证券机构 · 合规场景落地")
    rect(s, 36, 116, 400, 230, fill=CARD, line=BORDER)
    card_head(s, 50, 132, 372, "项目简介", RED, 9.5)
    txt(s, 50, 152, 372, 42, [P("一家持牌证券机构，AI 不是想上就能上，审计这一关必须先过。我们把整套系统装进客户自己的机房，业务人员说一句话就能起草尽调报告，每条结论带出处、每一步留痕。", 8.5, False, DARK, ls=1.25)])
    card_head(s, 50, 200, 372, "要解决的问题", RED, 9.5)
    txt(s, 50, 220, 372, 46, [P(
        "一份尽调报告靠人工写，2–3 天一份，占掉分析师大量时间；研报、募集书等文件散在各处，"
        "要用的时候找不到；业务经验靠师傅带徒弟，人一走就断档；几个系统各管一摊，数据不能互相用；"
        "持牌机构审计要求严，AI 不是想上就能上。", 8, False, DARK, ls=1.2)])
    card_head(s, 50, 272, 372, "我们做了什么", RED, 9.5)
    txt(s, 50, 292, 372, 46, [P(
        "整套系统装进客户自己的机房，数据不出公司；业务人员提一句需求，系统按模板"
        "自动填满报告，每条结论都标出处；对话中整理出的知识，走三级审批后才推给团队；"
        "每一步操作都留痕，可以查，审计这一关过得去。", 8, False, DARK, ls=1.2)])
    rect(s, 444, 116, 240, 148, fill=WHITE, line=BORDER)
    txt(s, 458, 126, 212, 16, [P("项目一览", 9, True, RED)])
    y = 148
    for k, v in [("行业", "证券 / 资管"), ("项目类型", "AI 助手 + 合规落地"),
                 ("规模", "5 个模块 · 全内网"), ("部署与合规", "装在自己机房 · 可审计")]:
        txt(s, 458, y, 212, 11, [P(k, 7, False, LIGHT)])
        txt(s, 458, y + 11, 212, 17, [P(v, 8.5, True, DARK, ls=1.05)])
        y += 29
    rect(s, 444, 272, 240, 74, fill=TINT, line=BORDER)
    txt(s, 456, 280, 216, 40, [P("80%", 28, True, RED)], anchor="m")
    txt(s, 456, 322, 216, 22, [P("尽调报告生成时间缩短\n2–3 天 → 数小时", 8, False, DARK, ls=1.15)])


def p12(s):
    base(s, page="20")
    header(s, "案例三 · 北京某证券机构", "痛点与应对：AI 如何应对严格的合规",
           "券商最怕的不是 AI 不好用，是 AI 不好管")
    xs = [36, 257, 478]
    heads = ["客户画像", "业务痛点", "合规怎么落地"]
    for i in range(3):
        rect(s, xs[i], 116, 206, 222, fill=WHITE, line=BORDER)
        card_head(s, xs[i] + 14, 128, 178, heads[i], RED if i == 2 else MID, 9.5)
    y = 152
    for k, v in [("客户", "证券公司"), ("行业", "证券 / 资管"), ("项目类型", "AI 智能助手系统"),
                 ("部署方式", "装在自己机房"), ("合规要求", "持牌机构 · 审计要求严")]:
        txt(s, 50, y, 178, 11, [P(k, 7, False, LIGHT)])
        txt(s, 50, y + 11, 178, 17, [P(v, 8.5, True, DARK, ls=1.05)])
        y += 37
    bullet_list(s, 271, 158, 178, [
        "尽调报告手工撰写，一份 2–3 天",
        "研报、募集书散在各处，要用时找不到",
        "业务知识靠人传人，沉淀不住",
        "审计要求高，AI 引入门槛大",
        "几个系统各管一摊，数据不能互相用"], gap_y=34, size=8.5)
    bullet_list(s, 492, 158, 178, [
        "全内网运行，数据不出公司",
        "每条结论都标出处",
        "知识推广先走三级审批",
        "五类角色分权，个人 / 团队 / 公司隔离",
        "操作全程留痕，日志不可篡改"], gap_y=34, size=8.5, color=DARK)


def p13(s):
    base(s, page="21")
    header(s, "案例三 · 北京某证券机构", "无所不在的合规智能体",
           "每一层都有对应的人管，也都有对应的权限管")
    layers = [
        ("1", "员工每天用的界面", "尽调工作台 · 业绩工作台 · 知识库浏览 · 个人知识整理 · 知识推广审批台"),
        ("2", "管理员后台", "用户管理 · 角色与权限 · 审计合规 · 运行状态看板 · 系统设置"),
        ("3", "会办事的 AI 助手", "连续对话（能嵌进别的系统）· 听懂问题 · 拆开任务 · 越用越准 · 不乱来"),
        ("4", "资料库", "资料存放 · 找资料 · 常用套路库 · 资料关系网（二期）· 过期资料有人管"),
        ("5", "公共设施", "统一登录 · 权限发放 · 统一的 AI 模型入口 · 操作留痕 · 敏感信息打码存储"),
    ]
    y = 114
    for code, name, body in layers:
        rect(s, 36, y, 648, 32, fill=WHITE, line=BORDER)
        rect(s, 36, y, 3, 32, fill=RED)
        txt(s, 46, y, 104, 32, [R([(code + "  ", True, RED), (name, True, DARK)], 9)], anchor="m")
        txt(s, 158, y, 516, 32, [P(body, 8.5, False, DARK, ls=1.1)], anchor="m")
        y += 36
    xs = [36, 257, 478]
    info = [("AI 尽调报告", "提一句需求，按模板自动填满，每条标出处"),
            ("知识沉淀", "对话自动整理成知识草稿，审批通过才推给团队"),
            ("合规保障", "完全在内网运行，操作全程留痕，满足持牌审计")]
    for i, (t, b) in enumerate(info):
        rect(s, xs[i], 300, 206, 52, fill=CARD, line=BORDER)
        txt(s, xs[i] + 14, 306, 178, 18, [P(t, 9.5, True, RED)])
        txt(s, xs[i] + 14, 326, 178, 22, [P(b, 8, False, DARK, ls=1.12)])




def p15(s):
    base(s, page="22")
    header(s, "成功案例 · 精细化工", "案例四：技术答复从 2 小时缩到 15 分钟",
           "某大型涂料与树脂制造企业 · 技术服务与客诉助手")
    rect(s, 36, 116, 400, 230, fill=CARD, line=BORDER)
    card_head(s, 50, 132, 372, "项目简介", RED, 9.5)
    txt(s, 50, 152, 372, 42, [P("一家做涂料与树脂的制造企业。我们把产品说明书、检测报告和客诉记录统一入库，做了技术服务和客诉两个助手，销售手上直接有了能讲的技术说法。", 8.5, False, DARK, ls=1.25)])
    card_head(s, 50, 200, 372, "要解决的问题", RED, 9.5)
    txt(s, 50, 220, 372, 46, [P(
        "客户问「这款漆能不能用在铝合金上」「耐盐雾多少小时」，答案散在产品说明书、检测报告、"
        "历史邮件里；技术服务工程师平均要花 2–3 小时才能凑出一份正式答复；同一类质量投诉反复发生，"
        "上次怎么解决的查不到；销售手上没有能讲的技术说法，每个问题都要转技术支持。", 8, False, DARK, ls=1.2)])
    card_head(s, 50, 272, 372, "我们做了什么", RED, 9.5)
    txt(s, 50, 292, 372, 46, [P(
        "把产品说明书、检测报告、配方档案、客诉记录、往来邮件统一收进一个库，扫描件和老文件"
        "机器自动识别成能搜的文字；建「技术服务助手」，客户问题进去就自动找到对应产品、适用条件、"
        "检测数据，出一份带出处的答复草稿；建「客诉助手」，输入客户描述的现象就列出历史相似案例；"
        "按人分权，配方只有核心团队能看到。", 8, False, DARK, ls=1.2)])
    rect(s, 444, 116, 240, 148, fill=WHITE, line=BORDER)
    txt(s, 458, 126, 212, 16, [P("项目一览", 9, True, RED)])
    y = 148
    for k, v in [("行业", "精细化工 · 涂料 / 树脂"), ("项目类型", "技术服务与客诉 AI 助手"),
                 ("规模", "2 个场景，已上线运行"), ("交付周期", "12 周完成交付")]:
        txt(s, 458, y, 212, 11, [P(k, 7, False, LIGHT)])
        txt(s, 458, y + 11, 212, 17, [P(v, 8.5, True, DARK, ls=1.05)])
        y += 29
    rect(s, 444, 272, 240, 74, fill=TINT, line=BORDER)
    txt(s, 452, 280, 224, 40, [P("2–3 小时 → 15 分钟", 17, True, RED)], anchor="m")
    txt(s, 456, 322, 216, 22, [P("技术答疑出稿时间（上线后实测）", 8, False, DARK)])


def pain_solution(s, page, tag, title, sub, rows, foot_red, foot_gray, third_w=206):
    base(s, page=page)
    header(s, tag, title, sub)
    xs = [36, 257, 478]
    heads = ["痛点（客户原话）", "我们的动作", "效果"]
    for i in range(3):
        rect(s, xs[i], 116, 206, 26, fill=CARD, line=BORDER)
        txt(s, xs[i] + 14, 116, 178, 26, [P(heads[i], 9, True, RED if i != 1 else MID)], anchor="m")
    y = 148
    for a, b, c in rows:
        for i, val in enumerate([a, b, c]):
            rect(s, xs[i], y, 206, 48, fill=WHITE, line=BORDER)
            col = RED if i == 2 else (DARK if i == 0 else DARK)
            bd = True if i == 2 else False
            txt(s, xs[i] + 14, y, 178, 48, [P(val, 8.5, bd, col, ls=1.18)], anchor="m")
        y += 52
    txt(s, 36, 316, 648, 20, [P(foot_red, 11, True, RED)])
    txt(s, 36, 340, 648, 18, [P(foot_gray, 8.5, False, MID)])


def p16(s):
    pain_solution(s, "23", "成功案例 · 精细化工",
                  "最费时间的三件事，各自对应一个动作",
                  "不讲 AI 原理，只讲谁省了多少时间",
                  [("答一个技术问题要翻三四个地方，半天就没了", "资料统一入库 + 带出处的问答",
                    "出稿 2–3 小时 → 15 分钟"),
                   ("同一类投诉年年有，上次的方案谁也记不住", "客诉助手自动匹配历史相似案例",
                    "首次响应时间 ↓60%"),
                   ("新人培养半年才算能上手", "老师傅的经验进系统，新人直接问",
                    "独立上岗 6 个月 → 3 个月")],
                  "↓70%　重复性技术问题转人工的比例",
                  "上线后客户没有更换任何现有系统，只在原有资料和流程上做加法。")


def result_page(s, page, tag, title, sub, stats, heads, cols, footnote=None):
    base(s, page=page)
    header(s, tag, title, sub)
    stat_cards(s, stats, y=116, w=156, h=66, gap=8, nsize=17)
    xs = [36, 257, 478]
    for i in range(3):
        rect(s, xs[i], 196, 206, 148, fill=CARD if i == 1 else WHITE, line=BORDER)
        if i == 1:
            rect(s, xs[i], 196, 206, 3, fill=RED)
        card_head(s, xs[i] + 14, 206, 178, heads[i], RED if i == 1 else MID, 9.5)
        bullet_list(s, xs[i] + 14, 236, 178, cols[i], gap_y=25,
                    size=8.5 if i == 1 else 8, color=DARK if i == 1 else MID)
    if footnote:
        txt(s, 36, 352, 648, 18, [P(footnote, 8, False, LIGHT)])


def p17(s):
    result_page(s, "24", "成功案例 · 交付成果与价值",
                "技术部从「救火队」变成「知识库」",
                "人手不变，能接的活多了",
                [("15 分钟", "技术答疑出稿\n原需 2–3 小时"), ("↓60%", "客诉首次响应时间"),
                 ("↓70%", "重复问题转人工比例"), ("12 周", "首批场景完成上线")],
                ["客服 / 销售侧", "技术侧", "管理侧"],
                [["销售自己就能答常见技术问题", "答复统一口径，不带个人说法",
                  "多语言客户问询同一套答案"],
                 ["工程师不再被打断，专注新配方", "客诉处理有历史依据，不重复分析",
                  "新人提问不再打扰老师傅"],
                 ["技术资料第一次变成公司资产", "权限分明，配方不外泄",
                  "老师傅经验留得下来，人走事不断"]],
                footnote="资料整理的成本是一次性的，收益从上线第一天起就开始累积。")


# ---- case 4 ----------------------------------------------------------------
def p18(s):
    base(s, page="29")
    header(s, "成功案例 · 汽车零部件", "案例五：一份 8D 报告，从 1–2 天到 2 小时",
           "某汽车零部件制造商 · 质量与供应商 AI 助手")
    rect(s, 36, 116, 400, 230, fill=CARD, line=BORDER)
    card_head(s, 50, 132, 372, "项目简介", RED, 9.5)
    txt(s, 50, 152, 372, 42, [P("一家有三个工厂的汽车零部件制造商。我们把质量记录、8D 报告和供应商来料数据统一入库，做了质量、报告、供应商三个助手，三个厂的经验第一次汇到一处。", 8.5, False, DARK, ls=1.25)])
    card_head(s, 50, 200, 372, "要解决的问题", RED, 9.5)
    txt(s, 50, 220, 372, 46, [P(
        "主机厂要的质量文件一份接一份：8D（质量问题整改报告）、PPAP（量产批准文件）、审核整改报告，全靠质量工程师手写，一份 8D 要 1–2 天；"
        "现场出质量异常，要翻历史同类问题找根因，主要靠老工程师的记忆；客户审核前一周全厂「翻箱倒柜」"
        "找资料补台账；三个工厂各管各的，同一个问题在不同厂重复踩坑。", 8, False, DARK, ls=1.2)])
    card_head(s, 50, 272, 372, "我们做了什么", RED, 9.5)
    txt(s, 50, 292, 372, 46, [P(
        "把 8D 报告、检验记录、审核报告、工艺文件、供应商来料数据统一入库，关键数据自动整理出来变成"
        "可比较的数据；建「质量助手」，输入现场现象就列出历史相似异常、当时定的根因和措施；"
        "建「报告助手」，按 8D / PPAP 模板自动生成初稿，引用真实记录和原始出处；按工厂和角色分权，"
        "总部能看全量。", 8, False, DARK, ls=1.2)])
    rect(s, 444, 116, 240, 148, fill=WHITE, line=BORDER)
    txt(s, 458, 126, 212, 16, [P("项目一览", 9, True, RED)])
    y = 148
    for k, v in [("行业", "汽车零部件 · 精密制造"), ("项目类型", "质量与供应商 AI 助手"),
                 ("规模", "3 个工厂 · 200+ 供应商已接入"), ("交付周期", "10 周完成交付")]:
        txt(s, 458, y, 212, 11, [P(k, 7, False, LIGHT)])
        txt(s, 458, y + 11, 212, 17, [P(v, 8.5, True, DARK, ls=1.05)])
        y += 29
    rect(s, 444, 272, 240, 74, fill=TINT, line=BORDER)
    txt(s, 452, 280, 224, 40, [P("1–2 天 → 2 小时", 17, True, RED)], anchor="m")
    txt(s, 456, 322, 216, 22, [P("8D 报告初稿时间（上线后实测）", 8, False, DARK)])


def p19(s):
    pain_solution(s, "30", "成功案例 · 汽车零部件",
                  "质量管理最费时间的三件事",
                  "省下来的是工程师的时间",
                  [("客户一投诉，我就得从头写一份 8D", "按模板自动出初稿，引用真实记录",
                    "8D 初稿 1–2 天 → 2 小时"),
                   ("同一个问题，上个月刚解决过", "自动匹配历史相似异常与措施",
                    "同类问题重复发生 ↓30%"),
                   ("审核前一周全厂找资料", "资料在线、按文件清单自动归集",
                    "审核前准备时间 ↓70%")],
                  "半天 → 10 分钟　供应商来料异常定位时间",
                  "这三个场景全部落在客户现有资料上完成，没有新增任何采集工作。")


def p20(s):
    result_page(s, "31", "成功案例 · 交付成果与价值",
                "三个工厂的质量经验，第一次汇到一处",
                "对接主机厂审核，手上得有数据，不能临时加班补",
                [("↓85%", "8D 报告编制时间\n原需 1–2 天"), ("↓70%", "审核前资料准备时间"),
                 ("↓30%", "同类问题重复发生率"), ("3 个工厂", "质量数据互通")],
                ["质量工程师", "管理层", "供应商管理"],
                [["报告有模板、有数据，不用从空白开始", "现场异常 10 分钟找到历史根因",
                  "三个厂的案例互相可查"],
                 ["一眼看到哪个厂、哪条线问题最集中", "审核前不再集中加班补资料",
                  "质量改进措施是否落地看得见"],
                 ["供应商表现有客观依据，不靠印象", "问题供应商自动浮出来，谈判有底",
                  "来料异常趋势可跟踪"]],
                footnote="三个工厂的质量数据一次打通，此后每一轮审核都直接用。")


# ---- case 5 ----------------------------------------------------------------
def p21(s):
    base(s, page="12")
    case_badge(s, 1)
    header(s, "案例一 · 杭州某化工公司", "几十年的研发积累，第一次能被「问」",
           "杭州某化工公司 · 知识库建设与研发辅助")
    rect(s, 36, 116, 400, 230, fill=CARD, line=BORDER)
    card_head(s, 50, 132, 372, "项目简介", RED, 9.5)
    txt(s, 50, 152, 372, 42, [P("一家做了几十年配方与工艺的化工企业。我们把它几十年的资料建成一个能问的知识库，再接到研发日常：研发新材料先查历史、找相似配方，不用再从零试一遍。", 8.5, False, DARK, ls=1.25)])
    card_head(s, 50, 200, 372, "要解决的问题", RED, 9.5)
    txt(s, 50, 220, 372, 46, [P(
        "这家企业做纺织化学助剂几十年，配方、试验、工艺资料是别人复制不了的家底。麻烦出在调用上："
        "研发新材料时，几年前做过的相似配方又得从零走一遍；境外母公司的技术文献只有"
        "少数人看得懂；各部门各自在用通用聊天工具，接不进公司数据，效果停在翻译和润色文案。",
        8, False, DARK, ls=1.2)])
    card_head(s, 50, 272, 372, "我们做了什么", RED, 9.5)
    txt(s, 50, 292, 372, 46, [P(
        "先把资料摸清、把场景排好序：哪些资料能进库、谁能看、从哪个部门先跑。"
        "\n把扫描件、老文档和外文文献转成可搜索的文字，把配方、试验和检测结果整理成表格。\n"
        "再把研发最常问的问题做成问答入口，每条答案都标出出自哪份文件。", 8, False, DARK, ls=1.2)])
    rect(s, 444, 116, 240, 148, fill=WHITE, line=BORDER)
    txt(s, 458, 124, 212, 16, [P("项目一览", 9, True, RED)])
    y = 144
    for k, v in [("行业", "化工 · 配方与工艺"), ("项目类型", "知识库建设 + 研发辅助"),
                 ("服务性质", "知识库 + 研发问答入口"), ("一期投入", "40 人天 · 周期 2 个月")]:
        txt(s, 458, y, 212, 10, [P(k, 7, False, LIGHT)])
        txt(s, 458, y + 10, 212, 16, [P(v, 8.5, True, DARK, ls=1.05)])
        y += 29
    rect(s, 444, 272, 240, 74, fill=TINT, line=BORDER)
    txt(s, 456, 278, 216, 40, [P("知识库", 26, True, RED)], anchor="m")
    txt(s, 456, 322, 216, 22, [P("几十年的资料、配方、试验记录都能被问，答案标出处", 8, False, DARK, ls=1.1)])


def p22(s):
    base(s, page="13")
    header(s, "案例一 · 杭州某化工公司", "痛点和应对",
           "六个痛点全部核对过现场情况，先从知识库下手")
    rect(s, 36, 116, 292, 218, fill=WHITE, line=BORDER)
    card_head(s, 50, 128, 264, "六个痛点（客户会说的原话）", MID, 9.5)
    bullets = [
        "试错成本高：历史批次调不出来，相似课题从零再做",
        "资料散落：找一份历史配方报告要问 2–3 个人",
        "机器读不了：扫描件、老文档、排版用 Excel、工艺曲线图",
        "外文文献：只有少数人看得懂，别人用不上",
        "专家被打断：一个问题被反复问，新人上手要数月",
        "通用工具使不上劲：接不进公司数据，只能润色和翻译",
    ]
    bullet_list(s, 50, 152, 264, bullets, gap_y=30, size=8.5, mark_w=12)
    rect(s, 340, 116, 344, 218, fill=CARD, line=BORDER)
    rect(s, 340, 116, 344, 3, fill=RED)
    card_head(s, 354, 128, 316, "知识库怎么建起来（从资料到能问）", RED, 9.5)
    paths = [
        ("1", "资料入库", "扫描件、老文档、外文文献统一识别成能搜的文字", "第 1 步"),
        ("2", "整理成表", "配方、原料、配比、检测结果从报告里抽成表格", "第 2 步"),
        ("3", "能搜能问", "关键词和说人话两种搜法，每条答案标出出处", "第 3 步"),
        ("4", "接到研发日常", "研发新材料先查历史，找相似配方、看当年怎么做", "第 4 步"),
    ]
    y = 152
    for code, name, body, ph in paths:
        txt(s, 354, y, 14, 20, [P(code, 9, True, RED)])
        txt(s, 372, y, 60, 20, [P(name, 8.5, True, DARK)])
        txt(s, 436, y, 179, 33, [P(body, 7.5, False, MID, ls=1.1)])
        txt(s, 618, y, 52, 20, [P(ph, 7.5, True, RED, align="r")], align="r")
        y += 45
    txt(s, 36, 344, 648, 18,
        [P("哪些资料进知识库、第一个研发入口从哪个部门先跑，这两件事都已经落地。",
           9, True, RED)])




def p24(s):
    base(s, page="31")
    header(s, "AI 中台", "AI 中台：把「各用各的 AI」变成「全公司共用一套底座」",
           "一次建好，后面每个场景都不用从头再来")
    rect(s, 36, 110, 648, 42, fill=CARD, line=BORDER)
    rect(s, 36, 110, 3, 42, fill=RED)
    txt(s, 50, 110, 620, 42, [P(
        "过去每个部门各自买 AI 工具，各花各的钱、各存各的数据、各踩各的坑。"
        "这层底座由我们建好、交到你手上：账号、花费、助手、计算能力四件事统一管起来，"
        "装在你选定的环境里。",
        8.5, False, DARK, ls=1.22)], anchor="m")
    txt(s, 36, 160, 648, 20,
        [P("像盖楼：地基打一次，上面想盖几层盖几层。每接一个新场景，只是多盖一层，"
           "不用重新挖地基、重新拉水电。", 10, True, RED)])
    layers = [
        ("应用层", "各部门具体用的东西", "研发问答 · 实验记录检索 · 配方比对 · 技术服务应答 · 新人培训助手 · 财务 / 销售 / 仓储…"),
        ("平台层", "建一次，各应用共用", "资料库与索引 · 带出处的检索 · 助手的组合与共享 · 提问模板与效果评测 · 权限与密级 · 操作留痕与监控"),
        ("数据层", "决定查得准不准", "格式识别 · 扫描件与图片转文字 · 配方与性能数据自动成表 · 多语种索引 · 内容校对 · 新增资料自动跟上"),
        ("公共设施层", "计算与通道", "本地服务器 · 接入 AI 模型的通道 · 公司账号统一登录 · 文件存储与权限 · 资料索引库 · 备份与运维"),
    ]
    y = 188
    for name, note, body in layers:
        rect(s, 36, y, 648, 36, fill=WHITE, line=BORDER)
        rect(s, 36, y, 3, 36, fill=RED)
        txt(s, 48, y, 118, 36, [R([(name, True, DARK), ("\n" + note, False, MID)], 8.5, ls=1.15)],
            anchor="m")
        txt(s, 176, y, 500, 36, [P(body, 8, False, DARK, ls=1.15)], anchor="m")
        y += 39
    txt(s, 36, 348, 648, 18,
        [P("资料进来时能不能被读懂，决定后面查得准不准。这一步没做好，后面换多贵的 AI 都补不回来。",
           8.5, False, MID)])


def p25(s):
    base(s, page="32")
    header(s, "AI 中台", "底座管四件事：都是钱和风险",
           "AI 项目做不成，问题多半出在这四件事上")
    cards = [
        ("① 账号与权限统一管理", "一套账号，按人、部门、项目分权。员工只能查看获准访问的配方、客户和财务数据。员工离职时可一键收回权限。",
         "像公司门禁：一张卡只能开该开的门"),
        ("② 用量统计与成本管控", "按人、团队、助手统计 AI 使用量与花费，超限自动提醒或拦截。费用不再是糊涂账，哪个部门用得值一眼看见。",
         "超了会自己拦下来"),
        ("③ 助手（技能和智能体）的组合与共享", "一次调好的助手，别的部门直接拿来用；多个助手可以配合完成跨部门任务。不做重复建设，好用的助手全公司扩散。",
         "不用每个部门各自再建一套"),
        ("④ 公共设施共用", "服务器、AI 模型调用通道、文件存储、资料索引建一次。新场景上线不用重买设备、重搭环境，周期短、投入小、风险低。",
         "像厂房和管线先建好，后面加产线直接接上"),
    ]
    for i, (t, b, a) in enumerate(cards):
        cx = 36 + (i % 2) * 331
        cy = 114 + (i // 2) * 120
        rect(s, cx, cy, 317, 112, fill=CARD, line=BORDER)
        rect(s, cx, cy, 317, 3, fill=RED)
        txt(s, cx + 14, cy + 14, 289, 22, [P(t, 11.5, True, RED)], anchor="m")
        txt(s, cx + 14, cy + 40, 289, 48, [P(b, 8.5, False, DARK, ls=1.2)])
        txt(s, cx + 14, cy + 90, 289, 18, [P(a, 8, True, MID)])
    txt(s, 36, 352, 648, 18,
        [P("这四项能力不产生「炫效果」，但决定 AI 能不能从一个部门的玩具，变成全公司的生产能力。",
           8.5, False, MID)])


def p26(s):
    base(s, page="33")
    header(s, "AI 中台", "行业、业务不同，底层是同一套可复用能力",
           "每个客户的数据、权限和运行环境独立管理")
    cases = ["化工 · 知识库", "外贸 · 产品库", "证券 · 合规", "供应链 · 销售链路"]
    x0, cw, gap = 140, 130, 6
    rect(s, 36, 116, 96, 34, fill=RED)
    txt(s, 36, 116, 96, 34, [P("底座能力", 8.5, True, WHITE, align="c")], anchor="m", align="c")
    for i, c in enumerate(cases):
        x = x0 + i * (cw + gap)
        rect(s, x, 116, cw, 34, fill=RED)
        txt(s, x, 116, cw, 34, [P(c, 8, True, WHITE, align="c")], anchor="m", align="c")
    data = [
        ("账号与权限统一", ["按部门与密级分开看", "四级权限分开管", "五类角色，资料分三级",
                        "客户群与销售分级"]),
        ("用量与成本管控", ["按部门与场景配额", "部门配额", "用户 / 团队 / 公司三级额度",
                        "按助手统计用量"]),
        ("助手的组合与共享", ["四期逐步增加应用", "多个助手配合干活", "能嵌进现有系统的对话功能",
                        "客户 / 销售 / 签约 / 订单四助手"]),
        ("公共设施共用", ["公共部分建一次，应用往上加", "沿用一期全部能力", "五层公共设施已建好",
                       "沿用同一套可复用能力"]),
    ]
    y = 154
    for name, vals in data:
        rect(s, 36, y, 96, 44, fill=CARD, line=BORDER)
        txt(s, 42, y, 84, 44, [P(name, 8, True, DARK, ls=1.1)], anchor="m")
        for i, v in enumerate(vals):
            x = x0 + i * (cw + gap)
            rect(s, x, y, cw, 44, fill=WHITE, line=BORDER)
            txt(s, x + 6, y, cw - 12, 44, [P(v, 7.5, False, DARK, ls=1.15)], anchor="m",
                align="c")
        y += 48
    txt(s, 36, 346, 648, 20,
        [P("SSW 已有成熟的 AI 底座，标准功能不重复收取软件授权费；我们按实际需要，收取部署、配置、系统对接和定制开发费用。所以第二个场景起周期更短。",
           10.5, True, RED)])


# ---- how we deliver --------------------------------------------------------
SV_CELL_XS = [162, 338, 514]
SV_CELL_W = 170


def sv_row(s, label, cells, y, rh=70, hi=False):
    rect(s, 36, y, 120, rh, fill=RED if hi else CARD, line=BORDER)
    if not hi:
        rect(s, 36, y, 3, rh, fill=RED)
    num, name = label
    txt(s, 44, y + 8, 106, 18, [P(num, 11, True, WHITE if hi else RED)])
    txt(s, 44, y + 28, 106, 36, [P(name, 9.5, True, WHITE if hi else DARK, ls=1.15)])
    for i, c in enumerate(cells):
        rect(s, SV_CELL_XS[i], y, SV_CELL_W, rh, fill=TINT if i == 2 else WHITE, line=BORDER)
        txt(s, SV_CELL_XS[i] + 8, y + 6, SV_CELL_W - 16, rh - 12,
            [P(c, 8, False, DARK, ls=1.24)])


def sv_rows_page(s, page, title, sub, rows, foot, hi_idx=None):
    base(s, page=page)
    header(s, "如何落地", title, sub)
    txt(s, 44, 110, 106, 18, [P("步骤", 8.5, True, MID)], anchor="m")
    for i, h in enumerate(["做法", "为什么必须做", "你拿到什么"]):
        rect(s, SV_CELL_XS[i], 110, SV_CELL_W, 18, fill=CARD, line=BORDER)
        txt(s, SV_CELL_XS[i] + 8, 110, SV_CELL_W - 16, 18,
            [P(h, 8.5, True, RED if i == 2 else MID)], anchor="m")
    y = 132
    for i, (label, cells) in enumerate(rows):
        sv_row(s, label, cells, y, rh=70, hi=(hi_idx == i))
        y += 76
    txt(s, 36, 362, 648, 16, [P(foot, 9.5, True, RED)])


def sv1(s):
    sv_rows_page(
        s, "52", "排场景，打地基，装资料",
        "第一步选错，后面所有投入都要重来",
        [(("第 1 步", "现状与场景盘点"),
          ["梳理资料、访谈各部门，优先选择投入小、见效快的场景",
           "场景选错，后面所有投入都要重来。我们见过直接让 AI 重写老系统的项目，最后整批返工",
           "一份排好序的场景清单，加上每一步的投入估算，先做哪个当场能定"]),
         (("第 2 步", "搭建 AI 底座"),
          ["把平台建一次：资料入库与索引、权限与密级、用量与成本、操作留痕、服务器与接入 AI 模型的通道",
           "跳过这一步，每增加一个场景都要重新搭建，重复花钱",
           "后续场景可沿用同一套系统，缩短建设周期，降低费用"]),
         (("第 3 步", "企业知识资产化"),
          ["说明书、报告、配方档案、客诉记录、往来邮件、扫描件统一收进一个库；"
           "识别成能搜的文字，原料、配比、检测结果自动整理成表格，再连成关系网",
           "通用大模型只学过公开互联网上的内容。你的配方、客诉记录、外文技术文献它一概没有，"
           "不搬进去，它只能帮你把话说漂亮",
           "公司几十年的积累第一次能「被问」，每条答案都能说出出自哪份文件"])],
        "公开互联网上的内容，通用大模型本来就有；真正只有你有的，是第 3 步搬进去的那部分。")


def sv2(s):
    sv_rows_page(
        s, "53", "落到业务、上线、演进",
        "系统上线只是开始，头三个月的调优决定它能不能活下来",
        [(("第 4 步", "业务场景开发"),
          ["按岗位做助手：技术答疑、客诉匹配、报告生成、尽调分析、质量追溯，"
           "每个助手配权限、配提问模板、配上验收指标",
           "通用工具接不进你的系统和数据，员工用两周就放下了，前面投入等于打了水漂",
           "助手落在具体岗位的日常里，省下多少工时算得出来，能直接写进立项汇报"]),
         (("第 5 步", "上线与交接"),
          ["先小范围真实业务试用，按答错、找漏的情况调优，再培训到每个岗位会用；"
           "同时交付文档、程序源代码和账号",
           "AI 不是装好就完事。头几周的问题不修，用户三天就弃用，前面所有投入都白花",
           "上线后真的有人在用；源代码、数据、账号都在你手里，不依赖我们也能继续跑"]),
         (("第 6 步", "运营与演进"),
          ["盯用量与效果数据、补资料缺口、按季度加新场景、整理资料并复核权限",
           "资料会旧、业务会变。没人维护的系统，一年后基本没人打开",
           "新场景越加越便宜；AI 从一次性项目变成公司自己的能力和团队"])],
        "六步可以一次委托，也可以先做前几步看结果，再决定后面怎么走。")


def sv3(s):
    base(s, page="54")
    header(s, "如何落地", "投入和产出",
           "投入可以分段，先走第一步再决定后面；不是每个场景都要走完六步")
    cols = [(36, 216, "步骤"), (252, 216, "大致周期"), (468, 216, "你要配合什么")]
    for x, w, h in cols:
        rect(s, x, 110, w, 22, fill=RED, line=RED)
        txt(s, x + 8, 110, w - 16, 22, [P(h, 8.5, True, WHITE)], anchor="m")
    rows = [
        ("1　现状与场景盘点", "约 1 个月", "各部门各抽 1 小时访谈"),
        ("2　部署 AI 底座", "约 1–2 个月", "提供服务器或机房条件、确认访问权限"),
        ("3　企业知识资产化", "约 1–3 个月", "提供资料、指定业务专家做校对"),
        ("4　业务场景开发", "每场景约 1–2 个月", "每个场景指定 1 位业务负责人"),
        ("5　上线与交接", "约 1 个月", "安排试用人员、组织参加培训"),
        ("6　运营与演进", "按季度", "每季度一次业务回顾"),
    ]
    y = 136
    for i, vals in enumerate(rows):
        fill = CARD if i % 2 == 0 else WHITE
        for (x, w, _), v in zip(cols, vals):
            rect(s, x, y, w, 30, fill=fill, line=BORDER)
            if x == 36:
                sz, bold, col = 9, True, DARK
            elif x == 252:
                sz, bold, col = 8.5, False, MID
            else:
                sz, bold, col = 8.5, False, DARK
            txt(s, x + 10, y, w - 20, 30, [P(v, sz, bold, col)], anchor="m")
        y += 30
    txt(s, 36, 322, 648, 20,
        [P("一个场景走完前五步约 115–270 人天（1 人天＝一位工程师干一天），不含系统对接；"
           "资料与配合程度不同，差别可到一倍。", 9.5, True, RED)])
    txt(s, 36, 344, 648, 18,
        [P("先做第 1 步拿到投入估算，再决定要不要往下走。"
           "案例里的 2–4 个月是单场景上线，完整路径按上表合计 5–9 个月。", 8.5, False, MID)])


# ---- speciality service line ----------------------------------------------
def svc_card(s, x, y, w, h, num, name, desc, bullets):
    rect(s, x, y, w, h, fill=WHITE, line=BORDER)
    rect(s, x, y, w, 3, fill=RED)
    card_head(s, x + 14, y + 12, w - 28, num + "　" + name, RED, 10)
    txt(s, x + 14, y + 34, w - 28, 14, [P(desc, 7.5, False, MID)])
    bullet_list(s, x + 14, y + 54, w - 28, bullets, gap_y=20, size=8, mark_w=12)


def sp1(s):
    base(s, page="35")
    header(s, "我们的服务", "在 AI 中台之上，我们对外提供四类服务",
           "横向四类能力，按你的业务场景落地；这四类是我们真正交付的东西")
    cards = [
        ("① 知识库搭建", "把资料变成能问的",
         ["扫描件、老文档统一入库", "配方、检测结果整理成表",
          "会说话的搜索，答案标出处", "按人分权限，分级查看"]),
        ("② 智能体定制", "按岗位做一个能干的助手",
         ["销售全链路、技术答疑", "客诉匹配、质量追溯",
          "报告生成、尽调分析", "配提问模板与验收指标"]),
        ("③ 系统集成", "让助手接进你现有的系统",
         ["接通 ERP / PLM / LIMS", "连通微信客户群与协作工具",
          "统一登录，沿用现有权限", "工作量单独评估"]),
        ("④ 驻场陪跑与运营", "人跟着业务一起跑",
         ["工程师长期在你现场", "上线后按答错、找漏调优",
          "按季度加场景、整资料", "把你的团队带起来"]),
    ]
    for i, (t, sub2, items) in enumerate(cards):
        x = 36 + i * 164
        rect(s, x, 112, 154, 218, fill=CARD, line=BORDER)
        rect(s, x, 112, 154, 3, fill=RED)
        txt(s, x + 12, 122, 130, 34, [P(t, 10.5, True, RED, ls=1.15)])
        txt(s, x + 12, 158, 130, 26, [P(sub2, 8.5, True, DARK, ls=1.15)])
        bullet_list(s, x + 12, 196, 130, items, gap_y=30, size=7.5, mark_w=10)
    txt(s, 36, 344, 648, 20,
        [P("这四类建在同一条技术线上，标准能力不重复建设；每个客户的数据、权限与运行环境独立管理。", 9.5, True, RED)])

def sp2(s):
    base(s, page="36")
    header(s, "我们的服务", "这些服务，落到你的哪些场景",
           "13 类场景都有方法沉淀，可以逐类评估；建议先选 1 个部门、1 条高频流程")
    left = [
        ("拓客与销售", ["销售与客户：线索、跟进、报价、签约、订单",
                    "市场营销：素材挖掘、竞品追踪、投放分析"]),
        ("研发 · 生产 · 供应链", ["研发管理：资料检索、实验记录、测试分析",
                            "项目管理：项目立项、进度管理、风险日报",
                            "生产运营：班前巡检、产能推演、异常记录",
                            "供应链采购：供应商画像、智能比价、采购预测",
                            "设备管理：智能运维、故障诊断、远程运维"]),
    ]
    right = [
        ("职能与后台", ["办公运营：智能客服、会议纪要、文档分析",
                    "合同法务：合同审查、风险识别、案件对比",
                    "财务管理：发票识别、报表生成、风险监控",
                    "人力资源：招聘助理、绩效助手、培训分析"]),
        ("经营与决策", ["经营决策：经营问答、经营报告、辅助分析",
                    "IT 研发运维：代码生成、自动测试、智能运维"]),
    ]
    for col, (x, blocks) in enumerate([(36, left), (368, right)]):
        rect(s, x, 116, 316, 222, fill=CARD if col == 0 else WHITE, line=BORDER)
        y = 116
        for gname, items in blocks:
            card_head(s, x + 16, y, 284, gname, RED, 9.5)
            y += 20
            bullet_list(s, x + 16, y, 284, items, gap_y=22, size=8, mark_w=11)
            y += len(items) * 22 + 10
    txt(s, 36, 348, 648, 20,
        [P("先选 1 个部门、1 条高频流程、1 个可量化指标做试点，跑通了再往下加。"
           "上面四类服务可以单独采，也可以组合。", 9.5, True, RED)])

def sp3(s):
    base(s, page="37")
    header(s, "我们的服务 · 举例", "智能体定制的例子：销售全链路，五段各管一段",
           "做企业客户的生意就三件事：线索、转化、交付。五段各管一段")
    stages = [
        ("①", "推广获客", ["品牌曝光与渠道触达", "新媒体代运营", "AI 搜索优化"], "线索量"),
        ("②", "线索运营", ["拉新、清洗、补全", "线索关联与打标", "动态监控"], "有效线索率"),
        ("③", "销售转化", ["沟通记录自动沉淀", "方案设计与报价", "样品申请与专家协同"], "转化率 · 人均产出 · 毛利"),
        ("④", "签约", ["按规则自动报价", "合同自动出稿", "条款与审批自动流转"], "签约周期"),
        ("⑤", "订单履约", ["群里一句指令下单", "拆单核价、查库存", "进度自动同步、异常上报"], "履约准时率"),
    ]
    for i, (num, name, items, metric) in enumerate(stages):
        x = 36 + i * 132
        rect(s, x, 116, 120, 140, fill=WHITE, line=BORDER)
        rect(s, x, 116, 120, 3, fill=RED)
        txt(s, x + 12, 126, 96, 14, [P(num, 9.5, True, RED)])
        txt(s, x + 12, 142, 96, 28, [P(name, 10.5, True, DARK, ls=1.1)])
        rect(s, x + 12, 176, 96, 0.8, fill=BORDER)
        bullet_list(s, x + 12, 184, 96, items, gap_y=23, size=7, mark_w=8)
        rect(s, x, 260, 120, 36, fill=TINT, line=BORDER)
        txt(s, x + 10, 260, 100, 36, [P(metric, 8.5, True, RED)], anchor="m")
    txt(s, 36, 306, 648, 20, [P("先上哪一段，取决于你现在最堵在哪一段。", 9.5, True, RED)])
    txt(s, 36, 330, 648, 18,
        [P("五段的产出都对应可验收的指标，不必一次全上。", 8.5, False, MID)])




def sp5(s):
    base(s, page="38")
    header(s, "我们的服务 · 举例", "同一个例子里，四个模块各管一段",
           "四个模块可以分别上线，也可以整条一起交")
    mods = [
        ("①", "客户助手", "直接面向下游客户的自助入口",
         ["客户自助下单、查进度", "产品资料与手册自助取", "重复问询量下降"]),
        ("②", "销售助手", "把销冠的做法复制给每一位销售",
         ["进群自动建档、跟进提醒", "把资料出成题，新人练完考完再上岗", "报价与催单自动推进，复盘自动沉淀"]),
        ("③", "签约助手", "报价与合同自动出稿，冲突点单独标出来",
         ["按规则自动报价", "按沟通记录自动起草合同", "账期这类冲突点自动触发审批"]),
        ("④", "订单助手", "群里一句指令，订单自己跑",
         ["群里一句指令下单", "拆单核价、查库存、定交期", "缺货延期异常自动上报"]),
    ]
    for i, (num, name, desc, items) in enumerate(mods):
        x = 36 + (i % 2) * 330
        y = 112 + (i // 2) * 116
        svc_card(s, x, y, 318, 106, num, name, desc, items)
    txt(s, 36, 348, 648, 20,
        [P("每一块都有已经在跑的真实界面。", 9.5, True, RED)])


def sp6(s):
    base(s, page="39")
    header(s, "我们的服务 · 推进方式", "找准切口，四步推进",
           "先立住、做出效果、养成能力、再谈上下游")
    cols = [(36, 112, "阶段"), (148, 384, "做什么"), (532, 152, "目标")]
    for x, w, h in cols:
        rect(s, x, 110, w, 24, fill=RED, line=RED)
        txt(s, x + 10, 110, w - 20, 24, [P(h, 9, True, WHITE)], anchor="m")
    rows = [
        ("1.0", "订单切入 · 代运营：从订单做起，先把获客、销售、履约接过来", "先立住"),
        ("2.0", "AI 助手 + 驻场工程师陪跑：不同角色的 AI 助手进群上岗，跟着业务一起迭代", "做出效果"),
        ("3.0", "客户自己做主：在平台上自己搭 AI 助手、改流程", "养成能力"),
        ("4.0", "产业上下游联动：和上下游一起用，走向行业专用 AI + 自己的服务器", "上下游协同"),
    ]
    y = 134
    for i, (ph, body, goal) in enumerate(rows):
        fill = CARD if i % 2 == 0 else WHITE
        rect(s, 36, y, 112, 38, fill=fill, line=BORDER)
        txt(s, 46, y, 92, 38, [P(ph, 12, True, RED)], anchor="m")
        rect(s, 148, y, 384, 38, fill=fill, line=BORDER)
        txt(s, 158, y, 364, 38, [P(body, 8.5, False, DARK, ls=1.14)], anchor="m")
        rect(s, 532, y, 152, 38, fill=TINT, line=BORDER)
        txt(s, 542, y, 132, 38, [P(goal, 9, True, RED)], anchor="m")
        y += 38
    rect(s, 36, 302, 648, 58, fill=CARD, line=BORDER)
    rect(s, 36, 302, 3, 58, fill=RED)
    for x, label, desc in [
        (52, "SSW 服务费", "部署、配置、系统对接和定制开发按工程师人天计费"),
        (374, "第三方费用", "服务器、机房和模型调用费用由客户直接支付"),
    ]:
        txt(s, x, 310, 294, 17, [P(label, 10.5, True, RED)])
        txt(s, x, 331, 294, 21, [P(desc, 9, False, DARK)])


# ---- differentiation -------------------------------------------------------
def p27(s):
    base(s, page="56")
    header(s, "为什么是我们", "为什么找我们",
           "AI 项目做不成，多数败在场景选错、数据没治、没人管得住")
    reasons = [
        ("01", "懂行业，才选得准场景",
         "30 年企业软件交付，覆盖化工、外贸制造、证券与食品供应链。AI 让资料和代码变便宜，"
         "先做哪一个、怎么拼起来，仍是行业判断"),
        ("02", "交过真正在业务里跑起来的系统",
         "装在自己机房、过得了持牌机构的审计、权限按级别分开、每一步都留痕，这些都有真实项目可查"),
        ("03", "底座能力已经建好，你少花一遍钱",
         "SSW 已有成熟的 AI 底座，标准功能不重复收软件授权费；落到你这里收的是部署、配置与定制"),
        ("04", "售前和交付是同一批人，全程不换班",
         "驻场工程师模式；杭州做交付、悉尼做技术把关；方案、开发、上线全程是同一批人"),
        ("05", "不锁死你：AI 模型可换、数据不出门、交付物归你",
         "AI 模型可以选国外的、国产的，也可以装在自己机房，交付物归你；"
         "交付后你具备独立运行能力，如需我们继续运营和扩场景，按季度另计"),
    ]
    y = 112
    for num, t, b in reasons:
        rect(s, 36, y, 648, 44, fill=CARD, line=BORDER)
        rect(s, 36, y, 3, 44, fill=RED)
        txt(s, 54, y, 34, 44, [P(num, 16, True, RED)], anchor="m")
        txt(s, 96, y + 7, 580, 18, [P(t, 11, True, DARK)], anchor="m")
        txt(s, 96, y + 25, 580, 16, [P(b, 8, False, MID, ls=1.05)], anchor="m")
        y += 48
    txt(s, 36, 352, 648, 20,
        [P("行业理解、工程能力、现成底座，三样同时具备的不多。", 10.5, True, RED)])


def psel(s):
    base(s, page="57")
    header(s, "选型参考", "什么样的企业，适合现在做",
           "对得上号的，几个月就能看到第一个结果")
    heads = ["对得上号，可以现在做", "先别急着做（先把前置条件解决掉）"]
    cols = [
        ["有一块别人拿不走的家底：配方、工艺、案例、客户资料",
         "有重复又耗人工的活：写报告、查资料、答问、审合同",
         "能指定一位业务负责人，也有明确的预算",
         "愿意先做一个场景，不要求全公司一次铺开"],
        ["资料整理成本较高，需要先评估",
         "想直接买一个通用 AI 工具，装上就见效",
         "期望三个月就替代掉一半员工",
         "一年内会换系统或大改组织，业务边界还在动"],
    ]
    xs = [36, 368]
    for i in range(2):
        rect(s, xs[i], 116, 316, 190, fill=CARD if i == 0 else WHITE, line=BORDER)
        rect(s, xs[i], 116, 316, 3, fill=RED)
        card_head(s, xs[i] + 16, 130, 284, heads[i], RED, 9.5)
        bullet_list(s, xs[i] + 16, 158, 284, cols[i], gap_y=32,
                    size=8.5, color=DARK if i == 0 else MID)
    rect(s, 36, 318, 648, 52, fill=TINT, line=BORDER)
    rect(s, 36, 318, 3, 52, fill=RED)
    txt(s, 52, 318, 620, 26,
        [P("判断一个场景能不能做，只看三件事：资料够不够、指标量不量得出来、有没有人负责。",
           10, True, RED)], anchor="m")
    txt(s, 52, 344, 620, 20,
        [P("三件齐了就能立项；缺哪一件，我们就先补哪一件，补完了再动。", 9, False, DARK)],
        anchor="m")


def p28(s):
    base(s, page="58")
    header(s, "合作模式", "分阶段推进",
           "四条腿走路，从哪一步开始都行")
    stages = [
        ("阶段一", "咨询规划", "摸清现状、排场景、算清账，出一份可报批的落地方案", "6 份报告 + 立项汇报材料"),
        ("阶段二", "小样验证", "用真实数据小范围试一遍，校准预期与指标", "验证报告 + 现场演示"),
        ("阶段三", "一期建设", "按功能清单建资料库，上线首批场景", "交付可用系统"),
        ("阶段四", "运营演进", "持续调优、整理资料、把做法复制到新场景", "驻场支持 · 方法交接"),
    ]
    for i, (ph, name, body, out) in enumerate(stages):
        x = 36 + i * 164
        rect(s, x, 116, 156, 202, fill=CARD, line=BORDER)
        rect(s, x, 116, 156, 3, fill=RED)
        txt(s, x + 12, 130, 132, 18, [P(ph, 8.5, True, MID)])
        txt(s, x + 12, 152, 132, 24, [P(name, 13, True, DARK)], anchor="m")
        txt(s, x + 12, 184, 132, 90, [P(body, 8.5, False, DARK, ls=1.28)])
        rect(s, x + 12, 280, 132, 1, fill=BORDER)
        txt(s, x + 12, 288, 132, 26, [P(out, 8.5, True, RED, ls=1.1)])
    notes = [
        ("委托方式", "四个阶段可分别委托，SSW 服务费按实际投入的驻场工程师人天计算"),
        ("第三方费用", "服务器、机房和模型调用费用由客户直接支付"),
        ("交接目标", "让你的团队学会与 AI 助手一起工作，并将做法用到下一个场景"),
    ]
    for i, (label, body) in enumerate(notes):
        y = 326 + i * 16
        txt(s, 36, y, 72, 15, [P(label, 9, True, MID)])
        txt(s, 115, y, 569, 15, [P(body, 9, False, MID)])


def p29(s):
    """正文结束页：浅灰底，红色只做块状点缀（CTA 红块 + 白卡红顶条）。后面两页是附录。"""
    base(s, bg="E9E9EC", page="60")
    W8 = "FFE8E8"                        # 红块上的次级白
    txt(s, 36, 34, 300, 16, [P("SSW AI 业务", 9, False, MID)])
    txt(s, 36, 60, 420, 64, [P("谢谢", 48, True, DARK)])
    txt(s, 36, 132, 420, 24, [P("你最想把哪项工作交给 AI？", 14, False, MID)])
    # CTA：一块红，白字
    rect(s, 36, 176, 420, 150, fill=WHITE, line=BORDER)
    rect(s, 36, 176, 4, 150, fill=RED)
    txt(s, 56, 192, 384, 30, [P("预约一次 60 分钟的沟通", 17, True, RED)])
    txt(s, 56, 228, 384, 34, [P(
        "聊一条最想改善的业务流程，我们一起判断第一个场景值不值得做。",
        9, False, DARK, ls=1.3)])
    rect(s, 56, 268, 384, 1, fill=BORDER)
    txt(s, 56, 278, 384, 40, [P(
        "带上三样就够：一个真实的业务问题 · 一份典型业务资料 · 一位业务负责人",
        9, False, MID, ls=1.3)])
    # 右侧白卡：联系我们
    rect(s, 470, 136, 214, 190, fill=WHITE, line=BORDER)
    rect(s, 470, 136, 214, 3, fill=RED)
    txt(s, 486, 152, 182, 16, [P("联系我们", 9, True, RED)])
    rect(s, 486, 172, 182, 1, fill=BORDER)
    txt(s, 486, 184, 182, 24, [P("www.ssw.cn", 13, True, DARK)])
    txt(s, 486, 210, 182, 24, [P("+86 571 8517 8910", 13, True, DARK)])
    rect(s, 486, 240, 182, 1, fill=BORDER)
    txt(s, 486, 250, 182, 66, [P(
        "正式诊断、项目范围与费用另行确认。技术细节见后两页附录。",
        7.5, False, MID, ls=1.32)])
def a01(s):
    base(s, page="A1")
    header(s, "附录一　术语对照", "同一个东西，两套说法", "给 IT 同事的对照页")
    items = [
        ("Agent / 智能体", "AI 助手 · 数字员工", "能自己拆任务、调工具的 AI 程序"),
        ("RAG", "查资料并给出处", "先查公司资料，再让 AI 据此回答"),
        ("GraphRAG / 知识图谱", "资料关系网", "把产品、认证、客户之间的关系连起来查"),
        ("向量库 / Embedding", "资料索引", "按意思也能搜到，不只看字面"),
        ("切片 / Chunking", "长文档怎么切开", "切得好才查得准，这项工作直接决定效果"),
        ("OCR / VLM", "扫描件与图片转文字", "把不能搜的文件变成能搜的"),
        ("MCP Server / 开放 API", "开放接口", "让别的系统和外部 AI 也能调用"),
        ("私有化部署", "装在自己机房", "数据不出公司网络"),
        ("SSO / AD", "用现有账号登录", "不新建一套用户名密码"),
        ("幻觉率", "答错 / 编造的比例", "AI 项目必须专测的质量指标"),
        ("BM25 + 语义双路", "关键词 + 说人话两种搜法", "两种一起用，找得又全又准"),
        ("token 用量", "AI 使用量", "按字数 / 次数计算，直接对应成本"),
    ]
    xs = [36, 367]
    for i, (tech, biz, note) in enumerate(items):
        x = xs[i // 6]
        y = 112 + (i % 6) * 38
        rect(s, x, y, 317, 34, fill=CARD, line=BORDER)
        txt(s, x + 12, y + 4, 293, 14,
            [R([(tech, True, RED), ("  →  ", False, LIGHT), (biz, True, DARK)], 8)])
        txt(s, x + 12, y + 19, 293, 14, [P(note, 7.5, False, MID)])
    txt(s, 36, 344, 648, 18,
        [P("本页是业务侧的通俗对照，不替代技术口径；词项以技术方案确认为准。", 8.5, False, MID)])


def a02(s):
    base(s, page="A2")
    header(s, "附录二　部署与集成", "给 IT 同事的细节页",
           "装法、算力、模型来源、集成运维")
    cards = [
        ("部署方式", ["放在云上：上线快、运维轻，适合不敏感的资料",
                   "装在自己机房：用你自己的服务器，数据不出公司",
                   "混着来：敏感资料在本地跑，通用问答走外部通道"]),
        ("计算方案", ["方案 A：买自己的服务器，一次性投入 + 按年折旧",
                   "方案 B：通常无需自购推理服务器，仍需评估部署、接入与集成成本",
                   "选型阶段给出六个维度的对比：投入·年度成本·性能·扩展·合规·运维"]),
        ("AI 模型从哪来", ["国外：Azure OpenAI、Anthropic Claude（需先确认能否合规连外网）",
                   "国产：通义千问、智谱、DeepSeek 等主流商用接口",
                   "装在自己机房：Qwen 等开源模型"]),
        ("集成与运维", ["身份：用公司现有账号登录，权限继承既有文件体系",
                    "系统：与 PLM / LIMS / ERP / OA 的数据对接（工作量单独评估）",
                    "工程：自动发布、异常告警、备份与灾难恢复"]),
    ]
    for i, (t, bs) in enumerate(cards):
        cx = 36 + (i % 2) * 331
        cy = 112 + (i // 2) * 120
        rect(s, cx, cy, 317, 112, fill=CARD, line=BORDER)
        rect(s, cx, cy, 317, 3, fill=RED)
        txt(s, cx + 14, cy + 12, 289, 20, [P(t, 11, True, RED)], anchor="m")
        bullet_list(s, cx + 14, cy + 38, 289, bs, gap_y=24, size=8, mark_w=12)
    txt(s, 36, 352, 648, 18,
        [P("选服务器、装在自己机房、自动发布与运维，SSW 可以做实施，也可以只做方案评审。",
           8.5, False, MID)])


# ---- section dividers ------------------------------------------------------
def divider(s, page, part, title, purpose, bullets, bsize=9.5, bstart=224, bstep=29):
    base(s, bg=CARD, topbar=False, footer=False)
    rect(s, 0, 0, 10, H, fill=RED)
    txt(s, 65, 76, 300, 24, [P(part, 10, False, MID)])
    rect(s, 65, 112, 79, 3, fill=RED)
    txt(s, 62, 126, 430, 46, [P(title, 30, True, DARK)])
    if purpose:
        txt(s, 62, 186, 470, 44, [P(purpose, 11, False, MID, ls=1.36)])
    y = bstart
    for b in bullets:
        rect(s, 62, y + 10, 6, 2, fill=RED)
        txt(s, 78, y, 520, bsize * 1.9, [P(b, bsize, False, DARK, ls=1.2)])
        y += bstep
    txt(s, 36, 383, 648, 14, [P(page, 7.5, False, LIGHT)], align="r")


def d1(s):
    divider(s, "03", "第一部分", "时代背景与公司",
            "这三十年，中国 B2B 的模式换了四代，痛点始终没变；我们也在这条路上",
            [])


def d2(s):
    divider(s, "10", "第二部分", "客户案例",
            "五个行业、五种做法，AI 落在哪条流程上、客户拿到什么结果",
            [])


def d3(s):
    divider(s, "30", "第三部分", "AI 中台：基础设施",
            "这一层基础设施由我们建好、交到你手上，后面所有 AI 场景都站在它上面跑",
            ["它不是概念，是一层实实在在的系统，装在你选定的环境里",
             "没有它约等于每接一个新场景都要重搭一遍地基",
             "它管四件事：账号与权限、用量与成本、助手的组合、公共设施",
             "上面 4 个案例，底层用的是同一套可复用能力"])


def d4(s):
    divider(s, "34", "第四部分", "我们提供哪些服务",
            "建在中台之上、我们对外提供的服务；横向四类能力，纵向按你的业务场景落地",
            ["横向四类：知识库搭建、智能体定制、系统集成、驻场陪跑与运营",
             "纵向 13 类场景：销售与客户、研发、生产、职能、经营决策……",
             "这四类建在同一条技术线上，标准能力不重复建设"])


def d5(s):
    divider(s, "51", "第六部分", "如何落地",
            "六步走完一个场景：做什么、为什么做、投入和产出",
            ["第 1–3 步：排场景 → 打底座 → 把资料装进去",
             "第 4–6 步：落到岗位上 → 上线交接 → 长期运营",
             "不是每个场景都要先建完整中台：小范围试点可以只做知识库加一个助手",
             "每一步的周期与投入都是区间，先走第 1 步再决定后面"])


def d6(s):
    divider(s, "55", "第七部分", "合作模式",
            "为什么找我们、什么样的企业适合现在做、四种合作方式",
            ["为什么找我们：五条可以逐条核实的理由",
             "对号入座：什么样的企业适合现在做，什么情况先别急",
             "四个阶段可以分段委托，按投入的驻场工程师人天计费"])


def c41(s):
    base(s, page="23")
    case_badge(s, 4)
    header(s, "案例四 · 某牛肉供应链企业", "从 13 人手动拉群，到 4 人管 1 万个群",
           "某牛肉供应链企业 · 销售全链路 AI 助手")
    rect(s, 36, 116, 400, 230, fill=CARD, line=BORDER)
    card_head(s, 50, 132, 372, "项目简介", RED, 9.5)
    txt(s, 50, 152, 372, 42, [P(
        "这家牛肉供应链企业的客户是餐饮企业和商超。\n获客、转化和接单交给 "
        "AI 助手后，运营团队从 13 人减到 4 人。", 8.5, False, DARK, ls=1.25)])
    card_head(s, 50, 200, 372, "要解决的问题", RED, 9.5)
    txt(s, 50, 220, 372, 46, [P(
        "下游客户分散，全靠销售人工一个个拉群、手动跟单；客户资料、沟通过程和报价记录都留在"
        "销售个人手里，人一走客户就带走；报价、催单、下单都要人盯，规模一大就撑不住；"
        "新人没人带，上手慢。", 8, False, DARK, ls=1.2)])
    card_head(s, 50, 272, 372, "我们做了什么", RED, 9.5)
    txt(s, 50, 292, 372, 46, [P(
        "把获客、转化、订单三段接过来：AI 自动建群、自动建档、自动跟进；报价与合同自动出稿，"
        "账期这类冲突点单独标出来给审批人看；群里一句指令就能下单，拆单核价、查库存、定交期，"
        "缺货延期自动上报。", 8, False, DARK, ls=1.2)])
    rect(s, 444, 116, 240, 148, fill=WHITE, line=BORDER)
    txt(s, 458, 124, 212, 16, [P("项目一览", 9, True, RED)])
    y = 144
    for k, v in [("行业", "食品供应链 · B 端食材"), ("项目类型", "销售全链路 AI 助手"),
                 ("规模", "1 万个客户群 · 4 人运营"), ("交付方式", "半年建成 · 边跑边加")]:
        txt(s, 458, y, 212, 10, [P(k, 7, False, LIGHT)])
        txt(s, 458, y + 10, 212, 16, [P(v, 8.5, True, DARK, ls=1.05)])
        y += 29
    rect(s, 444, 272, 240, 74, fill=TINT, line=BORDER)
    txt(s, 456, 278, 216, 40, [P("近 4 倍", 26, True, RED)], anchor="m")
    txt(s, 456, 322, 216, 22, [P("人均月产出从 60 万元增至 220 万元", 8, False, DARK, ls=1.1)])


def c42(s):
    base(s, page="24")
    header(s, "案例四 · 某牛肉供应链企业", "获客、转化、订单，三段 AI 渗透",
           "同一批人，管得住的客户群从 62 个变成 1 万个")
    cols = [(36, 96, "环节"), (140, 190, "原来怎么做"), (338, 228, "现在怎么做"), (574, 110, "结果")]
    for x, w, h in cols:
        rect(s, x, 116, w, 24, fill=RED, line=RED)
        txt(s, x + 8, 116, w - 16, 24, [P(h, 9, True, WHITE)], anchor="m")
    rows = [
        ("获客", "销售人工一个个拉群、发资料，加满就停",
         "AI 助手自动建群、自动发产品资料与活动，一个人管一批群",
         "62 个群 → 1 万个群"),
        ("转化", "跟单靠销售个人本事，报价靠互相砍价",
         "沟通自动建档、跟进自动提醒；报价与合同自动出稿，冲突点单独标出",
         "人均月产出\n60 万元 → 220 万元"),
        ("订单", "群里口头下单，人工核价、查库存、定交期",
         "群里一句指令下单，拆单核价、查库存、定交期；缺货延期自动上报",
         "4 人管住 1 万个群"),
    ]
    y = 146
    for i, vals in enumerate(rows):
        fill = CARD if i % 2 == 0 else WHITE
        for (x, w, _), v in zip(cols, vals):
            rect(s, x, y, w, 68, fill=fill, line=BORDER)
            if x == 36:
                sz, bold, col = 9.5, True, DARK
            elif x == 574:
                sz, bold, col = 8.5, True, RED
            else:
                sz, bold, col = 8, False, DARK
            txt(s, x + 8, y, w - 16, 68, [P(v, sz, bold, col, ls=1.18)], anchor="m")
        y += 68
    txt(s, 36, 356, 648, 20,
        [P("客户资料、沟通过程和报价记录都沉淀回公司，人走了客户还在。", 9.5, True, RED)])




def result_card(s, tag, title, sub, rows, extra, foot):
    """统一的结果卡页：4 行等高的案例对照 + 配套机制一行 + 结论一行。"""
    header(s, tag, title, sub)
    rect(s, 36, 106, 96, 24, fill=RED, line=RED)
    txt(s, 44, 106, 80, 24, [P("维度", 8.5, True, WHITE)], anchor="m")
    rect(s, 132, 106, 552, 24, fill=RED, line=RED)
    txt(s, 140, 106, 536, 24, [P("内容", 8.5, True, WHITE)], anchor="m")
    y = 130
    h = 48
    for i, (label, body) in enumerate(rows):
        fill = CARD if i % 2 == 0 else WHITE
        rect(s, 36, y, 96, h, fill=fill, line=BORDER)
        txt(s, 44, y, 80, h, [P(label, 9, True, RED)], anchor="m")
        rect(s, 132, y, 552, h, fill=fill, line=BORDER)
        txt(s, 140, y, 536, h, [P(body, 8.5, False, DARK, ls=1.24)], anchor="m")
        y += h
    txt(s, 36, y + 6, 648, 16, [P(extra, 8, False, MID)])
    txt(s, 36, 356, 648, 18, [P(foot, 9.5, True, RED)])


def p23(s):
    base(s, page="14")
    result_card(
        s, "案例一 · 杭州某化工公司", "AI 带来的改变",
        "原来的问题、AI 改在哪里、拿到什么结果",
        [("原来", "研发新材料时，几年前做过的相似配方又得从零走一遍；外文技术文献只有少数人看得懂；"
                 "各部门各自用通用聊天工具，接不进公司数据"),
         ("改造后", "资料统一入库并结构化，接到研发日常：研发新材料先查历史配方与试验记录，"
                  "每条答案都标出出自哪份文件"),
         ("实际结果", "知识库与研发问答入口已交付并投入使用；提效数字尚未统计"),
         ("项目状态", "已交付（知识库 + 研发问答入口）")],
        "配套机制：按人、按部门分权；答案带出处；哪些资料不能出公司、能不能连外网都已确认。",
        "后面的每个场景都可以单独立项，也可以整体交给同一批人往下做。")


def p09(s):
    base(s, page="17")
    result_card(
        s, "案例二 · 宁波某外贸工艺品厂家", "问题与成果",
        "资料以前只能存着，现在能查、能答、能拿来用",
        [("原来", "提案本事装在销售个人脑子里，人一走就没了；产品资料、认证、专利散在服务器和邮箱里，"
                 "只能按文件名搜，找不到、找不准"),
         ("改造后", "产品资料自动提取入库 → 会说话的搜索、答案带出处 → 传张图搜同款 → "
                  "一句话出带图提案与多语言文案"),
         ("实际结果", "提案时间缩短 80%+；7 个模块全面升级；产品、认证、专利、客户关系连成一张网"),
         ("项目状态", "已上线（一期产品库 + 二期 AI 升级均已交付）")],
        "配套机制：分几层加、不推倒重来；专测答错率与找全率；模型优先放自己机房；每一步操作留痕。",
        "一期打好的产品库，二期直接在上面加了 AI——顺序对了，钱不用花两遍。")


def p14(s):
    base(s, page="22")
    result_card(
        s, "案例三 · 北京某证券机构", "AI 带来的成果",
        "报告快八成，合规流程一样没少",
        [("原来", "尽调报告一份人工写 2–3 天，占掉分析师大量时间；研报、募集书散在各处；"
                 "持牌机构审计要求严，AI 不是想上就能上"),
         ("改造后", "整套系统装进客户自己的机房：一句话起草尽调报告，按模板填满、每条结论标出处；"
                  "知识推广走三级审批；操作全程留痕"),
         ("实际结果", "尽调报告生成时间缩短 80%（2–3 天 → 数小时）；5 个模块一体交付"),
         ("项目状态", "已上线（装在自己机房，完全在内网运行）")],
        "管得住 AI 的机制：五类角色分权 · 知识三级审批 · 操作留痕不可篡改 · 用量按人 / 团队 / 公司三级配额。",
        "AI 用得上，还得管得住，这是金融行业能落地的前提。")


def c43(s):
    base(s, page="25")
    result_card(
        s, "案例四 · 某牛肉供应链企业", "运营团队从 13 人减至 4 人，人均月产出增至 220 万元",
        "某牛肉供应链企业 · 获客、转化、订单三段交给 AI",
        [("原来", "客户靠人工一个个拉群、手动跟单；客户资料、报价记录和沟通记录都留在销售个人手里，"
                 "人一走客户就带走"),
         ("改造后", "获客、转化、订单三段交给 AI 助手：自动建群建档、报价与合同自动出稿、"
                  "群里一句指令下单"),
         ("实际结果", "运营人数 13 → 4；客户群 62 → 1 万个；人均月产出从 60 万元增至 220 万元"),
         ("项目状态", "已交付（半年建成，仍在往上加场景）")],
        "配套机制：客户资料、沟通过程和报价记录都沉淀回公司；缺货延期自动上报；四块可以分别上线。",
        "人没有增加，管得住的客户群扩大了 160 倍。")


def tok(s):
    base(s, page="40")
    header(s, "我们的服务 · AI 用量", "AI用量，我们帮你采购",
           "我们拿到的是渠道价，比你直接去官方买便宜")
    rect(s, 36, 108, 648, 46, fill=TINT, line=BORDER)
    rect(s, 36, 108, 4, 46, fill=RED)
    txt(s, 52, 108, 620, 46, [P(
        "AI 用量是长期运行成本的大头。官方直购没有折扣，量再大也是原价。"
        "我们拿到的是渠道价，比官方直购便宜，"
        "也省掉一家家开账号、对账、垫资的麻烦。", 8.5, False, DARK, ls=1.24)],
        anchor="m")
    cols = [(36, 380, "模型档位"), (416, 268, "给到你的价")]
    for x, w, h in cols:
        rect(s, x, 166, w, 24, fill=RED, line=RED)
        txt(s, x + 10, 166, w - 20, 24, [P(h, 9, True, WHITE)], anchor="m")
    rows = [
        ("国产主流模型（Kimi / 智谱 / DeepSeek 这类）", "官方价 7 折", False),
        ("视频生成类模型（有月消费保底）", "官方价 9 折", False),
        ("大额额度买断（Claude 这类）", "官方价 8.5 折", False),
        ("原厂直连 Key（部分国际模型）", "按官方价代管代付", True),
    ]
    y = 190
    for i, (name, price, plain) in enumerate(rows):
        fill = CARD if i % 2 == 0 else WHITE
        rect(s, 36, y, 380, 38, fill=fill, line=BORDER)
        txt(s, 48, y, 356, 38, [P(name, 9, False, DARK, ls=1.15)], anchor="m")
        rect(s, 416, y, 268, 38, fill=fill, line=BORDER)
        txt(s, 428, y, 244, 38,
            [P(price, 11, True, MID if plain else RED, ls=1.1)], anchor="m")
        y += 38
    txt(s, 36, 346, 648, 30,
        [P("每月一张清单：用多少、单价多少，一目了然。",
           9.5, True, RED),
         P("以上价格仅作参考，最终以实际报价为准。", 7.5, False, MID, sa=2)])


def sec(s):
    base(s, page="59")
    header(s, "数据安全", "安全：我们的四条承诺",
           "不用讲技术细节，这四条就够")
    cards = [
        ("① 数据不外流",
         "装在你自己的机房，数据不出公司网络；放在云上的，数据存在你专属的环境里，不与其他客户混用。"),
        ("② 谁能看，你定",
         "按人、按部门、按项目分权；配方、客户名单、财务数据各归其位；人一走，权限一键收回。"),
        ("③ 要出去，先打码",
         "需要连外网的问答，只发送打码后的内容；核心资料可以完全不连外网，就在自己机房跑。"),
        ("④ 做过什么，查得到",
         "谁在什么时候查了什么、导出了什么，全程留痕，日志可以导出来给你。"),
    ]
    for i, (t, b) in enumerate(cards):
        cx = 36 + (i % 2) * 331
        cy = 116 + (i // 2) * 120
        rect(s, cx, cy, 317, 106, fill=CARD, line=BORDER)
        rect(s, cx, cy, 317, 3, fill=RED)
        txt(s, cx + 16, cy + 16, 289, 24, [P(t, 12, True, RED)], anchor="m")
        txt(s, cx + 16, cy + 46, 289, 48, [P(b, 9, False, DARK, ls=1.26)])
    txt(s, 36, 352, 648, 20,
        [P("以上每一条都可以写进合同。我们不做只靠口头保证的安全。", 10, True, RED)])


def net(s):
    base(s, page="41")
    header(s, "我们的服务 · 网络接入", "合规访问境外互联网",
           "调用境外模型、海外 SaaS、海外资料，都需要这条线路")
    rect(s, 36, 108, 648, 46, fill=TINT, line=BORDER)
    rect(s, 36, 108, 4, 46, fill=RED)
    txt(s, 52, 108, 620, 46, [P(
        "有些场景必须访问境外：调用海外模型、用海外 SaaS 工具、查海外资料。"
        "这条线路要走合规通道，不是拉根宽带就能解决的。我们帮你把线路开通好，"
        "你按月付带宽费。", 8.5, False, DARK, ls=1.24)], anchor="m")
    cols = [(36, 380, "项目"), (416, 268, "说明")]
    for x, w, h in cols:
        rect(s, x, 166, w, 24, fill=RED, line=RED)
        txt(s, x + 10, 166, w - 20, 24, [P(h, 9, True, WHITE)], anchor="m")
    rows = [
        ("带宽", "10M", False),
        ("价格", "约 4,000 元 / 月", False),
        ("计费", "按月付费，按带宽计，不按流量", False),
        ("适用场景", "调用境外模型、海外 SaaS 工具、海外资料检索", True),
    ]
    y = 190
    for i, (k, v, plain) in enumerate(rows):
        fill = CARD if i % 2 == 0 else WHITE
        rect(s, 36, y, 380, 38, fill=fill, line=BORDER)
        txt(s, 48, y, 356, 38, [P(k, 9, False, DARK, ls=1.15)], anchor="m")
        rect(s, 416, y, 268, 38, fill=fill, line=BORDER)
        txt(s, 428, y, 244, 38,
            [P(v, 11, True, MID if plain else RED, ls=1.1)], anchor="m")
        y += 38
    txt(s, 36, 346, 648, 30,
        [P("线路怎么合规开通、要不要备案，我们在方案阶段一并给你。", 9.5, True, RED),
         P("以上为常规配置的参考价格，实际按线路与开通方案确定。", 7.5, False, MID, sa=2)])


def c5a(s):
    base(s, page="26")
    case_badge(s, 5)
    header(s, "案例五 · 某高校与科研机构", "高校已有 AI 能力，我们协助完成产品化交付",
           "高校与科研场景系列项目 · 产品规划、工程实施与部署交付")
    rect(s, 36, 116, 400, 230, fill=CARD, line=BORDER)
    card_head(s, 50, 132, 372, "项目简介", RED, 9.5)
    txt(s, 50, 152, 372, 44, [P(
        "校方已经具备 AI 研究与服务能力，并开展了早期应用原型建设。随着项目推进，"
        "建设范围逐步扩展到客户端、业务后台、AI 服务、数据管理和部署环境，"
        "需要有人持续统筹产品决策与工程交付。", 8.5, False, DARK, ls=1.2)])
    card_head(s, 50, 208, 372, "要解决的问题", RED, 9.5)
    txt(s, 50, 228, 372, 46, [P(
        "研究人员关注模型效果，业务使用方关注操作体验，开发团队需要清晰的规则和接口。"
        "原型进入实际使用后，任务进度、异常恢复、历史记录、权限和部署维护等问题陆续出现，"
        "需要结合产品目标和工程影响一起处理。", 8, False, DARK, ls=1.18)])
    card_head(s, 50, 274, 372, "我们做了什么", RED, 9.5)
    txt(s, 50, 294, 372, 46, [P(
        "SSW 顾问担任整体产品负责人，统筹需求澄清、版本优先级和跨团队协作，"
        "同时深入参与前后端开发、数据建模及 AI 服务集成，持续跟进客户环境中的部署与运行问题。",
        8, False, DARK, ls=1.18)])
    rect(s, 444, 116, 240, 148, fill=WHITE, line=BORDER)
    txt(s, 458, 124, 212, 16, [P("项目一览", 9, True, RED)])
    y = 144
    for k, v in [("行业", "高校与科研"), ("服务类型", "AI 应用产品化与工程交付"),
                 ("参与方式", "产品负责人统筹 + 开发与集成"),
                 ("建设范围", "业务平台、轻量应用、AI 集成与部署能力")]:
        txt(s, 458, y, 212, 10, [P(k, 7, False, LIGHT)])
        txt(s, 458, y + 10, 212, 17, [P(v, 8.5, True, DARK, ls=1.05)])
        y += 29
    rect(s, 444, 272, 240, 74, fill=TINT, line=BORDER)
    rect(s, 444, 272, 4, 74, fill=RED)
    txt(s, 458, 280, 212, 16, [P("核心价值", 9, True, RED)])
    txt(s, 458, 298, 212, 44, [P(
        "帮助校方将已有研究能力落实为完整的用户流程，并补齐持续使用和维护所需的工程能力。",
        8, False, DARK, ls=1.18)])


def c5m(s):
    base(s, page="27")
    header(s, "案例五 · 某高校与科研机构", "项目挑战与应对",
           "高校有研究与开发资源，正式应用还需要交付经验")
    rect(s, 36, 104, 648, 24, fill=TINT, line=BORDER)
    rect(s, 36, 104, 4, 24, fill=RED)
    txt(s, 48, 104, 624, 24,
        [P("我们在产品决策、工程实现和运行环境之间持续协调。", 9, False, DARK)], anchor="m")
    items = [
        ("1. 不同参与方对产品理解不同",
         "SSW 组织需求讨论，明确用户任务、业务规则和验收预期，将讨论结论落实为开发安排。"),
        ("2. 一个需求变更会影响多个模块",
         "SSW 评估页面、数据、接口和权限之间的关系，确定版本范围与实施顺序，协调相关模块同步调整。"),
        ("3. 原型能展示能力，完整使用流程仍需完善",
         "SSW 补齐文件上传、任务进度、结果核对、历史记录、报告下载和连接恢复，让用户能够完成实际工作。"),
        ("4. AI 服务、前后端与数据链路需要共同运行",
         "SSW 明确服务职责、数据交换和异常处理，沿请求、文件、数据及配置定位集成问题。"),
        ("5. 学校的服务器和网络条件各不相同",
         "SSW 根据实际环境推进镜像打包、离线导出、部署配置、数据库备份和故障排查。"),
    ]
    y = 136
    for i, (t, d) in enumerate(items):
        rect(s, 36, y, 648, 38, fill=CARD if i % 2 == 0 else WHITE, line=BORDER)
        rect(s, 36, y, 3, 38, fill=RED)
        txt(s, 48, y + 3, 620, 15, [P(t, 9.5, True, RED)])
        txt(s, 48, y + 20, 620, 15, [P(d, 8.5, False, DARK)])
        y += 38
    rect(s, 36, 334, 648, 34, fill=TINT, line=BORDER)
    rect(s, 36, 334, 4, 34, fill=RED)
    txt(s, 48, 334, 624, 34, [R([("协作分工　", True, RED),
        ("校方持续提供研究能力与业务判断；SSW 将这些输入转化为明确的产品行为、工程实现和交付安排。",
         False, DARK)], 8.5)], anchor="m")


def c5b(s):
    base(s, page="28")
    header(s, "案例五 · 某高校与科研机构", "建设成果与合作价值",
           "从交互原型，逐步建设到业务平台与轻量应用")
    rect(s, 36, 104, 648, 24, fill=TINT, line=BORDER)
    rect(s, 36, 104, 4, 24, fill=RED)
    txt(s, 48, 104, 624, 24,
        [P("成果体现在产品完整性、系统协作和部署能力上。", 9, False, DARK)], anchor="m")
    cards = [
        ("产品形态逐步完善",
         "系列项目从早期交互原型发展出业务平台与轻量应用，覆盖文档阅读与编辑、"
         "结构化结果展示、任务管理和报告下载等使用环节。"),
        ("系统基础随业务演进",
         "围绕业务变化调整数据关系、历史记录与权限规则，协调前端、后台和数据库保持一致，"
         "为后续功能扩展建立基础。"),
        ("交付延伸到运行环境",
         "围绕客户实际条件建设部署配置、版本管理、备份与故障排查能力，"
         "使开发工作能够衔接客户现场的运行要求。"),
        ("探索多种应用入口",
         "逐步形成 Web、桌面客户端及移动端基础实现，复用共享协议与核心逻辑，"
         "并推进多语言界面适配。"),
    ]
    for i, (t, d) in enumerate(cards):
        cx = 36 + (i % 2) * 331
        cy = 132 + (i // 2) * 96
        rect(s, cx, cy, 317, 88, fill=CARD, line=BORDER)
        rect(s, cx, cy, 317, 3, fill=RED)
        txt(s, cx + 16, cy + 14, 289, 20, [P(t, 10.5, True, RED)], anchor="m")
        txt(s, cx + 16, cy + 38, 289, 46, [P(d, 8.5, False, DARK, ls=1.24)])
    rect(s, 36, 328, 648, 46, fill=TINT, line=BORDER)
    rect(s, 36, 328, 4, 46, fill=RED)
    txt(s, 52, 334, 620, 16, [P("这一案例说明了什么", 9, True, RED)])
    txt(s, 52, 352, 620, 18, [P(
        "高校提供研究成果和开发力量。咨询团队协调各方、确定产品范围并负责交付，"
        "让学校能使用、扩展和维护这些应用。",
        8, False, DARK, ls=1.16)])


def c5c(s):
    base(s, page="29")
    result_card(
        s, "案例五 · 某高校与科研机构", "交付了什么",
        "原来的问题、我们改了什么、现在跑成什么样",
        [("原来", "研究成果与应用系统之间有距离；各方对产品的理解各不相同；开发任务彼此依赖；"
                 "部署环境各有条件限制"),
         ("改造后", "先把「用户从哪进、提交什么、怎么看懂结果、出问题怎么办、下一步做什么」定义清楚；"
                  "项目会议做需求澄清与方案取舍，把四方诉求收敛到同一份产品方案"),
         ("实际结果", "从交互原型长成完整业务平台 + 轻量化应用；Web、桌面端、移动端基础实现可用，"
                  "支持多语言界面；两种形态复用同一套核心逻辑与服务，尚未做量化统计"),
         ("项目状态", "已交付（平台与轻量应用均在运行）")],
        "配套机制：权限与账号规则在界面入口、操作按钮、后台接口和数据查询四处保持一致；"
        "账号、密码校验、输入检查、错误提示逐项维护。",
        "高校项目的共同点：先跑通一个明确场景，再逐步扩展。")


def ice(s):
    base(s, page="05")
    header(s, "为什么是现在", "企业 AI 落地，是一座冰山",
           "水面上的应用只是露出的一角，决定成败的都在水下")
    s.shapes.add_picture(ASSETS + "/iceberg_clean.png", Pt(36), Pt(106), width=Pt(314))

    items = [
        ("01", "只买了工具，没换流程",
         "AI 做成散装的单点功能，没嵌进获客、销售、履约的真实流程。"),
        ("02", "数据孤岛，知识没沉淀",
         "订单、客户、历史数据散在各自系统里，没沉淀成业务档案和知识库。"),
        ("03", "流程乱、权限乱，跑不通",
         "内部流程不标准化、职责不清，AI 接不到正确的流程。"),
        ("04", "没有陪跑，上线就没人管",
         "一次性交付后没人跟业务一起调，业务一变就停在试用期。"),
    ]
    y = 106
    for num, t, d in items:
        rect(s, 370, y, 314, 52, fill=CARD, line=BORDER)
        rect(s, 370, y, 3, 52, fill=RED)
        txt(s, 384, y + 6, 292, 12, [R([(num + "　", True, RED), (t, True, DARK)], 8.5)])
        txt(s, 384, y + 22, 292, 26, [P(d, 8, False, MID, ls=1.18)])
        y += 58

    txt(s, 370, 348, 314, 18,
        [P("AI 落地不是买工具，是把水下这五层建起来。", 9, True, RED)])


def era(s):
    base(s, page="04")
    header(s, "时代背景", "中国 B2B 三十年：模式在变，痛点没变",
           "同质化价格战、获客成本高企、数据孤岛、增长靠人——老问题从未真正解决")
    eras = [
        ("1990s–2005", "信息黄页时代", "线上名录、信息墙合起来", "同质化与价格战"),
        ("2005–2015", "平台交易时代", "B2B 平台兴起、交易撮合", "获客成本攀升"),
        ("2015–2020", "SaaS 与产业互联网", "工具与数字化进场", "数据孤岛与定制化陷阱"),
        ("2020-至今", "数字化与 AI 时代", "工具普及、AI 爆发", "业务数据散在各处\n企业难以留存和复用"),
    ]
    for i, (yrs, name, did, left) in enumerate(eras):
        x = 36 + i * 165
        rect(s, x, 118, 151, 212, fill=CARD, line=BORDER)
        rect(s, x, 118, 151, 3, fill=RED)
        txt(s, x + 8, 132, 135, 16, [P(yrs, 11, True, RED, align="c")], align="c")
        txt(s, x + 8, 154, 135, 20, [P(name, 11, True, DARK, align="c")], align="c")
        rect(s, x + 24, 182, 103, 1, fill=BORDER)
        txt(s, x + 12, 194, 127, 44, [P(did, 8, False, MID, ls=1.24, align="c")], align="c")
        txt(s, x + 8, 250, 135, 12, [P("留下的问题", 7, False, LIGHT, align="c")], align="c")
        txt(s, x + 12, 266, 127, 40,
            [P(left, 9, True, RED, ls=1.24, align="c")], align="c")
    txt(s, 36, 344, 648, 20,
        [P("客户记录留在个人手里，订单和研发资料散在不同系统，换个项目又得从头找", 10, True, RED, align="c")], align="c")


def platform(s, page, title, sub, scenes, cans, bottom, lead="我们提供落地咨询与实施",
            tag="国内 AI 平台落地"):
    base(s, page=page)
    header(s, tag, title, sub)
    txt(s, 36, 104, 330, 16, [P("企业最常见的五个场景", 9, True, RED)])
    y = 126
    for i, (name, desc) in enumerate(scenes):
        rect(s, 36, y, 330, 42, fill=CARD if i % 2 == 0 else WHITE, line=BORDER)
        txt(s, 46, y + 4, 310, 14,
            [R([("%02d　" % (i + 1), True, RED), (name, True, DARK)], 8.5)])
        txt(s, 46, y + 21, 310, 16, [P(desc, 8, False, MID)])
        y += 42
    txt(s, 390, 104, 294, 16, [P("落地要过的三道坎", 9, True, RED)])
    y = 126
    for name, desc in cans:
        rect(s, 390, y, 294, 70, fill=TINT, line=BORDER)
        rect(s, 390, y, 3, 70, fill=RED)
        txt(s, 402, y + 7, 272, 14, [P(name, 8.5, True, RED)])
        txt(s, 402, y + 24, 272, 42, [P(desc, 8, False, DARK, ls=1.18)])
        y += 70
    rect(s, 36, 348, 648, 28, fill=TINT, line=BORDER)
    rect(s, 36, 348, 4, 28, fill=RED)
    txt(s, 50, 348, 620, 28,
        [R([(lead + "　", True, RED), (bottom, False, DARK)], 8.5)], anchor="m")


def pwb(s):
    platform(
        s, "43",
        "WorkBuddy：桌面上的 AI，也能进企业微信",
        "腾讯的办公智能体；企业版可公有云、专享 VPC，也能私有化部署",
        [("项目管理", "每 4 小时汇总群聊、会议和文档，自动同步到项目看板"),
         ("跨系统协作", "群里的订单自动整理成表，写进 ERP，进度再推回群里"),
         ("产品研发", "把几十个客户群里的需求与现场照片，汇成研发灵感看板"),
         ("客户服务", "从沟通里识别意向客户，自动出客户看得懂的分析报告"),
         ("门店标准化", "把零散讨论变成统一规范，AI 按规范检查装修图纸")],
        [("权益类型选定就不可切换",
          "企业版分「Buddy One 套件」和「企业微信联合套件」两条路，创建后不能换，选错只能重建。"),
         ("默认模型的数据会上云",
          "用端内默认模型时，输入输出要上传云端处理；敏感场景要接本地模型，或走专享、私有化。"),
         ("权限规则得有人来配",
          "文件保护路径、命令与网络规则、部门级策略优先级——平台给的是能力，规则要按你的组织设计。")],
        "选型 · 权限设计 · 连接器与私有 MCP · 技能搭建 · 上线陪跑")


def pdb(s):
    platform(
        s, "44",
        "豆包办公：站在飞书里的 AI",
        "字节的办公 AI；豆包 + 飞书 + 火山引擎三者打包，不是一个 SaaS",
        [("企业知识问答", "问制度、问流程、问项目背景，答案附参考来源"),
         ("会议与群聊", "群聊重点、会议纪要、待办提取、周报，直接生成"),
         ("周期性报告", "定时自动跑，成果自动存云盘并推送到群里"),
         ("表格与经营分析", "解析业务数据，出结论和可视化图表"),
         ("电脑与浏览器自动化", "模拟人工操作本地软件；长任务放云上跑，关机也不断")],
        [("权限就是 AI 的边界",
          "官方明确只读你有权看的内容。飞书里权限理不清，AI 答出来的东西就不可信。"),
         ("连接器授权不能分项",
          "第三方实测反馈：授权页要么全开、要么不连，没有中间档；合规沟通要提前谈。"),
         ("企业版还在内测",
          "目前没有公开定价、也没有私有化说明，报什么价要以商务口径为准。")],
        "权限梳理 · 知识库建设 · 连接器打通 · 定时报告搭建 · 上线陪跑")


def pqw(s):
    platform(
        s, "45",
        "千问办公：站在钉钉里的 AI",
        "阿里的办公 AI；原生载体是钉钉，不是一个独立 App",
        [("制度问答", "制度文档入库，回答带出处，敏感制度按人群隔离"),
         ("会议全链路", "纪要、待办、日程、邮件，在 IM 侧一次打通"),
         ("经营看板", "连生意参谋这类数据源，出看板并定时推送到钉钉群"),
         ("Office 产物", "直接出 PPT、Word、Excel，还能在线改"),
         ("合同、HR 与财务", "批量比对合同、解析简历、跑月度财务简报")],
        [("组织 ID 填错，席位发到空壳",
          "席位绑在组织 ID 上：要从钉钉组织里进，或先在阿里云建企业，权限不够就卡在这一步。"),
         ("知识库得另买",
          "企业知识库要单独购买百炼知识库，再在千问办公里装技能接入，这笔钱容易漏报。"),
         ("专网只有旗舰版有",
          "需要专网的金融、医疗和政务客户，须选用支持专网部署的旗舰版。")],
        "选型与组织规划 · 知识库建设 · 数据源对接 · 合同财务场景 · 上线陪跑")


def dpl(s):
    divider(s, "42", "第五部分", "AI 平台落地：国内与海外",
            "不必从零开始搭中台\n按企业的办公生态、研发流程和数据合规要求，选择合适的 AI 平台",
            [])


def pcmp(s):
    base(s, page="46")
    header(s, "国内 AI 平台落地", "三个平台怎么选：先看你的办公底座",
           "模型差距不大，真正的区别在「AI 站在哪个系统里」")
    cols = [(36, 140, "对比维度"), (176, 169, "WorkBuddy"), (345, 169, "豆包办公"),
            (514, 169, "千问办公")]
    for x, w, h in cols:
        rect(s, x, 106, w, 26, fill=RED, line=RED)
        txt(s, x + 8, 106, w - 16, 26,
            [P(h, 9, True, WHITE, align="c" if x > 40 else "l")], anchor="m",
            align="c" if x > 40 else "l")
    rows = [
        ("办公底座", None, ("", "企业微信 / 自建通讯录"), ("", "飞书"), ("", "钉钉")),
        ("私有化部署", "check", ("✓", "企业私有化版"), ("—", "无相关支持"), ("—", "无相关支持")),
        ("专网 / VPC", "check", ("✓", "企业专享版"), ("—", "无相关支持"), ("✓", "旗舰版才有")),
        ("企业知识库", "check", ("✓", "内置腾讯乐享"), ("✓", "内置飞书知识库"), ("△", "需另购百炼")),
        ("可换模型", "check", ("✓", "混元 + 5 家 + 本地"), ("△", "以豆包系为主"), ("△", "以千问系为主")),
        ("企业版状态", "check", ("✓", "已商用"), ("△", "内测中"), ("✓", "已商用")),
    ]
    y = 132
    for i, (label, kind, *cells) in enumerate(rows):
        fill = CARD if i % 2 == 0 else WHITE
        rect(s, 36, y, 140, 31, fill=fill, line=BORDER)
        txt(s, 44, y, 124, 31, [P(label, 8.5, True, DARK)], anchor="m")
        for (x, w, _), (mark, note) in zip(cols[1:], cells):
            rect(s, x, y, w, 31, fill=fill, line=BORDER)
            if kind == "check":
                color = RED if mark == "✓" else (MID if mark == "△" else LIGHT)
                txt(s, x + 10, y, w - 20, 31,
                    [R([(mark + "  ", True, color), (note, False, DARK if mark != "—" else MID)], 8.5)],
                    anchor="m")
            else:
                txt(s, x + 10, y, w - 20, 31, [P(note, 8.5, True, RED)], anchor="m")
        y += 31
    # 推荐场景
    rect(s, 36, y, 140, 50, fill=TINT, line=BORDER)
    rect(s, 36, y, 3, 50, fill=RED)
    txt(s, 44, y, 124, 50, [P("推荐用在", 8.5, True, RED)], anchor="m")
    recs = ["已用企业微信；要私有化或强数据管控（制造、医药、政企）",
            "已用飞书；要电脑与浏览器自动化、长任务托管（互联网、消费）",
            "已用钉钉；要合同审阅、HR 与财务批处理（贸易、专业服务）"]
    for (x, w, _), t in zip(cols[1:], recs):
        rect(s, x, y, w, 50, fill=TINT, line=BORDER)
        txt(s, x + 10, y, w - 20, 50, [P(t, 8, False, DARK, ls=1.22)], anchor="m")


def case_badge(s, idx):
    txt(s, 519.8, 14.8, 149.8, 16,
        [P("新案例   %02d / 05" % idx, 11, True, RED, align="c")], align="c")


def nav(s):
    base(s, page="02")
    txt(s, 36, 18, 300, 20, [P("SSW AI 业务  /  阅读导航", 10, False, MID)])
    txt(s, 36, 43, 648, 36, [P("按你关心的问题，直接跳到相关内容", 24, True, DARK)])
    rect(s, 36, 87, 40, 3, fill=RED)
    txt(s, 85, 80, 576, 17, [P("点击章节或案例标签，即可跳转；后续页码均对应本版本", 9, False, MID)])
    secs = [
        ("01", "时代背景与公司", "企业痛点 · SSW 能力", "P03–09", 3),
        ("02", "客户案例", "五个行业的交付成果", "P10–29", 10),
        ("03", "AI 中台设施", "账号 · 数据 · 成本", "P30–33", 30),
        ("04", "我们提供哪些服务", "知识库 · 智能体 · 集成", "P34–41", 34),
        ("05", "国内外 AI 平台", "选型 · 落地 · 合规", "P42–50", 42),
        ("06", "如何落地", "实施六步 · 周期与投入", "P51–54", 51),
        ("07", "合作模式", "适用客户 · 合作阶段", "P55–58", 55),
        ("08", "数据安全", "权限 · 脱敏 · 留痕", "P59", 59),
    ]
    for i, (num, name, desc, rng, tgt) in enumerate(secs):
        x = 36 + (i % 4) * 165.6
        y = 114 + (i // 4) * 100
        _card = rect(s, x, y, 153, 86, fill=CARD, line=BORDER)
        _bar = rect(s, x, y, 3.5, 86, fill=RED)
        _t1 = txt(s, x + 12, y + 9, 30, 22, [P(num, 15, True, RED)])
        _t2 = txt(s, x + 12, y + 32, 131, 22, [P(name, 11, True, DARK)])
        _t3 = txt(s, x + 12, y + 56, 131, 16, [P(desc, 8, False, MID)])
        _t4 = txt(s, x + 98, y + 8, 43, 17, [P(rng, 8.5, False, LIGHT, align="r")], align="r")
        LINK_QUEUE.extend([(_card, tgt), (_bar, tgt), (_t1, tgt), (_t2, tgt), (_t3, tgt), (_t4, tgt)])
    txt(s, 36, 324, 106, 17, [P("案例快速跳转", 10, True, RED)])
    _cases = [("化工 · 知识库", 12), ("外贸 · 产品库", 15), ("金融 · 合规", 19),
              ("供应链 · 销售", 23), ("高校 · 科研", 26)]
    for i, (c, tgt) in enumerate(_cases):
        x = 136.8 + i * 109.5
        _r = rect(s, x, 320, 102, 25, fill=WHITE, line=BORDER)
        _t = txt(s, x, 320, 102, 25, [P(c, 9, True, DARK, align="c")], anchor="m", align="c")
        LINK_QUEUE.extend([(_r, tgt), (_t, tgt)])


def ov1(s):
    platform(
        s, "47",
        "ChatGPT：企业知识与日常工作的一站式 AI 助手",
        "OpenAI 企业 AI 工作空间；让知识查询、内容工作和数据分析落到岗位上",
        [("企业知识问答", "连接获授权的内部资料，回答问题并提供来源线索"),
         ("报告与文案", "整理文档、起草周报、售前材料与客户回复"),
         ("表格与数据分析", "分析业务数据，生成图表与初步结论"),
         ("跨系统资料查询", "通过经批准的应用获取项目、客户和知识数据"),
         ("部门常用助手", "将固定任务与企业规则做成可复用的工作流程")],
        [("权限先理清", "先确认员工身份、文件访问范围和应用授权，避免越权取数。"),
         ("数据与应用治理", "Business / Enterprise 默认不以业务数据训练模型；第三方应用仍须单独评估。"),
         ("使用地区有限制", "中国大陆不在 ChatGPT 公开支持地区；须核对人员位置、服务条款和数据出境。")],
        "场景梳理 · 企业知识接入 · 权限与应用配置 · 试点培训",
        lead="SSW 可提供落地支持",
        tag="海外 AI 平台落地")


def ov2(s):
    platform(
        s, "48",
        "Codex：把 AI 编码助手接入真实研发流程",
        "OpenAI 软件工程 Agent；读代码、改代码、跑测试，输出工程师可审查的变更",
        [("代码库理解", "梳理项目模块、调用关系与现有实现逻辑"),
         ("缺陷定位和修复", "复现问题、修改代码并补充回归验证"),
         ("功能开发", "按明确需求实现功能，生成可评审的代码差异"),
         ("测试与质量", "补充单元测试、运行检查、协助代码审阅"),
         ("多任务协作", "在隔离工作区并行推进任务，由工程师审核合并")],
        [("代码权限与密钥", "代码仓库、密钥、外部网络和运行环境都要事先设定访问边界。"),
         ("变更必须审查", "AI 生成的代码须经过测试、人工审核和发布流程。"),
         ("账号与运行范围", "云端与本地功能受方案和管理设置限制；同样需要核对 OpenAI 服务支持地区。")],
        "代码库评估 · 工作流接入 · 测试规范 · 交付质量验收",
        lead="SSW 可提供落地支持",
        tag="海外 AI 平台落地")


def ov3(s):
    platform(
        s, "49",
        "Claude：让长文档与复杂任务形成可用成果",
        "Anthropic 企业 AI 助手；适合知识阅读、分析、写作与部分研发协作",
        [("合同与制度梳理", "对照长文档、定位差异、整理需确认的条款"),
         ("研究资料分析", "整理报告、访谈记录与专业资料，形成结论草稿"),
         ("知识驱动写作", "依据项目资料起草提案、说明文档和业务总结"),
         ("跨团队协同", "围绕共享资料组织需求、问题和交付文档"),
         ("研发辅助", "配合 Claude Code 理解仓库、修改代码与运行测试")],
        [("授权与资料边界", "上传文件或连接系统前，核实使用权限、留存规则与数据处理条款。"),
         ("关键内容要复核", "研究结论、合同意见与代码变更需要业务专家或工程师验收。"),
         ("地区与部署限制", "Claude 受支持地区与主体资格约束；不能承诺中国大陆可直接开通。")],
        "需求评估 · 企业知识接入 · 工作流设计 · 试点与培训",
        lead="SSW 可提供落地支持",
        tag="海外 AI 平台落地")


def ovcmp(s):
    base(s, page="50")
    header(s, "海外 AI 平台落地", "三个海外 AI 平台怎么选",
           "按工作内容比较用途，再核对服务地区、组织身份与数据流向")
    cols = [(36, 140, "对比维度"), (176, 169, "ChatGPT"), (345, 169, "Codex"),
            (514, 169, "Claude")]
    for x, w, h in cols:
        rect(s, x, 106, w, 26, fill=RED, line=RED)
        txt(s, x + 8, 106, w - 16, 26,
            [P(h, 9, True, WHITE, align="c" if x > 40 else "l")], anchor="m",
            align="c" if x > 40 else "l")
    rows = [
        ("核心定位", "企业通用 AI 工作台", "软件工程 Agent", "知识与复杂任务助手"),
        ("主要入口", "工作区 · 应用", "桌面 · IDE · 云任务", "工作区 · Claude Code"),
        ("典型场景", "问答 · 报告 · 数据", "编码 · 修复 · 测试", "文档 · 协作 · 研发"),
        ("企业治理", "账号、权限和应用管理", "仓库、环境与审批", "身份、权限与留痕"),
        ("本地部署", "不承诺本地私有版", "本地与云任务需评估", "不承诺本地私有版"),
        ("大陆直接使用", "不在公开支持地区", "同受 OpenAI 地区限制", "不在公开支持地区"),
    ]
    y = 132
    for i, r in enumerate(rows):
        fill = CARD if i % 2 == 0 else WHITE
        rect(s, 36, y, 140, 31, fill=fill, line=BORDER)
        txt(s, 44, y, 124, 31, [P(r[0], 8.5, True, DARK)], anchor="m")
        for (x, w, _), t in zip(cols[1:], r[1:]):
            rect(s, x, y, w, 31, fill=fill, line=BORDER)
            txt(s, x + 10, y, w - 20, 31, [P(t, 8.5, False, DARK)], anchor="m")
        y += 31
    rect(s, 36, y, 140, 50, fill=TINT, line=BORDER)
    rect(s, 36, y, 3.5, 50, fill=RED)
    txt(s, 44, y, 124, 50, [P("推荐用在", 8.5, True, RED)], anchor="m")
    recs = ["合规的海外知识办公、分析与协作", "合规的海外研发团队和代码管理", "合规的海外知识工作及研发协作"]
    for (x, w, _), t in zip(cols[1:], recs):
        rect(s, x, y, w, 50, fill=TINT, line=BORDER)
        txt(s, x + 10, y, w - 20, 50, [P(t, 8, False, DARK, ls=1.22)], anchor="m")


def closing(s):
    base(s, bg=CARD, topbar=False, footer=False)
    rect(s, 0, 0, 10, H, fill=RED)
    txt(s, 65, 76, 300, 24, [P("SSW AI 业务", 10, False, MID)])
    rect(s, 65, 112, 79, 3, fill=RED)
    txt(s, 62, 150, 580, 50, [P("谢谢", 34, True, DARK)])
    txt(s, 65, 215, 580, 30, [P("你最想把哪项工作交给 AI？", 15, False, MID)])
    txt(s, 36, 383, 648, 14, [P("63", 7.5, False, LIGHT)], align="r")


# ================================================================ assemble
prs = Presentation()
prs.slide_width = Emu(9144000)
prs.slide_height = Emu(5143500)
BLANK = prs.slide_layouts[6]

BUILDERS = [p01, nav, d1, era, ice, p02, p03, p04, p05,
            d2, p06,                    # 第二部分分隔页 = 五个案例的引子
            p21, p22, p23,              # 案例一 化工企业 · 知识库与研发辅助
            p07, p08, p09, p10,         # 案例二 外贸厂家 · 产品库与提案辅助
            p11, p12, p13, p14,         # 案例三 证券公司 · 合规
            c41, c42, c43,              # 案例四 供应链企业 · 获客 / 转化 / 订单
            c5a, c5m, c5b, c5c,         # 案例五 高校与科研场景（4 页）
            d3, p24, p25, p26,
            d4, sp1, sp2, sp3, sp5, sp6, tok, net,
            dpl,                        # 第五部分 分隔页：AI 平台落地
            pwb, pdb, pqw, pcmp,         # 国内 AI 平台落地（3 页 + 对比页）
            ov1, ov2, ov3, ovcmp,        # 海外 AI 平台落地（3 页 + 对比页）
            d5, sv1, sv2, sv3,
            d6, p27, psel, p28,
            sec,                        # 数据安全（一页讲清）
            p29,                        # NEXT STEP 行动页
            a01, a02]
# 已下架备查，不再进 BUILDERS：
#   p15/p16/p17 涂料案例、p18/p19/p20 汽车零部件案例（Ken 2026-10-13 决定砍掉）
#   sp4 已改名 c43 升为案例四成果页

for fn in BUILDERS:
    fn(prs.slides.add_slide(BLANK))

# 案例段（P11 案例简介 + P12–P29 五个案例）左侧红条与页脚导航点
CASE_CHROME = {11: None,
               12: 1, 13: 1, 14: 1,
               15: 2, 16: 2, 17: 2, 18: 2,
               19: 3, 20: 3, 21: 3, 22: 3,
               23: 4, 24: 4, 25: 4,
               26: 5, 27: 5, 28: 5, 29: 5}
for _i, _sl in enumerate(prs.slides, 1):
    if _i in CASE_CHROME:
        case_chrome(_sl, CASE_CHROME[_i])

# 目录 / 阅读导航的真超链接（必须等全部页面建完再解析）
for _sh, _tgt in LINK_QUEUE:
    _n = len(prs.slides._sldIdLst)
    if 1 <= _tgt <= _n:
        _sh.click_action.target_slide = prs.slides[_tgt - 1]

prs.save(OUT)
print("slides:", len(prs.slides.__iter__.__self__._sldIdLst))
print("saved:", OUT)
