# -*- coding: utf-8 -*-
"""CIS 销售经理 · 华勤手机业务 · 转正答辩 PPT"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

# Warm paper / ink — avoid purple AI look
BG = RGBColor(0xF6, 0xF4, 0xF1)
INK = RGBColor(0x1B, 0x2A, 0x3A)
MUTED = RGBColor(0x5C, 0x6B, 0x7A)
ACCENT = RGBColor(0xB8, 0x95, 0x4A)
SALES = RGBColor(0x2A, 0x5A, 0x73)
SOFT = RGBColor(0xE8, 0xF0, 0xF4)
SOFT_G = RGBColor(0xE8, 0xEE, 0xE6)
SOFT_A = RGBColor(0xF4, 0xEE, 0xE0)
SOFT_R = RGBColor(0xF4, 0xEB, 0xE8)
CARD = RGBColor(0xFF, 0xFF, 0xFF)
LINE = RGBColor(0xD8, 0xD3, 0xCB)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
DARK = RGBColor(0x16, 0x24, 0x32)

FONT = "Microsoft YaHei"


def set_run_font(run, size, bold=False, color=INK, name=FONT):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = name
    rPr = run._r.get_or_add_rPr()
    for tag in ("a:ea", "a:cs"):
        el = rPr.find(qn(tag))
        if el is None:
            el = rPr.makeelement(qn(tag), {})
            rPr.insert(0, el)
        el.set("typeface", name)


def add_bg(slide, color=BG):
    sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    sh.fill.solid()
    sh.fill.fore_color.rgb = color
    sh.line.fill.background()
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, Inches(0.07))
    bar.fill.solid()
    bar.fill.fore_color.rgb = ACCENT
    bar.line.fill.background()
    return sh


def add_footer(slide, page, total=16):
    box = slide.shapes.add_textbox(Inches(0.55), Inches(7.18), Inches(10.2), Inches(0.25))
    p = box.text_frame.paragraphs[0]
    run = p.add_run()
    run.text = "转正答辩  ·  CIS 销售经理  ·  华勤手机业务"
    set_run_font(run, 10, False, MUTED)
    num = slide.shapes.add_textbox(Inches(11.6), Inches(7.18), Inches(1.2), Inches(0.25))
    p = num.text_frame.paragraphs[0]
    p.alignment = PP_ALIGN.RIGHT
    run = p.add_run()
    run.text = f"{page} / {total}"
    set_run_font(run, 10, False, MUTED)


def add_title(slide, title, subtitle=""):
    box = slide.shapes.add_textbox(Inches(0.55), Inches(0.22), Inches(12.2), Inches(0.48))
    p = box.text_frame.paragraphs[0]
    run = p.add_run()
    run.text = title
    set_run_font(run, 26, True, INK)
    if subtitle:
        sub = slide.shapes.add_textbox(Inches(0.55), Inches(0.70), Inches(12.2), Inches(0.32))
        p = sub.text_frame.paragraphs[0]
        run = p.add_run()
        run.text = subtitle
        set_run_font(run, 13, False, MUTED)


def round_rect(slide, l, t, w, h, fill, line=None, radius=0.08):
    sh = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, l, t, w, h)
    sh.adjustments[0] = radius
    sh.fill.solid()
    sh.fill.fore_color.rgb = fill
    if line is None:
        sh.line.fill.background()
    else:
        sh.line.color.rgb = line
        sh.line.width = Pt(1)
    return sh


def add_text(slide, l, t, w, h, lines, size=13, color=INK, bold=False, align=PP_ALIGN.LEFT, leading=6):
    box = slide.shapes.add_textbox(l, t, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    tf.clear()
    for i, text in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.space_after = Pt(leading)
        run = p.add_run()
        run.text = text
        is_bold = bold if not isinstance(text, tuple) else text[1]
        txt = text if not isinstance(text, tuple) else text[0]
        if isinstance(text, tuple):
            run.text = txt
            set_run_font(run, size, is_bold, color)
        else:
            set_run_font(run, size, bold, color)
    return box


def notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text


def set_table_font(table, size=12, header=True):
    for i, row in enumerate(table.rows):
        for cell in row.cells:
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            for p in cell.text_frame.paragraphs:
                p.space_before = Pt(2)
                p.space_after = Pt(2)
                for run in p.runs:
                    set_run_font(run, size, bold=(header and i == 0), color=INK if i else INK)
                    if header and i == 0:
                        run.font.color.rgb = WHITE


def shade_header(table, color=SALES):
    for cell in table.rows[0].cells:
        cell.fill.solid()
        cell.fill.fore_color.rgb = color


def shade_row(table, idx, color):
    for cell in table.rows[idx].cells:
        cell.fill.solid()
        cell.fill.fore_color.rgb = color


# =============================================================================
# 1 Cover
# =============================================================================
s = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(s, DARK)
gold = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(0.14), prs.slide_height)
gold.fill.solid()
gold.fill.fore_color.rgb = ACCENT
gold.line.fill.background()

add_text(s, Inches(0.9), Inches(1.55), Inches(11), Inches(0.4),
         ["试用期转正答辩"], 14, ACCENT, True)
add_text(s, Inches(0.9), Inches(2.05), Inches(11.5), Inches(1.1),
         ["把华勤做成可复制的手机 CIS 账号"], 34, WHITE, True)
add_text(s, Inches(0.9), Inches(3.25), Inches(11), Inches(0.5),
         ["CIS 销售经理  ·  负责华勤手机业务  ·  销售公司 CIS 芯片"], 16, RGBColor(0xC5, 0xD0, 0xD8))

round_rect(s, Inches(0.9), Inches(4.2), Inches(2.4), Inches(0.08), ACCENT, radius=0.0)

meta = [
    ("汇报人", "CIS 销售经理（华勤手机）"),
    ("客户定位", "华勤 · 手机 ODM · CIS 导入与定点"),
    ("汇报日期", "2026 年 8 月"),
]
for i, (k, v) in enumerate(meta):
    left = Inches(0.9 + i * 3.9)
    add_text(s, left, Inches(4.55), Inches(3.6), Inches(0.3), [k], 11, ACCENT, True)
    add_text(s, left, Inches(4.9), Inches(3.6), Inches(0.5), [v], 13, WHITE, False)

notes(s, "开场：先说明岗位——在华勤卖 CIS，不是卖一颗料，而是在品牌 AVL 框架下，把公司芯片推进到可量产。试用期重点是摸清决策链、建立窗口、把项目漏斗跑起来。")

# =============================================================================
# 2 Agenda
# =============================================================================
s = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(s)
add_title(s, "汇报大纲", "约 15 分钟  ·  先讲清楚做了什么，再判断是否胜任，最后看下一阶段怎么打")
add_footer(s, 2)

agenda = [
    ("01", "试用期工作总结", "客户经营 · 项目导入 · 内部协同",
     "华勤决策链、项目漏斗、送样验证、销售与 FAE 闭环"),
    ("02", "试用期自我评价", "胜任点 · 短板 · 匹配度",
     "用过程指标评价：窗口、漏斗、闭环，而不是尚未发生的定点金额"),
    ("03", "未来规划与展望", "90 天  ·  12 个月  ·  诉求",
     "把华勤做成样板账号，形成可复制到其他手机 ODM 的打法"),
]
for i, (num, title, sub, desc) in enumerate(agenda):
    top = Inches(1.25 + i * 1.8)
    round_rect(s, Inches(0.55), top, Inches(12.2), Inches(1.6), CARD, LINE)
    nbox = s.shapes.add_textbox(Inches(0.85), top + Inches(0.4), Inches(1.3), Inches(0.7))
    p = nbox.text_frame.paragraphs[0]
    run = p.add_run()
    run.text = num
    set_run_font(run, 28, True, ACCENT)
    add_text(s, Inches(2.3), top + Inches(0.28), Inches(9.8), Inches(0.4), [title], 20, INK, True)
    add_text(s, Inches(2.3), top + Inches(0.72), Inches(9.8), Inches(0.3), [sub], 13, SALES, False)
    add_text(s, Inches(2.3), top + Inches(1.05), Inches(9.8), Inches(0.35), [desc], 13, MUTED, False)

# =============================================================================
# 3 Role
# =============================================================================
s = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(s)
add_title(s, "岗位认知：在华勤卖 CIS，实际在做什么", "先对齐评价口径——试用期的产出是「可推进的项目」而不是「已经量产的金额」")
add_footer(s, 3)

points = [
    ("三层客户", "品牌方管规格与 AVL，华勤管落地节奏与第二源建议，模组厂管工程验证。销售要同时看见这三层，不能只对接一个窗口。"),
    ("一条链路", "情报 → 规格匹配 → 送样 → 模组/IQ 验证 → EVT/DVT → 定点 → 量产。试用期必须把前半段跑通，后半段用漏斗管理。"),
    ("销售产出", "项目情报、选型卡位、送样与问题闭环、价格与供给口径、客情与内部协同。芯片参数由 FAE 支撑，销售负责把节点往前推。"),
]
for i, (t, d) in enumerate(points):
    left = Inches(0.55 + i * 4.15)
    round_rect(s, left, Inches(1.25), Inches(3.95), Inches(4.55), CARD, LINE)
    tag = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, Inches(1.25), Inches(0.12), Inches(4.55))
    tag.fill.solid()
    tag.fill.fore_color.rgb = ACCENT if i == 0 else SALES
    tag.line.fill.background()
    add_text(s, left + Inches(0.35), Inches(1.55), Inches(3.4), Inches(0.5), [t], 18, INK, True)
    add_text(s, left + Inches(0.35), Inches(2.2), Inches(3.4), Inches(3.2), [d], 14, MUTED, False, leading=10)

notes(s, "强调：华勤是 ODM。很多时候品牌已经定了格科/思特威/OV。我们的切入点是第二源、供给风险对冲、以及新机型窗口。")

# =============================================================================
# 4 Overview
# =============================================================================
s = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(s)
add_title(s, "试用期工作总览", "四条线并行：摸清客户、卡位项目、建立机制、补齐专业")
add_footer(s, 4)

cards = [
    ("01  客户经营", SOFT, SALES,
     ["完成华勤手机 CIS 决策链梳理", "建立采购 / 项目 / 技术三线窗口", "明确 ODM 角色与品牌 AVL 约束", "形成稳定沟通节奏与拜访纪要"]),
    ("02  项目导入", SOFT_A, ACCENT,
     ["输出在研机型与摄像头规格地图", "2M 等规格对标成熟竞品卡位", "组织送样，协同 FAE 验证闭环", "项目分级（A/B/C）进入漏斗管理"]),
    ("03  内部协同", SOFT_G, RGBColor(0x4A, 0x6F, 0x5A),
     ["把客户语言翻译成内部任务", "与 FAE / 质量 / 计划对齐接口", "周报红黄绿，不报喜藏忧", "补齐导入包：规格、setting、FAQ"]),
    ("04  专业内功", RGBColor(0xF3, 0xEE, 0xF0), RGBColor(0x6B, 0x4E, 0x5A),
     ["吃透手机向 CIS 规格与竞品", "掌握送样、OTP、模组验证节奏", "学习华勤项目节点（EVT/DVT/MP）", "能独立完成常规拜访与纪要"]),
]
for i, (title, fill, accent, lines) in enumerate(cards):
    r, c = divmod(i, 2)
    left = Inches(0.55 + c * 6.35)
    top = Inches(1.22 + r * 2.8)
    round_rect(s, left, top, Inches(6.1), Inches(2.55), fill, LINE)
    add_text(s, left + Inches(0.3), top + Inches(0.22), Inches(5.5), Inches(0.4), [title], 16, accent, True)
    add_text(s, left + Inches(0.3), top + Inches(0.75), Inches(5.5), Inches(1.6),
             ["·  " + x for x in lines], 13, INK, False, leading=4)

# =============================================================================
# 5 Timeline
# =============================================================================
s = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(s)
add_title(s, "试用期节奏", "按「先结构、后项目、再机制」推进，避免一上来只追一颗料")
add_footer(s, 5)

# connector line
line = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1.3), Inches(1.55), Inches(10.7), Inches(0.06))
line.fill.solid()
line.fill.fore_color.rgb = LINE
line.line.fill.background()

months = [
    ("第 1 阶段", "融入与摸底",
     ["产品、竞品、导入流程学习", "华勤组织与决策链访谈", "识别在研机型与现有供应商", "确定主对接与备份窗口"]),
    ("第 2 阶段", "卡位与送样",
     ["规格对照表（我司 vs 成熟料）", "目标机型送样与模组协同", "FAE 陪访，问题 48h 闭环", "报价与交期口径内部对齐"]),
    ("第 3 阶段", "漏斗与复盘",
     ["项目 A/B/C 分级周更新", "验证进展与风险升级", "客情复盘：谁决策、谁影响", "形成华勤账号作战手册初稿"]),
]
for i, (ph, title, items) in enumerate(months):
    left = Inches(0.55 + i * 4.2)
    dot = s.shapes.add_shape(MSO_SHAPE.OVAL, left + Inches(1.7), Inches(1.42), Inches(0.32), Inches(0.32))
    dot.fill.solid()
    dot.fill.fore_color.rgb = ACCENT if i == 2 else SALES
    dot.line.fill.background()
    round_rect(s, left, Inches(2.05), Inches(3.95), Inches(4.55), CARD, LINE)
    add_text(s, left + Inches(0.3), Inches(2.25), Inches(3.35), Inches(0.3), [ph], 12, ACCENT, True)
    add_text(s, left + Inches(0.3), Inches(2.6), Inches(3.35), Inches(0.45), [title], 18, INK, True)
    add_text(s, left + Inches(0.3), Inches(3.2), Inches(3.35), Inches(3.0),
             ["·  " + x for x in items], 14, MUTED, False, leading=8)

# =============================================================================
# 6 Decision chain
# =============================================================================
s = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(s)
add_title(s, "客户经营：华勤决策链", "ODM 不是单一采购关系。每个角色要的东西不同，销售动作必须拆开")
add_footer(s, 6)

rows = [
    ["角色", "他关心什么", "试用期我做了什么"],
    ["项目经理", "节点、风险、能否赶上 EVT/DVT", "对齐机型节奏，不把未验证的料推进关键路径"],
    ["摄像头硬件", "规格匹配、脚位/封装、替代可行性", "输出对照表，明确能否第二源、改板成本"],
    ["影像 / 调试", "IQ、线性、噪声、暗角、验证结论", "拉 FAE 进现场，问题闭环进纪要"],
    ["采购", "价格、交期、供给安全、商务条款", "建立窗口，同步公司供给与报价原则"],
    ["质量 / 计划", "可靠性、客诉、上量节奏", "提前暴露封装/回流/OTP 等导入风险"],
]
table = s.shapes.add_table(len(rows), 3, Inches(0.55), Inches(1.2), Inches(12.2), Inches(5.35)).table
table.columns[0].width = Inches(2.2)
table.columns[1].width = Inches(4.6)
table.columns[2].width = Inches(5.4)
for r, row in enumerate(rows):
    for c, val in enumerate(row):
        cell = table.cell(r, c)
        cell.text = val
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        for p in cell.text_frame.paragraphs:
            for run in p.runs:
                set_run_font(run, 13, bold=(r == 0 or c == 0), color=WHITE if r == 0 else INK)
shade_header(table, SALES)
for r in range(1, 6):
    if r % 2 == 0:
        shade_row(table, r, RGBColor(0xF3, 0xF1, 0xED))

notes(s, "讲这一页时点出：华勤能推荐，但品牌 AVL 不过，项目仍进不了。所以 rel 品牌窗口是下一阶段任务，试用期先把华勤内部盘清楚是正确顺序。")

# =============================================================================
# 7 Funnel
# =============================================================================
s = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(s)
add_title(s, "项目推进：漏斗而不是「跟一颗料」", "试用期目标是把漏斗建起来，并让 A 类项目进入验证，而不是承诺立刻定点")
add_footer(s, 7)

steps = [
    ("1", "情报", "机型、像素档、现有供应商、开案窗口"),
    ("2", "匹配", "规格/封装/接口能否进短名单"),
    ("3", "送样", "芯片+模组路径，资料包齐全"),
    ("4", "验证", "IQ / 可靠性 / 产线，问题闭环"),
    ("5", "定点", "商务+质量+计划同时过"),
    ("6", "量产", "交期、品质、变更受控"),
]
for i, (n, t, d) in enumerate(steps):
    left = Inches(0.4 + i * 2.15)
    round_rect(s, left, Inches(1.3), Inches(2.05), Inches(2.35), CARD, LINE)
    add_text(s, left + Inches(0.12), Inches(1.45), Inches(1.8), Inches(0.4), [n], 20, ACCENT, True, align=PP_ALIGN.CENTER)
    add_text(s, left + Inches(0.12), Inches(1.9), Inches(1.8), Inches(0.4), [t], 16, INK, True, align=PP_ALIGN.CENTER)
    add_text(s, left + Inches(0.12), Inches(2.4), Inches(1.8), Inches(1.05), [d], 11, MUTED, False, align=PP_ALIGN.CENTER, leading=3)
    if i < 5:
        arr = s.shapes.add_textbox(left + Inches(1.92), Inches(2.15), Inches(0.28), Inches(0.3))
        p = arr.text_frame.paragraphs[0]
        run = p.add_run()
        run.text = ">"
        set_run_font(run, 14, True, ACCENT)

round_rect(s, Inches(0.55), Inches(3.9), Inches(6.0), Inches(2.7), SOFT, LINE)
add_text(s, Inches(0.8), Inches(4.1), Inches(5.5), Inches(0.4), ["试用期实际停在哪"], 16, SALES, True)
add_text(s, Inches(0.8), Inches(4.6), Inches(5.5), Inches(1.8), [
    "·  1–3 步：已跑通方法和窗口",
    "·  第 4 步：目标项目进入/准备进入验证",
    "·  5–6 步：纳入漏斗管理，不提前报喜",
    "·  每周更新状态，红灯当周升级",
], 14, INK, False, leading=6)

round_rect(s, Inches(6.8), Inches(3.9), Inches(5.95), Inches(2.7), SOFT_A, LINE)
add_text(s, Inches(7.05), Inches(4.1), Inches(5.5), Inches(0.4), ["主攻卡位"], 16, ACCENT, True)
add_text(s, Inches(7.05), Inches(4.6), Inches(5.5), Inches(1.8), [
    "·  手机前置 / 功能摄等 2M 规格档",
    "·  对标格科、思特威等成熟料做第二源",
    "·  话术：供给弹性、导入服务、风险对冲",
    "·  不硬刚像素纸面参数，打「可落地」",
], 14, INK, False, leading=6)

# =============================================================================
# 8 Results table
# =============================================================================
s = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(s)
add_title(s, "试用期成果盘点", "用可核验的过程结果说话。定点与收入作为下一阶段指标，不在试用期夸大")
add_footer(s, 8)

rows = [
    ["工作项", "状态", "结果说明"],
    ["华勤手机 CIS 组织与决策链地图", "已完成", "覆盖项目 / 硬件 / 影像 / 采购 / 质量，避免单点窗口"],
    ["主对接 + 备份窗口", "已建立", "采购与项目双线，沟通节奏可周/双周维持"],
    ["在研机型摄像头规格与竞品卡位表", "已输出", "看清格科/思特威等在哪些像素档占位、窗口在哪"],
    ["目标规格送样与 FAE 协同", "进行中", "资料包、陪访、问题闭环机制已跑通，验证结论跟进中"],
    ["项目漏斗（A/B/C）与周报", "已建立", "红黄绿状态，内部对齐，不报喜藏忧"],
    ["华勤账号作战手册初稿", "已形成", "决策链、话术、导入节奏、风险清单可交接、可复制"],
    ["品牌方 AVL 窗口", "未打开", "已识别约束，作为转正后 90 天重点，不回避"],
]
table = s.shapes.add_table(len(rows), 3, Inches(0.55), Inches(1.18), Inches(12.2), Inches(5.5)).table
table.columns[0].width = Inches(4.2)
table.columns[1].width = Inches(1.6)
table.columns[2].width = Inches(6.4)
for r, row in enumerate(rows):
    for c, val in enumerate(row):
        cell = table.cell(r, c)
        cell.text = val
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        for p in cell.text_frame.paragraphs:
            p.alignment = PP_ALIGN.CENTER if c == 1 else PP_ALIGN.LEFT
            for run in p.runs:
                col = INK
                if r == 0:
                    col = WHITE
                elif c == 1:
                    if val == "已完成" or val == "已建立" or val == "已输出" or val == "已形成":
                        col = RGBColor(0x2F, 0x6B, 0x4F)
                    elif val == "进行中":
                        col = RGBColor(0x8A, 0x6A, 0x1A)
                    elif val == "未打开":
                        col = RGBColor(0x8B, 0x4A, 0x3C)
                set_run_font(run, 12, bold=(r == 0 or c == 1), color=col)
shade_header(table, SALES)
for r in range(1, len(rows)):
    if r % 2 == 0:
        shade_row(table, r, RGBColor(0xF3, 0xF1, 0xED))

# =============================================================================
# 9 Collaboration
# =============================================================================
s = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(s)
add_title(s, "协同方式：销售对节点负责，FAE 对技术结论负责", "试用期刻意避免「销售自己讲参数、FAE 自己跟客户」两张皮")
add_footer(s, 9)

left_items = [
    "拜访前：规格问题清单，不空跑",
    "现场：销售控议程，FAE 答技术",
    "当天：纪要含问题、责任人、截止日",
    "48 小时：技术问题给到书面回复",
    "升级：卡住节点当周进内部会",
]
right_items = [
    "客户要的是「能不能赶上这台机」",
    "内部要的是「问题是否可关闭」",
    "导入风险提前讲：封装、CRA、OTP、上电时序，不把问题留到验证后期",
    "竞品动态进周报，供产品与计划决策",
]
round_rect(s, Inches(0.55), Inches(1.25), Inches(6.05), Inches(5.3), CARD, LINE)
add_text(s, Inches(0.85), Inches(1.5), Inches(5.5), Inches(0.45), ["销售 × FAE 工作约定"], 18, SALES, True)
add_text(s, Inches(0.85), Inches(2.15), Inches(5.5), Inches(4.0),
         ["·  " + x for x in left_items], 15, INK, False, leading=10)

round_rect(s, Inches(6.85), Inches(1.25), Inches(5.95), Inches(5.3), CARD, LINE)
add_text(s, Inches(7.15), Inches(1.5), Inches(5.4), Inches(0.45), ["对内对外同一套事实"], 18, ACCENT, True)
add_text(s, Inches(7.15), Inches(2.15), Inches(5.4), Inches(4.0),
         ["·  " + x for x in right_items], 15, INK, False, leading=10)

# =============================================================================
# 10 Gaps
# =============================================================================
s = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(s)
add_title(s, "试用期未完成项与原因", "转正答辩把短板放桌上，比把漏斗画满更重要")
add_footer(s, 10)

gaps = [
    ("品牌 AVL 尚未突破", "华勤可推荐，但品牌短名单不打开，项目到不了定点。试用期优先把 ODM 侧盘清，品牌窗口列入 90 天计划。"),
    ("尚无量产定点可报", "手机 CIS 导入周期长于试用期。评价应用漏斗质量、验证进度、客情厚度，而不是用尚未发生的收入。"),
    ("商务深度不够", "价格、账期、供给承诺还需要与产品、计划形成华勤专用口径，避免现场口头承诺。"),
    ("技术陪访依赖 FAE", "常规规格已能讲，OTP/IQ 细节仍需 FAE。转正后要能独立完成 70% 的售前问答，复杂问题再升级。"),
]
for i, (t, d) in enumerate(gaps):
    r, c = divmod(i, 2)
    left = Inches(0.55 + c * 6.35)
    top = Inches(1.25 + r * 2.7)
    round_rect(s, left, top, Inches(6.1), Inches(2.45), CARD, LINE)
    mark = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, Inches(0.12), Inches(2.45))
    mark.fill.solid()
    mark.fill.fore_color.rgb = RGBColor(0x8B, 0x4A, 0x3C)
    mark.line.fill.background()
    add_text(s, left + Inches(0.4), top + Inches(0.28), Inches(5.4), Inches(0.45), [t], 16, INK, True)
    add_text(s, left + Inches(0.4), top + Inches(0.85), Inches(5.4), Inches(1.35), [d], 14, MUTED, False, leading=6)

# =============================================================================
# 11 Self eval
# =============================================================================
s = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(s)
add_title(s, "自我评价：六个维度", "评自己，也给评委一个打分框架")
add_footer(s, 11)

dims = [
    ("客户洞察", "强", "能把华勤三层结构讲清楚，知道决策链和品牌约束"),
    ("目标管理", "中+", "漏斗和周报已建立；A 类项目节点还要更细"),
    ("沟通协同", "强", "对内对外同一套事实，FAE 约定可执行"),
    ("专业能力", "中", "规格与导入流程已入门，IQ/OTP 仍需加强"),
    ("抗压担当", "强", "长周期、多角色，不把问题藏到月底"),
    ("结果导向", "中+", "过程结果扎实；经营结果要靠下一阶段验证"),
]
for i, (name, level, desc) in enumerate(dims):
    r, c = divmod(i, 3)
    left = Inches(0.55 + c * 4.15)
    top = Inches(1.25 + r * 2.7)
    round_rect(s, left, top, Inches(3.95), Inches(2.45), CARD, LINE)
    add_text(s, left + Inches(0.28), top + Inches(0.28), Inches(2.2), Inches(0.4), [name], 16, INK, True)
    pill = round_rect(s, left + Inches(2.55), top + Inches(0.32), Inches(1.1), Inches(0.36), SOFT_A, None, 0.3)
    add_text(s, left + Inches(2.55), top + Inches(0.34), Inches(1.1), Inches(0.34), [level], 12, ACCENT, True, align=PP_ALIGN.CENTER)
    add_text(s, left + Inches(0.28), top + Inches(0.9), Inches(3.4), Inches(1.3), [desc], 13, MUTED, False, leading=6)

# =============================================================================
# 12 Self conclusion
# =============================================================================
s = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(s)
add_title(s, "综合评价", "结论先讲：匹配这个岗位，转正后要用项目结果把过程能力兑现")
add_footer(s, 12)

round_rect(s, Inches(0.55), Inches(1.25), Inches(12.2), Inches(1.35), SOFT, LINE)
add_text(s, Inches(0.85), Inches(1.45), Inches(11.6), Inches(1.0),
         ["胜任判断：能在华勤这种多层决策的 ODM 账号上，把 CIS 销售做成「可管理的项目」，而不是个人关系或一次性送样。这是转正的基本盘。"],
         16, INK, False, leading=8)

cols = [
    (SOFT_G, "适合留下的理由",
     ["客户结构理解快，不把 ODM 当渠道批发", "执行闭环清楚：纪要、责任人、升级", "对数字诚实，不把漏斗当业绩", "愿意补专业，不与 FAE 抢结论"]),
    (SOFT_R, "必须补齐的短板",
     ["品牌侧渗透几乎为零", "独立商务谈判经验不足", "技术问答还不能独立覆盖", "A 类项目的周节点仍偏粗"]),
]
for i, (fill, t, items) in enumerate(cols):
    left = Inches(0.55 + i * 6.35)
    round_rect(s, left, Inches(2.85), Inches(6.1), Inches(3.7), fill, LINE)
    add_text(s, left + Inches(0.35), Inches(3.05), Inches(5.4), Inches(0.45), [t], 18, INK, True)
    add_text(s, left + Inches(0.35), Inches(3.6), Inches(5.4), Inches(2.6),
             ["·  " + x for x in items], 15, INK, False, leading=8)

notes(s, "这一页主动请评委按短板提问。转正不是完美，是方向对、闭环在、短板有计划。")

# =============================================================================
# 13 90 days
# =============================================================================
s = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(s)
add_title(s, "未来 90 天：四件必须发生的事", "转正后第一个季度，用节点而不是愿望考核自己")
add_footer(s, 13)

q = [
    ("01", "验证出结论", "至少 1 个目标项目完成模组/IQ 验证，形成书面结论（过 / 有条件过 / 不过及原因）。"),
    ("02", "品牌线索", "画出品牌—华勤—模组三方图，拿到至少 1 个品牌侧或品牌指定模组的有效窗口线索。"),
    ("03", "商务口径", "与产品、计划锁定华勤专用报价、交期、供给原则，现场不再口头承诺。"),
    ("04", "机制固化", "华勤项目月评：漏斗、风险、FAE 资源、下一步动作，形成公司可看见的账号经营。"),
]
for i, (n, t, d) in enumerate(q):
    top = Inches(1.22 + i * 1.35)
    round_rect(s, Inches(0.55), top, Inches(12.2), Inches(1.22), CARD, LINE)
    add_text(s, Inches(0.8), top + Inches(0.35), Inches(1.0), Inches(0.5), [n], 20, ACCENT, True)
    add_text(s, Inches(1.9), top + Inches(0.18), Inches(10.4), Inches(0.4), [t], 16, INK, True)
    add_text(s, Inches(1.9), top + Inches(0.6), Inches(10.4), Inches(0.5), [d], 13, MUTED, False)

# =============================================================================
# 14 12 months
# =============================================================================
s = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(s)
add_title(s, "未来 12 个月：把华勤做成样板账号", "样板的意思是：可复制方法，而不只是「我跟过一个客户」")
add_footer(s, 14)

year = [
    ("项目结果", "1 个项目进入试产或定点路径；2–3 个项目稳定在验证池；杜绝只有拜访没有下一节点的空账号。"),
    ("关系结构", "从单一窗口升级到项目组级关系：采购、硬件、影像、质量都有可对话的人，且有备份。"),
    ("产品策略", "以 2M 等可切入规格做第二源破口，再向华勤在研更高规格机型要验证机会，不一次铺太多料号。"),
    ("组织资产", "沉淀《华勤 CIS 导入包》：规格对照、setting、OTP/封装注意事项、FAQ、竞品卡位，可交接。"),
    ("外溢价值", "打法复制到同类手机 ODM（如有规划），销售经理从「管一个客户」变成「管一类战场」。"),
    ("自我能力", "独立完成 70% 售前；复杂 IQ 仍协同 FAE；能主持华勤月度项目会。"),
]
for i, (t, d) in enumerate(year):
    r, c = divmod(i, 3)
    left = Inches(0.55 + c * 4.15)
    top = Inches(1.22 + r * 2.75)
    round_rect(s, left, top, Inches(3.95), Inches(2.5), CARD, LINE)
    add_text(s, left + Inches(0.28), top + Inches(0.25), Inches(3.4), Inches(0.45), [t], 16, SALES, True)
    add_text(s, left + Inches(0.28), top + Inches(0.8), Inches(3.4), Inches(1.45), [d], 13, MUTED, False, leading=5)

# =============================================================================
# 15 Ask and commit
# =============================================================================
s = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(s)
add_title(s, "转正承诺与支持诉求", "承诺可被检查；支持对应 90 天目标，不是要资源的空话")
add_footer(s, 15)

round_rect(s, Inches(0.55), Inches(1.22), Inches(6.1), Inches(5.35), CARD, LINE)
add_text(s, Inches(0.85), Inches(1.45), Inches(5.5), Inches(0.45), ["我承诺"], 18, SALES, True)
add_text(s, Inches(0.85), Inches(2.05), Inches(5.5), Inches(4.2), [
    "·  华勤项目状态周更新，红灯当周升级",
    "·  对外口径与对内事实一致，不藏风险",
    "·  不把送样数当业绩，只认下一节点",
    "·  90 天四件事按月复盘，未完成说明原因和下一步",
    "·  与 FAE 按约定分工，不甩锅、不抢功",
], 15, INK, False, leading=10)

round_rect(s, Inches(6.85), Inches(1.22), Inches(5.95), Inches(5.35), CARD, LINE)
add_text(s, Inches(7.15), Inches(1.45), Inches(5.4), Inches(0.45), ["请公司支持"], 18, ACCENT, True)
add_text(s, Inches(7.15), Inches(2.05), Inches(5.4), Inches(4.2), [
    "·  产品：明确可主推华勤的规格与 roadmap 时间",
    "·  FAE：华勤验证期保证档期，问题 48h 响应",
    "·  商务：第二源价格与供给策略，能进短名单",
    "·  必要时协助品牌侧协同，避免只在 ODM 空转",
    "·  质量/交付：导入风险清单与我同步对外讲",
], 15, INK, False, leading=10)

# =============================================================================
# 16 Close
# =============================================================================
s = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(s, DARK)
gold = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(0.14), prs.slide_height)
gold.fill.solid()
gold.fill.fore_color.rgb = ACCENT
gold.line.fill.background()

add_text(s, Inches(0.9), Inches(2.0), Inches(11.5), Inches(0.9),
         ["请批评指正"], 32, WHITE, True)
add_text(s, Inches(0.9), Inches(2.95), Inches(11.5), Inches(1.2),
         ["试用期把华勤的结构和漏斗立起来了。转正后，用验证结论和定点路径，把这个账号做成公司手机 CIS 的样板。"],
         16, RGBColor(0xC5, 0xD0, 0xD8), False, leading=8)

round_rect(s, Inches(0.9), Inches(4.5), Inches(2.2), Inches(0.08), ACCENT, radius=0.0)
add_text(s, Inches(0.9), Inches(4.85), Inches(11), Inches(0.8),
         ["CIS 销售经理  ·  华勤手机业务", "谢谢"], 14, ACCENT, False)

notes(s, "结束：开放提问。预判问题：1）为什么还没定点？2）和格科怎么打？3）华勤和品牌到底谁说了算？按前面三层客户结构回答。")

out = "/workspace/CIS销售经理-华勤手机业务-转正答辩.pptx"
prs.save(out)
print("saved", out, "slides", len(prs.slides))
