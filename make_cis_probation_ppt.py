# -*- coding: utf-8 -*-
"""CIS 销售经理 · 华勤手机 · 转正答辩
版式对齐「市场营销例会 - 销售与技术服务部周报」：深蓝页眉、进展表、问题与计划。
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

NAVY = RGBColor(0x1B, 0x36, 0x5D)
BLUE = RGBColor(0x2E, 0x5A, 0x88)
PALE = RGBColor(0xE6, 0xEE, 0xF6)
INK = RGBColor(0x1A, 0x2A, 0x3A)
MUTED = RGBColor(0x5A, 0x6A, 0x7A)
GOLD = RGBColor(0xC4, 0xA3, 0x5A)
GREEN = RGBColor(0x2F, 0x6B, 0x4F)
AMBER = RGBColor(0x8A, 0x6A, 0x1A)
RED = RGBColor(0x8B, 0x4A, 0x3C)
CARD = RGBColor(0xFF, 0xFF, 0xFF)
LINE = RGBColor(0xC5, 0xCD, 0xD6)
BG = RGBColor(0xF4, 0xF6, 0xF8)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
ROW = RGBColor(0xF3, 0xF6, 0xF9)
FONT = "Microsoft YaHei"
TOTAL = 16


def font(run, size, bold=False, color=INK):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = FONT
    rPr = run._r.get_or_add_rPr()
    for tag in ("a:ea", "a:cs"):
        el = rPr.find(qn(tag))
        if el is None:
            el = rPr.makeelement(qn(tag), {})
            rPr.insert(0, el)
        el.set("typeface", FONT)


def shape_fill(sh, color, line=False):
    sh.fill.solid()
    sh.fill.fore_color.rgb = color
    if line:
        sh.line.color.rgb = LINE
        sh.line.width = Pt(1)
    else:
        sh.line.fill.background()


def add_bg(slide):
    sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    shape_fill(sh, BG)


def add_header(slide, title, page):
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, Inches(0.82))
    shape_fill(bar, NAVY)
    gold = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, Inches(0.82), prs.slide_width, Inches(0.06))
    shape_fill(gold, GOLD)
    left = slide.shapes.add_textbox(Inches(0.45), Inches(0.12), Inches(9.6), Inches(0.32))
    p = left.text_frame.paragraphs[0]
    r = p.add_run()
    r.text = "销售与技术服务部  |  市场营销例会"
    font(r, 11, False, RGBColor(0xC5, 0xD0, 0xDC))
    t = slide.shapes.add_textbox(Inches(0.45), Inches(0.38), Inches(10.5), Inches(0.4))
    p = t.text_frame.paragraphs[0]
    r = p.add_run()
    r.text = title
    font(r, 20, True, WHITE)
    right = slide.shapes.add_textbox(Inches(10.4), Inches(0.18), Inches(2.5), Inches(0.5))
    p = right.text_frame.paragraphs[0]
    p.alignment = PP_ALIGN.RIGHT
    r = p.add_run()
    r.text = f"转正答辩  {page}/{TOTAL}"
    font(r, 11, False, RGBColor(0xC5, 0xD0, 0xDC))
    foot = slide.shapes.add_textbox(Inches(0.45), Inches(7.18), Inches(10.5), Inches(0.22))
    p = foot.text_frame.paragraphs[0]
    r = p.add_run()
    r.text = "华勤手机 CIS 账号  ·  内部资料"
    font(r, 10, False, MUTED)


def box(slide, l, t, w, h, lines, size=13, color=INK, bold=False, align=PP_ALIGN.LEFT, leading=5):
    tb = slide.shapes.add_textbox(l, t, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.clear()
    for i, text in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.space_after = Pt(leading)
        r = p.add_run()
        r.text = text
        font(r, size, bold, color)
    return tb


def card(slide, l, t, w, h, fill=CARD):
    sh = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, l, t, w, h)
    sh.adjustments[0] = 0.04
    shape_fill(sh, fill, line=True)
    return sh


def conclusion(slide, text):
    sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.45), Inches(1.02), Inches(12.4), Inches(0.38))
    shape_fill(sh, PALE)
    box(slide, Inches(0.6), Inches(1.06), Inches(12.1), Inches(0.32),
        ["结论  " + text], 12, BLUE, False)


def notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text


def fill_table(table, rows, header=NAVY, widths=None):
    if widths:
        for i, w in enumerate(widths):
            table.columns[i].width = w
    for r, row in enumerate(rows):
        for c, val in enumerate(row):
            cell = table.cell(r, c)
            cell.text = val
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.fill.solid()
            if r == 0:
                cell.fill.fore_color.rgb = header
            elif r % 2 == 0:
                cell.fill.fore_color.rgb = ROW
            else:
                cell.fill.fore_color.rgb = WHITE
            for p in cell.text_frame.paragraphs:
                p.space_before = Pt(2)
                p.space_after = Pt(2)
                for run in p.runs:
                    col = WHITE if r == 0 else INK
                    font(run, 12 if r else 12, bold=(r == 0 or c == 0), color=col)


def status_color(val):
    if val in ("已完成", "已建立", "已输出", "已形成", "绿"):
        return GREEN
    if val in ("进行中", "黄"):
        return AMBER
    if val in ("未打开", "未完成", "红"):
        return RED
    return INK


# =============================================================================
# 1 Cover
# =============================================================================
s = prs.slides.add_slide(prs.slide_layouts[6])
bg = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
shape_fill(bg, NAVY)
gold = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(0.16), prs.slide_height)
shape_fill(gold, GOLD)
box(s, Inches(0.8), Inches(1.45), Inches(11.5), Inches(0.35),
    ["销售与技术服务部  ·  市场营销例会专项"], 13, GOLD, True)
box(s, Inches(0.8), Inches(1.95), Inches(12), Inches(1.2),
    ["CIS 销售经理试用期转正答辩"], 32, WHITE, True)
box(s, Inches(0.8), Inches(3.15), Inches(12), Inches(0.45),
    ["负责华勤手机业务  ·  销售公司 CIS 芯片"], 16, RGBColor(0xC5, 0xD0, 0xDC))
line = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(3.8), Inches(2.4), Inches(0.06))
shape_fill(line, GOLD)

meta = [("汇报人", "（姓名）"), ("负责客户", "华勤 · 手机 ODM"), ("汇报日期", "2026 年 8 月"), ("材料性质", "转正答辩")]
for i, (k, v) in enumerate(meta):
    left = Inches(0.8 + i * 3.05)
    box(s, left, Inches(4.2), Inches(2.9), Inches(0.28), [k], 11, GOLD, True)
    box(s, left, Inches(4.5), Inches(2.9), Inches(0.4), [v], 14, WHITE)
box(s, Inches(0.8), Inches(5.4), Inches(11.5), Inches(0.7),
    ["核心结论：试用期把华勤做成可管理的 CIS 账号（决策链、窗口、漏斗、协同闭环）。",
     "转正后用验证结论和定点路径，把过程能力兑成项目结果。"], 13, RGBColor(0xC5, 0xD0, 0xDC), leading=4)
notes(s, "开场 30 秒：岗位是在华勤卖 CIS。华勤是 ODM，品牌管 AVL，模组厂做验证。试用期评价口径是可推进的项目，不是已经量产的金额。")

# =============================================================================
# 2 Agenda
# =============================================================================
s = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(s)
add_header(s, "汇报大纲", 2)
conclusion(s, "按周报习惯讲：先进展，再判断，最后计划和要支持的事项。约 15 分钟。")

agenda = [
    ("01", "试用期工作总结", "客户经营 · 项目导入 · 内部协同 · 专业补齐",
     "对应周报里的「本周期工作 + 项目进展」"),
    ("02", "试用期自我评价", "胜任点 · 短板 · 匹配度",
     "对应周报里的「问题与反思」，不回避未完成项"),
    ("03", "未来规划与展望", "90 天四件事 · 12 个月样板账号 · 支持诉求",
     "对应周报里的「下周/下阶段计划 + 需协调」"),
]
for i, (n, t, a, b) in enumerate(agenda):
    top = Inches(1.58 + i * 1.7)
    card(s, Inches(0.45), top, Inches(12.4), Inches(1.55))
    box(s, Inches(0.7), top + Inches(0.4), Inches(1.2), Inches(0.6), [n], 26, GOLD, True)
    box(s, Inches(2.1), top + Inches(0.22), Inches(10.2), Inches(0.4), [t], 18, INK, True)
    box(s, Inches(2.1), top + Inches(0.65), Inches(10.2), Inches(0.3), [a], 13, BLUE)
    box(s, Inches(2.1), top + Inches(1.0), Inches(10.2), Inches(0.35), [b], 12, MUTED)

# =============================================================================
# 3 Role
# =============================================================================
s = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(s)
add_header(s, "一、岗位与评价口径", 3)
conclusion(s, "在华勤卖 CIS，不是卖一颗料，是在品牌 AVL 框架下把公司芯片推进到可量产。")

items = [
    ("三层客户", "品牌管规格与 AVL；华勤管落地节奏和第二源建议；模组厂管工程验证。销售必须同时看见这三层。"),
    ("一条链路", "情报 → 规格匹配 → 送样 → 模组/IQ 验证 → EVT/DVT → 定点 → 量产。试用期要把前半段跑通，后半段用漏斗管。"),
    ("销售产出", "项目情报、选型卡位、送样与问题闭环、价格与供给口径、客情与内部协同。参数由 FAE 支撑，销售对节点负责。"),
    ("试用期怎么评", "看窗口、漏斗、闭环、风险是否暴露，不把尚未发生的定点金额当试用期 KPI。"),
]
for i, (t, d) in enumerate(items):
    r, c = divmod(i, 2)
    left = Inches(0.45 + c * 6.4)
    top = Inches(1.58 + r * 2.55)
    card(s, left, top, Inches(6.15), Inches(2.4))
    mark = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, Inches(0.1), Inches(2.4))
    shape_fill(mark, GOLD if i == 0 else BLUE)
    box(s, left + Inches(0.35), top + Inches(0.25), Inches(5.5), Inches(0.4), [t], 16, INK, True)
    box(s, left + Inches(0.35), top + Inches(0.8), Inches(5.5), Inches(1.35), [d], 13, MUTED, leading=6)
notes(s, "点题：格科/思特威/OV 往往已在品牌短名单。切入点是第二源、供给对冲、新机型窗口，不硬刚纸面参数。")

# =============================================================================
# 4 Overview
# =============================================================================
s = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(s)
add_header(s, "一、试用期工作总览", 4)
conclusion(s, "四条线并行：摸清客户、卡位项目、建立机制、补齐专业。")

cards = [
    ("01  客户经营", ["完成华勤手机 CIS 决策链梳理", "采购 / 项目 / 硬件三线窗口", "明确 ODM 与品牌 AVL 约束", "拜访纪要与沟通节奏可维持"]),
    ("02  项目导入", ["在研机型摄像头规格地图", "2M 等规格对标成熟竞品", "送样 + FAE 验证闭环", "A/B/C 漏斗进入周报"]),
    ("03  内部协同", ["客户语言翻译成内部任务", "与 FAE / 计划 / 质量对齐", "红黄绿状态，不报喜藏忧", "导入包：规格、setting、FAQ"]),
    ("04  专业内功", ["手机向 CIS 规格与竞品", "送样、OTP、模组验证节奏", "华勤 EVT / DVT / MP 节点", "能独立完成常规拜访与纪要"]),
]
for i, (title, lines) in enumerate(cards):
    r, c = divmod(i, 2)
    left = Inches(0.45 + c * 6.4)
    top = Inches(1.58 + r * 2.55)
    card(s, left, top, Inches(6.15), Inches(2.4), PALE if i % 2 == 0 else CARD)
    box(s, left + Inches(0.3), top + Inches(0.2), Inches(5.5), Inches(0.4), [title], 16, BLUE, True)
    box(s, left + Inches(0.3), top + Inches(0.7), Inches(5.5), Inches(1.5),
        ["·  " + x for x in lines], 13, INK, leading=4)

# =============================================================================
# 5 Timeline
# =============================================================================
s = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(s)
add_header(s, "一、试用期节奏（按周报阶段汇总）", 5)
conclusion(s, "先结构、后项目、再机制。避免一上来只追一颗料。")

line = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1.2), Inches(1.7), Inches(10.9), Inches(0.05))
shape_fill(line, LINE)
phases = [
    ("第 1 阶段", "融入与摸底", ["产品/竞品/导入流程学习", "华勤组织与决策链访谈", "识别在研机型与现有供应商", "确定主对接与备份窗口"]),
    ("第 2 阶段", "卡位与送样", ["规格对照（我司 vs 成熟料）", "目标机型送样与模组协同", "FAE 陪访，问题 48h 闭环", "报价与交期口径内部对齐"]),
    ("第 3 阶段", "漏斗与复盘", ["项目 A/B/C 分级周更新", "验证进展与风险升级", "客情：谁决策、谁影响", "华勤账号作战手册初稿"]),
]
for i, (ph, title, items) in enumerate(phases):
    left = Inches(0.45 + i * 4.25)
    dot = s.shapes.add_shape(MSO_SHAPE.OVAL, left + Inches(1.75), Inches(1.58), Inches(0.28), Inches(0.28))
    shape_fill(dot, GOLD if i == 2 else BLUE)
    card(s, left, Inches(2.05), Inches(4.05), Inches(4.7))
    box(s, left + Inches(0.28), Inches(2.22), Inches(3.5), Inches(0.3), [ph], 12, GOLD, True)
    box(s, left + Inches(0.28), Inches(2.55), Inches(3.5), Inches(0.4), [title], 18, INK, True)
    box(s, left + Inches(0.28), Inches(3.15), Inches(3.5), Inches(3.2),
        ["·  " + x for x in items], 14, MUTED, leading=8)

# =============================================================================
# 6 Decision chain
# =============================================================================
s = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(s)
add_header(s, "一、客户经营：华勤决策链", 6)
conclusion(s, "ODM 不是单一采购。每个角色要的东西不同，销售动作必须拆开。")

rows = [
    ["角色", "他关心什么", "试用期动作（周报口径）"],
    ["项目经理", "节点、风险、能否赶上 EVT/DVT", "对齐机型节奏，未验证料不推进关键路径"],
    ["摄像头硬件", "规格、封装/脚位、替代可行性", "输出对照表，明确第二源与改板成本"],
    ["影像 / 调试", "IQ、噪声、暗角、验证结论", "FAE 进现场，问题闭环写入纪要"],
    ["采购", "价格、交期、供给安全、商务条款", "建立窗口，同步公司供给与报价原则"],
    ["质量 / 计划", "可靠性、客诉、上量节奏", "提前暴露封装 / 回流 / OTP 导入风险"],
]
tbl = s.shapes.add_table(len(rows), 3, Inches(0.45), Inches(1.55), Inches(12.4), Inches(5.15)).table
fill_table(tbl, rows, widths=[Inches(2.1), Inches(4.7), Inches(5.6)])
notes(s, "华勤能推荐，品牌 AVL 不过仍定不了。试用期先盘清 ODM 内部是正确顺序，品牌窗口放 90 天。")

# =============================================================================
# 7 Project table (weekly style)
# =============================================================================
s = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(s)
add_header(s, "一、项目进展（周报表）", 7)
conclusion(s, "漏斗已建立。试用期停在送样/验证，不把定点写成既成事实。机型名现场按保密口径补充。")

rows = [
    ["项目", "规格卡位", "阶段", "状态", "本周期进展 / 风险", "下一步"],
    ["前置 2M 第二源", "对标成熟 1/5\" 2M", "送样 / 验证", "黄", "资料包与陪访机制已跑通，验证结论跟进", "出书面 IQ/模组结论"],
    ["功能摄 / 微距档", "2M 规格，看开案窗口", "匹配 / 送样", "黄", "规格可进短名单，机型节奏待项目组确认", "锁定开案节点"],
    ["新开案选型池", "在研机型摄像头地图", "情报 / 匹配", "绿", "竞品占位（格科/思特威/OV）已摸清", "滚动更新漏斗"],
    ["品牌 AVL 路径", "需品牌短名单", "未启动", "红", "约束已识别，窗口未打开", "要品牌侧或指定模组线索"],
]
tbl = s.shapes.add_table(len(rows), 6, Inches(0.35), Inches(1.55), Inches(12.6), Inches(4.35)).table
fill_table(tbl, rows, widths=[Inches(2.15), Inches(2.15), Inches(1.7), Inches(0.85), Inches(3.5), Inches(2.25)])
for r in range(1, 5):
    cell = tbl.cell(r, 3)
    for p in cell.text_frame.paragraphs:
        p.alignment = PP_ALIGN.CENTER
        for run in p.runs:
            font(run, 12, True, status_color(cell.text_frame.text))

box(s, Inches(0.45), Inches(6.05), Inches(12.4), Inches(0.9), [
    "主攻话术：供给弹性、导入服务、风险对冲，做「可落地的第二源」，不硬刚像素纸面参数。",
    "红灯当周升级；黄灯写清阻塞点；绿灯也要写下一节点，避免只有拜访没有下文。",
], 12, MUTED, leading=3)

# =============================================================================
# 8 Results
# =============================================================================
s = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(s)
add_header(s, "一、试用期成果盘点", 8)
conclusion(s, "用可核验的过程结果说话。定点与收入作为下一阶段指标。")

rows = [
    ["工作项", "状态", "结果说明"],
    ["华勤手机 CIS 组织与决策链地图", "已完成", "覆盖项目 / 硬件 / 影像 / 采购 / 质量，避免单点窗口"],
    ["主对接 + 备份窗口", "已建立", "采购与项目双线，沟通节奏可按周/双周维持"],
    ["在研机型摄像头规格与竞品卡位表", "已输出", "看清格科/思特威/OV 在哪些像素档占位、窗口在哪"],
    ["目标规格送样与 FAE 协同", "进行中", "资料包、陪访、48h 闭环已跑通，验证结论跟进中"],
    ["项目漏斗（A/B/C）纳入周报", "已建立", "红黄绿状态对内对齐，不报喜藏忧"],
    ["华勤账号作战手册初稿", "已形成", "决策链、话术、导入节奏、风险清单可交接"],
    ["品牌方 AVL 窗口", "未打开", "已识别约束，列入转正后 90 天，不回避"],
]
tbl = s.shapes.add_table(len(rows), 3, Inches(0.45), Inches(1.52), Inches(12.4), Inches(5.2)).table
fill_table(tbl, rows, widths=[Inches(4.3), Inches(1.5), Inches(6.6)])
for r in range(1, len(rows)):
    cell = tbl.cell(r, 1)
    for p in cell.text_frame.paragraphs:
        p.alignment = PP_ALIGN.CENTER
        for run in p.runs:
            font(run, 12, True, status_color(cell.text_frame.text))

# =============================================================================
# 9 Collaboration
# =============================================================================
s = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(s)
add_header(s, "一、销售 × 技术服务协同", 9)
conclusion(s, "销售对节点负责，FAE 对技术结论负责。避免两张皮。")
card(s, Inches(0.45), Inches(1.55), Inches(6.15), Inches(5.15))
box(s, Inches(0.75), Inches(1.75), Inches(5.5), Inches(0.4), ["工作约定（可进周报检查）"], 16, BLUE, True)
box(s, Inches(0.75), Inches(2.3), Inches(5.5), Inches(4.0), [
    "·  拜访前：规格问题清单，不空跑",
    "·  现场：销售控议程，FAE 答技术",
    "·  当天：纪要含问题、责任人、截止日",
    "·  48 小时：技术问题给到书面回复",
    "·  卡住节点：当周升级内部会",
], 15, INK, leading=10)
card(s, Inches(6.75), Inches(1.55), Inches(6.1), Inches(5.15), PALE)
box(s, Inches(7.05), Inches(1.75), Inches(5.5), Inches(0.4), ["对内对外同一套事实"], 16, BLUE, True)
box(s, Inches(7.05), Inches(2.3), Inches(5.5), Inches(4.0), [
    "·  客户要的是「能否赶上这台机」",
    "·  内部要的是「问题能否关闭」",
    "·  导入风险提前讲：封装、CRA、OTP、上电时序",
    "·  竞品动态进周报，供产品与计划决策",
    "·  不把送样次数写成业绩",
], 15, INK, leading=10)

# =============================================================================
# 10 Gaps
# =============================================================================
s = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(s)
add_header(s, "一、未完成项与风险（周报红灯）", 10)
conclusion(s, "转正答辩把短板放桌上，比把漏斗画满更重要。")
gaps = [
    ("品牌 AVL 尚未突破", "华勤可推荐，品牌短名单不打开则到不了定点。试用期优先盘清 ODM，品牌窗口列入 90 天。"),
    ("尚无量产定点可报", "手机 CIS 导入周期长于试用期。应用漏斗质量、验证进度、客情厚度评价，而不是尚未发生的收入。"),
    ("商务深度不够", "价格、账期、供给承诺还需与产品、计划形成华勤专用口径，避免现场口头承诺。"),
    ("售前仍依赖 FAE", "常规规格已能讲，OTP/IQ 细节仍需 FAE。转正后独立覆盖约 70% 售前问答。"),
]
for i, (t, d) in enumerate(gaps):
    r, c = divmod(i, 2)
    left = Inches(0.45 + c * 6.4)
    top = Inches(1.55 + r * 2.55)
    card(s, left, top, Inches(6.15), Inches(2.4))
    mark = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, Inches(0.1), Inches(2.4))
    shape_fill(mark, RED)
    box(s, left + Inches(0.35), top + Inches(0.25), Inches(5.5), Inches(0.4), [t], 16, INK, True)
    box(s, left + Inches(0.35), top + Inches(0.8), Inches(5.5), Inches(1.35), [d], 13, MUTED, leading=6)

# =============================================================================
# 11 Self eval
# =============================================================================
s = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(s)
add_header(s, "二、试用期自我评价", 11)
conclusion(s, "六个维度给评委一个打分框架。强项可复用，短板有计划。")
dims = [
    ("客户洞察", "强", "能把华勤三层结构讲清楚，知道决策链和品牌约束"),
    ("目标管理", "中+", "漏斗和周报已建立；A 类项目节点还要更细"),
    ("沟通协同", "强", "对内对外同一套事实，与 FAE 约定可执行"),
    ("专业能力", "中", "规格与导入流程已入门，IQ/OTP 仍需加强"),
    ("抗压担当", "强", "长周期、多角色，不把问题藏到月底"),
    ("结果导向", "中+", "过程结果扎实；经营结果靠下一阶段验证"),
]
for i, (name, level, desc) in enumerate(dims):
    r, c = divmod(i, 3)
    left = Inches(0.45 + c * 4.25)
    top = Inches(1.55 + r * 2.55)
    card(s, left, top, Inches(4.05), Inches(2.4))
    box(s, left + Inches(0.25), top + Inches(0.25), Inches(2.2), Inches(0.4), [name], 16, INK, True)
    pill = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left + Inches(2.55), top + Inches(0.28), Inches(1.2), Inches(0.36))
    pill.adjustments[0] = 0.4
    shape_fill(pill, PALE)
    box(s, left + Inches(2.55), top + Inches(0.3), Inches(1.2), Inches(0.32), [level], 12, BLUE, True, align=PP_ALIGN.CENTER)
    box(s, left + Inches(0.25), top + Inches(0.85), Inches(3.55), Inches(1.3), [desc], 13, MUTED, leading=5)

# =============================================================================
# 12 Conclusion of eval
# =============================================================================
s = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(s)
add_header(s, "二、综合评价与转正判断", 12)
conclusion(s, "匹配这个岗位。转正后要用项目结果把过程能力兑现。")
card(s, Inches(0.45), Inches(1.55), Inches(12.4), Inches(1.35), PALE)
box(s, Inches(0.7), Inches(1.75), Inches(11.9), Inches(1.05),
    ["胜任判断：能在华勤这种多层决策的 ODM 账号上，把 CIS 销售做成可管理的项目，而不是个人关系或一次性送样。这是转正的基本盘。"],
    15, INK, leading=6)
card(s, Inches(0.45), Inches(3.1), Inches(6.15), Inches(3.6))
box(s, Inches(0.75), Inches(3.3), Inches(5.5), Inches(0.4), ["适合留下的理由"], 16, GREEN, True)
box(s, Inches(0.75), Inches(3.85), Inches(5.5), Inches(2.6), [
    "·  客户结构理解快，不把 ODM 当渠道批发",
    "·  执行闭环清楚：纪要、责任人、升级",
    "·  对数字诚实，不把漏斗当业绩",
    "·  愿意补专业，不与 FAE 抢结论",
], 14, INK, leading=8)
card(s, Inches(6.75), Inches(3.1), Inches(6.1), Inches(3.6))
box(s, Inches(7.05), Inches(3.3), Inches(5.5), Inches(0.4), ["必须补齐的短板"], 16, RED, True)
box(s, Inches(7.05), Inches(3.85), Inches(5.5), Inches(2.6), [
    "·  品牌侧渗透几乎为零",
    "·  独立商务谈判经验不足",
    "·  技术问答还不能独立覆盖",
    "·  A 类项目的周节点仍偏粗",
], 14, INK, leading=8)
notes(s, "请评委按短板提问。转正不是完美，是方向对、闭环在、短板有计划。")

# =============================================================================
# 13 90 days
# =============================================================================
s = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(s)
add_header(s, "三、未来 90 天（下阶段周计划放大）", 13)
conclusion(s, "四件必须发生的事。用节点考核，不用愿望考核。")
q = [
    ("01", "验证出结论", "至少 1 个目标项目完成模组/IQ 验证，形成书面结论（过 / 有条件过 / 不过及原因）。"),
    ("02", "品牌线索", "画出品牌—华勤—模组三方图，拿到至少 1 个品牌侧或品牌指定模组的有效窗口线索。"),
    ("03", "商务口径", "与产品、计划锁定华勤专用报价、交期、供给原则，现场不再口头承诺。"),
    ("04", "机制固化", "华勤项目月评：漏斗、风险、FAE 资源、下一步，形成公司可看见的账号经营。"),
]
for i, (n, t, d) in enumerate(q):
    top = Inches(1.55 + i * 1.28)
    card(s, Inches(0.45), top, Inches(12.4), Inches(1.18))
    box(s, Inches(0.7), top + Inches(0.32), Inches(1.0), Inches(0.5), [n], 20, GOLD, True)
    box(s, Inches(1.85), top + Inches(0.16), Inches(10.6), Inches(0.35), [t], 16, INK, True)
    box(s, Inches(1.85), top + Inches(0.58), Inches(10.6), Inches(0.45), [d], 13, MUTED)

# =============================================================================
# 14 12 months
# =============================================================================
s = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(s)
add_header(s, "三、未来 12 个月：华勤做成样板账号", 14)
conclusion(s, "样板 = 可复制方法，而不只是「我跟过一个客户」。")
year = [
    ("项目结果", "1 个项目进入试产或定点路径；2–3 个项目稳定在验证池；杜绝只有拜访没有下一节点的空账号。"),
    ("关系结构", "升级到项目组级关系：采购、硬件、影像、质量都有可对话的人，且有备份。"),
    ("产品策略", "以 2M 等可切入规格做第二源破口，再向更高规格在研机型要验证机会，不一次铺太多料号。"),
    ("组织资产", "沉淀《华勤 CIS 导入包》：规格对照、setting、OTP/封装注意、FAQ、竞品卡位，可交接。"),
    ("外溢价值", "打法复制到同类手机 ODM（如有规划），从管一个客户变成管一类战场。"),
    ("自我能力", "独立完成约 70% 售前；复杂 IQ 仍协同 FAE；能主持华勤月度项目会。"),
]
for i, (t, d) in enumerate(year):
    r, c = divmod(i, 3)
    left = Inches(0.45 + c * 4.25)
    top = Inches(1.55 + r * 2.55)
    card(s, left, top, Inches(4.05), Inches(2.4))
    box(s, left + Inches(0.25), top + Inches(0.22), Inches(3.55), Inches(0.4), [t], 16, BLUE, True)
    box(s, left + Inches(0.25), top + Inches(0.75), Inches(3.55), Inches(1.45), [d], 13, MUTED, leading=5)

# =============================================================================
# 15 Commit & ask
# =============================================================================
s = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(s)
add_header(s, "三、转正承诺与需协调事项", 15)
conclusion(s, "承诺可被周报检查；支持对应 90 天目标。")
card(s, Inches(0.45), Inches(1.55), Inches(6.15), Inches(5.15))
box(s, Inches(0.75), Inches(1.75), Inches(5.5), Inches(0.4), ["我承诺"], 18, BLUE, True)
box(s, Inches(0.75), Inches(2.3), Inches(5.5), Inches(4.1), [
    "·  华勤项目状态进周报，红灯当周升级",
    "·  对外口径与对内事实一致，不藏风险",
    "·  不把送样数当业绩，只认下一节点",
    "·  90 天四件事按月复盘，未完成写原因和下一步",
    "·  与 FAE 按约定分工，不甩锅、不抢功",
], 15, INK, leading=10)
card(s, Inches(6.75), Inches(1.55), Inches(6.1), Inches(5.15), PALE)
box(s, Inches(7.05), Inches(1.75), Inches(5.5), Inches(0.4), ["请公司支持"], 18, GOLD, True)
box(s, Inches(7.05), Inches(2.3), Inches(5.5), Inches(4.1), [
    "·  产品：明确可主推华勤的规格与 roadmap",
    "·  FAE：验证期保证档期，问题 48h 响应",
    "·  商务：第二源价格与供给策略，能进短名单",
    "·  必要时协助品牌侧协同，避免只在 ODM 空转",
    "·  质量/交付：导入风险清单与我同步对外讲",
], 15, INK, leading=10)

# =============================================================================
# 16 Close
# =============================================================================
s = prs.slides.add_slide(prs.slide_layouts[6])
bg = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
shape_fill(bg, NAVY)
gold = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(0.16), prs.slide_height)
shape_fill(gold, GOLD)
box(s, Inches(0.8), Inches(2.05), Inches(11.5), Inches(0.8), ["请批评指正"], 32, WHITE, True)
box(s, Inches(0.8), Inches(2.95), Inches(11.5), Inches(1.15),
    ["试用期把华勤的结构和漏斗立起来了。转正后，用验证结论和定点路径，",
     "把这个账号做成公司手机 CIS 的样板，并按周报机制接受检查。"], 16, RGBColor(0xC5, 0xD0, 0xDC), leading=6)
line = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(4.45), Inches(2.2), Inches(0.06))
shape_fill(line, GOLD)
box(s, Inches(0.8), Inches(4.75), Inches(11), Inches(0.7),
    ["销售与技术服务部  ·  CIS 销售经理  ·  华勤手机业务", "谢谢"], 14, GOLD, leading=4)
notes(s, "预判提问：1）为什么还没定点？导入周期长于试用期，用漏斗和验证进度评价。2）和格科怎么打？第二源+供给+服务，不硬刚参数。3）华勤和品牌谁说了算？品牌 AVL，华勤落地，两边都要做。")

out = "/workspace/CIS销售经理-华勤手机业务-转正答辩.pptx"
prs.save(out)
print("saved", out, "slides", len(prs.slides))
