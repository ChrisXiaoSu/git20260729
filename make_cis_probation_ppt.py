# -*- coding: utf-8 -*-
"""转正答辩 PPT：沿用《市场营销例会-销售与技术服务部周报260824》模块
本期结论 → 销售工作 → 项目进度 → 技术服务 → 问题风险 → 自我评价 → 下阶段计划 → 需协调
"""
from pptx import Presentation
from pptx.util import Inches, Pt
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
TOTAL = 14


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


def fill(sh, color, lined=False):
    sh.fill.solid()
    sh.fill.fore_color.rgb = color
    if lined:
        sh.line.color.rgb = LINE
        sh.line.width = Pt(1)
    else:
        sh.line.fill.background()


def add_bg(s):
    sh = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    fill(sh, BG)


def header(s, title, page):
    bar = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, Inches(0.82))
    fill(bar, NAVY)
    g = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, Inches(0.82), prs.slide_width, Inches(0.06))
    fill(g, GOLD)
    tb = s.shapes.add_textbox(Inches(0.45), Inches(0.10), Inches(10.2), Inches(0.28))
    r = tb.text_frame.paragraphs[0].add_run()
    r.text = "销售与技术服务部  |  市场营销例会  |  周报框架 260824"
    font(r, 11, False, RGBColor(0xC5, 0xD0, 0xDC))
    tb = s.shapes.add_textbox(Inches(0.45), Inches(0.38), Inches(10.4), Inches(0.4))
    r = tb.text_frame.paragraphs[0].add_run()
    r.text = title
    font(r, 20, True, WHITE)
    tb = s.shapes.add_textbox(Inches(10.5), Inches(0.22), Inches(2.4), Inches(0.42))
    p = tb.text_frame.paragraphs[0]
    p.alignment = PP_ALIGN.RIGHT
    r = p.add_run()
    r.text = f"{page} / {TOTAL}"
    font(r, 11, False, RGBColor(0xC5, 0xD0, 0xDC))
    tb = s.shapes.add_textbox(Inches(0.45), Inches(7.18), Inches(12.4), Inches(0.22))
    r = tb.text_frame.paragraphs[0].add_run()
    r.text = "转正答辩专题  ·  华勤手机 CIS  ·  内部资料"
    font(r, 10, False, MUTED)


def box(s, l, t, w, h, lines, size=13, color=INK, bold=False, align=PP_ALIGN.LEFT, leading=5):
    tb = s.shapes.add_textbox(l, t, w, h)
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


def card(s, l, t, w, h, color=CARD):
    sh = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, l, t, w, h)
    sh.adjustments[0] = 0.04
    fill(sh, color, True)
    return sh


def banner(s, text):
    sh = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.45), Inches(1.02), Inches(12.4), Inches(0.38))
    fill(sh, PALE)
    box(s, Inches(0.6), Inches(1.06), Inches(12.1), Inches(0.32), ["本期结论  " + text], 12, BLUE)


def notes(s, text):
    s.notes_slide.notes_text_frame.text = text


def status_color(val):
    if val in ("已完成", "已建立", "绿", "强"):
        return GREEN
    if val in ("进行中", "黄", "中", "中+"):
        return AMBER
    if val in ("未打开", "未完成", "红"):
        return RED
    return INK


def table(s, rows, l, t, w, h, widths, center_cols=()):
    tbl = s.shapes.add_table(len(rows), len(rows[0]), l, t, w, h).table
    for i, ww in enumerate(widths):
        tbl.columns[i].width = ww
    for r, row in enumerate(rows):
        for c, val in enumerate(row):
            cell = tbl.cell(r, c)
            cell.text = val
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.fill.solid()
            cell.fill.fore_color.rgb = NAVY if r == 0 else (ROW if r % 2 == 0 else WHITE)
            for p in cell.text_frame.paragraphs:
                p.space_before = Pt(2)
                p.space_after = Pt(2)
                if c in center_cols:
                    p.alignment = PP_ALIGN.CENTER
                for run in p.runs:
                    col = WHITE if r == 0 else (status_color(val) if c in center_cols and r else INK)
                    font(run, 12, bold=(r == 0 or c == 0 or (c in center_cols and r > 0)), color=col)
    return tbl


# 1 Cover — 周报封面信息架构
s = prs.slides.add_slide(prs.slide_layouts[6])
sh = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
fill(sh, NAVY)
g = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(0.16), prs.slide_height)
fill(g, GOLD)
box(s, Inches(0.8), Inches(1.35), Inches(12), Inches(0.32),
    ["市场营销例会  ·  销售与技术服务部周报框架（260824）"], 13, GOLD, True)
box(s, Inches(0.8), Inches(1.85), Inches(12), Inches(1.1),
    ["CIS 销售经理试用期转正答辩"], 32, WHITE, True)
box(s, Inches(0.8), Inches(3.05), Inches(12), Inches(0.4),
    ["负责华勤手机业务  ·  销售公司 CIS 芯片"], 16, RGBColor(0xC5, 0xD0, 0xDC))
ln = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(3.65), Inches(2.3), Inches(0.06))
fill(ln, GOLD)
meta = [("汇报人", "________"), ("部门", "销售与技术服务部"), ("客户", "华勤 · 手机 ODM"), ("日期", "2026 年 8 月 24 日")]
for i, (k, v) in enumerate(meta):
    left = Inches(0.8 + i * 3.05)
    box(s, left, Inches(4.05), Inches(2.9), Inches(0.26), [k], 11, GOLD, True)
    box(s, left, Inches(4.35), Inches(2.9), Inches(0.4), [v], 14, WHITE)
box(s, Inches(0.8), Inches(5.15), Inches(11.6), Inches(1.0), [
    "本期结论：试用期完成华勤账号的决策链、窗口、漏斗和销售×技术服务闭环。",
    "转正后按周报机制接受检查，用验证结论和定点路径兑现经营结果。",
], 14, RGBColor(0xC5, 0xD0, 0xDC), leading=4)
notes(s, "开场：沿用部门周报结构汇报。岗位是在华勤卖 CIS。试用期看过程指标，不看尚未发生的定点金额。")

# 2 目录 = 周报模块
s = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(s)
header(s, "汇报目录（沿用周报模块）", 2)
banner(s, "先工作进展，再自我评价，最后下阶段计划与需协调。约 15 分钟。")
items = [
    ("01", "试用期工作总结", "销售工作 · 项目进度跟踪 · 技术服务 · 问题与风险"),
    ("02", "试用期自我评价", "六维能力 · 胜任判断 · 短板不回避"),
    ("03", "未来规划与展望", "90 天节点 · 12 个月样板账号 · 需协调事项"),
]
for i, (n, t, d) in enumerate(items):
    top = Inches(1.58 + i * 1.7)
    card(s, Inches(0.45), top, Inches(12.4), Inches(1.55))
    box(s, Inches(0.7), top + Inches(0.4), Inches(1.2), Inches(0.55), [n], 26, GOLD, True)
    box(s, Inches(2.1), top + Inches(0.28), Inches(10), Inches(0.4), [t], 20, INK, True)
    box(s, Inches(2.1), top + Inches(0.8), Inches(10), Inches(0.4), [d], 14, MUTED)

# 3 本期结论
s = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(s)
header(s, "本期核心结论", 3)
banner(s, "华勤是 ODM：品牌管 AVL，华勤管落地，模组厂做验证。销售要同时看见这三层。")
rows = [
    ["周报栏目", "试用期结果", "状态"],
    ["销售工作", "决策链、采购/项目/硬件窗口、竞品卡位表", "已完成"],
    ["项目进度跟踪", "2M 第二源进入送样/验证；漏斗 A/B/C 进周报", "进行中"],
    ["技术服务", "陪访、48h 闭环、导入风险前置（封装/OTP/CRA）", "已建立"],
    ["问题与风险", "品牌 AVL 未打开；尚无定点；商务口径未锁", "红灯在册"],
    ["评价口径", "窗口 / 漏斗 / 闭环，不把未发生的收入当试用期 KPI", "对齐中"],
]
table(s, rows, Inches(0.45), Inches(1.55), Inches(12.4), Inches(5.15),
      [Inches(2.6), Inches(8.0), Inches(1.8)], center_cols=(2,))
notes(s, "强调：格科/思特威/OV 多已在品牌短名单。切入点是第二源、供给对冲、新机型窗口。")

# 4 一、销售工作
s = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(s)
header(s, "一、试用期工作总结  |  销售工作", 4)
banner(s, "先结构、后项目。把华勤从客户名称做成可经营账号。")
rows = [
    ["角色", "关心什么", "试用期动作"],
    ["项目经理", "EVT/DVT 节点与风险", "对齐机型节奏，未验证料不进关键路径"],
    ["摄像头硬件", "规格、封装、第二源可行性", "输出对照表，明确改板成本"],
    ["影像 / 调试", "IQ 与验证结论", "拉 FAE 进现场，结论进纪要"],
    ["采购", "价格、交期、供给", "建立窗口，同步公司报价与供给原则"],
    ["质量 / 计划", "可靠性、上量", "提前暴露封装/回流/OTP 风险"],
]
table(s, rows, Inches(0.45), Inches(1.55), Inches(12.4), Inches(5.15),
      [Inches(2.2), Inches(4.4), Inches(5.8)])

# 5 一、项目进度
s = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(s)
header(s, "一、试用期工作总结  |  项目进度跟踪", 5)
banner(s, "红黄绿按周报更新。机型名现场按保密口径替换。")
rows = [
    ["项目", "规格卡位", "阶段", "灯", "本周期进展 / 风险", "下一步"],
    ["前置 2M 第二源", "对标成熟 1/5\" 2M", "送样/验证", "黄", "资料包与陪访已跑通，结论跟进中", "出书面验证结论"],
    ["功能摄/微距档", "2M，看开案窗口", "匹配/送样", "黄", "规格可进短名单，节奏待项目组确认", "锁定开案节点"],
    ["新开案选型池", "在研机型摄像头地图", "情报/匹配", "绿", "格科/思特威/OV 占位已摸清", "滚动更新漏斗"],
    ["品牌 AVL 路径", "需进入品牌短名单", "未启动", "红", "约束已识别，窗口未打开", "要品牌或指定模组线索"],
]
table(s, rows, Inches(0.35), Inches(1.52), Inches(12.6), Inches(4.35),
      [Inches(2.15), Inches(2.2), Inches(1.55), Inches(0.7), Inches(3.75), Inches(2.25)],
      center_cols=(3,))
box(s, Inches(0.45), Inches(6.05), Inches(12.4), Inches(0.85), [
    "主攻：供给弹性 + 导入服务 + 风险对冲，做可落地的第二源，不硬刚纸面参数。",
    "绿灯也要写下一节点。黄灯写阻塞点。红灯当周升级。不把送样次数写成业绩。",
], 12, MUTED, leading=3)

# 6 一、技术服务
s = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(s)
header(s, "一、试用期工作总结  |  技术服务", 6)
banner(s, "销售对节点负责，FAE 对技术结论负责。部门周报里销售与技术服务本来就是一块。")
card(s, Inches(0.45), Inches(1.55), Inches(6.15), Inches(5.15))
box(s, Inches(0.75), Inches(1.75), Inches(5.5), Inches(0.4), ["工作约定（可进周报检查）"], 16, BLUE, True)
box(s, Inches(0.75), Inches(2.35), Inches(5.5), Inches(4.0), [
    "·  拜访前出具规格问题清单，不空跑",
    "·  现场销售控议程，FAE 答技术",
    "·  当天纪要：问题 / 责任人 / 截止日",
    "·  技术问题 48 小时书面回复",
    "·  卡住节点当周升级内部会",
], 15, INK, leading=10)
card(s, Inches(6.75), Inches(1.55), Inches(6.1), Inches(5.15), PALE)
box(s, Inches(7.05), Inches(1.75), Inches(5.5), Inches(0.4), ["对内对外同一套事实"], 16, BLUE, True)
box(s, Inches(7.05), Inches(2.35), Inches(5.5), Inches(4.0), [
    "·  客户要的是：这台机赶不赶得上",
    "·  内部要的是：问题关不关得掉",
    "·  风险前置：封装、CRA、OTP、上电时序",
    "·  竞品动态进周报，给产品/计划决策",
    "·  导入包：规格对照、setting、FAQ",
], 15, INK, leading=10)

# 7 一、问题与风险 + 成果
s = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(s)
header(s, "一、试用期工作总结  |  问题与风险", 7)
banner(s, "红灯进周报。转正答辩不藏未完成项。")
rows = [
    ["工作项 / 风险", "状态", "说明"],
    ["决策链地图（项目/硬件/影像/采购/质量）", "已完成", "避免单点窗口"],
    ["主对接 + 备份窗口", "已建立", "采购与项目双线，周/双周节奏"],
    ["规格与竞品卡位表", "已输出", "看清成熟料占位和开案窗口"],
    ["送样与 FAE 协同", "进行中", "机制已跑通，验证结论跟进"],
    ["项目漏斗纳入周报", "已建立", "红黄绿对内对齐"],
    ["品牌 AVL 窗口", "未打开", "90 天重点，不回避"],
    ["量产定点", "未完成", "导入周期长于试用期，不提前报喜"],
    ["华勤专用商务口径", "未完成", "价格/交期/供给需与产品计划锁定"],
]
table(s, rows, Inches(0.45), Inches(1.5), Inches(12.4), Inches(5.2),
      [Inches(5.0), Inches(1.6), Inches(5.8)], center_cols=(1,))

# 8 二、自我评价
s = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(s)
header(s, "二、试用期自我评价", 8)
banner(s, "给评委打分框架。强项可复用，短板写进下阶段计划。")
dims = [
    ("客户洞察", "强", "能讲清品牌—华勤—模组三层，知道谁决策、哪里会卡住"),
    ("目标管理", "中+", "漏斗和周报已建立；A 类项目周节点还要更细"),
    ("沟通协同", "强", "对内对外同一套事实，与 FAE 约定可执行"),
    ("专业能力", "中", "规格与导入流程已入门，IQ/OTP 仍需 FAE"),
    ("抗压担当", "强", "长周期、多角色，红灯当周暴露，不藏到月底"),
    ("结果导向", "中+", "过程结果扎实；定点与收入靠下一阶段兑现"),
]
for i, (n, lv, d) in enumerate(dims):
    r, c = divmod(i, 3)
    left = Inches(0.45 + c * 4.25)
    top = Inches(1.55 + r * 2.55)
    card(s, left, top, Inches(4.05), Inches(2.4))
    box(s, left + Inches(0.22), top + Inches(0.22), Inches(2.3), Inches(0.4), [n], 16, INK, True)
    pill = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left + Inches(2.55), top + Inches(0.26), Inches(1.2), Inches(0.36))
    pill.adjustments[0] = 0.4
    fill(pill, PALE)
    box(s, left + Inches(2.55), top + Inches(0.28), Inches(1.2), Inches(0.32), [lv], 12, BLUE, True, align=PP_ALIGN.CENTER)
    box(s, left + Inches(0.22), top + Inches(0.8), Inches(3.6), Inches(1.35), [d], 13, MUTED, leading=5)

# 9 二、综合评价
s = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(s)
header(s, "二、综合评价与转正判断", 9)
banner(s, "匹配岗位。转正后用项目结果兑现过程能力。")
card(s, Inches(0.45), Inches(1.55), Inches(12.4), Inches(1.4), PALE)
box(s, Inches(0.7), Inches(1.72), Inches(11.9), Inches(1.1),
    ["胜任：能在华勤这种多层决策的 ODM 上，把 CIS 销售做成可管理的项目，而不是个人关系或一次性送样。",
     "对「为何还没定点」：导入跨 EVT/DVT，周期长于试用期；竞品已在品牌短名单。试用期 KPI 应看窗口、漏斗、验证和风险暴露。"],
    14, INK, leading=6)
card(s, Inches(0.45), Inches(3.15), Inches(6.15), Inches(3.55))
box(s, Inches(0.75), Inches(3.35), Inches(5.5), Inches(0.4), ["适合留下"], 16, GREEN, True)
box(s, Inches(0.75), Inches(3.9), Inches(5.5), Inches(2.5), [
    "·  不把 ODM 当渠道批发",
    "·  纪要、责任人、升级闭环清楚",
    "·  不把漏斗当业绩",
    "·  不与 FAE 抢技术结论",
], 14, INK, leading=8)
card(s, Inches(6.75), Inches(3.15), Inches(6.1), Inches(3.55))
box(s, Inches(7.05), Inches(3.35), Inches(5.5), Inches(0.4), ["必须补齐"], 16, RED, True)
box(s, Inches(7.05), Inches(3.9), Inches(5.5), Inches(2.5), [
    "·  品牌侧渗透几乎为零",
    "·  独立商务谈判不足",
    "·  售前还不能覆盖约 70%",
    "·  A 类项目周节点仍偏粗",
], 14, INK, leading=8)

# 10 三、90天 = 周报「下周计划」放大
s = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(s)
header(s, "三、未来规划  |  下阶段计划（90 天）", 10)
banner(s, "对应周报「下周计划」。四件必须发生，用节点考核。")
q = [
    ("01", "验证出结论", "至少 1 个目标项目完成模组/IQ 验证，书面结论：过 / 有条件过 / 不过及原因。"),
    ("02", "品牌线索", "画出品牌—华勤—模组三方图，拿到至少 1 个品牌侧或指定模组的有效窗口。"),
    ("03", "商务口径", "与产品、计划锁定华勤专用报价、交期、供给原则，现场不再口头承诺。"),
    ("04", "机制固化", "华勤项目月评：漏斗、风险、FAE 资源、下一步，继续进周报红黄绿。"),
]
for i, (n, t, d) in enumerate(q):
    top = Inches(1.55 + i * 1.28)
    card(s, Inches(0.45), top, Inches(12.4), Inches(1.18))
    box(s, Inches(0.7), top + Inches(0.32), Inches(1.0), Inches(0.5), [n], 20, GOLD, True)
    box(s, Inches(1.85), top + Inches(0.16), Inches(10.6), Inches(0.35), [t], 16, INK, True)
    box(s, Inches(1.85), top + Inches(0.58), Inches(10.6), Inches(0.45), [d], 13, MUTED)

# 11 三、12个月
s = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(s)
header(s, "三、未来规划  |  12 个月展望", 11)
banner(s, "把华勤做成样板账号：方法可复制，而不只是跟过一个客户。")
year = [
    ("项目结果", "1 个项目进入试产或定点路径；2–3 个稳定在验证池；杜绝只有拜访没有下一节点。"),
    ("关系结构", "项目组级关系：采购、硬件、影像、质量都有窗口和备份。"),
    ("产品策略", "2M 第二源破口，再向更高规格在研机型要验证机会，不一次铺太多料号。"),
    ("组织资产", "《华勤 CIS 导入包》：规格对照、setting、OTP/封装、FAQ、竞品，可交接。"),
    ("外溢价值", "打法复制到同类手机 ODM（如有规划），从管一个客户到管一类战场。"),
    ("自我能力", "独立完成约 70% 售前；复杂 IQ 协同 FAE；能主持华勤月度项目会。"),
]
for i, (t, d) in enumerate(year):
    r, c = divmod(i, 3)
    left = Inches(0.45 + c * 4.25)
    top = Inches(1.55 + r * 2.55)
    card(s, left, top, Inches(4.05), Inches(2.4))
    box(s, left + Inches(0.25), top + Inches(0.22), Inches(3.55), Inches(0.4), [t], 16, BLUE, True)
    box(s, left + Inches(0.25), top + Inches(0.75), Inches(3.55), Inches(1.4), [d], 13, MUTED, leading=5)

# 12 需协调 = 周报最后一页
s = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(s)
header(s, "三、未来规划  |  承诺与需协调事项", 12)
banner(s, "对应周报最后一页「需支持」。承诺可被周报检查。")
card(s, Inches(0.45), Inches(1.55), Inches(6.15), Inches(5.15))
box(s, Inches(0.75), Inches(1.75), Inches(5.5), Inches(0.4), ["我承诺"], 18, BLUE, True)
box(s, Inches(0.75), Inches(2.3), Inches(5.5), Inches(4.1), [
    "·  华勤项目进周报，红灯当周升级",
    "·  对外口径与对内事实一致",
    "·  不把送样数当业绩，只认下一节点",
    "·  90 天四件事按月复盘，未完成写原因",
    "·  与 FAE 按约定分工，不甩锅、不抢功",
], 15, INK, leading=10)
card(s, Inches(6.75), Inches(1.55), Inches(6.1), Inches(5.15), PALE)
box(s, Inches(7.05), Inches(1.75), Inches(5.5), Inches(0.4), ["请公司支持"], 18, GOLD, True)
box(s, Inches(7.05), Inches(2.3), Inches(5.5), Inches(4.1), [
    "·  产品：可主推华勤的规格与 roadmap",
    "·  FAE：验证期档期，问题 48h 响应",
    "·  商务：能进短名单的第二源价格与供给",
    "·  必要时协助品牌侧，避免只在 ODM 空转",
    "·  质量/交付：导入风险清单同步对外",
], 15, INK, leading=10)

# 13 口头 3 分钟提纲
s = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(s)
header(s, "答辩口头提纲（10–12 分钟）", 13)
banner(s, "按周报顺序讲，评委打断时回到「灯号和下一步」。")
rows = [
    ["段落", "时间", "要讲完的一句话"],
    ["开场 / 口径", "1 min", "在华勤卖 CIS，看三层客户；试用期看漏斗不看未发生的定点。"],
    ["工作总结", "4 min", "窗口和决策链已立；2M 第二源在验证；销售×FAE 闭环可检查。"],
    ["问题红灯", "2 min", "AVL 未破、无定点、商务口径未锁，已进 90 天计划。"],
    ["自我评价", "2 min", "匹配岗位；短板是品牌、商务、独立售前。"],
    ["规划与要支持", "2 min", "90 天四件事 + 周报检查；请产品/FAE/商务按节点给支持。"],
]
table(s, rows, Inches(0.45), Inches(1.55), Inches(12.4), Inches(4.4),
      [Inches(2.4), Inches(1.5), Inches(8.5)], center_cols=(1,))
box(s, Inches(0.45), Inches(6.15), Inches(12.4), Inches(0.7), [
    "预判提问：①为何没定点 → 周期与 AVL。②怎么打格科 → 第二源+供给+服务。③华勤和品牌谁说了算 → 品牌短名单，华勤落地，两边都要做。",
], 13, MUTED)

# 14 Close
s = prs.slides.add_slide(prs.slide_layouts[6])
sh = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
fill(sh, NAVY)
g = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(0.16), prs.slide_height)
fill(g, GOLD)
box(s, Inches(0.8), Inches(2.1), Inches(12), Inches(0.8), ["请批评指正"], 32, WHITE, True)
box(s, Inches(0.8), Inches(3.0), Inches(11.6), Inches(1.1), [
    "试用期把华勤的结构和漏斗立起来了。转正后按销售与技术服务部周报机制接受检查，",
    "用验证结论和定点路径，把这个账号做成公司手机 CIS 的样板。",
], 16, RGBColor(0xC5, 0xD0, 0xDC), leading=6)
ln = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(4.4), Inches(2.2), Inches(0.06))
fill(ln, GOLD)
box(s, Inches(0.8), Inches(4.7), Inches(11), Inches(0.7),
    ["销售与技术服务部  ·  CIS 销售经理  ·  华勤手机业务", "谢谢"], 14, GOLD, leading=4)
notes(s, "结束即提问。回答一律落到：现状灯号、原因、下一步、需要谁。")

out = "/workspace/CIS销售经理-华勤手机业务-转正答辩.pptx"
prs.save(out)
print("saved", out, "slides", len(prs.slides))
