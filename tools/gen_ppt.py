#!/usr/bin/env python3
"""生成 IoA 平台 B-EP1 答辩 PPT"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
import os

# ═══════════════════════════════════════════════════════
#  配色体系
# ═══════════════════════════════════════════════════════
CYAN = RGBColor(0x00, 0xF0, 0xFF)
PINK = RGBColor(0xFF, 0x2E, 0x93)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
DARK = RGBColor(0x0A, 0x0E, 0x17)
DARKER = RGBColor(0x06, 0x08, 0x10)
GRAY = RGBColor(0x88, 0x92, 0xA0)
LIGHT_GRAY = RGBColor(0x55, 0x5D, 0x6E)
GREEN = RGBColor(0x00, 0xE6, 0x76)
YELLOW = RGBColor(0xFF, 0xD7, 0x00)
RED = RGBColor(0xFF, 0x45, 0x55)
BG = DARK
CARD_BG = RGBColor(0x12, 0x18, 0x24)
BORDER = RGBColor(0x1E, 0x28, 0x36)

FONT_TITLE = '微软雅黑'
FONT_BODY = '微软雅黑'

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

# ═══════════════════════════════════════════════════════
#  辅助函数
# ═══════════════════════════════════════════════════════
def add_blank_slide():
    layout = prs.slide_layouts[6]  # blank
    slide = prs.slides.add_slide(layout)
    # 深色背景
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = BG
    return slide

def add_text_box(slide, left, top, width, height, text, font_size=18,
                 color=WHITE, bold=False, align=PP_ALIGN.LEFT, font_name=FONT_BODY):
    txBox = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(font_size)
    p.font.color.rgb = color
    p.font.bold = bold
    p.font.name = font_name
    p.alignment = align
    return txBox

def add_multi_text(slide, left, top, width, height, lines, default_size=16,
                   default_color=WHITE, line_spacing=1.3):
    """lines: list of (text, font_size, color, bold) or just str"""
    txBox = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = txBox.text_frame
    tf.word_wrap = True
    for i, line in enumerate(lines):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.space_after = Pt(6)
        if isinstance(line, str):
            p.text = line
            p.font.size = Pt(default_size)
            p.font.color.rgb = default_color
            p.font.bold = False
        else:
            text, size, color, bold = line[0], line[1] if len(line) > 1 else default_size, \
                                       line[2] if len(line) > 2 else default_color, \
                                       line[3] if len(line) > 3 else False
            p.text = text
            p.font.size = Pt(size)
            p.font.color.rgb = color
            p.font.bold = bold
        p.font.name = FONT_BODY
    return txBox

def add_card(slide, left, top, width, height, fill_color=CARD_BG):
    """圆角矩形卡片"""
    shape = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(left), Inches(top), Inches(width), Inches(height)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    shape.line.fill.background()
    # 调整圆角
    shape.adjustments[0] = 0.05
    return shape

def add_rect(slide, left, top, width, height, fill_color=CARD_BG):
    shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(left), Inches(top), Inches(width), Inches(height)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    shape.line.fill.background()
    return shape

def add_line(slide, x1, y1, x2, y2, color=BORDER, width=1.5):
    connector = slide.shapes.add_connector(
        1,  # straight
        Inches(x1), Inches(y1), Inches(x2), Inches(y2)
    )
    connector.line.color.rgb = color
    connector.line.width = Pt(width)
    return connector

def add_arrow_label(slide, left, top, text, color=CYAN):
    """小标签 + 箭头装饰"""
    shape = slide.shapes.add_shape(
        MSO_SHAPE.CHEVRON,
        Inches(left), Inches(top), Inches(0.7), Inches(0.35)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()
    tf = shape.text_frame
    tf.word_wrap = False
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(9)
    p.font.color.rgb = DARK
    p.font.bold = True
    p.font.name = FONT_BODY
    p.alignment = PP_ALIGN.CENTER
    return shape

def add_cyan_bar(slide, left, top, width, height=0.04):
    """青色装饰条"""
    shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(left), Inches(top), Inches(width), Inches(height)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = CYAN
    shape.line.fill.background()
    return shape

def add_page_number(slide, num, total):
    add_text_box(slide, 12.5, 7.1, 0.7, 0.3,
                 f'{num}/{total}', font_size=10, color=GRAY, align=PP_ALIGN.RIGHT)

def add_footer(slide):
    """底部装饰线"""
    add_rect(slide, 0, 7.3, 13.333, 0.01, CYAN)

def title_slide_content(slide, main_title, subtitle='', page_num=None, total=None):
    """通用标题区"""
    add_cyan_bar(slide, 0.8, 0.55, 0.4)
    add_text_box(slide, 1.4, 0.4, 10, 0.6, main_title,
                 font_size=30, color=WHITE, bold=True, font_name=FONT_TITLE)
    if subtitle:
        add_text_box(slide, 1.4, 0.95, 10, 0.35, subtitle,
                     font_size=14, color=GRAY)
    if page_num:
        add_page_number(slide, page_num, total)
    add_footer(slide)

TOTAL_PAGES = 13  # will update after final count

# ═══════════════════════════════════════════════════════
#  封面
# ═══════════════════════════════════════════════════════
slide = add_blank_slide()
# 装饰：左侧大色块
add_rect(slide, 0, 0, 0.08, 7.5, CYAN)
# 顶部横线
add_rect(slide, 0, 0, 13.333, 0.008, CYAN)
# 底部渐变区域
add_rect(slide, 0, 5.5, 13.333, 2, CARD_BG)
add_rect(slide, 0, 5.5, 13.333, 0.02, PINK)

# 作品名称
add_text_box(slide, 1.2, 1.5, 11, 1.0, 'IoA 分布式网络运维协同平台',
             font_size=44, color=WHITE, bold=True, font_name=FONT_TITLE)
# 副标题
add_multi_text(slide, 1.2, 2.6, 11, 1.0, [
    ('基于 MCP 协议的多智能体网络故障自愈系统', 22, CYAN, False),
])
# 标签
add_text_box(slide, 1.2, 3.4, 4, 0.4, 'B-EP1  智能体互联网创新攻关赛项',
             font_size=16, color=GRAY)
# 底部信息
add_multi_text(slide, 1.2, 5.8, 8, 1.5, [
    ('2026 中国高校计算机大赛 — 网络技术挑战赛', 14, GRAY, False),
    ('Python + FastAPI + LangGraph + DeepSeek + MCP', 12, LIGHT_GRAY, False),
])

# ═══════════════════════════════════════════════════════
#  目录
# ═══════════════════════════════════════════════════════
slide = add_blank_slide()
title_slide_content(slide, '汇报大纲', 'CONTENTS', 2, TOTAL_PAGES)

toc_items = [
    ('01', '项目背景与痛点', '传统运维为何需要 AI 化改造'),
    ('02', '系统架构总览', '四层架构 + 9 Agent 集群'),
    ('03', 'MCP 协议落地', '7 工具 · SSE 传输 · Windows 兼容'),
    ('04', 'Agent 协同与 DAG', '编排流程 · 调度引擎 · Bandit 路由'),
    ('05', '差异化修复策略', '6 种故障 · primary+fallback · 闭环验证'),
    ('06', '界面与测试', '五面板 Dashboard · 138 用例 · 并发稳定'),
    ('07', '创新点总结', '五大创新 · 比赛契合度'),
]

for i, (num, title, desc) in enumerate(toc_items):
    row = i // 2
    col = i % 2
    x = 0.8 + col * 6.2
    y = 1.6 + row * 1.3

    # 编号圆圈
    circle = slide.shapes.add_shape(
        MSO_SHAPE.OVAL, Inches(x), Inches(y + 0.05), Inches(0.5), Inches(0.5)
    )
    circle.fill.solid()
    circle.fill.fore_color.rgb = CYAN if i % 2 == 0 else PINK
    circle.line.fill.background()
    tf = circle.text_frame
    p = tf.paragraphs[0]
    p.text = num
    p.font.size = Pt(16)
    p.font.color.rgb = DARK
    p.font.bold = True
    p.font.name = FONT_BODY
    p.alignment = PP_ALIGN.CENTER

    add_text_box(slide, x + 0.7, y, 5, 0.35, title,
                 font_size=18, color=WHITE, bold=True)
    add_text_box(slide, x + 0.7, y + 0.4, 5, 0.3, desc,
                 font_size=12, color=GRAY)

TOTAL_PAGES = 11
# Fix page number on slide 2
for shape in slide.shapes:
    if hasattr(shape, 'text') and '/13' in shape.text:
        shape.text = f'2/{TOTAL_PAGES}'
        break

# ═══════════════════════════════════════════════════════
#  第3页：项目背景与痛点
# ═══════════════════════════════════════════════════════
slide = add_blank_slide()
title_slide_content(slide, '传统网络运维的三大痛点', '为什么需要智能体驱动的运维协同？', 3, TOTAL_PAGES)

pain_points = [
    ('⏱️ 响应延迟高', '从告警到修复平均耗时数十分钟\n人工跨系统操作，MTTR 居高不下', RED),
    ('🔗 跨域协同难', '4 域网络设备配置异构\n排障需反复切换工具，信息孤岛严重', YELLOW),
    ('🧠 经验难复用', '运维知识散落在个人脑中\n无法沉淀为系统级可调用能力', GRAY),
]

for i, (title, desc, accent) in enumerate(pain_points):
    x = 0.8 + i * 4.1
    add_card(slide, x, 1.6, 3.8, 3.5)
    # 顶部色条
    add_rect(slide, x, 1.6, 3.8, 0.06, accent)
    add_text_box(slide, x + 0.3, 1.9, 3.2, 0.4, title,
                 font_size=20, color=WHITE, bold=True)
    add_multi_text(slide, x + 0.3, 2.5, 3.2, 2.0, [
        (desc, 14, GRAY, False),
    ])

# 底部结论
add_card(slide, 0.8, 5.5, 11.7, 1.2, CARD_BG)
add_text_box(slide, 1.2, 5.7, 11, 0.8,
             '→ 核心命题：如何用智能体（Agent）替代人工操作，用标准化协议打通异构系统，实现网络故障全闭环自愈？',
             font_size=16, color=CYAN, bold=False)

# ═══════════════════════════════════════════════════════
#  第4页：系统架构总览
# ═══════════════════════════════════════════════════════
slide = add_blank_slide()
title_slide_content(slide, '系统架构：四层架构 + 九 Agent 集群', 'Python + FastAPI + LangGraph + DeepSeek + MCP', 4, TOTAL_PAGES)

# 四层架构 - 纵向排列
layers = [
    ('Agent 层', '9 Agent 集群', ['orchestrator ×1', 'monitor ×4', 'diagnoser ×1', 'repairer ×1', 'verifier ×1', 'reporter ×1'], CYAN),
    ('中间件层', 'IoAP 消息总线 + DAG 调度器', ['NL 意图解析 · 模板匹配', '拓扑排序 · 依赖调度 · retry', '语义路由 (UCB1 Bandit)'], PINK),
    ('协议层', 'MCP Server (port 9000)', ['7 个标准化工具 · SSE 传输', 'A2A Agent-to-Agent 协议', 'AutoToolClient 双协议自动降级'], CYAN),
    ('模拟层', 'Network Simulator (port 8001)', ['4 域拓扑 · 6 种故障注入', '实时指标生成 · WebSocket 推送', 'REST API + 状态管理'], PINK),
]

for i, (name, subtitle, details, accent) in enumerate(layers):
    y = 1.5 + i * 1.4
    # 层背景
    add_card(slide, 0.8, y, 11.7, 1.2, CARD_BG)
    # 左侧标签
    label = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(y), Inches(1.8), Inches(1.2)
    )
    label.fill.solid()
    label.fill.fore_color.rgb = accent
    label.line.fill.background()
    tf = label.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = name
    p.font.size = Pt(16)
    p.font.color.rgb = DARK
    p.font.bold = True
    p.font.name = FONT_BODY
    p.alignment = PP_ALIGN.CENTER
    p2 = tf.add_paragraph()
    p2.text = subtitle
    p2.font.size = Pt(10)
    p2.font.color.rgb = DARK
    p2.font.name = FONT_BODY
    p2.alignment = PP_ALIGN.CENTER

    # 详情
    detail_text = '    '.join(f'• {d}' for d in details)
    add_text_box(slide, 3.0, y + 0.15, 9, 0.9, detail_text,
                 font_size=13, color=GRAY)

# 右侧端口标注
ports = [
    ('8000', 'Middleware', 1.55),
    ('8001', 'Simulator', 2.95),
    ('9000', 'MCP Server', 4.35),
    ('GUI', 'Dashboard', 5.75),
]
for port, label, y_pos in ports:
    add_text_box(slide, 11.0, y_pos, 2, 0.3, f':{port}',
                 font_size=11, color=CYAN, font_name='Consolas')

# ═══════════════════════════════════════════════════════
#  第5页：MCP 协议落地
# ═══════════════════════════════════════════════════════
slide = add_blank_slide()
title_slide_content(slide, 'MCP 协议：7 工具 + SSE 传输 + 双协议降级', 'Model Context Protocol — Agent 与模拟器之间的标准化协议层', 5, TOTAL_PAGES)

# 左侧：MCP 工具列表
add_card(slide, 0.8, 1.5, 5.8, 5.2)
add_text_box(slide, 1.1, 1.65, 5, 0.35, 'MCP Server — 7 个标准化工具', font_size=16, color=CYAN, bold=True)

tools = [
    ('get_metrics', '单域指标采集'),
    ('get_all_metrics', '全域指标汇总'),
    ('get_topology', '拓扑结构查询'),
    ('inject_fault', '故障注入 (6 种)'),
    ('clear_fault', '指定故障清除'),
    ('list_faults', '活跃故障列表'),
    ('execute_repair', '修复操作 (5 策略)'),
]
for i, (tool, desc) in enumerate(tools):
    y = 2.3 + i * 0.58
    # 工具名
    add_text_box(slide, 1.3, y, 2.5, 0.3, tool,
                 font_size=13, color=WHITE, bold=True, font_name='Consolas')
    # 描述
    add_text_box(slide, 3.9, y, 2.5, 0.3, desc,
                 font_size=12, color=GRAY)
    # 分隔线
    if i < len(tools) - 1:
        add_rect(slide, 1.3, y + 0.45, 5.0, 0.005, BORDER)

# 右侧：协议特性
add_card(slide, 7.0, 1.5, 5.5, 5.2)

add_text_box(slide, 7.3, 1.65, 5, 0.35, '关键技术点', font_size=16, color=PINK, bold=True)

features = [
    ('SSE 传输模式',
     'Server-Sent Events 长连接推送\nMCP Python SDK 原生支持'),
    ('双协议自动降级',
     'AutoToolClient 优先 MCP\n不可用时自动回退 HTTP'),
    ('Windows 兼容修复',
     'CRLF → LF 行尾转换\nConnection: close 解决 keep-alive'),
    ('标准化 JSON-RPC',
     '符合 MCP 2025-03-26 协议规范\ntools/list · tools/call'),
]
for i, (title, desc) in enumerate(features):
    y = 2.2 + i * 1.15
    add_text_box(slide, 7.5, y, 4.5, 0.3, title,
                 font_size=14, color=CYAN, bold=True)
    add_multi_text(slide, 7.5, y + 0.35, 4.5, 0.7, [
        (desc, 12, GRAY, False),
    ])

# ═══════════════════════════════════════════════════════
#  第6页：Agent 协同与 DAG
# ═══════════════════════════════════════════════════════
slide = add_blank_slide()
title_slide_content(slide, 'Agent 协同：从自然语言到全闭环修复', 'Orchestrator → DAG → Scheduler → Agent → MCP Tool', 6, TOTAL_PAGES)

# 流程图 - 横向
flow_steps = [
    ('NL 输入', '华东链路拥塞\n请全流程修复', CYAN),
    ('意图解析', 'LLM 提取\n域+故障类型', CYAN),
    ('模板匹配', 'full_remediation\n5-stage DAG', PINK),
    ('DAG 调度', '拓扑排序\n依赖+并行', PINK),
    ('Agent 执行', '9 Agent 协同\nMCP 工具调用', CYAN),
    ('闭环验证', 'Verifier 复检\n达标确认', CYAN),
]

for i, (title, desc, color) in enumerate(flow_steps):
    x = 0.5 + i * 2.1
    add_card(slide, x, 1.8, 1.9, 2.4, CARD_BG)
    add_rect(slide, x, 1.8, 1.9, 0.05, color)
    add_text_box(slide, x + 0.1, 2.0, 1.7, 0.35, title,
                 font_size=14, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
    add_multi_text(slide, x + 0.1, 2.5, 1.7, 1.2, [
        (desc, 11, GRAY, False),
    ])
    # 箭头
    if i < len(flow_steps) - 1:
        add_text_box(slide, x + 1.92, 2.6, 0.25, 0.3, '→',
                     font_size=22, color=CYAN, align=PP_ALIGN.CENTER)

# 下半部分：DAG 引擎细节
add_card(slide, 0.8, 4.7, 5.8, 2.3)
add_text_box(slide, 1.1, 4.85, 5, 0.35, 'DAG 调度引擎', font_size=16, color=PINK, bold=True)
dag_details = [
    '• 模板系统：full_remediation / diagnose_only / health_check',
    '• 拓扑排序：自动解析 depends_on 依赖关系',
    '• 并行执行：无依赖节点并发调度（如 4 域 monitor 并发）',
    '• 失败重试：max_retries 机制 + 指数退避',
    '• 超时控制：per-node timeout_ms（默认 60-120s）',
]
add_multi_text(slide, 1.1, 5.35, 5.3, 1.8, [(d, 12, GRAY, False) for d in dag_details])

add_card(slide, 7.0, 4.7, 5.5, 2.3)
add_text_box(slide, 7.3, 4.85, 5, 0.35, 'Bandit 智能路由', font_size=16, color=CYAN, bold=True)
bandit_details = [
    '• UCB1 算法：在线学习各 Agent 执行成功率',
    '• 探索-利用平衡：新 Agent 获得尝试机会',
    '• 动态权重：根据历史延迟和成功率自适应调整',
    '• 跨域路由：同 capability 多实例自动选择最优',
]
add_multi_text(slide, 7.3, 5.35, 5.5, 1.8, [(d, 12, GRAY, False) for d in bandit_details])

# ═══════════════════════════════════════════════════════
#  第7页：差异化修复策略
# ═══════════════════════════════════════════════════════
slide = add_blank_slide()
title_slide_content(slide, '差异化修复：6 种故障 × primary + fallback 双策略', '六种故障类型全覆盖 + 闭环验证确保真实修复', 7, TOTAL_PAGES)

# 故障策略表
strategies = [
    ('link_congestion', '链路拥塞', 'traffic_shape', 'route_switch', '流量整形 → 路由切换'),
    ('link_outage', '链路中断', 'link_failover', 'route_switch', '链路切换 → 备选路由'),
    ('cpu_overload', 'CPU 过载', 'restart_service', 'traffic_shape', '服务重启 → 流量限速'),
    ('ddos', 'DDoS 攻击', 'acl_deploy', 'traffic_shape', 'ACL 规则 → 流量整形'),
    ('misconfig', '配置错误', 'restart_service', 'traffic_shape', '配置恢复 → 兜底整形'),
    ('device_failure', '设备故障', 'link_failover', 'traffic_shape', '链路切换 → 流量降级'),
]

# 表头
headers = ['故障类型', 'Primary 策略', 'Fallback 策略', '策略链']
col_x = [0.8, 3.2, 5.8, 8.4]
col_w = [2.2, 2.4, 2.4, 4.3]
for j, (h, x, w) in enumerate(zip(headers, col_x, col_w)):
    add_rect(slide, x, 1.5, w, 0.45, CYAN)
    add_text_box(slide, x + 0.1, 1.52, w - 0.2, 0.4, h,
                 font_size=13, color=DARK, bold=True, align=PP_ALIGN.CENTER)

for i, (eng, cn, primary, fallback, chain) in enumerate(strategies):
    y = 2.05 + i * 0.7
    bg_color = CARD_BG if i % 2 == 0 else RGBColor(0x16, 0x1E, 0x2A)
    for j, (val, x, w) in enumerate(zip([f'{eng}\n({cn})', primary, fallback, chain], col_x, col_w)):
        add_rect(slide, x, y, w, 0.62, bg_color)
        align = PP_ALIGN.CENTER if j > 0 else PP_ALIGN.LEFT
        size = 12 if j > 0 else 11
        color = CYAN if j in [1, 2] else (WHITE if j == 0 else GRAY)
        add_text_box(slide, x + 0.1, y + 0.08, w - 0.2, 0.5, val,
                     font_size=size, color=color, bold=(j <= 2), align=align)

# 底部验证闭环
add_card(slide, 0.8, 6.35, 11.7, 0.7, CARD_BG)
add_text_box(slide, 1.1, 6.45, 11, 0.5,
             '🔄 闭环验证：修复完成 → Verifier 复检指标 → 达标 ✓ / 未达标 → 回退 diagnose+repair 重试',
             font_size=14, color=CYAN)

# ═══════════════════════════════════════════════════════
#  第8页：界面展示
# ═══════════════════════════════════════════════════════
slide = add_blank_slide()
title_slide_content(slide, 'GUI 仪表盘：五面板一站式运维指挥中心', '拓扑可视化 → 实时指标 → Agent 状态 → DAG 记录 → 消息流', 8, TOTAL_PAGES)

panels = [
    ('🗺️ 拓扑面板', 'vis-network\n4 域拓扑实时渲染\n故障域颜色标记', CYAN),
    ('📊 指标面板', 'Chart.js\n延迟趋势 + 丢包率\n4 域切换查看', PINK),
    ('🤖 Agent 面板', '9 Agent 状态\n在线 + 域归属 +\n实时负载监控', CYAN),
    ('📋 DAG 面板', '执行流程记录\n节点展开查看详情\n每步耗时追踪', PINK),
    ('💬 消息流面板', 'IoAP 协议交互\n实时消息滚动\ncorrelation_id 追踪', CYAN),
]

for i, (title, desc, accent) in enumerate(panels):
    x = 0.5 + i * 2.55
    add_card(slide, x, 1.5, 2.35, 3.5, CARD_BG)
    add_rect(slide, x, 1.5, 2.35, 0.05, accent)
    add_text_box(slide, x + 0.15, 1.7, 2.05, 0.35, title,
                 font_size=15, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
    add_multi_text(slide, x + 0.15, 2.2, 2.05, 2.0, [
        (desc, 12, GRAY, False),
    ])

# 底部说明
add_card(slide, 0.8, 5.4, 11.7, 1.3)
add_text_box(slide, 1.1, 5.55, 11, 0.3,
             '设计理念：暗色"指挥中心"风格 · 霓虹粉 #FF2E93 + 青色 #00F0FF 双主色 · Material Icons 图标体系',
             font_size=13, color=GRAY)
add_text_box(slide, 1.1, 5.95, 11, 0.5,
             '技术栈：原生 JavaScript · vis-network 拓扑图 · Chart.js 指标图 · 全 SPA 单页应用 · 实时 WebSocket 推送',
             font_size=12, color=LIGHT_GRAY)

# ═══════════════════════════════════════════════════════
#  第9页：测试结果
# ═══════════════════════════════════════════════════════
slide = add_blank_slide()
title_slide_content(slide, '测试覆盖：138 用例 · 并发稳定 · 全链路闭环', 'pytest · 并发测试 · MCP 协议验证 · 端到端集成', 9, TOTAL_PAGES)

# 四大测试维度
test_cards = [
    ('138', '单元 + 集成测试', 'pytest 全部通过\n19 skipped (playwright)\n0 failed', GREEN),
    ('3/3', '并发稳定性', '3 域 3 故障并发执行\n全部 DAG 完成\n无竞态 / 死锁', CYAN),
    ('7/7', 'MCP 协议验证', '7 工具全部可用\nSSE + JSON-RPC\nWindows 兼容修复', PINK),
    ('100%', '端到端全链路', '故障注入 → NL → DAG\n→ 诊断 → 修复 → 验证\n→ 报告 全闭环', GREEN),
]

for i, (big, title, desc, accent) in enumerate(test_cards):
    x = 0.5 + i * 3.2
    add_card(slide, x, 1.5, 3.0, 4.3, CARD_BG)
    add_rect(slide, x, 1.5, 3.0, 0.06, accent)

    # 大数字
    add_text_box(slide, x + 0.2, 1.8, 2.6, 1.0, big,
                 font_size=48, color=accent, bold=True, align=PP_ALIGN.CENTER, font_name='Consolas')

    add_text_box(slide, x + 0.2, 2.9, 2.6, 0.35, title,
                 font_size=16, color=WHITE, bold=True, align=PP_ALIGN.CENTER)

    add_multi_text(slide, x + 0.3, 3.5, 2.4, 2.0, [
        (desc, 13, GRAY, False),
    ])

# 底部补充
add_text_box(slide, 0.8, 6.2, 11.7, 0.5,
             '稳定性：服务器连续运行无崩溃 · 平均 DAG 执行 8-12 秒 · Node.js --check 语法检查通过',
             font_size=13, color=GRAY, align=PP_ALIGN.CENTER)

# ═══════════════════════════════════════════════════════
#  第10页：创新点总结
# ═══════════════════════════════════════════════════════
slide = add_blank_slide()
title_slide_content(slide, '五大创新点', '技术特色与比赛赛题契合度分析', 10, TOTAL_PAGES)

innovations = [
    ('01', 'MCP/A2A 双重协议标准化', '7 个 MCP 工具 · SSE 传输 · Windows 兼容修复 · 双协议自动降级', CYAN),
    ('02', '差异化修复策略引擎', '6 种故障 · primary+fallback 双层策略 · 闭环验证防"假修复"', PINK),
    ('03', 'UCB1 Bandit 智能路由', '多臂老虎机在线学习 · 探索-利用平衡 · 动态权重收敛', CYAN),
    ('04', 'LangGraph 工作流编排', 'NL 意图解析 → 多步推理链 · 条件分支 + 错误恢复 · 优于单步 prompt', PINK),
    ('05', '双总线 + 跨域弹性架构', 'MemoryBus(开发) / NATS(生产) 零配置切换 · 按域水平扩展 Agent', CYAN),
]

for i, (num, title, desc, accent) in enumerate(innovations):
    y = 1.5 + i * 1.1
    add_card(slide, 0.8, y, 11.7, 0.95, CARD_BG)

    # 编号圆圈
    circle = slide.shapes.add_shape(
        MSO_SHAPE.OVAL, Inches(1.0), Inches(y + 0.2), Inches(0.55), Inches(0.55)
    )
    circle.fill.solid()
    circle.fill.fore_color.rgb = accent
    circle.line.fill.background()
    tf = circle.text_frame
    p = tf.paragraphs[0]
    p.text = num
    p.font.size = Pt(16)
    p.font.color.rgb = DARK
    p.font.bold = True
    p.font.name = FONT_BODY
    p.alignment = PP_ALIGN.CENTER

    add_text_box(slide, 1.8, y + 0.15, 4, 0.35, title,
                 font_size=18, color=WHITE, bold=True)
    add_text_box(slide, 1.8, y + 0.55, 10, 0.3, desc,
                 font_size=13, color=GRAY)

# ═══════════════════════════════════════════════════════
#  第11页：致谢
# ═══════════════════════════════════════════════════════
slide = add_blank_slide()
add_rect(slide, 0, 0, 13.333, 0.008, CYAN)
add_rect(slide, 0, 7.492, 13.333, 0.008, PINK)

add_text_box(slide, 0, 2.2, 13.333, 1.0, '感谢聆听',
             font_size=52, color=WHITE, bold=True, align=PP_ALIGN.CENTER, font_name=FONT_TITLE)

add_text_box(slide, 0, 3.4, 13.333, 0.5, 'IoA 分布式网络运维协同平台',
             font_size=24, color=CYAN, align=PP_ALIGN.CENTER, font_name=FONT_TITLE)

add_multi_text(slide, 0, 4.3, 13.333, 1.5, [
    ('B-EP1 智能体互联网创新攻关赛项', 16, GRAY, False),
    ('', 8, GRAY, False),
    ('MCP + Agent + DAG  ·  全闭环故障自愈  ·  9 Agent 集群', 14, LIGHT_GRAY, False),
])

add_text_box(slide, 0, 6.5, 13.333, 0.4, 'https://github.com/xiabai2004/IoA-Platform',
             font_size=13, color=GRAY, align=PP_ALIGN.CENTER, font_name='Consolas')

# ═══════════════════════════════════════════════════════
#  保存
# ═══════════════════════════════════════════════════════
output_dir = r'C:\Users\HZR\Desktop\IoA 分布式网络运维协同平台\docs'
os.makedirs(output_dir, exist_ok=True)
output_path = os.path.join(output_dir, 'IoA_平台_答辩PPT.pptx')
prs.save(output_path)
print(f'PPT saved: {output_path}')
print(f'Slides: {len(prs.slides)}')
print(f'File size: {os.path.getsize(output_path)} bytes')
