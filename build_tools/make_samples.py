"""生成测试样本：刻意做出各种"排版很糟糕"的文档，用来验证优化效果。

用法:  python build_tools/make_samples.py [输出目录]
"""

from __future__ import annotations

import os
import sys

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, RGBColor, Cm
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, "..", "samples")


def _messy_font(run, size=9, family="Calibri", color=None):
    run.font.name = family
    run.font.size = Pt(size)
    if color:
        run.font.color.rgb = color


def make_word_tender(path: str) -> None:
    doc = Document()
    # 故意使用极窄页边距 + 混乱字号颜色

    p = doc.add_paragraph()
    _messy_font(p.add_run("××市 2026 年度智慧园区信息化建设项目"), 12, "華康少女",
                RGBColor(0xFF, 0x00, 0x00))
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT

    doc.add_paragraph()

    def h(text, size=12, family="Calibri", color=None):
        p = doc.add_paragraph()
        r = p.add_run(text)
        _messy_font(r, size, family, color)
        return p

    def body(text):
        p = doc.add_paragraph()
        r = p.add_run(text)
        _messy_font(r, 9, "Calibri", RGBColor(0x33, 0x33, 0x33))
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.0
        return p

    h("第一章 招标公告", 11, "楷体", RGBColor(0x00, 0x00, 0xFF))
    body("本招标项目××市智慧园区信息化建设项目已由××市发展和改革委员会以×发改投〔2026〕"
         "18号文批准建设，资金来源为财政拨款与自筹资金，招标人为××市产业园区管理委员会。"
         "项目已具备招标条件，现对该项目的施工进行公开招标。")
    body("项目概况与招标范围：本项目总建筑面积约 98000 平方米，主要包括综合布线、"
         "安防监控、能耗管理、数据中心机房、园区运营管理平台等软硬件系统集成。"
         "计划工期 180 个日历天。招标范围包含深化设计、设备采购、安装调试、验收及质保服务。")
    h("1.1 投标人资格要求", 10)
    body("1.1.1 具有独立法人资格，具备电子与智能化工程专业承包二级及以上资质；")
    body("1.1.2 近三年（2023 年至今）承担过单个合同金额不低于 3000 万元的同类项目业绩不少于 2 项；")
    body("1.1.3 具有良好的商业信誉和健全的财务会计制度，未被列入失信被执行人名单。")

    h("第二章 评标办法", 11, "仿宋", RGBColor(0x80, 0x00, 0x80))
    body("本次评标采用综合评估法，满分 100 分，其中技术分 50 分、商务分 20 分、价格分 30 分。")
    body("评标委员会依照评分细则逐项打分，去掉一个最高分和一个最低分后取平均值为最终得分。")

    # 未设样式的表格：无底色、无表头重复、列宽拥挤
    rows = [
        ["序号", "评分项", "分值", "评分标准"],
        ["1", "技术方案", "25", "方案完整性、先进性、可实施性评价"],
        ["2", "项目团队", "10", "项目经理资质、团队人员配置与经历"],
        ["3", "实施计划", "10", "工期安排合理性、里程碑设置、风险应对"],
        ["4", "售后服务", "5", "响应时效、驻场安排、备品备件保障"],
        ["5", "投标报价", "30", "满足招标文件要求且最低价为基准价"],
        ["6", "企业业绩", "15", "同类项目数量与合同金额"],
        ["7", "体系认证", "5", "ISO9001 / ISO27001 / CMMI3 及以上"],
    ]
    table = doc.add_table(rows=len(rows), cols=4)
    for i, row in enumerate(rows):
        for j, val in enumerate(row):
            cell = table.cell(i, j)
            cell.text = ""
            run = cell.paragraphs[0].add_run(val)
            _messy_font(run, 8, "Calibri")

    h("第三章 投标文件编制要求", 11)
    body("投标文件应使用 A4 纸张双面打印，正本一份，副本四份，电子版一份（U 盘存储）。"
         "投标文件应编制目录及页码，装订成册并加盖骑缝章。")
    body("投标文件正本与副本必须一致，如出现不一致以正本为准。")
    doc.save(path)


def make_word_notice(path: str) -> None:
    doc = Document()
    p = doc.add_paragraph("关于做好2026年国庆节期间安全生产工作的通知")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _messy_font(p.runs[0], 10, "Calibri")

    doc.add_paragraph("各部门、各子公司：")
    doc.add_paragraph(
        "为切实做好国庆节期间安全生产工作，防范和遏制各类生产安全事故，"
        "根据上级有关文件精神，结合我单位实际，现就有关事项通知如下："
    )
    doc.add_paragraph("一、提高思想认识，压实安全责任")
    doc.add_paragraph(
        "各部门要充分认识节假日期间安全生产工作的极端重要性，主要负责人要靠前指挥，"
        "层层压实责任，确保各项措施落到实处。"
    )
    doc.add_paragraph("二、突出重点领域，开展隐患排查")
    doc.add_paragraph(
        "重点对消防设施、电气线路、电梯设备、特种作业场所等开展一次全覆盖排查，"
        "建立隐患台账，实行销号管理。"
    )
    doc.add_paragraph("三、强化值班值守，畅通信息报送")
    doc.add_paragraph(
        "严格落实领导带班和 24 小时值班制度，遇有突发事件第一时间报告，"
        "不得迟报、漏报、瞒报。"
    )
    doc.add_paragraph()
    p = doc.add_paragraph("××集团有限公司办公室")
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p = doc.add_paragraph("2026年9月20日")
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    doc.save(path)


def make_word_product(path: str) -> None:
    doc = Document()
    p = doc.add_paragraph("SmartGate 智能网关产品介绍")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _messy_font(p.runs[0], 11, "Calibri")
    doc.add_paragraph("1. 产品概述")
    doc.add_paragraph(
        "SmartGate 是一款面向园区与楼宇的智能物联网网关，支持多协议接入、"
        "边缘计算与云端协同，可显著降低现场部署与运维成本。"
    )
    doc.add_paragraph("2. 核心特性")
    doc.add_paragraph("• 支持 Modbus / BACnet / MQTT 等 12 种协议")
    doc.add_paragraph("• 内置边缘规则引擎，毫秒级响应")
    doc.add_paragraph("• 工业级宽温设计，-20℃ ~ 70℃ 稳定运行")
    doc.add_paragraph("• 双电源冗余与断网续传")
    doc.add_paragraph("3. 技术规格")
    rows = [
        ["参数", "规格", "备注"],
        ["处理器", "四核 ARM Cortex-A55 1.8GHz", ""],
        ["内存 / 存储", "2GB DDR4 / 16GB eMMC", "可扩展"],
        ["以太网口", "2 × 10/100/1000M", "支持冗余"],
        ["串口", "RS485 × 4, RS232 × 2", "光电隔离"],
        ["协议支持", "Modbus, BACnet, OPC UA, MQTT", "内置驱动"],
    ]
    t = doc.add_table(rows=len(rows), cols=3)
    for i, row in enumerate(rows):
        for j, val in enumerate(row):
            t.cell(i, j).text = val
            for r in t.cell(i, j).paragraphs[0].runs:
                _messy_font(r, 8)
    doc.add_paragraph("4. 应用场景")
    doc.add_paragraph(
        "广泛应用于智慧园区、智慧工厂、楼宇自控、能源管理等领域，支持按需定制。"
    )
    doc.save(path)


def make_excel_tender(path: str) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "评分明细表"
    ws["A1"] = "××项目综合评分明细表"
    ws.merge_cells("A1:F1")
    ws["A1"].font = Font(size=9, name="Calibri")
    ws["A1"].alignment = Alignment(horizontal="left")

    headers = ["序号", "评分项", "分值", "评分标准", "投标人A", "投标人B"]
    for j, h in enumerate(headers, 1):
        c = ws.cell(row=2, column=j, value=h)
        c.font = Font(size=8)
        c.alignment = Alignment(horizontal="left")
    data = [
        [1, "技术方案完整性", 25, "对项目需求全覆盖，方案先进可实施", 22, 20],
        [2, "项目团队配置", 10, "项目经理具备高级职称及同类项目经验", 8, 9],
        [3, "实施计划合理", 10, "工期安排、里程碑、风险应对措施", 9, 7],
        [4, "售后服务保障", 5, "2小时响应，驻场服务，备件保障", 4, 5],
        [5, "投标报价", 30, "以最低有效报价为基准价计算", 28.5, 30],
        [6, "企业同类业绩", 15, "近三年同类项目合同及验收材料", 12, 13],
        [7, "体系认证", 5, "ISO9001、ISO27001、CMMI3及以上", 5, 3],
    ]
    for i, row in enumerate(data, 3):
        for j, v in enumerate(row, 1):
            c = ws.cell(row=i, column=j, value=v)
            c.font = Font(size=8)
    # 超长文本列，测试列宽适配
    ws.cell(row=11, column=1, value="备注")
    ws.cell(row=11, column=2,
            value="本评分表仅用于内部评标参考，最终结果以评标委员会签署的评标报告为准；"
                  "各项评分需由评标委员会三名以上成员独立打分后取平均值。")
    ws.merge_cells("B11:F11")
    wb.save(path)


def make_excel_inventory(path: str) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "设备清单"
    for j, h in enumerate(["序号", "设备名称", "型号", "数量", "单价(元)", "金额(元)"], 1):
        ws.cell(row=1, column=j, value=h)
    items = [
        ("1", "核心交换机", "S6520X-30QC-EI", 4, 28600, 114400),
        ("2", "汇聚交换机", "S5560X-30C-EI", 12, 9800, 117600),
        ("3", "接入交换机", "S5130S-28S-EI", 96, 2350, 225600),
        ("4", "防火墙", "SecPath F1080", 2, 68000, 136000),
        ("5", "无线 AP", "WA6320-D", 320, 860, 275200),
        ("6", "光模块 10G", "SFP-XG-LX-SM1310", 240, 320, 76800),
    ]
    for i, item in enumerate(items, 2):
        for j, v in enumerate(item, 1):
            c = ws.cell(row=i, column=j, value=v)
            c.font = Font(size=8)
        ws.cell(row=i, column=6, value=f"=D{i}*E{i}")
    ws.cell(row=9, column=1, value="合计")
    ws.merge_cells("A9:E9")
    ws.cell(row=9, column=6, value="=SUM(F2:F7)")
    wb.save(path)


def make_word_contract(path: str) -> None:
    doc = Document()
    p = doc.add_paragraph("设备采购及安装合同")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _messy_font(p.runs[0], 10)
    doc.add_paragraph("甲方：××科技有限公司")
    doc.add_paragraph("乙方：××电子工程有限公司")
    doc.add_paragraph("根据《中华人民共和国民法典》及相关法律法规，甲乙双方本着平等自愿、"
                      "诚实信用的原则，就设备采购及安装事宜达成如下协议：")
    doc.add_paragraph("第一条 标的物")
    doc.add_paragraph("乙方向甲方供应下列设备并负责安装调试，具体清单附后。")
    doc.add_paragraph("第二条 合同价款")
    doc.add_paragraph("本合同总价款为人民币壹佰贰拾万元整（¥1,200,000.00），含税含运费。")
    doc.add_paragraph("第三条 交付与验收")
    doc.add_paragraph("乙方应于合同签订后 60 日内完成交付，甲方应在收到货物后 7 日内组织验收。")
    doc.add_paragraph("第四条 质量保证")
    doc.add_paragraph("设备质保期为验收合格之日起 36 个月，质保期内非人为损坏由乙方免费维修。")
    doc.add_paragraph("第五条 违约责任")
    doc.add_paragraph("任何一方违约，应向守约方支付合同总价款 5% 的违约金。")
    doc.add_paragraph("甲方（盖章）：　　　　　　乙方（盖章）：")
    doc.add_paragraph("签订日期：　　年　　月　　日")
    doc.save(path)


def make_word_process(path: str) -> None:
    doc = Document()
    p = doc.add_paragraph("网络设备巡检作业流程")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph("一、目的")
    doc.add_paragraph("规范日常巡检操作，及时发现并处置设备隐患，保障网络稳定运行。")
    doc.add_paragraph("二、适用范围")
    doc.add_paragraph("适用于园区核心、汇聚及接入层网络设备的日常巡检工作。")
    doc.add_paragraph("三、作业步骤")
    doc.add_paragraph("1. 准备阶段")
    doc.add_paragraph("携带笔记本电脑、千兆网线、串口线、标签打印机等工具。")
    doc.add_paragraph("2. 现场检查")
    doc.add_paragraph("检查设备指示灯状态、风扇噪声、机柜温湿度及线缆标签完整性。")
    doc.add_paragraph("3. 登录核查")
    doc.add_paragraph("通过 SSH 登录设备，查看 CPU、内存、温度、接口错误包计数。")
    doc.add_paragraph("4. 记录归档")
    doc.add_paragraph("填写巡检记录表，异常项 24 小时内上报并跟踪闭环。")
    doc.save(path)


def make_word_spec(path: str) -> None:
    doc = Document()
    p = doc.add_paragraph("园区网络系统技术说明")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph("1 概述")
    doc.add_paragraph("本技术说明描述园区网络系统的总体架构、关键技术与性能指标要求。")
    doc.add_paragraph("1.1 设计原则")
    doc.add_paragraph("系统应具备高可用、可扩展、易维护等特性，满足未来 5 年业务增长需求。")
    doc.add_paragraph("1.2 总体架构")
    doc.add_paragraph("采用三层架构：核心层、汇聚层、接入层。核心层部署两台高可用交换机做 VRRP 冗余。")
    doc.add_paragraph("2 技术指标")
    doc.add_paragraph("2.1 性能要求")
    doc.add_paragraph("核心单板吞吐量不低于 2Tbps，时延小于 10μs。")
    doc.add_paragraph("2.2 可靠性要求")
    doc.add_paragraph("关键链路可用性不低于 99.99%，故障切换时间小于 50ms。")
    doc.save(path)


def make_word_messy(path: str) -> None:
    """极端混乱文档：全篇同字号、无页眉、颜色刺眼、行距 0.8。"""
    doc = Document()
    for i in range(1, 40):
        p = doc.add_paragraph()
        if i % 7 == 1:
            text = f"第{i // 7 + 1}节 内容说明"
            color = RGBColor(0xFF, 0x00, 0xFF)
        else:
            text = ("这是一段用于测试自动识别能力的中文正文段落，包含数字 123 与英文 "
                    "English words，用于验证字体、字号、行距、缩进与分页是否被正确优化。")
            color = RGBColor(0x00, 0x80, 0x00)
        r = p.add_run(text)
        _messy_font(r, 8, "Times New Roman", color)
        p.paragraph_format.line_spacing = 0.8
        p.paragraph_format.left_indent = Cm(0.3 * (i % 4))
    doc.save(path)


def main() -> None:
    os.makedirs(OUT, exist_ok=True)
    jobs = [
        (make_word_tender, "招标文件_原始.docx"),
        (make_word_notice, "发布通知_原始.docx"),
        (make_word_contract, "采购合同_原始.docx"),
        (make_word_product, "产品介绍_原始.docx"),
        (make_word_process, "技术流程_原始.docx"),
        (make_word_spec, "技术说明_原始.docx"),
        (make_word_messy, "混乱排版_原始.docx"),
        (make_excel_tender, "评分明细_原始.xlsx"),
        (make_excel_inventory, "设备清单_原始.xlsx"),
    ]
    for fn, name in jobs:
        path = os.path.join(OUT, name)
        fn(path)
        print("生成", path, os.path.getsize(path), "bytes")


if __name__ == "__main__":
    main()
