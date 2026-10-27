import datetime
import math
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.properties import PageSetupProperties

FONT = "微软雅黑"
C_TITLE = "FF665200"
C_INFO = "FFFFF9E3"
C_HEAD = "FFF2F2F2"
C_LINE = "FFD9DEE7"
C_ZEBRA = "FFF7F9FC"
C_PHASE = "FFEAF2FF"

thin = Side(style="thin", color=C_LINE)
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

TASKS = [
 ("启动与前置条件确认", [
  (1, "设备到货清关、开箱验货及IT环境确认（网段、IP、防火墙、端口、远程工具）", "《到货清点单》《IT环境确认单》", 2, 0, "递缇项目经理", "双方", "待客户确认",
   "设备到达墨西哥并完成清关，具备开箱验收条件", "扫码枪实到10把（缺5把）；一体机后到；清关由甲方负责"),
 ]),
 ("现场硬件与网络环境搭建", [
  (2, "服务器机柜就位，安装服务器、显示器、键鼠、UPS、交换机", "机柜与设备安装记录", 2, 2, "以弗德安装组", "以弗德（报价含2人8天）", "未开始",
   "机柜、服务器、UPS、交换机到位，供电正常", "交换机、一体机、扫码枪、大屏、支架、辅材由以弗德提供"),
  (3, "服务器上电，检查硬件状态、硬盘、内存、网口、RAID", "上电自检记录", 1, 4, "递缇实施工程师", "递缇", "未开始",
   "服务器到位，供电正常", "机房接地与UPS带载需先确认"),
  (4, "安装Windows Server 2016，激活、打补丁，设置主机名、固定IP、防火墙、远程桌面", "系统安装记录", 1, 5, "递缇实施工程师", "递缇", "未开始",
   "服务器硬件自检正常，安装介质可用", ""),
  (5, "安装SQL Server 2019/2016及以上，配置实例、账号、端口、备份策略", "数据库配置记录", 1, 6, "递缇实施工程师", "递缇", "未开始",
   "Windows Server安装完成", "协议约定MS SQL Server 2019"),
  (6, "配置网络交换机，划分/确认MES网段，连接服务器、工位一体机、大屏、PLC及设备", "网络拓扑与IP规划表", 2, 4, "递缇实施工程师", "甲方提供交换机，双方配置", "未开始",
   "交换机到位，最终用户网络参数确认", "交换机为甲方供货，需甲方开放配置权限"),
  (7, "安装15台工位一体机及支架，上电，配置网络", "15站安装上电记录", 4, 6, "以弗德安装组", "以弗德（一体机15台已到货）", "未开始",
   "一体机、支架到位，工位供电与网络到位", "★一体机原未发运、后补发到货，集中安装，窗口后移"),
  (8, "安装并调试扫码枪：到场10把有线/无线连接、扫码识别、条码类型；缺口5把按借用/串口枪/补发处置", "扫码枪调试记录", 3, 8, "递缇实施工程师", "递缇+甲方", "待客户确认",
   "扫码枪到位，工位一体机系统就绪", "★实到10把，缺5把；斑马清关受阻，建议递缇人员随行带1把串口枪，其余向其他工位借用或本地采购"),
  (9, "安装75寸大屏电视+盒式PC，配置浏览器、网络、看板显示", "大屏显示确认", 1, 10, "以弗德安装组", "以弗德", "未开始",
   "大屏、盒式PC到位，网络连通", ""),
  (10, "配合调试上线打码/标签打印：标签模板、条码规则、总成二维码打印", "打码打印测试记录", 2, 10, "递缇实施工程师", "甲方提供打码设备", "待客户确认",
   "打印/打码设备到位，条码规则确认", "★合同仅含标签打印与条码规则，打码硬件未在供货清单，需甲方提供"),
  (11, "检查桥架、网线、电源线等安装辅材，确认工位网络与供电稳定", "辅材与布线检查表", 2, 5, "以弗德安装组", "以弗德", "未开始",
   "辅材到位，现场施工完成", ""),
 ]),
 ("MES软件部署与基础配置", [
  (12, "部署MES服务端：软件包、组件、依赖环境，初始化数据库", "服务端部署清单", 2, 12, "递缇实施工程师", "递缇", "未开始",
   "服务器系统、SQL Server就绪", ""),
  (13, "部署工控机客户端到15个工位一体机，配置启动项、客户端参数", "客户端部署记录", 2, 12, "递缇实施工程师", "递缇", "未开始",
   "工位一体机系统与网络就绪", "一体机后到货，客户端镜像建议用同型号测试机提前验证"),
  (14, "配置系统语言：中文、西班牙文切换", "双语切换确认", 1, 14, "递缇实施工程师", "递缇", "未开始",
   "MES服务端部署完成", "协议要求软件支持中西文切换"),
  (15, "配置用户管理：用户、角色、权限、员工、工位技能、班组长/质检/出库等角色", "权限配置确认单", 1, 14, "递缇实施工程师", "递缇", "未开始",
   "最终用户提供组织架构与权限需求", ""),
  (16, "配置主动轮工艺：产品型号、编号、组件、厂商、重量/扭力/动平衡上下限", "主动轮工艺配置单", 1, 15, "递缇实施工程师", "递缇", "未开始",
   "最终用户提供主动轮工艺参数", ""),
  (17, "配置从动轮工艺：产品型号、编号、组件、厂商、工艺参数上下限", "从动轮工艺配置单", 1, 16, "递缇实施工程师", "递缇", "未开始",
   "最终用户提供从动轮工艺参数", ""),
  (18, "配置条码规则、标签模板、上线打印功能", "条码规则与标签模板", 1, 16, "递缇实施工程师", "递缇", "未开始",
   "条码规则、标签模板确认", ""),
  (19, "配置工艺路线、工序、工位、启用状态、异常关联工位、干涉工位跳过逻辑", "工艺路线配置单", 1, 17, "递缇实施工程师", "递缇", "未开始",
   "工艺路线、工位定义确认", ""),
  (20, "配置电子作业指导书：PDM图、预装提示、按车型/工位/工序切换", "ESOP配置确认", 1, 18, "递缇实施工程师", "递缇", "未开始",
   "PDM图、作业指导书资料齐全", ""),
  (21, "配置拧紧工艺：拧紧枪、程序号、拧紧数量、不合格次数、使能控制、数据采集", "拧紧工艺配置单", 1, 18, "递缇实施工程师", "递缇", "未开始",
   "拧紧枪接口协议、程序号确认", "ATLAS拧紧枪1把（主动轮）"),
  (22, "配置设备采集接口：伺服压机5台、拧紧枪1把、检测设备2台、PLC通讯参数", "接口配置清单", 2, 19, "递缇实施工程师", "递缇", "未开始",
   "设备接口协议、通讯参数确认", "压机5台为CVT二线；国外调试费用未含，需甲方/设备商配合"),
  (23, "配置订单管理：生产计划、工单、计划编号规则、终端客户、产品型号、数量", "订单规则配置单", 1, 19, "递缇实施工程师", "递缇", "未开始",
   "订单管理规则确认", ""),
  (24, "配置质量报表、SPC统计、电子看板、日志管理、追溯查询", "报表与看板配置单", 1, 20, "递缇实施工程师", "递缇", "未开始",
   "报表格式、看板内容确认", ""),
  (25, "配置MES与工厂MES/SAP预留接口、PLC标准接口协议", "接口文档", 1, 21, "递缇实施工程师", "递缇", "未开始",
   "工厂MES/SAP接口文档、PLC协议确认", "预留接口，联调视甲方系统可用性"),
 ]),
 ("工位单机与设备采集调试", [
  (26, "服务器本地功能验证：登录、订单、工艺、报表、追溯、用户权限、日志", "功能验证记录", 1, 22, "递缇实施工程师", "递缇", "未开始",
   "基础配置完成", ""),
  (27, "逐工位验证客户端：启动、扫码、SOP显示、防错、过站、数据上传", "工位签认表×15", 2, 23, "递缇实施工程师", "递缇", "未开始",
   "工位客户端部署完成，网络连通", "缺枪工位先用手工录入模式，签认表标注"),
  (28, "上线打码/条码绑定测试：总成二维码生成、打印、扫描、绑定", "打码绑定测试记录", 1, 23, "递缇实施工程师", "双方", "未开始",
   "打码设备、扫码枪就绪", ""),
  (29, "零件条码采集测试：零件二维码、批次码、重复校验、数量校验、特征码校验", "条码防错测试记录", 1, 24, "递缇实施工程师", "递缇", "未开始",
   "零件条码规则确认，扫码枪就绪", "★缺5把扫码枪的工位无法完成条码防错验证"),
  (30, "过程数据采集测试：进站、出站、过站信息、工单号、序列号、工站、人员、时间", "过站数据核对", 1, 25, "递缇实施工程师", "递缇", "未开始",
   "工位客户端与PLC通讯正常", ""),
  (31, "拧紧数据采集测试：扭矩、角度、曲线、合格状态、螺栓遗漏查询", "拧紧曲线核对", 1, 26, "递缇实施工程师", "递缇", "未开始",
   "拧紧枪通讯正常，程序号配置完成", "ATLAS 1把（主动轮）"),
  (32, "伺服压机数据采集验证：压力、位移、通讯稳定性、数据上传", "压机采集记录", 2, 26, "甲方/设备商主责，递缇远程协助", "甲方/设备商", "待客户确认",
   "压机通电、网络连通、接口协议开放", "★报价明确5台压机仅国内调试、不含国外调试费用，现场责任界面需书面确认"),
  (33, "检测设备数据采集验证：2台检测设备数据上传、结果绑定", "检测数据记录", 1, 27, "递缇实施工程师", "递缇", "未开始",
   "检测设备接口开放、通讯正常", ""),
  (34, "电子看板验证：订单内容、计划产量、实际产量、不良率、工位状态、SPC图", "看板上线确认", 1, 28, "递缇实施工程师", "递缇", "未开始",
   "看板数据源配置完成，大屏就绪", ""),
  (35, "质量报告与追溯验证：正反向追溯、树形展示、报表导出Excel、工号姓名显示", "追溯报表核对", 1, 28, "递缇实施工程师", "递缇", "未开始",
   "质量数据采集完成", ""),
  (36, "性能验证：操作延时≤1秒、数据库查询≤4秒、并发访问、故障恢复、数据不丢失", "性能测试报告", 1, 29, "递缇实施工程师", "递缇", "未开始",
   "系统稳定运行，测试数据准备", "判定方法需双方事先约定"),
 ]),
 ("全线联调与试生产", [
  (37, "MES与PLC联调：工件到位、保存数据、保存完成、验证完成、设备监控、生产控制", "PLC联调记录", 2, 30, "递缇实施工程师", "双方", "未开始",
   "单机调试完成，PLC通讯正常", ""),
  (38, "MES与工厂MES/SAP接口联调（如有）", "接口联调记录", 1, 31, "递缇实施工程师", "双方", "未开始",
   "工厂MES/SAP接口可用", "甲方系统不具备时列为遗留项"),
  (39, "全线空运行：订单下发→打码→扫码→装配→数据采集→过站→下线→追溯", "空运行记录", 1, 32, "双方", "双方", "未开始",
   "各工位、设备、PLC联调完成", ""),
  (40, "带产品试生产：小批量验证，记录MES问题、设备采集问题、网络问题", "试生产问题清单", 2, 33, "双方", "双方", "未开始",
   "空运行无重大异常，试生产物料准备", "试生产物料由甲方准备"),
  (41, "异常场景测试：设备异常、质量异常、数据异常、安灯报警、不合格拦截、高权限跳站", "异常场景测试记录", 1, 33, "递缇实施工程师", "递缇", "未开始",
   "异常处理逻辑配置完成", ""),
  (42, "生产报表、SPC、看板、日志、用户权限综合验证", "综合验证记录", 1, 35, "递缇实施工程师", "递缇", "未开始",
   "试生产数据积累", ""),
  (43, "分角色培训：操作工、班组长、质量、维护，并交付用户操作手册", "《培训计划》《用户操作手册》", 2, 35, "递缇实施工程师", "递缇", "未开始",
   "系统功能验证完成", "★原计划缺培训任务，按交付物要求补充"),
  (44, "预验收：按技术协议在杭州国内工厂进行；如已完成则做资料归档与遗留项确认", "《项目预验收报告》", 1, 36, "双方", "双方", "待确认",
   "杭州阶段单机部署完成", "★需确认预验收是否已随一体机后到而顺延"),
  (45, "现场实施完成确认与交接：遗留问题清单、备件与工具移交", "现场实施完成确认书", 1, 37, "递缇项目经理", "双方", "未开始",
   "试生产与培训完成", ""),
 ]),
 ("试运行、验收与运维", [
  (46, "远程调试：缺陷修复、参数优化、陪产支持（1人1个月）", "调试周报/问题闭环记录", 20, 38, "递缇实施工程师", "递缇（报价含1人1个月）", "未开始",
   "现场交接完成，远程通道可用", "按墨西哥时差排班"),
  (47, "试运行：连续生产，问题闭环，免费修改与升级", "《系统运行报告》", 90, 38, "双方", "双方", "未开始",
   "远程调试启动", "生产3个月或1000台，以先到为准"),
  (48, "终验收：按技术协议逐项验收", "《项目终验收报告》", 1, 128, "双方", "双方", "未开始",
   "试运行达标", "不组织验收则视为合格；争议期间除外"),
  (49, "质保与运维：终验收合格起1年，7×24支持，4小时内出方案", "《系统运维报告》", 365, 129, "递缇", "递缇", "未开始",
   "终验收合格", "质保期1年"),
 ]),
]

# 国内已完成一轮调试：原第4条（Windows Server）与第5条（SQL Server）合并为“系统与数据库复核”
MERGED4 = ("服务器系统与数据库复核（国内已安装）：核对主机名、固定IP、防火墙、远程桌面、实例、账号、端口、备份策略；异常时按镜像或配置基线恢复",
           "系统与数据库复核记录", 1, 5, "递缇实施工程师", "递缇", "未开始",
           "服务器硬件自检正常；国内已完成系统与数据库安装，具备镜像或配置基线",
           "★国内已调过一轮，原“装Windows Server”“装SQL Server”两条合并为复核；若硬件变更再重新安装")
_new, _n = [], 0
for _ph, _items in TASKS:
    _lst = []
    for _t in _items:
        if _t[0] == 5:
            continue
        _n += 1
        _lst.append((_n,) + MERGED4 if _t[0] == 4 else (_n,) + _t[1:])
    _new.append((_ph, _lst))
TASKS = _new

PHASE_COLOR = {
 "启动与前置条件确认": "FFD9E2F3",
 "现场硬件与网络环境搭建": "FFD7EDD9",
 "MES软件部署与基础配置": "FFFFE7C7",
 "工位单机与设备采集调试": "FFE4D5F7",
 "全线联调与试生产": "FFCCEFF2",
 "试运行、验收与运维": "FFF9D6D9",
}

wb = Workbook()
ws = wb.active
ws.title = "计划"

# ---------- 说明 sheet（先建，便于公式引用） ----------
ws0 = wb.create_sheet("说明与修订")
ws0.sheet_view.showGridLines = False
ws0.merge_cells("A1:C1")
ws0["A1"] = "墨西哥CVT总成装配线MES系统 — 现场实施计划（说明与修订）"
ws0["A1"].font = Font(name=FONT, size=14, bold=True, color="FFFFFFFF")
ws0["A1"].fill = PatternFill("solid", fgColor=C_TITLE)
ws0["A1"].alignment = Alignment(horizontal="center", vertical="center")
ws0.row_dimensions[1].height = 38

ws0["A3"] = "一、基线参数"
ws0["A3"].font = Font(name=FONT, size=11, bold=True)
ws0["A4"], ws0["B4"], ws0["C4"] = "参数", "值", "说明"
ws0["B5"] = datetime.datetime(2026, 11, 2)
ws0["B5"].number_format = "yyyy-mm-dd"
ws0["A5"] = "计划开始日"
ws0["C5"] = "★修改此格，计划页所有日期与甘特周次自动联动"
ws0["A6"] = "现场实施周期（自然日）"
ws0["B6"] = "=MAX(计划!G5:G49)-MIN(计划!F5:F49)+1"
ws0["C6"] = "到货清点至现场交接（含并行作业）"
ws0["A7"] = "试运行周期（自然日）"
ws0["B7"] = 90
ws0["C7"] = "生产3个月或1000台，以先到为准"
ws0["A8"] = "远程调试投入"
ws0["B8"] = "1人×1个月"
ws0["C8"] = "报价已含"
for r in range(4, 9):
    for c in range(1, 4):
        cell = ws0.cell(r, c)
        cell.border = BORDER
        cell.font = Font(name=FONT, size=10, bold=(r == 4))
        cell.alignment = Alignment(vertical="center", wrap_text=True)
        if r == 4:
            cell.fill = PatternFill("solid", fgColor=C_HEAD)
ws0["B5"].font = Font(name=FONT, size=10, bold=True, color="FF0000FF")
ws0["B5"].fill = PatternFill("solid", fgColor="FFFFF2CC")

ws0["A10"] = "二、关键里程碑"
ws0["A10"].font = Font(name=FONT, size=11, bold=True)
mh = ["里程碑", "计划日期", "判定标准"]
for i, v in enumerate(mh):
    c = ws0.cell(11, i + 1, v)
    c.border = BORDER
    c.fill = PatternFill("solid", fgColor=C_HEAD)
    c.font = Font(name=FONT, size=10, bold=True)
    c.alignment = Alignment(horizontal="center", vertical="center")
MS_START = 12
MILESTONES = [
 ("到货清点完成", 1, "开箱验货单签认"),
 ("机房上电完成", 3, "服务器自检通过"),
 ("15工位客户端就绪", 12, "15台可登录"),
 ("MES配置完成", 24, "基础数据与工艺确认"),
 ("单机与采集调试完成", 35, "采集数据入库"),
 ("试生产完成", 39, "问题清单闭环"),
 ("现场交接", 44, "完成确认书签署"),
 ("终验收", 47, "按技术协议逐项"),
]

ws0["A21"] = "三、对原表格已填内容的修订（逐条）"
ws0["A21"].font = Font(name=FONT, size=11, bold=True)
ws0["A22"], ws0["B22"], ws0["C22"] = "序号", "原内容 / 问题", "修订"
for i in range(1, 4):
    c = ws0.cell(22, i)
    c.border = BORDER
    c.fill = PatternFill("solid", fgColor=C_HEAD)
    c.font = Font(name=FONT, size=10, bold=True)
    c.alignment = Alignment(horizontal="center", vertical="center")
FIXES = [
 ("“安装15把扫码枪、调试有线/无线连接”", "实到10把，缺5把：改为“10把安装调试＋5把缺口处置（借用/串口枪/补发）”，并标“待客户确认”"),
 ("“安装15台一体机及支架”", "一体机原未发运、后补发：改为“集中安装”，安装窗口后移，前置条件加“到货时间确认”"),
 ("“伺服压机数据采集验证”未注范围", "报价明确5台压机“国内调试，不含国外调试费用”：责任方改为甲方/设备商主责，递缇远程协助"),
 ("“安装上线打印/打码设备”", "合同仅含“标签打印、条码规则管理”，打码硬件未在供货清单：改为“配合调试”，设备由甲方提供"),
 ("无培训任务", "按交付物要求补充“分角色培训＋《用户操作手册》”，2天"),
 ("无预验收/终验收/质保节点", "补充预验收（杭州）、终验收、质保与7×24运维，共3条"),
 ("无文档交付对应", "文档按合同交付物清单单独提交（《项目实施方案》《系统需求报告》《蓝图设计》《详细设计方案》《实施报告》《培训计划》《用户操作手册》《系统上线报告》《安装配置维护手册》《预验收/终验收报告》《运行/运维报告》），计划页不再单列"),
 ("“周期（Day）”整列空白", "全部补齐：现场实施38天（含并行），远程调试20天，试运行90天"),
 ("“负责人”整列空白", "按角色补齐：递缇项目经理/实施工程师、以弗德安装组、甲方设备商"),
 ("状态描述仅“待对接”", "统一状态字典并加下拉：未开始/进行中/已完成/受阻/待客户确认/待确认"),
 ("无责任边界", "新增“责任方”列：交换机与安装由甲方（以弗德），软件与远程调试由递缇"),
 ("无远程调试任务", "补为独立任务：1人1个月，与试运行并行"),
 ("无开工前置（签证/安全/差旅）", "在D1任务与备注中体现安全培训、作业许可、签证与差旅"),
 ("无15工位逐站签认", "新增“工位签认表”页，15站逐项签认并标注扫码枪状态"),
 ("标题日期2026-10-15", "已过期旧日期：改为“计划开始日”参数（默认2026-11-02），全表公式联动"),
 ("性能指标未给判定方法", "在输出物列明确≤1秒/≤4秒、并发、故障恢复、数据不丢失的测试记录要求"),
 ("“安装Windows Server 2016”“安装SQL Server”拆为两条独立任务", "国内已完成一轮调试：合并为“服务器系统与数据库复核”1天，保留镜像/配置基线恢复兜底；仅硬件变更时才重新安装"),
]
for i, (a, b) in enumerate(FIXES):
    r = 23 + i
    ws0.cell(r, 1, i + 1).alignment = Alignment(horizontal="center", vertical="center")
    ws0.cell(r, 2, a)
    ws0.cell(r, 3, b)
    for cc in range(1, 4):
        x = ws0.cell(r, cc)
        x.border = BORDER
        x.font = Font(name=FONT, size=10)
        x.alignment = Alignment(vertical="top", wrap_text=True)
    ws0.row_dimensions[r].height = 40

start = 23 + len(FIXES) + 1
ws0.cell(start, 1, "四、待客户确认事项（影响工期与验收）").font = Font(name=FONT, size=11, bold=True)
CONFIRM = [
 "5把扫码枪的补齐责任与时间：是否同意“借用＋串口枪＋手工录入”临时方案",
 "缺枪工位优先级：哪5个工位可暂缓启用扫码防错",
 "递缇人员赴墨西哥的签证与差旅费用归属（报价仅含远程调试1人1个月）",
 "5台伺服压机国外调试界面：合同不含，是否由甲方/设备商负责，递缇仅远程协助",
 "上线打码设备由谁提供、到场时间",
 "预验收是否已在杭州完成；一体机后到货是否需补做",
 "现场网络与VPN开放时间、服务器与交换机到场时间",
]
for i, t in enumerate(CONFIRM):
    r = start + 1 + i
    ws0.cell(r, 1, i + 1).alignment = Alignment(horizontal="center", vertical="center")
    ws0.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
    ws0.cell(r, 1, f"{i+1}. {t}")
    ws0.cell(r, 1).border = BORDER
    ws0.cell(r, 1).font = Font(name=FONT, size=10)
    ws0.cell(r, 1).alignment = Alignment(vertical="center", wrap_text=True, horizontal="left")
    ws0.row_dimensions[r].height = 22

PHASE_DESC = {
 "启动与前置条件确认": "清点、IT 环境确认",
 "现场硬件与网络环境搭建": "机房、布线、一体机、扫码枪、大屏",
 "MES软件部署与基础配置": "服务端/客户端、工艺、权限、报表",
 "工位单机与设备采集调试": "逐站验证、设备采集、性能",
 "全线联调与试生产": "PLC/SAP、空跑、试生产、培训",
 "试运行、验收与运维": "远程调试、试运行、终验收、质保",
}
lg = start + 9
ws0.cell(lg, 1, "五、甘特配色图例").font = Font(name=FONT, size=11, bold=True)
for i, (p, col) in enumerate(PHASE_COLOR.items()):
    r3 = lg + 1 + i
    c = ws0.cell(r3, 1)
    c.fill = PatternFill("solid", fgColor=col)
    c.border = BORDER
    ws0.cell(r3, 2, p).font = Font(name=FONT, size=10)
    ws0.cell(r3, 3, PHASE_DESC[p]).font = Font(name=FONT, size=9)
    for cc in range(1, 4):
        ws0.cell(r3, cc).border = BORDER
        ws0.cell(r3, cc).alignment = Alignment(vertical="center", wrap_text=True)
    ws0.row_dimensions[r3].height = 20

ws0.column_dimensions["A"].width = 26
ws0.column_dimensions["B"].width = 52
ws0.column_dimensions["C"].width = 78

# ---------- 计划 sheet ----------
HEAD = ["阶段", "序号", "工作内容", "周期\n(Day)", "计划开始", "计划完成", "负责人", "责任方", "状态描述", "前置条件 / 依赖", "备注（差异与风险）"]
WIDTH = [16, 6, 60, 8, 11, 11, 15, 17, 12, 34, 36]

ws.merge_cells("A1:K1")
ws["A1"] = "墨西哥CVT总成装配线MES系统 — 现场实施计划表"
ws["A1"].font = Font(name=FONT, size=15, bold=True, color="FFFFFFFF")
ws["A1"].fill = PatternFill("solid", fgColor=C_TITLE)
ws["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
ws.row_dimensions[1].height = 40

ws.merge_cells("A2:K2")
ws["A2"] = "项目：CVT总成装配线MES系统（墨西哥）   合同编号：CDT-ZB251022-EPHOD-CVT-MES-01   甲方：重庆以弗德智能装备有限公司   乙方：上海递缇智能系统有限公司"
ws["A2"].font = Font(name=FONT, size=9)
ws["A2"].fill = PatternFill("solid", fgColor=C_INFO)
ws["A2"].alignment = Alignment(horizontal="left", vertical="center")
ws.row_dimensions[2].height = 26

ws.merge_cells("A3:K3")
ws["A3"] = "版本 V2.1（现场实施版）   编制：上海递缇项目管理   计划开始日：见“说明与修订”页（B5，可一键联动）   现场实施：38天   远程调试：1人×1个月   试运行：3个月或1000台"
ws["A3"].font = Font(name=FONT, size=9, bold=True)
ws["A3"].fill = PatternFill("solid", fgColor=C_INFO)
ws["A3"].alignment = Alignment(horizontal="left", vertical="center")
ws.row_dimensions[3].height = 26

for i, h in enumerate(HEAD, 1):
    c = ws.cell(4, i, h)
    c.fill = PatternFill("solid", fgColor=C_HEAD)
    c.font = Font(name=FONT, size=10, bold=True)
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c.border = BORDER
ws.row_dimensions[4].height = 38
for i, w in enumerate(WIDTH, 1):
    ws.column_dimensions[get_column_letter(i)].width = w

r = 5
phase_ranges = []
ROW = {}
for phase, items in TASKS:
    ps = r
    for (no, work, out, days, off, owner, party, status, pre, note) in items:
        ROW[no] = r
        ws.cell(r, 2, no).alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(r, 3, work)
        ws.cell(r, 4, days).alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(r, 5, f"='说明与修订'!$B$5+{off}").number_format = "yyyy-mm-dd"
        ws.cell(r, 6, f"=E{r}+D{r}-1").number_format = "yyyy-mm-dd"
        ws.cell(r, 7, owner)
        ws.cell(r, 8, party)
        ws.cell(r, 9, status).alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.cell(r, 10, pre)
        ws.cell(r, 11, note)
        for c in range(2, 12):
            cell = ws.cell(r, c)
            cell.border = BORDER
            cell.font = Font(name=FONT, size=9, bold=(no in (4, 7, 31, 39, 42, 43)))
            cell.alignment = Alignment(vertical="center", wrap_text=True,
                                       horizontal=("center" if c in (2, 4, 8, 9) else "left"))
        ws.row_dimensions[r].height = max(
            30,
            min(84, max(math.ceil(len(str(work)) / 28), math.ceil(len(str(pre)) / 16),
                        math.ceil(len(str(note)) / 17)) * 15 + 8)
        )
        r += 1
    pe = r - 1
    phase_ranges.append((ps, pe, phase))
    if pe > ps:
        ws.merge_cells(start_row=ps, start_column=1, end_row=pe, end_column=1)
    c = ws.cell(ps, 1, phase)
    c.font = Font(name=FONT, size=10, bold=True)
    c.fill = PatternFill("solid", fgColor=C_PHASE)
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for rr in range(ps, pe + 1):
        ws.cell(rr, 1).border = BORDER
        ws.cell(rr, 1).fill = PatternFill("solid", fgColor=C_PHASE)

last = r - 1
# 斑马纹
for rr in range(5, last + 1):
    if (rr - 5) % 2 == 1:
        for c in range(2, 12):
            if ws.cell(rr, c).fill.fgColor.rgb in (None, "00000000"):
                ws.cell(rr, c).fill = PatternFill("solid", fgColor=C_ZEBRA)

# 状态色
STATUS_FILL = {
    "待客户确认": "FFFDE7E9",
    "待确认": "FFFFF4E5",
    "受阻": "FFFDE7E9",
    "已完成": "FFE6F4EA",
    "进行中": "FFE8F0FE",
}
for rr in range(5, last + 1):
    v = ws.cell(rr, 9).value
    if v in STATUS_FILL:
        ws.cell(rr, 9).fill = PatternFill("solid", fgColor=STATUS_FILL[v])

dv = DataValidation(type="list", formula1='"未开始,进行中,已完成,受阻,待客户确认,待确认"', allow_blank=True)
ws.add_data_validation(dv)
dv.add(f"I5:I{last}")

# 隐藏偏移列
for rr in range(5, last + 1):
    ws.cell(rr, 12).value = next(t[4] for ph, items in TASKS for t in items if t[0] == ws.cell(rr, 2).value)
ws.column_dimensions["L"].hidden = True
# 日期公式改为引用隐藏偏移列
for rr in range(5, last + 1):
    ws.cell(rr, 5).value = f"='说明与修订'!$B$5+L{rr}"

# 合计行
tr = last + 1
ws.merge_cells(start_row=tr, start_column=1, end_row=tr, end_column=4)
ws.cell(tr, 1, "合计：现场实施 38 天（含并行作业）；远程调试 20 天；试运行 90 天；质保 365 天")
ws.cell(tr, 1).font = Font(name=FONT, size=10, bold=True)
ws.cell(tr, 1).alignment = Alignment(horizontal="left", vertical="center")
for c in range(1, 12):
    ws.cell(tr, c).border = Border(left=thin, right=thin, bottom=thin, top=Side(style="medium", color="FFAAB4C5"))
    ws.cell(tr, c).fill = PatternFill("solid", fgColor="FFEAF2FF")
ws.row_dimensions[tr].height = 26

ws.freeze_panes = "C5"
ws.auto_filter.ref = f"A4:K{last}"

# 里程碑（依赖计划页实际行号）
for i, (name, tno, crit) in enumerate(MILESTONES):
    rr = MS_START + i
    ws0.cell(rr, 1, name)
    c = ws0.cell(rr, 2, f"=计划!F{ROW[tno]}")
    c.number_format = "yyyy-mm-dd"
    ws0.cell(rr, 3, crit)
    for cc in range(1, 4):
        x = ws0.cell(rr, cc)
        x.border = BORDER
        x.font = Font(name=FONT, size=10)
        x.alignment = Alignment(vertical="center", wrap_text=True)
ws0["B6"] = f"=MAX(计划!F5:F{ROW[44]})-MIN(计划!E5:E{ROW[44]})+1"

# ---------- 甘特 ----------
wg = wb.create_sheet("甘特（周）")
wg.sheet_view.showGridLines = False
WEEKS = 20
wg.merge_cells(start_row=1, start_column=1, end_row=1, end_column=4 + WEEKS)
wg["A1"] = "现场实施甘特（周）｜D1 = 计划开始日"
wg["A1"].font = Font(name=FONT, size=14, bold=True, color="FFFFFFFF")
wg["A1"].fill = PatternFill("solid", fgColor=C_TITLE)
wg["A1"].alignment = Alignment(horizontal="center", vertical="center")
wg.row_dimensions[1].height = 34

ghead = ["序号", "阶段", "工作内容", "周期"]
for i, h in enumerate(ghead, 1):
    c = wg.cell(2, i, h)
    c.fill = PatternFill("solid", fgColor=C_HEAD)
    c.font = Font(name=FONT, size=9, bold=True)
    c.border = BORDER
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
for k in range(WEEKS):
    c = wg.cell(2, 5 + k)
    c.value = f"=TEXT('说明与修订'!$B$5+7*{k},\"mm/dd\")"
    c.fill = PatternFill("solid", fgColor=C_HEAD)
    c.font = Font(name=FONT, size=8, bold=True)
    c.border = BORDER
    c.alignment = Alignment(horizontal="center", vertical="center", text_rotation=90)
wg.row_dimensions[2].height = 40
for i, w in enumerate([5, 15, 42, 6], 1):
    wg.column_dimensions[get_column_letter(i)].width = w
for k in range(WEEKS):
    wg.column_dimensions[get_column_letter(5 + k)].width = 4.2

gr = 3
for phase, items in TASKS:
    for (no, work, out, days, off, owner, party, status, pre, note) in items:
        if no == 48:
            continue
        wg.cell(gr, 1, no).alignment = Alignment(horizontal="center", vertical="center")
        wg.cell(gr, 2, phase)
        wg.cell(gr, 3, work)
        wg.cell(gr, 4, days).alignment = Alignment(horizontal="center", vertical="center")
        ws_ = off // 7 + 1
        we = (off + days - 1) // 7 + 1
        for k in range(1, WEEKS + 1):
            cell = wg.cell(gr, 4 + k)
            cell.border = BORDER
            if ws_ <= k <= min(we, WEEKS):
                cell.fill = PatternFill("solid", fgColor=PHASE_COLOR[phase])
            cell.font = Font(name=FONT, size=8)
        for c in range(1, 5):
            cell = wg.cell(gr, c)
            cell.border = BORDER
            cell.font = Font(name=FONT, size=8)
            cell.alignment = Alignment(vertical="center", wrap_text=(c == 3),
                                       horizontal=("center" if c in (1, 4) else "left"))
        wg.row_dimensions[gr].height = 26
        gr += 1

wg.cell(gr + 1, 1, "注：色块对应阶段，明细见“计划”页；质保期1年另计。").font = Font(name=FONT, size=8, italic=True)

# ---------- 工位签认表 ----------
ww = wb.create_sheet("工位签认表")
ww.sheet_view.showGridLines = False
wh = ["工位", "一体机到货", "支架安装", "上电", "网络", "客户端", "扫码枪状态", "ESOP", "扫码/手工录入", "签认人", "日期"]
ww.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(wh))
ww["A1"] = "15工位安装 / 上电 / 客户端 / 扫码 逐站签认表（★W11–W15 扫码枪缺口）"
ww["A1"].font = Font(name=FONT, size=13, bold=True, color="FFFFFFFF")
ww["A1"].fill = PatternFill("solid", fgColor=C_TITLE)
ww["A1"].alignment = Alignment(horizontal="center", vertical="center")
ww.row_dimensions[1].height = 32
for i, h in enumerate(wh, 1):
    c = ww.cell(2, i, h)
    c.fill = PatternFill("solid", fgColor=C_HEAD)
    c.font = Font(name=FONT, size=10, bold=True)
    c.border = BORDER
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
ww.row_dimensions[2].height = 34
for i, w in enumerate([8, 12, 10, 8, 8, 9, 20, 8, 15, 10, 11], 1):
    ww.column_dimensions[get_column_letter(i)].width = w
for i in range(15):
    r = 3 + i
    ww.cell(r, 1, f"W{i+1:02d}").alignment = Alignment(horizontal="center", vertical="center")
    for c in range(2, 12):
        ww.cell(r, c).border = BORDER
        ww.cell(r, c).font = Font(name=FONT, size=9)
        ww.cell(r, c).alignment = Alignment(horizontal="center", vertical="center")
    ww.cell(r, 2, "□")
    ww.cell(r, 3, "□")
    ww.cell(r, 4, "□")
    ww.cell(r, 5, "□")
    ww.cell(r, 6, "□")
    ww.cell(r, 7, "已到货" if i < 10 else "缺（借用/串口/补发）")
    ww.cell(r, 8, "□")
    ww.cell(r, 9, "扫码" if i < 10 else "手工录入")
    ww.cell(r, 10, "")
    ww.cell(r, 11, "")
    if i % 2 == 1:
        for c in range(1, 12):
            ww.cell(r, c).fill = PatternFill("solid", fgColor=C_ZEBRA)
    if i >= 10:
        ww.cell(r, 7).fill = PatternFill("solid", fgColor="FFFDE7E9")
        ww.cell(r, 9).fill = PatternFill("solid", fgColor="FFFFF4E5")
    ww.row_dimensions[r].height = 24
note_r = 19
ww.merge_cells(start_row=note_r, start_column=1, end_row=note_r, end_column=11)
ww.cell(note_r, 1, "说明：扫码枪实到10把；缺口5把按“借用其他工位/递缇随行带1把串口枪（RS232直连一体机COM口）/本地采购或国内补发”处置。缺枪工位先以手工录入序列号＋双人复核上线，扫码枪到位后仅改配置，无需二次开发。")
ww.cell(note_r, 1).font = Font(name=FONT, size=9, color="FFC00000")
ww.cell(note_r, 1).alignment = Alignment(vertical="center", wrap_text=True)
ww.row_dimensions[note_r].height = 42

# ---------- 硬件清单 ----------
wh2 = wb.create_sheet("硬件清单")
wh2.sheet_view.showGridLines = False
H = ["序号", "名称", "内容", "品牌", "单位", "数量", "提供方", "到货状态", "备注"]
wh2.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(H))
wh2["A1"] = "硬件与软件清单（含供货方、到货状态）"
wh2["A1"].font = Font(name=FONT, size=13, bold=True, color="FFFFFFFF")
wh2["A1"].fill = PatternFill("solid", fgColor=C_TITLE)
wh2["A1"].alignment = Alignment(horizontal="center", vertical="center")
wh2.row_dimensions[1].height = 32
for i, h in enumerate(H, 1):
    c = wh2.cell(2, i, h)
    c.fill = PatternFill("solid", fgColor=C_HEAD)
    c.font = Font(name=FONT, size=10, bold=True)
    c.border = BORDER
    c.alignment = Alignment(horizontal="center", vertical="center")
for i, w in enumerate([6, 16, 52, 12, 6, 7, 16, 20, 34], 1):
    wh2.column_dimensions[get_column_letter(i)].width = w

HW = [
 ("一", "机房硬件服务器设备", "", "", "", "", "", "", ""),
 (1, "塔式服务器", "联想（Lenovo）塔式服务器至强E2224G 32G内存丨256G固态+2*1TSATA硬盘", "Lenovo", "台", 1, "递缇", "待清点", ""),
 (2, "显示器及键鼠", "21寸", "三星", "台", 1, "递缇", "待清点", ""),
 (3, "数据库软件", "SQL Server 2012 R2/2016及以上", "微软", "套", 1, "递缇", "待清点", ""),
 (4, "操作系统", "Windows Server 2016", "微软", "套", 1, "递缇", "待清点", ""),
 (5, "服务器机柜", "42U服务器机柜，600*600*2000", "图腾", "台", 1, "递缇", "待清点", ""),
 (6, "网络交换机", "24口千兆以太网+4千兆光纤口网络", "华为", "台", 1, "重庆以弗德", "待清点", "需甲方开放配置权限"),
 (7, "UPS电源", "3KVA", "APC/SANTAK", "台", 1, "递缇", "待清点", ""),
 ("二", "现场硬件设备", "", "", "", "", "", "", ""),
 (1, "作业指导显示工位", "", "", "", "", "", "", ""),
 (1.1, "工位一体机", "PPC-3321QK-HR，J6412/8G内存/256G SSD/3*RS232/2*USB/Lan*2/无风扇/win10 Iot", "华北工控", "台", 15, "重庆以弗德", "已到货（晚到）", "★原未发运后补发，集中安装"),
 (2, "扫码枪", "DS8178-DP/USB口 有线/无线二维码扫描枪", "斑马", "台", 15, "重庆以弗德", "10把（缺5把）", "★缺口5把：借用/串口枪/补发"),
 (4, "工位显示大屏幕", "75寸电视+盒式PC", "小米", "台", 1, "重庆以弗德", "待清点", ""),
 (5, "一体机安装支架", "定制一体机安装支架", "", "台", 15, "重庆以弗德", "待清点", ""),
 (6, "设备安装辅材", "桥架，网线，电源线等", "", "套", 15, "重庆以弗德", "待清点", ""),
 ("三", "软件模块", "", "", "", "", "", "", ""),
 (1, "订单管理", "生产订单管理、计划排产、生产执行、生产报工管理", "", "套", 1, "递缇", "—", ""),
 (2, "上线打印功能", "标签打印，条码规则管理", "", "套", 1, "递缇", "—", "打码硬件不在供货清单"),
 (3, "零件条码采集", "零件条码识别，零件条码绑定", "", "套", 1, "递缇", "—", ""),
 (4, "过程数据管理", "产品进站数据，出站数据采集", "", "套", 1, "递缇", "—", ""),
 (5, "工艺管理", "工艺信息配置、操作工序管理", "", "套", 1, "递缇", "—", ""),
 (6, "质量报告管理", "生产质量数据报告", "", "套", 1, "递缇", "—", ""),
 (7, "电子看板", "产量统计，质量看板，合格率", "", "套", 1, "递缇", "—", ""),
 (8, "统计分析模块", "合格率统计，生产数据统计报表", "", "套", 1, "递缇", "—", ""),
 (9, "设备数据采集模块", "", "", "套", 1, "递缇", "—", ""),
 (9.1, "与伺服压机通讯接口", "伺服压机数据采集", "", "套", 5, "递缇", "—", "★5台，CVT二线；国内调试，不含国外调试费用"),
 (9.2, "拧紧控制", "1套拧紧工具的数据采集", "", "套", 1, "递缇", "—", "ATLAS拧紧枪（主动轮）"),
 (9.3, "电子操作指导", "工艺电子操作指导", "", "套", 1, "递缇", "—", ""),
 (9.4, "与检测设备通讯接口", "检测设备数据采集", "", "套", 2, "递缇", "—", ""),
]
hr = 3
for row in HW:
    for i, v in enumerate(row, 1):
        c = wh2.cell(hr, i, v)
        c.border = BORDER
        c.font = Font(name=FONT, size=9, bold=isinstance(row[0], str))
        c.alignment = Alignment(vertical="center", wrap_text=True,
                                horizontal=("center" if i in (1, 5, 6, 7, 8) else "left"))
    if isinstance(row[0], str):
        for i in range(1, len(H) + 1):
            wh2.cell(hr, i).fill = PatternFill("solid", fgColor=C_PHASE)
    elif row[7] == "10把（缺5把）":
        wh2.cell(hr, 8).fill = PatternFill("solid", fgColor="FFFDE7E9")
    elif row[7] == "已到货（晚到）":
        wh2.cell(hr, 8).fill = PatternFill("solid", fgColor="FFFFF4E5")
    wh2.row_dimensions[hr].height = 36
    hr += 1

wb._sheets = [wb[n] for n in ["计划", "甘特（周）", "工位签认表", "硬件清单", "说明与修订"]]

for sh in wb.worksheets:
    sh.page_setup.orientation = "landscape"
    sh.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    sh.page_setup.fitToWidth = 1
    sh.page_setup.fitToHeight = 0
    sh.print_options.horizontalCentered = True
    sh.page_margins.left = 0.3
    sh.page_margins.right = 0.3
    sh.page_margins.top = 0.4
    sh.page_margins.bottom = 0.4
ws.print_title_rows = "1:4"
wg.print_title_rows = "1:2"
ww.print_title_rows = "1:2"
wh2.print_title_rows = "1:2"
for sh in (ww, wh2):
    sh.page_setup.fitToHeight = 1

wb.save("/workspace/CVT-MES项目资料/04_墨西哥CVT线现场实施计划_V2.xlsx")
print("saved, last row", last)
