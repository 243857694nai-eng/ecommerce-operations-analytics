"""Windows Excel 报表：复用完整校验和指标，先核验内容，再无覆盖发布。"""

import io
import os
import tempfile
from datetime import datetime
from decimal import Decimal
from pathlib import Path
from uuid import uuid4
from xml.etree.ElementTree import tostring

from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.axis import DateAxis
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.styles import Alignment, Font, PatternFill

from metrics import MetricsReport, calculate_metrics, format_money
from order_validator import BEIJING_TIME


SHEET_NAMES = ("运营总览", "商品销量排行", "数据校验摘要")
VISUAL_SHEET_NAMES = ("销售趋势", "商品销售分析", "退款分析")
MONEY_FORMAT = '#,##0.00'
COUNT_FORMAT = '#,##0'
DEFAULT_OUTPUT_DIR = Path(__file__).parent / "output"


class ExcelReportError(Exception):
    """导出未完成；不能据此认定已经生成正式报表。"""


def _text(value: str) -> str:
    # Excel 的文本长度有限；openpyxl 默认会截断，必须先拒绝，避免静默丢字。
    if len(value.encode("utf-16-le")) // 2 > 32767:
        raise ExcelReportError("文本超过 Excel 单元格长度限制，拒绝截断导出。")
    return value


def _count(value: int) -> int:
    if len(str(abs(value))) > 15:
        raise ExcelReportError("件数或订单数超过 Excel 的安全数值精度，拒绝导出。")
    return value


def _money(cents) -> Decimal | str:
    if cents is None:
        return format_money(None)
    # 只调用已有展示舍入，不重新计算客单价或任何业务指标。
    value = Decimal(format_money(cents))
    if len(value.as_tuple().digits) > 15:
        raise ExcelReportError("金额（含两位小数）超过 Excel 的安全数值精度，拒绝导出。")
    return value


def _append(sheet, values):
    values = tuple(_text(value) if isinstance(value, str) else value for value in values)
    sheet.append(values)
    for cell in sheet[sheet.max_row]:
        if isinstance(cell.value, str):
            # '='、错误码、URL 等一律保存为普通文本；不生成公式或自动链接。
            cell.data_type = "s"
            cell.hyperlink = None


def _style(sheet, widths):
    sheet.freeze_panes = "A2"
    sheet.sheet_view.showGridLines = False
    sheet.row_dimensions[1].height = 30
    for cell in sheet[1]:
        cell.font = Font(name="微软雅黑", bold=True, color="FFFFFF", size=11)
        cell.fill = PatternFill("solid", fgColor="244A64")
        cell.alignment = Alignment(vertical="center", wrap_text=True)
    for column, width in widths.items():
        sheet.column_dimensions[column].width = width
    for row in sheet.iter_rows(min_row=2):
        sheet.row_dimensions[row[0].row].height = 42
        for cell in row:
            cell.font = Font(name="微软雅黑", size=11)
            cell.alignment = Alignment(vertical="center", wrap_text=True,
                                       horizontal="right" if cell.data_type == "n" else "left")
            if cell.row % 2 == 0:
                cell.fill = PatternFill("solid", fgColor="EFF5F8")
    sheet.print_options.horizontalCentered = True
    sheet.sheet_properties.pageSetUpPr.fitToPage = True
    sheet.page_setup.orientation = "landscape"
    sheet.page_setup.paperSize = sheet.PAPERSIZE_A4
    sheet.page_setup.fitToWidth = 1
    sheet.page_setup.fitToHeight = 0
    sheet.print_title_rows = "1:1"


def _build_workbook(report: MetricsReport, source_name: str, generated_at: datetime) -> Workbook:
    workbook = Workbook()
    overview = workbook.active
    overview.title = SHEET_NAMES[0]
    _append(overview, ("指标", "数值", "单位", "统计口径"))
    rows = (
        ("商品明细数", _count(report.line_count), "条", "一行是一条商品明细，不等于一笔订单"),
        ("唯一订单数", _count(report.total_order_count), "笔", "包含未支付和取消订单；按订单编号去重"),
        ("支付金额", _money(report.payment_cents), "元", "包含后来全额、部分退款订单的历史支付"),
        ("退款金额", _money(report.refund_cents), "元", "已完成的退款金额"),
        ("退款后净销售额", _money(report.net_sales_cents), "元", "支付金额 − 退款金额；不等于利润"),
        ("支付订单数", _count(report.paid_order_count), "笔", "包含后来全额、部分退款的唯一已支付订单"),
        ("净有效订单数", _count(report.net_order_count), "笔", "已支付订单中，整单净金额大于零"),
        ("支付客单价", _money(report.paid_average_cents), "元/笔", "支付金额 ÷ 支付订单数；仅展示时四舍五入"),
        ("净客单价", _money(report.net_average_cents), "元/笔", "净销售额 ÷ 净有效订单数；仅展示时四舍五入"),
    )
    for values in rows:
        _append(overview, values)
        overview.cell(overview.max_row, 2).number_format = MONEY_FORMAT if values[2].startswith("元") else COUNT_FORMAT
    _style(overview, {"A": 24, "B": 23, "C": 12, "D": 64})

    ranking = workbook.create_sheet(SHEET_NAMES[1])
    _append(ranking, ("排名", "商品编号", "商品名称", "商品类别", "净成交件数"))
    for rank, product in enumerate(report.products, 1):
        _append(ranking, (_count(rank), _text(product.product_id), _text(product.product_name),
                          _text(product.category), _count(product.net_quantity)))
        ranking.cell(ranking.max_row, 1).number_format = COUNT_FORMAT
        ranking.cell(ranking.max_row, 5).number_format = COUNT_FORMAT
    ranking.auto_filter.ref = f"A1:E{ranking.max_row}"
    _style(ranking, {"A": 10, "B": 18, "C": 34, "D": 20, "E": 18})

    summary = workbook.create_sheet(SHEET_NAMES[2])
    _append(summary, ("项目", "内容"))
    source = report.source
    is_excel = source is not None and source.file_format == "XLSX"
    for values in (
        ("输入 Excel 文件名" if is_excel else "输入 CSV 文件名", _text(source.file_name if source else source_name)),
        ("校验结果", "整份数据通过既有严格校验；未丢弃异常行或自动补值"),
        ("商品明细数", _count(report.line_count)),
        ("唯一订单数", _count(report.total_order_count)),
        ("金额口径", "人民币元；历史支付包含全额退款订单；净销售额 = 支付 − 已完成退款"),
        ("订单与客单价口径", "支付与净额口径分别使用对应订单数；零分母显示无法计算"),
        ("商品排行口径", "购买件数 − 明确退款件数；只退差价不扣件数；并列按商品编号字符串升序"),
        ("数据时间边界", "订单状态快照；缺少支付和退款发生时间，不是每日资金流水"),
        ("报表生成时间（北京时间 UTC+08:00）", generated_at),
        ("使用提示", "请使用虚构模拟订单练习。修改 Excel 后需重新读取并确认映射，本报表不会自动更新" if is_excel else
         "本项目附带 CSV 为虚构模拟数据。修改 CSV 后需重新运行程序，本报表不会自动更新"),
    ):
        _append(summary, values)
    # 追加来源行，不移动原有指标、生成时间及其格式位置，保留 CSV 报表兼容性。
    _append(summary, ("输入文件格式", source.file_format if source else "CSV"))
    if is_excel:
        for values in (("输入工作表", source.sheet_name), ("表头原始行号", source.header_row),
                       ("输入快照 SHA256", source.sha256),
                       ("快照说明", "结果对应映射确认时的内存快照；源文件后来变化需重新运行并确认")):
            _append(summary, values)
            if values[0] == "表头原始行号":
                summary.cell(summary.max_row, 2).number_format = COUNT_FORMAT
    summary["B4"].number_format = COUNT_FORMAT
    summary["B5"].number_format = COUNT_FORMAT
    summary["B10"].number_format = "yyyy-mm-dd hh:mm:ss"
    _style(summary, {"A": 44, "B": 100})
    # 文字较长的两列保持足够高度；不靠缩小字体挤入单元格。
    for index in range(6, 10):
        summary.row_dimensions[index].height = 54
    return workbook


def _analysis_style(sheet, widths, header_row, note):
    """原有三张表不移动；新分析表用独立标题、口径提示及数据表头。"""
    _style(sheet, widths)
    last = max(len(row) for row in sheet.iter_rows())
    sheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=last)
    sheet.merge_cells(start_row=2, start_column=1, end_row=2, end_column=last)
    sheet.merge_cells(start_row=3, start_column=1, end_row=3, end_column=last)
    sheet.row_dimensions[1].height = 32
    sheet.row_dimensions[3].height = 36
    sheet.cell(2, 1, note)
    sheet.row_dimensions[2].height = 64
    sheet.cell(2, 1).alignment = Alignment(wrap_text=True, vertical="center")
    sheet.cell(2, 1).font = Font(name="微软雅黑", size=11, color="8A4400")
    for cell in sheet[header_row]:
        cell.font = Font(name="微软雅黑", bold=True, color="FFFFFF", size=11)
        cell.fill = PatternFill("solid", fgColor="244A64")
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    sheet.row_dimensions[header_row].height = 42
    sheet.freeze_panes = f"A{header_row + 1}"
    sheet.print_title_rows = f"1:{header_row}"
    sheet.sheet_properties.tabColor = "244A64"
    sheet.sheet_view.zoomScale = 90


def _line_chart(sheet, title, header_row, end_row, columns, anchor):
    if end_row <= header_row:
        return
    chart = LineChart()
    chart.title = title
    chart.title.overlay = False
    chart.style = 13
    chart.width, chart.height = 21, 12
    chart.varyColors = False  # 单条退款曲线也只保留一个系列图例，不按每个日期生成图例。
    # 单位和日期归组写入图表标题；省去部分 Excel 版本会与刻度重叠的轴标题。
    chart.y_axis.numFmt = MONEY_FORMAT
    chart.y_axis.crossAx = 500
    chart.x_axis = DateAxis(crossAx=100)
    chart.x_axis.axPos = "b"
    chart.y_axis.axPos = "l"
    chart.x_axis.number_format = "yyyy-mm-dd"
    chart.x_axis.majorTimeUnit = "days"
    chart.legend.position = "b"
    chart.legend.overlay = False
    for axis in (chart.x_axis, chart.y_axis):
        axis.delete = False
        axis.tickLblPos = "nextTo"
    chart.y_axis.majorGridlines.spPr = GraphicalProperties()
    chart.y_axis.majorGridlines.spPr.line.solidFill = "E2E8F0"
    for column in columns:
        chart.add_data(Reference(sheet, min_col=column, min_row=header_row, max_row=end_row),
                       titles_from_data=True)
    chart.set_categories(Reference(sheet, min_col=1, min_row=header_row + 1, max_row=end_row))
    for series, color in zip(chart.series, ("2878B5", "E58B2A", "2B8A6E")):
        series.graphicalProperties.line.solidFill = color
        series.graphicalProperties.line.width = 26000
        series.marker.symbol = "circle"  # 单日数据仍能看到实际数据点。
        series.marker.size = 5
        series.marker.graphicalProperties.solidFill = color
        series.marker.graphicalProperties.line.solidFill = color
        series.smooth = False
    sheet.add_chart(chart, anchor)


def _display_ratio(value):
    # 精确 Fraction 保留在模型中；这里只生成展示用、四舍五入到 0.01% 的数值。
    return format_money(None) if value is None else Decimal(format_money(value * 10000)) / 100


def _add_visual_analysis(workbook: Workbook, report: MetricsReport):
    trend = workbook.create_sheet(VISUAL_SHEET_NAMES[0])
    _append(trend, ("销售趋势｜模拟数据",))
    _append(trend, (None,))
    _append(trend, ("仅展示源数据中出现的下单日期；缺失日期不补零，折线只连接已观测点。",))
    _append(trend, ("下单日期", "支付金额（元）", "退款金额（元）", "净成交金额（元）"))
    for day in report.daily:
        _append(trend, (datetime(day.day.year, day.day.month, day.day.day), _money(day.payment_cents),
                        _money(day.refund_cents), _money(day.net_sales_cents)))
        trend.cell(trend.max_row, 1).number_format = "yyyy-mm-dd"
        for column in range(2, 5):
            trend.cell(trend.max_row, column).number_format = MONEY_FORMAT
    _analysis_style(trend, {"A": 28, "B": 22, "C": 22, "D": 24}, 4,
                    "虚构模拟订单。按下单日期归组的支付、累计退款及净额快照；缺少支付/退款发生时间，不能解释为当日资金流水。净成交金额沿用净销售额口径，不等于利润。")
    _line_chart(trend, "销售快照（按下单日期归组，元｜模拟数据）", 4, trend.max_row, (2, 3, 4), "F4")

    products = workbook.create_sheet(VISUAL_SHEET_NAMES[1])
    _append(products, ("商品销售分析｜模拟数据",))
    _append(products, (None,))
    _append(products, ("完整排行保留所有已支付商品；图表最多展示净销量前10。",))
    _append(products, ("净销量排名", "商品编号", "商品名称", "类别", "支付件数", "退款件数", "净成交件数",
                       "支付金额（元）", "退款金额（元）", "净销售金额（元）", "图表标签（商品 / 编号）"))
    for rank, product in enumerate(report.products, 1):
        _append(products, (_count(rank), product.product_id, product.product_name, product.category,
                           _count(product.paid_quantity), _count(product.refund_quantity), _count(product.net_quantity),
                           _money(product.payment_cents), _money(product.refund_cents), _money(product.net_sales_cents),
                           f"{product.product_name} / {product.product_id}"))
        for column in (1, 5, 6, 7):
            products.cell(products.max_row, column).number_format = COUNT_FORMAT
        for column in (8, 9, 10):
            products.cell(products.max_row, column).number_format = MONEY_FORMAT
    end = products.max_row
    _analysis_style(products, {"A": 14, "B": 16, "C": 28, "D": 18, "E": 14, "F": 14, "G": 16,
                               "H": 20, "I": 20, "J": 22, "K": 40}, 4,
                    "虚构模拟订单。销量=支付件数−明确退款件数；金额同时列出支付、退款和净额，排名沿用净销量降序、商品编号字符串升序。同名不同编号不合并，金额不代表利润。")
    products.auto_filter.ref = f"A4:K{end}"
    if report.products:
        chart = BarChart()
        chart.type = "bar"  # 中文商品名较长，横向条形图比竖向柱形图更易读。
        chart.title = "净成交件数前10（件｜模拟数据）"
        chart.title.overlay = False
        chart.style = 10
        chart.width, chart.height = 26, 15
        chart.y_axis.numFmt = COUNT_FORMAT
        chart.x_axis.scaling.orientation = "maxMin"
        chart.x_axis.axPos = "l"
        chart.y_axis.axPos = "b"
        for axis in (chart.x_axis, chart.y_axis):
            axis.delete = False
            axis.tickLblPos = "nextTo"
        chart.y_axis.majorGridlines.spPr = GraphicalProperties()
        chart.y_axis.majorGridlines.spPr.line.solidFill = "E2E8F0"
        chart.legend = None
        chart.add_data(Reference(products, min_col=7, min_row=4, max_row=min(end, 14)), titles_from_data=True)
        chart.set_categories(Reference(products, min_col=11, min_row=5, max_row=min(end, 14)))
        products.add_chart(chart, "M4")
    else:
        _append(products, ("没有已支付商品，不生成空排行图。",))

    refunds = workbook.create_sheet(VISUAL_SHEET_NAMES[2])
    _append(refunds, ("退款分析｜模拟数据",))
    _append(refunds, (None,))
    _append(refunds, ("金额率和订单率分母不同；都不是退货件数率，也不是申请退款率。",))
    _append(refunds, ("指标", "数值", "口径 / 单位", "说明"))
    for values in (
        ("退款总金额", _money(report.refund_cents), "元", "累计已完成退款，不含申请中金额"),
        ("支付订单数", _count(report.paid_order_count), "笔", "包括之后全额、部分退款的唯一订单"),
        ("退款订单数", _count(report.refund_order_count), "笔", "整单已发生退款，一单多商品仅算一笔"),
        ("全额退款订单数", _count(report.full_refund_order_count), "笔", "按原有已全额退款状态去重"),
        ("部分退款订单数", _count(report.partial_refund_order_count), "笔", "按原有部分退款状态去重；包括只退差价"),
        ("退款金额率", _display_ratio(report.refund_amount_ratio), "退款金额 ÷ 支付金额", "分母为零时无法计算；展示舍入到0.01%"),
        ("退款订单率", _display_ratio(report.refund_order_ratio), "退款订单数 ÷ 支付订单数", "分母为零时无法计算；展示舍入到0.01%"),
    ):
        _append(refunds, values)
        if isinstance(values[1], Decimal):
            refunds.cell(refunds.max_row, 2).number_format = "0.00%" if "率" in values[0] else MONEY_FORMAT
        elif isinstance(values[1], int):
            refunds.cell(refunds.max_row, 2).number_format = COUNT_FORMAT
    _append(refunds, (None,))
    _append(refunds, ("退款趋势按下单日期归组，不能判断哪一天实际发生了退款。",))
    _append(refunds, ("缺失日期不补零；同一订单的全部商品行属于同一下单日期。",))
    _append(refunds, ("下单日期", "退款金额（元）", "退款订单数（笔）", "支付订单数（笔）"))
    for day in report.daily:
        _append(refunds, (datetime(day.day.year, day.day.month, day.day.day), _money(day.refund_cents),
                          _count(day.refund_order_count), _count(day.paid_order_count)))
        refunds.cell(refunds.max_row, 1).number_format = "yyyy-mm-dd"
        refunds.cell(refunds.max_row, 2).number_format = MONEY_FORMAT
    _analysis_style(refunds, {"A": 28, "B": 22, "C": 32, "D": 46}, 4,
                    "虚构模拟订单。退款按订单当前状态及明细累计已完成金额统计；没有退款发生日期，图表展示的是下单日期归组的退款快照，不是每日实际退款流水。")
    for cell in refunds[15]:
        cell.font = Font(name="微软雅黑", bold=True, color="FFFFFF", size=11)
        cell.fill = PatternFill("solid", fgColor="244A64")
    for row in (13, 14):
        refunds.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
    _line_chart(refunds, "退款快照（按下单日期归组，元｜模拟数据）", 15, refunds.max_row, (2,), "F15")
    _append(workbook["数据校验摘要"], ("可视化说明", "5C 模拟数据分析；按下单日期归组的快照；退款率分母见退款分析；图表和数据均为静态结果"))


def _serialize_and_verify(workbook: Workbook) -> bytes:
    """先在内存完成 XLSX，并回读检查每个值和关键格式，不产生正式文件。"""
    with io.BytesIO() as buffer:
        workbook.save(buffer)
        buffer.seek(0)
        restored = load_workbook(buffer, data_only=False)
        try:
            if (workbook.sheetnames not in (list(SHEET_NAMES), list(SHEET_NAMES + VISUAL_SHEET_NAMES)) or
                    restored.sheetnames != workbook.sheetnames):
                raise ExcelReportError("生成的工作表清单异常。")
            for original, saved in zip(workbook.worksheets, restored.worksheets):
                # 图表引用只能指向本次工作簿数据；回读验证其序列、标题、轴及样式没有丢失或改变。
                if [tostring(chart.to_tree()) for chart in original._charts] != [
                        tostring(chart.to_tree()) for chart in saved._charts]:
                    raise ExcelReportError("图表内容或数据引用回读不一致。")
                if (original.max_row, original.max_column, original.freeze_panes, original.auto_filter.ref) != (
                        saved.max_row, saved.max_column, saved.freeze_panes, saved.auto_filter.ref):
                    raise ExcelReportError("工作表结构或表头设置回读不一致。")
                for row in original:
                    for expected in row:
                        actual = saved[expected.coordinate]
                        if actual.data_type == "f" or actual.hyperlink is not None:
                            raise ExcelReportError("发现非预期公式或链接，拒绝导出。")
                        if isinstance(expected.value, Decimal):
                            equal = actual.data_type == "n" and Decimal(str(actual.value)) == expected.value
                        else:
                            equal = actual.value == expected.value and actual.data_type == expected.data_type
                        if not equal or actual.number_format != expected.number_format:
                            raise ExcelReportError(f"{saved.title}!{actual.coordinate} 的内容或数值精度回读不一致。")
        finally:
            restored.close()
        return buffer.getvalue()


def _write_payload(stream, payload: bytes):
    if stream.write(payload) != len(payload):
        raise OSError("报表临时文件未完整写入。")
    stream.flush()
    os.fsync(stream.fileno())


def _publish_no_replace(temporary: Path, destination: Path):
    # Windows 的 rename 在目标存在时拒绝执行；不能用会覆盖目标的 replace。
    if os.name != "nt":
        raise ExcelReportError("当前无覆盖发布流程仅支持 Windows；其他系统尚未实现。")
    os.rename(temporary, destination)


def export_excel(csv_path: str | Path, output_path: str | Path | None = None, *, visual: bool = False) -> Path:
    """完整校验 → 复用指标 → 内存回读 → 临时文件 → Windows 无覆盖发布。"""
    report = calculate_metrics(csv_path)  # 校验异常直接向上传递；此前不创建输出目录。
    generated_at = datetime.now(BEIJING_TIME).replace(tzinfo=None, microsecond=0)
    destination = Path(output_path) if output_path is not None else (
        DEFAULT_OUTPUT_DIR / f"运营报表_{generated_at:%Y%m%d_%H%M%S}_{uuid4().hex[:8]}.xlsx"
    )
    temporary = None
    workbook = None
    try:
        # 在创建文件前确定路径，正式发布后不再做可能失败的路径解析。
        destination = destination.absolute()
        if destination.suffix.lower() != ".xlsx":
            raise ExcelReportError("输出文件必须使用 .xlsx 扩展名。")
        if os.path.lexists(destination):
            raise ExcelReportError(f"目标已经存在，拒绝覆盖：{destination}")
        workbook = _build_workbook(report, Path(csv_path).name, generated_at)
        if visual:
            _add_visual_analysis(workbook, report)
        payload = _serialize_and_verify(workbook)
        workbook.close()
        workbook = None
        destination.parent.mkdir(parents=True, exist_ok=True)
        with tempfile.NamedTemporaryFile(mode="wb", prefix=".excel_report_", suffix=".tmp",
                                         dir=destination.parent, delete=False) as stream:
            temporary = Path(stream.name)
            _write_payload(stream, payload)
        _publish_no_replace(temporary, destination)
        temporary = None
        return destination
    except ExcelReportError:
        raise
    except Exception as error:
        # 只在导出边界包装异常，不返回成功；原始原因保留在异常链中。
        raise ExcelReportError(f"导出失败，未发布本次正式报表：{error}") from error
    finally:
        if workbook is not None:
            workbook.close()
        if temporary is not None:
            try:
                temporary.unlink(missing_ok=True)
            except OSError as error:
                raise ExcelReportError(f"导出未完成，临时文件无法清理，请检查：{temporary}") from error
