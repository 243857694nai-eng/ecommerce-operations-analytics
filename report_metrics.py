"""中文只读指标报告入口；不保存文件、不修改 CSV 或 Excel。"""

import argparse
from pathlib import Path

from csv_reader import CsvValidationError
from metrics import MetricsReport, calculate_metrics, format_money
from order_loader import InputCancelled, InputValidationError


def render_report(report: MetricsReport) -> str:
    lines = [
        "运营指标报告（当前 V1 口径；订单状态快照，非每日资金流水）",
        f"商品明细：{report.line_count} 条",
        f"唯一订单：{report.total_order_count} 笔（包含未支付和取消订单）",
        f"支付金额：{format_money(report.payment_cents)} 元",
        f"已完成退款金额：{format_money(report.refund_cents)} 元",
        f"净销售额：{format_money(report.net_sales_cents)} 元",
        f"支付订单数：{report.paid_order_count} 笔（包含全额退款订单）",
        f"净有效订单数：{report.net_order_count} 笔（整单净金额大于零）",
    ]
    for label, average, formula in (
        ("支付客单价", report.paid_average_cents, "支付金额 ÷ 支付订单数"),
        ("净客单价", report.net_average_cents, "净销售额 ÷ 净有效订单数"),
    ):
        unit = " 元/笔" if average is not None else ""
        lines.append(f"{label}：{format_money(average)}{unit}（{formula}）")
    lines.extend([
        "",
        "商品销量排行（净成交件数降序，并列按商品编号字符串升序）",
        "商品编号 | 商品名称 | 支付件数 | 退款件数 | 净成交件数 | 净销售金额（元）",
    ])
    for product in report.products:
        lines.append(
            f"{product.product_id} | {product.product_name} | {product.paid_quantity} | "
            f"{product.refund_quantity} | {product.net_quantity} | {format_money(product.net_sales_cents)}"
        )
    if not report.products:
        lines.append("暂无已支付商品。")
    lines.extend([
        f"商品净成交件数合计：{report.net_quantity} 件",
        "净销售额不等于利润；退款件数使用明确记录，不根据退款金额推算。",
        "计算只读入 Excel 内存快照；未修改原始文件。" if report.source and report.source.file_format == "XLSX" else
        "计算只读入 CSV；未修改原始文件。本项目样本全部为虚构模拟数据。",
    ])
    if report.source and report.source.file_format == "XLSX":
        lines.extend([f"输入工作表：{report.source.sheet_name!r}；表头原始行号：{report.source.header_row}",
                      f"输入快照 SHA256：{report.source.sha256}"])
    return "\n".join(lines)


def main(argv=None):
    parser = argparse.ArgumentParser(description="先严格校验，再只读计算 V1 运营指标。", allow_abbrev=False)
    parser.add_argument("csv_file", nargs="?", type=Path,
                        default=Path(__file__).parent / "data" / "sample_orders.csv",
                        help="CSV 或 .xlsx 路径；Excel 需交互确认，省略时仍使用原模拟 CSV")
    args = parser.parse_args(argv)
    try:
        report = calculate_metrics(args.csv_file)
        text = render_report(report)
    except CsvValidationError as error:
        print(f"校验未通过，共发现 {len(error.issues)} 个问题：")
        for issue in error.issues:
            print(f"- {issue}")
        print("拒绝输出正式指标或部分统计；未修改原始 CSV。")
        return 1
    except InputValidationError as error:
        print(f"Excel 校验未通过：{error}")
        print("拒绝输出正式指标或部分统计；未修改原始 Excel。")
        return 1
    except InputCancelled:
        print("已取消，未计算或输出正式指标。")
        return 2
    except ModuleNotFoundError as error:
        if error.name not in {"openpyxl", "et_xmlfile"}:
            raise
        print("缺少 Excel 依赖，请用本项目 .venv\\Scripts\\python.exe 运行；程序不会自动安装软件。")
        return 2
    except Exception as error:
        print(f"计算未完成（{type(error).__name__}），已停止，未输出正式指标。")
        return 1
    print(text)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
