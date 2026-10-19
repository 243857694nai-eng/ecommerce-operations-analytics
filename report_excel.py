"""中文命令行入口：校验通过后导出 Excel，不自动覆盖或安装依赖。"""

import argparse
from pathlib import Path

from csv_reader import CsvValidationError
from order_loader import InputCancelled, InputValidationError


def main(argv=None):
    parser = argparse.ArgumentParser(description="严格校验 CSV 或经确认映射的 XLSX 后导出三工作表报表。", allow_abbrev=False)
    parser.add_argument("csv_file", nargs="?", type=Path,
                        default=Path(__file__).parent / "data" / "sample_orders.csv",
                        help="CSV 或 .xlsx 路径；Excel 需交互确认，默认仍使用原模拟 CSV")
    parser.add_argument("--output", type=Path, help="新建 .xlsx 的路径；默认保存到项目 output 目录；拒绝覆盖")
    args = parser.parse_args(argv)
    try:
        from excel_report import ExcelReportError, export_excel
    except ModuleNotFoundError as error:
        if error.name not in {"openpyxl", "et_xmlfile"}:
            raise
        print("缺少 Excel 导出依赖。请使用本项目 .venv\\Scripts\\python.exe 运行；程序不会自动安装软件。")
        return 2
    try:
        destination = export_excel(args.csv_file, args.output)
    except CsvValidationError as error:
        print(f"校验未通过，共发现 {len(error.issues)} 个问题：")
        for issue in error.issues:
            print(f"- {issue}")
        print("未生成 Excel 报表，也未返回部分统计；原始 CSV 未修改。")
        return 1
    except ExcelReportError as error:
        print(f"Excel 导出未完成：{error}")
        return 1
    except InputValidationError as error:
        print(f"Excel 校验未通过：{error}")
        print("未生成 Excel 报表，也未返回部分统计；原始 Excel 未修改。")
        return 1
    except InputCancelled:
        print("已取消，未生成本次 Excel 报表。")
        return 2
    except (EOFError, KeyboardInterrupt):
        print("操作已中断，未报告成功；若已进入保存阶段，请检查目标文件。")
        return 2
    except Exception as error:
        print(f"操作未完成（{type(error).__name__}），已停止，未报告导出成功。")
        return 1
    print(f"Excel 报表已生成：{destination}")
    if args.csv_file.suffix.lower() == ".xlsx":
        print("包含运营总览、商品销量排行、数据校验摘要；原始 Excel 未修改。")
    else:
        print("包含运营总览、商品销量排行、数据校验摘要；原始 CSV 未修改。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
