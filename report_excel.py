"""中文命令行入口：校验通过后导出 Excel，不自动覆盖或安装依赖。"""

import argparse
from pathlib import Path

from csv_reader import CsvValidationError


def main(argv=None):
    parser = argparse.ArgumentParser(description="严格校验 CSV 后导出三工作表 Excel 报表。", allow_abbrev=False)
    parser.add_argument("csv_file", nargs="?", type=Path,
                        default=Path(__file__).parent / "data" / "sample_orders.csv",
                        help="CSV 路径，默认使用项目的虚构模拟订单")
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
    print(f"Excel 报表已生成：{destination}")
    print("包含运营总览、商品销量排行、数据校验摘要；原始 CSV 未修改。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
