"""中文命令行入口：只读取和校验，不输出正式运营指标。"""

import argparse
from pathlib import Path

from csv_reader import CsvValidationError
from order_validator import load_and_validate


def main(argv=None):
    parser = argparse.ArgumentParser(description="只读检查 V1 订单 CSV，不修改原始文件。", allow_abbrev=False)
    parser.add_argument("csv_file", nargs="?", type=Path,
                        default=Path(__file__).parent / "data" / "sample_orders.csv",
                        help="CSV 路径；省略时检查项目的模拟订单文件")
    args = parser.parse_args(argv)
    try:
        records = load_and_validate(args.csv_file)
    except CsvValidationError as error:
        print(f"校验未通过，共发现 {len(error.issues)} 个问题：")
        for issue in error.issues:
            print(f"- {issue}")
        print("未返回部分数据，未修改 CSV，也未填补或删除任何记录。")
        return 1
    print(f"校验通过：读取 {len(records)} 条商品明细。")
    print("原始 CSV 未修改；当前阶段不计算运营指标。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
