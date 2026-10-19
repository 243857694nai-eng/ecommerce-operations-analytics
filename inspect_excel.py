"""5A 中文交互入口：明确选择 → 映射预览 → 确认 → 只校验，不生成报表。"""

import argparse
from pathlib import Path

from csv_reader import REQUIRED_FIELDS
from field_mapping import LINE_MONEY_FIELDS, MappingError, analyze_headers, confirm_mapping, resolve_mapping
from order_loader import InputCancelled, InputValidationError, OrderSource, ValidatedOrders


def _display_header(value):
    # 表头也可能包含客户信息或终端控制字符，仅显示有限、安全转义后的表头。
    return ascii(value[:60]) if any(ord(char) < 32 or ord(char) == 127 for char in value) else repr(value[:60])


def inspect_and_validate(path: str | Path) -> ValidatedOrders:
    """复用原交互，不打印业务结果；仅在确认且整份校验通过后返回。"""
    from excel_reader import ExcelImportError, header_candidates, read_snapshot, select_sheet, validate_selection
    mapping_location = ""
    try:
        snapshot = read_snapshot(path)
        print("已读取内存快照，源文件 SHA256：" + snapshot.sha256)
        print("不会默认选择活动工作表。可随时输入 q 取消。")
        for index, sheet in enumerate(snapshot.sheets, 1):
            print(f"{index}. {_display_header(sheet.name)}，{len(sheet.rows)} 行（含表头与空白），状态 {sheet.state}")
        answer = input("请选择工作表序号：")
        if answer == "q":
            raise InputCancelled()
        if not answer.isascii() or not answer.isdecimal() or not 1 <= int(answer) <= len(snapshot.sheets):
            raise MappingError("工作表序号无效，未选择任何默认表。")
        sheet = snapshot.sheets[int(answer) - 1]
        suggestions = header_candidates(sheet)
        print("候选表头行（只作提示，不自动选择）：" + str(suggestions))
        answer = input("请输入单行表头的原始行号：")
        if answer == "q":
            raise InputCancelled()
        if not answer.isascii() or not answer.isdecimal():
            raise MappingError("表头行号必须是正整数。")
        selection = select_sheet(snapshot, sheet.name, int(answer))
        mapping_location = f"工作表 {_display_header(sheet.name)} 第 {selection.header_row} 行表头："
        matches = analyze_headers(selection.headers)
        print("标准字段：" + "、".join(REQUIRED_FIELDS))
        print("先逐列展示候选。输入标准英文字段名；输入 - 明确忽略；回车仅接受唯一候选。")
        print("金额须为人民币元，优惠/支付/退款须已分摊到本明细；支付为退款前历史支付，退款为累计已完成退款。")
        choices = {}
        meanings = set()
        for match in matches:
            flag = "（重复表头，须明确选择）" if match.duplicate else ""
            print(f"第 {match.column} 列 {_display_header(match.title)}{flag} → 候选 {match.candidates or '无'}")
            answer = input("映射选择：")
            if answer == "q":
                raise InputCancelled()
            if answer == "-":
                choices[match.column] = None
            elif answer:
                choices[match.column] = answer
            elif len(match.candidates) == 1 and not match.duplicate:
                choices[match.column] = match.candidates[0]
            else:
                raise MappingError(f"第 {match.column} 列必须明确选择或忽略。")
            field = choices[match.column]
            if field is not None and (match.needs_meaning_confirmation or (
                    field in LINE_MONEY_FIELDS and field not in match.candidates)):
                answer = input("请核实：这是人民币元、分摊到本明细的该项金额，非重复整单总额？输入 明细 确认：")
                if answer != "明细":
                    raise InputCancelled()
                meanings.add(match.column)
        plan = resolve_mapping(selection.headers, choices, confirmed_meanings=meanings, context=selection.context)
        print("最终映射预览（忽略列不进入订单校验器）：")
        for column, (header, field) in enumerate(zip(plan.headers, plan.assignments), 1):
            print(f"第 {column} 列 {_display_header(header)} → {field or '明确忽略'}")
        print("必需字段齐全不代表业务口径正确。未知状态不会自动翻译；缺失退款件数不会补零。")
        if input("确认上述映射、明细金额粒度及北京时间规则后，输入 确认 开始整份校验：") != "确认":
            raise InputCancelled()
        records = validate_selection(selection, confirm_mapping(plan, True))
        return ValidatedOrders(tuple(records), OrderSource(Path(path).name, "XLSX", sheet.name,
                                                          selection.header_row, snapshot.sha256))
    except (ExcelImportError, MappingError) as error:
        location = mapping_location if isinstance(error, MappingError) else ""
        raise InputValidationError(f"{location}{error}") from error
    except (EOFError, KeyboardInterrupt) as error:
        raise InputCancelled() from error


def main(argv=None):
    parser = argparse.ArgumentParser(description="只读检查模拟 Excel 订单；不计算指标或生成报表。", allow_abbrev=False)
    parser.add_argument("xlsx_file", type=Path, help="普通 .xlsx 文件路径；原文件不修改")
    args = parser.parse_args(argv)
    try:
        orders = inspect_and_validate(args.xlsx_file)
    except ModuleNotFoundError as error:
        if error.name not in {"openpyxl", "et_xmlfile"}:
            raise
        print("缺少 Excel 读取依赖，请用本项目 .venv\\Scripts\\python.exe 运行；程序不会安装软件。")
        return 2
    except InputValidationError as error:
        print(f"未完成校验：{error}")
        print("未返回部分合格结果，未修改原文件，未生成报表。")
        return 1
    except InputCancelled:
        return _cancel()
    except Exception as error:
        # 明确报告程序异常及失败退出；不把可能包含订单原文的异常详情写入终端。
        print(f"程序异常（{type(error).__name__}），已停止；未返回校验成功结果，未生成报表。")
        return 1
    records = orders.records
    print(f"整份校验通过：{len(records)} 条商品明细，{len({row.order_id for row in records})} 个唯一订单编号。")
    print("结果仅对应上述 SHA256 的内存快照；未修改原文件，未计算运营指标，未生成报表。")
    return 0


def _cancel():
    print("已取消，未返回校验成功结果，未修改文件或生成报表。")
    return 2


if __name__ == "__main__":
    raise SystemExit(main())
