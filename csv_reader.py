"""只读 CSV 读取：检查编码、表头、行结构，并保留原始物理行号。"""

import csv
import io
import re
from collections import Counter
from dataclasses import dataclass
from pathlib import Path


REQUIRED_FIELDS = (
    "order_id", "order_line_id", "order_time", "product_id", "product_name",
    "category", "quantity", "unit_price", "discount_amount", "payment_amount",
    "order_status", "refund_amount", "refund_quantity", "channel",
)


@dataclass(frozen=True)
class ValidationIssue:
    line_number: int | None
    field: str
    reason: str
    end_line_number: int | None = None

    def __str__(self):
        if self.line_number is None:
            location = "文件"
        elif self.end_line_number and self.end_line_number != self.line_number:
            location = f"第 {self.line_number}–{self.end_line_number} 行"
        else:
            location = f"第 {self.line_number} 行"
        return f"{location} [{self.field}]：{self.reason}"


class CsvValidationError(Exception):
    """发现错误时携带问题清单，不返回可被误用的部分数据。"""

    def __init__(self, issues):
        self.issues = tuple(issues)
        super().__init__("\n".join(str(issue) for issue in self.issues))


@dataclass(frozen=True)
class RawCsvRow:
    line_number: int
    end_line_number: int
    values: dict[str, str]


def read_csv(path: str | Path) -> list[RawCsvRow]:
    """读取 UTF-8（有无 BOM 均可），不改写文件或修复任何内容。"""
    try:
        content = Path(path).read_bytes()
    except (OSError, ValueError) as error:
        raise CsvValidationError([ValidationIssue(None, "文件路径", f"无法读取：{error}")]) from error
    try:
        text = content.decode("utf-8-sig")
    except UnicodeDecodeError as error:
        # 解码错误发生在某个字节；此前的换行数量可以定位物理行。
        line = len(re.findall(rb"\r\n|\r|\n", error.object[:error.start])) + 1
        raise CsvValidationError([ValidationIssue(line, "编码", "不是有效的 UTF-8 文件；不会猜测编码或替换乱码。")]) from error

    reader = csv.reader(io.StringIO(text, newline=""), strict=True)
    try:
        header = next(reader)
    except StopIteration:
        raise CsvValidationError([ValidationIssue(1, "表头", "CSV 为空，缺少表头和商品明细。")])
    except csv.Error as error:
        raise CsvValidationError([ValidationIssue(1, "CSV 格式", f"无法解析表头，请检查引号和分隔结构：{error}", reader.line_num)]) from error

    issues = []
    if any(not value or value != value.strip() for value in header):
        issues.append(ValidationIssue(1, "表头", "字段名不能为空或包含前后空白；不会自动修正。"))
    for field, count in Counter(header).items():
        if count > 1:
            issues.append(ValidationIssue(1, "表头", f"字段 {field!r} 重复出现。"))
    missing = sorted(set(REQUIRED_FIELDS) - set(header))
    extra = sorted(set(header) - set(REQUIRED_FIELDS))
    if missing:
        issues.append(ValidationIssue(1, "表头", "缺少必需字段：" + "、".join(missing)))
    if extra:
        issues.append(ValidationIssue(1, "表头", "当前 V1 不接受未知字段：" + "、".join(extra)))
    if issues:
        raise CsvValidationError(issues)

    rows = []
    while True:
        start = reader.line_num + 1
        try:
            values = next(reader)
        except StopIteration:
            break
        except csv.Error as error:
            issues.append(ValidationIssue(start, "CSV 格式", f"无法解析记录，请检查引号、分隔符或字段长度：{error}", reader.line_num))
            break  # 结构已经不可信，不能猜测后面的记录边界。
        if not values:
            issues.append(ValidationIssue(start, "明细", "存在空白记录，不能静默跳过。", reader.line_num))
        elif len(values) != len(header):
            issues.append(ValidationIssue(start, "列数", f"应为 {len(header)} 列，实际为 {len(values)} 列。", reader.line_num))
        else:
            rows.append(RawCsvRow(start, reader.line_num, dict(zip(header, values))))
    if not rows and not issues:
        issues.append(ValidationIssue(None, "明细", "文件只有表头，没有商品明细。"))
    if issues:
        raise CsvValidationError(issues)
    return rows
