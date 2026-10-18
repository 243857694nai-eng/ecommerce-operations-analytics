"""按当前 V1 口径校验订单，不计算运营指标，不补值或删除异常行。"""

import re
from collections import defaultdict
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from pathlib import Path

from csv_reader import CsvValidationError, RawCsvRow, ValidationIssue, read_csv


PAID_STATUSES = frozenset({"已支付", "已全额退款", "部分退款"})
STATUSES = PAID_STATUSES | {"未支付", "已取消"}
CATEGORIES = frozenset({"文具用品", "家居收纳", "日用生活", "数码配件"})
CHANNELS = frozenset({"平台搜索", "直播间", "内容推荐", "社群推广"})
BEIJING_TIME = timezone(timedelta(hours=8))


@dataclass(frozen=True)
class OrderLine:
    source_line: int
    order_id: str
    order_line_id: str
    order_time: datetime
    product_id: str
    product_name: str
    category: str
    quantity: int
    unit_price_cents: int
    discount_amount_cents: int
    payment_amount_cents: int
    order_status: str
    refund_amount_cents: int
    refund_quantity: int
    channel: str


def _parse_row(raw: RawCsvRow, issues: list[ValidationIssue]) -> OrderLine | None:
    values = raw.values
    start_count = len(issues)

    def problem(field, reason):
        issues.append(ValidationIssue(raw.line_number, field, reason, raw.end_line_number))

    for field, value in values.items():
        if not value.strip():
            problem(field, "必需值不能为空或全是空白。")
        elif value != value.strip():
            problem(field, "包含前后空白；请核实数据，程序不会自动去除。")
        elif any(ord(character) < 32 or ord(character) == 127 for character in value):
            problem(field, "不允许换行、制表符或其他控制字符。")
    if len(issues) != start_count:
        return None

    def integer(field, positive=False):
        pattern = r"[1-9][0-9]*" if positive else r"(?:0|[1-9][0-9]*)"
        if not re.fullmatch(pattern, values[field]):
            problem(field, "必须是正整数。" if positive else "必须是非负整数。")
            return None
        try:
            return int(values[field])
        except ValueError:
            problem(field, "数字过长，超出当前 Python 的整数解析范围。")
            return None

    def money(field):
        if not re.fullmatch(r"(?:0|[1-9][0-9]*)\.[0-9]{2}", values[field]):
            problem(field, "金额必须为非负十进制数并恰有两位小数，例如 19.90；不接受指数、千分位或 NaN。")
            return None
        try:
            yuan, cents = values[field].split(".")
            return int(yuan) * 100 + int(cents)
        except ValueError:
            problem(field, "金额数字过长，超出当前 Python 的整数解析范围。")
            return None

    quantity = integer("quantity", positive=True)
    refund_quantity = integer("refund_quantity")
    price = money("unit_price")
    discount = money("discount_amount")
    payment = money("payment_amount")
    refund = money("refund_amount")
    moment = None
    if not re.fullmatch(r"[0-9]{4}-[0-9]{2}-[0-9]{2} [0-9]{2}:[0-9]{2}:[0-9]{2}", values["order_time"]):
        problem("order_time", "日期格式必须为 YYYY-MM-DD HH:MM:SS，例如 2026-10-01 10:07:11。")
    else:
        try:
            moment = datetime.strptime(values["order_time"], "%Y-%m-%d %H:%M:%S").replace(tzinfo=BEIJING_TIME)
        except ValueError:
            problem("order_time", "日期或时间不存在，请检查月份、天数和时分秒。")
    for field, allowed in (("order_status", STATUSES), ("category", CATEGORIES), ("channel", CHANNELS)):
        if values[field] not in allowed:
            problem(field, "不属于当前 V1 允许的取值：" + "、".join(sorted(allowed)))
    if len(issues) != start_count:
        return None

    if price <= 0:
        problem("unit_price", "单价必须大于零，当前 V1 不支持零元商品。")
    if discount > quantity * price:
        problem("discount_amount", "优惠金额超过本行数量乘单价。")
    if refund > payment:
        problem("refund_amount", "退款金额超过本行实际支付金额。")
    if refund_quantity > quantity:
        problem("refund_quantity", "退款件数超过本行购买数量。")
    if refund_quantity > 0 and refund == 0:
        problem("refund_quantity", "存在退款商品件数，但退款金额为零，与当前 V1 口径矛盾。")

    status = values["order_status"]
    if status in PAID_STATUSES:
        if payment <= 0:
            problem("payment_amount", "已支付类明细的实际支付必须大于零。")
        if payment != quantity * price - discount:
            problem("payment_amount", "实际支付必须等于数量乘单价减去本行优惠。")
    else:
        for field, value in (("payment_amount", payment), ("refund_amount", refund), ("refund_quantity", refund_quantity)):
            if value != 0:
                problem(field, "未支付和支付前取消订单的支付、退款及退款件数必须为零。")
    if status == "已支付" and (refund != 0 or refund_quantity != 0):
        problem("order_status", "已支付未退款状态不能存在退款金额或退款件数。")
    if status == "已全额退款":
        if refund != payment:
            problem("refund_amount", "全额退款明细的退款金额必须等于实际支付。")
        if refund_quantity != quantity:
            problem("refund_quantity", "当前 V1 全额退款明细的退款件数必须等于购买数量。")
    if len(issues) != start_count:
        return None

    return OrderLine(raw.line_number, values["order_id"], values["order_line_id"], moment,
                     values["product_id"], values["product_name"], values["category"], quantity,
                     price, discount, payment, status, refund, refund_quantity, values["channel"])


def validate_rows(rows: list[RawCsvRow]) -> list[OrderLine]:
    """只在所有明细通过时返回类型明确的数据；任何错误均拒绝整份输入。"""
    if not rows:
        raise CsvValidationError([ValidationIssue(None, "明细", "没有可校验的商品明细。")])
    issues = []
    records = []
    seen_ids = {}
    order_first = {}
    product_first = {}
    groups = defaultdict(list)
    invalid_orders = set()
    for raw in rows:
        order_id = raw.values["order_id"]
        detail_id = raw.values["order_line_id"]
        if detail_id in seen_ids:
            issues.append(ValidationIssue(raw.line_number, "order_line_id",
                                          f"明细编号重复，首次出现在第 {seen_ids[detail_id]} 行。", raw.end_line_number))
            invalid_orders.add(order_id)
        else:
            seen_ids[detail_id] = raw.line_number
        record = _parse_row(raw, issues)
        if record is None:
            invalid_orders.add(order_id)
            continue

        first = order_first.setdefault(order_id, record)
        for field in ("order_time", "order_status", "channel"):
            if getattr(record, field) != getattr(first, field):
                issues.append(ValidationIssue(raw.line_number, field,
                                              f"同一订单 {order_id} 的信息与第 {first.source_line} 行不一致。", raw.end_line_number))
                invalid_orders.add(order_id)
        product = product_first.setdefault(record.product_id, record)
        for field, attribute in (("product_name", "product_name"), ("category", "category"), ("unit_price", "unit_price_cents")):
            if getattr(record, attribute) != getattr(product, attribute):
                issues.append(ValidationIssue(raw.line_number, field,
                                              f"同一商品编号的信息与第 {product.source_line} 行不一致（当前 V1 规则）。", raw.end_line_number))
                invalid_orders.add(order_id)
        groups[order_id].append(record)
        records.append(record)

    for order_id, group in groups.items():
        if order_id in invalid_orders:
            continue  # 避免对不完整的错误订单求和后，再报出误导性的状态错误。
        if group[0].order_status == "部分退款":
            payment = sum(row.payment_amount_cents for row in group)
            refund = sum(row.refund_amount_cents for row in group)
            if not 0 < refund < payment:
                issues.append(ValidationIssue(group[0].source_line, "order_status",
                                              f"订单 {order_id} 标为部分退款，但整单退款必须大于零且小于整单支付。"))
    if issues:
        raise CsvValidationError(issues)
    return records


def load_and_validate(path: str | Path) -> list[OrderLine]:
    return validate_rows(read_csv(path))
