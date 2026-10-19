"""只在整份输入校验通过后计算指标；金额以整数分和精确分数保存。"""

from collections import defaultdict
from dataclasses import dataclass, field
from datetime import date
from fractions import Fraction
from pathlib import Path

from order_validator import PAID_STATUSES
from order_loader import OrderSource, load_orders


@dataclass(frozen=True)
class ProductMetrics:
    product_id: str
    product_name: str
    category: str
    paid_quantity: int
    refund_quantity: int
    payment_cents: int
    refund_cents: int

    @property
    def net_quantity(self) -> int:
        return self.paid_quantity - self.refund_quantity

    @property
    def net_sales_cents(self) -> int:
        return self.payment_cents - self.refund_cents


@dataclass(frozen=True)
class DailyMetrics:
    """按北京时间下单日期归组的当前订单快照，不是资金发生日流水。"""

    day: date
    payment_cents: int
    refund_cents: int
    paid_order_count: int
    refund_order_count: int

    @property
    def net_sales_cents(self) -> int:
        return self.payment_cents - self.refund_cents


@dataclass(frozen=True)
class MetricsReport:
    line_count: int
    total_order_count: int
    payment_cents: int
    refund_cents: int
    paid_order_count: int
    net_order_count: int
    products: tuple[ProductMetrics, ...]
    # 来源用于追溯，不参与业务指标值的等价比较。
    source: OrderSource | None = field(default=None, compare=False)
    daily: tuple[DailyMetrics, ...] = ()
    full_refund_order_count: int = 0
    partial_refund_order_count: int = 0

    @property
    def refund_order_count(self) -> int:
        return self.full_refund_order_count + self.partial_refund_order_count

    @property
    def refund_amount_ratio(self) -> Fraction | None:
        return Fraction(self.refund_cents, self.payment_cents) if self.payment_cents else None

    @property
    def refund_order_ratio(self) -> Fraction | None:
        return Fraction(self.refund_order_count, self.paid_order_count) if self.paid_order_count else None

    @property
    def net_sales_cents(self) -> int:
        return self.payment_cents - self.refund_cents

    @property
    def paid_average_cents(self) -> Fraction | None:
        if self.paid_order_count == 0:
            return None
        return Fraction(self.payment_cents, self.paid_order_count)

    @property
    def net_average_cents(self) -> Fraction | None:
        if self.net_order_count == 0:
            return None
        return Fraction(self.net_sales_cents, self.net_order_count)

    @property
    def net_quantity(self) -> int:
        return sum(product.net_quantity for product in self.products)


def calculate_metrics(path: str | Path) -> MetricsReport:
    """公共入口：CSV 沿用原路径，Excel 必须先映射确认并完整校验，无部分报表。"""
    orders = load_orders(path)
    records = orders.records
    # 在这里之前不计算指标。复用现有行级金额、状态和整单一致性规则。
    order_net = defaultdict(int)
    product_lines = defaultdict(list)
    payment = 0
    refund = 0
    daily_amounts = defaultdict(lambda: [0, 0])
    daily_paid_orders = defaultdict(set)
    daily_refund_orders = defaultdict(set)
    paid_statuses = {}
    for row in records:
        day = row.order_time.date()
        day_amounts = daily_amounts[day]  # 未支付或取消订单所在日期保留为零，不填造缺失日期。
        if row.order_status not in PAID_STATUSES:
            continue
        payment += row.payment_amount_cents
        refund += row.refund_amount_cents
        day_amounts[0] += row.payment_amount_cents
        day_amounts[1] += row.refund_amount_cents
        daily_paid_orders[day].add(row.order_id)
        if row.refund_amount_cents > 0:
            daily_refund_orders[day].add(row.order_id)
        paid_statuses[row.order_id] = row.order_status
        # 每个订单只有一个键；一单多商品仍只算一笔订单。
        order_net[row.order_id] += row.payment_amount_cents - row.refund_amount_cents
        product_lines[row.product_id].append(row)

    products = []
    for product_id, rows in product_lines.items():
        products.append(ProductMetrics(
            product_id=product_id,
            product_name=rows[0].product_name,
            category=rows[0].category,
            paid_quantity=sum(row.quantity for row in rows),
            refund_quantity=sum(row.refund_quantity for row in rows),
            payment_cents=sum(row.payment_amount_cents for row in rows),
            refund_cents=sum(row.refund_amount_cents for row in rows),
        ))
    # 并列时按商品编号的字符串升序；不依赖输入行的先后顺序。
    products.sort(key=lambda product: (-product.net_quantity, product.product_id))
    return MetricsReport(
        line_count=len(records),
        total_order_count=len({row.order_id for row in records}),
        payment_cents=payment,
        refund_cents=refund,
        paid_order_count=len(order_net),
        net_order_count=sum(net > 0 for net in order_net.values()),
        products=tuple(products),
        source=orders.source,
        daily=tuple(DailyMetrics(day, amounts[0], amounts[1], len(daily_paid_orders[day]),
                                 len(daily_refund_orders[day]))
                    for day, amounts in sorted(daily_amounts.items())),
        full_refund_order_count=sum(status == "已全额退款" for status in paid_statuses.values()),
        partial_refund_order_count=sum(status == "部分退款" for status in paid_statuses.values()),
    )


def format_money(cents: int | Fraction | None) -> str:
    """展示元，按四舍五入保留两位；不改变原始金额或精确客单价。"""
    if cents is None:
        return "无法计算（分母为零）"
    value = Fraction(cents)
    # 用整数比较是否达到半分，避免 float 误差和 Decimal 默认精度限制。
    rounded, remainder = divmod(abs(value.numerator), value.denominator)
    if remainder * 2 >= value.denominator:
        rounded += 1
    sign = "-" if value < 0 and rounded else ""
    yuan, fraction = divmod(rounded, 100)
    return f"{sign}{yuan}.{fraction:02d}"
