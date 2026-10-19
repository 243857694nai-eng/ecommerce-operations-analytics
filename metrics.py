"""只在整份输入校验通过后计算指标；金额以整数分和精确分数保存。"""

from collections import defaultdict
from dataclasses import dataclass, field
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
    for row in records:
        if row.order_status not in PAID_STATUSES:
            continue
        payment += row.payment_amount_cents
        refund += row.refund_amount_cents
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
