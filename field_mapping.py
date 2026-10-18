"""只负责表头与标准字段的对应；不读取文件，不校验或修改订单。"""

from collections import Counter
from dataclasses import dataclass, replace

from csv_reader import REQUIRED_FIELDS


ALIASES = {
    "order_id": ("订单编号", "订单号"),
    "order_line_id": ("订单明细编号", "明细编号"),
    "order_time": ("下单时间",),
    "product_id": ("商品编号",),
    "product_name": ("商品名称",),
    "category": ("商品类别",),
    "quantity": ("购买数量",),
    "unit_price": ("商品单价（元）",),
    "discount_amount": ("明细优惠金额（元）",),
    "payment_amount": ("明细支付金额（元）",),
    "order_status": ("订单状态",),
    "refund_amount": ("明细退款金额（元）",),
    "refund_quantity": ("退款件数",),
    "channel": ("下单渠道",),
}
# 只提供待确认候选，绝不把这些名称当作已确定的金额粒度。
AMBIGUOUS_ALIASES = {
    "实付金额": ("payment_amount",),
    "订单优惠": ("discount_amount",),
    "退款金额": ("refund_amount",),
}
LINE_MONEY_FIELDS = frozenset({"discount_amount", "payment_amount", "refund_amount"})


class MappingError(ValueError):
    """映射不完整、含歧义或尚未确认；不能继续校验。"""


@dataclass(frozen=True)
class HeaderMatch:
    column: int  # Excel 列位置从 1 开始，重复表头仍各占一项。
    title: str
    candidates: tuple[str, ...]
    needs_meaning_confirmation: bool
    duplicate: bool


@dataclass(frozen=True)
class MappingPlan:
    context: str
    headers: tuple[str, ...]
    assignments: tuple[str | None, ...]  # None 是明确忽略，并非缺省决定。
    confirmed_meanings: tuple[int, ...]
    confirmed: bool = False


def analyze_headers(headers) -> tuple[HeaderMatch, ...]:
    """只对表头去除首尾空白；不改变大小写、标点或字段值。"""
    headers = tuple(headers)
    if not headers or any(not isinstance(title, str) or not title.strip() for title in headers):
        raise MappingError("表头必须是非空文本；不能猜测空白或数字表头的含义。")
    titles = tuple(title.strip() for title in headers)
    counts = Counter(titles)
    result = []
    for column, title in enumerate(titles, 1):
        candidates = tuple(field for field in REQUIRED_FIELDS if title == field or title in ALIASES[field])
        ambiguous = title in AMBIGUOUS_ALIASES
        if ambiguous:
            candidates = AMBIGUOUS_ALIASES[title]
        result.append(HeaderMatch(column, title, candidates, ambiguous, counts[title] > 1))
    return tuple(result)


def resolve_mapping(headers, choices=None, *, confirmed_meanings=(), context="") -> MappingPlan:
    """choices 用列号明确选择标准字段或 None；未知额外列不能默认忽略。"""
    headers = tuple(headers)
    matches = analyze_headers(headers)
    choices = {} if choices is None else dict(choices)
    meanings = frozenset(confirmed_meanings)
    valid_columns = set(range(1, len(headers) + 1))
    if any(type(column) is not int or column not in valid_columns for column in (*choices, *meanings)):
        raise MappingError("选择中包含非法列位置。")
    assignments = []
    for match in matches:
        if match.column in choices:
            field = choices[match.column]
        elif len(match.candidates) == 1 and not match.duplicate:
            field = match.candidates[0]
        else:
            raise MappingError(f"第 {match.column} 列需明确映射或忽略；不会猜测或丢弃。")
        if field is not None and field not in REQUIRED_FIELDS:
            raise MappingError(f"第 {match.column} 列选择了未知标准字段。")
        # 手动把未知列改映射到金额字段，同样必须确认明细粒度。
        manual_money = field in LINE_MONEY_FIELDS and field not in match.candidates
        if field is not None and (match.needs_meaning_confirmation or manual_money) and match.column not in meanings:
            raise MappingError(f"第 {match.column} 列的金额含义未确认：须为人民币元、分摊到本明细的金额。")
        assignments.append(field)
    counts = Counter(field for field in assignments if field is not None)
    repeated = [field for field, count in counts.items() if count > 1]
    if repeated:
        locations = [f"{field}（列 {','.join(str(i) for i, f in enumerate(assignments, 1) if f == field)}）"
                     for field in repeated]
        raise MappingError("多个源列匹配同一标准字段，须明确选择：" + "、".join(locations))
    missing = [field for field in REQUIRED_FIELDS if field not in counts]
    if missing:
        raise MappingError("缺少必需字段：" + "、".join(missing) + "；确认不能补造缺失数据。")
    return MappingPlan(context, headers, tuple(assignments), tuple(sorted(meanings)))


def confirm_mapping(plan: MappingPlan, approved: bool) -> MappingPlan:
    if approved is not True:
        raise MappingError("用户未明确确认映射，已停止，未开始订单校验。")
    return replace(plan, confirmed=True)
