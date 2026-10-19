"""统一输入适配：CSV 保持原路径；XLSX 复用 5A 交互与整份校验。"""

from dataclasses import dataclass
from pathlib import Path

from order_validator import OrderLine, load_and_validate


class InputCancelled(Exception):
    """用户取消或交互中断；不能继续计算或导出。"""


class InputValidationError(ValueError):
    """Excel 读取、映射或校验失败，不携带部分合格结果。"""


@dataclass(frozen=True)
class OrderSource:
    file_name: str
    file_format: str
    sheet_name: str | None = None
    header_row: int | None = None
    sha256: str | None = None


@dataclass(frozen=True)
class ValidatedOrders:
    """内部传递完整校验结果；这不是防范恶意 Python 调用者的权限系统。"""

    records: tuple[OrderLine, ...]
    source: OrderSource


def load_orders(path: str | Path) -> ValidatedOrders:
    path = Path(path)
    if path.suffix.lower() == ".xlsx":
        # 延迟导入，CSV 校验和指标仍可仅用标准库运行，无 Excel 依赖也不弹出交互。
        from inspect_excel import inspect_and_validate
        return inspect_and_validate(path)
    # 原有 CSV 接口并未要求特定后缀，继续保留这一行为，不新增静默格式猜测。
    return ValidatedOrders(tuple(load_and_validate(path)), OrderSource(path.name, "CSV"))
