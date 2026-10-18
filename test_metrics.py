"""运营指标测试：仅使用原始模拟数据（只读）和隔离临时 CSV。"""

import contextlib
import csv
import hashlib
import io
import subprocess
import sys
import tempfile
import unittest
from dataclasses import FrozenInstanceError
from fractions import Fraction
from pathlib import Path

from csv_reader import CsvValidationError, REQUIRED_FIELDS
from metrics import calculate_metrics, format_money
from report_metrics import main, render_report


ROOT = Path(__file__).parent
SAMPLE = ROOT / "data" / "sample_orders.csv"


def row(order_id="A1", product_id="P1", **changes):
    """构造明确数值的模拟明细，不调用正式计算函数生成验收答案。"""
    values = dict(zip(REQUIRED_FIELDS, (
        order_id, order_id + "-01", "2026-10-01 10:00:00", product_id,
        "模拟商品" + product_id, "文具用品", "1", "10.00", "0.00", "10.00",
        "已支付", "0.00", "0", "平台搜索",
    )))
    values.update(changes)
    return values


class MetricsTests(unittest.TestCase):
    def setUp(self):
        temporary = tempfile.TemporaryDirectory(prefix="metrics_test_", dir=ROOT)
        self.addCleanup(temporary.cleanup)
        self.folder = Path(temporary.name)
        self.path = self.folder / "模拟订单.csv"

    def write(self, rows, headers=REQUIRED_FIELDS):
        with self.path.open("w", encoding="utf-8-sig", newline="") as stream:
            writer = csv.DictWriter(stream, fieldnames=headers)
            writer.writeheader()
            writer.writerows(rows)
        return self.path

    def calculate(self, rows):
        return calculate_metrics(self.write(rows))

    def test_sample_detail_and_unique_order_counts(self):
        report = calculate_metrics(SAMPLE)
        self.assertEqual((report.line_count, report.total_order_count), (80, 40))
        self.assertEqual((report.paid_order_count, report.net_order_count), (32, 28))

    def test_sample_payment_refund_and_net_sales(self):
        report = calculate_metrics(SAMPLE)
        self.assertEqual((report.payment_cents, report.refund_cents, report.net_sales_cents),
                         (372457, 90086, 282371))

    def test_sample_two_averages_keep_separate_denominators(self):
        report = calculate_metrics(SAMPLE)
        self.assertEqual(report.paid_average_cents, Fraction(372457, 32))
        self.assertEqual(report.net_average_cents, Fraction(282371, 28))
        self.assertEqual(format_money(report.paid_average_cents), "116.39")
        self.assertEqual(format_money(report.net_average_cents), "100.85")

    def test_sample_complete_product_ranking_matches_static_baseline(self):
        report = calculate_metrics(SAMPLE)
        actual = [(p.product_id, p.paid_quantity, p.refund_quantity, p.net_quantity, p.net_sales_cents)
                  for p in report.products]
        self.assertEqual(actual, [
            ("P003", 17, 0, 17, 50295), ("P001", 16, 0, 16, 30997),
            ("P007", 24, 8, 16, 26334), ("P005", 24, 12, 12, 70172),
            ("P006", 12, 0, 12, 29360), ("P010", 12, 0, 12, 26000),
            ("P008", 12, 2, 10, 7634), ("P004", 8, 0, 8, 27600),
            ("P009", 12, 4, 8, 7854), ("P002", 5, 0, 5, 6125),
        ])
        self.assertEqual(report.net_quantity, 116)

    def test_product_totals_reconcile_with_whole_report(self):
        report = calculate_metrics(SAMPLE)
        self.assertEqual(sum(p.payment_cents for p in report.products), report.payment_cents)
        self.assertEqual(sum(p.refund_cents for p in report.products), report.refund_cents)
        self.assertEqual(sum(p.net_sales_cents for p in report.products), report.net_sales_cents)
        self.assertEqual(sum(p.paid_quantity for p in report.products), 142)
        self.assertEqual(sum(p.refund_quantity for p in report.products), 26)

    def test_multi_product_order_is_counted_once_even_when_rows_are_separated(self):
        report = self.calculate([
            row(quantity="2", payment_amount="20.00"), row("B1"),
            row(product_id="P2", order_line_id="A1-02"),
        ])
        self.assertEqual((report.line_count, report.total_order_count, report.paid_order_count), (3, 2, 2))
        self.assertEqual((report.payment_cents, report.net_quantity), (4000, 4))
        self.assertEqual(format_money(report.paid_average_cents), "20.00")

    def test_same_order_same_product_distinct_details_are_not_dropped(self):
        report = self.calculate([row(), row(order_line_id="A1-02")])
        self.assertEqual((report.paid_order_count, report.payment_cents), (1, 2000))
        self.assertEqual(len(report.products), 1)
        self.assertEqual(report.products[0].net_quantity, 2)

    def test_full_refund_keeps_historical_payment_but_removes_net_order(self):
        report = self.calculate([row(), row("B1", order_status="已全额退款",
                                          refund_amount="10.00", refund_quantity="1")])
        self.assertEqual((report.payment_cents, report.refund_cents, report.net_sales_cents), (2000, 1000, 1000))
        self.assertEqual((report.paid_order_count, report.net_order_count, report.net_quantity), (2, 1, 1))
        self.assertEqual(format_money(report.paid_average_cents), "10.00")
        self.assertEqual(format_money(report.net_average_cents), "10.00")

    def test_partial_refund_preserves_one_net_order_and_uses_explicit_units(self):
        report = self.calculate([row(quantity="3", payment_amount="30.00", order_status="部分退款",
                                     refund_amount="8.23", refund_quantity="1")])
        self.assertEqual((report.paid_order_count, report.net_order_count), (1, 1))
        self.assertEqual((report.net_sales_cents, report.net_quantity), (2177, 2))

    def test_partial_order_with_a_fully_refunded_line_is_evaluated_as_whole_order(self):
        report = self.calculate([
            row(order_status="部分退款", refund_amount="10.00", refund_quantity="1"),
            row(product_id="P2", order_line_id="A1-02", order_status="部分退款"),
        ])
        self.assertEqual((report.paid_order_count, report.net_order_count), (1, 1))
        self.assertEqual((report.net_sales_cents, report.net_quantity), (1000, 1))
        self.assertEqual([p.product_id for p in report.products], ["P2", "P1"])

    def test_unpaid_and_cancelled_rows_do_not_contribute_payment_or_units(self):
        report = self.calculate([
            row(), row("B1", "P2", order_status="未支付", payment_amount="0.00"),
            row("C1", "P3", order_status="已取消", payment_amount="0.00"),
        ])
        self.assertEqual((report.total_order_count, report.paid_order_count), (3, 1))
        self.assertEqual((report.payment_cents, report.net_quantity), (1000, 1))
        self.assertEqual([p.product_id for p in report.products], ["P1"])

    def test_price_only_refund_changes_money_but_does_not_remove_units(self):
        report = self.calculate([row(quantity="2", payment_amount="20.00", order_status="部分退款",
                                     refund_amount="10.00", refund_quantity="0")])
        self.assertEqual((report.net_sales_cents, report.net_quantity), (1000, 2))
        self.assertEqual(report.products[0].refund_quantity, 0)

    def test_net_order_is_based_on_amount_not_on_remaining_units(self):
        report = self.calculate([row(order_status="部分退款", refund_amount="5.00", refund_quantity="1")])
        self.assertEqual((report.net_sales_cents, report.net_quantity, report.net_order_count), (500, 0, 1))

    def test_same_name_with_different_ids_is_not_merged(self):
        report = self.calculate([row(product_name="同名商品"),
                                 row("B1", "P2", product_name="同名商品")])
        self.assertEqual([p.product_id for p in report.products], ["P1", "P2"])
        self.assertEqual([p.net_quantity for p in report.products], [1, 1])

    def test_tie_sort_is_product_id_order_independent_of_input_order(self):
        rows = [row("B1", "P2"), row("A1", "P1")]
        first = self.calculate(rows)
        second = self.calculate(list(reversed(rows)))
        self.assertEqual(first, second)
        self.assertEqual([p.product_id for p in first.products], ["P1", "P2"])

    def test_product_id_tie_break_is_explicitly_string_order(self):
        report = self.calculate([row("A1", "P2"), row("B1", "P10")])
        self.assertEqual([p.product_id for p in report.products], ["P10", "P2"])

    def test_tenths_sum_exactly_without_binary_float(self):
        report = self.calculate([
            row(unit_price="0.10", payment_amount="0.10"),
            row("B1", "P2", unit_price="0.20", payment_amount="0.20"),
        ])
        self.assertEqual(report.payment_cents, 30)
        self.assertIsInstance(report.payment_cents, int)
        self.assertEqual(report.paid_average_cents, Fraction(15))
        self.assertEqual(format_money(report.payment_cents), "0.30")

    def test_average_stays_exact_and_half_cent_rounds_up_only_when_displayed(self):
        report = self.calculate([
            row(unit_price="1.00", payment_amount="1.00"),
            row("B1", "P2", unit_price="1.01", payment_amount="1.01"),
        ])
        self.assertEqual(report.paid_average_cents, Fraction(201, 2))
        self.assertEqual(format_money(report.paid_average_cents), "1.01")
        render_report(report)
        self.assertEqual(report.paid_average_cents, Fraction(201, 2))
        self.assertEqual(report.payment_cents, 201)

    def test_display_rounding_below_at_and_above_half_cent(self):
        for value, expected in ((Fraction(10049, 100), "1.00"), (Fraction(201, 2), "1.01"),
                                (Fraction(10051, 100), "1.01"), (Fraction(1999, 2), "10.00")):
            with self.subTest(value=value):
                self.assertEqual(format_money(value), expected)

    def test_large_amount_exceeding_decimal_default_precision_is_exact(self):
        amount = "123456789012345678901234567890.12"
        report = self.calculate([row(quantity="2", unit_price=amount,
                                     payment_amount="246913578024691357802469135780.24")])
        self.assertEqual(report.payment_cents, 24691357802469135780246913578024)
        self.assertEqual(format_money(report.paid_average_cents), "246913578024691357802469135780.24")

    def test_zero_paid_orders_in_valid_input_has_no_division_error(self):
        report = self.calculate([
            row(order_status="未支付", payment_amount="0.00"),
            row("B1", order_status="已取消", payment_amount="0.00"),
        ])
        self.assertEqual((report.payment_cents, report.refund_cents, report.net_sales_cents), (0, 0, 0))
        self.assertEqual((report.paid_order_count, report.net_order_count, report.net_quantity), (0, 0, 0))
        self.assertIsNone(report.paid_average_cents)
        self.assertIsNone(report.net_average_cents)
        self.assertEqual(report.products, ())
        text = render_report(report)
        self.assertEqual(text.count("无法计算（分母为零）"), 2)
        self.assertIn("暂无已支付商品", text)

    def test_all_fully_refunded_orders_have_zero_net_sales_and_unavailable_net_average(self):
        report = self.calculate([row(order_status="已全额退款", refund_amount="10.00", refund_quantity="1")])
        self.assertEqual((report.paid_order_count, report.net_order_count), (1, 0))
        self.assertEqual((report.net_sales_cents, report.net_quantity), (0, 0))
        self.assertEqual(format_money(report.paid_average_cents), "10.00")
        self.assertIsNone(report.net_average_cents)
        self.assertEqual(len(report.products), 1)  # 保留支付过但净成交为零的商品。

    def test_empty_and_header_only_files_still_fail_existing_validation(self):
        self.write([])
        with self.assertRaises(CsvValidationError):
            calculate_metrics(self.path)
        self.path.write_bytes(b"")
        with self.assertRaises(CsvValidationError):
            calculate_metrics(self.path)

    def test_late_invalid_row_refuses_any_official_or_partial_result(self):
        self.write([row(), row("B1", quantity="0")])
        with self.assertRaises(CsvValidationError) as caught:
            calculate_metrics(self.path)
        self.assertIn("第 3 行 [quantity]", str(caught.exception))
        self.assertFalse(hasattr(caught.exception, "report"))
        output = io.StringIO()
        with contextlib.redirect_stdout(output):
            self.assertEqual(main([str(self.path)]), 1)
        text = output.getvalue()
        self.assertIn("拒绝输出正式指标或部分统计", text)
        for forbidden in ("运营指标报告", "支付金额：", "净销售额：", "商品销量排行"):
            self.assertNotIn(forbidden, text)

    def test_invalid_whole_order_status_refuses_report(self):
        self.write([row(order_status="部分退款")])
        with self.assertRaises(CsvValidationError) as caught:
            calculate_metrics(self.path)
        self.assertIn("整单退款必须大于零", str(caught.exception))

    def test_missing_refund_quantity_is_rejected_instead_of_guessed(self):
        self.write([row(order_status="部分退款", refund_amount="5.00", refund_quantity="")])
        with self.assertRaises(CsvValidationError) as caught:
            calculate_metrics(self.path)
        self.assertIn("refund_quantity", str(caught.exception))

    def test_read_error_is_reported_without_metrics(self):
        output = io.StringIO()
        with contextlib.redirect_stdout(output):
            self.assertEqual(main([str(self.folder / "不存在.csv")]), 1)
        self.assertIn("无法读取", output.getvalue())
        self.assertNotIn("支付金额：", output.getvalue())

    def test_utf8_bom_and_reordered_headers_use_existing_reader(self):
        self.write([row()], headers=tuple(reversed(REQUIRED_FIELDS)))
        report = calculate_metrics(self.path)
        self.assertEqual((report.payment_cents, report.net_quantity), (1000, 1))

    def test_sample_and_temporary_inputs_are_not_modified_on_success_or_failure(self):
        before = hashlib.sha256(SAMPLE.read_bytes()).digest()
        calculate_metrics(SAMPLE)
        self.assertEqual(hashlib.sha256(SAMPLE.read_bytes()).digest(), before)
        for values in ([row()], [row(quantity="0")]):
            self.write(values)
            content = self.path.read_bytes()
            try:
                calculate_metrics(self.path)
            except CsvValidationError:
                pass
            self.assertEqual(self.path.read_bytes(), content)
        self.assertEqual([p.name for p in self.folder.iterdir()], [self.path.name])

    def test_result_and_product_values_are_immutable(self):
        report = self.calculate([row()])
        with self.assertRaises(FrozenInstanceError):
            report.payment_cents = 0
        with self.assertRaises(FrozenInstanceError):
            report.products[0].paid_quantity = 0
        self.assertIsInstance(report.products, tuple)

    def test_cli_default_sample_from_another_working_directory(self):
        result = subprocess.run([sys.executable, "-B", "-X", "utf8", str(ROOT / "report_metrics.py")],
                                cwd=self.folder, capture_output=True, text=True, encoding="utf-8", timeout=20)
        self.assertEqual(result.returncode, 0, result.stderr)
        for expected in ("商品明细：80 条", "唯一订单：40 笔", "支付金额：3724.57 元",
                         "已完成退款金额：900.86 元", "净销售额：2823.71 元", "支付订单数：32 笔",
                         "净有效订单数：28 笔", "支付客单价：116.39 元/笔", "净客单价：100.85 元/笔",
                         "商品净成交件数合计：116 件"):
            self.assertIn(expected, result.stdout)
        self.assertEqual(list(self.folder.iterdir()), [])

    def test_cli_invalid_input_process_exits_with_no_official_metrics(self):
        self.write([row(), row("B1", payment_amount="错误金额")])
        content = self.path.read_bytes()
        result = subprocess.run([sys.executable, "-B", "-X", "utf8", str(ROOT / "report_metrics.py"), str(self.path)],
                                cwd=self.folder, capture_output=True, text=True, encoding="utf-8", timeout=20)
        self.assertEqual(result.returncode, 1, result.stderr)
        self.assertIn("第 3 行 [payment_amount]", result.stdout)
        self.assertIn("拒绝输出正式指标或部分统计", result.stdout)
        self.assertNotIn("运营指标报告", result.stdout)
        self.assertNotIn("支付金额：", result.stdout)
        self.assertEqual(self.path.read_bytes(), content)

    def test_imports_do_not_read_data_or_write_files(self):
        script = '''from unittest.mock import patch
with patch("pathlib.Path.read_bytes", side_effect=AssertionError("unexpected read")), patch("pathlib.Path.write_bytes", side_effect=AssertionError("unexpected write")), patch("pathlib.Path.write_text", side_effect=AssertionError("unexpected write")):
    import metrics, report_metrics
print("IMPORT_SAFE")
'''
        result = subprocess.run([sys.executable, "-B", "-X", "utf8", "-c", script], cwd=ROOT,
                                capture_output=True, text=True, encoding="utf-8", timeout=20)
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertIn("IMPORT_SAFE", result.stdout)


if __name__ == "__main__":
    unittest.main(verbosity=2)
