"""正常和异常 CSV 的隔离测试；不会改写项目的原始样本。"""

import contextlib
import csv
import hashlib
import io
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from csv_reader import CsvValidationError, REQUIRED_FIELDS, read_csv
from order_validator import load_and_validate, validate_rows
from validate_orders import main


ROOT = Path(__file__).parent


def sample_row(**changes):
    row = dict(zip(REQUIRED_FIELDS, (
        "A001", "A001-01", "2026-10-01 10:07:11", "P001", "学习文具", "文具用品",
        "2", "19.90", "3.98", "35.82", "已支付", "0.00", "0", "平台搜索",
    )))
    row.update(changes)
    return row


def second_row(**changes):
    row = sample_row(order_line_id="A001-02", product_id="P002", product_name="学习笔记本",
                     quantity="1", unit_price="12.50", discount_amount="1.25", payment_amount="11.25")
    row.update(changes)
    return row


class CsvValidationTests(unittest.TestCase):
    def setUp(self):
        temporary = tempfile.TemporaryDirectory(prefix="csv_validation_", dir=ROOT)
        self.addCleanup(temporary.cleanup)
        self.folder = Path(temporary.name)
        self.path = self.folder / "模拟订单.csv"

    def write(self, rows=None, headers=REQUIRED_FIELDS, encoding="utf-8-sig"):
        if rows is None:
            rows = [sample_row(), second_row()]
        stream = io.StringIO(newline="")
        writer = csv.writer(stream)
        writer.writerow(headers)
        writer.writerows([row.get(field, "") for field in headers] for row in rows)
        self.path.write_bytes(stream.getvalue().encode(encoding))
        return self.path

    def rejected(self, field=None, line=None):
        before = self.path.read_bytes()
        with self.assertRaises(CsvValidationError) as caught:
            load_and_validate(self.path)
        self.assertEqual(self.path.read_bytes(), before)
        self.assertFalse(hasattr(caught.exception, "records"))
        if field is not None:
            self.assertTrue(any(issue.field == field and (line is None or issue.line_number == line)
                                for issue in caught.exception.issues), str(caught.exception))
        return caught.exception

    def test_original_sample_passes_without_any_byte_change(self):
        source = ROOT / "data" / "sample_orders.csv"
        before = hashlib.sha256(source.read_bytes()).digest()
        records = load_and_validate(source)
        self.assertEqual(len(records), 80)
        self.assertEqual(records[0].source_line, 2)
        self.assertEqual(records[0].unit_price_cents, 1990)
        self.assertEqual(records[0].order_time.utcoffset().total_seconds(), 8 * 3600)
        self.assertEqual(hashlib.sha256(source.read_bytes()).digest(), before)

    def test_utf8_with_and_without_bom_are_supported(self):
        for encoding in ("utf-8", "utf-8-sig"):
            with self.subTest(encoding=encoding):
                self.write(encoding=encoding)
                self.assertEqual(len(load_and_validate(self.path)), 2)

    def test_header_order_can_change_without_misreading_values(self):
        self.write(headers=tuple(reversed(REQUIRED_FIELDS)))
        records = load_and_validate(self.path)
        self.assertEqual(records[0].product_name, "学习文具")
        self.assertEqual(records[0].payment_amount_cents, 3582)

    def test_reader_preserves_text_and_physical_line_numbers(self):
        self.write([sample_row(product_name="学习,文具"), second_row()])
        records = load_and_validate(self.path)
        self.assertEqual(records[0].product_name, "学习,文具")
        self.assertEqual([record.source_line for record in records], [2, 3])

    def test_missing_required_header(self):
        self.write(headers=tuple(field for field in REQUIRED_FIELDS if field != "refund_quantity"))
        error = self.rejected("表头", 1)
        self.assertIn("refund_quantity", str(error))

    def test_duplicate_header(self):
        self.write(headers=REQUIRED_FIELDS + ("quantity",))
        self.assertIn("重复", str(self.rejected("表头", 1)))

    def test_unknown_or_empty_or_padded_header(self):
        for field in ("unexpected", "", " order_id"):
            with self.subTest(field=field):
                self.write(headers=(field,) + REQUIRED_FIELDS[1:])
                self.rejected("表头", 1)

    def test_empty_file_and_header_only_are_rejected(self):
        for content in (b"", b"\xef\xbb\xbf"):
            self.path.write_bytes(content)
            self.rejected("表头", 1)
        self.write([])
        self.rejected("明细")

    def test_too_few_or_too_many_columns(self):
        for suffix in (",extra",):
            self.write()
            lines = self.path.read_text(encoding="utf-8-sig").splitlines()
            lines[1] += suffix
            self.path.write_text("\n".join(lines), encoding="utf-8-sig")
            self.rejected("列数", 2)
        self.write()
        lines = self.path.read_text(encoding="utf-8-sig").splitlines()
        lines[2] = lines[2].rsplit(",", 1)[0]
        self.path.write_text("\n".join(lines), encoding="utf-8-sig")
        self.rejected("列数", 3)

    def test_blank_record_is_not_silently_skipped(self):
        self.write()
        text = self.path.read_text(encoding="utf-8-sig")
        self.path.write_text(text.replace("\n", "\n\n", 1), encoding="utf-8-sig")
        self.rejected("明细", 2)

    def test_malformed_quoted_csv(self):
        self.path.write_text(",".join(REQUIRED_FIELDS) + '\n"未闭合的字段\n', encoding="utf-8-sig")
        self.rejected("CSV 格式", 2)

    def test_wrong_encoding_reports_its_line(self):
        self.write(encoding="gbk")
        self.rejected("编码", 2)

    def test_missing_file_directory_and_permission_error(self):
        with self.assertRaises(CsvValidationError):
            read_csv(self.folder / "不存在.csv")
        with self.assertRaises(CsvValidationError):
            read_csv(self.folder)
        with self.assertRaises(CsvValidationError):
            read_csv("\x00.csv")
        with patch.object(Path, "read_bytes", side_effect=PermissionError("模拟拒绝读取")):
            with self.assertRaises(CsvValidationError) as caught:
                read_csv(self.path)
        self.assertIn("模拟拒绝读取", str(caught.exception))

    def test_every_required_field_rejects_empty_value(self):
        for field in REQUIRED_FIELDS:
            with self.subTest(field=field):
                self.write([sample_row(**{field: ""})])
                self.rejected(field, 2)

    def test_whitespace_is_rejected_instead_of_normalized(self):
        for field, value in (("order_id", " A001"), ("product_name", "学习文具 "),
                             ("payment_amount", "35.82 "), ("quantity", " ")):
            with self.subTest(field=field):
                self.write([sample_row(**{field: value})])
                self.rejected(field, 2)

    def test_control_characters_and_multiline_values_have_useful_locations(self):
        self.write([sample_row(product_name="第一行\n第二行"), second_row(quantity="0")])
        error = self.rejected("product_name", 2)
        self.assertTrue(any(issue.line_number == 4 and issue.field == "quantity" for issue in error.issues))
        self.assertIn("2–3", str(error))
        for value in ("学习\t文具", "学习\x00文具", "学习\x7f文具"):
            self.write([sample_row(product_name=value)])
            self.rejected("product_name", 2)

    def test_duplicate_detail_id_references_first_occurrence(self):
        self.write([sample_row(), second_row(order_line_id="A001-01")])
        error = self.rejected("order_line_id", 3)
        self.assertIn("首次出现在第 2 行", str(error))

    def test_same_order_and_product_can_have_distinct_detail_ids(self):
        self.write([sample_row(), sample_row(order_line_id="A001-02")])
        self.assertEqual(len(load_and_validate(self.path)), 2)

    def test_invalid_quantity_and_refund_quantity_types(self):
        for field, values in (("quantity", ["0", "-1", "1.5", "2.0", "2e1", "２", "01"]),
                              ("refund_quantity", ["-1", "1.5", "０", "00"])):
            for value in values:
                with self.subTest(field=field, value=value):
                    self.write([sample_row(**{field: value})])
                    self.rejected(field, 2)

    def test_positive_quantity_is_not_limited_to_sample_four(self):
        self.write([sample_row(quantity="5", payment_amount="95.52")])
        self.assertEqual(load_and_validate(self.path)[0].quantity, 5)

    def test_invalid_money_formats_for_each_money_column(self):
        for field in ("unit_price", "discount_amount", "payment_amount", "refund_amount"):
            for value in ("-1.00", "1", "1.0", "1.000", "1e2", "NaN", "Infinity", "1,000.00", ".50", "+1.00", "01.00", "１.00"):
                with self.subTest(field=field, value=value):
                    self.write([sample_row(**{field: value})])
                    self.rejected(field, 2)

    def test_large_money_keeps_exact_cents_without_float_or_decimal_rounding(self):
        amount = "123456789012345678901234567890.12"
        self.write([sample_row(quantity="1", unit_price=amount, discount_amount="0.00", payment_amount=amount)])
        record = load_and_validate(self.path)[0]
        self.assertEqual(record.unit_price_cents, 12345678901234567890123456789012)
        self.assertEqual(record.payment_amount_cents, record.unit_price_cents)

    def test_integer_parser_failure_is_a_field_error(self):
        self.write([sample_row(quantity="9" * 5000)])
        self.rejected("quantity", 2)

    def test_zero_price_and_excessive_discount(self):
        self.write([sample_row(unit_price="0.00", discount_amount="0.00", payment_amount="0.00")])
        self.rejected("unit_price", 2)
        self.write([sample_row(discount_amount="39.81")])
        self.rejected("discount_amount", 2)

    def test_payment_must_match_quantity_price_and_discount(self):
        self.write([sample_row(payment_amount="35.81")])
        self.rejected("payment_amount", 2)

    def test_refund_amount_and_quantity_cannot_exceed_payment_and_quantity(self):
        self.write([sample_row(order_status="部分退款", refund_amount="35.83")])
        self.rejected("refund_amount", 2)
        self.write([sample_row(order_status="部分退款", refund_amount="5.00", refund_quantity="3")])
        self.rejected("refund_quantity", 2)

    def test_refund_units_require_nonzero_refund_amount(self):
        self.write([sample_row(order_status="部分退款", refund_quantity="1")])
        self.rejected("refund_quantity", 2)

    def test_unknown_state_category_and_channel_are_rejected(self):
        for field, value in (("order_status", "退款申请中"), ("category", "食品"), ("channel", "未知渠道")):
            self.write([sample_row(**{field: value})])
            self.rejected(field, 2)

    def test_invalid_date_formats_and_impossible_calendar_dates(self):
        for value in ("2026-2-01 10:00:00", "2026/10/01 10:00:00", "2026-02-30 10:00:00",
                      "2026-10-01 24:00:00", "2026-10-01", "0000-10-01 10:00:00", "2026-10-01 10:00:00Z"):
            with self.subTest(value=value):
                self.write([sample_row(order_time=value)])
                self.rejected("order_time", 2)

    def test_valid_leap_day_and_other_ids_are_not_sample_date_locked(self):
        self.write([sample_row(order_id="OTHER", order_line_id="ROW-1", order_time="2024-02-29 23:59:59")])
        self.assertEqual(load_and_validate(self.path)[0].order_time.year, 2024)

    def test_unpaid_and_cancelled_orders_need_zero_payment_and_refunds(self):
        for status in ("未支付", "已取消"):
            self.write([sample_row(order_status=status, payment_amount="0.00")])
            self.assertEqual(len(load_and_validate(self.path)), 1)
            self.write([sample_row(order_status=status)])
            self.rejected("payment_amount", 2)
            self.write([sample_row(order_status=status, payment_amount="0.00", refund_amount="0.01")])
            self.rejected("refund_amount", 2)

    def test_paid_without_refund_state_rejects_any_refund(self):
        self.write([sample_row(refund_amount="1.00")])
        self.rejected("order_status", 2)

    def test_fully_refunded_orders_require_both_full_amount_and_units(self):
        self.write([sample_row(order_status="已全额退款", refund_amount="35.82", refund_quantity="2")])
        self.assertEqual(len(load_and_validate(self.path)), 1)
        self.write([sample_row(order_status="已全额退款", refund_amount="30.00", refund_quantity="2")])
        self.rejected("refund_amount", 2)
        self.write([sample_row(order_status="已全额退款", refund_amount="35.82", refund_quantity="1")])
        self.rejected("refund_quantity", 2)

    def test_partial_refund_zero_or_full_order_total_is_rejected(self):
        self.write([sample_row(order_status="部分退款"), second_row(order_status="部分退款")])
        self.rejected("order_status", 2)
        self.write([sample_row(order_status="部分退款", refund_amount="35.82", refund_quantity="2"),
                    second_row(order_status="部分退款", refund_amount="11.25", refund_quantity="1")])
        self.rejected("order_status", 2)

    def test_partial_order_may_include_a_fully_refunded_line(self):
        self.write([sample_row(order_status="部分退款", refund_amount="35.82", refund_quantity="2"),
                    second_row(order_status="部分退款")])
        self.assertEqual(len(load_and_validate(self.path)), 2)

    def test_partial_refund_with_or_without_refunded_units_is_valid(self):
        for quantity in ("0", "1"):
            self.write([sample_row(order_status="部分退款", refund_amount="5.00", refund_quantity=quantity)])
            self.assertEqual(load_and_validate(self.path)[0].refund_quantity, int(quantity))

    def test_order_time_channel_and_state_must_match_first_row(self):
        for changes, field in (({"order_time": "2026-10-02 10:07:11"}, "order_time"),
                                ({"channel": "社群推广"}, "channel"),
                                ({"order_status": "未支付", "payment_amount": "0.00"}, "order_status")):
            self.write([sample_row(), second_row(**changes)])
            self.assertIn("第 2 行", str(self.rejected(field, 3)))

    def test_same_product_name_category_and_v1_price_must_be_consistent(self):
        for changes, field in (({"product_name": "另一名称"}, "product_name"),
                                ({"category": "日用生活"}, "category"),
                                ({"unit_price": "20.00", "payment_amount": "36.02"}, "unit_price")):
            self.write([sample_row(), sample_row(order_id="B001", order_line_id="B001-01", **changes)])
            self.rejected(field, 3)

    def test_multiple_bad_rows_report_each_location_and_no_partial_records(self):
        self.write([sample_row(quantity="0"), second_row(payment_amount="bad")])
        error = self.rejected()
        self.assertTrue(any(issue.line_number == 2 and issue.field == "quantity" for issue in error.issues))
        self.assertTrue(any(issue.line_number == 3 and issue.field == "payment_amount" for issue in error.issues))

    def test_empty_row_list_is_not_a_successful_validation(self):
        with self.assertRaises(CsvValidationError):
            validate_rows([])

    def test_cli_success_and_failure_exit_codes_and_messages(self):
        self.write()
        output = io.StringIO()
        with contextlib.redirect_stdout(output):
            self.assertEqual(main([str(self.path)]), 0)
        self.assertIn("读取 2 条商品明细", output.getvalue())
        self.write([sample_row(quantity="0")])
        output = io.StringIO()
        with contextlib.redirect_stdout(output):
            self.assertEqual(main([str(self.path)]), 1)
        self.assertIn("第 2 行 [quantity]", output.getvalue())
        self.assertIn("未返回部分数据", output.getvalue())
        self.assertNotIn("校验通过：", output.getvalue())

    def test_cli_default_sample_is_independent_of_working_directory(self):
        result = subprocess.run([sys.executable, "-B", "-X", "utf8", str(ROOT / "validate_orders.py")],
                                cwd=self.folder, capture_output=True, text=True, encoding="utf-8", timeout=20)
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertIn("读取 80 条商品明细", result.stdout)
        self.assertEqual(list(self.folder.iterdir()), [])

    def test_imports_do_not_read_data_or_write_files(self):
        script = '''from unittest.mock import patch
with patch("pathlib.Path.read_bytes", side_effect=AssertionError("unexpected read")), patch("pathlib.Path.write_bytes", side_effect=AssertionError("unexpected write")), patch("pathlib.Path.write_text", side_effect=AssertionError("unexpected write")):
    import csv_reader, order_validator, validate_orders
print("IMPORT_SAFE")
'''
        result = subprocess.run([sys.executable, "-B", "-X", "utf8", "-c", script], cwd=ROOT,
                                capture_output=True, text=True, timeout=20)
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertIn("IMPORT_SAFE", result.stdout)


if __name__ == "__main__":
    unittest.main(verbosity=2)
