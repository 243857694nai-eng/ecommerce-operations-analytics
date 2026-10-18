"""字段映射纯函数测试：不读写订单或创建工作簿。"""

import unittest
from dataclasses import FrozenInstanceError, replace

from csv_reader import REQUIRED_FIELDS
from field_mapping import ALIASES, MappingError, analyze_headers, confirm_mapping, resolve_mapping


class FieldMappingTests(unittest.TestCase):
    def test_english_standard_fields_match_exactly(self):
        self.assertEqual(resolve_mapping(REQUIRED_FIELDS).assignments, REQUIRED_FIELDS)

    def test_all_explicit_chinese_aliases(self):
        for field, aliases in ALIASES.items():
            for alias in aliases:
                with self.subTest(field=field, alias=alias):
                    headers = tuple(alias if item == field else item for item in REQUIRED_FIELDS)
                    self.assertEqual(resolve_mapping(headers).assignments, REQUIRED_FIELDS)

    def test_reordered_headers_preserve_column_positions(self):
        headers = tuple(reversed(REQUIRED_FIELDS))
        self.assertEqual(resolve_mapping(headers).assignments, headers)

    def test_only_outer_header_whitespace_is_normalized(self):
        headers = tuple(" " + item + "\t" for item in REQUIRED_FIELDS)
        self.assertEqual(resolve_mapping(headers).assignments, REQUIRED_FIELDS)
        for title in ("Order_ID", "订单 编号", "销售额", "商品单价(元)"):
            self.assertEqual(analyze_headers((title,))[0].candidates, ())

    def test_duplicate_headers_are_not_lost_in_a_dictionary(self):
        matches = analyze_headers(REQUIRED_FIELDS + ("order_id",))
        self.assertEqual([item.column for item in matches if item.duplicate], [1, 15])
        with self.assertRaisesRegex(MappingError, "明确映射"):
            resolve_mapping(REQUIRED_FIELDS + ("order_id",))

    def test_duplicate_headers_can_be_explicitly_selected_and_ignored(self):
        plan = resolve_mapping(REQUIRED_FIELDS + ("order_id",), {1: "order_id", 15: None})
        self.assertEqual(plan.assignments, REQUIRED_FIELDS + (None,))

    def test_two_distinct_aliases_for_one_field_require_resolution(self):
        headers = REQUIRED_FIELDS + ("订单号",)
        with self.assertRaisesRegex(MappingError, "多个源列.*order_id"):
            resolve_mapping(headers)
        self.assertIsNone(resolve_mapping(headers, {15: None}).assignments[-1])

    def test_missing_refund_quantity_cannot_be_filled_by_confirmation(self):
        with self.assertRaisesRegex(MappingError, "refund_quantity"):
            resolve_mapping(tuple(field for field in REQUIRED_FIELDS if field != "refund_quantity"))

    def test_extra_customer_column_needs_explicit_ignore(self):
        headers = REQUIRED_FIELDS + ("客户资料",)
        with self.assertRaisesRegex(MappingError, "第 15 列"):
            resolve_mapping(headers)
        self.assertIsNone(resolve_mapping(headers, {15: None}).assignments[-1])

    def test_unknown_header_can_be_explicitly_mapped_without_fuzzy_guess(self):
        headers = ("外部订单标识",) + REQUIRED_FIELDS[1:]
        with self.assertRaises(MappingError):
            resolve_mapping(headers)
        self.assertEqual(resolve_mapping(headers, {1: "order_id"}).assignments, REQUIRED_FIELDS)

    def test_ambiguous_money_requires_separate_meaning_confirmation(self):
        for title, field in (("实付金额", "payment_amount"), ("订单优惠", "discount_amount"),
                             ("退款金额", "refund_amount")):
            with self.subTest(title=title):
                column = REQUIRED_FIELDS.index(field) + 1
                headers = tuple(title if item == field else item for item in REQUIRED_FIELDS)
                with self.assertRaisesRegex(MappingError, "金额含义未确认"):
                    resolve_mapping(headers)
                self.assertEqual(resolve_mapping(headers, confirmed_meanings={column}).assignments, REQUIRED_FIELDS)

    def test_manual_unknown_money_mapping_also_needs_meaning_confirmation(self):
        headers = tuple("未知金额" if field == "payment_amount" else field for field in REQUIRED_FIELDS)
        with self.assertRaisesRegex(MappingError, "金额含义未确认"):
            resolve_mapping(headers, {10: "payment_amount"})
        self.assertEqual(resolve_mapping(headers, {10: "payment_amount"}, confirmed_meanings={10}).assignments,
                         REQUIRED_FIELDS)

    def test_unknown_standard_target_is_rejected(self):
        with self.assertRaisesRegex(MappingError, "未知标准字段"):
            resolve_mapping(REQUIRED_FIELDS, {1: "customer_id"})

    def test_ignored_required_column_is_still_missing(self):
        with self.assertRaisesRegex(MappingError, "缺少必需字段.*order_id"):
            resolve_mapping(REQUIRED_FIELDS, {1: None})

    def test_invalid_column_positions_are_rejected(self):
        for column in (0, -1, 15, "1", True):
            with self.subTest(column=column), self.assertRaises(MappingError):
                resolve_mapping(REQUIRED_FIELDS, {column: "order_id"})

    def test_non_text_empty_and_blank_headers_are_rejected(self):
        for title in (None, 1, True, "", " "):
            with self.subTest(title=title), self.assertRaises(MappingError):
                analyze_headers((title,))

    def test_default_plan_is_unconfirmed_and_context_is_preserved(self):
        plan = resolve_mapping(REQUIRED_FIELDS, context="snapshot/sheet/header")
        self.assertFalse(plan.confirmed)
        self.assertEqual(confirm_mapping(plan, True), replace(plan, confirmed=True))
        self.assertEqual(plan.context, "snapshot/sheet/header")

    def test_cancellation_and_truthy_values_are_not_explicit_confirmation(self):
        for approved in (False, None, 1, "确认"):
            with self.subTest(approved=approved), self.assertRaisesRegex(MappingError, "未明确确认"):
                confirm_mapping(resolve_mapping(REQUIRED_FIELDS), approved)

    def test_confirmed_mapping_is_immutable(self):
        plan = confirm_mapping(resolve_mapping(REQUIRED_FIELDS), True)
        with self.assertRaises(FrozenInstanceError):
            plan.confirmed = False


if __name__ == "__main__":
    unittest.main()
