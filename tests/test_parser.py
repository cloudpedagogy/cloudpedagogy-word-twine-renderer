import unittest
from types import SimpleNamespace

from word_to_twine import (
    STYLE_DIRECTIVE_MAP,
    choice_to_markup,
    para_from_docx,
    parse_kv,
    setter_expr,
    style_name,
    sugar_expr,
)


class TestParser(unittest.TestCase):
    def test_parse_kv_empty_and_simple(self):
        head, kv = parse_kv("")
        self.assertIsNone(head)
        self.assertEqual(kv, {})

        head, kv = parse_kv("Just a head")
        self.assertEqual(head, "Just a head")
        self.assertEqual(kv, {})

        head, kv = parse_kv("Head | key1=val1 | key2 = val2")
        self.assertEqual(head, "Head")
        self.assertEqual(kv, {"key1": "val1", "key2": "val2"})

    def test_sugar_expr_logical_operators(self):
        # AND -> &&, OR -> ||, NOT -> !
        self.assertEqual(sugar_expr("a AND b"), "$a && $b")
        self.assertEqual(sugar_expr("a OR b"), "$a || $b")
        self.assertEqual(sugar_expr("NOT a"), "!$a")
        self.assertEqual(sugar_expr("a and b or not c"), "$a && $b || !$c")

    def test_sugar_expr_keywords_and_literals(self):
        # Keywords and numbers must not be prefixed with $
        self.assertEqual(sugar_expr("true AND false AND null"), "true && false && null")
        self.assertEqual(sugar_expr("score >= 10"), "$score >= 10")
        self.assertEqual(sugar_expr("count == 3.14"), "$count == 3.14")
        self.assertEqual(sugar_expr("Math.max(score, 5)"), "Math.max($score, 5)")

    def test_sugar_expr_quoted_strings(self):
        # Quoted string contents should not be altered
        expr = "role == 'admin' AND status == \"active\""
        self.assertEqual(sugar_expr(expr), "$role == 'admin' && $status == \"active\"")

    def test_setter_expr(self):
        self.assertEqual(setter_expr("score += 2; level = 1"), "$score += 2;  $level = 1")

    def test_choice_to_markup_simplified(self):
        cid, label, target, typ, cost = choice_to_markup("Review the evidence -> Evidence")
        self.assertEqual(label, "Review the evidence")
        self.assertEqual(target, "Evidence")
        self.assertEqual(typ, "single")
        self.assertEqual(cid, "review_the_evidence")
        self.assertEqual(cost, "")

        cid, label, target, typ, cost = choice_to_markup("Check records -> Results | cost=4", forced_type="multi")
        self.assertEqual(label, "Check records")
        self.assertEqual(target, "Results")
        self.assertEqual(typ, "multi")
        self.assertEqual(cid, "check_records")
        self.assertEqual(cost, "4")

    def test_choice_to_markup_legacy(self):
        payload = "id=opt1 | type=multi | label=Check records | target=Results | cost=4"
        cid, label, target, typ, cost = choice_to_markup(payload)
        self.assertEqual(cid, "opt1")
        self.assertEqual(label, "Check records")
        self.assertEqual(target, "Results")
        self.assertEqual(typ, "multi")
        self.assertEqual(cost, "4")

    def test_style_name(self):
        class MockStyle:
            name = "Scenario Choice"

        p = SimpleNamespace(style=MockStyle())
        self.assertEqual(style_name(p), "Scenario Choice")

        p_no_name = SimpleNamespace(style=SimpleNamespace(name=None))
        self.assertEqual(style_name(p_no_name), "Normal")

        p_err = SimpleNamespace()
        self.assertEqual(style_name(p_err), "Normal")

    def test_para_from_docx_empty(self):
        p = SimpleNamespace(text="   ", style=SimpleNamespace(name="Normal"))
        self.assertIsNone(para_from_docx(p))

    def test_para_from_docx_mapped_styles(self):
        for style, expected_directive in STYLE_DIRECTIVE_MAP.items():
            p = SimpleNamespace(text="Test payload", style=SimpleNamespace(name=style))
            converted = para_from_docx(p)
            self.assertIsNotNone(converted)
            self.assertEqual(converted.text, f"{expected_directive} :: Test payload")
            self.assertEqual(converted.style, style)

    def test_para_from_docx_existing_directive_not_duplicated(self):
        p = SimpleNamespace(text="Choice :: Test payload", style=SimpleNamespace(name="Scenario Choice"))
        converted = para_from_docx(p)
        self.assertIsNotNone(converted)
        self.assertEqual(converted.text, "Choice :: Test payload")


if __name__ == "__main__":
    unittest.main()

