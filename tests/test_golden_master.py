import re
import unittest
from pathlib import Path

from word_to_twine import (
    compile_html,
    embedded_template_html,
    load_docx,
    story_to_twee,
    validate,
)


class TestGoldenMaster(unittest.TestCase):
    def setUp(self):
        self.root = Path(__file__).resolve().parents[1]
        self.demo_docx = self.root / "demo" / "research_data_incident_scenario.docx"
        self.literature_docx = self.root / "literature_review_scenario.docx"
        self.ref_twee = self.root / "output" / "twine" / "research_data_incident_scenario.twee"
        self.ref_html = self.root / "output" / "twine" / "research_data_incident_scenario.html"

    def test_research_data_incident_twee_exact_match(self):
        story = load_docx(self.demo_docx)
        errors = validate(story)
        self.assertEqual(errors, [], f"Validation errors on reference demo docx: {errors}")

        generated_twee = story_to_twee(story)
        expected_twee = self.ref_twee.read_text(encoding="utf-8")

        self.assertEqual(
            generated_twee,
            expected_twee,
            "Generated Twee does not match tracked reference output exactly.",
        )

    def test_research_data_incident_html_normalized_ifid_exact_match(self):
        story = load_docx(self.demo_docx)
        generated_html = compile_html(story)
        expected_html = self.ref_html.read_text(encoding="utf-8")

        # The IFID is randomly generated on each compilation via uuid.uuid4().
        # Normalize ONLY the ifid attribute without altering any other HTML, CSS or JS content.
        ifid_pattern = re.compile(r'ifid="[A-Fa-f0-9-]+"')
        norm_generated = ifid_pattern.sub('ifid="NORMALIZED-IFID"', generated_html)
        norm_expected = ifid_pattern.sub('ifid="NORMALIZED-IFID"', expected_html)

        self.assertEqual(
            norm_generated,
            norm_expected,
            "Generated HTML does not match tracked reference HTML after IFID normalization.",
        )

    def test_embedded_sugarcube_runtime_intact(self):
        template = embedded_template_html()
        self.assertIn("SugarCube", template)
        self.assertIn("<tw-storydata", template)
        self.assertIn("twine-user-stylesheet", template)
        self.assertIn("twine-user-script", template)

    def test_literature_review_scenario_conversion(self):
        self.assertTrue(self.literature_docx.is_file())
        story = load_docx(self.literature_docx)
        errors = validate(story)
        self.assertEqual(errors, [], f"Validation errors on literature review scenario: {errors}")

        twee = story_to_twee(story)
        self.assertIn(":: StoryTitle\nLiterature Review Scenario", twee)
        self.assertIn(":: Start", twee)

        html = compile_html(story)
        self.assertIn("<title>Literature Review Scenario</title>", html)
        self.assertIn('name="Start"', html)


if __name__ == "__main__":
    unittest.main()
