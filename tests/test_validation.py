import tempfile
import unittest
from pathlib import Path

from docx import Document

from word_to_twine import Para, Passage, Story, load_docx, validate


class TestValidation(unittest.TestCase):
    def test_validate_clean_story(self):
        story = Story(
            title="Clean Story",
            start="Intro",
            passages=[
                Passage("Intro", [Para("[[Continue->Next]]", "Normal")]),
                Passage("Next", [Para("The end.", "Normal")]),
            ],
        )
        errors = validate(story)
        self.assertEqual(errors, [])

    def test_validate_duplicate_passage_titles(self):
        story = Story(
            title="Duplicate Story",
            start="Intro",
            passages=[
                Passage("Intro", []),
                Passage("Intro", []),
            ],
        )
        errors = validate(story)
        self.assertIn("Duplicate Heading 2 passage names found.", errors)

    def test_validate_missing_start_passage(self):
        story = Story(
            title="Missing Start",
            start="NonExistentStart",
            passages=[
                Passage("Intro", []),
            ],
        )
        errors = validate(story)
        self.assertTrue(any("StartPassage 'NonExistentStart' is not a Heading 2 passage" in e for e in errors))

    def test_validate_broken_wiki_link(self):
        story = Story(
            title="Broken Link",
            start="Intro",
            passages=[
                Passage("Intro", [Para("Click here [[Go to room->SecretRoom]]", "Normal")]),
            ],
        )
        errors = validate(story)
        self.assertTrue(any("Broken link in 'Intro': target 'SecretRoom' does not exist" in e for e in errors))

    def test_validate_broken_choice_target(self):
        story = Story(
            title="Broken Choice",
            start="Intro",
            passages=[
                Passage("Intro", [Para("Choice :: Open door -> Hallway", "Normal")]),
            ],
        )
        errors = validate(story)
        self.assertTrue(any("Broken Choice target in 'Intro': 'Hallway' does not exist" in e for e in errors))

    def test_load_docx_no_heading_2_raises_value_error(self):
        with tempfile.TemporaryDirectory() as td:
            doc_path = Path(td) / "empty.docx"
            doc = Document()
            doc.add_heading("Just a Title", level=1)
            doc.add_paragraph("No passage headings here.")
            doc.save(doc_path)

            with self.assertRaises(ValueError) as ctx:
                load_docx(doc_path)
            self.assertIn("No Heading 2 passages found", str(ctx.exception))


if __name__ == "__main__":
    unittest.main()

