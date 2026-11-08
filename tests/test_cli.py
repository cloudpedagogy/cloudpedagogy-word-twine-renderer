import io
import tempfile
import unittest
from contextlib import redirect_stderr, redirect_stdout
from pathlib import Path

from docx import Document

from word_to_twine import main


class TestCLI(unittest.TestCase):
    def test_cli_missing_input_file(self):
        stderr = io.StringIO()
        with redirect_stderr(stderr):
            code = main(["non_existent_file.docx"])
        self.assertEqual(code, 1)
        self.assertIn("not found", stderr.getvalue())

    def test_cli_non_docx_file(self):
        with tempfile.NamedTemporaryFile(suffix=".txt") as tmp:
            stderr = io.StringIO()
            with redirect_stderr(stderr):
                code = main([tmp.name])
            self.assertEqual(code, 1)
            self.assertIn("must be a .docx file", stderr.getvalue())

    def test_cli_invalid_docx_content(self):
        with tempfile.NamedTemporaryFile(suffix=".docx") as tmp:
            tmp.write(b"Not a real docx zip archive")
            tmp.flush()
            stderr = io.StringIO()
            with redirect_stderr(stderr):
                code = main([tmp.name])
            self.assertEqual(code, 1)
            self.assertIn("Error loading document", stderr.getvalue())

    def test_cli_validation_failure(self):
        with tempfile.TemporaryDirectory() as td:
            doc_path = Path(td) / "broken.docx"
            doc = Document()
            doc.add_heading("Broken Scenario", level=1)
            doc.add_heading("Screen One", level=2)
            doc.add_paragraph("[[Link to Nowhere->MissingScreen]]")
            doc.save(doc_path)

            stderr = io.StringIO()
            with redirect_stderr(stderr):
                code = main([str(doc_path)])
            self.assertEqual(code, 2)
            self.assertIn("Validation failed:", stderr.getvalue())
            self.assertIn("MissingScreen", stderr.getvalue())

    def test_cli_successful_conversion(self):
        fixture = Path("demo/research_data_incident_scenario.docx")
        with tempfile.TemporaryDirectory() as td:
            out_dir = Path(td) / "output"
            stdout = io.StringIO()
            with redirect_stdout(stdout):
                code = main([str(fixture), "--output-dir", str(out_dir)])
            self.assertEqual(code, 0)
            self.assertTrue((out_dir / "research_data_incident_scenario.twee").is_file())
            self.assertTrue((out_dir / "research_data_incident_scenario.html").is_file())

    def test_cli_twee_only_flag(self):
        fixture = Path("demo/research_data_incident_scenario.docx")
        with tempfile.TemporaryDirectory() as td:
            out_dir = Path(td) / "output"
            stdout = io.StringIO()
            with redirect_stdout(stdout):
                code = main([str(fixture), "--output-dir", str(out_dir), "--twee-only"])
            self.assertEqual(code, 0)
            self.assertTrue((out_dir / "research_data_incident_scenario.twee").is_file())
            self.assertFalse((out_dir / "research_data_incident_scenario.html").is_file())


if __name__ == "__main__":
    unittest.main()

