import unittest

from word_to_twine import (
    Para,
    Passage,
    Story,
    compile_html,
    render_passage,
    render_plain_para,
    story_to_twee,
)


class TestPassageRenderer(unittest.TestCase):
    def test_render_plain_para_heading3(self):
        p1 = Para("Section Subheading", "Heading 3")
        self.assertEqual(render_plain_para(p1), "<h3>Section Subheading</h3>")

        p2 = Para("### Markdown Style Subheading", "Normal")
        self.assertEqual(render_plain_para(p2), "<h3>Markdown Style Subheading</h3>")

    def test_render_plain_para_author_note(self):
        p = Para("Author note text here", "Author Note")
        self.assertEqual(render_plain_para(p), '<div class="author-note">Author note text here</div>')

    def test_render_plain_para_normal_and_empty(self):
        p = Para("Ordinary learner text", "Normal")
        self.assertEqual(render_plain_para(p), "<p>Ordinary learner text</p>")

        p_empty = Para("   ", "Normal")
        self.assertEqual(render_plain_para(p_empty), "")

    def test_render_passage_lists(self):
        passage = Passage(
            title="ListTest",
            items=[
                Para("First bullet", "List Bullet"),
                Para("Second bullet", "List Bullet"),
                Para("First numbered", "List Number"),
                Para("Second numbered", "List Number"),
            ],
        )
        rendered = render_passage(passage)
        self.assertIn('<ul class="scenario-list"><li>First bullet</li><li>Second bullet</li></ul>', rendered)
        self.assertIn('<ol class="scenario-list"><li>First numbered</li><li>Second numbered</li></ol>', rendered)

    def test_render_passage_single_choice_with_onchoose(self):
        passage = Passage(
            title="ChoiceTest",
            items=[
                Para("Choice :: Investigate records -> Investigation", "Normal"),
                Para("Set :: investigated=true", "Normal"),
                Para("Add :: score += 5", "Normal"),
            ],
        )
        rendered = render_passage(passage)
        expected = (
            '<div class="scenario-action">'
            '[[Investigate records->Investigation][$investigated=true;  $score += 5]]'
            '</div>'
        )
        self.assertEqual(rendered, expected)

    def test_render_passage_multichoice_and_gate(self):
        passage = Passage(
            title="MultiTest",
            items=[
                Para("CostRule :: parallel-max-selected", "Normal"),
                Para("Gate :: $unlocked", "Normal"),
                Para("GateFailureTarget :: LockedScreen", "Normal"),
                Para("MultiChoice :: Option A -> NextScreen | cost=2", "Normal"),
                Para("MultiChoice :: Option B -> NextScreen | cost=5", "Normal"),
            ],
        )
        rendered = render_passage(passage)
        self.assertIn('class="scenario-multi"', rendered)
        self.assertIn('data-choice="option_a"', rendered)
        self.assertIn('data-choice="option_b"', rendered)
        self.assertIn('onclick="ScenarioEngine.submitMulti', rendered)
        self.assertIn('parallel-max-selected', rendered)
        self.assertIn('LockedScreen', rendered)

    def test_render_passage_conditionals(self):
        passage = Passage(
            title="CondTest",
            items=[
                Para("If :: score > 10 AND completed", "Normal"),
                Para("High score achieved!", "Normal"),
                Para("ElseIf :: score == 10", "Normal"),
                Para("Mid score achieved!", "Normal"),
                Para("Else", "Normal"),
                Para("Low score achieved!", "Normal"),
                Para("EndIf", "Normal"),
            ],
        )
        rendered = render_passage(passage)
        self.assertIn("<<if $score > 10 && $completed>>", rendered)
        self.assertIn("<p>High score achieved!</p>", rendered)
        self.assertIn("<<elseif $score == 10>>", rendered)
        self.assertIn("<p>Mid score achieved!</p>", rendered)
        self.assertIn("<<else>>", rendered)
        self.assertIn("<p>Low score achieved!</p>", rendered)
        self.assertIn("<</if>>", rendered)

    def test_render_passage_table_and_rows(self):
        passage = Passage(
            title="TableTest",
            items=[
                Para("Table :: headers=Summary", "Normal"),
                Para("Row :: 1 | Click start | Game begins", "Normal"),
                Para("Row :: 2 | Answer question | Points earned", "Normal"),
            ],
        )
        rendered = render_passage(passage)
        self.assertIn('<table class="scenario-table">', rendered)
        self.assertIn('<th>Summary</th>', rendered)
        self.assertIn('<td>1</td><td>Click start</td><td>Game begins</td>', rendered)
        self.assertIn('<td>2</td><td>Answer question</td><td>Points earned</td>', rendered)
        self.assertIn('</tbody></table>', rendered)

    def test_render_passage_chart(self):
        passage = Passage(
            title="ChartTest",
            items=[
                Para("Chart :: data=10,25,50 | x=Month | y=Progress", "Normal"),
            ],
        )
        rendered = render_passage(passage)
        self.assertIn('<div class="scenario-chart"', rendered)
        self.assertIn('aria-label="Progress by Month"', rendered)
        self.assertIn('<div class="scenario-bar"', rendered)

    def test_render_passage_outcomes(self):
        passage = Passage(
            title="OutcomeTest",
            items=[
                Para("Outcome :: Great Success | when=solved AND verified", "Normal"),
                Para("You did amazing.", "Normal"),
                Para("Outcome :: Partial Success | when=solved", "Normal"),
                Para("You solved part of it.", "Normal"),
                Para("Outcome :: title=Failure | default", "Normal"),
                Para("Try again.", "Normal"),
            ],
        )
        rendered = render_passage(passage)
        self.assertIn("<<if $solved && $verified>>", rendered)
        self.assertIn("<h3>Great Success</h3>", rendered)
        self.assertIn("<<elseif $solved>>", rendered)
        self.assertIn("<h3>Partial Success</h3>", rendered)
        self.assertIn("<<else>>", rendered)
        self.assertIn("<h3>Failure</h3>", rendered)
        self.assertIn("<</if>>", rendered)

    def test_render_passage_calculations_and_clamp(self):
        passage = Passage(
            title="MathTest",
            items=[
                Para("Calculate :: total = ceil(base * 1.5)", "Normal"),
                Para("Clamp :: total between 0 and 100", "Normal"),
                Para("ScoreDisplay :: Total Points | value=total | max=100", "Normal"),
                Para("Feedback :: good | Excellent effort!", "Normal"),
            ],
        )
        rendered = render_passage(passage)
        self.assertIn("<<set $total = Math.ceil($base * 1.5)>>", rendered)
        self.assertIn("<<set $total = Math.max(0, Math.min(100, $total))>>", rendered)
        self.assertIn('<div class="score-row"><strong>Total Points:</strong> <<print $total>> / 100</div>', rendered)
        self.assertIn('ScenarioEngine.feedback("good", "Excellent effort!");', rendered)

    def test_story_to_twee_and_compile_html(self):
        story = Story(
            title="Demo Scenario",
            start="StartPassage",
            passages=[
                Passage(title="StartPassage", items=[Para("Welcome to the demo.", "Normal")]),
                Passage(title="EndPassage", items=[Para("The end.", "Normal")]),
            ],
        )
        twee = story_to_twee(story)
        self.assertIn(":: StoryTitle\nDemo Scenario", twee)
        self.assertIn(":: StoryInit [script]", twee)
        self.assertIn(":: StartPassage", twee)
        self.assertIn(":: EndPassage", twee)

        html = compile_html(story)
        self.assertIn("<title>Demo Scenario</title>", html)
        self.assertIn('<tw-storydata name="Demo Scenario"', html)
        self.assertIn('name="StartPassage"', html)
        self.assertIn('name="EndPassage"', html)
        self.assertIn("window.ScenarioEngine =", html)


if __name__ == "__main__":
    unittest.main()
