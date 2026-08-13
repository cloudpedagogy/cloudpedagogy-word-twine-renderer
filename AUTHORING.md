# Word-to-Twine Scenario Authoring Manual

This guide explains how to create interactive branching scenarios in
Microsoft Word and convert them to Twee and standalone SugarCube HTML
with `word_to_twine.py`.

The authoring model is generic. It can be used for clinical, public
health, ethics, research, policy, management, professional-practice, and
other decision-based learning scenarios.

> **Recommended workflow:** keep the Word document as the source of
> truth. Generate Twee and HTML from Word. Twine is optional and is best
> used for visualisation, QA, and debugging.

## 1. Authoring model

The system has three layers. Most authors only need the first.

  -----------------------------------------------------------------------
  Level                   Authoring features      Typical use
  ----------------------- ----------------------- -----------------------
  Basic                   Word headings,          Ordinary branching
                          paragraphs, lists,      scenarios
                          Scenario Choice styles  

  Interactive             MultiChoice, Cost, Set, Remember decisions,
                          Add, Feedback           scores, resources,
                                                  multi-select tasks

  Advanced                If, Gate, Calculate,    Complex simulations,
                          Outcome and related     conditional evidence,
                          directives              calculations,
                                                  personalised endings
  -----------------------------------------------------------------------

## 2. Minimal Word structure

A basic scenario can be created almost entirely with normal Word
structure:

-   **Heading 1** --- scenario title.
-   **Heading 2** --- new scenario screen / Twine passage.
-   **Heading 3** --- subheading within a passage.
-   **Normal** --- learner-facing paragraphs.
-   Word bullets and numbered lists --- learner-facing lists.
-   **Scenario Choice** --- a single branching decision.
-   **Scenario MultiChoice** --- an option in a multiple-selection
    decision.

Example:

``` text
Heading 1: Research Ethics Scenario

Heading 2: Start

A researcher asks to use identifiable participant data for a new purpose.

[Scenario Choice]
Review the consent documentation -> Consent Review

[Scenario Choice]
Approve immediately -> Immediate Approval

Heading 2: Consent Review

The original consent form does not clearly cover the proposed secondary use.

[Scenario Choice]
Request further review -> Governance Review
```

The text after `->` must exactly match the destination Heading 2.

## 3. Scenario Word styles

  ----------------------------------------------------------------------------------------------
  Word style              Meaning                     Example paragraph content
  ----------------------- --------------------------- ------------------------------------------
  Scenario Choice         Single decision/navigation  `Review the evidence -> Evidence Review`

  Scenario MultiChoice    Checkbox-style option       `Check records -> Results`

  Scenario Cost           Cost of the preceding       `4`
                          choice                      

  Scenario Set            Set state when the          `reviewed=true`
                          preceding choice is         
                          selected                    

  Scenario Add            Increase or decrease        `score+=3`
                          numeric state               

  Scenario Feedback       Store feedback/debrief      `good \| Good decision.`
                          information                 

  Scenario Condition      Conditional logic           `If :: reviewed`

  Scenario Gate           Progression requirement     `Gate :: reviewed AND evidence`

  Scenario Outcome        Conditional ending          `Outcome :: Successful \| when=...`

  Scenario Advanced       Advanced                    `Calculate :: ...`
  Setting                 configuration/calculation   

  Author Note             Author/developer            Not shown to learners
                          documentation               
  ----------------------------------------------------------------------------------------------

The converter also retains support for the older explicit directive
syntax, so existing scenarios do not have to be rewritten.

## 4. Single choices

Apply the **Scenario Choice** Word style to the paragraph:

``` text
Review the evidence -> Evidence Review
```

The converter creates the internal Twine link automatically. Authors do
not need to create IDs.

## 5. Multiple-selection choices

Apply **Scenario MultiChoice** to each option. Consecutive options
pointing to the same results passage are rendered as a checkbox group.

``` text
[Scenario MultiChoice]
Review records -> Results

[Scenario Cost]
4

[Scenario Set]
recordsReviewed=true

[Scenario Add]
decisionScore+=3

[Scenario MultiChoice]
Interview stakeholders -> Results

[Scenario Cost]
3

[Scenario Set]
interviewed=true
```

`Scenario Cost`, `Scenario Set`, and `Scenario Add` immediately
following a choice belong to that choice.

## 6. Remembering decisions

Use **Scenario Set** to store a value:

``` text
evidenceReviewed=true
```

or:

``` text
approach='consultation'
```

Use **Scenario Add** to change a numeric value:

``` text
time+=4
score+=3
budget-=10
```

Use short, meaningful variable names without spaces.

## 7. Conditional content

Use conditions only when later content genuinely depends on earlier
decisions.

``` text
If :: evidenceReviewed
The learner now sees the evidence summary.
Else
The learner reaches this decision without reviewing the evidence.
EndIf
```

Readable Boolean operators are supported:

``` text
If :: consultation AND riskAssessment
Both safeguards are in place.
EndIf
```

You may use `AND`, `OR`, and `NOT`. Comparisons such as `time <= 60` and
`score >= 10` are also supported.

## 8. Feedback

Apply **Scenario Feedback**:

``` text
good | You checked the evidence before acting.
```

Other useful feedback types include:

``` text
bad | You acted before checking the available evidence.
info | There may be more than one defensible approach.
```

Feedback can be used later in a debrief.

## 9. Gates

A gate prevents normal progression until required state has been
collected.

``` text
Gate :: consentChecked AND riskAssessed
GateFailureTarget :: Missing Information
```

Use gates sparingly. Ordinary branching is often sufficient.

## 10. Time, resources, status and scoring

These features are optional.

``` text
Budget :: time | start=0 | max=72 | unit=hours
Status :: cases | start=23 | label=Cases
Score :: decision | max=70
```

A simple branching scenario does not need any of these.

## 11. Calculations

Advanced simulations can calculate derived values:

``` text
Calculate :: efficiency = 15 - ceil(max(0,time-48)/4)
Clamp :: efficiency between 0 and 15
Calculate :: total = decision + efficiency
```

Common functions include `ceil`, `floor`, `round`, `max`, `min`, and
`abs`.

Calculations are normally best configured by a learning technologist or
technically confident author.

## 12. Outcomes

Use **Scenario Outcome** for conditional endings:

``` text
Outcome :: Strong outcome | when=consulted AND riskAssessed AND score>=10
```

A fallback outcome can be:

``` text
Outcome :: Unresolved outcome | default
```

Put more specific outcomes before broader outcomes, with the default
last.

## 13. Tables and charts

Advanced scenarios can display structured evidence:

``` text
Table :: headers=Area|Population|Cases|Rate
Row :: A|18200|12|66
Row :: B|26700|32|120
```

Example chart:

``` text
Chart :: type=bar | x=Day | y=Cases | data=1,2,4,7,11,14
```

## 14. Author notes

Use the **Author Note** Word style for instructions, rationale,
maintenance notes, or other information intended for authors rather than
learners.

Author notes are not shown in the learner-facing scenario.

## 15. Converting a Word scenario

From the repository root:

``` bash
python word_to_twine.py input/scenario.docx --output-dir output
```

The converter writes:

``` text
output/scenario.twee
output/scenario.html
```

The HTML file is the normal learner-facing deliverable. The Twee file is
useful for portability, version control, inspection, and optional use
with Twine.

To generate only Twee:

``` bash
python word_to_twine.py input/scenario.docx --output-dir output --twee-only
```

## 16. Using Twine

Twine is optional. The converter can generate standalone SugarCube HTML
directly.

Twine is useful as a **visual story-map and QA tool**. Importing or
opening the generated story can help a learning technologist inspect:

-   passage structure;
-   branches and connections;
-   dead ends;
-   unexpectedly dense sections;
-   navigation logic.

### Source-of-truth rule

Use this workflow:

``` text
Word -> word_to_twine.py -> Twee + HTML -> optional Twine inspection
```

If Twine reveals a problem, correct the Word document and regenerate.

Avoid making substantive changes only in Twine. Twine edits are not
automatically written back into the Word authoring document, so
maintaining both independently creates version drift.

## 17. Recommended QA workflow

1.  Create or update the Word scenario.
2.  Run the converter.
3.  Open the generated HTML in a browser.
4.  Test every important decision path.
5.  Optionally inspect the story structure in Twine.
6.  Correct problems in Word.
7.  Regenerate.
8.  Repeat until the scenario is ready to publish.

## 18. Common problems

  -----------------------------------------------------------------------
  Problem                 Likely cause            Fix
  ----------------------- ----------------------- -----------------------
  Broken target           Choice target does not  Correct the target or
                          match a Heading 2       heading

  Choice effect not       Set/Add/Cost is         Keep effects
  applied                 detached from its       immediately after the
                          choice                  choice

  Conditional macro error Incomplete              Ensure every `If` has
                          If/Else/EndIf block     an `EndIf`

  Unexpected ending       Earlier state triggered Review Set/Add actions
                          an outcome              and outcome order

  Input file not found    Wrong filename/path     Check the file exists
                                                  in `input/`

  No HTML                 Validation or converter Review terminal output
                          error                   

  Large visual gaps       Old runtime/output or   Rebuild with the
                          unnecessary blank       current converter
                          content                 
  -----------------------------------------------------------------------

## 19. Suggested responsibilities

  -----------------------------------------------------------------------
  Role                                Typical responsibilities
  ----------------------------------- -----------------------------------
  Academic / subject expert           Content, decisions, feedback,
                                      pedagogical consequences, basic
                                      styles

  Learning technologist               Branching design, state,
                                      conditions, scoring, gates,
                                      calculations, QA

  Developer / platform owner          Converter, SugarCube runtime,
                                      CSS/theme, accessibility defaults,
                                      validation
  -----------------------------------------------------------------------

A learning technologist should be able to build substantial scenarios
with this system. Basic and intermediate authoring is structured
learning design rather than programming. Advanced calculations and
nested conditions require more technical confidence.

## 20. Authoring principles

-   Use normal Word formatting whenever normal Word formatting is
    sufficient.
-   Use scenario styles only when interactive behaviour is required.
-   Keep passage titles short, unique, and meaningful.
-   Prefer meaningful variable names such as `consentChecked` rather
    than `x1`.
-   Use conditions only when content genuinely depends on earlier
    actions.
-   Keep Word as the maintained source of truth.
-   Treat Twee and HTML as generated outputs.
-   Use Twine primarily for visualisation, QA, and debugging.
-   Test every major route before release.

## Quick reference

  Need                        Authoring method
  --------------------------- ------------------------------
  Scenario title              Heading 1
  New screen                  Heading 2
  Subheading                  Heading 3
  Single choice               Scenario Choice
  Multiple selection          Scenario MultiChoice
  Choice cost                 Scenario Cost
  Remember state              Scenario Set
  Change score/resource       Scenario Add
  Conditional content         `If / ElseIf / Else / EndIf`
  Stored feedback             Scenario Feedback
  Required information        `Gate`
  Different ending            Scenario Outcome
  Derived score               `Calculate / Clamp`
  Table                       `Table / Row`
  Chart                       `Chart`
  Non-learner documentation   Author Note

## Minimal complete example

``` text
Heading 1: Data Access Scenario

Heading 2: Request

A researcher asks for access to a dataset containing potentially identifiable information.

[Scenario Choice]
Review the governance documentation -> Governance

[Scenario Choice]
Approve immediately -> Immediate Approval

Heading 2: Governance

The documentation requires a data-access review.

[Scenario Set]
governanceReviewed=true

[Scenario Feedback]
good | You checked the governance requirements before deciding.

[Scenario Choice]
Continue -> Decision

Heading 2: Immediate Approval

[Scenario Feedback]
bad | You approved access without reviewing the governance requirements.

[Scenario Choice]
Continue -> Decision

Heading 2: Decision

If :: governanceReviewed
You can now make an informed decision.
Else
You are making the decision without checking the governance requirements.
EndIf
```
