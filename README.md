# CloudPedagogy Word-to-Twine Renderer

Convert structured Microsoft Word documents into interactive branching
scenarios, simulations, and decision-based learning activities.

Authors work in Word. The converter reads Word headings, paragraph
styles, and optional advanced directives, then generates both:

-   Twine-compatible Twee (`.twee`);
-   standalone SugarCube HTML (`.html`) that can be opened directly in a
    browser.

Twine is **optional**: it can be useful for visualising and inspecting
the story map, but it is not required to generate the final HTML.

## Why this approach?

The project is designed around a simple source-of-truth model:

``` text
Word -> word_to_twine.py -> Twee + standalone HTML
                              |
                              +-> optional Twine visualisation / QA
```

Content and branching logic remain maintainable in Word rather than
being locked into one-off HTML applications.

## Features

-   Word-based scenario authoring;
-   Heading 1 for the scenario title;
-   Heading 2 for passages/screens;
-   Heading 3 for subheadings;
-   normal Word paragraphs, bullets, and numbered lists;
-   Word styles for single choices and multiple-selection choices;
-   state variables and remembered decisions;
-   time/resource costs;
-   scoring;
-   conditional content using `If / ElseIf / Else / EndIf`;
-   readable `AND`, `OR`, and `NOT` conditions;
-   feedback and debrief information;
-   progression gates;
-   tables and simple charts;
-   calculations and score clamping;
-   conditional outcomes;
-   responsive standalone SugarCube HTML;
-   Twee output for portability and optional Twine inspection;
-   validation of passage/choice targets;
-   backward compatibility with the earlier explicit directive syntax.

## Authoring manual

See **[AUTHORING.md](docs/AUTHORING.md)** for the complete authoring
guide, including basic branching, Word styles, state, conditions,
scoring, gates, calculations, outcomes, QA, and the recommended Twine
workflow.

## Example scenarios

A repository can use a structure such as:

``` text
.
├── README.md
├── requirements.txt
├── word_to_twine.py
├── docs/
│   └── AUTHORING.md
├── input/
│   └── demo_branching_scenario.docx
├── examples/
│   └── advanced/
│       └── mtaa_saba_scenario.docx
└── output/
```

`demo_branching_scenario.docx` is intended as the beginner example. A
more complex scenario such as Mtaa Saba can demonstrate advanced state,
multi-select decisions, gating, calculations, scoring, and conditional
outcomes.

## Requirements

-   Python 3.10 or later;
-   `python-docx`.

Install dependencies from the repository's `requirements.txt` where
provided.

## Installation

From the repository root:

``` bash
python3 -m venv .venv
source .venv/bin/activate
python3 -m pip install --upgrade pip
python3 -m pip install -r requirements.txt
```

On Windows PowerShell:

``` powershell
.\.venv\Scripts\Activate.ps1
```

## Quick start

Put a Word scenario in the `input/` folder and run:

``` bash
python word_to_twine.py input/demo_branching_scenario.docx --output-dir output
```

The converter should create:

``` text
output/demo_branching_scenario.twee
output/demo_branching_scenario.html
```

Open the generated HTML on macOS:

``` bash
open output/demo_branching_scenario.html
```

On Windows, open the HTML file normally in a browser.

To generate only Twee:

``` bash
python word_to_twine.py input/demo_branching_scenario.docx --output-dir output --twee-only
```

## Basic Word structure

The simplest scenario requires very little syntax.

  Word content/style     Purpose
  ---------------------- ---------------------------
  Heading 1              Scenario title
  Heading 2              New passage/screen
  Heading 3              Subheading
  Normal                 Learner-facing text
  Scenario Choice        Single branching choice
  Scenario MultiChoice   Multiple-selection option
  Scenario Set           Remember a value
  Scenario Add           Change a score/resource
  Scenario Feedback      Store feedback

Example:

``` text
Heading 1: Data Sharing Decision

Heading 2: The Request

A collaborator asks for access to a dataset.

[Scenario Choice]
Review the documentation -> Review the Evidence

[Scenario Choice]
Approve immediately -> Immediate Approval

Heading 2: Review the Evidence

The documentation requires a governance review.

[Scenario Set]
evidenceReviewed=true

[Scenario Choice]
Continue -> Final Decision
```

For the full syntax and advanced features, see **[the authoring
manual](docs/AUTHORING.md)**.

## Advanced scenario logic

The converter can support richer simulations where required.

Example conditional content:

``` text
If :: evidenceReviewed
You can make an informed decision.
Else
You have not yet reviewed the evidence.
EndIf
```

Readable conditions are supported:

``` text
If :: consentChecked AND riskAssessed
...
EndIf
```

Example calculation:

``` text
Calculate :: efficiency = 15 - ceil(max(0,time-48)/4)
```

Example conditional outcome:

``` text
Outcome :: Strong outcome | when=consulted AND riskAssessed
Outcome :: Unresolved outcome | default
```

These features are optional. A basic branching scenario does not require
variables, scoring, calculations, or outcomes.

## Using Twine

Twine is not required for the normal build process because
`word_to_twine.py` generates standalone HTML directly.

Twine can still be useful for:

-   visualising the passage/story map;
-   inspecting branches;
-   identifying dead ends;
-   reviewing complex navigation;
-   debugging a scenario.

The recommended rule is:

> **Word is the source of truth.**

If Twine reveals a problem, make the correction in Word and regenerate
the outputs. Editing only in Twine creates a separate version that is
not automatically written back to the Word source.

## Output files

For:

``` text
input/my_scenario.docx
```

running:

``` bash
python word_to_twine.py input/my_scenario.docx --output-dir output
```

creates:

``` text
output/my_scenario.twee
output/my_scenario.html
```

The `.html` file is the normal learner-facing deliverable.

The `.twee` file is useful for portability, version control, inspection,
and Twine/Twee workflows.

## Validation

Before generating outputs, the converter checks scenario structure,
including:

-   duplicate Heading 2 passage names;
-   whether the start passage exists;
-   broken Twine-style links;
-   broken `Choice` / `MultiChoice` targets.

A validation failure is reported in the terminal so the Word source can
be corrected and regenerated.

## Troubleshooting

### `FileNotFoundError`

The input filename or path is wrong. Check the contents of the input
folder:

``` bash
ls input
```

Then use the exact filename:

``` bash
python word_to_twine.py input/my_scenario.docx --output-dir output
```

### Broken choice target

The text after `->` must exactly match a Heading 2 passage title.

### Conditional error

Check that every:

``` text
If ::
```

has a corresponding:

``` text
EndIf
```

### Unexpected outcome

The scenario is stateful. An earlier `Scenario Set`, `Scenario Add`, or
choice may have triggered the outcome. Review the state changes and
outcome order.

## Accessibility and QA

The generated scenario should still be tested before publication.

Recommended checks include:

-   keyboard navigation;
-   screen-reader behaviour;
-   heading structure;
-   descriptive choice text;
-   colour contrast;
-   responsive/mobile layout;
-   meaningful alternative text for informative images;
-   captions/transcripts for media;
-   every important branching route;
-   restart/save behaviour where used.

Accessibility and responsive behaviour should be provided by the shared
converter/runtime wherever possible rather than being reimplemented by
each author.

## Recommended roles

**Academic / subject expert:** content, decisions, feedback, pedagogical
consequences.

**Learning technologist:** branching design, state, conditions, scoring,
gates, calculations, QA.

**Developer / platform owner:** converter, SugarCube runtime, styling,
accessibility defaults, validation.

## Source of truth and generated files

Treat Word documents as maintained source files.

Treat `.twee` and `.html` as generated outputs that can be recreated
from Word.

This avoids divergence between a Word version, a Twine version, and
manually edited HTML.

## Licence

Add the repository licence in `LICENSE`. If the project is released
under the MIT Licence, include the standard MIT licence text there.
