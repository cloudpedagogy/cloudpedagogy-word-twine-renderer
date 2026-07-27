# CloudPedagogy Word-to-Twine Renderer

Convert structured Microsoft Word documents into Twine-compatible Twee files for branching scenarios, interactive narratives and decision-based learning activities.

The converter allows authors to design a scenario in a familiar Word document. Word headings define the story and its passages, while simple Twine link notation defines the choices between them.

## Live demo

[Open the live demonstration](http://cloudpedagogy-word-twine-renderer.s3-website.eu-west-2.amazonaws.com/Vaccine_Effectiveness_Slide_Test_slides.html#/title-slide)

## Features

- converts `.docx` files to Twee (`.twee`);
- uses Word Heading styles to define the story structure;
- supports branching choices and links between passages;
- supports YouTube, Panopto, iframe and image embed markers;
- processes a single document or multiple documents with a glob pattern;
- checks links and reports references to missing passages;
- creates output directories automatically;
- supports optional protection against overwriting existing files.

## Project files

```text
.
├── README.md
├── requirements.txt
├── word_to_twine.py
├── literature_review_twine_demo.docx
└── output/
    └── twine/
```

## Requirements

- Python 3.10 or later
- [Mammoth](https://github.com/mwilliamson/python-mammoth)

The converter itself does not require Twine to generate a `.twee` file. To edit and publish the resulting story visually, use the free [Twine application](https://twinery.org/). Because Twine imports compiled story HTML rather than raw `.twee` source, use the free [Tweego compiler](https://www.motoslave.net/tweego/) to compile the generated file first.

## Installation

Open Terminal, move into the repository root, and create a virtual environment:

```bash
cd cloudpedagogy-word-twine-renderer
python3 -m venv .venv
source .venv/bin/activate
python3 -m pip install --upgrade pip
python3 -m pip install -r requirements.txt
```

If you downloaded the repository as a ZIP file, extract it first, open Terminal,
and use `cd` to enter the extracted folder. For example:

```bash
cd ~/Downloads/cloudpedagogy-word-twine-renderer-main
```

Then run the virtual-environment and installation commands shown above.

On Windows PowerShell, activate the virtual environment with:

```powershell
.\.venv\Scripts\Activate.ps1
```

All remaining commands can be run from the repository root.

## Word document structure

The converter interprets the document as follows:

| Word content | Purpose |
|---|---|
| Heading 1 | Story title |
| Heading 2 | Twine passage or node |
| Normal body text | Passage content |
| Heading 3–6 | Subheadings inside the current passage |

Use one Heading 1 at the start of the document and a Heading 2 for every passage.

### Example

```text
Literature Review Scenario                 [Heading 1]

Start                                      [Heading 2]
You are beginning a literature review.

[[Define the question->Focused Question]]
[[Start searching->Search Too Soon]]

Focused Question                           [Heading 2]
You clarify the question before searching.

[[Develop a search strategy->Search Strategy]]
```

The target name in each link must exactly match a Heading 2 passage title.

## Twine links

The following link formats are supported:

```text
[[Choice text->Target Passage]]
[[Target Passage]]
```

For a labelled choice, the text before `->` is shown to the learner and the text after it identifies the destination passage.

## Media embeds

Place a media marker on its own line in the body of a passage.

### YouTube

```text
YouTubeEmbed :: https://www.youtube.com/watch?v=VIDEO_ID
```

YouTube watch and `youtu.be` links are converted to embeddable URLs.

### Panopto

```text
PanoptoEmbed :: https://example.cloud.panopto.eu/Panopto/Pages/Viewer.aspx?id=VIDEO_ID
```

The converter changes the Panopto viewer URL to its corresponding embed URL.

### Iframe

```text
IFrameEmbed :: https://example.com/embed/
```

Only embed content from sources you trust and verify that the site permits iframe embedding.

### Image

```text
ImageEmbed :: images/example.png
```

The image path is retained in the generated Twee source. Ensure the referenced image is available at the same relative location when the story is compiled and published. Generated image tags currently use an empty `alt` attribute, so add appropriate alternative text during the final accessibility review when the image conveys meaning.

## Run the demonstration

From the repository root:

```bash
python3 word_to_twine.py \
  --input literature_review_twine_demo.docx \
  --output output/twine/literature_review_twine_demo.twee
```

Expected confirmation:

```text
Created: output/twine/literature_review_twine_demo.twee
```

If `--output` is omitted, the converter uses `output/twine/` by default:

```bash
python3 word_to_twine.py \
  --input literature_review_twine_demo.docx
```

## Open the generated story in Twine

Twine is a free, open-source application available as a desktop download and as a browser-based application:

- [Download or open Twine](https://twinery.org/)
- [Download Tweego](https://www.motoslave.net/tweego/)

The converter creates a `.twee` source file. Twine does not directly import raw `.twee` files, so first compile the file into a Twine story HTML file with Tweego.

### 1. Compile the `.twee` file

After installing Tweego, run the following command from the repository root:

```bash
tweego \
  -o output/twine/literature_review_twine_demo.html \
  output/twine/literature_review_twine_demo.twee
```

This creates:

```text
output/twine/literature_review_twine_demo.html
```

You can open this HTML file directly in a web browser to test the story.

### 2. Import the compiled story into Twine

1. Open the Twine desktop application or the browser-based version of Twine.
2. Go to the story library.
3. Choose **Import From File**.
4. Select `output/twine/literature_review_twine_demo.html`.
5. Open the imported story to inspect or edit its passages.
6. Use **Build → Publish to File** when you are ready to export the finished story.

The exact menu wording can vary slightly between Twine versions, but the workflow is the same: compile `.twee` to `.html` with Tweego, then import the compiled HTML into Twine.

## Convert another Word document

```bash
python3 word_to_twine.py \
  --input path/to/scenario.docx \
  --output output/twine/scenario.twee
```

## Convert multiple documents

Use `--input-glob` to process multiple `.docx` files:

```bash
python3 word_to_twine.py \
  --input-glob "examples/**/*.docx" \
  --output-dir output/twine
```

Keep the glob pattern in quotation marks so the converter receives it unchanged.

## Prevent overwriting

By default, an existing output file is replaced. To stop instead of overwriting it:

```bash
python3 word_to_twine.py \
  --input literature_review_twine_demo.docx \
  --output output/twine/literature_review_twine_demo.twee \
  --no-overwrite
```

## Command-line options

| Option | Description |
|---|---|
| `--input PATH` | Convert one `.docx` file |
| `--input-glob PATTERN` | Convert multiple matching `.docx` files |
| `--output PATH` | Set the `.twee` output path for one input file |
| `--output-dir PATH` | Set the output directory; default: `output/twine` |
| `--no-overwrite` | Refuse to overwrite an existing output file |
| `-h`, `--help` | Display command help |

`--input` and `--input-glob` are mutually exclusive, and one of them is required.

## Link validation

After conversion, the script compares every Twine link target with the available passage titles. A missing target produces a warning such as:

```text
Warnings:
  - Link points to missing passage: [[Continue->Missing Passage]]
```

The output file is still created, allowing the source document to be corrected and converted again.

## Troubleshooting

### `python3: command not found`

Install Python 3 and reopen Terminal.

### `Missing dependency: mammoth`

Activate the virtual environment and install Mammoth:

```bash
source .venv/bin/activate
python3 -m pip install -r requirements.txt
```

### `No headings found`

Apply Word's built-in **Heading 1** style to the story title and **Heading 2** to every passage title. Making text bold or increasing its font size is not sufficient.

### `No Heading 2 passages found`

The document contains a title but no passages. Apply the **Heading 2** style to each passage or node title.

### Missing-passage warnings

Check spelling, spacing and capitalisation. Each link target must match its destination Heading 2 title exactly.

### Twine will not import the `.twee` file

Compile the `.twee` file to `.html` with Tweego, then use Twine's **Import From File** option to import the resulting HTML file.

### `tweego: command not found`

Confirm that Tweego has been downloaded and that its executable is available on your system's `PATH`. Alternatively, run Tweego using the full path to the downloaded executable.

## Accessibility and quality assurance

Review the compiled story before publication. In particular:

- test every choice and return path;
- check keyboard navigation in the selected Twine story format;
- add meaningful alternative text for informative images;
- confirm that videos include captions or transcripts;
- provide descriptive link and choice text;
- verify colour contrast and responsive behaviour;
- check that embedded services are permitted by local privacy, security and content policies.

## Limitations

- the output is Twee source rather than a standalone HTML file;
- Word formatting is simplified during conversion;
- only Heading 1 and Heading 2 have structural story roles;
- embedded images are referenced rather than copied;
- custom Twine story-format features are not added automatically;
- broken links are reported as warnings rather than stopping conversion.

## Licence

Add the repository's licence here. If the project is released under the MIT Licence, include a `LICENSE` file in the repository.
