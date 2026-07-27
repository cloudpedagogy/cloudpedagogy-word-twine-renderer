#!/usr/bin/env python3
"""
Word → Twine/Twee converter

Expected Word structure:
    Heading 1 = Story title
    Heading 2 = Passage / node title
    Body text under each Heading 2 = passage content

Supported Twine links:
    [[Choice text->Target Passage]]
    [[Target Passage]]

Supported media markers:
    YouTubeEmbed :: https://www.youtube.com/watch?v=VIDEO_ID
    PanoptoEmbed :: https://lshtm.cloud.panopto.eu/Panopto/Pages/Viewer.aspx?id=...
    IFrameEmbed :: https://example.com/embed/...
    ImageEmbed :: images/example.png

Example:
    python src/course_generator/tools/word_to_twine.py \
      --input imports/twine/literature_review_twine_demo.docx \
      --output output/twine/literature_review_twine_demo.twee
"""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path
from urllib.parse import parse_qs, urlparse

try:
    import mammoth
except ImportError:
    print(
        "Missing dependency: mammoth\n\n"
        "Install it with:\n"
        "  pip install mammoth\n",
        file=sys.stderr,
    )
    raise


HEADING_RE = re.compile(r"<h([1-6])>(.*?)</h\1>", re.IGNORECASE | re.DOTALL)
BLOCK_RE = re.compile(r"(<h[1-6]>.*?</h[1-6]>)", re.IGNORECASE | re.DOTALL)


def clean_text(value: str) -> str:
    value = value.replace("&amp;", "&")
    value = value.replace("&lt;", "<")
    value = value.replace("&gt;", ">")
    value = value.replace("&quot;", '"')
    value = value.replace("&#x27;", "'")
    value = value.replace("\r\n", "\n").replace("\r", "\n")
    value = re.sub(r"\n{3,}", "\n\n", value)
    return value.strip()


def strip_html(value: str) -> str:
    value = re.sub(r"<br\s*/?>", "\n", value, flags=re.IGNORECASE)
    value = re.sub(r"</p\s*>", "\n\n", value, flags=re.IGNORECASE)
    value = re.sub(r"<p\s*>", "", value, flags=re.IGNORECASE)
    value = re.sub(r"</li\s*>", "\n", value, flags=re.IGNORECASE)
    value = re.sub(r"<li\s*>", "- ", value, flags=re.IGNORECASE)
    value = re.sub(r"</?(ul|ol)\s*>", "\n", value, flags=re.IGNORECASE)

    # Convert Word/Mammoth hyperlinks to plain URL text if possible.
    value = re.sub(
        r'<a\s+[^>]*href="([^"]+)"[^>]*>(.*?)</a>',
        lambda m: m.group(1) if m.group(1).strip() else m.group(2),
        value,
        flags=re.IGNORECASE | re.DOTALL,
    )

    value = re.sub(r"<[^>]+>", "", value)
    return clean_text(value)


def youtube_embed_url(url: str) -> str:
    parsed = urlparse(url)

    if "youtube.com" in parsed.netloc:
        query = parse_qs(parsed.query)
        video_id = query.get("v", [""])[0]
        if video_id:
            return f"https://www.youtube.com/embed/{video_id}"

        if parsed.path.startswith("/embed/"):
            return url

    if "youtu.be" in parsed.netloc:
        video_id = parsed.path.strip("/").split("/")[0]
        if video_id:
            return f"https://www.youtube.com/embed/{video_id}"

    return url


def panopto_embed_url(url: str) -> str:
    return url.replace("/Pages/Viewer.aspx", "/Pages/Embed.aspx")


def iframe_html(src: str, width: str = "720", height: str = "405") -> str:
    return (
        f'<iframe width="{width}" height="{height}" '
        f'src="{src}" '
        f'frameborder="0" allowfullscreen></iframe>'
    )


def image_html(src: str) -> str:
    return f'<img src="{src}" alt="" style="max-width: 100%;">'


def replace_media_embeds(text: str) -> str:
    def replace_youtube(match: re.Match) -> str:
        url = match.group(1).strip()
        return iframe_html(youtube_embed_url(url))

    def replace_panopto(match: re.Match) -> str:
        url = match.group(1).strip()
        return iframe_html(panopto_embed_url(url))

    def replace_iframe(match: re.Match) -> str:
        url = match.group(1).strip()
        return iframe_html(url)

    def replace_image(match: re.Match) -> str:
        src = match.group(1).strip()
        return image_html(src)

    text = re.sub(
        r"^YouTubeEmbed\s*::\s*(.+)$",
        replace_youtube,
        text,
        flags=re.MULTILINE,
    )
    text = re.sub(
        r"^PanoptoEmbed\s*::\s*(.+)$",
        replace_panopto,
        text,
        flags=re.MULTILINE,
    )
    text = re.sub(
        r"^IFrameEmbed\s*::\s*(.+)$",
        replace_iframe,
        text,
        flags=re.MULTILINE,
    )
    text = re.sub(
        r"^ImageEmbed\s*::\s*(.+)$",
        replace_image,
        text,
        flags=re.MULTILINE,
    )

    return text


def convert_docx_to_html(docx_path: Path) -> str:
    with docx_path.open("rb") as docx_file:
        result = mammoth.convert_to_html(docx_file)

    if result.messages:
        print(f"\nMammoth messages for {docx_path}:")
        for message in result.messages:
            print(f"  - {message}")

    return result.value


def split_html_by_headings(html: str) -> list[tuple[int, str, str]]:
    parts = BLOCK_RE.split(html)
    sections: list[tuple[int, str, str]] = []

    current_level: int | None = None
    current_title: str | None = None
    current_body: list[str] = []

    for part in parts:
        if not part.strip():
            continue

        heading_match = HEADING_RE.match(part.strip())

        if heading_match:
            if current_level is not None and current_title is not None:
                sections.append((current_level, current_title, "\n".join(current_body)))

            current_level = int(heading_match.group(1))
            current_title = strip_html(heading_match.group(2))
            current_body = []
        else:
            if current_level is not None:
                current_body.append(part)

    if current_level is not None and current_title is not None:
        sections.append((current_level, current_title, "\n".join(current_body)))

    return sections


def validate_twine_links(twee: str) -> list[str]:
    passage_titles = set(
        match.group(1).strip()
        for match in re.finditer(r"^::\s+(.+?)\s*$", twee, flags=re.MULTILINE)
        if match.group(1).strip() != "StoryTitle"
    )

    warnings: list[str] = []

    for link in re.findall(r"\[\[(.+?)\]\]", twee):
        if "->" in link:
            target = link.split("->", 1)[1].strip()
        elif "<-" in link:
            target = link.split("<-", 1)[0].strip()
        else:
            target = link.strip()

        if target and target not in passage_titles:
            warnings.append(f"Link points to missing passage: [[{link}]]")

    return warnings


def docx_to_twee(docx_path: Path) -> str:
    html = convert_docx_to_html(docx_path)
    sections = split_html_by_headings(html)

    if not sections:
        raise ValueError(
            f"No headings found in {docx_path}. "
            "Use Heading 1 for the story title and Heading 2 for passages."
        )

    story_title = docx_path.stem.replace("_", " ").replace("-", " ").title()
    passages: list[tuple[str, str]] = []

    current_passage_title: str | None = None
    current_passage_body: list[str] = []

    for level, title, body_html in sections:
        body_text = replace_media_embeds(strip_html(body_html))

        if level == 1:
            story_title = title or story_title
            continue

        if level == 2:
            if current_passage_title is not None:
                passages.append(
                    (
                        current_passage_title,
                        clean_text("\n\n".join(current_passage_body)),
                    )
                )

            current_passage_title = title
            current_passage_body = []

            if body_text:
                current_passage_body.append(body_text)

            continue

        if current_passage_title is not None:
            heading_line = "#" * min(level, 6) + " " + title
            current_passage_body.append(heading_line)

            if body_text:
                current_passage_body.append(body_text)

    if current_passage_title is not None:
        passages.append(
            (
                current_passage_title,
                clean_text("\n\n".join(current_passage_body)),
            )
        )

    if not passages:
        raise ValueError(
            f"No Heading 2 passages found in {docx_path}. "
            "Use Heading 2 for each Twine node."
        )

    lines: list[str] = []
    lines.append(":: StoryTitle")
    lines.append(story_title)
    lines.append("")

    for title, body in passages:
        lines.append(f":: {title}")
        lines.append(body)
        lines.append("")

    return "\n".join(lines).rstrip() + "\n"


def convert_one(input_path: Path, output_path: Path, overwrite: bool = True) -> None:
    if output_path.exists() and not overwrite:
        raise FileExistsError(f"Output already exists: {output_path}")

    output_path.parent.mkdir(parents=True, exist_ok=True)

    twee = docx_to_twee(input_path)
    warnings = validate_twine_links(twee)

    output_path.write_text(twee, encoding="utf-8")

    print(f"Created: {output_path}")

    if warnings:
        print("Warnings:")
        for warning in warnings:
            print(f"  - {warning}")


def derive_output_path(input_path: Path, output_dir: Path) -> Path:
    return output_dir / f"{input_path.stem}.twee"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Convert Word .docx files into Twine/Twee files."
    )

    mode = parser.add_mutually_exclusive_group(required=True)
    mode.add_argument("--input", type=Path, help="Path to a single .docx file.")
    mode.add_argument(
        "--input-glob",
        help='Glob pattern for multiple .docx files, e.g. "imports/twine/**/*.docx".',
    )

    parser.add_argument(
        "--output",
        type=Path,
        help="Output .twee file path. Use with --input.",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=Path("output/twine"),
        help="Output directory for generated .twee files. Default: output/twine",
    )
    parser.add_argument(
        "--no-overwrite",
        action="store_true",
        help="Do not overwrite existing .twee files.",
    )

    return parser.parse_args()


def main() -> None:
    args = parse_args()
    overwrite = not args.no_overwrite

    if args.input:
        input_path = args.input

        if not input_path.exists():
            raise FileNotFoundError(f"Input file not found: {input_path}")

        if input_path.suffix.lower() != ".docx":
            raise ValueError(f"Input must be a .docx file: {input_path}")

        output_path = args.output or derive_output_path(input_path, args.output_dir)
        convert_one(input_path, output_path, overwrite=overwrite)
        return

    input_paths = sorted(Path().glob(args.input_glob))

    if not input_paths:
        raise FileNotFoundError(f"No files matched: {args.input_glob}")

    for input_path in input_paths:
        if input_path.suffix.lower() != ".docx":
            continue

        output_path = derive_output_path(input_path, args.output_dir)
        convert_one(input_path, output_path, overwrite=overwrite)


if __name__ == "__main__":
    main()