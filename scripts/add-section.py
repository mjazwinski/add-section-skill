"""Add a new section (heading + body text) to a Word (.docx) document,
reusing the same paragraph styles (formatting) as the existing sections.

Usage:
    python add_section.py --doc <path-to-.docx> --name "Section Name" --body "Section body text"

Notes:
 - "Body text" may contain multiple paragraphs separated by newlines ("\n" or "\r\n").
 - Each line becomes its own paragraph using the detected body style.
 - The heading style is detected by looking at the existing section headings in the
   document (paragraphs whose style name starts with "Heading"). The most commonly
   used heading level/style is reused so the new section looks consistent with the
   others. If no heading styles are found, "Heading 1" is used as a fallback.
 - The body style is detected by looking at the paragraph(s) that immediately follow
   each detected heading. The most common style among those is reused for the new
   section's body. If none is found, "Normal" is used as a fallback.
 - The new section is appended at the end of the document.
"""

import argparse
import sys
from collections import Counter

from docx import Document


def detect_heading_style(document):
    """Return the most frequently used heading style name among existing sections."""
    heading_styles = [
        p.style.name
        for p in document.paragraphs
        if p.style and p.style.name and p.style.name.lower().startswith("heading") and p.text.strip()
    ]
    if not heading_styles:
        return "Heading 1"
    return Counter(heading_styles).most_common(1)[0][0]


def detect_body_style(document, heading_style_name):
    """Return the most common style of the first non-empty paragraph after each
    heading that uses heading_style_name."""
    paragraphs = document.paragraphs
    body_styles = []
    for idx, p in enumerate(paragraphs):
        if p.style.name == heading_style_name and p.text.strip():
            for follower in paragraphs[idx + 1:]:
                if follower.style and follower.style.name.lower().startswith("heading"):
                    break
                if follower.text.strip():
                    body_styles.append(follower.style.name)
                    break
    if not body_styles:
        return "Normal"
    return Counter(body_styles).most_common(1)[0][0]


def add_section(doc_path, section_name, section_body, heading_style=None):
    document = Document(doc_path)

    heading_style = heading_style or detect_heading_style(document)
    body_style = detect_body_style(document, heading_style)

    document.add_paragraph(section_name, style=heading_style)

    for line in section_body.splitlines() or [section_body]:
        document.add_paragraph(line, style=body_style)

    document.save(doc_path)
    print(f"Added section '{section_name}' using heading style '{heading_style}' "
          f"and body style '{body_style}'.")


def main():
    parser = argparse.ArgumentParser(description="Add a new section to a Word document.")
    parser.add_argument("--doc", required=True, help="Path to the .docx document to edit")
    parser.add_argument("--name", required=True, help="Name/title of the new section")
    parser.add_argument("--body", required=True, help="Body text of the new section")
    parser.add_argument("--heading-style", default=None,
                        help="Force a specific heading style (e.g. 'Heading 1') instead of "
                             "auto-detecting the most common one used in the document")
    args = parser.parse_args()

    try:
        add_section(args.doc, args.name, args.body, args.heading_style)
    except Exception as exc:  # surface a clear error to the caller
        print(f"Error adding section: {exc}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
