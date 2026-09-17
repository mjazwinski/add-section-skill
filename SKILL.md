name: add-ms-word-docx-document-section
description: Use this skill to add a new section (heading + body text) to a Word document, matching the formatting of the document's existing sections
user-invocable: true
disable-model-invocation: false
argument-hint: section name, section body
---

This skill is invoked once the section name and section body text have already been specified.

Prerequisites:
1. You must know the path to the .docx document being edited. If you don't know it, ask the user.
2. You must have the new section's name and body text. Those must already be present in the context
   (this skill is invoked after they are specified). If either is missing, ask the user.

The skill script (`scripts/add-section.py`, next to this SKILL.md file) detects the heading and body
paragraph styles used by the existing sections in the document and reuses them for the new section, so
the new section's formatting (font, style, etc.) matches the rest of the document. The new section is
appended at the end of the document.

Steps to run it:

1. Ensure a Python virtual environment exists. If not, create one, e.g.:

   python -m venv .venv

2. Activate the venv:
   - Windows (PowerShell): `.venv\Scripts\Activate.ps1`
   - Linux/macOS: `source .venv/bin/activate`
3. Ensure `python-docx` is installed in the venv; install it if missing:

   python -m pip show python-docx || python -m pip install python-docx

4. Run the script, passing the document path, section name, and section body as arguments:

   python <path to this skill's scripts/add-section.py> --doc "<document path>" --name "<section name>" --body "<section body text>"

5. Confirm the script printed a success message naming the heading/body styles it reused, and let the
   user know the section was added. If it fails, report the error output to the user.
