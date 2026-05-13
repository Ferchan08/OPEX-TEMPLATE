"""
fill_opex_template.py
=====================
Fills the OpEx Charter PPTX template with data from a JSON file.

Usage:
    python fill_opex_template.py --json data.json --template template.pptx --output output.pptx

The JSON must contain the keys defined in FIELD_MAP below.
"""

import argparse
import json
import sys
import zipfile
import shutil
import os
import re
from pathlib import Path
from lxml import etree

# ---------------------------------------------------------------------------
# Namespaces
# ---------------------------------------------------------------------------
NSMAP = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}


def ns(prefix, tag):
    return f"{{{NSMAP[prefix]}}}{tag}"


# ---------------------------------------------------------------------------
# Shape name → JSON key mapping
# ---------------------------------------------------------------------------
FIELD_MAP = {
    "ProjectNameBox":      "project_name",
    "ProblemStatementBox": "problem_statement",
    "ProjectLeaderBox":    "project_leader",
    "TeamMembersBox":      "team_members",
    "ProjectObjectivesBox":"project_objectives",
    "ProjectTimingBox":    "project_timing",
    "ProjectSponsorsBox":  "project_sponsors",
    "OpExMasterBox":       "opex_master",
    "InScopeBox":          "in_scope",
    "OutOfScopeBox":       "out_of_scope",
    "ProjectImpactBox":    "project_impact_3y",
    "FinanceBox":          "finance",
    # Timeline date boxes
    "Define":              "define",
    "Measure":             "measure",
    "Analyze":             "analyze",
    "Improve":             "improve",
    "Control":             "control",
}

# For date fields, show only the date portion (YYYY-MM-DD → DD/MM/YYYY)
DATE_FIELDS = {"define", "measure", "analyze", "improve", "control"}


def format_value(key, value):
    """Format a JSON value for display in the slide."""
    if key in DATE_FIELDS and value:
        # Convert ISO date to DD/MM/YYYY
        try:
            from datetime import datetime
            dt = datetime.strptime(str(value), "%Y-%m-%d")
            return dt.strftime("%d/%m/%Y")
        except ValueError:
            return str(value)
    return str(value) if value is not None else ""


# ---------------------------------------------------------------------------
# XML helpers
# ---------------------------------------------------------------------------

def get_first_run_rpr(txBody):
    """Return the rPr element of the first run in the text body (to copy formatting)."""
    for para in txBody.findall(f".//{ns('a','r')}"):
        rpr = para.find(ns("a", "rPr"))
        if rpr is not None:
            return rpr
    return None


def set_text_in_shape(shape_elem, new_text, label_text=None):
    """
    Replace the body text of a shape with new_text.
    If label_text is given, the first paragraph keeps the label (bold/italic)
    and a new paragraph with new_text is appended.
    """
    txBody = shape_elem.find(f".//{ns('p','txBody')}")
    if txBody is None:
        return

    # Collect existing paragraphs to grab formatting reference
    existing_paras = txBody.findall(ns("a", "p"))
    ref_rpr = get_first_run_rpr(txBody)

    if label_text is None:
        # ----------------------------------------------------------------
        # No label: replace ALL text content with new_text in first para
        # ----------------------------------------------------------------
        # Remove all existing paragraphs
        for p in existing_paras:
            txBody.remove(p)

        # Build a single paragraph with the new text
        para = etree.SubElement(txBody, ns("a", "p"))
        pPr = etree.SubElement(para, ns("a", "pPr"))
        pPr.set("defTabSz", "685772")
        etree.SubElement(pPr, ns("a", "defRPr"))

        run = etree.SubElement(para, ns("a", "r"))
        rPr = etree.SubElement(run, ns("a", "rPr"))
        rPr.set("lang", "es-MX")
        if ref_rpr is not None:
            # Copy size/font settings from original
            if ref_rpr.get("sz"):
                rPr.set("sz", ref_rpr.get("sz"))
            latin = ref_rpr.find(ns("a", "latin"))
            if latin is not None:
                import copy
                run.append(copy.deepcopy(latin))
        rPr.set("dirty", "0")

        t = etree.SubElement(run, ns("a", "t"))
        t.text = new_text
        return

    # ----------------------------------------------------------------
    # Has label: keep first paragraph intact, add value paragraph below
    # ----------------------------------------------------------------
    # Remove all paragraphs AFTER the first (the label)
    for p in existing_paras[1:]:
        txBody.remove(p)

    # Add a value paragraph
    para = etree.SubElement(txBody, ns("a", "p"))
    pPr = etree.SubElement(para, ns("a", "pPr"))
    pPr.set("defTabSz", "685772")
    etree.SubElement(pPr, ns("a", "defRPr"))

    run = etree.SubElement(para, ns("a", "r"))
    rPr = etree.SubElement(run, ns("a", "rPr"))
    rPr.set("lang", "es-MX")
    if ref_rpr is not None:
        if ref_rpr.get("sz"):
            rPr.set("sz", ref_rpr.get("sz"))
        import copy
        latin = ref_rpr.find(ns("a", "latin"))
        if latin is not None:
            run.append(copy.deepcopy(latin))
    rPr.set("dirty", "0")

    t = etree.SubElement(run, ns("a", "t"))
    t.text = new_text


# ---------------------------------------------------------------------------
# Shapes that have a bold/italic label on the first paragraph
# (keep the label paragraph, add value paragraph below)
# ---------------------------------------------------------------------------
LABEL_SHAPES = {
    "ProblemStatementBox": "Problem Statement:",
    "ProjectLeaderBox":    "Project Leader:",
    "TeamMembersBox":      "Team Members:",
    "ProjectObjectivesBox":"Project Objective(s):",
    "ProjectTimingBox":    "Project Timing",
    "ProjectSponsorsBox":  "Project Sponsor(s):",
    "OpExMasterBox":       "OpEx Master:",
    "InScopeBox":          "In Scope:",
    "OutOfScopeBox":       "Out of Scope:",
    "ProjectImpactBox":    "Project Impact (first 3 years):",
    "FinanceBox":          "Finance:",
}

# Date shapes contain only the date text (no label on slide)
DATE_SHAPE_NAMES = {"Define", "Measure", "Analyze", "Improve", "Control"}


# ---------------------------------------------------------------------------
# Core fill function
# ---------------------------------------------------------------------------

def fill_template(template_path: str, data: dict, output_path: str):
    """
    Open the template PPTX, fill shapes from data, write to output_path.
    """
    # Work on a copy
    tmp_path = output_path + ".tmp.pptx"
    shutil.copy2(template_path, tmp_path)

    with zipfile.ZipFile(tmp_path, "r") as zin:
        names = zin.namelist()
        slide_names = [n for n in names if re.match(r"ppt/slides/slide\d+\.xml$", n)]

    if not slide_names:
        raise ValueError("No slides found in the PPTX file.")

    # We expect a single slide (the charter)
    slide_xml_path = slide_names[0]

    with zipfile.ZipFile(tmp_path, "r") as zin:
        slide_bytes = zin.read(slide_xml_path)

    tree = etree.fromstring(slide_bytes)

    # Iterate all shapes
    for sp in tree.findall(f".//{ns('p','sp')}"):
        cNvPr = sp.find(f".//{ns('p','cNvPr')}")
        if cNvPr is None:
            continue
        shape_name = cNvPr.get("name", "")

        if shape_name not in FIELD_MAP:
            continue

        json_key = FIELD_MAP[shape_name]
        raw_value = data.get(json_key, "")
        value = format_value(json_key, raw_value)

        if shape_name in DATE_SHAPE_NAMES:
            # Date shapes: just replace all text with formatted date
            set_text_in_shape(sp, value, label_text=None)
        elif shape_name == "ProjectNameBox":
            # Project name: no label in the shape, just the project name
            set_text_in_shape(sp, value, label_text=None)
        else:
            label = LABEL_SHAPES.get(shape_name)
            set_text_in_shape(sp, value, label_text=label)

    modified_xml = etree.tostring(tree, xml_declaration=True, encoding="UTF-8", standalone=True)

    # Rewrite the ZIP with the modified slide
    final_path = output_path
    with zipfile.ZipFile(tmp_path, "r") as zin:
        with zipfile.ZipFile(final_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == slide_xml_path:
                    zout.writestr(item, modified_xml)
                else:
                    zout.writestr(item, zin.read(item.filename))

    os.remove(tmp_path)
    print(f"✅ Filled presentation saved to: {final_path}")


# ---------------------------------------------------------------------------
# CLI entry point
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="Fill OpEx Charter PPTX template from JSON data.")
    parser.add_argument("--json",     required=True, help="Path to the input JSON file")
    parser.add_argument("--template", required=True, help="Path to the PPTX template file")
    parser.add_argument("--output",   required=True, help="Path for the filled output PPTX")
    args = parser.parse_args()

    # Load JSON
    json_path = Path(args.json)
    if not json_path.exists():
        print(f"❌ JSON file not found: {json_path}", file=sys.stderr)
        sys.exit(1)

    with open(json_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    # Fill template
    fill_template(
        template_path=args.template,
        data=data,
        output_path=args.output,
    )


if __name__ == "__main__":
    main()
