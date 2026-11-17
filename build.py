#!/usr/bin/env python3
"""
Build index.html for the INSII focal contact map.

The Excel workbook is read from the locally synchronized
SharePoint/OneDrive folder.

Expected folder structure:

    Contacts/
    ├── Web GSP SID _ Contact Network.xlsx
    └── insii-contact-map/
        ├── build.py
        ├── index.html
        ├── data/
        │   └── country_bbox.json
        ├── src/
        │   └── template.html
        └── vendor/
            └── maplibre-gl-4.7.1/
                ├── maplibre-gl.css
                └── maplibre-gl.js

Usage:

    python build.py

Optional:

    python build.py --sheet Website
    python build.py --out index.html

The script reads ONLY:

    ISO (first column; header may be blank)
    Country
    Lat
    Lon
    Region
    Sub Region
    INSII Contact
    INSII (Institution)

No email-address column is read.

Requires:

    Python 3.9+
    openpyxl

Install if necessary:

    python -m pip install openpyxl
"""

import argparse
import json
import re
import sys
from pathlib import Path

try:
    import openpyxl
except ImportError:
    sys.exit(
        "openpyxl is required.\n"
        "Install it with:\n\n"
        "    python -m pip install openpyxl"
    )


# ============================================================
# Paths
# ============================================================

# Folder containing this build.py file:
#
# ...\INSII\Contacts\insii-contact-map

ROOT = Path(__file__).resolve().parent


# The Excel workbook is expected one folder above build.py:
#
# ...\INSII\Contacts\Web GSP SID _ Contact Network.xlsx

WORKBOOK = ROOT.parent / "Web GSP SID _ Contact Network.xlsx"


# Default generated page:
#
# ...\INSII\Contacts\insii-contact-map\index.html

DEFAULT_OUTPUT = ROOT / "index.html"


# ============================================================
# Website-sheet configuration
# ============================================================

WANTED = {
    "country": "country",
    "lat": "lat",
    "lon": "lon",
    "region": "region",
    "sub region": "subregion",
    "insii contact": "contact",
    "insii (institution)": "institution",
}

EMPTY = {
    "",
    "NA",
    "N/A",
    "NONE",
    "NULL",
    "UNSPECIFIED",
}

EMAIL_PATTERN = re.compile(
    r"[\w.+-]+@[\w-]+(?:\.[\w-]+)+",
    flags=re.IGNORECASE,
)


# ============================================================
# Cleaning functions
# ============================================================

def clean(value):
    """Clean text values read from Excel."""

    if value is None:
        return None

    text = str(value).replace("\xa0", " ")
    text = re.sub(r"\s+", " ", text).strip()

    if text.upper() in EMPTY:
        return None

    return text


def normalize_header(value):
    """Normalize an Excel column heading."""

    if value is None:
        return ""

    return re.sub(
        r"\s+",
        " ",
        str(value).replace("\xa0", " "),
    ).strip().lower()


def numeric_coordinate(value):
    """Return a numeric latitude/longitude or None."""

    if value is None or isinstance(value, bool):
        return None

    if isinstance(value, (int, float)):
        return float(value)

    try:
        return float(str(value).strip())
    except (TypeError, ValueError):
        return None


# ============================================================
# Locate workbook
# ============================================================

def find_workbook():
    """
    Find the locally synchronized Excel workbook.

    First checks the expected location immediately above the
    insii-contact-map folder.

    If it is not there, searches the parent folder for a matching
    workbook name.
    """

    if WORKBOOK.exists():
        return WORKBOOK

    parent = ROOT.parent

    matches = list(
        parent.glob("Web GSP SID _ Contact Network*.xlsx")
    )

    # Ignore Excel temporary lock files such as:
    # ~$Web GSP SID _ Contact Network.xlsx

    matches = [
        path
        for path in matches
        if not path.name.startswith("~$")
    ]

    if len(matches) == 1:
        return matches[0]

    if len(matches) > 1:
        sys.exit(
            "More than one matching Excel workbook was found:\n\n"
            + "\n".join(
                f"    {path}"
                for path in matches
            )
            + "\n\nPlease keep only one matching workbook "
              "or update WORKBOOK in build.py."
        )

    sys.exit(
        "The Excel workbook could not be found.\n\n"
        "Expected location:\n\n"
        f"    {WORKBOOK}\n\n"
        "Make sure the SharePoint/OneDrive folder has synchronized "
        "the workbook to this computer."
    )


# ============================================================
# Read Excel data
# ============================================================

def read_rows(xlsx_path, sheet):
    """
    Read approved public fields from the Website worksheet.

    ISO is always taken from the first column.

    All other values are identified using column headings.
    """

    try:
        workbook = openpyxl.load_workbook(
            xlsx_path,
            data_only=True,
            read_only=True,
        )

    except PermissionError:
        sys.exit(
            "The Excel workbook could not be opened.\n\n"
            "It may currently be locked or unavailable through "
            "OneDrive.\n\n"
            "Try closing Excel, wait a few seconds for OneDrive "
            "to finish synchronizing, and run build.py again."
        )

    except Exception as exc:
        sys.exit(
            "The Excel workbook could not be opened:\n\n"
            f"    {xlsx_path}\n\n"
            f"Details: {exc}"
        )

    try:

        if sheet not in workbook.sheetnames:

            sys.exit(
                f'Sheet "{sheet}" was not found.\n\n'
                f"Available sheets:\n"
                + "\n".join(
                    f"    {name}"
                    for name in workbook.sheetnames
                )
            )

        worksheet = workbook[sheet]

        iterator = worksheet.iter_rows(
            values_only=True
        )

        try:
            first_row = next(iterator)

        except StopIteration:
            sys.exit(
                f'Sheet "{sheet}" is empty.'
            )

        header = [
            normalize_header(value)
            for value in first_row
        ]

        missing = [
            column
            for column in WANTED
            if column not in header
        ]

        if missing:

            sys.exit(
                f'Missing required columns in sheet "{sheet}":\n\n'
                + "\n".join(
                    f"    {column}"
                    for column in missing
                )
                + "\n\nFound headers:\n\n"
                + ", ".join(
                    heading if heading else "[blank]"
                    for heading in header
                )
            )

        indexes = {
            output_key: header.index(excel_header)
            for excel_header, output_key
            in WANTED.items()
        }

        output = []
        problems = []

        seen_iso = {}

        for row_number, row in enumerate(
            iterator,
            start=2,
        ):

            # Ignore completely empty rows.

            if not row:
                continue

            if not any(
                value is not None
                for value in row
            ):
                continue

            # ISO is intentionally read from column A,
            # whose heading may be blank.

            iso = clean(
                row[0]
                if len(row) > 0
                else None
            )

            if not iso:
                continue

            iso = iso.upper()

            record = {
                "iso": iso
            }

            for output_key, column_index in indexes.items():

                value = (
                    row[column_index]
                    if column_index < len(row)
                    else None
                )

                if output_key in ("lat", "lon"):

                    record[output_key] = (
                        numeric_coordinate(value)
                    )

                else:

                    record[output_key] = clean(
                        value
                    )

            # --------------------------------------------
            # Validate coordinates
            # --------------------------------------------

            if (
                record["lat"] is None
                or record["lon"] is None
            ):

                problems.append(
                    f"row {row_number} ({iso}): "
                    "lat/lon not numeric - skipped"
                )

                continue

            if not -90 <= record["lat"] <= 90:

                problems.append(
                    f"row {row_number} ({iso}): "
                    f"latitude {record['lat']} "
                    "outside -90 to 90 - skipped"
                )

                continue

            if not -180 <= record["lon"] <= 180:

                problems.append(
                    f"row {row_number} ({iso}): "
                    f"longitude {record['lon']} "
                    "outside -180 to 180 - skipped"
                )

                continue

            # --------------------------------------------
            # Require country name
            # --------------------------------------------

            if not record["country"]:

                problems.append(
                    f"row {row_number} ({iso}): "
                    "no country name - skipped"
                )

                continue

            # --------------------------------------------
            # Prevent duplicate ISO codes
            # --------------------------------------------

            if iso in seen_iso:

                previous_row = seen_iso[iso]

                sys.exit(
                    f"Duplicate ISO code {iso} found.\n\n"
                    f"Rows: {previous_row} and {row_number}\n\n"
                    "Nothing was written."
                )

            seen_iso[iso] = row_number

            output.append(record)

        return output, problems

    finally:

        workbook.close()


# ============================================================
# Privacy check
# ============================================================

def check_for_email_addresses(data):
    """
    Safety check.

    Even though build.py reads only approved columns from Website,
    refuse to generate the public page if an email-like string
    appears anywhere in the extracted data.
    """

    blob = json.dumps(
        data,
        ensure_ascii=False,
        separators=(",", ":"),
    )

    match = EMAIL_PATTERN.search(blob)

    if match:

        sys.exit(
            "SAFETY CHECK FAILED\n\n"
            "An email-like string was found in the data that "
            "would be inserted into the public map:\n\n"
            f"    {match.group(0)}\n\n"
            "Nothing was written.\n\n"
            "Please remove the email address from the Website "
            "sheet and build again."
        )

    return blob


# ============================================================
# Local project files
# ============================================================

def read_required_file(path, description):
    """Read a required local project file."""

    if not path.exists():

        sys.exit(
            f"Required {description} was not found:\n\n"
            f"    {path}"
        )

    try:

        return path.read_text(
            encoding="utf-8"
        )

    except OSError as exc:

        sys.exit(
            f"Could not read {description}:\n\n"
            f"    {path}\n\n"
            f"Details: {exc}"
        )


# ============================================================
# Build HTML
# ============================================================

def build_html(data, output_path):

    bbox_path = (
        ROOT
        / "data"
        / "country_bbox.json"
    )

    template_path = (
        ROOT
        / "src"
        / "template.html"
    )

    vendor_path = (
        ROOT
        / "vendor"
        / "maplibre-gl-4.7.1"
    )

    css_path = (
        vendor_path
        / "maplibre-gl.css"
    )

    js_path = (
        vendor_path
        / "maplibre-gl.js"
    )

    # --------------------------------------------
    # Bounding boxes
    # --------------------------------------------

    bbox_text = read_required_file(
        bbox_path,
        "country bounding-box file",
    )

    try:

        bbox = json.loads(
            bbox_text
        )

    except json.JSONDecodeError as exc:

        sys.exit(
            "country_bbox.json contains invalid JSON:\n\n"
            f"    {exc}"
        )

    # --------------------------------------------
    # Template
    # --------------------------------------------

    template = read_required_file(
        template_path,
        "HTML template",
    )

    # --------------------------------------------
    # MapLibre
    # --------------------------------------------

    maplibre_css = read_required_file(
        css_path,
        "MapLibre CSS file",
    )

    maplibre_js = read_required_file(
        js_path,
        "MapLibre JavaScript file",
    )

    # --------------------------------------------
    # Verify template
    # --------------------------------------------

    required_placeholders = {
        "/*MAPLIBRE_CSS*/",
        "/*MAPLIBRE_JS*/",
        "/*DATA*/",
        "/*BBOX*/",
    }

    missing_placeholders = [
        placeholder
        for placeholder
        in required_placeholders
        if placeholder not in template
    ]

    if missing_placeholders:

        sys.exit(
            "template.html is missing required "
            "placeholder(s):\n\n"
            + "\n".join(
                f"    {placeholder}"
                for placeholder
                in missing_placeholders
            )
        )

    # --------------------------------------------
    # Public-data safety check
    # --------------------------------------------

    blob = check_for_email_addresses(
        data
    )

    # Prevent embedded content from prematurely
    # closing HTML script elements.

    safe_maplibre_js = re.sub(
        r"</script",
        r"<\/script",
        maplibre_js,
        flags=re.IGNORECASE,
    )

    safe_blob = blob.replace(
        "</",
        "<\\/"
    )

    # --------------------------------------------
    # Generate HTML
    # --------------------------------------------

    html = (
        template

        .replace(
            "/*MAPLIBRE_CSS*/",
            maplibre_css,
        )

        .replace(
            "/*MAPLIBRE_JS*/",
            safe_maplibre_js,
        )

        .replace(
            "/*DATA*/",
            safe_blob,
        )

        .replace(
            "/*BBOX*/",
            json.dumps(
                bbox,
                ensure_ascii=False,
                separators=(",", ":"),
            ),
        )
    )

    # --------------------------------------------
    # Write atomically
    # --------------------------------------------

    temporary_output = (
        output_path.with_suffix(
            output_path.suffix + ".tmp"
        )
    )

    try:

        temporary_output.write_text(
            html,
            encoding="utf-8",
        )

        temporary_output.replace(
            output_path
        )

    except OSError as exc:

        try:
            temporary_output.unlink(
                missing_ok=True
            )
        except OSError:
            pass

        sys.exit(
            "Could not write the generated map:\n\n"
            f"    {output_path}\n\n"
            f"Details: {exc}"
        )

    no_bbox = [
        record["iso"]
        for record in data
        if record["iso"] not in bbox
    ]

    return no_bbox


# ============================================================
# Command-line arguments
# ============================================================

def parse_arguments():

    parser = argparse.ArgumentParser(
        description=(
            "Build the INSII contact map from the "
            "locally synchronized SharePoint workbook."
        )
    )

    parser.add_argument(
        "--sheet",
        default="Website",
        help=(
            'Worksheet to read. '
            'Default: "Website".'
        ),
    )

    parser.add_argument(
        "--out",
        default=str(DEFAULT_OUTPUT),
        help=(
            "Output HTML file. "
            f"Default: {DEFAULT_OUTPUT}"
        ),
    )

    return parser.parse_args()


# ============================================================
# Main
# ============================================================

def main():

    arguments = parse_arguments()

    output_path = (
        Path(arguments.out)
        .expanduser()
        .resolve()
    )

    # --------------------------------------------
    # Find synced SharePoint workbook
    # --------------------------------------------

    workbook_path = find_workbook()

    print(
        "Using synchronized SharePoint workbook:"
    )

    print(
        f"  {workbook_path}"
    )

    print()

    # --------------------------------------------
    # Read Website sheet
    # --------------------------------------------

    print(
        f'Reading sheet "{arguments.sheet}"...'
    )

    data, problems = read_rows(
        workbook_path,
        arguments.sheet,
    )

    if not data:

        sys.exit(
            f'No valid country records were found '
            f'in sheet "{arguments.sheet}".\n\n'
            "Nothing was written."
        )

    print(
        f"Found {len(data)} valid country records."
    )

    # --------------------------------------------
    # Build map
    # --------------------------------------------

    print(
        "Building INSII contact map..."
    )

    no_bbox = build_html(
        data,
        output_path,
    )

    print()

    print(
        f"Wrote {output_path}"
    )

    print(
        f"Countries: {len(data)}"
    )

    # --------------------------------------------
    # Quality-control reporting
    # --------------------------------------------

    missing_institution = [
        record["country"]
        for record in data
        if not record["institution"]
    ]

    missing_contact = [
        record["country"]
        for record in data
        if not record["contact"]
    ]

    if missing_institution:

        print()

        print(
            "Countries without institution:"
        )

        print(
            "  "
            + ", ".join(
                missing_institution
            )
        )

    if missing_contact:

        print()

        print(
            "Countries without contact name:"
        )

        print(
            "  "
            + ", ".join(
                missing_contact
            )
        )

    if no_bbox:

        print()

        print(
            "Countries without a zoom extent "
            "(map will use the fixed view):"
        )

        print(
            "  "
            + ", ".join(
                no_bbox
            )
        )

    if problems:

        print()

        print(
            "Warnings:"
        )

        for problem in problems:

            print(
                f"  {problem}"
            )

    print()

    print(
        "Build completed successfully."
    )


if __name__ == "__main__":
    main()