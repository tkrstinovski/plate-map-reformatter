#!/usr/bin/env python3
"""
96-well plate reformatter with plate-name prompt.

Run:
    python plate_reformatter.py

Dependencies:
    python -m pip install openpyxl xlrd

The program accepts .xlsx and legacy Excel 97-2003 .xls input files, then saves
a reformatted .xlsx workbook beside the source workbook.
"""

import re
import math
import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox
from pathlib import Path
from collections import OrderedDict
import shutil
import subprocess
import tempfile

import openpyxl
try:
    import xlrd  # Optional fallback for reading legacy .xls files directly
except ImportError:
    xlrd = None
from datetime import datetime, timedelta
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter

NROWS, NCOLS = 8, 12
SOURCE_PREFIX_FIELDS = (
    "Study Number", "Custom ID", "Sample Number", "Group Number", "Matrix/Split"
)
NUMERIC_TEXT_PATTERN = r"(?:(?:\d{1,3}(?:,\d{3})+)|\d+)(?:\.\d+)?"
DILUTION_CANDIDATE_PATTERN = r"[+-]?(?:(?:\d[\d,]*)(?:\.\d+)?|\.\d+)(?:[eE][+-]?\d+)?"
FILL = {
    "standard": "D9EAF7",
    "qc": "E2F0D9",
    "control": "E4DFEC",
    "sample": "FCE4D6",
    "empty": "F2F2F2",
    "header": "E7E6E6",
}
thin = Side(style="thin", color="000000")
medium = Side(style="medium", color="000000")
grid = Border(left=thin, right=thin, top=thin, bottom=thin)
outer = Border(left=medium, right=medium, top=medium, bottom=medium)


def txt(value):
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()


def get_study_number(ws, data):
    title = " ".join(txt(ws.cell(r, c).value) for r in range(1, min(ws.max_row, 3) + 1)
                     for c in range(1, min(ws.max_column, 3) + 1))
    match = re.search(
        r"\bStudy\s+(?:Number|No\.?)\s*[:#]?\s*([A-Za-z0-9][A-Za-z0-9_-]*)\b",
        title,
        re.I,
    )
    if not match:
        match = re.search(
            r"\bStudy\s+(?!Number\b|No\.?\b)([A-Za-z0-9][A-Za-z0-9_-]*)\b",
            title,
            re.I,
        )
    if match:
        return clean_study_number(match.group(1))
    for row in data:
        for value in row:
            match = re.match(
                r"^((?:\d{6,}|[A-Za-z]{1,4}-\d{4,})(?:_[A-Za-z0-9-]+)?)\b",
                value,
            )
            if match:
                return clean_study_number(match.group(1))
    return ""


def clean_study_number(value):
    """Remove a recognized assay identifier suffix from a study identifier."""
    return re.sub(r"_(?:TSH|ELISA|CYTO)$", "", value, flags=re.I)


class MissingInformationError(ValueError):
    """Raised before output creation when required source data is missing."""

    def __init__(self, missing):
        self.missing = missing
        super().__init__("Required information is missing; no workbook was created.")


class InvalidInformationError(ValueError):
    """Raised when a user-supplied value fails a required format check."""

    def __init__(self, invalid):
        self.invalid = invalid
        super().__init__("One or more values are invalid; no workbook was created.")


def is_numeric_value(value):
    value = user_text(value)
    # Accept ordinary positive numbers and optional correctly grouped thousands
    # separators; intentionally reject scientific notation and other symbols.
    if not re.fullmatch(NUMERIC_TEXT_PATTERN, value):
        return False
    try:
        number = float(value.replace(",", ""))
        return math.isfinite(number) and number > 0
    except (TypeError, ValueError, OverflowError):
        return False


def normalize_dilution_value(value):
    """Remove thousands commas and redundant leading zeros for display."""
    value = user_text(value).replace(",", "")
    integer_part, decimal_point, fractional_part = value.partition(".")
    integer_part = integer_part.lstrip("0") or "0"
    return integer_part + (decimal_point + fractional_part if decimal_point else "")


def user_text(value):
    """Keep user-entered text intact apart from surrounding whitespace."""
    return "" if value is None else str(value).strip()


def inspect_workbook(input_file):
    """Find missing study, source-entry, or per-well lookup fields without output."""
    wb_in, src = load_source_workbook(input_file)
    try:
        data_row, data_col, warnings = find_source_plate(src)
        raw = [
            [txt(src.cell(data_row + r, data_col + c).value) for c in range(NCOLS)]
            for r in range(NROWS)
        ]
        warnings.extend(sample_identity_conflict_warnings(raw))
        study = get_study_number(src, raw)

        samples, _ = group_horizontal_samples(raw)
        labels = {key: label_for_index(i) for i, key in enumerate(samples)}

        missing = []
        if not study:
            missing.append({
                "kind": "study_number", "well": "", "label": "",
                "source_entry": "", "field": "Study number", "value": "",
            })
        for key, coords in samples.items():
            for r, c in coords:
                well = f"{chr(65 + r)}{c + 1}"
                metadata = parse_source_metadata(raw[r][c])
                if not metadata["prefix_complete"]:
                    missing.append({
                        "kind": "source_entry", "well": well, "label": labels[key],
                        "source_entry": raw[r][c],
                        "field": "Incomplete source metadata prefix", "value": "",
                        "reason": (
                            "Expected leading fields are Study Number, Custom ID, Sample Number, "
                            "Group Number, and Matrix/Split. The exact missing component cannot "
                            "be identified safely, so enter the complete corrected source entry."
                        ),
                    })
                    continue
                for field in ("Sample Number", "Test Period", "Dilution Fold"):
                    value = metadata[field]
                    invalid_dilution = (
                        field == "Dilution Fold" and value and not is_numeric_value(value)
                    )
                    if not value or invalid_dilution:
                        missing.append({
                            "kind": "well", "well": well, "label": labels[key],
                            "source_entry": raw[r][c], "field": field, "value": value,
                            "reason": (
                                f"Existing dilution value {value!r} is invalid."
                                if invalid_dilution else "Required value is missing."
                            ),
                        })
        return {"study": study, "warnings": warnings, "missing": missing}
    finally:
        wb_in.close()


def find_duplicate_lookup_groups(input_file, well_values=None):
    """Find separate lookup rows sharing sample number, period, and dilution.

    Corrections are applied before horizontal grouping, matching workbook
    creation behavior. Only distinct resulting lookup rows are compared.
    """
    well_values = well_values or {}
    wb_in, src = load_source_workbook(input_file)
    try:
        data_row, data_col, _ = find_source_plate(src)
        raw = [
            [txt(src.cell(data_row + r, data_col + c).value) for c in range(NCOLS)]
            for r in range(NROWS)
        ]
    finally:
        wb_in.close()

    completed = [row.copy() for row in raw]
    fields_by_coord = {}
    for c in range(NCOLS):
        for r in range(NROWS):
            if classify(raw[r][c])[0] != "sample":
                continue
            well = f"{chr(65 + r)}{c + 1}"
            corrections = well_values.get(well, {})
            source_override = user_text(corrections.get("Source metadata entry", ""))
            if source_override:
                entry = " ".join(source_override.split())
            else:
                extracted = dict(zip(
                    ("Sample Number", "Test Period", "Dilution Fold"),
                    parse_fields(raw[r][c]),
                ))
                required = {}
                for field, old_value in extracted.items():
                    if not old_value or (field == "Dilution Fold" and not is_numeric_value(old_value)):
                        value = user_text(corrections.get(field, ""))
                        if value:
                            required[field] = value
                entry = complete_source_entry(raw[r][c], required)
            completed[r][c] = entry
            metadata = parse_source_metadata(entry)
            dilution = metadata["Dilution Fold"]
            if is_numeric_value(dilution):
                dilution = normalize_dilution_value(dilution)
            fields_by_coord[(r, c)] = (
                metadata["Sample Number"], metadata["Test Period"], dilution
            )

    samples, _ = group_horizontal_samples(completed)
    rows_by_key = {}
    for sample_index, (key, coords) in enumerate(samples.items()):
        display_key = fields_by_coord[coords[0]]
        if not all(display_key):
            continue
        comparison_key = tuple(value.casefold() for value in display_key)
        rows_by_key.setdefault(comparison_key, {"metadata": display_key, "rows": []})["rows"].append({
            "label": label_for_index(sample_index),
            "wells": [f"{chr(65 + r)}{c + 1}" for r, c in coords],
            "source_entry": completed[coords[0][0]][coords[0][1]],
        })

    return [
        {
            "sample_number": group["metadata"][0],
            "test_period": group["metadata"][1],
            "dilution": group["metadata"][2],
            "rows": group["rows"],
        }
        for group in rows_by_key.values() if len(group["rows"]) > 1
    ]


def classify(value):
    if not value:
        return "empty", ""
    if value.lower() == "reagent" or "control blank" in value.lower():
        return "control", "Control Blank"
    # Per the requested cleanup, retain the STD/QC designation and discard
    # everything after its designation, including replicate/trailing numbers.
    m = re.search(
        r"(?i)\b(t?QC|HQC|MQC|LQC|EQC|STD)\s*-?\s*(\d+)(?:\s*-\s*(\d+))?",
        value,
    )
    if m:
        label = f"{m.group(1)}{m.group(2)}"
        # A dashed second number is part of a QC designation (for example,
        # QC1-1), while later trailing/replicate text is still discarded.
        if m.group(3) and "qc" in m.group(1).lower():
            label += f"-{m.group(3)}"
        return ("qc" if "qc" in m.group(1).lower() else "standard"), label
    return "sample", value


def group_horizontal_samples(values):
    """Group identical sample entries only across contiguous horizontal wells."""
    samples = OrderedDict()
    group_for_coord = {}
    for c in range(NCOLS):
        for r in range(NROWS):
            if classify(values[r][c])[0] != "sample":
                continue
            if (
                c > 0
                and classify(values[r][c - 1])[0] == "sample"
                and values[r][c] == values[r][c - 1]
            ):
                group_key = group_for_coord[(r, c - 1)]
                samples[group_key].append((r, c))
            else:
                group_key = (r, c)
                samples[group_key] = [(r, c)]
            group_for_coord[(r, c)] = group_key
    return samples, group_for_coord


def sample_identity_conflict_warnings(values):
    """Flag different source entries sharing a sample number and test period."""
    entries_by_metadata = {}
    for c in range(NCOLS):
        for r in range(NROWS):
            entry = values[r][c]
            if classify(entry)[0] != "sample":
                continue
            metadata = parse_source_metadata(entry)
            sample_number = metadata["Sample Number"]
            test_period = metadata["Test Period"]
            if not sample_number or not test_period:
                continue
            signature = (sample_number, test_period)
            entries_by_metadata.setdefault(signature, {}).setdefault(entry, []).append(
                f"{chr(65 + r)}{c + 1}"
            )

    warnings = []
    for (sample_number, test_period), entries in entries_by_metadata.items():
        if len(entries) < 2:
            continue
        competitors = "; ".join(
            f"{', '.join(wells)}: {entry}"
            for entry, wells in entries.items()
        )
        warnings.append(
            f"Different source entries share Sample Number {sample_number!r} and "
            f"Test Period {test_period!r}: {competitors}. The entries remain separate; "
            "review their metadata against the source workbook."
        )
    return warnings


def parse_source_metadata(value):
    """Read the expected source-entry fields while retaining the raw tokens."""
    tokens = value.split()
    dilution = ""
    content_end = len(tokens)
    if tokens and re.fullmatch(DILUTION_CANDIDATE_PATTERN, tokens[-1]):
        dilution = tokens[-1]
        content_end -= 1

    period_start = next(
        (i for i, token in enumerate(tokens[:content_end])
         if token.lower() in {"day", "pnd", "week"}),
        None,
    )
    if period_start is None:
        prefix = tokens[:min(5, content_end)]
        period_tokens = tokens[5:content_end] if content_end > 5 else []
    else:
        prefix = tokens[:period_start]
        period_tokens = tokens[period_start:content_end]

    values = dict(zip(SOURCE_PREFIX_FIELDS[:4], prefix[:4]))
    values["Matrix/Split"] = " ".join(prefix[4:]) if len(prefix) >= 5 else ""
    for field in SOURCE_PREFIX_FIELDS:
        values.setdefault(field, "")
    values["Test Period"] = " ".join(period_tokens)
    values["Dilution Fold"] = dilution
    values["prefix_complete"] = (
        len(prefix) >= len(SOURCE_PREFIX_FIELDS)
        and all(values[field] for field in SOURCE_PREFIX_FIELDS)
    )
    values["prefix_tokens"] = prefix
    return values


def source_entry_issues(value):
    """Return required metadata still missing from a user-corrected entry."""
    if classify(txt(value))[0] != "sample":
        return ["Sample source entry"]
    metadata = parse_source_metadata(txt(value))
    issues = []
    if not metadata["prefix_complete"]:
        issues.append("Study Number, Custom ID, Sample Number, Group Number, and Matrix/Split")
    else:
        for field in SOURCE_PREFIX_FIELDS:
            if not metadata[field]:
                issues.append(field)
    if not metadata["Test Period"]:
        issues.append("Test Period")
    if not metadata["Dilution Fold"]:
        issues.append("Dilution Fold")
    elif not is_numeric_value(metadata["Dilution Fold"]):
        issues.append("Dilution Fold (must be a number greater than zero)")
    return issues


def parse_fields(value):
    metadata = parse_source_metadata(value)
    return metadata["Sample Number"], metadata["Test Period"], metadata["Dilution Fold"]


def complete_source_entry(value, corrections):
    """Insert user-supplied missing fields into a well's source identity."""
    tokens = value.split()
    _, existing_period, _ = parse_fields(value)

    if "Sample Number" in corrections:
        while len(tokens) < 2:
            tokens.append("")
        tokens[2:3] = user_text(corrections["Sample Number"]).split()

    if "Test Period" in corrections and not existing_period:
        # The source layout places the test period after matrix and before the
        # final dilution value. Insert before a trailing token when present.
        insert_at = len(tokens) - 1 if len(tokens) > 5 else len(tokens)
        tokens[insert_at:insert_at] = user_text(corrections["Test Period"]).split()

    if "Dilution Fold" in corrections:
        corrected_dilution = user_text(corrections["Dilution Fold"])
        if tokens and re.fullmatch(DILUTION_CANDIDATE_PATTERN, tokens[-1]):
            tokens[-1] = corrected_dilution
        else:
            tokens.append(corrected_dilution)

    return " ".join(token for token in tokens if token)


def find_source_plate(ws):
    # Source examples have column numbers 1..12 in row 2 and well data in rows 3..10.
    candidates = []
    for header_row in range(1, min(ws.max_row, 8) + 1):
        values = [ws.cell(header_row, c).value for c in range(1, ws.max_column + 1)]
        numeric = [c for c, v in enumerate(values, 1)
                   if isinstance(v, (int, float)) and float(v).is_integer() and 1 <= int(v) <= 12]
        if len(set(int(ws.cell(header_row, c).value) for c in numeric)) >= 10:
            # Usually A is row label, B:M are wells; tolerate headers starting in A.
            start_col = min(numeric)
            if start_col == 2:
                data_start_col = 2
            elif start_col == 1 and ws.max_column >= 13:
                data_start_col = 2
            else:
                data_start_col = start_col
            candidates.append((header_row + 1, data_start_col, header_row))
    if candidates:
        # Preserve the existing first-match behavior, but report other plausible
        # header rows so the user can review the selection.
        data_start_row, data_start_col, selected_header = candidates[0]
        competing_rows = [candidate[2] for candidate in candidates[1:]]
        warnings = []
        if competing_rows:
            all_rows = [selected_header, *competing_rows]
            warnings.append(
                f"Multiple possible plate header rows were found (worksheet rows "
                f"{', '.join(map(str, all_rows))}). The processor used row {selected_header}; "
                "review the result against the source workbook."
            )

        header_numbers = []
        for column in range(1, ws.max_column + 1):
            value = ws.cell(selected_header, column).value
            if isinstance(value, (int, float)) and float(value).is_integer():
                header_numbers.append((int(value), column))
        found_numbers = {number for number, _ in header_numbers}
        missing_numbers = [number for number in range(1, NCOLS + 1) if number not in found_numbers]
        extra_numbers = [number for number, _ in header_numbers if number < 1 or number > NCOLS]
        duplicate_numbers = sorted({
            number for number, _ in header_numbers
            if sum(1 for candidate, _ in header_numbers if candidate == number) > 1
        })
        ordered_numbers = [
            number for number, _ in sorted(header_numbers, key=lambda item: item[1])
            if 1 <= number <= NCOLS
        ]
        if missing_numbers or extra_numbers or duplicate_numbers or ordered_numbers != list(range(1, NCOLS + 1)):
            details = []
            if missing_numbers:
                details.append(f"missing column number(s) {', '.join(map(str, missing_numbers))}")
            if extra_numbers:
                details.append(f"extra column number(s) {', '.join(map(str, sorted(set(extra_numbers))))}")
            if duplicate_numbers:
                details.append(f"repeated column number(s) {', '.join(map(str, duplicate_numbers))}")
            if ordered_numbers != list(range(1, NCOLS + 1)) and not missing_numbers:
                details.append("column numbers are not in 1–12 order")
            warnings.append(
                "The detected plate header does not contain exactly one column numbered "
                f"1–12 ({'; '.join(details)}). Review the worksheet layout and output."
            )

        row_label_column = data_start_col - 1 if data_start_col > 1 else None
        if row_label_column is not None:
            row_labels = [
                txt(ws.cell(data_start_row + offset, row_label_column).value).upper()
                for offset in range(NROWS)
            ]
            expected_row_labels = [chr(65 + offset) for offset in range(NROWS)]
            if row_labels != expected_row_labels:
                warnings.append(
                    "The detected plate area does not have row labels A–H in order "
                    f"(found {', '.join(label or '(blank)' for label in row_labels)}). "
                    "Review the worksheet layout and output."
                )
        elif ws.max_row < data_start_row + NROWS - 1:
            available = max(0, ws.max_row - data_start_row + 1)
            warnings.append(
                f"The detected plate has only {available} data row(s); 8 are expected. "
                "Review the worksheet layout and output."
            )
        else:
            empty_rows = [
                data_start_row + offset
                for offset in range(NROWS)
                if all(
                    ws.cell(data_start_row + offset, data_start_col + column).value in (None, "")
                    for column in range(NCOLS)
                )
            ]
            if empty_rows:
                warnings.append(
                    "The detected plate has an entirely blank data row at worksheet row(s) "
                    f"{', '.join(map(str, empty_rows))}. This could be an intentionally empty "
                    "plate row or a missing row; review the source."
                )

        extra_rows = []
        populated_rows = []
        first_extra_row = data_start_row + NROWS
        for row in range(first_extra_row, ws.max_row + 1):
            label = (
                txt(ws.cell(row, row_label_column).value).upper()
                if row_label_column is not None else ""
            )
            populated = sum(
                ws.cell(row, data_start_col + offset).value not in (None, "")
                for offset in range(NCOLS)
            )
            if populated:
                populated_rows.append(row)
                if len(label) == 1 and "I" <= label <= "Z":
                    extra_rows.append(label)
        if populated_rows:
            row_text = ", ".join(map(str, populated_rows))
            label_text = f" (labeled {', '.join(extra_rows)})" if extra_rows else ""
            warnings.append(
                "Non-empty cells were found below the expected 8-row plate area at worksheet "
                f"row(s) {row_text}{label_text}. The output includes only rows A–H; "
                "review the source."
            )

        extra_columns = []
        first_extra_column = data_start_col + NCOLS
        for column in range(first_extra_column, ws.max_column + 1):
            if any(
                ws.cell(data_start_row + offset, column).value not in (None, "")
                for offset in range(NROWS)
            ):
                extra_columns.append(column)
        if extra_columns:
            extra_column_names = ", ".join(
                openpyxl.utils.get_column_letter(column) for column in extra_columns
            )
            warnings.append(
                "Non-empty cells were found beyond the 12-column plate area in worksheet "
                f"column(s) {extra_column_names}. The output includes only columns 1–12; "
                "review the source."
            )
        return data_start_row, data_start_col, warnings
    raise ValueError("Could not find the plate header row (expected numbered columns 1–12).")


def label_for_index(index):
    # A..Z, AA, BB, CC ... ZZ, AAA, BBB ... (repeat the same letter).
    if index < 26:
        return chr(65 + index)
    index -= 26
    length = 2
    while index >= 26:
        index -= 26
        length += 1
    return chr(65 + index) * length


def safe_name(value):
    value = re.sub(r'[<>:"/\\\\|?*]+', "_", txt(value))
    value = re.sub(r"\s+", "_", value).strip(" ._")
    return value or "Plate"


def load_source_workbook(input_file):
    """Load .xlsx directly and read binary Excel 97-2003 .xls with xlrd.

    Legacy .xls files are deliberately routed to xlrd first. They must never be
    passed directly to openpyxl, which does not support the .xls format.
    """
    path = Path(input_file)
    suffix = path.suffix.lower()

    if suffix == ".xlsx":
        try:
            wb_in = openpyxl.load_workbook(path, data_only=True)
            return wb_in, wb_in.active
        except Exception as exc:
            raise ValueError(f"Could not open this .xlsx workbook: {exc}") from exc

    if suffix != ".xls":
        raise ValueError(
            "Unsupported file type. Please select an .xlsx or Excel 97-2003 .xls workbook."
        )

    # Most reliable path for classic binary .xls files. xlrd 2.x supports .xls.
    if xlrd is None:
        raise ValueError(
            "This is a legacy .xls file, but the xlrd package is not installed.\n\n"
            "Close this window and run the updated Run_Reformatter.bat, or open Command "
            "Prompt in this folder and run:\npython -m pip install --upgrade xlrd openpyxl"
        )

    try:
        legacy = xlrd.open_workbook(str(path), on_demand=True)
        if legacy.nsheets < 1:
            raise ValueError("The .xls workbook contains no worksheets.")
        sheet = legacy.sheet_by_index(0)
        wb_in = openpyxl.Workbook()
        src = wb_in.active
        src.title = (sheet.name[:31] or "Source")
        for r in range(sheet.nrows):
            for c in range(sheet.ncols):
                cell = sheet.cell(r, c)
                value = cell.value
                if cell.ctype == xlrd.XL_CELL_DATE:
                    try:
                        value = xlrd.xldate.xldate_as_datetime(value, legacy.datemode)
                    except Exception:
                        pass
                elif cell.ctype in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK, xlrd.XL_CELL_ERROR):
                    value = None
                elif cell.ctype == xlrd.XL_CELL_NUMBER and float(value).is_integer():
                    value = int(value)
                src.cell(r + 1, c + 1, value)
        legacy.release_resources()
        return wb_in, src
    except Exception as exc:
        raise ValueError(
            "Could not read this legacy .xls file with xlrd. Confirm that it is a real "
            "Excel 97-2003 workbook and is not password-protected.\n\n"
            f"Details: {exc}"
        ) from exc


def reformat(
    input_file, plate_name, well_values=None, study_override="",
    duplicate_groups_approved=False,
):
    well_values = well_values or {}
    inspection = inspect_workbook(input_file)
    missing = []
    invalid = []
    # User-entered study identifiers are preserved as typed (apart from
    # surrounding whitespace); suffix cleanup applies only to source extraction.
    normalized_study_override = user_text(study_override)
    if not inspection["study"] and not normalized_study_override:
        missing.append(inspection["missing"][0])
    for item in inspection["missing"]:
        if item["kind"] == "source_entry":
            entered = user_text(
                well_values.get(item["well"], {}).get("Source metadata entry", "")
            )
            if not entered:
                missing.append(item)
            else:
                issues = source_entry_issues(entered)
                if issues:
                    invalid.append({
                        **item, "value": entered,
                        "reason": "Corrected entry is still missing or has invalid metadata: "
                                  + ", ".join(issues),
                    })
            continue
        if item["kind"] != "well":
            continue
        entered = user_text(well_values.get(item["well"], {}).get(item["field"], ""))
        if not entered:
            missing.append(item)
        elif item["field"] == "Dilution Fold" and not is_numeric_value(entered):
            invalid.append({**item, "value": entered})
    if missing:
        raise MissingInformationError(missing)
    if invalid:
        raise InvalidInformationError(invalid)

    wb_in, src = load_source_workbook(input_file)
    data_row, data_col, warnings = find_source_plate(src)

    raw = []
    for r in range(NROWS):
        row = []
        for c in range(NCOLS):
            row.append(txt(src.cell(data_row + r, data_col + c).value))
        raw.append(row)

    study = get_study_number(src, raw) or normalized_study_override
    review_notes = []
    if not inspection["study"] and normalized_study_override:
        review_notes.append({
            "type": "Manual completion", "well": "",
            "field": "Study number", "source_value": "",
            "value": normalized_study_override,
            "displayed_value": normalized_study_override,
            "details": "Study number was entered by the user.",
        })

    # Apply corrections before grouping, assigning labels, or merging wells.
    required_fields = {
        (item["well"], item["field"])
        for item in inspection["missing"]
        if item["kind"] == "well"
    }
    completed_raw = [row.copy() for row in raw]
    sample_metadata = {}
    pending_review_notes = []
    for c in range(NCOLS):
        for r in range(NROWS):
            kind, _ = classify(raw[r][c])
            if kind == "sample":
                well = f"{chr(65 + r)}{c + 1}"
                source_entry_override = user_text(
                    well_values.get(well, {}).get("Source metadata entry", "")
                )
                if source_entry_override:
                    completed_entry = " ".join(source_entry_override.split())
                    completed_raw[r][c] = completed_entry
                    metadata = parse_source_metadata(completed_entry)
                    fields = {
                        field: metadata[field]
                        for field in ("Sample Number", "Test Period", "Dilution Fold")
                    }
                    if is_numeric_value(fields["Dilution Fold"]):
                        fields["Dilution Fold"] = normalize_dilution_value(fields["Dilution Fold"])
                    sample_metadata[(r, c)] = fields
                    pending_review_notes.append({
                        "type": "Manual completion", "well": well,
                        "field": "Source metadata entry", "source_value": raw[r][c],
                        "value": completed_entry, "displayed_value": completed_entry,
                        "details": (
                            "The user supplied a complete source entry because the missing "
                            "metadata field could not be identified safely. Verify it against "
                            "the source workbook."
                        ),
                    })
                    continue
                fields = dict(zip(
                    ("Sample Number", "Test Period", "Dilution Fold"),
                    parse_fields(raw[r][c]),
                ))
                corrections = {}
                for field, source_value in fields.items():
                    if (well, field) not in required_fields:
                        continue
                    entered = user_text(well_values.get(well, {}).get(field, ""))
                    corrections[field] = entered
                    fields[field] = entered
                    pending_review_notes.append({
                        "type": "Manual completion", "well": well,
                        "field": field, "source_value": source_value,
                        "value": entered,
                        "displayed_value": (
                            normalize_dilution_value(entered)
                            if field == "Dilution Fold" else entered
                        ),
                        "details": (
                            "Value was entered by the user; the lookup normalizes commas and leading zeros. "
                            "Verify against the source workbook."
                            if field == "Dilution Fold"
                            else "Value was entered by the user; verify against the source workbook."
                        ),
                    })
                completed_entry = complete_source_entry(raw[r][c], corrections)
                completed_raw[r][c] = completed_entry
                if is_numeric_value(fields["Dilution Fold"]):
                    fields["Dilution Fold"] = normalize_dilution_value(fields["Dilution Fold"])
                sample_metadata[(r, c)] = fields

    samples, sample_group_for_coord = group_horizontal_samples(completed_raw)
    sample_labels = {key: label_for_index(i) for i, key in enumerate(samples)}
    warnings.extend(sample_identity_conflict_warnings(completed_raw))
    review_notes.extend(pending_review_notes)
    manually_completed_wells = sorted({
        note["well"] for note in review_notes
        if note["type"] == "Manual completion" and note.get("well")
    })
    manually_completed_fields = sorted({
        note["field"] for note in review_notes
        if note["type"] == "Manual completion"
    })
    if manually_completed_wells:
        study_text = (
            " The study number was also entered manually."
            if "Study number" in manually_completed_fields else ""
        )
        warnings.append(
            "Missing metadata was completed for well(s) "
            f"{', '.join(manually_completed_wells)}. Review the original and entered values "
            "on the Review Notes sheet before using the workbook." + study_text
        )
    elif any(note["field"] == "Study number" for note in review_notes):
        warnings.append(
            "The study number was entered manually. Review the value on the Review Notes "
            "sheet against the source workbook."
        )
    # Control Blank numbers increment for each horizontal run of adjacent Reagent wells.
    control_labels = {}
    control_num = 0
    for r in range(NROWS):
        c = 0
        while c < NCOLS:
            kind, _ = classify(raw[r][c])
            if kind != "control":
                c += 1
                continue
            end = c + 1
            while end < NCOLS and classify(raw[r][end])[0] == "control":
                end += 1
            control_num += 1
            for cc in range(c, end):
                control_labels[(r, cc)] = f"Control Blank {control_num}"
            c = end

    out = openpyxl.Workbook()
    ws = out.active
    ws.title = "Plate Map + Lookup"

    # Study number and plate name are the only frozen rows (freeze panes at A3).
    # Keep these headings plain, without fill or borders.
    no_border = Border()
    ws.merge_cells("A1:M1")
    ws["A1"] = f"Study Number: {study}"
    ws["A1"].font = Font(name="Arial", size=14, bold=True)
    ws["A1"].alignment = Alignment(horizontal="left", vertical="center")
    ws["A1"].fill = PatternFill(fill_type=None)
    ws["A1"].border = no_border
    ws.row_dimensions[1].height = 25

    ws.merge_cells("A2:M2")
    ws["A2"] = f"Plate Name: {plate_name}"
    ws["A2"].font = Font(name="Arial", size=14, bold=True)
    ws["A2"].alignment = Alignment(horizontal="left", vertical="center")
    ws["A2"].fill = PatternFill(fill_type=None)
    ws["A2"].border = no_border
    ws.row_dimensions[2].height = 25

    # Column headers and plate map.
    for c in range(NCOLS):
        cell = ws.cell(3, c + 2, c + 1)
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor=FILL["header"])
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = grid
    for r in range(NROWS):
        row_cell = ws.cell(4 + r, 1, chr(65 + r))
        row_cell.font = Font(bold=True)
        row_cell.fill = PatternFill("solid", fgColor=FILL["header"])
        row_cell.alignment = Alignment(horizontal="center", vertical="center")
        row_cell.border = grid

        for c in range(NCOLS):
            kind, normalized = classify(raw[r][c])
            cell = ws.cell(4 + r, 2 + c)
            cell.border = grid
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.fill = PatternFill("solid", fgColor=FILL[kind])
            if kind == "empty":
                cell.value = "Empty"
            elif kind == "control":
                cell.value = control_labels.get((r, c), "Control Blank")
            elif kind in ("standard", "qc"):
                cell.value = normalized
            else:
                cell.value = sample_labels[sample_group_for_coord[(r, c)]]

    # Merge adjacent wells when their classifications and displayed labels match.
    # This intentionally combines distinct source entries that normalize to the
    # same displayed standard or QC label.
    for r in range(NROWS):
        c = 0
        while c < NCOLS:
            cell = ws.cell(4 + r, 2 + c)
            value = cell.value
            kind, _ = classify(raw[r][c])
            if kind == "empty":
                c += 1
                continue
            end = c + 1
            while end < NCOLS:
                next_kind, _ = classify(raw[r][end])
                next_value = ws.cell(4 + r, 2 + end).value
                if next_kind != kind or next_value != value:
                    break
                end += 1
            if end - c > 1:
                ws.merge_cells(start_row=4+r, start_column=2+c, end_row=4+r, end_column=1+end)
                ws.cell(4+r, 2+c).border = grid
                # Outline the merged region after merging.
                for cc in range(c, end):
                    ws.cell(4+r, 2+cc).border = grid
            c = end

    # Resolve lookup metadata by the corrected source identity. Identical
    # completed entries share a lookup row only when they are horizontal neighbors.
    lookup_groups = []
    for key, coords in samples.items():
        metadata_key = tuple(sample_metadata[coord][field] for coord in coords[:1]
                             for field in ("Sample Number", "Test Period", "Dilution Fold"))
        wells = [f"{chr(65 + r)}{c + 1}" for r, c in coords]
        lookup_groups.append((sample_labels[key], wells, metadata_key))

    repeated_lookup_details = {}
    for label, wells, metadata_key in lookup_groups:
        if all(metadata_key):
            comparison_key = tuple(value.casefold() for value in metadata_key)
            repeated_lookup_details.setdefault(
                comparison_key, {"metadata": metadata_key, "rows": []}
            )["rows"].append((label, wells))
    for group in repeated_lookup_details.values():
        sample_number, test_period, dilution = group["metadata"]
        rows = group["rows"]
        if len(rows) > 1:
            competing_rows = "; ".join(
                f"{label} (wells {', '.join(wells)})" for label, wells in rows
            )
            warnings.append(
                "Separate lookup rows share Sample Number "
                f"{sample_number!r}, Test Period {test_period!r}, and Dilution "
                f"{dilution!r}: {competing_rows}. Review these rows against the source workbook."
            )
            if duplicate_groups_approved:
                review_notes.append({
                    "type": "Duplicate lookup approval",
                    "well": ", ".join(well for _, wells in rows for well in wells),
                    "field": "Sample Number / Test Period / Dilution Fold",
                    "source_value": "",
                    "value": "Approved to continue",
                    "displayed_value": "Approved",
                    "details": (
                        f"Approved after review: Sample Number {sample_number!r}, Test Period "
                        f"{test_period!r}, Dilution {dilution!r}. Competing lookup rows: "
                        f"{competing_rows}."
                    ),
                })

    # Lookup table below plate map.
    lookup_title_row = 14
    ws.merge_cells(start_row=lookup_title_row, start_column=1, end_row=lookup_title_row, end_column=5)
    title_cell = ws.cell(lookup_title_row, 1, "Sample Lookup")
    title_cell.font = Font(bold=True, size=11)
    title_cell.fill = PatternFill(fill_type=None)
    title_cell.border = Border(left=thin, right=thin, bottom=thin)

    headers = ["Label", "Well(s)", "Sample Number", "Test Period", "Dilution Fold"]
    # The lookup uses merged column groups so every plate well column B:M
    # keeps the same width while the lookup fields remain readable.
    groups = [(1, 2), (3, 4), (5, 6), (7, 10), (11, 13)]
    for header, (c1, c2) in zip(headers, groups):
        ws.merge_cells(start_row=lookup_title_row + 1, start_column=c1,
                       end_row=lookup_title_row + 1, end_column=c2)
        cell = ws.cell(lookup_title_row + 1, c1, header)
        cell.font = Font(bold=True)
        cell.fill = PatternFill(fill_type=None)
        cell.border = grid
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for cc in range(c1, c2 + 1):
            ws.cell(lookup_title_row + 1, cc).border = grid

    # Alternate subtle neutral fills across sample lookup data rows for readability.
    # Because merged cells can discard styles applied to non-anchor cells, apply
    # banding to the anchor cells after each merge and then reapply it in a final pass.
    lookup_row_fills = ("FFFFFF", "E6E6E6")  # white, slightly darker grey

    # Keep wells ordered by row-major position in the lookup; sample labels themselves are assigned column-major.
    for sample_index, (label, wells, metadata) in enumerate(lookup_groups):
        idx = lookup_title_row + 2 + sample_index
        row_fill = PatternFill(fill_type="solid", fgColor=lookup_row_fills[sample_index % 2])
        sample_no, test_period, dilution = metadata
        values = [label, ", ".join(wells), sample_no, test_period, dilution]
        for value, (c1, c2) in zip(values, groups):
            ws.merge_cells(start_row=idx, start_column=c1, end_row=idx, end_column=c2)
            cell = ws.cell(idx, c1, value)
            cell.border = grid
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            cell.fill = row_fill
            for cc in range(c1, c2 + 1):
                ws.cell(idx, cc).border = grid
                ws.cell(idx, cc).fill = row_fill

    # Reapply alternating fills after all merged ranges are created. Excel stores
    # the visible fill for each merged field on its top-left (anchor) cell.
    for sample_index in range(len(lookup_groups)):
        idx = lookup_title_row + 2 + sample_index
        fill = PatternFill(fill_type="solid", fgColor=lookup_row_fills[sample_index % 2])
        for c1, c2 in groups:
            ws.cell(idx, c1).fill = fill
            # Set edge cells too so the stripe remains continuous in viewers that
            # render merged ranges using individual cell styles.
            for cc in range(c1 + 1, c2 + 1):
                ws.cell(idx, cc).fill = fill

    # All 12 well columns stay equal width; lookup fields span multiple columns.
    ws.column_dimensions["A"].width = 10
    for c in range(2, 14):
        ws.column_dimensions[get_column_letter(c)].width = 7
    ws.row_dimensions[lookup_title_row + 1].height = 30
    ws.row_dimensions[3].height = 20
    for r in range(4, 12):
        ws.row_dimensions[r].height = 28

    ws.freeze_panes = "A3"
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.page_setup.orientation = "landscape"
    ws.print_area = f"A1:M{max(lookup_title_row + 1 + len(lookup_groups), 12)}"

    if review_notes or warnings:
        notes_ws = out.create_sheet("Review Notes")
        notes_ws.append([
            "Type", "Well", "Field", "Source value",
            "Entered value", "Displayed value", "Details",
        ])
        for note in review_notes:
            notes_ws.append([
                note["type"], note["well"], note["field"],
                note.get("source_value", ""), note["value"],
                note.get("displayed_value", note["value"]), note["details"],
            ])
        for warning in warnings:
            notes_ws.append(["Processing warning", "", "", "", "", "", warning])
        notes_ws.freeze_panes = "A2"
        notes_ws.auto_filter.ref = notes_ws.dimensions
        for column, width in {
            "A": 22, "B": 12, "C": 12, "D": 20, "E": 18,
            "F": 18, "G": 18, "H": 78,
        }.items():
            notes_ws.column_dimensions[column].width = width
        for cell in notes_ws[1]:
            cell.font = Font(bold=True)
            cell.fill = PatternFill("solid", fgColor=FILL["header"])

    output_path = Path(input_file).with_name(f"{safe_name(study)}_{safe_name(plate_name)}.xlsx")
    # Avoid overwriting the source or an existing file without warning.
    if output_path.resolve() == Path(input_file).resolve():
        output_path = Path(input_file).with_name(f"{safe_name(study)}_{safe_name(plate_name)}_reformatted.xlsx")
    out.save(output_path)
    return output_path, study, len(samples), warnings


def main():
    root = tk.Tk()
    root.withdraw()
    root.update()

    selected = filedialog.askopenfilename(
        title="Select source plate Excel workbook",
        filetypes=[("Excel workbooks (.xlsx, .xls)", "*.xlsx *.xls"), ("All files", "*.*")]
    )
    if not selected:
        return

    plate_name = simpledialog.askstring("Plate Name", "Enter the plate name:")
    if plate_name is None or not plate_name.strip():
        messagebox.showwarning("Plate name required", "No plate name was entered. No file was created.")
        return

    try:
        output, study, count, warnings = reformat(selected, plate_name.strip())
        warning_text = ""
        if warnings:
            warning_text = "\n\nWarnings for review:\n- " + "\n- ".join(warnings)
        messagebox.showinfo(
            "Reformat complete",
            f"Study Number: {study}\nPlate Name: {plate_name.strip()}\n"
            f"Grouped sample labels: {count}\n\nSaved file:\n{output}{warning_text}"
        )
    except Exception as exc:
        messagebox.showerror("Reformat failed", f"The workbook could not be reformatted:\n\n{exc}")
        raise


if __name__ == "__main__":
    main()
