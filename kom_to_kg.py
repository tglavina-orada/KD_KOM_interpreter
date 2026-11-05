"""
Convert invoiced quantities measured in KOM (pieces) into kilograms.

For every row whose sales unit ("Prodajna jedinica") is KOM, the unit weight is
read from the item description ("Opis"), e.g.

    BROK SLANI FILETI INĆUNA 150g            -> 0.15 kg / kom
    BROK M. FIL. INĆUNA S LIMUNOM 150g/8 kom -> 0.15 kg / kom
    INĆUN MARINIRANI 4,5kg  U MARINADI       -> 4.5  kg / kom
    Brancin filet 90-130g, zamrznuti - 1kg   -> 1    kg / kom  (90-130g is a size range)
    ORATA EVISCERATA 300/400,ATM,PAIO        -> 0.7  kg / kom  (pair: 2 x 350 g midpoint)

Rows already in KG get a factor of 1. Two columns are appended:
    "KG po jedinici"  - kg per one sales unit
    "Količina (KG)"   - Fakturirana količina * KG po jedinici

Usage:
    python kom_to_kg.py "input.xlsx" [-o output.xlsx]
"""

from __future__ import annotations

import argparse
import re
import sys
from copy import copy
from pathlib import Path

import openpyxl
from openpyxl.styles import PatternFill

COL_DESC = "Opis"
COL_QTY = "Fakturirana količina"
COL_UNIT = "Prodajna jedinica"
COL_MATERIAL = "Materijal"
OUT_PER_UNIT = "KG po jedinici"
OUT_TOTAL = "Količina (KG)"

# Manual overrides by material number, for descriptions that cannot be parsed.
# Example: {"1500123": 0.25}
OVERRIDES: dict[str, float] = {}

# A number followed by g / kg, e.g. "150g", "4,5kg", "200 g", "2 kg".
# Not preceded by a digit, '-' or '/' so size ranges like "90-130g" or
# "300/400g" are skipped. Not followed by a letter (so "kg" in words is ignored).
WEIGHT_RE = re.compile(
    r"(?<![\d/\-])(\d+(?:[.,]\d+)?)\s*(kg|g)(?![a-zčćšđž])",
    re.IGNORECASE,
)
# Size range such as "300/400" used for whole fish sold in pairs.
RANGE_RE = re.compile(r"(\d+)\s*/\s*(\d+)")
PAIR_RE = re.compile(r"\b(PAIO|PAR)\b", re.IGNORECASE)


def weight_from_description(desc: str) -> float | None:
    """Return kg per piece parsed from the description, or None if unknown."""
    if not desc:
        return None
    matches = WEIGHT_RE.findall(desc)
    if matches:
        # Take the last explicit weight (e.g. "... 90-130g, zamrznuti - 1kg").
        value, unit = matches[-1]
        kg = float(value.replace(",", "."))
        if unit.lower() == "g":
            kg /= 1000
        return round(kg, 6)

    # Fallback: pair of whole fish with a size range, e.g. "300/400,ATM,PAIO".
    if PAIR_RE.search(desc):
        rng = RANGE_RE.search(desc)
        if rng:
            lo, hi = int(rng.group(1)), int(rng.group(2))
            return round(2 * (lo + hi) / 2 / 1000, 6)
    return None


def to_number(value) -> float | None:
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)):
        return float(value)
    try:
        return float(str(value).strip().replace(",", "."))
    except ValueError:
        return None


def process(input_path: Path, output_path: Path) -> None:
    wb = openpyxl.load_workbook(input_path)
    ws = wb.active

    header = [c.value for c in ws[1]]
    try:
        i_desc = header.index(COL_DESC)
        i_qty = header.index(COL_QTY)
        i_unit = header.index(COL_UNIT)
    except ValueError as e:
        sys.exit(f"Missing expected column: {e}")
    i_mat = header.index(COL_MATERIAL) if COL_MATERIAL in header else None

    # Append (or reuse, if the script is run on its own output) the two columns.
    def col_for(name: str) -> int:
        if name in header:
            return header.index(name) + 1
        header.append(name)
        return len(header)

    c_per_unit = col_for(OUT_PER_UNIT)
    c_total = col_for(OUT_TOTAL)
    for c, name in ((c_per_unit, OUT_PER_UNIT), (c_total, OUT_TOTAL)):
        cell = ws.cell(row=1, column=c, value=name)
        src = ws.cell(row=1, column=i_qty + 1)
        cell.font = copy(src.font)
        cell.fill = copy(src.fill)
        cell.alignment = copy(src.alignment)
        cell.border = copy(src.border)

    warn_fill = PatternFill("solid", fgColor="FFC7CE")
    unresolved: dict[str, int] = {}
    kom_rows = 0

    for row in ws.iter_rows(min_row=2):
        unit = str(row[i_unit].value or "").strip().upper()
        if not unit:
            continue
        desc = str(row[i_desc].value or "").strip()
        qty = to_number(row[i_qty].value)
        material = str(row[i_mat].value or "").strip() if i_mat is not None else ""

        if unit == "KG":
            factor = 1.0
        elif unit == "KOM":
            kom_rows += 1
            factor = OVERRIDES.get(material) or weight_from_description(desc)
        else:
            factor = None

        r = row[0].row
        per_cell = ws.cell(row=r, column=c_per_unit)
        tot_cell = ws.cell(row=r, column=c_total)
        if factor is None:
            unresolved[f"{unit} | {material} | {desc}"] = unresolved.get(f"{unit} | {material} | {desc}", 0) + 1
            per_cell.value = None
            tot_cell.value = None
            per_cell.fill = warn_fill
            tot_cell.fill = warn_fill
            continue
        per_cell.value = factor
        tot_cell.value = round(qty * factor, 6) if qty is not None else None
        per_cell.number_format = "0.000"
        tot_cell.number_format = "#,##0.000"

    ws.column_dimensions[openpyxl.utils.get_column_letter(c_per_unit)].width = 15
    ws.column_dimensions[openpyxl.utils.get_column_letter(c_total)].width = 15

    wb.save(output_path)

    print(f"Saved: {output_path}")
    print(f"KOM rows converted: {kom_rows - sum(v for k, v in unresolved.items() if k.startswith('KOM'))}/{kom_rows}")
    if unresolved:
        print("\nCould not determine weight (highlighted red; add to OVERRIDES by material number):")
        for k, n in sorted(unresolved.items()):
            print(f"  {n:5d} x  {k}")


def main() -> None:
    ap = argparse.ArgumentParser(description="Convert KOM quantities to KG using the weight in the item description.")
    ap.add_argument("input", type=Path, help="Input .xlsx file (SAP export)")
    ap.add_argument("-o", "--output", type=Path, help="Output .xlsx (default: <input>_KG.xlsx)")
    args = ap.parse_args()

    out = args.output or args.input.with_name(args.input.stem + "_KG.xlsx")
    process(args.input, out)


if __name__ == "__main__":
    main()
