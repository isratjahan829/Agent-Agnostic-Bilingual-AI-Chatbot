#!/usr/bin/env python3
"""Score returned hallucination-annotation workbooks.

Usage:
    python scripts/score_annotation.py artifacts/reviewer/annotation_package

Reads annotation_A1.xlsx / A2 / A3 (the "Items" sheet, "label" column), and
prints the per-category rates, the majority-label rate, Fleiss' kappa and the
share of unanimous items — the three numbers Section 3.6.1 needs.
"""
from __future__ import annotations

import sys
from collections import Counter
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from openpyxl import load_workbook

from banglafingpt.eval.annotation import CATEGORIES
from banglafingpt.eval.human_eval import fleiss_kappa, interpret_kappa


def read_labels(path: Path) -> dict[str, str]:
    ws = load_workbook(path, read_only=True)["Items"]
    rows = ws.iter_rows(values_only=True)
    header = [str(c or "") for c in next(rows)]
    i_id, i_label = header.index("item_id"), header.index("label")
    out = {}
    for row in rows:
        if row is None or row[i_id] is None:
            continue
        label = (row[i_label] or "").strip()
        if label:
            out[str(row[i_id])] = label
    return out


def main(folder: str) -> int:
    d = Path(folder)
    sheets = {p.stem.split("_")[-1]: read_labels(p) for p in sorted(d.glob("annotation_*.xlsx"))}
    if not sheets:
        print(f"no annotation_*.xlsx found in {d}")
        return 1

    for name, labels in sheets.items():
        bad = sorted({v for v in labels.values()} - set(CATEGORIES))
        print(f"{name}: {len(labels)} labelled" + (f"  UNRECOGNISED: {bad}" if bad else ""))
    common = set.intersection(*(set(s) for s in sheets.values()))
    if not common:
        print("\nno item is labelled by every annotator yet — nothing to score")
        return 1
    print(f"\n{len(common)} items labelled by all {len(sheets)} annotators\n")

    matrix, unanimous, majority = [], 0, Counter()
    for item in sorted(common):
        votes = [sheets[a][item] for a in sheets]
        matrix.append([votes.count(c) for c in CATEGORIES])
        if len(set(votes)) == 1:
            unanimous += 1
        majority[Counter(votes).most_common(1)[0][0]] += 1

    n = len(common)
    print("Majority label distribution")
    for c in CATEGORIES:
        print(f"  {c:20s} {majority[c]:4d}  {100 * majority[c] / n:5.1f}%")

    kappa = fleiss_kappa(matrix)
    print(f"\nFleiss' kappa       {kappa:.3f}  ({interpret_kappa(kappa)})")
    print(f"Unanimous items     {unanimous}/{n}  {100 * unanimous / n:.1f}%")
    print("\nPaste into Section 3.6.1: "
          f"Fleiss' kappa = {kappa:.2f}, with {100 * unanimous / n:.0f}% of items unanimous.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1] if len(sys.argv) > 1 else "artifacts/reviewer/annotation_package"))
