# Evaluation-Slot-Comparison-Tool

The Evaluation Slot Comparison Tool is an internal desktop application built in Python that automates the comparison of two Excel evaluation files generated from MTM (Mark-to-Market) runs. Designed for the Petrotech and Investment teams, this tool eliminates the need for manual cross-referencing by producing a structured, multi-sheet Excel report in just a few clicks.

## What It Does

Upload two evaluation files and the tool produces an Excel report covering:

- **Summary** — UWI counts by slot category side by side with deltas and a full list of added/dropped UWIs
- **LTD** — Side-by-side key comparisons with match flags
- **DUC / Permit / PDP** — Unique UWI lists with Slot ID, Well Name and match flags
- **Asset Analysis** — Net Acreage and Royalty Rate + ORRI comparison per asset with deltas
- **Category Changes** — UWIs that shifted slot category between runs (e.g. DUC → PDP)
- **Raw Data** — Unmodified Slots Allocations sheet from both files for audit purposes

## How to Run

```bash
pip install pandas openpyxl
python "Fund Comparison.py"
```

A popup window will guide you through the rest — no command line needed after launch.

## Input Requirements

- File format: `.xlsx` or `.xls`
- Both files must contain a sheet named **`Slots Allocations`**
- Filenames should include `eval_XXXXX` for short name extraction

## Built With

- Python · pandas · openpyxl · tkinter

---

> 📖 Full user guide available on [Confluence](#)
