# Evaluation-Slot-Comparison-Tool

A desktop tool that compares two Excel Evaluation Slot Detail files and generates a structured multi-sheet report — built to support the Petrotech and Investment teams in tracking changes between MTM runs.

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
