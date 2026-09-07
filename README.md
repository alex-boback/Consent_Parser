# Consent parser

Define any number of IRBs, choose which Consent values count for each, and merge
consenting participants by Computing ID. Requires Python 3.10+.

## Input schema

Every CSV or Excel sheet must contain these headers:

```text
First name,Last name,email,Computing ID,Subject ID,Professor,Consent
```

Header order and capitalization do not matter. Extra columns are ignored.
Supported inputs: `.csv`, `.xlsx`, and `.xlsm`.

## Configure and run

Copy `irbs.example.json` to `irbs.json`. Add one entry per IRB:

```json
[
  {"name": "6614", "file": "6614.xlsx", "consent_values": ["Both"]},
  {"name": "6881", "file": "6881.xlsx", "consent_values": ["Both", "Survey Only"]}
]
```

These are example consent rules; edit them to match your studies. Only the values
explicitly listed count as consent. Blank Consent never qualifies.
File paths are relative to the JSON file, or may be absolute.
Each IRB must have a unique name.

```powershell
python -m pip install -r requirements.txt
python consent_parser.py irbs.json -o consented.csv
```

An entry can optionally include `"sheet": "Participants"` and `"header_row": 2`.
Defaults are the first worksheet and header row 1. Multiple entries can use
different worksheets in the same workbook. Do not specify `sheet` for CSV files.

## Results

The output has one row per consenting Computing ID, with these columns:

```text
First name,Last name,email,Computing ID,Subject ID,Professor,IRBs
```

- Include people who consent to at least one configured IRB.
- Merge only matching, nonempty Computing IDs; email and names are never join keys.
- Comparisons ignore case and repeated/leading/trailing whitespace.
- Only consenting rows contribute participant details. Blank details can be filled
  from another consenting row. Identical duplicates are collapsed.
- `IRBs` lists the studies the person consented to, separated by semicolons.
- Conflicting nonblank participant details raise an error identifying the field,
  Computing ID, and IRB. Subject IDs and Professors must agree across merged rows.
- A consenting row without Computing ID raises an error.

Results are sorted by last name, first name, then Computing ID. Existing output
files are never overwritten. This uses any qualifying row, not latest-consent
status; resolve withdrawals in your input data first.

Excel formulas and error cells are rejected; use values-only inputs. Text IDs
and simple numeric zero-padding formats such as `00000` retain leading zeros.

## Python usage

```python
from consent_parser import IRB, parse_consent_files, write_csv

irbs = [
    IRB("6614", "6614.xlsx", ["Both"]),
    IRB("6881", "6881.xlsx", ["Both", "Survey Only"], sheet="Participants"),
    IRB("9000", "9000.csv", ["Yes"]),
]
people = parse_consent_files(irbs)
write_csv(people, "consented.csv")
```

Use `parse_irb(irb)` to get consenting rows for just one IRB.
The previous fixed two-file API and command have been replaced by this interface.
Run tests with `python -m unittest -v`.
