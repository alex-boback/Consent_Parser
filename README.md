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
First name,Last name,email,Computing ID,Professor,IRBs,Subject ID (6614),Subject ID (6881)
```

- Each configured IRB gets a `Subject ID (<IRB name>)` column, in configuration order.
  It is blank when the person has no consenting row or no Subject ID for that IRB.
  Subject IDs are never copied between IRBs.
- Include people who consent to at least one configured IRB.
- Merge only matching, nonempty Computing IDs; email and names are never join keys.
- Comparisons ignore case and repeated/leading/trailing whitespace.
- The text `None` (ignoring case and surrounding whitespace) is treated as blank.
  A real value from another consenting row fills it regardless of input order;
  if no real value exists, the output cell stays blank.
- Only consenting rows contribute participant details. Blank details can be filled
  from another consenting row. Identical duplicates are collapsed.
- `IRBs` lists the studies the person consented to, separated by semicolons.
- Conflicting nonblank participant details raise an error identifying the field,
  Computing ID, and IRB. Professors must agree across merged rows. Subject IDs
  may differ between IRBs; conflicting IDs within the same IRB still raise an error.
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
write_csv(people, "consented.csv", irbs)
```

Pass `irbs` to `write_csv` to retain all IRB columns even when the results are empty.

Use `parse_irb(irb)` to get consenting rows for just one IRB.
The previous fixed two-file API and command have been replaced by this interface.
Run tests with `python -m unittest -v`.

## Filter response data by consented Computing IDs

Use `response_parser.py` to keep response rows whose Computing ID appears in
the consent parser's output:

```powershell
python response_parser.py responses.xlsx consented.csv -o filtered_responses.xlsx
```

Both inputs accept `.csv`, `.xlsx`, or `.xlsm` files and must have exactly one
`Computing ID` column. Other columns may contain any response data. The output
is an Excel (`.xlsx`) file containing the original response headers and matching rows, in
their original order. Repeated responses are retained. Response text, including
whitespace and literal `None` answers, is preserved.

IDs and the Computing ID header are matched without regard to capitalization or
repeated/leading/trailing whitespace, as in the consent parser. Blank IDs and
`None` IDs never match. Membership in the consent file is the only filter; all
IRBs in that file qualify. No consent columns are added to the response data.

By default, each input uses its first worksheet and first row as the header.
For other worksheets or header positions:

```powershell
python response_parser.py responses.xlsx consented.csv -o filtered_responses.xlsx --response-sheet "Responses" --response-header-row 2
```

`--consent-sheet` and `--consent-header-row` configure the ID-list input in the
same way. Worksheet options apply only to Excel inputs. Excel inputs must be
values-only; formulas and error cells are rejected. Output stores values as text
in a worksheet named `Responses`, preserving leading zeros and literal answers
without copying source workbook formatting. The default output filename is
`filtered_responses.xlsx`. Existing output files are never overwritten. If no
responses match, the output contains only the response header.
