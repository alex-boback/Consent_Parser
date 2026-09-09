````markdown
# Consent Parser

The Consent Parser combines participant consent data from any number of IRBs and merges consenting participants using their **Computing ID**.

Each IRB can define which `Consent` values count as consent.

## Requirements

- Python 3.10+
- Git
- Input files in `.csv`, `.xlsx`, or `.xlsm` format

---

## 1. Clone the Repository

### Open a Terminal

Open **PowerShell**, **Command Prompt**, **Git Bash**, or another terminal.

### Navigate to the Folder Where You Want the Project

For example:

```powershell
cd C:\Code
```

### Clone the Repository

Run:

```powershell
git clone https://github.com/alex-boback/Consent_Parser.git
```


### Enter the Project Directory

```powershell
cd repository-name
```

### Verify Git Is Installed

You can check whether Git is installed by running:

```powershell
git --version
```

If Git is installed, this command will display the installed version.

---

## 2. Install Python Dependencies

From inside the project directory, run:

```powershell
python -m pip install -r requirements.txt
```

---

## 3. Prepare Input Files

### Required Headers

Every CSV or Excel sheet must contain the following headers:

```text
First name, Last name, email, Computing ID, Subject ID, Professor, Consent
```

The parser expects **all of these columns to be present**.

Header capitalization and order do not matter.

Extra columns are ignored.

### Handling Missing Fields

#### First Name, Last Name, Email, and Professor

If one of these fields is unavailable:

- Add the required column to the sheet.
- Fill missing values with `"None"`.
- `"None"` is treated as a blank value by the parser.
- If another consenting row with the same Computing ID contains a real value for that field, that value will be used in the output.
- If no other consenting row contains a value, the output field will remain blank.

#### Subject ID

If Subject IDs are not already available:

- Generate a unique Subject ID for each row.
- In Excel, one simple option is:

```excel
=ROW()-1
```

Subject IDs are tracked separately for each IRB. A participant may therefore have a different Subject ID for each IRB in which their Computing ID appears.

#### Computing ID

A Computing ID is **required** for every consenting entry.

- Remove entries that do not have a Computing ID before running the parser.
- The parser uses Computing ID as the primary key for matching and joining participants across files.
- Names and email addresses are never used as join keys.

### Supported File Types

The parser supports:

- `.csv`
- `.xlsx`
- `.xlsm`

---

## 4. Configure the IRBs

Copy:

```text
irbs.example.json
```

to:

```text
irbs.json
```

Add one entry for each IRB.

Example:

```json
[
  {
    "name": "6614",
    "file": "6614.xlsx",
    "consent_values": ["Both"]
  },
  {
    "name": "6881",
    "file": "6881.xlsx",
    "consent_values": ["Both", "Survey Only"]
  }
]
```

### Configuration Fields

Each IRB entry contains:

- `"name"` — Unique name used to identify the IRB.
- `"file"` — Path to the CSV or Excel file.
- `"consent_values"` — Values in the `Consent` column that count as consent.

Only values explicitly listed in `"consent_values"` count as consent.

For example:

```json
"consent_values": ["Both", "Survey Only"]
```

means that only rows containing `Both` or `Survey Only` qualify.

A blank `Consent` value never qualifies.

### File Paths

File paths may be:

- Relative to the location of the JSON configuration file, or
- Absolute paths

Each IRB must have a unique `"name"`.

### Excel Sheet Options

Excel entries may optionally specify a worksheet and header row:

```json
{
  "name": "6881",
  "file": "6881.xlsx",
  "consent_values": ["Both", "Survey Only"],
  "sheet": "Participants",
  "header_row": 2
}
```

If these options are omitted:

- The first worksheet is used.
- Header row 1 is used.

Different entries may reference different worksheets within the same workbook.

Do **not** specify `"sheet"` for CSV files.

---

## 5. Run the Parser

Run:

```powershell
python consent_parser.py irbs.json -o consented.csv
```

Where:

- `irbs.json` is your IRB configuration file.
- `consented.csv` is the desired output file.

Existing output files are **never overwritten**.

---

## 6. Results

The output contains one row for each unique consenting Computing ID.

For example, with IRBs `6614` and `6881`, the output columns would be:

```text
First name,Last name,email,Computing ID,Professor,IRBs,Subject ID (6614),Subject ID (6881)
```

### IRB Columns

Each configured IRB receives its own Subject ID column:

```text
Subject ID (<IRB name>)
```

For example:

```text
Subject ID (6614)
Subject ID (6881)
```

IRB columns appear in the same order as the IRBs in the configuration file.

A Subject ID field is blank when:

- The participant has no consenting row for that IRB, or
- The participant has no Subject ID for that IRB.

Subject IDs are **never copied between IRBs**.

### Participant Inclusion

A person is included in the output if they consent to **at least one configured IRB**.

Only consenting rows contribute participant information.

### Matching Participants

Participants are merged using matching, nonempty Computing IDs.

The parser does **not** use:

- First name
- Last name
- Email

as join keys.

### Value Comparisons

Comparisons ignore:

- Capitalization
- Leading whitespace
- Trailing whitespace
- Repeated whitespace

The text `"None"`, ignoring case and surrounding whitespace, is treated as blank.

For example, the following are treated as blank:

```text
None
none
 NONE
```

### Filling Missing Participant Information

If one consenting row contains `"None"` or a blank value while another consenting row for the same Computing ID contains a real value, the real value is used.

For example:

```text
IRB 6614: Professor = None
IRB 6881: Professor = Smith
```

produces:

```text
Professor = Smith
```

This behavior does not depend on input file order.

If no real value exists, the output field remains blank.

### Duplicate Rows

Identical duplicate records are collapsed.

### IRBs Column

The `IRBs` column lists every IRB for which the participant has a qualifying consent row.

Multiple IRBs are separated by semicolons.

For example:

```text
6614;6881
```

### Conflicting Data

Conflicting nonblank participant information raises an error.

The error identifies:

- The conflicting field
- The Computing ID
- The IRB

Professor values must agree across merged consenting rows.

Subject IDs may differ between different IRBs.

For example, this is valid:

```text
Subject ID (6614) = 104
Subject ID (6881) = 271
```

However, conflicting Subject IDs **within the same IRB** raise an error.

### Missing Computing IDs

A consenting row without a Computing ID raises an error.

### Sorting

Results are sorted by:

1. Last name
2. First name
3. Computing ID

### Consent Status

The parser uses **any qualifying consent row**.

It does not determine the participant's latest consent status.

If a participant has withdrawn consent, resolve the withdrawal in the input data before running the parser.

### Excel Formulas

Excel formulas and error cells are rejected.

Use **values-only** input files.

### Leading Zeros

Text IDs and simple numeric zero-padding formats retain leading zeros.

For example, an Excel format such as:

```text
00000
```

will preserve IDs such as:

```text
00123
```

---

## 7. Python Usage

The parser can also be used directly from Python.

```python
from consent_parser import IRB, parse_consent_files, write_csv

irbs = [
    IRB("6614", "6614.xlsx", ["Both"]),
    IRB(
        "6881",
        "6881.xlsx",
        ["Both", "Survey Only"],
        sheet="Participants"
    ),
    IRB("9000", "9000.csv", ["Yes"]),
]

people = parse_consent_files(irbs)

write_csv(people, "consented.csv", irbs)
```

Pass `irbs` to `write_csv` to retain all configured IRB columns even when the results are empty.

To retrieve consenting rows for only one IRB, use:

```python
parse_irb(irb)
```

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
