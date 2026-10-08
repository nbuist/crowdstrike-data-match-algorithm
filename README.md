# crowdstrike-data-match-algorithm

Created for internship project in collaboration with Claude.

CrowdStrike / Inventory hostname matcher.

## Expected Layout

Row 1 contains the headers.

| Column | Header   | Source                         |
|--------|----------|--------------------------------|
| A      | Hostname | CrowdStrike                    |
| B      | Last Name| Inventory                      |
| C      | Name     | Inventory                      |
| D      | ID#      | Inventory (5-character HSI)    |

A hostname matches when its last 5 characters equal an ID# in column D (case-insensitive, surrounding spaces ignored).

## Usage

```bash
# Adds/replaces a "Matched" tab in data.xlsx
python match_hostnames.py data.xlsx

# Writes to a new file instead
python match_hostnames.py data.xlsx -o result.xlsx

# Pick the source tab (default: first tab)
python match_hostnames.py data.xlsx --sheet Sheet1

# Writes data_matched.csv
python match_hostnames.py data.csv
```

## CSV Output

Every row has a `Status` column:

| Status                     | Meaning                                              |
|----------------------------|------------------------------------------------------|
| `Matched`                  | Hostname whose last 5 chars match an ID#             |
| `Unmatched - hostname only`| Hostname with no matching ID#                        |
| `Unmatched - inventory only`| Inventory item no hostname matched                  |

## Requirements

- [openpyxl](https://pypi.org/project/openpyxl/)

```bash
pip install openpyxl
```
