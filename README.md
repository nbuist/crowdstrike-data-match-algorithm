# crowdstrike-data-match-algorithm
created for internship project in collaboration with claude

CrowdStrike / Inventory hostname matcher.

Expected layout (row 1 = headers):
    Column A: Hostname (from CrowdStrike)
    Column B: Last Name (inventory)
    Column C: Name      (inventory)
    Column D: ID#       (inventory, 5-character HSI)

A hostname matches when its last 5 characters equal an ID# in column D
(case-insensitive, surrounding spaces ignored).

Usage:
    python match_hostnames.py data.xlsx                 # adds/replaces a "Matched" tab in data.xlsx
    python match_hostnames.py data.xlsx -o result.xlsx  # writes to a new file instead
    python match_hostnames.py data.xlsx --sheet Sheet1  # pick the source tab (default: first tab)
    python match_hostnames.py data.csv                  # writes data_matched.csv

CSV output has a Status column on every row:
    Matched                      - hostname whose last 5 chars match an ID#
    Unmatched - hostname only    - hostname with no matching ID#
    Unmatched - inventory only   - inventory item no hostname matched

Requires: openpyxl  (pip install openpyxl)
