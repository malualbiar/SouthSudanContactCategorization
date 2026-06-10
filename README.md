# Data Parsing and Multi-Sheet Excel Export Challenge

## Problem Description

You are required to write a Python utility script to read, parse, and categorize an unstructured data sheet of South Sudan mobile phone numbers (`south_sudan_phone_numbers.xls`). The input source contains completely shuffled rows of varying network groups. 

Your objective is to map these numbers into specialized network bins within a python dictionary framework, and sequentially commit those classifications to isolated worksheets within a final target workbook (`.xlsx`).

---

## Technical Requirements

### Libraries
Your solution must exclusively utilize the following Python modules:
* `pandas` (for file processing, column ingestion, and string slicing)
* `openpyxl` (as the engine driving the complex, multi-sheet workbook output export)

### Categorization Matrix

Your logic must map data arrays to distinct target dictionary keys based on matching digit patterns:

| Target Dictionary Key | Matching Pattern Condition | Expected Array Contents |
| :--- | :--- | :--- |
| `"MTN Corporate 1"` | Numbers starting with `211922903` | `[211922903xxx, ...]` |
| `"MTN Corporate 2"` | Numbers starting with `211922904` | `[211922904xxx, ...]` |
| `"MTN Corporate 3"` | Numbers starting with `211922905` | `[211922905xxx, ...]` |
| `"MTN Numbers"` | Numbers starting with `21192` (excluding Corporate 1, 2, 3) | `[21192xxxxxxx, ...]` |
| `"Zain Numbers"` | Numbers starting with `21191` | `[21191xxxxxxx, ...]` |
| `"Digital Numbers"` | Numbers starting with `21198` | `[21198xxxxxxx, ...]` |

---

## Input / Output Protocol

### Input Format
The program accepts an unstructured raw file `south_sudan_phone_numbers.xls` containing columns:
* `Prefix Group`: Mask type rule (e.g. `211922903xxx`).
* `Phone Number`: Mixed numbers stored as literal string fields.

### Intermediate Dictionary Object
Your processing logic must populate a structured Python dictionary schema exactly like this:

```python
output = {
    "MTN Corporate 1": [211922903001, 211922903458, ...],
    "MTN Corporate 2": [211922904107, 211922904293, ...],
    "MTN Corporate 3": [211922905477, 211922905892, ...],
    "MTN Numbers":     [211924619379, 211922097463, ...],
    "Zain Numbers":    [211910769941, 211910609383, ...],
    "Digital Numbers": [211980893074, 211987515328, ...]
}