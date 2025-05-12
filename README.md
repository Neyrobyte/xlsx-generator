**English** | [Русский](README.ru.md)
<p align="center">
  <img src="logo.svg" alt="" width="256">
</p>

---

# XLSX Generator  
A tool to generate formatted Excel files from simple text data. Designed as an AI helper utility.

## For Users

### Script Purpose  
This script converts text files into formatted Excel (XLSX) files with automatic styling:
- Automatic data type detection (numbers, dates, text)
- Color-coded columns by data type
- Header and border formatting
- Automatic column width adjustment

### Requirements  
- Python 3.6 or newer  
- Required dependency: `openpyxl`  

### Installation  
```bash
pip install openpyxl
```

### Usage  
1. Create a text file named `xlsx_generate.txt` in the same folder as the script  
2. Fill it with data using the format described below  
3. Run the script: `python3 xlsx.py`  
4. The output will be saved to `storage/shared/download/output.xlsx`  

### Input File Format  
The file should contain data in the following format:
```
sheet:ListName1
Tutle1;Title2;Title3
value1;value2;value3
42;2023-01-01;Text

sheet:ListName2
OtherTitle1;OtherTitle2
100;Example Text
```

#### Format Specifications:
- Sheets are separated by empty lines  
- Sheet names are specified with `sheet:SheetName`  
- Data fields are delimited by semicolons (`;`)  
- Special cell formats supported:  
  - Dates: `YYYY-MM-DD`, `DD/MM/YYYY`, `DD.MM.YYYY`  
  - Excel formulas: `{"value": "=SUM(A1:A10)"}`  
  - JSON format for complex values: `{"value": "text"}`  

### Output Features  
The script generates an Excel file with:
- Gray headers with bold font  
- Light gray first column  
- Numeric columns with alternating pale blue shades  
- Text columns with alternating pale green and orange shades  
- Date columns with pale purple background  
- Thick external borders  
- Auto-adjusted column widths  

## For Developers  

### Script Architecture  
Main functions:  
1. `parse_input_file` - parses the input text file  
2. `apply_default_styles` - applies styles to Excel sheets  
3. `generate_xlsx` - main file generation function  

### Style Configuration  
Styles are defined in the `DEFAULT_STYLES` dictionary. Customizable options:  
- Fill colors  
- Font settings  
- Borders  
- Text alignment  

### Extending Functionality  
1. **Adding new data types**:  
   - Implement new type-checking functions (similar to `is_number` and `is_date`)  
   - Add corresponding styles to `DEFAULT_STYLES`  
   - Modify logic in `apply_default_styles`  

2. **Changing input format**:  
   - Modify the `parse_input_file` function  
   - Update user documentation  

3. **Adding new styles**:  
   - Use methods from `openpyxl.styles`  
   - Add new parameters to `DEFAULT_STYLES`  

### Testing  
To test changes:  
1. Create a test input file  
2. Run the script  
3. Verify the output file  

### Error Handling  
The script handles basic errors:  
- Missing input file  
- JSON parsing errors  
- Date formatting errors  

To add new error handling, modify the `try-except` block in `generate_xlsx`.  

### Dependencies  
- `openpyxl` - for Excel file operations  
- Python standard libraries: `datetime`, `re`, `json`, `os`  

### Compatibility  
Tested on:  
- Python 3.6+  
- OpenPyXL 3.0+

## **AI Documentation (for Neural Networks)**  

This section is intended for AI models that will analyze the repository and generate input data for the `xlsx-generator` script.  

#### **1. Input Data Structure**  
The script expects a JSON structure describing the Excel table. Format:  
```json
{
  "filename": "output_filename.xlsx",
  "sheets": [
    {
      "name": "SheetName",
      "headers": ["Header1", "Header2", ...],
      "rows": [
        ["Value1", "Value2", ...],
        ...
      ],
      "styles": {
        "header": { ... },
        "cells": { ... }
      }
    }
  ]
}
```  

#### **2. Field Requirements**  
- **`filename`** (string):  
  - Must end with `.xlsx`.  
  - If omitted, defaults to `output.xlsx`.  

- **`sheets`** (array of objects):  
  - Each object defines a separate sheet.  
  - **`name`** (string): Sheet name (if empty, defaults to `Sheet1`, `Sheet2`, etc.).  
  - **`headers`** (array of strings): Column headers (optional).  
  - **`rows`** (array of arrays): Table data. Each sub-array represents a row.  
  - **`styles`** (object, optional): Styling rules.  

#### **3. Styling (`styles`)**  
The script supports basic formatting via `openpyxl`. Example:  
```json
"styles": {
  "header": {
    "font": {"bold": true, "color": "FF0000"},
    "fill": {"patternType": "solid", "fgColor": "FFFF00"}
  },
  "cells": {
    "font": {"size": 12},
    "alignment": {"horizontal": "center"}
  }
}
```  
Allowed parameters:  
- **`font`**: `bold`, `italic`, `size`, `color` (HEX format).  
- **`fill`**: `patternType`, `fgColor` (background color).  
- **`alignment`**: `horizontal`, `vertical`.  

#### **4. Generation Rules**  
- If there are fewer rows than headers, empty cells are filled with `null`.  
- If a row contains more values than headers, excess data is ignored.  
- The script automatically adjusts column widths to fit content.  

#### **5. Example AI Task**  
**User Request:**  
> "Generate a table with two sheets: the first named 'Report' with headers ['ID', 'Name', 'Score'] and 3 data rows; the second named 'Stats' with a header ['Average Score'] and one row with the calculated value."  

**Generated Input (JSON):**  
```json
{
  "filename": "report.xlsx",
  "sheets": [
    {
      "name": "Report",
      "headers": ["ID", "Name", "Score"],
      "rows": [
        [1, "Alex", 85],
        [2, "Maria", 92],
        [3, "Ivan", 78]
      ]
    },
    {
      "name": "Stats",
      "headers": ["Average Score"],
      "rows": [[85]]
    }
  ]
}
```  

#### **6. Key Notes**  
- The AI **must** validate the JSON structure before generation.  
- If the user requests calculations (e.g., "average score"), the AI should compute them and insert results into `rows`.  
- For complex styling, `styles` can be omitted—the script will use default settings.
