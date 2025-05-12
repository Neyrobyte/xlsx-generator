<pre>
    _   __                      __          __     
   / | / /__  __  ___________  / /_  __  __/ /____ 
  /  |/ / _ \/ / / / ___/ __ \/ __ \/ / / / __/ _ \
 / /|  /  __/ /_/ / /  / /_/ / /_/ / /_/ / /_/  __/
/_/ |_/\___/\__, /_/   \____/_.___/\__, /\__/\___/ 
           /____/                 /____/           
</pre>



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
Header1;Header2;Header3
value1;value2;value3
42;2023-01-01;Text

sheet:ListName2
OtherHeader1;OtherHeader2
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
