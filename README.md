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

### 1.1. Input JSON File Format

The program uses `xlsx_data.json` file with the following structure:

```json
{
  "sheets": [
    {
      "name": "Sheet name",
      "columnWidths": [numbers],
      "rowHeights": [numbers],
      "data": [
        ["Cell1", "Cell2", ...],
        ["Cell1", 123, ...],
        ...
      ]
    }
  ]
}
```

### 1.2. Configuration Elements

#### Sheet Parameters:
- `name` (optional) - worksheet name
- `columnWidths` (optional) - array of column widths (in characters)
- `rowHeights` (optional) - array of row heights (in points)
- `data` (required) - 2D array of worksheet data

### 1.3. Cell Data Types

The program automatically detects data types:
- **Strings**: `"text"`
- **Numbers**: `123`, `45.67`
- **Boolean values**: `true`, `false`
- **Formulas**: `"=SUM(A1:A10)"` (starts with =)
- **Dates**: strings in `"YYYY-MM-DD"` format

### 1.4. Automatic Formatting

The program applies styles automatically:
1. **Header row**:
   - Gray background
   - Bold text
   - Center alignment
   - Text wrapping

2. **First column**:
   - Light gray background
   - Bold text
   - Center alignment

3. **Numeric cells**:
   - Even columns: pale blue
   - Odd columns: pale cyan

4. **Text cells**:
   - Even columns: pale green
   - Odd columns: pale orange

5. **Date cells**:
   - Pale purple background
   - "YYYY-MM-DD" format

### 1.5. Examples

#### Simple table:
```json
{
  "sheets": [
    {
      "name": "Report",
      "data": [
        ["Product", "Quantity", "Price"],
        ["Apples", 150, 25.50],
        ["Pears", 80, 32.75],
        ["Total", "=SUM(B2:B3)", "=SUM(C2:C3)"]
      ]
    }
  ]
}
```

#### Table with dates:
```json
{
  "sheets": [
    {
      "name": "Schedule",
      "columnWidths": [20, 15, 15],
      "data": [
        ["Event", "Date", "Days"],
        ["Project start", "2023-01-10", 0],
        ["First milestone", "2023-02-15", 36],
        ["Completion", "2023-05-20", 130]
      ]
    }
  ]
}
```

## 2. For Developers: Program Architecture

### 2.1. Overall Structure

The program consists of one main class `XlsxGenerator` with methods:
- `main()` - entry point
- `generate()` - main generation method
- Helper methods for style creation, sheet processing and data handling

### 2.2. Key Components

#### 2.2.1. Input Processing
- `readJsonData()` - JSON reading and parsing
- Uses `org.json` library

#### 2.2.2. Excel Generation
- Based on Apache POI (`XSSFWorkbook`)
- Supports:
  - Various data types
  - Formulas
  - Styles and formatting
  - Auto-sizing

#### 2.2.3. Style System
- Styles created during initialization
- Cached in Map for reuse
- Automatically applied based on:
  - Cell position
  - Data type
  - Column parity

### 2.3. Error Handling

- Basic exception handling
- Debug mode (`--debug`) for verbose output

### 2.4. Extensibility

#### 2.4.1. Adding New Styles
1. Add color to `COLORS` array
2. Create new style in `createDefaultStyles()`
3. Add application logic in `applyCellStyle()`

#### 2.4.2. Supporting New Data Types
Modify `applyCellStyle()` to recognize new types

### 2.5. Key Dependencies

- **Apache POI** (v5.2.3+) - Excel manipulation
- **org.json** (v20231013+) - JSON processing

### 2.6. Maintenance Recommendations

1. **Testing**:
   - Verify all data type handling
   - Test edge cases (empty data, large files)

2. **Logging**:
   - Add more detailed logging
   - Consider integrating SLF4J

3. **Optimization**:
   - For large files consider streaming
   - Optimize style creation

4. **Security**:
   - Input JSON validation
   - File size limitations

### 2.7. Extension Example

To add hyperlink support:

```java
// In createSheet() method:
if (cellData instanceof String && ((String)cellData).startsWith("http")) {
    CreationHelper helper = workbook.getCreationHelper();
    Hyperlink link = helper.createHyperlink(HyperlinkType.URL);
    link.setAddress((String)cellData);
    cell.setHyperlink(link);
}
```
