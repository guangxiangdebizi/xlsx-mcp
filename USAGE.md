# 📖 Usage Guide

This guide demonstrates how to use the XLSX Reader MCP tools with practical examples.

## 🚀 Getting Started

### 1. Start the MCP Server

```bash
# Build the project
npm run build

# Start the server (stdio mode)
node build/index.js

# Or start SSE server
npm run sse
```

### 2. Configure Claude Desktop

Add to your Claude Desktop configuration file:

**For Stdio Mode:**
```json
{
  "mcpServers": {
    "xlsx-reader": {
      "command": "node",
      "args": ["C:/path/to/xlsx-reader-mcp/build/index.js"]
    }
  }
}
```

**For SSE Mode:**
```json
{
  "mcpServers": {
    "xlsx-reader": {
      "type": "sse",
      "url": "http://localhost:3100/sse",
      "timeout": 600
    }
  }
}
```

## 🛠️ Tool Examples

### Reading Excel Files

#### Basic File Reading
```
Read the Excel file "sales_data.xlsx"
```

#### Reading Specific Sheet
```
Read the "Summary" sheet from "quarterly_report.xlsx"
```

#### Reading Cell Range
```
Read cells A1:E10 from "data.xlsx" and format as JSON
```

#### Different Output Formats
```
Read "inventory.xlsx" and format as CSV
```

#### Reading with Row Limits
```
Read "large_dataset.xlsx" but only show first 50 rows
```

### Analyzing Excel Files

#### Basic Analysis
```
Analyze the structure of "customer_data.xlsx"
```

#### Detailed Analysis with Preview
```
Analyze "financial_report.xlsx" and show the first 10 rows of each sheet
```

#### Analysis without Data Types
```
Analyze "large_dataset.xlsx" but skip data type analysis for faster processing
```

## 📊 Sample Outputs

### Reading Excel File (Table Format)

```
📊 **Excel File Analysis**

**File:** ADSL_test.xlsx
**Sheet:** ADSL
**Available Sheets:** ADSL
**Data Rows:** 2
**Data Columns:** 11
**Format:** table
**Range:** Entire sheet

## Data Content

| Dataset | Label                           | Description                    | Structure              | Class    | Keys             | Purpose                                                                                                                                                                                                                        | Comment                                                                                                    | Pages    | Study | Document |
|---------|---------------------------------|--------------------------------|------------------------|----------|------------------|--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|------------------------------------------------------------------------------------------------------------|----------|-------|-----------|
| ADSL    | Subject-Level Analysis Dataset  | SUBJECT LEVEL ANALYSIS DATASET | One record per subject | Analysis | STUDYID, USUBJID | This dataset is based on SDTM domain DM plus additional derived subject-level variables needed for analyses; it includes all enrolled subjects. Screen Failures are excluded since they are not needed for this study analysis | If DC domain is used, SUBJID should be used to merge SDTM datasets instead of USUBJID for code generation. |          |       | Y         |
```

### Excel File Analysis

```
# 📊 Excel File Analysis Report

## 📁 File Information
- **File:** ADS_test.xlsx
- **Full Path:** C:/Users/26214/Desktop/MyProject/sascode-agent/ADS_test.xlsx
- **Size:** 12.34 KB
- **Format:** .XLSX
- **Last Modified:** 12/19/2024, 10:30:00 AM

## 📋 Workbook Overview
- **Total Sheets:** 3
- **Sheet Names:** ADSL, Documents, Sheet1

## 📄 Sheet 1: "ADSL"

### Basic Information
- **Dimensions:** 2 rows × 11 columns
- **Data Rows:** 2
- **Has Data:** Yes

### Column Analysis

**Dataset** (Column 1)
- Primary Type: string
- Non-empty: 1, Empty: 0
- Unique Values: 1
- Type Distribution: string: 1

**Label** (Column 2)
- Primary Type: string
- Non-empty: 1, Empty: 0
- Unique Values: 1
- Type Distribution: string: 1

### Data Preview (First 5 rows)

```
| Dataset | Label                          | Description                    | Structure              | Class    |
|---------|--------------------------------|--------------------------------|------------------------|-----------|
| ADSL    | Subject-Level Analysis Dataset | SUBJECT LEVEL ANALYSIS DATASET | One record per subject | Analysis |
```
```

## 🎯 Advanced Usage Tips

### 1. Working with Large Files
- Use the `analyze_xlsx` tool first to understand file structure
- Read specific sheets or ranges to avoid memory issues
- Set `analyzeDataTypes: false` for faster analysis of large files
- Use `maxRows` parameter to limit output size (default: 100, max: 1000)
- Enable `skipEmptyRows: false` only when you need to preserve exact row structure

### 2. Data Processing Workflow
1. **Analyze** the file structure first
2. **Identify** the sheets and ranges you need
3. **Read** specific data with appropriate format
4. **Process** the data as needed

### 3. Format Selection
- **Table**: Best for human-readable output and reports
- **JSON**: Ideal for programmatic processing
- **CSV**: Perfect for data exchange and import/export

### 4. Performance Optimization
- **Default Limits**: Tools automatically limit to 100 rows by default
- **Preview Control**: Analysis preview is limited to 5 rows by default (max: 20)
- **Custom Limits**: Adjust `maxRows` and `previewRows` based on your needs

### 5. Search Functionality
- **Column Search**: Use `searchColumn` to specify which column to search (name, index, or Excel letter)
- **Search Value**: Set `searchValue` to find specific data
- **Search Types**: Choose from "exact", "contains", "startsWith", "endsWith" matching
- **Combined with Limits**: Search results respect `maxRows` parameter

## Example Usage Scenarios

### 1. Quick Data Preview
```
Analyze the structure of "sales_data.xlsx"
Read "sales_data.xlsx" but only show first 10 rows
```

### 2. Working with Large Files
```
Analyze "large_dataset.xlsx" and show the first 3 rows of each sheet
Read "large_dataset.xlsx" but only show first 50 rows
```

### 3. Specific Data Extraction
```
Read the "Q4 Results" sheet from "report.xlsx"
Read cells A1:E20 from "report.xlsx"
```

### 4. Search and Filter Data
```
Read "data.xlsx" and search for exact match "ADSL" in column A
Read "employees.xlsx" and search for "Sales" in Department column
Read "products.xlsx" and find SKUs starting with "ABC"
Read "data.xlsx" and search for "Active" in column 3, limit to 20 rows
```

### 6. Error Handling
The tools provide detailed error messages for:
- File not found
- Invalid file formats
- Missing sheets
- Invalid cell ranges

## 🔧 Troubleshooting

### Common Issues

1. **"File not found" error**
   - Check the file path is correct
   - Use absolute paths when possible
   - Ensure the file exists and is accessible

2. **"Sheet not found" error**
   - Use `analyze_xlsx` to see available sheet names
   - Check for exact spelling and case sensitivity

3. **"Invalid range" error**
   - Use Excel-style ranges (e.g., "A1:C10")
   - Ensure the range exists in the sheet

4. **Memory issues with large files**
   - Read specific ranges instead of entire sheets
   - Use smaller preview row counts
   - Disable data type analysis for very large files
   - Reduce `maxRows` parameter for very large datasets

## 📝 Best Practices

1. **Always analyze first**: Use `analyze_xlsx` to understand file structure
2. **Use specific ranges**: Read only the data you need
3. **Choose appropriate formats**: Select the output format that matches your use case
4. **Handle errors gracefully**: Check for error responses in your workflow
5. **Test with sample data**: Verify your approach with smaller files first
5. **Optimize for performance**: Use appropriate `maxRows` limits for your use case
6. **Control preview size**: Adjust `previewRows` based on how much data you need to see

---

*For more examples and advanced usage, check the [README.md](README.md) file.*