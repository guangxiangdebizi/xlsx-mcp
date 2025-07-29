# 📊 XLSX Reader MCP

> A powerful Model Context Protocol (MCP) server for reading and analyzing Excel documents with advanced features.

<a href="https://glama.ai/mcp/servers/@guangxiangdebizi/xlsx-mcp">
  <img width="380" height="200" src="https://glama.ai/mcp/servers/@guangxiangdebizi/xlsx-mcp/badge" alt="XLSX Reader MCP server" />
</a>

[![License](https://img.shields.io/badge/License-Apache%202.0-blue.svg)](https://opensource.org/licenses/Apache-2.0)
[![TypeScript](https://img.shields.io/badge/TypeScript-5.3+-blue.svg)](https://www.typescriptlang.org/)
[![MCP SDK](https://img.shields.io/badge/MCP%20SDK-0.6.0-green.svg)](https://github.com/modelcontextprotocol/sdk)

## ✨ Features

- 📖 **Read Excel Files**: Support for `.xlsx`, `.xls`, and `.xlsm` formats
- 🎯 **Flexible Data Access**: Read entire sheets, specific ranges, or targeted cells
- 📊 **Multiple Output Formats**: JSON, CSV, and formatted table outputs
- 🔍 **Advanced Analysis**: Detailed file structure, data type analysis, and statistics
- 📋 **Sheet Management**: Work with multiple sheets and get comprehensive overviews
- 🎨 **Beautiful Formatting**: Clean, readable output with proper table formatting
- ⚡ **High Performance**: Efficient processing of large Excel files
- 🚫 **Smart Filtering**: Automatically skip empty rows for cleaner output
- 📏 **Configurable Limits**: Control maximum rows and preview size for optimal performance

## 🚀 Quick Start

### Installation

```bash
# Clone the repository
git clone https://github.com/guangxiangdebizi/xlsx-reader-mcp.git
cd xlsx-reader-mcp

# Install dependencies
npm install

# Build the project
npm run build
```

### Usage with Claude Desktop

#### Method 1: Stdio Mode (Recommended for Development)

Add to your Claude Desktop configuration:

```json
{
  "mcpServers": {
    "xlsx-reader": {
      "command": "node",
      "args": ["path/to/xlsx-reader-mcp/build/index.js"]
    }
  }
}
```

#### Method 2: SSE Mode (Recommended for Production)

```bash
# Start the SSE server
npm run sse
```

Then add to Claude Desktop configuration:

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

## 🛠️ Available Tools

### 1. `read_xlsx` - Excel File Reader

Read and extract data from Excel files with flexible formatting options and search capabilities.

**Parameters:**
- `filePath` (required): Path to the Excel file
- `sheetName` (optional): Specific sheet to read
- `range` (optional): Cell range (e.g., "A1:C10")
- `format` (optional): Output format - "json", "csv", or "table" (default)
- `includeHeaders` (optional): Include headers in output (default: true)
- `maxRows` (optional): Maximum rows to return (default: 100, max: 1000)
- `searchColumn` (optional): Column to search in (column name, index, or Excel letter)
- `searchValue` (optional): Value to search for
- `searchType` (optional): Search type - "exact", "contains", "startsWith", "endsWith" (default: "exact")

**Note:** Empty rows are automatically filtered out to provide cleaner output.

**Example Usage:**
```
Read the Excel file "data.xlsx" and show the first sheet as a table
```

### 2. `analyze_xlsx` - Excel File Analyzer

Perform comprehensive analysis of Excel files including structure, data types, and statistics.

**Parameters:**
- `filePath` (required): Path to the Excel file
- `includePreview` (optional): Include data preview (default: true)
- `previewRows` (optional): Number of preview rows (default: 5, max: 20)
- `analyzeDataTypes` (optional): Analyze column data types (default: true)

**Example Usage:**
```
Analyze the structure and content of "report.xlsx"
```

## 📖 Examples

### Basic File Reading

```
Read the Excel file "sales_data.xlsx"
```

### Reading Specific Sheet and Range

```
Read cells A1:E10 from the "Summary" sheet in "quarterly_report.xlsx" and format as JSON
```

### Reading with Row Limits

```
Read "large_dataset.xlsx" but limit to first 50 rows
```

### Searching Data

```
Search for "ADSL" in column A of "data.xlsx" with exact match
```

```
Find all rows containing "John" in the "Name" column of "employees.xlsx"
```

```
Search for values starting with "ABC" in column 3 of "products.xlsx"
```

### Comprehensive File Analysis

```
Analyze "customer_database.xlsx" and show detailed information about all sheets
```

## 🏗️ Project Structure

```
src/
├── index.ts              # MCP server entry point
└── tools/
    ├── xlsx-reader.ts    # Excel file reading tool
    └── xlsx-analyzer.ts  # Excel file analysis tool
```

## 🔧 Development

### Scripts

- `npm run build` - Build the TypeScript project
- `npm run dev` - Watch mode for development
- `npm start` - Start the MCP server
- `npm run sse` - Start SSE server on port 3100

### Dependencies

- **[@modelcontextprotocol/sdk](https://www.npmjs.com/package/@modelcontextprotocol/sdk)**: MCP SDK for server implementation
- **[xlsx](https://www.npmjs.com/package/xlsx)**: Excel file processing library

## 📄 License

This project is licensed under the Apache License 2.0 - see the [LICENSE](LICENSE) file for details.

## 👨‍💻 Author

**Xingyu Chen**
- 📧 Email: [guangxiangdebizi@gmail.com](mailto:guangxiangdebizi@gmail.com)
- 🐙 GitHub: [@guangxiangdebizi](https://github.com/guangxiangdebizi/)
- 💼 LinkedIn: [Xingyu Chen](https://www.linkedin.com/in/xingyu-chen-b5b3b0313/)
- 📦 NPM: [@xingyuchen](https://www.npmjs.com/~xingyuchen)

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## 🙏 Acknowledgments

- Thanks to the [Model Context Protocol](https://github.com/modelcontextprotocol) team for the excellent SDK
- Built with [SheetJS](https://sheetjs.com/) for robust Excel file processing

---

*Made with ❤️ for the MCP community*