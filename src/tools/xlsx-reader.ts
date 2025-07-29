import XLSX from 'xlsx';
import * as fs from 'fs';
import * as path from 'path';

export const xlsxReader = {
  name: "read_xlsx",
  description: "Read and extract data from Excel (xlsx) files. Supports reading specific sheets, cell ranges, and converting data to various formats.",
  parameters: {
    type: "object",
    properties: {
      filePath: {
        type: "string",
        description: "Path to the Excel file to read"
      },
      sheetName: {
        type: "string",
        description: "Name of the specific sheet to read (optional, defaults to first sheet)"
      },
      range: {
        type: "string",
        description: "Cell range to read (e.g., 'A1:C10', optional, defaults to entire sheet)"
      },
      format: {
        type: "string",
        enum: ["json", "csv", "table"],
        description: "Output format for the data (default: table)"
      },
      includeHeaders: {
        type: "boolean",
        description: "Whether to include headers in the output (default: true)"
      },
      maxRows: {
        type: "number",
        description: "Maximum number of rows to return (default: 100, max: 1000)"
      },
      searchColumn: {
        type: "string",
        description: "Column name or index to search in (e.g., 'A', '1', or 'Name')"
      },
      searchValue: {
        type: "string",
        description: "Value to search for in the specified column"
      },
      searchType: {
        type: "string",
        enum: ["exact", "contains", "startsWith", "endsWith"],
        description: "Type of search to perform (default: exact)"
      }
    },
    required: ["filePath"]
  },
  
  async run(args: { 
    filePath: string; 
    sheetName?: string; 
    range?: string; 
    format?: "json" | "csv" | "table";
    includeHeaders?: boolean;
    maxRows?: number;
    searchColumn?: string;
    searchValue?: string;
    searchType?: "exact" | "contains" | "startsWith" | "endsWith";
  }) {
    try {
      // Parameter validation
      if (!args.filePath) {
        throw new Error("File path is required");
      }

      // Check if file exists
      if (!fs.existsSync(args.filePath)) {
        throw new Error(`File not found: ${args.filePath}`);
      }

      // Check file extension
      const ext = path.extname(args.filePath).toLowerCase();
      if (!['.xlsx', '.xls', '.xlsm'].includes(ext)) {
        throw new Error("File must be an Excel file (.xlsx, .xls, or .xlsm)");
      }

      // Read the workbook
      const workbook = XLSX.readFile(args.filePath);
      
      // Get sheet names
      const sheetNames = workbook.SheetNames;
      if (sheetNames.length === 0) {
        throw new Error("No sheets found in the workbook");
      }

      // Determine which sheet to read
      let targetSheet: string;
      if (args.sheetName) {
        if (!sheetNames.includes(args.sheetName)) {
          throw new Error(`Sheet '${args.sheetName}' not found. Available sheets: ${sheetNames.join(', ')}`);
        }
        targetSheet = args.sheetName;
      } else {
        targetSheet = sheetNames[0];
      }

      // Get the worksheet
      const worksheet = workbook.Sheets[targetSheet];
      
      // Extract data based on format
      const format = args.format || "table";
      const includeHeaders = args.includeHeaders !== false;
      const maxRows = Math.min(args.maxRows || 100, 1000); // Default 100, max 1000
      const skipEmptyRows = true; // Always skip empty rows
      
      let data: any;
      let output: string;
      let originalTotalRows = 0;
      let nonEmptyDataRows = 0;
      let skippedEmptyRows = 0;
      let isDataTruncated = false;
      let filteredData: any[] = [];
      
      // Helper function to check if a row is empty
      const isEmptyRow = (row: any[]): boolean => {
        return row.every(cell => cell === '' || cell == null || cell === undefined);
      };
      
      // Helper function to perform search
      const performSearch = (data: any[][], searchColumn: string, searchValue: string, searchType: string) => {
        if (data.length === 0) return { filteredData: data, matchedRows: 0 };
        
        const headers = data[0];
        let columnIndex = -1;
        
        // Find column index
        if (/^\d+$/.test(searchColumn)) {
          // Numeric index (1-based)
          columnIndex = parseInt(searchColumn) - 1;
        } else if (/^[A-Z]+$/i.test(searchColumn)) {
          // Excel column letter (A, B, C, etc.)
          columnIndex = searchColumn.toUpperCase().charCodeAt(0) - 65;
        } else {
          // Column name
          columnIndex = headers.findIndex((header: any) => 
            String(header).toLowerCase() === searchColumn.toLowerCase()
          );
        }
        
        if (columnIndex === -1 || columnIndex >= headers.length) {
          throw new Error(`Column '${searchColumn}' not found. Available columns: ${headers.join(', ')}`);
        }
        
        // Filter rows based on search criteria
        const headerRow = [headers];
        const dataRows = data.slice(1);
        const matchedRows = dataRows.filter((row: any[]) => {
          const cellValue = String(row[columnIndex] || '').toLowerCase();
          const searchVal = searchValue.toLowerCase();
          
          switch (searchType) {
            case "contains":
              return cellValue.includes(searchVal);
            case "startsWith":
              return cellValue.startsWith(searchVal);
            case "endsWith":
              return cellValue.endsWith(searchVal);
            case "exact":
            default:
              return cellValue === searchVal;
          }
        });
        
        return {
          filteredData: [...headerRow, ...matchedRows],
          matchedRows: matchedRows.length
        };
      };

      if (args.range) {
        // Read specific range
        const range = XLSX.utils.decode_range(args.range);
        const rangeData: any[][] = [];
        
        for (let row = range.s.r; row <= range.e.r; row++) {
          const rowData: any[] = [];
          for (let col = range.s.c; col <= range.e.c; col++) {
            const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
            const cell = worksheet[cellAddress];
            rowData.push(cell ? cell.v : '');
          }
          rangeData.push(rowData);
        }
        data = rangeData;
      } else {
        // Read entire sheet
        let rawData = XLSX.utils.sheet_to_json(worksheet, { 
          header: 1,
          defval: ''
        }) as any[][];
        
        // Get sheet statistics before filtering
         originalTotalRows = rawData.length;
         
         // Always filter empty rows (except header)
         const headerRow = rawData.length > 0 ? [rawData[0]] : [];
         const dataRows = rawData.slice(1).filter(row => !isEmptyRow(row));
         filteredData = [...headerRow, ...dataRows];
         nonEmptyDataRows = dataRows.length;
         skippedEmptyRows = originalTotalRows - filteredData.length;
         
         // Apply search filter if specified
         if (args.searchColumn && args.searchValue) {
           const searchResults = performSearch(filteredData, args.searchColumn, args.searchValue, args.searchType || "exact");
           filteredData = searchResults.filteredData;
           nonEmptyDataRows = searchResults.matchedRows;
         }
         
         // Limit the number of rows for display
         const totalRowsToShow = includeHeaders ? maxRows + 1 : maxRows;
         data = filteredData.slice(0, totalRowsToShow);
         isDataTruncated = filteredData.length > totalRowsToShow;
      }

      // Format output
      switch (format) {
        case "json":
          if (includeHeaders && data.length > 0) {
            const headers = data[0];
            const rows = data.slice(1);
            const jsonData = rows.map((row: any[]) => {
              const obj: any = {};
              headers.forEach((header: string, index: number) => {
                obj[header || `Column${index + 1}`] = row[index] || '';
              });
              return obj;
            });
            output = JSON.stringify(jsonData, null, 2);
          } else {
            output = JSON.stringify(data, null, 2);
          }
          break;
          
        case "csv":
          output = data.map((row: any[]) => 
            row.map((cell: any) => 
              typeof cell === 'string' && cell.includes(',') ? `"${cell}"` : cell
            ).join(',')
          ).join('\n');
          break;
          
        default: // table
          if (data.length === 0) {
            output = "No data found in the specified range.";
          } else {
            const maxCols = Math.max(...data.map((row: any[]) => row.length));
            const colWidths = Array(maxCols).fill(0);
            
            // Calculate column widths
            data.forEach((row: any[]) => {
              row.forEach((cell: any, index: number) => {
                const cellStr = String(cell || '');
                colWidths[index] = Math.max(colWidths[index], cellStr.length);
              });
            });
            
            // Format table
            const formatRow = (row: any[]) => {
              return '| ' + row.map((cell: any, index: number) => {
                const cellStr = String(cell || '');
                return cellStr.padEnd(colWidths[index]);
              }).join(' | ') + ' |';
            };
            
            const separator = '|' + colWidths.map(width => '-'.repeat(width + 2)).join('|') + '|';
            
            output = data.map((row: any[], index: number) => {
              const formattedRow = formatRow(row);
              if (index === 0 && includeHeaders) {
                return formattedRow + '\n' + separator;
              }
              return formattedRow;
            }).join('\n');
          }
          break;
      }

      // Prepare summary with sheet statistics
      const summaryParts = [
        `📊 **Excel File Analysis**`,
        ``,
        `**File:** ${path.basename(args.filePath)}`,
        `**Sheet:** ${targetSheet}`,
        `**Available Sheets:** ${sheetNames.join(', ')}`,
        `**Format:** ${format}`,
        args.range ? `**Range:** ${args.range}` : `**Range:** Entire sheet`,
        ``,
        `## 📊 Sheet Information`
      ];
      
      if (!args.range) {
          summaryParts.push(
            `- **Total Rows:** ${originalTotalRows} (including header)`,
            `- **Data Rows:** ${nonEmptyDataRows} (excluding header)`
          );
          if (skippedEmptyRows > 0) {
            summaryParts.push(`- **Empty Rows Filtered:** ${skippedEmptyRows}`);
          }
          if (args.searchColumn && args.searchValue) {
            summaryParts.push(
              `- **🔍 Search Applied:** Column '${args.searchColumn}' ${args.searchType || 'exact'} '${args.searchValue}'`,
              `- **Search Results:** ${nonEmptyDataRows} matching rows found`
            );
          }
          if (isDataTruncated) {
            summaryParts.push(`- **⚠️ Display Limit:** Showing first ${maxRows} rows (${filteredData.length - (includeHeaders ? maxRows + 1 : maxRows)} more rows available)`);
          }
      } else {
        summaryParts.push(
          `- **Displayed Rows:** ${data.length}`,
          `- **Displayed Columns:** ${data.length > 0 ? Math.max(...data.map((row: any[]) => row.length)) : 0}`
        );
      }
      
      summaryParts.push(
        ``,
        `## Data Content`,
        ``,
        output
      );
      
      const summary = summaryParts.join('\n');
      
      return {
        content: [{
          type: "text",
          text: summary
        }]
      };
      
    } catch (error: any) {
      return {
        content: [{
          type: "text",
          text: `❌ **Error reading Excel file:** ${error.message}`
        }],
        isError: true
      };
    }
  }
};