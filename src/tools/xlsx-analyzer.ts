import XLSX from 'xlsx';
import * as fs from 'fs';
import * as path from 'path';

export const xlsxAnalyzer = {
  name: "analyze_xlsx",
  description: "Analyze Excel files to get detailed information about structure, data types, statistics, and sheet contents.",
  parameters: {
    type: "object",
    properties: {
      filePath: {
        type: "string",
        description: "Path to the Excel file to analyze"
      },
      includePreview: {
        type: "boolean",
        description: "Whether to include a preview of the data (default: true)"
      },
      previewRows: {
        type: "number",
        description: "Number of rows to preview (default: 5, max: 20)"
      },
      analyzeDataTypes: {
        type: "boolean",
        description: "Whether to analyze data types in each column (default: true)"
      }
    },
    required: ["filePath"]
  },
  
  async run(args: { 
    filePath: string; 
    includePreview?: boolean;
    previewRows?: number;
    analyzeDataTypes?: boolean;
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

      // Get file stats
      const stats = fs.statSync(args.filePath);
      const fileSize = (stats.size / 1024).toFixed(2); // KB

      // Read the workbook
      const workbook = XLSX.readFile(args.filePath);
      const sheetNames = workbook.SheetNames;
      
      const includePreview = args.includePreview !== false;
      const previewRows = Math.min(args.previewRows || 5, 20); // Default 5, max 20
      const analyzeDataTypes = args.analyzeDataTypes !== false;

      // Analyze each sheet
      const sheetAnalysis: any[] = [];
      
      for (const sheetName of sheetNames) {
        const worksheet = workbook.Sheets[sheetName];
        const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');
        
        // Get sheet dimensions
        const totalRows = range.e.r + 1;
        const totalCols = range.e.c + 1;
        
        // Convert to array for analysis
        let rawData = XLSX.utils.sheet_to_json(worksheet, { 
          header: 1,
          defval: ''
        }) as any[][];
        
        // Helper function to check if a row is empty
        const isEmptyRow = (row: any[]): boolean => {
          return row.every(cell => cell === '' || cell == null || cell === undefined);
        };
        
        // Get original statistics
        const originalTotalRows = rawData.length;
        
        // Filter out completely empty rows for analysis
        const data = rawData.filter((row, index) => {
          // Always keep the first row (header)
          if (index === 0) return true;
          // Filter out empty rows
          return !isEmptyRow(row);
        });
        
        const dataRows = data.length > 0 ? data.length - 1 : 0; // Exclude header
        const skippedEmptyRows = originalTotalRows - data.length;
        
        // Analyze data types if requested
        let columnAnalysis: any[] = [];
        if (analyzeDataTypes && data.length > 0) {
          const headers = data[0] || [];
          const dataRows = data.slice(1);
          
          columnAnalysis = headers.map((header: any, colIndex: number) => {
            const columnData = dataRows.map(row => row[colIndex]).filter(cell => cell !== '' && cell != null);
            
            if (columnData.length === 0) {
              return {
                column: colIndex + 1,
                header: header || `Column ${colIndex + 1}`,
                dataType: 'empty',
                nonEmptyCount: 0,
                emptyCount: dataRows.length,
                uniqueValues: 0
              };
            }
            
            // Analyze data types
            const types = {
              number: 0,
              string: 0,
              date: 0,
              boolean: 0
            };
            
            columnData.forEach(cell => {
              if (typeof cell === 'number') {
                types.number++;
              } else if (typeof cell === 'boolean') {
                types.boolean++;
              } else if (cell instanceof Date || (typeof cell === 'string' && !isNaN(Date.parse(cell)))) {
                types.date++;
              } else {
                types.string++;
              }
            });
            
            // Determine primary type
            const primaryType = Object.entries(types).reduce((a, b) => types[a[0] as keyof typeof types] > types[b[0] as keyof typeof types] ? a : b)[0];
            
            // Get unique values count
            const uniqueValues = new Set(columnData).size;
            
            return {
              column: colIndex + 1,
              header: header || `Column ${colIndex + 1}`,
              dataType: primaryType,
              nonEmptyCount: columnData.length,
              emptyCount: dataRows.length - columnData.length,
              uniqueValues,
              typeDistribution: types
            };
          });
        }
        
        // Get preview data
        let preview = '';
        if (includePreview && data.length > 0) {
          // Limit preview data and ensure we don't exceed available data
          const maxPreviewRows = Math.min(previewRows + 1, data.length, 21); // +1 for header, max 21 total
          const previewData = data.slice(0, maxPreviewRows);
          const maxCols = Math.max(...previewData.map(row => row.length));
          const colWidths = Array(maxCols).fill(0);
          
          // Calculate column widths
          previewData.forEach(row => {
            row.forEach((cell: any, index: number) => {
              const cellStr = String(cell || '');
              colWidths[index] = Math.max(colWidths[index], Math.min(cellStr.length, 20)); // Max width 20
            });
          });
          
          // Format preview table
          const formatRow = (row: any[]) => {
            return '| ' + row.map((cell: any, index: number) => {
              const cellStr = String(cell || '');
              const truncated = cellStr.length > 20 ? cellStr.substring(0, 17) + '...' : cellStr;
              return truncated.padEnd(colWidths[index]);
            }).join(' | ') + ' |';
          };
          
          const separator = '|' + colWidths.map(width => '-'.repeat(width + 2)).join('|') + '|';
          
          preview = previewData.map((row: any[], index: number) => {
            const formattedRow = formatRow(row);
            if (index === 0) {
              return formattedRow + '\n' + separator;
            }
            return formattedRow;
          }).join('\n');
        }
        
        sheetAnalysis.push({
          name: sheetName,
          dimensions: {
            rows: totalRows,
            columns: totalCols
          },
          originalTotalRows,
          dataRows,
          skippedEmptyRows,
          hasData: data.length > 0,
          columnAnalysis,
          preview
        });
      }
      
      // Generate comprehensive report
      const report = [
        `# 📊 Excel File Analysis Report`,
        ``,
        `## 📁 File Information`,
        `- **File:** ${path.basename(args.filePath)}`,
        `- **Full Path:** ${args.filePath}`,
        `- **Size:** ${fileSize} KB`,
        `- **Format:** ${ext.toUpperCase()}`,
        `- **Last Modified:** ${stats.mtime.toLocaleString()}`,
        ``,
        `## 📋 Workbook Overview`,
        `- **Total Sheets:** ${sheetNames.length}`,
        `- **Sheet Names:** ${sheetNames.join(', ')}`,
        ``
      ];
      
      // Add detailed sheet analysis
      sheetAnalysis.forEach((sheet, index) => {
        report.push(`## 📄 Sheet ${index + 1}: "${sheet.name}"`);
        report.push(``);
        report.push(`### Basic Information`);
        report.push(`- **Total Rows:** ${sheet.originalTotalRows} (including header)`);
        report.push(`- **Data Rows:** ${sheet.dataRows} (excluding header)`);
        if (sheet.skippedEmptyRows > 0) {
          report.push(`- **Empty Rows Filtered:** ${sheet.skippedEmptyRows}`);
        }
        report.push(`- **Columns:** ${sheet.dimensions.columns}`);
        report.push(`- **Dimensions:** ${sheet.dimensions.rows} rows × ${sheet.dimensions.columns} columns (after filtering)`);
        report.push(`- **Has Data:** ${sheet.hasData ? 'Yes' : 'No'}`);
        report.push(``);
        
        if (analyzeDataTypes && sheet.columnAnalysis.length > 0) {
          report.push(`### Column Analysis`);
          report.push(``);
          sheet.columnAnalysis.forEach((col: any) => {
            report.push(`**${col.header}** (Column ${col.column})`);
            report.push(`- Primary Type: ${col.dataType}`);
            report.push(`- Non-empty: ${col.nonEmptyCount}, Empty: ${col.emptyCount}`);
            report.push(`- Unique Values: ${col.uniqueValues}`);
            if (col.typeDistribution) {
              const dist = Object.entries(col.typeDistribution)
                .filter(([_, count]) => (count as number) > 0)
                .map(([type, count]) => `${type}: ${count}`)
                .join(', ');
              if (dist) report.push(`- Type Distribution: ${dist}`);
            }
            report.push(``);
          });
        }
        
        if (includePreview && sheet.preview) {
          report.push(`### Data Preview (First ${previewRows} rows)`);
          report.push(``);
          report.push('```');
          report.push(sheet.preview);
          report.push('```');
          report.push(``);
        }
        
        if (index < sheetAnalysis.length - 1) {
          report.push(`---`);
          report.push(``);
        }
      });
      
      return {
        content: [{
          type: "text",
          text: report.join('\n')
        }]
      };
      
    } catch (error: any) {
      return {
        content: [{
          type: "text",
          text: `❌ **Error analyzing Excel file:** ${error.message}`
        }],
        isError: true
      };
    }
  }
};