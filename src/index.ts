import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { CallToolRequestSchema, ListToolsRequestSchema } from "@modelcontextprotocol/sdk/types.js";

// Import business tools
import { xlsxReader } from "./tools/xlsx-reader.js";
import { xlsxAnalyzer } from "./tools/xlsx-analyzer.js";

// Create MCP server
const server = new Server({
  name: "xlsx-reader-mcp",
  version: "1.0.0",
}, {
  capabilities: { tools: {} }
});

// Tool registration
server.setRequestHandler(ListToolsRequestSchema, async () => {
  return {
    tools: [
      {
        name: xlsxReader.name,
        description: xlsxReader.description,
        inputSchema: xlsxReader.parameters
      },
      {
        name: xlsxAnalyzer.name,
        description: xlsxAnalyzer.description,
        inputSchema: xlsxAnalyzer.parameters
      }
    ]
  };
});

// Tool call handling
server.setRequestHandler(CallToolRequestSchema, async (request) => {
  switch (request.params.name) {
    case "read_xlsx":
      return await xlsxReader.run(request.params.arguments as any);
    case "analyze_xlsx":
      return await xlsxAnalyzer.run(request.params.arguments as any);
    default:
      throw new Error(`Unknown tool: ${request.params.name}`);
  }
});

// Start server
async function main() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error("XLSX Reader MCP server running on stdio");
}

main().catch(console.error);