// mcp/server.ts
import { McpPlugin } from '@microsoft/teams.mcp';
import { z } from 'zod';

export const mcpServerPlugin = new McpPlugin({
  name: 'mcpServer',
  description: 'Provides developer community insights like GitHub issues and StackOverflow discussions',
  inspector: 'http://localhost:5173?proxyPort=9000' // Optional for local MCP DevTools
})
.tool(
  'echo',
  'Echoes back whatever input is given',
  {
    input: z.string().describe('The text to echo')
  },
  {
  title: "Echo Utility",
  idempotentHint: true
  }, // <- Add empty annotations object if no hints
  async ({ input }) => ({
    content: [{ type: 'text', text: `You said "${input}"` }]
  })
);

// ✅ Starts the internal HTTP plugin and exposes `/mcp`
mcpServerPlugin.onInit();
mcpServerPlugin.onStart({ port: 3978 });