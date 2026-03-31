import express from "express";
import { randomUUID } from "crypto";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { SSEServerTransport } from "@modelcontextprotocol/sdk/server/sse.js";

// Import your existing tool registration from index.ts
// We re-use the same server setup but expose it via HTTP+SSE
import { createLarkMcpServer } from "./server.js";

const app = express();
const PORT = process.env.PORT || 3000;

// Store active SSE transports
const transports = new Map<string, SSEServerTransport>();

// SSE endpoint — Claude.ai connects here
app.get("/sse", async (req, res) => {
  const sessionId = randomUUID();
  console.log(`New SSE connection: ${sessionId}`);

  res.setHeader("Content-Type", "text/event-stream");
  res.setHeader("Cache-Control", "no-cache");
  res.setHeader("Connection", "keep-alive");
  res.setHeader("Access-Control-Allow-Origin", "*");

  const transport = new SSEServerTransport(`/messages?sessionId=${sessionId}`, res);
  transports.set(sessionId, transport);

  const server = await createLarkMcpServer();
  await server.connect(transport);

  req.on("close", () => {
    transports.delete(sessionId);
    console.log(`SSE connection closed: ${sessionId}`);
  });
});

// Message endpoint — Claude.ai posts tool calls here
app.post("/messages", express.json(), async (req, res) => {
  const sessionId = req.query.sessionId as string;
  const transport = transports.get(sessionId);

  if (!transport) {
    return res.status(404).json({ error: "Session not found" });
  }

  await transport.handlePostMessage(req, res);
});

// Health check
app.get("/", (req, res) => {
  res.json({ status: "ok", message: "Lark MCP Server running" });
});

app.listen(PORT, () => {
  console.log(`Lark MCP HTTP/SSE server running on port ${PORT}`);
});
