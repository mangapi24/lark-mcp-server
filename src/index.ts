import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { SSEServerTransport } from "@modelcontextprotocol/sdk/server/sse.js";
import { CallToolRequestSchema, ListToolsRequestSchema } from "@modelcontextprotocol/sdk/types.js";
import { z } from "zod";
import * as lark from "@larksuiteoapi/node-sdk";
import { isRecurrenceInRange } from './utils/recurrence.js';
import { createClient } from '@supabase/supabase-js';
import express from "express";

// Create Lark client instance
const client = new lark.Client({
  appId: process.env.LARK_APP_ID!,
  appSecret: process.env.LARK_APP_SECRET!,
  loggerLevel: "error" as any
});

// Create Supabase client instance
const supabaseUrl = process.env.SUPABASE_URL!
const supabaseKey = process.env.SUPABASE_KEY!
const supabase = createClient(supabaseUrl, supabaseKey)

// Validate schemas
const schemas = {
  toolInputs: {
    send_message_to_user: z.object({
      content: z.string()
    }),
    list_events: z.object({
      start_time: z.string(),
      end_time: z.string(),
      calendar_id: z.string()
    }),
    create_event: z.object({
      summary: z.string(),
      description: z.string().optional(),
      start_time: z.string(),
      end_time: z.string(),
      location: z.string().optional(),
      need_notification: z.boolean().optional()
    }),
    add_attendees: z.object({
      event_id: z.string(),
      attendees: z.array(
        z.object({
          type: z.enum(["user", "chat", "resource", "third_party"]),
          user_id: z.string().optional(),
          chat_id: z.string().optional(),
          resource_id: z.string().optional(),
          third_party_email: z.string().optional(),
          is_optional: z.boolean().optional(),
          operate_id: z.string().optional(),
          approval_reason: z.string().optional()
        })
      ),
      need_notification: z.boolean().optional()
    }),
    search_user_in_supabase: z.object({
      name_query: z.string().min(1)
    })
  }
}

// Tool definition
const TOOL_DEFINITIONS = [
  {
    name: "send_message_to_user",
    description: "Send a message to the user on Lark",
    inputSchema: {
      type: "object",
      properties: {
        content: {
          type: "string",
          description: "Content of the message"
        },
      },
      required: ["content"]
    },
  },
  {
    name: "list_events",
    description: "List events from the user on Lark",
    inputSchema: {
      type: "object",
      properties: {
        calendar_id: {
          type: "string",
          description: "ID of the calendar to list events from, default is the user's own calendar id please just use primary; If need a non-personal calendar_id please use search_user_in_supabase to find this person's calendar_id"
        },
        start_time: {
          type: "string",
          description: "Start time in ISO format with UTC+8 timezone (e.g. 2024-03-20T10:00:00+08:00)"
        },
        end_time: {
          type: "string",
          description: "End time in ISO format with UTC+8 timezone (e.g. 2024-03-20T11:00:00+08:00)"
        },
      },
    }
  },
  {
    name: "create_event",
    description: "Create a calendar event on Lark",
    inputSchema: {
      type: "object",
      properties: {
        summary: {
          type: "string",
          description: "Event title or summary"
        },
        description: {
          type: "string",
          description: "Event description (optional)"
        },
        start_time: {
          type: "string",
          description: "Event start time in ISO format with UTC+8 timezone (e.g. 2024-03-20T10:00:00+08:00)"
        },
        end_time: {
          type: "string",
          description: "Event end time in ISO format with UTC+8 timezone (e.g. 2024-03-20T11:00:00+08:00)"
        },
        location: {
          type: "string",
          description: "Event location (optional)"
        },
        need_notification: {
          type: "boolean",
          description: "Whether to send notification to participants (default: true)"
        }
      },
      required: ["summary", "start_time", "end_time"]
    },
  },
  {
    name: "add_attendees",
    description: "Add attendees to a calendar event on Lark",
    inputSchema: {
      type: "object",
      properties: {
        event_id: {
          type: "string",
          description: "ID of the event to add attendees to"
        },
        attendees: {
          type: "array",
          description: "List of attendees to add to the event",
          items: {
            type: "object",
            properties: {
              type: {
                type: "string",
                description: "Type of attendee: 'user', 'chat', 'resource', or 'third_party'",
                enum: ["user", "chat", "resource", "third_party"]
              },
              user_id: {
                type: "string",
                description: "User ID when type is 'user',please use search_user_in_supabase to find this person's user_id"
              },
              chat_id: {
                type: "string",
                description: "Chat/Group ID when type is 'chat'"
              },
              resource_id: {
                type: "string",
                description: "Resource (meeting room) ID when type is 'resource'"
              },
              third_party_email: {
                type: "string",
                description: "Email address when type is 'third_party'"
              },
              is_optional: {
                type: "boolean",
                description: "Whether the attendee is optional (default: false)"
              },
              operate_id: {
                type: "string",
                description: "Operator ID for room booking contact"
              },
              approval_reason: {
                type: "string",
                description: "Reason for booking a room"
              }
            },
            required: ["type"]
          }
        },
        need_notification: {
          type: "boolean",
          description: "Whether to send notifications to added attendees (default: true)"
        }
      },
      required: ["event_id", "attendees"]
    }
  },
  {
    name: "search_user_in_supabase",
    description: "Search for a user in Supabase database by partial name and get their user ID and calendar ID",
    inputSchema: {
      type: "object",
      properties: {
        name_query: {
          type: "string",
          description: "Partial name to search for in the user_name field (case insensitive)"
        }
      },
      required: ["name_query"]
    },
  }
]

// Tool implementation handlers
const toolHandlers = {
  async send_message_to_user(args: unknown) {
    const { content } = schemas.toolInputs.send_message_to_user.parse(args);
    try {
      const messageContent = JSON.stringify({ text: content });
      const result = await client.im.message.create({
          params: { receive_id_type: "user_id" },
          data: {
            receive_id: process.env.LARK_USER_ID!,
            msg_type: "text",
            content: messageContent
          },
      });
      if (!result || result.code !== 0) {
        return { content: [{ type: "text" as const, text: `Failed to send message: ${result?.msg || "Unknown error"}` }] };
      }
      return { content: [{ type: "text" as const, text: `Message sent successfully! Message ID: ${result.data?.message_id || "unknown"}` }] };
    } catch (error) {
      return { content: [{ type: "text" as const, text: `Error sending message: ${error instanceof Error ? error.message : "Unknown error"}` }] };
    }
  },

  async list_events(args: unknown) {
    const { start_time, end_time, calendar_id } = schemas.toolInputs.list_events.parse(args);
    const startUnix = Math.floor(new Date(start_time).getTime() / 1000).toString();
    const endUnix = Math.floor(new Date(end_time).getTime() / 1000).toString();
    const result = await client.calendar.v4.calendarEvent.list({
      path: { calendar_id: calendar_id === "primary" ? process.env.LARK_CALENDAR_ID! : calendar_id },
      params: { page_size: 1000, start_time: startUnix, end_time: endUnix }
    }, lark.withUserAccessToken(process.env.LARK_USER_ACCESS_TOKEN!));
    if (!result || result.code !== 0) {
      return { content: [{ type: "text" as const, text: `Failed to list events: ${result?.msg || "Unknown error"}` }] };
    }
    const allEvents = result.data?.items || [];
    const events = allEvents.filter(event =>
      (event.status !== "cancelled" && event.start_time?.timestamp && event.end_time?.timestamp &&
        parseInt(event.start_time.timestamp) >= parseInt(startUnix) && parseInt(event.end_time.timestamp) <= parseInt(endUnix)) ||
      (event.recurrence !== "" && event.status !== "cancelled" &&
        isRecurrenceInRange(event.recurrence || "", { startTime: parseInt(event.start_time?.timestamp || "0"), endTime: parseInt(event.end_time?.timestamp || "0") }, { startTime: parseInt(startUnix), endTime: parseInt(endUnix) }))
    );
    const formattedEvents = events.map(event => ({
      summary: event.summary || "",
      organizer: event.event_organizer?.display_name || "",
      status: event.status || "unknown",
      startTime: new Date(parseInt(event.start_time?.timestamp || "0") * 1000).toLocaleString(),
      endTime: new Date(parseInt(event.end_time?.timestamp || "0") * 1000).toLocaleString(),
      rerecurrence: event.recurrence
    }));
    return { content: [{ type: "text" as const, text: events.length ? `Found ${events.length} active events:\n\n${JSON.stringify(formattedEvents, null, 2)}` : "No active events found in the given time range" }] };
  },

  async create_event(args: unknown) {
    const { summary, description, start_time, end_time, location, need_notification } = schemas.toolInputs.create_event.parse(args);
    try {
      const startTimestamp = Math.floor(Date.parse(start_time) / 1000).toString();
      const endTimestamp = Math.floor(Date.parse(end_time) / 1000).toString();
      const requestData: any = {
        summary, need_notification: need_notification ?? true,
        start_time: { timestamp: startTimestamp, timezone: "Asia/Shanghai" },
        end_time: { timestamp: endTimestamp, timezone: "Asia/Shanghai" },
        attendee_ability: "can_modify_event"
      };
      if (description) requestData.description = description;
      if (location) requestData.location = { name: location };
      const result = await client.calendar.v4.calendarEvent.create({
        path: { calendar_id: process.env.LARK_CALENDAR_ID! },
        data: requestData
      }, lark.withUserAccessToken(process.env.LARK_USER_ACCESS_TOKEN!));
      if (!result || result.code !== 0) {
        return { content: [{ type: "text" as const, text: `Failed to create event: ${result?.msg || "Unknown error"}` }] };
      }
      const eventId = result.data?.event?.event_id || "unknown";
      if (eventId !== "unknown") {
        await client.calendar.v4.calendarEventAttendee.create({
          path: { calendar_id: process.env.LARK_CALENDAR_ID!, event_id: eventId },
          params: { user_id_type: "user_id" },
          data: { attendees: [{ type: "user", user_id: process.env.LARK_USER_ID!, is_optional: false }], need_notification: false }
        }, lark.withUserAccessToken(process.env.LARK_USER_ACCESS_TOKEN!));
      }
      return { content: [{ type: "text" as const, text: `Event "${result.data?.event?.summary || summary}" created successfully!\nEvent ID: ${eventId}` }] };
    } catch (error) {
      return { content: [{ type: "text" as const, text: `Error creating event: ${error instanceof Error ? error.message : "Unknown error"}` }] };
    }
  },

  async add_attendees(args: unknown) {
    const { event_id, attendees, need_notification } = schemas.toolInputs.add_attendees.parse(args);
    try {
      const transformedAttendees = attendees.map(attendee => {
        const t: any = { type: attendee.type, is_optional: attendee.is_optional || false };
        if (attendee.type === "user") t.user_id = attendee.user_id;
        if (attendee.type === "chat") t.chat_id = attendee.chat_id;
        if (attendee.type === "resource") t.resource_id = attendee.resource_id;
        if (attendee.type === "third_party") t.third_party_email = attendee.third_party_email;
        if (attendee.operate_id) t.operate_id = attendee.operate_id;
        if (attendee.approval_reason) t.approval_reason = attendee.approval_reason;
        return t;
      });
      const result = await client.calendar.v4.calendarEventAttendee.create({
        path: { calendar_id: process.env.LARK_CALENDAR_ID!, event_id },
        params: { user_id_type: "user_id" },
        data: { attendees: transformedAttendees, need_notification: need_notification ?? true }
      }, lark.withUserAccessToken(process.env.LARK_USER_ACCESS_TOKEN!));
      if (!result || result.code !== 0) {
        return { content: [{ type: "text" as const, text: `Failed to add attendees: ${result?.msg || "Unknown error"}` }] };
      }
      return { content: [{ type: "text" as const, text: `Successfully added ${result.data?.attendees?.length || 0} attendee(s) to the event` }] };
    } catch (error) {
      return { content: [{ type: "text" as const, text: `Error adding attendees: ${error instanceof Error ? error.message : "Unknown error"}` }] };
    }
  },

  async search_user_in_supabase(args: unknown) {
    const { name_query } = schemas.toolInputs.search_user_in_supabase.parse(args);
    try {
      const { data, error } = await supabase.from('lark_members').select('user_id, user_name, user_main_calendar_id').ilike('user_name', `%${name_query}%`);
      if (error) return { content: [{ type: "text" as const, text: `Error querying Supabase: ${error.message}` }] };
      if (!data || data.length === 0) return { content: [{ type: "text" as const, text: `No users found with name containing "${name_query}".` }] };
      return { content: [{ type: "text" as const, text: `Found ${data.length} user(s) matching "${name_query}":\n\n${JSON.stringify(data.map(u => ({ user_id: u.user_id, user_name: u.user_name, calendar_id: u.user_main_calendar_id })), null, 2)}` }] };
    } catch (error) {
      return { content: [{ type: "text" as const, text: `Error searching users: ${error instanceof Error ? error.message : "Unknown error"}` }] };
    }
  }
};

// Create server instance
const server = new Server({ name: "lark-mcp-server", version: "1.0.0" }, { capabilities: { tools: {} } });

server.setRequestHandler(ListToolsRequestSchema, async () => ({ tools: TOOL_DEFINITIONS }));

server.setRequestHandler(CallToolRequestSchema, async (request) => {
  const { name, arguments: args } = request.params;
  const handler = toolHandlers[name as keyof typeof toolHandlers];
  if (!handler) throw new Error(`Unknown tool: ${name}`);
  return await handler(args);
});

// HTTP/SSE server for Claude.ai custom connector
async function main() {
  const requiredEnvVars = ['LARK_APP_ID','LARK_APP_SECRET','LARK_USER_ID','LARK_CALENDAR_ID','LARK_USER_ACCESS_TOKEN','SUPABASE_KEY','SUPABASE_URL'];
  const missingVars = requiredEnvVars.filter(v => !process.env[v]);
  if (missingVars.length > 0) {
    console.error(`Missing required environment variables: ${missingVars.join(', ')}`);
    process.exit(1);
  }

  const app = express();
  app.use(express.json());

  const transports: Record<string, SSEServerTransport> = {};

  app.get("/sse", async (req, res) => {
    console.error("New SSE connection");
    const transport = new SSEServerTransport("/messages", res);
    transports[transport.sessionId] = transport;
    res.on("close", () => delete transports[transport.sessionId]);
    await server.connect(transport);
  });

  app.post("/messages", async (req, res) => {
    const sessionId = req.query.sessionId as string;
    const transport = transports[sessionId];
    if (transport) await transport.handlePostMessage(req, res);
    else res.status(404).send("Session not found");
  });

  app.get("/", (_req, res) => res.json({ status: "ok", message: "Lark MCP Server running" }));

  const PORT = process.env.PORT || 3000;
  app.listen(PORT, () => console.error(`Lark MCP HTTP/SSE server running on port ${PORT}`));
}

main().catch(err => { console.error("Fatal error:", err); process.exit(1); });
