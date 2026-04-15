import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import {
  type MeetingSearchCriteria,
  rankMeetings,
  type ScoredMeeting,
} from "../lib/meeting-search.js";
import type { GraphService } from "../services/graph.js";
import type { CallTranscript, OnlineMeeting, TranscriptContent } from "../types/graph.js";

// ---------- Constants --------------------------------------------------------

export const FIND_MEETINGS_DEFAULT_LIMIT = 10;
export const FIND_MEETINGS_MAX_LIMIT = 25;
// Candidate cap is constrained by the N+1 query pattern: each calendar event
// requires a follow-up /me/onlineMeetings?$filter=joinWebUrl eq '...' lookup
// to resolve to an OnlineMeeting with a stable id usable for transcript calls.
export const FIND_MEETINGS_CANDIDATE_CAP = 50;
export const FIND_MEETINGS_DEFAULT_WINDOW_DAYS = 30;
export const LIST_MEETINGS_DEFAULT_TOP = 20;
export const LIST_MEETINGS_MAX_TOP = 50;

// ---------- Shared helpers ---------------------------------------------------

function isJoinUrl(value: string): boolean {
  return value.trim().startsWith("https://teams.microsoft.com/");
}

function defaultWindow(now: Date): { start: string; end: string } {
  const end = new Date(now);
  const start = new Date(now);
  start.setUTCDate(start.getUTCDate() - FIND_MEETINGS_DEFAULT_WINDOW_DAYS);
  return { start: start.toISOString(), end: end.toISOString() };
}

interface CalendarEvent {
  id?: string;
  subject?: string;
  onlineMeeting?: { joinUrl?: string } | null;
  isOnlineMeeting?: boolean;
}

/**
 * Lists the signed-in user's organized Teams meetings via the calendar endpoint.
 *
 * Microsoft Graph does NOT support listing /me/onlineMeetings directly — that
 * endpoint is lookup-only (filter by joinWebUrl or videoTeleconferenceId, or
 * GET by id). The correct delegated-auth path is to list /me/events with
 * isOnlineMeeting=true, then resolve each event's joinUrl to an OnlineMeeting.
 *
 * This is an N+1 pattern: 1 events call + N join-URL lookups. Callers should
 * cap `top` accordingly.
 */
async function fetchOrganizedMeetingsViaCalendar(
  graphService: GraphService,
  startDateTime: string | undefined,
  endDateTime: string | undefined,
  top: number
): Promise<OnlineMeeting[]> {
  const client = await graphService.getClient();

  const filterParts = ["isOnlineMeeting eq true"];
  if (startDateTime) filterParts.push(`start/dateTime ge '${startDateTime}'`);
  if (endDateTime) filterParts.push(`start/dateTime le '${endDateTime}'`);
  const eventsFilter = filterParts.join(" and ");

  const eventsResponse = (await client
    .api("/me/events")
    .filter(eventsFilter)
    .orderby("start/dateTime desc")
    .top(top)
    .select(
      "id,subject,start,end,organizer,attendees,onlineMeeting,isOnlineMeeting,onlineMeetingProvider"
    )
    .get()) as { value?: CalendarEvent[] };

  const events = eventsResponse?.value ?? [];

  const joinUrls = events
    .map((e) => e.onlineMeeting?.joinUrl)
    .filter((url): url is string => typeof url === "string" && url.length > 0);

  const resolved: OnlineMeeting[] = [];
  for (const joinUrl of joinUrls) {
    try {
      const response = (await client
        .api("/me/onlineMeetings")
        .filter(`joinWebUrl eq '${joinUrl}'`)
        .get()) as { value?: OnlineMeeting[] };
      const items = response?.value ?? [];
      if (items.length > 0 && items[0]) {
        resolved.push(items[0]);
      }
    } catch (error) {
      console.error(`Failed to resolve onlineMeeting for joinUrl ${joinUrl}:`, error);
    }
  }

  return resolved;
}

// ---------- listOnlineMeetings ----------------------------------------------

export interface ListOnlineMeetingsParams {
  startDateTime?: string;
  endDateTime?: string;
  subjectContains?: string;
  top?: number;
}

export async function listOnlineMeetings(
  graphService: GraphService,
  params: ListOnlineMeetingsParams
): Promise<OnlineMeeting[]> {
  const top = Math.min(params.top ?? LIST_MEETINGS_DEFAULT_TOP, LIST_MEETINGS_MAX_TOP);

  let items = await fetchOrganizedMeetingsViaCalendar(
    graphService,
    params.startDateTime,
    params.endDateTime,
    top
  );

  if (params.subjectContains) {
    const needle = params.subjectContains.toLowerCase();
    items = items.filter((m) => (m.subject ?? "").toLowerCase().includes(needle));
  }

  return items;
}

// ---------- getOnlineMeeting ------------------------------------------------

export interface GetOnlineMeetingParams {
  meetingIdOrJoinUrl: string;
}

export async function getOnlineMeeting(
  graphService: GraphService,
  params: GetOnlineMeetingParams
): Promise<OnlineMeeting | null> {
  const client = await graphService.getClient();
  const value = params.meetingIdOrJoinUrl.trim();

  if (isJoinUrl(value)) {
    const response = (await client
      .api("/me/onlineMeetings")
      .filter(`joinWebUrl eq '${value}'`)
      .get()) as { value?: OnlineMeeting[] };
    const items: OnlineMeeting[] = response?.value ?? [];
    return items.length > 0 ? (items[0] ?? null) : null;
  }

  const meeting = (await client.api(`/me/onlineMeetings/${value}`).get()) as OnlineMeeting | null;
  return meeting ?? null;
}

// ---------- listMeetingTranscripts ------------------------------------------

export interface ListMeetingTranscriptsParams {
  meetingId: string;
}

export async function listMeetingTranscripts(
  graphService: GraphService,
  params: ListMeetingTranscriptsParams
): Promise<CallTranscript[]> {
  const client = await graphService.getClient();
  const response = (await client
    .api(`/me/onlineMeetings/${params.meetingId}/transcripts`)
    .get()) as { value?: CallTranscript[] };
  return response?.value ?? [];
}

// ---------- getMeetingTranscriptContent -------------------------------------

export interface GetMeetingTranscriptContentParams {
  meetingId: string;
  transcriptId: string;
}

export async function getMeetingTranscriptContent(
  graphService: GraphService,
  params: GetMeetingTranscriptContentParams
): Promise<TranscriptContent> {
  const client = await graphService.getClient();
  const raw = await client
    .api(`/me/onlineMeetings/${params.meetingId}/transcripts/${params.transcriptId}/content`)
    .header("Accept", "text/vtt")
    .get();

  let content: string;
  if (typeof raw === "string") {
    content = raw;
  } else if (raw != null) {
    content = String(raw);
  } else {
    content = "";
  }

  return { format: "vtt", content };
}

// ---------- findMeetings (fuzzy multi-criteria) -----------------------------

export interface FindMeetingsParams extends MeetingSearchCriteria {
  limit?: number;
}

export async function findMeetings(
  graphService: GraphService,
  params: FindMeetingsParams,
  nowFactory: () => Date = () => new Date()
): Promise<ScoredMeeting[]> {
  const now = nowFactory();
  const limit = Math.min(params.limit ?? FIND_MEETINGS_DEFAULT_LIMIT, FIND_MEETINGS_MAX_LIMIT);

  // Expand date range to a sensible default if neither bound is provided.
  let start = params.startDateTime;
  let end = params.endDateTime;
  if (!start && !end) {
    const win = defaultWindow(now);
    start = win.start;
    end = win.end;
  }

  const candidates = await fetchOrganizedMeetingsViaCalendar(
    graphService,
    start,
    end,
    FIND_MEETINGS_CANDIDATE_CAP
  );

  const criteria: MeetingSearchCriteria = {};
  if (params.query !== undefined) criteria.query = params.query;
  if (params.subjectContains !== undefined) criteria.subjectContains = params.subjectContains;
  if (params.participantNameOrEmail !== undefined)
    criteria.participantNameOrEmail = params.participantNameOrEmail;
  if (start !== undefined) criteria.startDateTime = start;
  if (end !== undefined) criteria.endDateTime = end;

  return rankMeetings(candidates, criteria, now, limit);
}

// ---------- MCP tool registration -------------------------------------------

const DELEGATED_SCOPE_NOTE =
  " Note: Microsoft Graph only returns meetings the signed-in user organized. Meetings you only attended will not appear.";

export function registerMeetingTools(
  server: McpServer,
  graphService: GraphService,
  _readOnly: boolean
) {
  // find_meetings (primary user-facing entry point)
  server.registerTool(
    "find_meetings",
    {
      title: "Find Meetings",
      description:
        "Find Teams meetings using fuzzy multi-criteria search. Use this FIRST when the user references a meeting by name, attendees, or an approximate date. Supports free-text query, subject substring, participant name/email, and a date range. Returns the top matches ranked by a weighted score with human-readable match reasons." +
        DELEGATED_SCOPE_NOTE,
      inputSchema: {
        query: z
          .string()
          .optional()
          .describe("Free-text fuzzy query matched against meeting subject (e.g. 'q4 kickof')"),
        subjectContains: z
          .string()
          .optional()
          .describe("Case-insensitive substring the subject must contain"),
        participantNameOrEmail: z
          .string()
          .optional()
          .describe(
            "Substring match against any participant's display name, UPN, or email address"
          ),
        startDateTime: z
          .string()
          .optional()
          .describe("ISO 8601 start of the search window. Defaults to 30 days ago."),
        endDateTime: z
          .string()
          .optional()
          .describe("ISO 8601 end of the search window. Defaults to now."),
        limit: z
          .number()
          .min(1)
          .max(FIND_MEETINGS_MAX_LIMIT)
          .optional()
          .default(FIND_MEETINGS_DEFAULT_LIMIT)
          .describe(
            `Max results to return. Default ${FIND_MEETINGS_DEFAULT_LIMIT}, max ${FIND_MEETINGS_MAX_LIMIT}.`
          ),
      },
      annotations: {
        readOnlyHint: true,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    async (args) => {
      try {
        const results = await findMeetings(graphService, args as FindMeetingsParams);
        return {
          content: [
            {
              type: "text" as const,
              text: JSON.stringify(results, null, 2),
            },
          ],
        };
      } catch (error: unknown) {
        const errorMessage = error instanceof Error ? error.message : "Unknown error occurred";
        return {
          content: [
            {
              type: "text" as const,
              text: `❌ Error: ${errorMessage}`,
            },
          ],
        };
      }
    }
  );

  // list_online_meetings
  server.registerTool(
    "list_online_meetings",
    {
      title: "List Online Meetings",
      description:
        "List the signed-in user's online meetings chronologically. Prefer find_meetings when the user has a specific meeting in mind." +
        DELEGATED_SCOPE_NOTE,
      inputSchema: {
        startDateTime: z.string().optional().describe("ISO 8601 lower bound on startDateTime"),
        endDateTime: z.string().optional().describe("ISO 8601 upper bound on startDateTime"),
        subjectContains: z
          .string()
          .optional()
          .describe("Optional case-insensitive subject substring filter (client-side)"),
        top: z
          .number()
          .min(1)
          .max(LIST_MEETINGS_MAX_TOP)
          .optional()
          .default(LIST_MEETINGS_DEFAULT_TOP)
          .describe(
            `Max meetings to return. Default ${LIST_MEETINGS_DEFAULT_TOP}, max ${LIST_MEETINGS_MAX_TOP}.`
          ),
      },
      annotations: {
        readOnlyHint: true,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    async (args) => {
      try {
        const results = await listOnlineMeetings(graphService, args as ListOnlineMeetingsParams);
        return {
          content: [
            {
              type: "text" as const,
              text: JSON.stringify(results, null, 2),
            },
          ],
        };
      } catch (error: unknown) {
        const errorMessage = error instanceof Error ? error.message : "Unknown error occurred";
        return {
          content: [
            {
              type: "text" as const,
              text: `❌ Error: ${errorMessage}`,
            },
          ],
        };
      }
    }
  );

  // get_online_meeting
  server.registerTool(
    "get_online_meeting",
    {
      title: "Get Online Meeting",
      description:
        "Fetch a single online meeting by meeting ID or by Teams join URL (auto-detected)." +
        DELEGATED_SCOPE_NOTE,
      inputSchema: {
        meetingIdOrJoinUrl: z
          .string()
          .describe(
            "Either a meeting ID returned by find_meetings or list_online_meetings, or a full Teams join URL starting with https://teams.microsoft.com/"
          ),
      },
      annotations: {
        readOnlyHint: true,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    async (args) => {
      try {
        const meeting = await getOnlineMeeting(graphService, args);
        if (!meeting) {
          return {
            content: [
              {
                type: "text" as const,
                text: "No meeting found. If you used a join URL, confirm it matches exactly and that the meeting was organized by the signed-in user.",
              },
            ],
          };
        }
        return {
          content: [
            {
              type: "text" as const,
              text: JSON.stringify(meeting, null, 2),
            },
          ],
        };
      } catch (error: unknown) {
        const errorMessage = error instanceof Error ? error.message : "Unknown error occurred";
        return {
          content: [
            {
              type: "text" as const,
              text: `❌ Error: ${errorMessage}`,
            },
          ],
        };
      }
    }
  );

  // list_meeting_transcripts
  server.registerTool(
    "list_meeting_transcripts",
    {
      title: "List Meeting Transcripts",
      description:
        "List all available transcripts for a specific meeting. Returns an empty array if the meeting was not recorded with transcription enabled.",
      inputSchema: {
        meetingId: z
          .string()
          .describe(
            "Meeting ID returned by find_meetings, list_online_meetings, or get_online_meeting"
          ),
      },
      annotations: {
        readOnlyHint: true,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    async (args) => {
      try {
        const transcripts = await listMeetingTranscripts(graphService, args);
        return {
          content: [
            {
              type: "text" as const,
              text: JSON.stringify(transcripts, null, 2),
            },
          ],
        };
      } catch (error: unknown) {
        const errorMessage = error instanceof Error ? error.message : "Unknown error occurred";
        return {
          content: [
            {
              type: "text" as const,
              text: `❌ Error: ${errorMessage}`,
            },
          ],
        };
      }
    }
  );

  // get_meeting_transcript_content
  server.registerTool(
    "get_meeting_transcript_content",
    {
      title: "Get Meeting Transcript Content",
      description:
        "Fetch the full WebVTT content of a specific meeting transcript, including speaker attribution (e.g. <v Alice Example>). Use list_meeting_transcripts first to get the transcriptId.",
      inputSchema: {
        meetingId: z.string().describe("Meeting ID"),
        transcriptId: z.string().describe("Transcript ID from list_meeting_transcripts"),
      },
      annotations: {
        readOnlyHint: true,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    async (args) => {
      try {
        const result = await getMeetingTranscriptContent(graphService, args);
        return {
          content: [
            {
              type: "text" as const,
              text: result.content || "(empty transcript)",
            },
          ],
        };
      } catch (error: unknown) {
        const errorMessage = error instanceof Error ? error.message : "Unknown error occurred";
        return {
          content: [
            {
              type: "text" as const,
              text: `❌ Error: ${errorMessage}`,
            },
          ],
        };
      }
    }
  );
}
