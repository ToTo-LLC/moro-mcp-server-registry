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
export const FIND_MEETINGS_CANDIDATE_CAP = 200;
export const FIND_MEETINGS_DEFAULT_WINDOW_DAYS = 30;
export const LIST_MEETINGS_DEFAULT_TOP = 25;
export const LIST_MEETINGS_MAX_TOP = 50;

// ---------- Shared helpers ---------------------------------------------------

function isJoinUrl(value: string): boolean {
  return value.trim().startsWith("https://teams.microsoft.com/");
}

function buildDateFilter(start?: string, end?: string): string | undefined {
  const parts: string[] = [];
  if (start) parts.push(`startDateTime ge ${start}`);
  if (end) parts.push(`startDateTime le ${end}`);
  return parts.length > 0 ? parts.join(" and ") : undefined;
}

function defaultWindow(now: Date): { start: string; end: string } {
  const end = new Date(now);
  const start = new Date(now);
  start.setUTCDate(start.getUTCDate() - FIND_MEETINGS_DEFAULT_WINDOW_DAYS);
  return { start: start.toISOString(), end: end.toISOString() };
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

  const client = await graphService.getClient();
  let request: any = client.api("/me/onlineMeetings");

  const filter = buildDateFilter(params.startDateTime, params.endDateTime);
  if (filter) {
    request = request.filter(filter);
  }
  request = request.top(top);

  const response = (await request.get()) as { value?: OnlineMeeting[] };
  let items: OnlineMeeting[] = response?.value ?? [];

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

  const client = await graphService.getClient();
  let request: any = client.api("/me/onlineMeetings");
  const filter = buildDateFilter(start, end);
  if (filter) request = request.filter(filter);
  request = request.top(FIND_MEETINGS_CANDIDATE_CAP);

  const response = (await request.get()) as { value?: OnlineMeeting[] };
  const candidates: OnlineMeeting[] = response?.value ?? [];

  const criteria: MeetingSearchCriteria = {};
  if (params.query !== undefined) criteria.query = params.query;
  if (params.subjectContains !== undefined) criteria.subjectContains = params.subjectContains;
  if (params.participantNameOrEmail !== undefined)
    criteria.participantNameOrEmail = params.participantNameOrEmail;
  if (start !== undefined) criteria.startDateTime = start;
  if (end !== undefined) criteria.endDateTime = end;

  return rankMeetings(candidates, criteria, now, limit);
}
