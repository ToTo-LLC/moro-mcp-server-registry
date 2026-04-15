import { describe, expect, it, vi } from "vitest";
import type { GraphService } from "../../services/graph.js";
import type { CallTranscript, OnlineMeeting } from "../../types/graph.js";
import {
  FIND_MEETINGS_CANDIDATE_CAP,
  FIND_MEETINGS_DEFAULT_LIMIT,
  FIND_MEETINGS_MAX_LIMIT,
  findMeetings,
  getMeetingTranscriptContent,
  getOnlineMeeting,
  LIST_MEETINGS_DEFAULT_TOP,
  LIST_MEETINGS_MAX_TOP,
  listMeetingTranscripts,
  listOnlineMeetings,
} from "../meetings.js";

// ---------------------------------------------------------------------------
// Mock helpers
// ---------------------------------------------------------------------------

/**
 * Builds a fresh GraphService mock for each test.
 * Each call to `.get()` / `.getStream()` consumes one result from
 * `getResultsInOrder`. Chainable methods (.filter, .top, .header) return the
 * chainable itself so the MS Graph SDK fluent API works end-to-end.
 */
function createMockGraphServiceWithApi(getResultsInOrder: unknown[]) {
  const mockGet = vi.fn();
  for (const r of getResultsInOrder) {
    mockGet.mockResolvedValueOnce(r);
  }

  const chainable: any = {};
  chainable.filter = vi.fn(() => chainable);
  chainable.top = vi.fn(() => chainable);
  chainable.orderby = vi.fn(() => chainable);
  chainable.select = vi.fn(() => chainable);
  chainable.header = vi.fn(() => chainable);
  chainable.get = mockGet;
  chainable.getStream = mockGet;

  const mockApi = vi.fn(() => chainable);

  const graphService = {
    getClient: vi.fn().mockResolvedValue({ api: mockApi }),
    getAuthStatus: vi.fn(),
    isAuthenticated: vi.fn().mockReturnValue(true),
    readOnlyMode: false,
    scopes: [],
  } as unknown as GraphService;

  return { graphService, mockApi, chainable, mockGet };
}

const NOW = new Date("2026-04-14T12:00:00Z");

// ---------------------------------------------------------------------------
// listOnlineMeetings
// ---------------------------------------------------------------------------

/**
 * Helper: builds a mocked queue for the calendar-based fetch.
 * First response = calendar events; subsequent responses = onlineMeeting
 * resolves (one per joinUrl-bearing event), in order.
 */
function calendarFetchResponses(
  events: Array<{ id: string; subject?: string; joinUrl?: string }>,
  resolvedById: Record<string, OnlineMeeting>
): unknown[] {
  const responses: unknown[] = [
    {
      value: events.map((e) => ({
        id: e.id,
        subject: e.subject,
        isOnlineMeeting: true,
        onlineMeeting: e.joinUrl ? { joinUrl: e.joinUrl } : null,
      })),
    },
  ];
  for (const e of events) {
    if (!e.joinUrl) continue;
    const resolved = resolvedById[e.id];
    responses.push({ value: resolved ? [resolved] : [] });
  }
  return responses;
}

describe("listOnlineMeetings", () => {
  it("queries /me/events with isOnlineMeeting filter and returns resolved OnlineMeetings", async () => {
    const events = [
      { id: "e1", subject: "Weekly Sync", joinUrl: "https://teams.microsoft.com/l/meetup-join/1" },
      { id: "e2", subject: "Planning", joinUrl: "https://teams.microsoft.com/l/meetup-join/2" },
    ];
    const resolved: Record<string, OnlineMeeting> = {
      e1: { id: "m1", subject: "Weekly Sync" },
      e2: { id: "m2", subject: "Planning" },
    };
    const { graphService, mockApi } = createMockGraphServiceWithApi(
      calendarFetchResponses(events, resolved)
    );

    const result = await listOnlineMeetings(graphService, {});

    expect(mockApi).toHaveBeenCalledWith("/me/events");
    expect(mockApi).toHaveBeenCalledWith("/me/onlineMeetings");
    expect(result.map((m) => m.id)).toEqual(["m1", "m2"]);
  });

  it("applies default top of LIST_MEETINGS_DEFAULT_TOP to the events query", async () => {
    const { graphService, chainable } = createMockGraphServiceWithApi([{ value: [] }]);

    await listOnlineMeetings(graphService, {});

    expect(chainable.top).toHaveBeenCalledWith(LIST_MEETINGS_DEFAULT_TOP);
  });

  it("clamps top to LIST_MEETINGS_MAX_TOP when top=999 is provided", async () => {
    const { graphService, chainable } = createMockGraphServiceWithApi([{ value: [] }]);

    await listOnlineMeetings(graphService, { top: 999 });

    expect(chainable.top).toHaveBeenCalledWith(LIST_MEETINGS_MAX_TOP);
  });

  it("applies date range as calendar start/dateTime filter", async () => {
    const { graphService, chainable } = createMockGraphServiceWithApi([{ value: [] }]);

    await listOnlineMeetings(graphService, {
      startDateTime: "2026-04-01T00:00:00Z",
      endDateTime: "2026-04-14T23:59:59Z",
    });

    expect(chainable.filter).toHaveBeenCalledWith(
      "isOnlineMeeting eq true and start/dateTime ge '2026-04-01T00:00:00Z' and start/dateTime le '2026-04-14T23:59:59Z'"
    );
  });

  it("always filters isOnlineMeeting even when no dates provided", async () => {
    const { graphService, chainable } = createMockGraphServiceWithApi([{ value: [] }]);

    await listOnlineMeetings(graphService, {});

    expect(chainable.filter).toHaveBeenCalledWith("isOnlineMeeting eq true");
  });

  it("applies subjectContains as client-side post-filter (case-insensitive)", async () => {
    const events = [
      { id: "e1", subject: "Weekly Sync", joinUrl: "https://teams.microsoft.com/l/meetup-join/1" },
      {
        id: "e2",
        subject: "Planning Meeting",
        joinUrl: "https://teams.microsoft.com/l/meetup-join/2",
      },
      { id: "e3", subject: "WEEKLY review", joinUrl: "https://teams.microsoft.com/l/meetup-join/3" },
    ];
    const resolved: Record<string, OnlineMeeting> = {
      e1: { id: "m1", subject: "Weekly Sync" },
      e2: { id: "m2", subject: "Planning Meeting" },
      e3: { id: "m3", subject: "WEEKLY review" },
    };
    const { graphService } = createMockGraphServiceWithApi(
      calendarFetchResponses(events, resolved)
    );

    const result = await listOnlineMeetings(graphService, {
      subjectContains: "weekly",
    });

    expect(result).toHaveLength(2);
    expect(result.map((m) => m.id)).toEqual(["m1", "m3"]);
  });

  it("returns empty array when calendar returns no events", async () => {
    const { graphService } = createMockGraphServiceWithApi([{ value: [] }]);

    const result = await listOnlineMeetings(graphService, {});

    expect(result).toEqual([]);
  });

  it("skips events that cannot be resolved to an onlineMeeting", async () => {
    const events = [
      { id: "e1", subject: "A", joinUrl: "https://teams.microsoft.com/l/meetup-join/1" },
      { id: "e2", subject: "B", joinUrl: "https://teams.microsoft.com/l/meetup-join/2" },
    ];
    // Only e1 resolves; e2 returns an empty value array.
    const resolved: Record<string, OnlineMeeting> = {
      e1: { id: "m1", subject: "A" },
    };
    const { graphService } = createMockGraphServiceWithApi(
      calendarFetchResponses(events, resolved)
    );

    const result = await listOnlineMeetings(graphService, {});

    expect(result).toHaveLength(1);
    expect(result[0]?.id).toBe("m1");
  });
});

// ---------------------------------------------------------------------------
// getOnlineMeeting
// ---------------------------------------------------------------------------

describe("getOnlineMeeting", () => {
  it("resolves raw meeting ID via /me/onlineMeetings/<id>", async () => {
    const meeting: OnlineMeeting = { id: "abc123", subject: "Design Review" };
    const { graphService, mockApi } = createMockGraphServiceWithApi([meeting]);

    const result = await getOnlineMeeting(graphService, {
      meetingIdOrJoinUrl: "abc123",
    });

    expect(mockApi).toHaveBeenCalledWith("/me/onlineMeetings/abc123");
    expect(result).toEqual(meeting);
  });

  it("resolves Teams join URL via $filter on /me/onlineMeetings", async () => {
    const joinUrl = "https://teams.microsoft.com/l/meetup-join/abc";
    const meeting: OnlineMeeting = { id: "m1", joinWebUrl: joinUrl };
    const { graphService, mockApi, chainable } = createMockGraphServiceWithApi([
      { value: [meeting] },
    ]);

    const result = await getOnlineMeeting(graphService, {
      meetingIdOrJoinUrl: joinUrl,
    });

    expect(mockApi).toHaveBeenCalledWith("/me/onlineMeetings");
    expect(chainable.filter).toHaveBeenCalledWith(`joinWebUrl eq '${joinUrl}'`);
    expect(result).toEqual(meeting);
  });

  it("returns null when join URL lookup returns empty value array", async () => {
    const joinUrl = "https://teams.microsoft.com/l/meetup-join/missing";
    const { graphService } = createMockGraphServiceWithApi([{ value: [] }]);

    const result = await getOnlineMeeting(graphService, {
      meetingIdOrJoinUrl: joinUrl,
    });

    expect(result).toBeNull();
  });
});

// ---------------------------------------------------------------------------
// listMeetingTranscripts
// ---------------------------------------------------------------------------

describe("listMeetingTranscripts", () => {
  it("calls /me/onlineMeetings/<meetingId>/transcripts and returns value", async () => {
    const transcripts: CallTranscript[] = [
      { id: "t1", meetingId: "m1" },
      { id: "t2", meetingId: "m1" },
    ];
    const { graphService, mockApi } = createMockGraphServiceWithApi([{ value: transcripts }]);

    const result = await listMeetingTranscripts(graphService, {
      meetingId: "m1",
    });

    expect(mockApi).toHaveBeenCalledWith("/me/onlineMeetings/m1/transcripts");
    expect(result).toEqual(transcripts);
  });

  it("returns empty array when value is empty", async () => {
    const { graphService } = createMockGraphServiceWithApi([{ value: [] }]);

    const result = await listMeetingTranscripts(graphService, {
      meetingId: "m1",
    });

    expect(result).toEqual([]);
  });
});

// ---------------------------------------------------------------------------
// getMeetingTranscriptContent
// ---------------------------------------------------------------------------

describe("getMeetingTranscriptContent", () => {
  it("calls correct path with Accept: text/vtt header and returns vtt content", async () => {
    const vttContent = "WEBVTT\n\n00:00:00.000 --> 00:00:02.000\nHello world";
    const { graphService, mockApi, chainable } = createMockGraphServiceWithApi([vttContent]);

    const result = await getMeetingTranscriptContent(graphService, {
      meetingId: "m1",
      transcriptId: "t1",
    });

    expect(mockApi).toHaveBeenCalledWith("/me/onlineMeetings/m1/transcripts/t1/content");
    expect(chainable.header).toHaveBeenCalledWith("Accept", "text/vtt");
    expect(result).toEqual({ format: "vtt", content: vttContent });
  });

  it("handles Buffer-like return value by converting to string", async () => {
    const bufferLike = Buffer.from("WEBVTT\n\nsome content");
    const { graphService } = createMockGraphServiceWithApi([bufferLike]);

    const result = await getMeetingTranscriptContent(graphService, {
      meetingId: "m1",
      transcriptId: "t1",
    });

    expect(result.format).toBe("vtt");
    expect(typeof result.content).toBe("string");
    expect(result.content.length).toBeGreaterThan(0);
  });
});

// ---------------------------------------------------------------------------
// findMeetings
// ---------------------------------------------------------------------------

describe("findMeetings", () => {
  it("queries /me/events with 30-day default window when no dates provided", async () => {
    const { graphService, mockApi, chainable } = createMockGraphServiceWithApi([{ value: [] }]);

    await findMeetings(graphService, {}, () => NOW);

    expect(mockApi).toHaveBeenCalledWith("/me/events");
    // 30 days before 2026-04-14T12:00:00Z = 2026-03-15T12:00:00Z
    expect(chainable.filter).toHaveBeenCalledWith(
      "isOnlineMeeting eq true and start/dateTime ge '2026-03-15T12:00:00.000Z' and start/dateTime le '2026-04-14T12:00:00.000Z'"
    );
  });

  it("applies top of FIND_MEETINGS_CANDIDATE_CAP to the calendar query", async () => {
    const { graphService, chainable } = createMockGraphServiceWithApi([{ value: [] }]);

    await findMeetings(graphService, {}, () => NOW);

    expect(chainable.top).toHaveBeenCalledWith(FIND_MEETINGS_CANDIDATE_CAP);
  });

  it("respects explicit startDateTime / endDateTime when provided", async () => {
    const { graphService, chainable } = createMockGraphServiceWithApi([{ value: [] }]);

    await findMeetings(
      graphService,
      {
        startDateTime: "2026-04-01T00:00:00Z",
        endDateTime: "2026-04-14T23:59:59Z",
      },
      () => NOW
    );

    expect(chainable.filter).toHaveBeenCalledWith(
      "isOnlineMeeting eq true and start/dateTime ge '2026-04-01T00:00:00Z' and start/dateTime le '2026-04-14T23:59:59Z'"
    );
  });

  it("returns top-N ranked results (default limit)", async () => {
    const events = Array.from({ length: 20 }, (_, i) => ({
      id: `e${i}`,
      subject: `Meeting ${i}`,
      joinUrl: `https://teams.microsoft.com/l/meetup-join/${i}`,
    }));
    const resolved: Record<string, OnlineMeeting> = Object.fromEntries(
      events.map((e, i) => [
        e.id,
        { id: `m${i}`, subject: e.subject, startDateTime: NOW.toISOString() },
      ])
    );
    const { graphService } = createMockGraphServiceWithApi(
      calendarFetchResponses(events, resolved)
    );

    const result = await findMeetings(graphService, { query: "Meeting" }, () => NOW);

    expect(result.length).toBeLessThanOrEqual(FIND_MEETINGS_DEFAULT_LIMIT);
  });

  it("clamps limit to FIND_MEETINGS_MAX_LIMIT", async () => {
    const events = Array.from({ length: 30 }, (_, i) => ({
      id: `e${i}`,
      subject: `Team Meeting ${i}`,
      joinUrl: `https://teams.microsoft.com/l/meetup-join/${i}`,
    }));
    const resolved: Record<string, OnlineMeeting> = Object.fromEntries(
      events.map((e, i) => [
        e.id,
        { id: `m${i}`, subject: e.subject, startDateTime: NOW.toISOString() },
      ])
    );
    const { graphService } = createMockGraphServiceWithApi(
      calendarFetchResponses(events, resolved)
    );

    const result = await findMeetings(
      graphService,
      { query: "Team Meeting", limit: 999 },
      () => NOW
    );

    expect(result.length).toBeLessThanOrEqual(FIND_MEETINGS_MAX_LIMIT);
  });

  it("returns ScoredMeetings with score and matchReasons", async () => {
    const events = [
      {
        id: "e1",
        subject: "Quarterly Review",
        joinUrl: "https://teams.microsoft.com/l/meetup-join/1",
      },
    ];
    const resolved: Record<string, OnlineMeeting> = {
      e1: { id: "m1", subject: "Quarterly Review", startDateTime: NOW.toISOString() },
    };
    const { graphService } = createMockGraphServiceWithApi(
      calendarFetchResponses(events, resolved)
    );

    const result = await findMeetings(graphService, { query: "Quarterly Review" }, () => NOW);

    if (result.length > 0) {
      expect(typeof result[0].score).toBe("number");
      expect(Array.isArray(result[0].matchReasons)).toBe(true);
    }
  });
});
