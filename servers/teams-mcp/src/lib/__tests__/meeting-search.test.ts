import { describe, expect, it } from "vitest";
import type { OnlineMeeting } from "../../types/graph.js";
import { rankMeetings, scoreMeeting } from "../meeting-search.js";

const baseMeeting = (overrides: Partial<OnlineMeeting> = {}): OnlineMeeting => ({
  id: "1",
  subject: "Default Subject",
  startDateTime: "2026-04-10T15:00:00Z",
  endDateTime: "2026-04-10T16:00:00Z",
  joinWebUrl: "https://teams.microsoft.com/l/meetup-join/default",
  participants: {
    organizer: {
      identity: {
        user: { displayName: "Alice Example", userPrincipalName: "alice@example.com" },
      },
    },
    attendees: [
      {
        identity: {
          user: { displayName: "Bob Builder", userPrincipalName: "bob@example.com" },
        },
      },
    ],
  },
  ...overrides,
});

const NOW = new Date("2026-04-14T12:00:00Z");

describe("scoreMeeting", () => {
  it("returns 0 when no criteria provided", () => {
    const r = scoreMeeting(baseMeeting(), {}, NOW);
    expect(r.score).toBe(0);
    expect(r.matchReasons).toEqual([]);
  });

  it("subject substring match adds weight 3 and a matchReason", () => {
    const r = scoreMeeting(
      baseMeeting({ subject: "Q2 Kickoff Call" }),
      { subjectContains: "kickoff" },
      NOW
    );
    expect(r.score).toBeGreaterThanOrEqual(3);
    expect(r.matchReasons.some((m) => m.toLowerCase().includes("subject"))).toBe(true);
  });

  it("participant email match adds weight 2", () => {
    const r = scoreMeeting(baseMeeting(), { participantNameOrEmail: "bob@example.com" }, NOW);
    expect(r.score).toBeGreaterThanOrEqual(2);
    expect(r.matchReasons.some((m) => m.toLowerCase().includes("bob"))).toBe(true);
  });

  it("participant display name match adds weight 2 (case-insensitive)", () => {
    const r = scoreMeeting(baseMeeting(), { participantNameOrEmail: "alice" }, NOW);
    expect(r.score).toBeGreaterThanOrEqual(2);
  });

  it("ranks subject match above participant-only match", () => {
    const subjectHit = scoreMeeting(
      baseMeeting({ subject: "Kickoff meeting", id: "A" }),
      { subjectContains: "kickoff", participantNameOrEmail: "someone@other.com" },
      NOW
    );
    const participantHit = scoreMeeting(
      baseMeeting({ subject: "Weekly sync", id: "B" }),
      { subjectContains: "kickoff", participantNameOrEmail: "bob@example.com" },
      NOW
    );
    expect(subjectHit.score).toBeGreaterThan(participantHit.score);
  });

  it("recency decay gives slight bump to recent meetings", () => {
    const recent = scoreMeeting(
      baseMeeting({
        subject: "Kickoff",
        startDateTime: "2026-04-13T15:00:00Z",
        id: "R",
      }),
      { subjectContains: "kickoff" },
      NOW
    );
    const old = scoreMeeting(
      baseMeeting({
        subject: "Kickoff",
        startDateTime: "2026-01-15T15:00:00Z",
        id: "O",
      }),
      { subjectContains: "kickoff" },
      NOW
    );
    expect(recent.score).toBeGreaterThan(old.score);
  });

  it("free-text query fuzzy-matches misspellings in the subject", () => {
    const r = scoreMeeting(baseMeeting({ subject: "Kickoff Call" }), { query: "kickof" }, NOW);
    expect(r.score).toBeGreaterThan(0);
  });
});

describe("rankMeetings", () => {
  it("returns top N sorted by score descending", () => {
    const meetings: OnlineMeeting[] = [
      baseMeeting({ id: "low", subject: "Unrelated standup" }),
      baseMeeting({
        id: "high",
        subject: "Client kickoff call",
        startDateTime: "2026-04-13T15:00:00Z",
      }),
      baseMeeting({
        id: "mid",
        subject: "Weekly kickoff review",
        startDateTime: "2026-04-06T15:00:00Z",
      }),
    ];
    const ranked = rankMeetings(meetings, { subjectContains: "kickoff" }, NOW, 2);
    expect(ranked).toHaveLength(2);
    expect(ranked[0].id).toBe("high");
    expect(ranked[0].score).toBeGreaterThan(ranked[1].score);
  });

  it("filters out zero-score candidates when any text criterion was given", () => {
    const meetings: OnlineMeeting[] = [
      baseMeeting({ id: "1", subject: "Marketing sync" }),
      baseMeeting({ id: "2", subject: "Kickoff call" }),
    ];
    const ranked = rankMeetings(meetings, { subjectContains: "kickoff" }, NOW, 10);
    expect(ranked.map((m) => m.id)).toEqual(["2"]);
  });

  it("returns all candidates sorted by recency when no text criteria given", () => {
    const meetings: OnlineMeeting[] = [
      baseMeeting({ id: "old", startDateTime: "2026-01-01T00:00:00Z" }),
      baseMeeting({ id: "new", startDateTime: "2026-04-13T00:00:00Z" }),
    ];
    const ranked = rankMeetings(meetings, {}, NOW, 10);
    expect(ranked[0].id).toBe("new");
  });
});
