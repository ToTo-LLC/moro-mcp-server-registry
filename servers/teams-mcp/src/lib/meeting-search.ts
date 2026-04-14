import Fuse from "fuse.js";
import type { MeetingParticipant, OnlineMeeting } from "../types/graph.js";

export interface MeetingSearchCriteria {
  query?: string;
  subjectContains?: string;
  participantNameOrEmail?: string;
  startDateTime?: string;
  endDateTime?: string;
}

export interface ScoredMeeting extends OnlineMeeting {
  score: number;
  matchReasons: string[];
}

const WEIGHT_SUBJECT = 3;
const WEIGHT_PARTICIPANT = 2;
const WEIGHT_RECENCY = 1;

const RECENCY_WINDOW_MS = 1000 * 60 * 60 * 24 * 90; // 90 days

function participantsAsStrings(meeting: OnlineMeeting): string[] {
  const out: string[] = [];
  const push = (p: MeetingParticipant | null | undefined): void => {
    if (!p) return;
    const user = p.identity?.user;
    if (user?.displayName) out.push(user.displayName);
    if (user?.userPrincipalName) out.push(user.userPrincipalName);
    if (p.upn) out.push(p.upn);
  };
  push(meeting.participants?.organizer ?? null);
  for (const a of meeting.participants?.attendees ?? []) push(a);
  return out;
}

function recencyScore(meeting: OnlineMeeting, now: Date): number {
  if (!meeting.startDateTime) return 0;
  const start = new Date(meeting.startDateTime).getTime();
  const delta = Math.abs(now.getTime() - start);
  if (delta >= RECENCY_WINDOW_MS) return 0;
  // Linear decay from 1.0 at now to 0.0 at the window edge.
  return WEIGHT_RECENCY * (1 - delta / RECENCY_WINDOW_MS);
}

export function scoreMeeting(
  meeting: OnlineMeeting,
  criteria: MeetingSearchCriteria,
  now: Date
): { score: number; matchReasons: string[] } {
  let score = 0;
  const matchReasons: string[] = [];

  const subject = (meeting.subject ?? "").toLowerCase();

  if (criteria.subjectContains) {
    const needle = criteria.subjectContains.toLowerCase();
    if (subject.includes(needle)) {
      // Base weight + small density bonus (needle/subject ratio) to break ties.
      const density = needle.length / (subject.length || 1);
      score += WEIGHT_SUBJECT + density;
      matchReasons.push(`subject contains "${criteria.subjectContains}"`);
    }
  }

  if (criteria.query) {
    // Fuzzy match the query against subject only via Fuse.
    const fuse = new Fuse([{ subject: meeting.subject ?? "" }], {
      keys: ["subject"],
      includeScore: true,
      threshold: 0.4,
    });
    const hits = fuse.search(criteria.query);
    if (hits.length > 0) {
      // Fuse score: 0 = perfect, 1 = worst. Invert and weight by subject.
      const fuseScore = hits[0].score ?? 1;
      const contribution = WEIGHT_SUBJECT * (1 - fuseScore);
      if (contribution > 0) {
        score += contribution;
        matchReasons.push(`subject fuzzy-matches "${criteria.query}"`);
      }
    }
  }

  if (criteria.participantNameOrEmail) {
    const needle = criteria.participantNameOrEmail.toLowerCase();
    const strings = participantsAsStrings(meeting).map((s) => s.toLowerCase());
    const hit = strings.find((s) => s.includes(needle));
    if (hit) {
      score += WEIGHT_PARTICIPANT;
      matchReasons.push(`attendee match: ${hit}`);
    }
  }

  // Only apply recency as a tiebreaker when at least one real criterion hit.
  if (matchReasons.length > 0) {
    score += recencyScore(meeting, now);
  }

  return { score, matchReasons };
}

export function rankMeetings(
  meetings: OnlineMeeting[],
  criteria: MeetingSearchCriteria,
  now: Date,
  limit: number
): ScoredMeeting[] {
  const hasTextCriteria = Boolean(
    criteria.query || criteria.subjectContains || criteria.participantNameOrEmail
  );

  const scored: ScoredMeeting[] = meetings.map((m) => {
    const { score, matchReasons } = scoreMeeting(m, criteria, now);
    return { ...m, score, matchReasons };
  });

  const filtered = hasTextCriteria
    ? scored.filter((m) => m.score > 0)
    : scored.map((m) => ({
        ...m,
        // No criteria: pure recency sort, populate a minimal score so ordering is stable.
        score: recencyScore(m, now),
        matchReasons: [],
      }));

  filtered.sort((a, b) => b.score - a.score);
  return filtered.slice(0, limit);
}
