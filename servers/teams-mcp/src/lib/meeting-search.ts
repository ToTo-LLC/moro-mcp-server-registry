import Fuse from "fuse.js";
import type { MeetingParticipant, OnlineMeeting } from "../types/graph.js";

export interface MeetingSearchCriteria {
  query?: string;
  subjectContains?: string;
  participantNameOrEmail?: string;
  startDateTime?: string;
  endDateTime?: string;
}

/**
 * A meeting enriched with its search score and human-readable match reasons.
 * Both `score` and `matchReasons` are included in find_meetings tool output
 * per spec — the LLM uses `matchReasons` to explain to the user why a given
 * meeting was picked. Do not strip them when serializing.
 */
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

/**
 * Build a single Fuse index over every meeting's subject and return a map
 * from meeting id to a fuse score contribution in [0, WEIGHT_SUBJECT].
 *
 * Pre-computing this once per rankMeetings() call is a perf optimization:
 * the naive approach constructs a Fuse index per meeting, which is O(N)
 * index builds for an N-candidate search and becomes measurable at the
 * 200-candidate cap used by find_meetings.
 */
function buildFuzzySubjectContributions(
  meetings: OnlineMeeting[],
  query: string
): Map<string, number> {
  const indexed = meetings.map((m, i) => ({
    refIndex: i,
    id: m.id,
    subject: m.subject ?? "",
  }));
  const fuse = new Fuse(indexed, {
    keys: ["subject"],
    includeScore: true,
    threshold: 0.4, // allows ~2 character transpositions on short strings
  });
  const hits = fuse.search(query);
  const out = new Map<string, number>();
  for (const hit of hits) {
    const fuseScore = hit.score ?? 1;
    const contribution = WEIGHT_SUBJECT * (1 - fuseScore);
    if (contribution > 0) {
      out.set(hit.item.id, contribution);
    }
  }
  return out;
}

export function scoreMeeting(
  meeting: OnlineMeeting,
  criteria: MeetingSearchCriteria,
  now: Date,
  fuzzyContributions?: Map<string, number>
): { score: number; matchReasons: string[] } {
  let score = 0;
  const matchReasons: string[] = [];

  const subject = (meeting.subject ?? "").toLowerCase();

  if (criteria.subjectContains) {
    const needle = criteria.subjectContains.toLowerCase();
    if (subject.includes(needle)) {
      score += WEIGHT_SUBJECT;
      matchReasons.push(`subject contains "${criteria.subjectContains}"`);
    }
  }

  if (criteria.query) {
    // Prefer pre-computed contributions from a batch Fuse index when available
    // (rankMeetings builds one per call). Fall back to a one-off index for
    // direct scoreMeeting() callers such as unit tests.
    let contribution: number | undefined;
    if (fuzzyContributions) {
      contribution = fuzzyContributions.get(meeting.id);
    } else {
      const fuse = new Fuse([{ subject: meeting.subject ?? "" }], {
        keys: ["subject"],
        includeScore: true,
        threshold: 0.4,
      });
      const hits = fuse.search(criteria.query);
      if (hits.length > 0) {
        const fuseScore = hits[0].score ?? 1;
        const oneOff = WEIGHT_SUBJECT * (1 - fuseScore);
        if (oneOff > 0) contribution = oneOff;
      }
    }
    if (contribution !== undefined && contribution > 0) {
      score += contribution;
      matchReasons.push(`subject fuzzy-matches "${criteria.query}"`);
    }
  }

  if (criteria.participantNameOrEmail) {
    const needle = criteria.participantNameOrEmail.toLowerCase();
    // participantsAsStrings may include both p.upn and user.userPrincipalName
    // for the same participant. Duplicates are harmless here because we only
    // use the first match, and the matchReason string reflects whichever form
    // hit first.
    const strings = participantsAsStrings(meeting).map((s) => s.toLowerCase());
    const hit = strings.find((s) => s.includes(needle));
    if (hit) {
      score += WEIGHT_PARTICIPANT;
      matchReasons.push(`attendee match: ${hit}`);
    }
  }

  // Recency is a tiebreaker only — never a sole contributor. Apply it only
  // when at least one real criterion matched.
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

  // Pre-compute the fuzzy contribution map once if a query is present.
  const fuzzyContributions = criteria.query
    ? buildFuzzySubjectContributions(meetings, criteria.query)
    : undefined;

  const scored: ScoredMeeting[] = meetings.map((m) => {
    const { score, matchReasons } = scoreMeeting(m, criteria, now, fuzzyContributions);
    return { ...m, score, matchReasons };
  });

  const filtered = hasTextCriteria
    ? scored.filter((m) => m.score > 0)
    : scored.map((m) => ({
        ...m,
        // scoreMeeting blocked recency behind matchReasons.length > 0, so every
        // score here is 0. Recompute raw recency so the no-criteria branch sorts
        // meaningfully (most recent first).
        score: recencyScore(m, now),
        matchReasons: [],
      }));

  filtered.sort((a, b) => b.score - a.score);
  return filtered.slice(0, limit);
}
