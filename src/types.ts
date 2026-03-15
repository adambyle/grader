export interface FeedbackItem {
  id: string;
  label: string;
  points: number; // negative = deduction, positive = bonus, 0 = neutral
}

export interface AppliedFeedback {
  itemId: string;
  labelOverride?: string;
  pointsOverride?: number;
}

export interface AdHocEntry {
  id: string;
  label: string;
  points: number;
}

export type LatePolicyType =
  | "none"
  | "percent-per-day"
  | "flat-per-day"
  | "zero-if-late";

export interface LatePolicy {
  type: LatePolicyType;
  amount: number; // % or flat points per day (ignored for zero-if-late/none)
  maxPenalty: number; // cap on total deduction (0 = no cap)
  deadline: string; // ISO date string, e.g. "2025-03-10"
}

export interface Submission {
  // From CSV (preserved for export)
  identifier: string;
  fullName: string;
  idNumber: string;
  email: string;
  maxGrade: number;
  submissionDate?: Date; // parsed from "Last modified (submission)" if present

  // Grading state
  appliedFeedback: AppliedFeedback[];
  adHocFeedback: AdHocEntry[];
  manualGradeOverride?: number; // if set, bypasses computed grade
  isMissing: boolean;
  daysLate?: number; // manually entered or computed from submissionDate
  latePenaltyWaived?: boolean; // grader can waive the late penalty
  markedPerfect?: boolean; // set by the Perfect button, cleared when any feedback is applied
}

export interface AutoTexts {
  missing: string;
  missingPoints: number;
  perfect: string;
}

export interface Project {
  assignmentName: string;
  feedbackItems: FeedbackItem[];
  submissions: Submission[];
  autoTexts: AutoTexts;
  capAtMax: boolean;
  latePolicy: LatePolicy;
}

export type SortKey = "name" | "email";
export type SortDir = "asc" | "desc";

export interface AppState {
  project: Project | null;
  selectedSubmissionId: string | null; // single-select (email)
  selectedSubmissionIds: Set<string>; // multi-select (emails)
  lastClickedId: string | null; // for shift-click range
  sortKey: SortKey;
  sortDir: SortDir;
  dirty: boolean;
  fileHandle: FileSystemFileHandle | null;
}
