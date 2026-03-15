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
  isMissing: boolean; // no submission date and status indicates missing
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
}

export type SortKey = 'name' | 'email';
export type SortDir = 'asc' | 'desc';

export interface AppState {
  project: Project | null;
  selectedSubmissionId: string | null; // email as id
  highlightedItemId: string | null;
  sortKey: SortKey;
  sortDir: SortDir;
  dirty: boolean;
  fileHandle: FileSystemFileHandle | null;
}
