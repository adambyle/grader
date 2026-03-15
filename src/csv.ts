import type { Submission, Project } from "./types";

// Minimal CSV parser that handles quoted fields (including embedded commas/newlines)
export function parseCSV(text: string): string[][] {
  // Strip BOM
  if (text.charCodeAt(0) === 0xfeff) text = text.slice(1);

  const rows: string[][] = [];
  let i = 0;

  while (i < text.length) {
    const row: string[] = [];

    while (i < text.length) {
      if (text[i] === '"') {
        // Quoted field
        i++; // skip opening quote
        let field = "";
        while (i < text.length) {
          if (text[i] === '"' && text[i + 1] === '"') {
            field += '"';
            i += 2;
          } else if (text[i] === '"') {
            i++; // skip closing quote
            break;
          } else {
            field += text[i++];
          }
        }
        row.push(field);
      } else {
        // Unquoted field
        let field = "";
        while (
          i < text.length &&
          text[i] !== "," &&
          text[i] !== "\n" &&
          text[i] !== "\r"
        ) {
          field += text[i++];
        }
        row.push(field.trim());
      }

      if (i < text.length && text[i] === ",") {
        i++; // skip comma, continue row
      } else {
        break; // end of row
      }
    }

    // Skip \r\n or \n
    if (i < text.length && text[i] === "\r") i++;
    if (i < text.length && text[i] === "\n") i++;

    if (row.length > 1 || (row.length === 1 && row[0] !== "")) {
      rows.push(row);
    }
  }

  return rows;
}

// Try to parse a date from the Moodle "Last modified (submission)" field.
// Returns undefined if the field is empty, '-', or unparseable.
export function parseSubmissionDate(raw: string): Date | undefined {
  if (!raw || raw === "-" || raw.trim() === "") return undefined;
  const d = new Date(raw);
  if (!isNaN(d.getTime())) return d;
  return undefined;
}

// Detect whether the status field indicates a missing submission.
// Moodle puts "No submission" in the status; we also treat missing date as a signal
// but only when there's no grade yet.
export function detectMissing(
  statusRaw: string,
  submissionDate?: Date,
): boolean {
  if (submissionDate) return false;
  return /no submission/i.test(statusRaw);
}

const EXPECTED_HEADERS = [
  "Identifier",
  "Full name",
  "ID number",
  "Email address",
  "Status",
  "Grade",
  "Maximum Grade",
  "Grade can be changed",
  "Last modified (submission)",
  "Online text",
  "Last modified (grade)",
  "Feedback comments",
];

export function importCSV(text: string): {
  submissions: Submission[];
  assignmentName: string;
} {
  const rows = parseCSV(text);
  if (rows.length < 2) throw new Error("CSV has no data rows.");

  const headers = rows[0];

  // Build a column index map (case-insensitive)
  const col = (name: string): number => {
    const idx = headers.findIndex(
      (h) => h.trim().toLowerCase() === name.toLowerCase(),
    );
    return idx;
  };

  const colFullName = col("Full name");
  const colIdNumber = col("ID number");
  const colEmail = col("Email address");
  const colStatus = col("Status");
  const colMaxGrade = col("Maximum Grade");
  const colLastModSub = col("Last modified (submission)");
  const colIdentifier = col("Identifier");

  if (colFullName === -1 || colEmail === -1) {
    throw new Error(
      "CSV is missing required columns (Full name, Email address).",
    );
  }

  const submissions: Submission[] = [];

  for (let r = 1; r < rows.length; r++) {
    const row = rows[r];
    if (row.length < 4) continue;

    const rawLastMod = colLastModSub >= 0 ? (row[colLastModSub] ?? "") : "";
    const submissionDate = parseSubmissionDate(rawLastMod);
    const statusRaw = colStatus >= 0 ? (row[colStatus] ?? "") : "";
    const isMissing = detectMissing(statusRaw, submissionDate);

    const maxGrade = colMaxGrade >= 0 ? parseFloat(row[colMaxGrade]) || 20 : 20;

    submissions.push({
      identifier: colIdentifier >= 0 ? row[colIdentifier] : "",
      fullName: row[colFullName] ?? "",
      idNumber: colIdNumber >= 0 ? (row[colIdNumber] ?? "") : "",
      email: row[colEmail] ?? "",
      maxGrade,
      submissionDate,
      appliedFeedback: [],
      adHocFeedback: [],
      isMissing,
    });
  }

  return { submissions, assignmentName: "" };
}

// Generate the human-readable late penalty label for a submission
export function lateFeedbackLabel(sub: Submission, project: Project): string {
  const lp = project.latePolicy;
  const days = sub.daysLate ?? 0;
  const dayStr = `${days} day${days !== 1 ? "s" : ""} late`;
  if (lp.type === "percent-per-day") {
    const pct = parseFloat((lp.amount * days).toFixed(1));
    return `${dayStr} (-${pct}%)`;
  } else if (lp.type === "flat-per-day") {
    const pts = parseFloat((lp.amount * days).toFixed(1));
    return `${dayStr} (-${pts} pts)`;
  } else if (lp.type === "zero-if-late") {
    return "Late submission";
  }
  return dayStr;
}

// Compute late penalty points for a submission (negative number or 0)
export function computeLatePenalty(sub: Submission, project: Project): number {
  const lp = project.latePolicy;
  if (lp.type === "none" || sub.latePenaltyWaived) return 0;
  const days = sub.daysLate ?? 0;
  if (days <= 0) return 0;
  let penalty = 0;
  if (lp.type === "percent-per-day") {
    penalty = -(sub.maxGrade * (lp.amount / 100) * days);
  } else if (lp.type === "flat-per-day") {
    penalty = -(lp.amount * days);
  } else if (lp.type === "zero-if-late") {
    penalty = -sub.maxGrade; // will be floored to 0 by computeGrade
  }
  if (lp.maxPenalty > 0) penalty = Math.max(penalty, -lp.maxPenalty);
  return penalty;
}

// Compute the effective grade for a submission
export function computeGrade(sub: Submission, project: Project): number {
  let grade = sub.maxGrade;

  if (sub.manualGradeOverride !== undefined) {
    grade = sub.manualGradeOverride;
  } else {
    for (const af of sub.appliedFeedback) {
      const item = project.feedbackItems.find((f) => f.id === af.itemId);
      const pts =
        af.pointsOverride !== undefined
          ? af.pointsOverride
          : (item?.points ?? 0);
      grade += pts;
    }
    for (const ah of sub.adHocFeedback) {
      if (ah.label.trim()) grade += ah.points; // only count entries with content
    }
    // Apply auto-text for missing — but not if a late penalty is already handling the grade,
    // and not if the submission has been explicitly marked perfect
    const hasLatePenalty = computeLatePenalty(sub, project) !== 0;
    if (
      sub.isMissing &&
      !sub.markedPerfect &&
      sub.appliedFeedback.length === 0 &&
      sub.adHocFeedback.length === 0 &&
      !hasLatePenalty
    ) {
      grade += project.autoTexts.missingPoints;
    }
    // Late penalty
    grade += computeLatePenalty(sub, project);
  }

  if (project.capAtMax) grade = Math.min(grade, sub.maxGrade);
  return Math.max(0, grade);
}

// Format a number to exactly 1 decimal place for CSV output
function fmt1(n: number): string {
  return parseFloat(n.toFixed(1)).toFixed(1);
}

// Build the feedback comment string for a submission (matching output format)
export function buildFeedbackComment(
  sub: Submission,
  project: Project,
): string {
  const parts: string[] = [];

  const grade = computeGrade(sub, project);

  // Auto-missing: if missing, no feedback applied, and no late penalty taking over
  const hasLatePenalty = computeLatePenalty(sub, project) !== 0;
  if (
    sub.isMissing &&
    !sub.markedPerfect &&
    sub.appliedFeedback.length === 0 &&
    sub.adHocFeedback.length === 0 &&
    !hasLatePenalty
  ) {
    const delta = grade - sub.maxGrade;
    parts.push(
      delta === 0
        ? project.autoTexts.missing
        : `${fmt1(delta)}: ${project.autoTexts.missing}`,
    );
  } else {
    for (const af of sub.appliedFeedback) {
      const item = project.feedbackItems.find((f) => f.id === af.itemId);
      if (!item) continue;
      const pts =
        af.pointsOverride !== undefined ? af.pointsOverride : item.points;
      const label =
        af.labelOverride !== undefined ? af.labelOverride : item.label;
      parts.push(pts === 0 ? label : `${fmt1(pts)}: ${label}`);
    }
    for (const ah of sub.adHocFeedback) {
      if (!ah.label.trim()) continue;
      parts.push(
        ah.points === 0 ? ah.label : `${fmt1(ah.points)}: ${ah.label}`,
      );
    }
    // Late penalty entry
    const latePts = computeLatePenalty(sub, project);
    if (latePts !== 0 && !sub.latePenaltyWaived) {
      parts.push(`${fmt1(latePts)}: ${lateFeedbackLabel(sub, project)}`);
    }
    // Perfect auto-text: only when explicitly marked
    if (sub.markedPerfect && parts.length === 0) {
      parts.push(project.autoTexts.perfect);
    }
  }

  if (parts.length === 0) return "";
  return parts.join("; ");
}

export function exportCSV(project: Project): string {
  const headers = EXPECTED_HEADERS;
  const rows: string[][] = [headers];

  for (const sub of project.submissions) {
    const grade = computeGrade(sub, project);
    const feedback = buildFeedbackComment(sub, project);
    const now = new Date().toLocaleString("en-US", {
      weekday: "long",
      year: "numeric",
      month: "long",
      day: "numeric",
      hour: "numeric",
      minute: "2-digit",
    });

    rows.push([
      sub.identifier,
      sub.fullName,
      sub.idNumber,
      sub.email,
      "No submission - Graded -  - ",
      fmt1(grade),
      sub.maxGrade.toString(),
      "Yes",
      sub.submissionDate
        ? sub.submissionDate.toLocaleString("en-US", {
            weekday: "long",
            year: "numeric",
            month: "long",
            day: "numeric",
            hour: "numeric",
            minute: "2-digit",
          })
        : "-",
      "",
      now,
      feedback,
    ]);
  }

  return rows
    .map((row) =>
      row
        .map((cell) => {
          if (
            cell.includes(",") ||
            cell.includes('"') ||
            cell.includes("\n") ||
            cell.includes("'")
          ) {
            return '"' + cell.replace(/"/g, '""') + '"';
          }
          return cell;
        })
        .join(","),
    )
    .join("\r\n");
}
