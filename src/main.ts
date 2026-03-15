import "./style.css";
import type {
  AppState,
  FeedbackItem,
  Project,
  Submission,
  AdHocEntry,
} from "./types";
import { importCSV, exportCSV, computeGrade } from "./csv";
import {
  saveToLocalStorage,
  loadFromLocalStorage,
  saveToFile,
  loadFromFile,
  downloadBlob,
  projectToJSON,
} from "./persistence";
import { uid } from "./uid";

// ─── Dark mode ───────────────────────────────────────────────────────────────

function initDarkMode() {
  const stored = localStorage.getItem("grader_theme");
  const prefersDark = window.matchMedia("(prefers-color-scheme: dark)").matches;
  const dark = stored ? stored === "dark" : prefersDark;
  document.documentElement.classList.toggle("dark", dark);
}

function toggleDarkMode() {
  const isDark = document.documentElement.classList.toggle("dark");
  localStorage.setItem("grader_theme", isDark ? "dark" : "light");
  const icon = document.getElementById("theme-icon");
  if (icon) icon.textContent = isDark ? "☀" : "☾";
}

initDarkMode();

// ─── State ───────────────────────────────────────────────────────────────────

const state: AppState = {
  project: null,
  selectedSubmissionId: null,
  highlightedItemId: null,
  sortKey: "name",
  sortDir: "asc",
  dirty: false,
  fileHandle: null,
};

function defaultProject(): Project {
  return {
    assignmentName: "",
    feedbackItems: [],
    submissions: [],
    autoTexts: {
      missing: "Missing submission",
      missingPoints: -20,
      perfect: "Great work!",
    },
    capAtMax: true,
  };
}

// ─── Dirty / autosave ────────────────────────────────────────────────────────

let autosaveTimer: ReturnType<typeof setTimeout> | null = null;

function markDirty() {
  state.dirty = true;
  updateDirtyIndicator();
  if (autosaveTimer) clearTimeout(autosaveTimer);
  autosaveTimer = setTimeout(async () => {
    if (!state.project) return;
    saveToLocalStorage(state.project);
    if (state.fileHandle) {
      try {
        await saveToFile(state.project, state);
        state.dirty = false;
        updateDirtyIndicator();
      } catch {
        /* ignore */
      }
    }
  }, 800);
}

function updateDirtyIndicator() {
  const el = document.getElementById("dirty-indicator");
  if (el) el.textContent = state.dirty ? "●" : "";
}

window.addEventListener("beforeunload", (e) => {
  if (state.dirty) {
    e.preventDefault();
    e.returnValue = "";
  }
});

// ─── Render orchestration ────────────────────────────────────────────────────

function render() {
  renderSidebar();
  renderTable();
  renderDetail();
}

// Auto-resize all textareas to fit their content (no scrollbars, wraps naturally)
function autoResizeTextareas(root: HTMLElement = document.body) {
  root
    .querySelectorAll<HTMLTextAreaElement>(
      "textarea.fi-label, textarea.dfi-lbl",
    )
    .forEach((ta) => {
      ta.style.height = "auto";
      ta.style.height = ta.scrollHeight + "px";
    });
}

// ─── Sidebar ─────────────────────────────────────────────────────────────────

function renderSidebar() {
  const el = document.getElementById("sidebar")!;
  const p = state.project;

  el.innerHTML = `
    <div class="sidebar-top">
      <div class="sidebar-title-row">
        <div class="app-title">Grader</div>
        <button id="btn-theme" class="btn-icon" title="Toggle dark mode"><span id="theme-icon">${document.documentElement.classList.contains("dark") ? "☀" : "☾"}</span></button>
      </div>
      <div class="project-controls">
        <button id="btn-new-csv">Import CSV</button>
        <button id="btn-load-json">Load</button>
        <button id="btn-save-json" ${!p ? "disabled" : ""}>Save<span id="dirty-indicator" class="dirty"></span></button>
        <button id="btn-export-csv" ${!p ? "disabled" : ""}>Export CSV</button>
      </div>
      ${
        p
          ? `
      <div class="meta-row">
        <input id="assignment-name" type="text" placeholder="Assignment name" value="${esc(p.assignmentName)}" />
      </div>
      <div class="meta-row">
        <label class="inline-check"><input type="checkbox" id="cap-at-max" ${p.capAtMax ? "checked" : ""} /> Cap score at max grade</label>
      </div>
      `
          : ""
      }
    </div>

    ${
      p
        ? `
    <div class="sidebar-section">
      <div class="section-header">Auto-text</div>
      <div class="auto-text-block">
        <div class="at-row">
          <span class="at-label">Missing</span>
          <input type="number" id="at-missing-pts" class="pts-input" value="${p.autoTexts.missingPoints}" step="0.5" />
          <input type="text" id="at-missing-label" class="flex-input" value="${esc(p.autoTexts.missing)}" />
        </div>
        <div class="at-row">
          <span class="at-label">Perfect</span>
          <input type="text" id="at-perfect-label" class="flex-input" value="${esc(p.autoTexts.perfect)}" />
        </div>
      </div>
    </div>

    <div class="sidebar-section feedback-section">
      <div class="section-header">
        Reusable feedback
        <button id="btn-add-item" class="btn-icon" title="Add feedback item">+</button>
      </div>
      <div class="feedback-hint">Define items here, then apply them per submission →</div>
      <ul class="feedback-list" id="feedback-list">
        ${p.feedbackItems.map((item) => renderFeedbackItem(item)).join("")}
      </ul>
    </div>
    `
        : `
    <div class="empty-state">
      <p>Import a Moodle CSV or load a saved project.</p>
    </div>
    `
    }
  `;

  // localStorage recovery notice
  if (!p) {
    const saved = loadFromLocalStorage();
    if (saved) {
      const notice = document.createElement("div");
      notice.className = "ls-notice";
      notice.innerHTML = `Unsaved session found. <button id="btn-recover">Restore</button>`;
      el.appendChild(notice);
      document.getElementById("btn-recover")?.addEventListener("click", () => {
        if (
          state.dirty &&
          !confirm("Load recovered session? Current work will be lost.")
        )
          return;
        state.project = saved;
        state.dirty = false;
        render();
      });
    }
  }

  bindSidebarEvents();
  autoResizeTextareas(el);
}

function renderFeedbackItem(item: FeedbackItem): string {
  const isHighlighted = state.highlightedItemId === item.id;
  const affectedCount =
    state.project?.submissions.filter((s) =>
      s.appliedFeedback.some((af) => af.itemId === item.id),
    ).length ?? 0;
  return `
    <li class="feedback-item ${isHighlighted ? "fi-highlighted" : ""}" data-id="${item.id}">
      <input type="number" class="pts-input fi-pts" value="${item.points}" step="0.5" data-id="${item.id}" />
      <textarea class="fi-label" data-id="${item.id}" placeholder="Feedback text" rows="1">${esc(item.label)}</textarea>
      <button class="fi-count-btn ${affectedCount > 0 ? "has-count" : ""}" data-id="${item.id}" title="Highlight affected submissions">${affectedCount || ""}</button>
      <button class="btn-icon fi-delete" data-id="${item.id}">×</button>
    </li>
  `;
}

function bindSidebarEvents() {
  const p = state.project;

  document
    .getElementById("btn-theme")
    ?.addEventListener("click", toggleDarkMode);
  document
    .getElementById("btn-new-csv")
    ?.addEventListener("click", handleImportCSV);
  document
    .getElementById("btn-load-json")
    ?.addEventListener("click", handleLoadJSON);
  document
    .getElementById("btn-save-json")
    ?.addEventListener("click", handleSaveJSON);
  document
    .getElementById("btn-export-csv")
    ?.addEventListener("click", handleExportCSV);

  if (!p) return;

  (
    document.getElementById("assignment-name") as HTMLInputElement
  )?.addEventListener("input", (e) => {
    p.assignmentName = (e.target as HTMLInputElement).value;
    markDirty();
  });

  (document.getElementById("cap-at-max") as HTMLInputElement)?.addEventListener(
    "change",
    (e) => {
      p.capAtMax = (e.target as HTMLInputElement).checked;
      markDirty();
      renderTable();
      renderDetail();
    },
  );

  (
    document.getElementById("at-missing-pts") as HTMLInputElement
  )?.addEventListener("input", (e) => {
    p.autoTexts.missingPoints =
      parseFloat((e.target as HTMLInputElement).value) || 0;
    markDirty();
    renderTable();
  });
  (
    document.getElementById("at-missing-label") as HTMLInputElement
  )?.addEventListener("input", (e) => {
    p.autoTexts.missing = (e.target as HTMLInputElement).value;
    markDirty();
  });
  (
    document.getElementById("at-perfect-label") as HTMLInputElement
  )?.addEventListener("input", (e) => {
    p.autoTexts.perfect = (e.target as HTMLInputElement).value;
    markDirty();
  });

  document.getElementById("btn-add-item")?.addEventListener("click", () => {
    const item: FeedbackItem = { id: uid(), label: "", points: -1 };
    p.feedbackItems.push(item);
    markDirty();
    renderSidebar();
    renderDetail(); // new item should appear in detail list immediately
    const labels =
      document.querySelectorAll<HTMLTextAreaElement>("textarea.fi-label");
    labels[labels.length - 1]?.focus();
  });

  const list = document.getElementById("feedback-list");

  list?.addEventListener("keydown", (e) => {
    const target = e.target as HTMLTextAreaElement | HTMLInputElement;
    if (e.key !== "Enter") return;
    e.preventDefault(); // prevent newline in textarea
    const id = (target as HTMLElement).dataset.id;
    if (!id) return;
    if (target.classList.contains("fi-label") && target.value.trim()) {
      const newItem: FeedbackItem = { id: uid(), label: "", points: -1 };
      p.feedbackItems.push(newItem);
      markDirty();
      renderSidebar();
      renderDetail();
      const labels =
        document.querySelectorAll<HTMLTextAreaElement>("textarea.fi-label");
      labels[labels.length - 1]?.focus();
    } else if (target.classList.contains("fi-pts")) {
      const labelTA = list.querySelector<HTMLTextAreaElement>(
        `textarea.fi-label[data-id="${id}"]`,
      );
      labelTA?.focus();
    }
  });

  // Blur on an empty label: silently remove the item
  list?.addEventListener("focusout", (e) => {
    const target = e.target as HTMLInputElement;
    if (!target.classList.contains("fi-label")) return;
    const id = target.dataset.id;
    if (!id) return;
    const item = p.feedbackItems.find((f) => f.id === id);
    if (!item || item.label.trim()) return; // only delete if still empty
    // Don't delete if focus is moving to another element within the same item
    const related = (e as FocusEvent).relatedTarget as HTMLElement | null;
    if (related?.dataset.id === id) return;
    p.feedbackItems = p.feedbackItems.filter((f) => f.id !== id);
    markDirty();
    renderSidebar();
    renderDetail();
  });

  list?.addEventListener("input", (e) => {
    const target = e.target as HTMLInputElement | HTMLTextAreaElement;
    if (target instanceof HTMLTextAreaElement) {
      target.style.height = "auto";
      target.style.height = target.scrollHeight + "px";
    }
    const id = (target as HTMLElement).dataset.id;
    if (!id) return;
    const item = p.feedbackItems.find((f) => f.id === id);
    if (!item) return;
    if (target.classList.contains("fi-pts")) {
      item.points = parseFloat((target as HTMLInputElement).value) || 0;
    } else if (target.classList.contains("fi-label")) {
      item.label = (target as HTMLTextAreaElement).value;
    }
    markDirty();
    updateFeedbackCountBadge(id);
    highlightTableRows();
    if (state.selectedSubmissionId) {
      const sel = p.submissions.find(
        (s) => s.email === state.selectedSubmissionId,
      );
      if (sel?.appliedFeedback.some((af) => af.itemId === id)) renderDetail();
    }
  });

  list?.addEventListener("click", (e) => {
    const target = e.target as HTMLElement;
    const id = target.dataset.id;
    if (!id) return;

    if (target.classList.contains("fi-delete")) {
      const item = p.feedbackItems.find((f) => f.id === id);
      const isApplied = p.submissions.some((s) =>
        s.appliedFeedback.some((af) => af.itemId === id),
      );
      const isEmpty = !item?.label.trim();
      if (
        !isEmpty &&
        isApplied &&
        !confirm(
          "Delete this feedback item? It will be removed from all submissions.",
        )
      )
        return;
      p.feedbackItems = p.feedbackItems.filter((f) => f.id !== id);
      for (const sub of p.submissions) {
        sub.appliedFeedback = sub.appliedFeedback.filter(
          (af) => af.itemId !== id,
        );
      }
      if (state.highlightedItemId === id) state.highlightedItemId = null;
      markDirty();
      renderSidebar();
      renderTable();
      renderDetail();
    } else if (target.classList.contains("fi-count-btn")) {
      state.highlightedItemId = state.highlightedItemId === id ? null : id;
      renderSidebar();
      highlightTableRows();
    }
  });
}

function updateFeedbackCountBadge(itemId: string) {
  const count =
    state.project?.submissions.filter((s) =>
      s.appliedFeedback.some((af) => af.itemId === itemId),
    ).length ?? 0;
  const btn = document.querySelector<HTMLElement>(
    `.fi-count-btn[data-id="${itemId}"]`,
  );
  if (btn) {
    btn.textContent = count > 0 ? String(count) : "";
    btn.classList.toggle("has-count", count > 0);
  }
}

// ─── File operations ──────────────────────────────────────────────────────────

async function handleImportCSV() {
  if (state.project && state.dirty) {
    if (!confirm("Import a new CSV? Unsaved changes will be lost.")) return;
  }
  const input = document.createElement("input");
  input.type = "file";
  input.accept = ".csv";
  input.onchange = async () => {
    const file = input.files?.[0];
    if (!file) return;
    const text = await file.text();
    try {
      const { submissions } = importCSV(text);
      const project = defaultProject();
      project.submissions = submissions;
      project.assignmentName = file.name.replace(/\.csv$/i, "");
      state.project = project;
      state.dirty = true;
      state.fileHandle = null;
      state.selectedSubmissionId = null;
      state.highlightedItemId = null;
      saveToLocalStorage(project);
      render();
    } catch (err: any) {
      alert("Failed to import CSV: " + err.message);
    }
  };
  input.click();
}

async function handleLoadJSON() {
  if (state.project && state.dirty) {
    if (!confirm("Load a project? Unsaved changes will be lost.")) return;
  }
  try {
    const result = await loadFromFile();
    if (!result) return;
    state.project = result.project;
    state.fileHandle = result.handle;
    state.dirty = false;
    state.selectedSubmissionId = null;
    state.highlightedItemId = null;
    render();
  } catch (err: any) {
    alert("Failed to load: " + err.message);
  }
}

async function handleSaveJSON() {
  if (!state.project) return;
  try {
    const handle = await saveToFile(state.project, state);
    if (handle) {
      state.fileHandle = handle;
      state.dirty = false;
      updateDirtyIndicator();
    }
  } catch {
    downloadBlob(
      `${state.project.assignmentName || "grading"}.json`,
      projectToJSON(state.project),
      "application/json",
    );
  }
}

function handleExportCSV() {
  if (!state.project) return;
  const csv = exportCSV(state.project);
  downloadBlob(`${state.project.assignmentName || "grades"}_output.csv`, csv);
}

// ─── Table ────────────────────────────────────────────────────────────────────

function sortedSubmissions(): Submission[] {
  if (!state.project) return [];
  return [...state.project.submissions].sort((a, b) => {
    const av = state.sortKey === "name" ? a.fullName : a.email;
    const bv = state.sortKey === "name" ? b.fullName : b.email;
    const cmp = av.localeCompare(bv);
    return state.sortDir === "asc" ? cmp : -cmp;
  });
}

function renderTable() {
  const el = document.getElementById("table-container")!;
  const p = state.project;
  if (!p) {
    el.innerHTML = "";
    return;
  }

  const subs = sortedSubmissions();

  const arrow = (key: string) =>
    state.sortKey === key
      ? `<span class="sarrow">${state.sortDir === "asc" ? "↑" : "↓"}</span>`
      : `<span class="sarrow dim">↕</span>`;

  el.innerHTML = `
    <table class="sub-table">
      <thead>
        <tr>
          <th class="sortable col-name" data-sort="name">Name ${arrow("name")}</th>
          <th class="sortable col-email" data-sort="email">Email ${arrow("email")}</th>
          <th class="col-grade">Grade</th>
          <th class="col-summary">Feedback</th>
          <th class="col-status">Done</th>
        </tr>
      </thead>
      <tbody>
        ${subs.map((sub) => renderRow(sub, p)).join("")}
      </tbody>
    </table>
  `;

  bindTableEvents();
  highlightTableRows();
}

function isGraded(sub: Submission): boolean {
  return (
    sub.appliedFeedback.length > 0 ||
    sub.adHocFeedback.length > 0 ||
    sub.manualGradeOverride !== undefined
  );
}

function rowStatus(
  sub: Submission,
  _p: Project,
): { cls: string; label: string } {
  if (isGraded(sub)) return { cls: "st-done", label: "✓" };
  return { cls: "st-todo", label: "·" };
}

function renderRow(sub: Submission, p: Project): string {
  const grade = computeGrade(sub, p);
  const isSelected = sub.email === state.selectedSubmissionId;
  const { cls, label } = rowStatus(sub, p);
  const graded = isGraded(sub);
  const summary = graded ? buildFeedbackSummary(sub, p) : "";

  return `<tr class="sub-row ${isSelected ? "row-selected" : ""} ${!graded ? "row-ungraded" : ""}"
      data-email="${esc(sub.email)}">
    <td class="cell-name">${esc(sub.fullName)}</td>
    <td class="cell-email">${esc(sub.email)}</td>
    <td class="cell-grade"><span class="grade-num">${fmtGrade(grade)}</span><span class="grade-max"> / ${sub.maxGrade}</span></td>
    <td class="cell-summary"><span class="summary-text">${esc(summary)}</span></td>
    <td class="cell-status"><span class="st-badge ${cls}">${label}</span></td>
  </tr>`;
}

// One-line plain-text summary of feedback for the table cell
function buildFeedbackSummary(sub: Submission, p: Project): string {
  const parts: string[] = [];
  for (const af of sub.appliedFeedback) {
    const item = p.feedbackItems.find((f) => f.id === af.itemId);
    if (!item) continue;
    const label =
      af.labelOverride !== undefined ? af.labelOverride : item.label;
    parts.push(label);
  }
  for (const ah of sub.adHocFeedback) {
    if (ah.label) parts.push(ah.label);
  }
  return parts.join("; ");
}

function fmtGrade(g: number): string {
  // Always one decimal to prevent column width shifting
  return Number.isInteger(g) ? g.toFixed(1) : g.toFixed(1);
}

function bindTableEvents() {
  document.querySelectorAll<HTMLElement>(".sortable").forEach((th) => {
    th.addEventListener("click", () => {
      const key = th.dataset.sort as "name" | "email";
      if (state.sortKey === key)
        state.sortDir = state.sortDir === "asc" ? "desc" : "asc";
      else {
        state.sortKey = key;
        state.sortDir = "asc";
      }
      renderTable();
    });
  });

  document.querySelectorAll<HTMLElement>(".sub-row").forEach((row) => {
    row.addEventListener("click", () => {
      const email = row.dataset.email!;
      if (state.selectedSubmissionId === email) {
        state.selectedSubmissionId = null;
      } else {
        state.selectedSubmissionId = email;
      }
      document
        .querySelectorAll(".sub-row")
        .forEach((r) => r.classList.remove("row-selected"));
      if (state.selectedSubmissionId) row.classList.add("row-selected");
      renderDetail();
    });
  });
}

function highlightTableRows() {
  const hid = state.highlightedItemId;
  document.querySelectorAll<HTMLElement>(".sub-row").forEach((row) => {
    const sub = state.project?.submissions.find(
      (s) => s.email === row.dataset.email,
    );
    const affected =
      !!hid && !!sub?.appliedFeedback.some((af) => af.itemId === hid);
    row.classList.toggle("row-highlight", affected);
  });
}

// ─── Detail panel ─────────────────────────────────────────────────────────────

function renderDetail() {
  const el = document.getElementById("detail-panel")!;
  const p = state.project;
  if (!p || !state.selectedSubmissionId) {
    el.innerHTML = `<div class="detail-empty">Select a submission to grade it.</div>`;
    return;
  }
  const sub = p.submissions.find((s) => s.email === state.selectedSubmissionId);
  if (!sub) {
    el.innerHTML = "";
    return;
  }

  const grade = computeGrade(sub, p);
  const isOverride = sub.manualGradeOverride !== undefined;

  el.innerHTML = `
    <div class="detail-hdr">
      <div>
        <div class="detail-name">${esc(sub.fullName)}</div>
        <div class="detail-email">${esc(sub.email)}</div>
      </div>
      <button class="btn-icon close-detail" title="Close">×</button>
    </div>

    <div class="grade-row">
      <div class="grade-label">Grade</div>
      <div class="grade-controls">
        <span class="grade-big ${isOverride ? "grade-overridden" : ""}" id="grade-display">${fmtGrade(grade)}</span>
        <span class="grade-max-detail"> / ${sub.maxGrade}</span>
        <input type="number" id="manual-grade" class="pts-input manual-grade-input"
          value="${isOverride ? sub.manualGradeOverride : fmtGrade(grade)}"
          min="0" step="0.5" ${!isOverride ? "disabled" : ""} />
        <button class="btn-icon grade-lock-btn" id="btn-grade-lock"
          title="${isOverride ? "Revert to computed grade" : "Set grade manually"}"
        >${isOverride ? "🔓" : "🔒"}</button>
        <button class="btn-text" id="btn-perfect" title="Assign full marks with perfect auto-text">✓ Perfect</button>
      </div>
      ${isOverride ? `<div class="override-note">Manual override active</div>` : ""}
    </div>

    <div class="detail-section">
      <div class="detail-sub">
        Applied feedback
        <span class="detail-sub-hint">— check to apply; values can be edited per submission</span>
      </div>
      <ul class="detail-fi-list" id="detail-fi-list">
        ${(() => {
          const labelled = p.feedbackItems.filter((item) => item.label.trim());
          if (labelled.length === 0)
            return `<li class="detail-empty-hint">No reusable feedback defined yet. Add items in the left panel.</li>`;
          return labelled.map((item) => renderDetailFI(item, sub)).join("");
        })()}
      </ul>
    </div>

    <div class="detail-section">
      <div class="detail-sub adhoc-hdr">
        One-off notes
        <button class="btn-icon" id="btn-add-adhoc" title="Add a custom note for this submission">+</button>
      </div>
      <ul class="detail-adhoc-list" id="detail-adhoc-list">
        ${sub.adHocFeedback.map((ah) => renderAdHoc(ah)).join("")}
      </ul>
    </div>

    <div class="detail-section feedback-summary-section">
      <div class="detail-sub">Feedback summary</div>
      ${renderFeedbackSummaryList(sub, p)}
    </div>
  `;

  bindDetailEvents(sub, p);
  autoResizeTextareas(el);
}

function renderDetailFI(item: FeedbackItem, sub: Submission): string {
  const af = sub.appliedFeedback.find((a) => a.itemId === item.id);
  const applied = !!af;
  const pts =
    af?.pointsOverride !== undefined ? af.pointsOverride : item.points;
  const label = af?.labelOverride !== undefined ? af.labelOverride : item.label;
  const hasOverride = !!(
    af?.labelOverride !== undefined || af?.pointsOverride !== undefined
  );

  return `<li class="dfi ${applied ? "dfi-applied" : ""}" data-id="${item.id}">
    <input type="checkbox" class="dfi-check" data-id="${item.id}" ${applied ? "checked" : ""} />
    <input type="number" class="pts-input dfi-pts" data-id="${item.id}"
      value="${pts}" step="0.5" ${!applied ? "disabled" : ""} />
    <textarea class="dfi-lbl" data-id="${item.id}"
      ${!applied ? "disabled" : ""} placeholder="text" rows="1">${esc(label)}</textarea>
    ${hasOverride ? `<button class="btn-icon dfi-revert" data-id="${item.id}" title="Revert to global">↺</button>` : ""}
  </li>`;
}

function renderAdHoc(ah: AdHocEntry): string {
  return `<li class="adhoc-item" data-ahid="${ah.id}">
    <input type="number" class="pts-input ah-pts" data-ahid="${ah.id}" value="${ah.points}" step="0.5" />
    <input type="text" class="ah-lbl" data-ahid="${ah.id}" value="${esc(ah.label)}" placeholder="Note" />
    <button class="btn-icon ah-del" data-ahid="${ah.id}">×</button>
  </li>`;
}

function bindDetailEvents(sub: Submission, p: Project) {
  document.querySelector(".close-detail")?.addEventListener("click", () => {
    state.selectedSubmissionId = null;
    document
      .querySelectorAll(".sub-row")
      .forEach((r) => r.classList.remove("row-selected"));
    renderDetail();
  });

  // Lock/unlock grade
  document.getElementById("btn-grade-lock")?.addEventListener("click", () => {
    if (sub.manualGradeOverride !== undefined) {
      delete sub.manualGradeOverride;
    } else {
      sub.manualGradeOverride = computeGrade(sub, p);
    }
    markDirty();
    renderDetail();
    updateRow(sub, p);
  });

  // Perfect score shortcut
  document.getElementById("btn-perfect")?.addEventListener("click", () => {
    // Clear any applied feedback so the perfect auto-text shows cleanly
    sub.appliedFeedback = [];
    sub.adHocFeedback = [];
    sub.manualGradeOverride = sub.maxGrade;
    markDirty();
    renderDetail();
    updateRow(sub, p);
  });

  (
    document.getElementById("manual-grade") as HTMLInputElement
  )?.addEventListener("input", (e) => {
    sub.manualGradeOverride = parseFloat((e.target as HTMLInputElement).value);
    markDirty();
    refreshGrade(sub, p);
    updateRow(sub, p);
    refreshPreview(sub, p);
  });

  // Feedback items list
  const fiList = document.getElementById("detail-fi-list");
  fiList?.addEventListener("change", (e) => {
    const t = e.target as HTMLInputElement;
    if (!t.classList.contains("dfi-check")) return;
    const id = t.dataset.id!;
    if (t.checked) {
      sub.appliedFeedback.push({ itemId: id });
    } else {
      sub.appliedFeedback = sub.appliedFeedback.filter(
        (af) => af.itemId !== id,
      );
    }
    markDirty();
    updateFeedbackCountBadge(id);
    highlightTableRows();
    renderDetail();
    updateRow(sub, p);
  });

  fiList?.addEventListener("input", (e) => {
    const t = e.target as HTMLInputElement;
    const id = t.dataset.id;
    if (!id) return;
    const af = sub.appliedFeedback.find((a) => a.itemId === id);
    if (!af) return;
    if (t.classList.contains("dfi-pts"))
      af.pointsOverride = parseFloat(t.value);
    if (t.classList.contains("dfi-lbl")) af.labelOverride = t.value;
    markDirty();
    refreshGrade(sub, p);
    updateRow(sub, p);
    refreshPreview(sub, p);
  });

  fiList?.addEventListener("click", (e) => {
    const t = e.target as HTMLElement;
    if (!t.classList.contains("dfi-revert")) return;
    const id = t.dataset.id!;
    const af = sub.appliedFeedback.find((a) => a.itemId === id);
    if (af) {
      delete af.labelOverride;
      delete af.pointsOverride;
    }
    markDirty();
    renderDetail();
    updateRow(sub, p);
  });

  // Ad-hoc
  document.getElementById("btn-add-adhoc")?.addEventListener("click", () => {
    sub.adHocFeedback.push({ id: uid(), label: "", points: -1 });
    markDirty();
    renderDetail();
    const labels = document.querySelectorAll<HTMLInputElement>(".ah-lbl");
    labels[labels.length - 1]?.focus();
  });

  const ahList = document.getElementById("detail-adhoc-list");

  ahList?.addEventListener("keydown", (e) => {
    if (e.key !== "Enter") return;
    e.preventDefault();
    const target = e.target as HTMLInputElement;
    if (target.classList.contains("ah-lbl") && target.value.trim()) {
      sub.adHocFeedback.push({ id: uid(), label: "", points: -1 });
      markDirty();
      renderDetail();
      const labels = document.querySelectorAll<HTMLInputElement>(".ah-lbl");
      labels[labels.length - 1]?.focus();
    } else if (target.classList.contains("ah-pts")) {
      const ahid = target.dataset.ahid;
      const lbl = ahList.querySelector<HTMLInputElement>(
        `.ah-lbl[data-ahid="${ahid}"]`,
      );
      lbl?.focus();
    }
  });
  ahList?.addEventListener("input", (e) => {
    const t = e.target as HTMLInputElement;
    const ahid = t.dataset.ahid;
    if (!ahid) return;
    const ah = sub.adHocFeedback.find((a) => a.id === ahid);
    if (!ah) return;
    if (t.classList.contains("ah-pts")) ah.points = parseFloat(t.value) || 0;
    if (t.classList.contains("ah-lbl")) ah.label = t.value;
    markDirty();
    refreshGrade(sub, p);
    updateRow(sub, p);
    refreshPreview(sub, p);
  });

  ahList?.addEventListener("click", (e) => {
    const t = e.target as HTMLElement;
    if (!t.classList.contains("ah-del")) return;
    const ahid = t.dataset.ahid!;
    sub.adHocFeedback = sub.adHocFeedback.filter((a) => a.id !== ahid);
    markDirty();
    renderDetail();
    updateRow(sub, p);
  });
}

function refreshGrade(sub: Submission, p: Project) {
  const grade = computeGrade(sub, p);
  const display = document.getElementById("grade-display");
  if (display) display.textContent = fmtGrade(grade);
  // If not in override mode, also sync the disabled input value so it shows current computed
  if (sub.manualGradeOverride === undefined) {
    const input = document.getElementById("manual-grade") as HTMLInputElement;
    if (input) input.value = fmtGrade(grade);
  }
}

function fmtPts(pts: number): string {
  if (pts > 0) return `+${pts % 1 === 0 ? pts.toFixed(1) : pts}`;
  if (pts === 0) return "0";
  return pts % 1 === 0 ? pts.toFixed(1) : String(pts);
}

function renderFeedbackSummaryList(sub: Submission, p: Project): string {
  const items: Array<{ pts: number; label: string; positive: boolean }> = [];

  if (
    sub.isMissing &&
    sub.appliedFeedback.length === 0 &&
    sub.adHocFeedback.length === 0 &&
    sub.manualGradeOverride === undefined
  ) {
    items.push({
      pts: p.autoTexts.missingPoints,
      label: p.autoTexts.missing,
      positive: false,
    });
  } else {
    for (const af of sub.appliedFeedback) {
      const item = p.feedbackItems.find((f) => f.id === af.itemId);
      if (!item) continue;
      const pts =
        af.pointsOverride !== undefined ? af.pointsOverride : item.points;
      const label =
        af.labelOverride !== undefined ? af.labelOverride : item.label;
      items.push({ pts, label, positive: pts > 0 });
    }
    for (const ah of sub.adHocFeedback) {
      items.push({ pts: ah.points, label: ah.label, positive: ah.points > 0 });
    }
    // Perfect auto-text: shown when grade equals max and no other feedback
    const grade = computeGrade(sub, p);
    if (items.length === 0 && grade >= sub.maxGrade) {
      items.push({ pts: 0, label: p.autoTexts.perfect, positive: false });
    }
  }

  if (items.length === 0) {
    return `<div class="summary-empty">No feedback applied yet.</div>`;
  }

  return `<ul class="feedback-summary-list">${items
    .map(
      (it) =>
        `<li class="fsl-item ${it.positive ? "fsl-positive" : it.pts < 0 ? "fsl-negative" : "fsl-neutral"}">
      <span class="fsl-pts">${fmtPts(it.pts)}</span>
      <span class="fsl-label">${esc(it.label)}</span>
    </li>`,
    )
    .join("")}</ul>`;
}

function refreshPreview(sub: Submission, p: Project) {
  const section = document.querySelector(".feedback-summary-section");
  if (!section) return;
  const existing = section.querySelector(
    ".feedback-summary-list, .summary-empty",
  );
  if (existing) existing.outerHTML = renderFeedbackSummaryList(sub, p);
  else
    section.insertAdjacentHTML("beforeend", renderFeedbackSummaryList(sub, p));
}

function updateRow(sub: Submission, p: Project) {
  const row = document.querySelector<HTMLElement>(
    `.sub-row[data-email="${sub.email}"]`,
  );
  if (!row) return;
  const grade = computeGrade(sub, p);
  const gradeCell = row.querySelector(".cell-grade");
  if (gradeCell)
    gradeCell.innerHTML = `<span class="grade-num">${fmtGrade(grade)}</span><span class="grade-max"> / ${sub.maxGrade}</span>`;
  const summaryCell = row.querySelector(".summary-text");
  if (summaryCell)
    summaryCell.textContent = isGraded(sub) ? buildFeedbackSummary(sub, p) : "";
  const { cls, label } = rowStatus(sub, p);
  const badge = row.querySelector(".st-badge");
  if (badge) {
    badge.className = `st-badge ${cls}`;
    badge.textContent = label;
  }
  row.classList.toggle("row-ungraded", !isGraded(sub));
}

// ─── Helpers ──────────────────────────────────────────────────────────────────

function esc(str: string): string {
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

// ─── Bootstrap ───────────────────────────────────────────────────────────────

document.querySelector<HTMLDivElement>("#app")!.innerHTML = `
  <div id="layout">
    <aside id="sidebar"></aside>
    <main id="main">
      <div id="table-container"></div>
    </main>
    <section id="detail-panel"></section>
  </div>
`;

render();
