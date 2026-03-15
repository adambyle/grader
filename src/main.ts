import "./style.css";
import type {
  AppState,
  FeedbackItem,
  Project,
  Submission,
  AdHocEntry,
  LatePolicyType,
} from "./types";
import {
  importCSV,
  exportCSV,
  computeGrade,
  computeLatePenalty,
  lateFeedbackLabel,
} from "./csv";
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
  selectedSubmissionIds: new Set(),
  lastClickedId: null,
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
    latePolicy: {
      type: "none",
      amount: 10,
      maxPenalty: 0,
      deadline: "",
    },
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

// JSON round-trips serialize Dates as strings — restore them after any load
function rehydrateDates(project: Project) {
  for (const sub of project.submissions) {
    if (sub.submissionDate && !(sub.submissionDate instanceof Date)) {
      const d = new Date(sub.submissionDate as unknown as string);
      sub.submissionDate = isNaN(d.getTime()) ? undefined : d;
    }
  }
}

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

    <div class="sidebar-section">
      <div class="section-header">Late policy</div>
      <div class="late-policy-block">
        <div class="at-row">
          <span class="at-label">Type</span>
          <select id="lp-type" class="lp-select">
            <option value="none" ${p.latePolicy.type === "none" ? "selected" : ""}>None</option>
            <option value="percent-per-day" ${p.latePolicy.type === "percent-per-day" ? "selected" : ""}>% per day</option>
            <option value="flat-per-day" ${p.latePolicy.type === "flat-per-day" ? "selected" : ""}>Points per day</option>
            <option value="zero-if-late" ${p.latePolicy.type === "zero-if-late" ? "selected" : ""}>Zero if late</option>
          </select>
        </div>
        ${
          p.latePolicy.type !== "none" && p.latePolicy.type !== "zero-if-late"
            ? `
        <div class="at-row">
          <span class="at-label">Amount</span>
          <input type="number" id="lp-amount" class="pts-input" value="${p.latePolicy.amount}" min="0" step="0.5" />
          <span class="lp-unit">${p.latePolicy.type === "percent-per-day" ? "% / day" : "pts / day"}</span>
        </div>`
            : ""
        }
        ${
          p.latePolicy.type !== "none"
            ? `
        <div class="at-row">
          <span class="at-label">Max</span>
          <input type="number" id="lp-max" class="pts-input" value="${p.latePolicy.maxPenalty}" min="0" step="0.5" title="Maximum penalty (0 = no cap)" />
          <span class="lp-unit">pts cap</span>
        </div>
        ${
          p.submissions.some((s) => s.submissionDate)
            ? `
        <div class="at-row">
          <span class="at-label">Deadline</span>
          <input type="date" id="lp-deadline" class="flex-input" value="${p.latePolicy.deadline}" />
        </div>`
            : ""
        }`
            : ""
        }
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
        rehydrateDates(saved);
        state.dirty = false;
        render();
      });
    }
  }

  bindSidebarEvents();
  autoResizeTextareas(el);
}

function renderFeedbackItem(item: FeedbackItem): string {
  const affectedCount =
    state.project?.submissions.filter((s) =>
      s.appliedFeedback.some((af) => af.itemId === item.id),
    ).length ?? 0;
  return `
    <li class="feedback-item" data-id="${item.id}" draggable="true">
      <span class="fi-drag" title="Drag to reorder">⠿</span>
      <input type="number" class="pts-input fi-pts" value="${item.points}" step="0.5" data-id="${item.id}" />
      <textarea class="fi-label" data-id="${item.id}" placeholder="Feedback text" rows="1">${esc(item.label)}</textarea>
      <button class="fi-count-btn ${affectedCount > 0 ? "has-count" : ""}" data-id="${item.id}" title="Select affected submissions">${affectedCount || ""}</button>
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
    // Missing points affect grade, so re-render full table; also refresh open detail
    renderTable();
    if (state.selectedSubmissionId) renderDetail();
  });
  (
    document.getElementById("at-missing-label") as HTMLInputElement
  )?.addEventListener("input", (e) => {
    p.autoTexts.missing = (e.target as HTMLInputElement).value;
    markDirty();
    // Update summary column for all missing/ungraded submissions
    for (const sub of p.submissions) {
      if (sub.isMissing) updateRow(sub, p);
    }
    if (state.selectedSubmissionId) renderDetail();
  });
  (
    document.getElementById("at-perfect-label") as HTMLInputElement
  )?.addEventListener("input", (e) => {
    p.autoTexts.perfect = (e.target as HTMLInputElement).value;
    markDirty();
    // Update summary column for all marked-perfect submissions
    for (const sub of p.submissions) {
      if (sub.markedPerfect) updateRow(sub, p);
    }
    if (state.selectedSubmissionId) renderDetail();
  });

  // Late policy
  (document.getElementById("lp-type") as HTMLSelectElement)?.addEventListener(
    "change",
    (e) => {
      p.latePolicy.type = (e.target as HTMLSelectElement)
        .value as LatePolicyType;
      markDirty();
      // Changing policy type can activate/deactivate a penalty — clear perfect on affected subs
      for (const sub of p.submissions) {
        if (sub.markedPerfect && computeLatePenalty(sub, p) !== 0)
          clearPerfect(sub);
      }
      renderSidebar();
      renderTable();
      if (state.selectedSubmissionId) renderDetail();
    },
  );
  (document.getElementById("lp-amount") as HTMLInputElement)?.addEventListener(
    "input",
    (e) => {
      p.latePolicy.amount =
        parseFloat((e.target as HTMLInputElement).value) || 0;
      for (const sub of p.submissions) {
        if (sub.markedPerfect && computeLatePenalty(sub, p) !== 0)
          clearPerfect(sub);
      }
      markDirty();
      renderTable();
      if (state.selectedSubmissionId) renderDetail();
    },
  );
  (document.getElementById("lp-max") as HTMLInputElement)?.addEventListener(
    "input",
    (e) => {
      p.latePolicy.maxPenalty =
        parseFloat((e.target as HTMLInputElement).value) || 0;
      for (const sub of p.submissions) {
        if (sub.markedPerfect && computeLatePenalty(sub, p) !== 0)
          clearPerfect(sub);
      }
      markDirty();
      renderTable();
      if (state.selectedSubmissionId) renderDetail();
    },
  );
  (
    document.getElementById("lp-deadline") as HTMLInputElement
  )?.addEventListener("input", (e) => {
    p.latePolicy.deadline = (e.target as HTMLInputElement).value;
    // Recompute daysLate for all submissions with a parsed date, unless manually overridden
    for (const sub of p.submissions) {
      if (sub.submissionDate && p.latePolicy.deadline && !sub.daysLateManual) {
        const deadline = new Date(p.latePolicy.deadline);
        deadline.setHours(23, 59, 59, 999);
        const diffMs = sub.submissionDate.getTime() - deadline.getTime();
        sub.daysLate = diffMs > 0 ? Math.ceil(diffMs / 86400000) : 0;
      }
      // Clear perfect if the updated deadline makes this submission late
      if (sub.markedPerfect && computeLatePenalty(sub, p) !== 0)
        clearPerfect(sub);
    }
    markDirty();
    renderTable();
    if (state.selectedSubmissionId) renderDetail();
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

  // When a feedback item gains a label and focus leaves it, refresh the detail panel
  // so the group panel (and single panel) picks up the newly visible item.
  list?.addEventListener("focusout", (e) => {
    const target = e.target as HTMLElement;
    if (!target.classList.contains("fi-label")) return;
    const id = target.dataset.id;
    if (!id) return;
    const item = p.feedbackItems.find((f) => f.id === id);
    if (!item || !item.label.trim()) return; // empty items handled by blur-to-delete below
    const related = (e as FocusEvent).relatedTarget as HTMLElement | null;
    if (related?.dataset.id === id) return; // focus staying within same item
    // Refresh detail so newly labelled items appear in applied feedback lists
    renderDetail();
  });

  // Blur on an empty label: silently remove the item.
  // We use a mousedown flag to reliably detect clicks into the detail panel,
  // since relatedTarget can be null for checkboxes in some browsers.
  let detailMousedownPending = false;
  document.getElementById("detail-panel")?.addEventListener(
    "mousedown",
    () => {
      detailMousedownPending = true;
      // Reset after a short window — long enough for focusout to fire
      setTimeout(() => {
        detailMousedownPending = false;
      }, 300);
    },
    true,
  ); // capture phase so it fires before focusout

  list?.addEventListener("focusout", (e) => {
    const target = e.target as HTMLInputElement;
    if (!target.classList.contains("fi-label")) return;
    const id = target.dataset.id;
    if (!id) return;
    const item = p.feedbackItems.find((f) => f.id === id);
    if (!item || item.label.trim()) return;
    const related = (e as FocusEvent).relatedTarget as HTMLElement | null;
    if (related?.dataset.id === id) return; // focus within same item
    const doDelete = () => {
      if (!p.feedbackItems.find((f) => f.id === id)) return;
      if (p.feedbackItems.find((f) => f.id === id)?.label.trim()) return;
      p.feedbackItems = p.feedbackItems.filter((f) => f.id !== id);
      markDirty();
      renderSidebar();
      renderDetail();
    };
    if (detailMousedownPending) {
      setTimeout(doDelete, 0);
    } else {
      doDelete();
    }
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
    // Update every row that has this item applied
    for (const sub of p.submissions) {
      if (sub.appliedFeedback.some((af) => af.itemId === id)) {
        updateRow(sub, p);
      }
    }
    if (state.selectedSubmissionId) {
      const sel = p.submissions.find(
        (s) => s.email === state.selectedSubmissionId,
      );
      if (sel?.appliedFeedback.some((af) => af.itemId === id)) renderDetail();
    } else if (state.selectedSubmissionIds.size > 1) {
      const anyAffected = [...state.selectedSubmissionIds].some((email) =>
        p.submissions
          .find((s) => s.email === email)
          ?.appliedFeedback.some((af) => af.itemId === id),
      );
      if (anyAffected) renderDetail();
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
      markDirty();
      renderSidebar();
      renderTable();
      renderDetail();
    } else if (target.classList.contains("fi-count-btn")) {
      selectAffectedSubmissions(id);
    }
  });

  // Apply-drag: dragging a feedback item onto a table row applies it.
  // dragstart sets 'text/feedback-id' so the table drop handler can identify it.
  list?.addEventListener("dragstart", (e) => {
    const li = (e.target as HTMLElement).closest<HTMLElement>(
      "li.feedback-item[data-id]",
    );
    if (!li) return;
    e.dataTransfer!.setData("text/feedback-id", li.dataset.id!);
    e.dataTransfer!.effectAllowed = "all";
    document
      .getElementById("table-container")
      ?.classList.add("table-drop-active");
  });

  list?.addEventListener("dragend", () => {
    document
      .getElementById("table-container")
      ?.classList.remove("table-drop-active");
    document
      .querySelectorAll(".row-drop-target")
      .forEach((el) => el.classList.remove("row-drop-target"));
  });

  const tableContainer = document.getElementById("table-container")!;

  tableContainer.addEventListener("dragover", (e) => {
    if (!e.dataTransfer!.types.includes("text/feedback-id")) return;
    e.preventDefault();
    e.dataTransfer!.dropEffect = "copy";
    const row = (e.target as HTMLElement).closest<HTMLElement>(".sub-row");
    document
      .querySelectorAll(".row-drop-target")
      .forEach((el) => el.classList.remove("row-drop-target"));
    if (row) row.classList.add("row-drop-target");
  });

  tableContainer.addEventListener("dragleave", (e) => {
    if (!tableContainer.contains(e.relatedTarget as Node)) {
      document
        .querySelectorAll(".row-drop-target")
        .forEach((el) => el.classList.remove("row-drop-target"));
    }
  });

  tableContainer.addEventListener("drop", (e) => {
    e.preventDefault();
    document
      .querySelectorAll(".row-drop-target")
      .forEach((el) => el.classList.remove("row-drop-target"));
    document
      .getElementById("table-container")
      ?.classList.remove("table-drop-active");
    const itemId = e.dataTransfer!.getData("text/feedback-id");
    if (!itemId) return;
    const row = (e.target as HTMLElement).closest<HTMLElement>(".sub-row");
    if (!row) return;
    const email = row.dataset.email!;
    const item = p.feedbackItems.find((f) => f.id === itemId);
    if (!item?.label.trim()) return;
    const isMultiTarget =
      state.selectedSubmissionIds.size > 1 &&
      state.selectedSubmissionIds.has(email);
    const targets = isMultiTarget
      ? p.submissions.filter((s) => state.selectedSubmissionIds.has(s.email))
      : [p.submissions.find((s) => s.email === email)!].filter(Boolean);
    for (const sub of targets) applyFeedbackItem(sub, itemId, p);
    renderDetail();
  });
}

// ─── Feedback apply/remove helpers ───────────────────────────────────────────
// Single source of truth for applying/removing a feedback item from a submission.

function applyFeedbackItem(sub: Submission, itemId: string, p: Project) {
  if (sub.appliedFeedback.some((af) => af.itemId === itemId)) return; // already applied
  clearPerfect(sub);
  sub.appliedFeedback.push({ itemId });
  updateFeedbackCountBadge(itemId);
  updateRow(sub, p);
  markDirty();
}

function removeFeedbackItem(sub: Submission, itemId: string, p: Project) {
  if (!sub.appliedFeedback.some((af) => af.itemId === itemId)) return; // not applied
  sub.appliedFeedback = sub.appliedFeedback.filter(
    (af) => af.itemId !== itemId,
  );
  updateFeedbackCountBadge(itemId);
  updateRow(sub, p);
  markDirty();
}

// Replace highlight-on-badge-click with select-affected-submissions
function selectAffectedSubmissions(itemId: string) {
  if (!state.project) return;
  const affected = state.project.submissions
    .filter((s) => s.appliedFeedback.some((af) => af.itemId === itemId))
    .map((s) => s.email);
  if (affected.length === 0) return;
  // Single result → single-select; multiple → multi-select
  state.selectedSubmissionIds = new Set(affected.length > 1 ? affected : []);
  state.selectedSubmissionId = affected.length === 1 ? affected[0] : null;
  state.lastClickedId = affected[affected.length - 1];
  document.querySelectorAll<HTMLElement>(".sub-row").forEach((r) => {
    const em = r.dataset.email!;
    const sel =
      state.selectedSubmissionIds.has(em) || state.selectedSubmissionId === em;
    r.classList.toggle("row-selected", sel);
  });
  renderDetail();
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
      // Default missing penalty to full max grade of the assignment
      if (submissions.length > 0) {
        project.autoTexts.missingPoints = -submissions[0].maxGrade;
      }
      state.project = project;
      state.dirty = true;
      state.fileHandle = null;
      state.selectedSubmissionId = null;
      state.selectedSubmissionIds = new Set();
      state.lastClickedId = null;
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
    rehydrateDates(state.project);
    state.fileHandle = result.handle;
    state.dirty = false;
    state.selectedSubmissionId = null;
    state.selectedSubmissionIds = new Set();
    state.lastClickedId = null;
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
  const hasAnyDate = subs.some((s) => s.submissionDate);
  const showDateCol = hasAnyDate; // only show when we have real parsed dates

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
          ${showDateCol ? `<th class="col-date">Submitted</th>` : ""}
          <th class="col-summary">Feedback</th>
        </tr>
      </thead>
      <tbody>
        ${subs.map((sub) => renderRow(sub, p, showDateCol)).join("")}
      </tbody>
    </table>
  `;

  bindTableEvents();
}

// Any grading action that isn't explicitly "mark perfect" should unmark it
function clearPerfect(sub: Submission) {
  sub.markedPerfect = false;
}

function isGraded(sub: Submission, p?: Project): boolean {
  if (sub.markedPerfect) return true;
  if (sub.appliedFeedback.length > 0) return true;
  if (sub.adHocFeedback.some((ah) => ah.label.trim())) return true; // only non-empty
  if (sub.manualGradeOverride !== undefined) return true;
  if (
    p &&
    p.latePolicy.type !== "none" &&
    (sub.daysLate ?? 0) > 0 &&
    !sub.latePenaltyWaived
  )
    return true;
  return false;
}

function renderRow(sub: Submission, p: Project, showDate = false): string {
  const grade = computeGrade(sub, p);
  const isSelectedSingle = sub.email === state.selectedSubmissionId;
  const isSelectedMulti = state.selectedSubmissionIds.has(sub.email);
  const isSelected = isSelectedSingle || isSelectedMulti;
  const graded = isGraded(sub, p);
  const summary = buildFeedbackSummary(sub, p);
  const daysLate = sub.daysLate ?? 0;
  const isLate = daysLate > 0 && !sub.latePenaltyWaived;

  let dateCell = "";
  if (showDate) {
    if (sub.submissionDate) {
      const fmt = sub.submissionDate.toLocaleDateString("en-US", {
        month: "short",
        day: "numeric",
      });
      dateCell = `<td class="cell-date ${isLate ? "cell-late" : ""}">${fmt}${daysLate > 0 ? ` <span class="late-days">(${daysLate}d)</span>` : ""}</td>`;
    } else if (daysLate > 0) {
      dateCell = `<td class="cell-date cell-late"><span class="late-days">${daysLate}d late</span></td>`;
    } else {
      dateCell = `<td class="cell-date"></td>`;
    }
  }

  return `<tr class="sub-row ${isSelected ? "row-selected" : ""} ${!graded ? "row-ungraded" : ""}"
      data-email="${esc(sub.email)}">
    <td class="cell-name">${esc(sub.fullName)}</td>
    <td class="cell-email">${esc(sub.email)}</td>
    <td class="cell-grade"><span class="grade-num">${fmtGrade(grade)}</span><span class="grade-max"> / ${sub.maxGrade}</span></td>
    ${dateCell}
    <td class="cell-summary"><span class="summary-text">${esc(summary)}</span></td>
  </tr>`;
}

// One-line plain-text summary of feedback for the table cell
function buildFeedbackSummary(sub: Submission, p: Project): string {
  const parts: string[] = [];

  // Missing with no feedback applied and no late penalty (but not if explicitly marked perfect)
  const latePts = computeLatePenalty(sub, p);
  if (
    sub.isMissing &&
    !sub.markedPerfect &&
    sub.appliedFeedback.length === 0 &&
    sub.adHocFeedback.length === 0 &&
    sub.manualGradeOverride === undefined &&
    latePts === 0
  ) {
    return p.autoTexts.missing;
  }

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

  // Late penalty in table summary
  if (latePts !== 0) {
    parts.push(lateFeedbackLabel(sub, p));
  }

  // Perfect: only shown when explicitly marked
  if (parts.length === 0 && sub.markedPerfect) {
    return p.autoTexts.perfect;
  }

  return parts.join("; ");
}

function fmtGrade(g: number): string {
  // Always one decimal, rounded, to prevent column width shifting and floating point display
  return parseFloat(g.toFixed(1)).toFixed(1);
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

  const rows = document.querySelectorAll<HTMLElement>(".sub-row");
  const orderedEmails = Array.from(rows).map((r) => r.dataset.email!);

  rows.forEach((row) => {
    row.addEventListener("click", (e) => {
      const email = row.dataset.email!;
      const isShift = (e as MouseEvent).shiftKey;
      const isCtrl = (e as MouseEvent).ctrlKey || (e as MouseEvent).metaKey;

      if (isShift && state.lastClickedId) {
        // Range select from lastClickedId to this row
        const a = orderedEmails.indexOf(state.lastClickedId);
        const b = orderedEmails.indexOf(email);
        const [lo, hi] = a < b ? [a, b] : [b, a];
        for (let i = lo; i <= hi; i++)
          state.selectedSubmissionIds.add(orderedEmails[i]);
        state.selectedSubmissionId = null;
      } else if (isCtrl) {
        // Toggle individual
        if (state.selectedSubmissionIds.has(email)) {
          state.selectedSubmissionIds.delete(email);
        } else {
          state.selectedSubmissionIds.add(email);
          // Also absorb single-select into the set
          if (state.selectedSubmissionId) {
            state.selectedSubmissionIds.add(state.selectedSubmissionId);
            state.selectedSubmissionId = null;
          }
        }
        state.lastClickedId = email;
      } else {
        // Plain click: single select, clear multi
        state.selectedSubmissionIds.clear();
        if (state.selectedSubmissionId === email) {
          state.selectedSubmissionId = null;
          state.lastClickedId = null;
        } else {
          state.selectedSubmissionId = email;
          state.lastClickedId = email;
        }
      }

      // Sync row highlight classes
      document.querySelectorAll<HTMLElement>(".sub-row").forEach((r) => {
        const em = r.dataset.email!;
        const sel =
          state.selectedSubmissionIds.has(em) ||
          (!state.selectedSubmissionIds.size &&
            state.selectedSubmissionId === em);
        r.classList.toggle("row-selected", sel);
      });

      renderDetail();
    });
  });
}

// ─── Detail panel ─────────────────────────────────────────────────────────────

function renderDetail() {
  if (state.selectedSubmissionIds.size > 1) {
    renderGroupDetail();
  } else if (state.selectedSubmissionIds.size === 1) {
    // Collapsed multi-select — treat as single
    state.selectedSubmissionId = [...state.selectedSubmissionIds][0];
    state.selectedSubmissionIds.clear();
    renderSingleDetail();
  } else {
    renderSingleDetail();
  }
}

function renderGroupDetail() {
  const el = document.getElementById("detail-panel")!;
  const p = state.project!;
  const emails = Array.from(state.selectedSubmissionIds);
  const subs = emails
    .map((em) => p.submissions.find((s) => s.email === em)!)
    .filter(Boolean);

  el.innerHTML = `
    <div class="detail-hdr">
      <div class="detail-name">${subs.length} submissions selected</div>
      <button class="btn-icon close-detail" title="Clear selection">×</button>
    </div>

    <div class="group-names">
      ${subs.map((s) => `<span class="group-name-tag">${esc(s.fullName)}</span>`).join("")}
    </div>

    <div class="grade-row">
      <div class="grade-label">Group actions</div>
      <div class="grade-controls">
        <button class="btn-perfect-action" id="btn-group-perfect">★ Mark perfect</button>
        <button class="btn-clear-action" id="btn-group-clear">✕ Clear all</button>
      </div>
    </div>

    ${
      p.latePolicy.type !== "none"
        ? (() => {
            // Only show when none of the selected subs have auto-computed dates
            // (i.e. all are manual-entry candidates)
            const anyHasDate = subs.some(
              (s) => s.submissionDate && !s.daysLateManual,
            );
            const daysValues = subs.map((s) => s.daysLate ?? 0);
            const allSame = daysValues.every((d) => d === daysValues[0]);
            return `
      <div class="detail-section">
        <div class="detail-sub">Late penalty</div>
        <div class="late-controls">
          <label class="late-days-label">Days late:</label>
          <input type="number" id="group-days-late" class="pts-input"
            value="${allSame ? daysValues[0] : ""}"
            placeholder="${allSame ? "" : "mixed"}" min="0" step="1" />
          ${anyHasDate ? `<button class="btn-icon" id="btn-group-reset-days" title="Reset all to computed values">↺</button>` : ""}
        </div>
      </div>`;
          })()
        : ""
    }

    <div class="detail-section">
      <div class="detail-sub">
        Apply feedback
        <span class="detail-sub-hint">— changes apply to all selected</span>
      </div>
      <ul class="detail-fi-list" id="group-fi-list">
        ${(() => {
          const labelled = p.feedbackItems.filter((item) => item.label.trim());
          if (labelled.length === 0)
            return `<li class="detail-empty-hint">No reusable feedback defined yet.</li>`;
          return labelled.map((item) => renderGroupFI(item, subs)).join("");
        })()}
      </ul>
    </div>
  `;

  bindGroupDetailEvents(subs, p);
}

function renderGroupFI(item: FeedbackItem, subs: Submission[]): string {
  const appliedCount = subs.filter((s) =>
    s.appliedFeedback.some((af) => af.itemId === item.id),
  ).length;
  const allApplied = appliedCount === subs.length;
  const someApplied = appliedCount > 0 && !allApplied;

  return `<li class="dfi ${allApplied ? "dfi-applied" : someApplied ? "dfi-mixed" : ""}" data-id="${item.id}">
    <input type="checkbox" class="gfi-check" data-id="${item.id}"
      ${allApplied ? "checked" : ""} ${someApplied ? 'data-indeterminate="true"' : ""} />
    <span class="gfi-label">${esc(item.label)}</span>
    <span class="gfi-pts">${fmtPts(item.points)}</span>
    <span class="gfi-count">${appliedCount}/${subs.length}</span>
  </li>`;
}

function bindGroupDetailEvents(subs: Submission[], p: Project) {
  document.querySelector(".close-detail")?.addEventListener("click", () => {
    state.selectedSubmissionIds.clear();
    state.selectedSubmissionId = null;
    state.lastClickedId = null;
    document
      .querySelectorAll(".sub-row")
      .forEach((r) => r.classList.remove("row-selected"));
    renderDetail();
  });

  document
    .getElementById("btn-group-perfect")
    ?.addEventListener("click", () => {
      const affectedItemIds = new Set(
        subs.flatMap((s) => s.appliedFeedback.map((af) => af.itemId)),
      );
      for (const sub of subs) {
        sub.appliedFeedback = [];
        sub.adHocFeedback = [];
        delete sub.manualGradeOverride;
        sub.daysLate = 0;
        sub.latePenaltyWaived = false;
        sub.markedPerfect = true;
        updateRow(sub, p);
      }
      affectedItemIds.forEach((id) => updateFeedbackCountBadge(id));
      markDirty();
      renderGroupDetail();
    });

  document.getElementById("btn-group-clear")?.addEventListener("click", () => {
    const affectedItemIds = new Set(
      subs.flatMap((s) => s.appliedFeedback.map((af) => af.itemId)),
    );
    for (const sub of subs) {
      sub.appliedFeedback = [];
      sub.adHocFeedback = [];
      delete sub.manualGradeOverride;
      sub.markedPerfect = false;
      sub.daysLate = 0;
      sub.daysLateManual = false;
      sub.latePenaltyWaived = false;
      updateRow(sub, p);
    }
    affectedItemIds.forEach((id) => updateFeedbackCountBadge(id));
    markDirty();
    renderGroupDetail();
  });

  // Set indeterminate state on mixed checkboxes (can't do in HTML)
  document
    .querySelectorAll<HTMLInputElement>(".gfi-check[data-indeterminate]")
    .forEach((cb) => {
      cb.indeterminate = true;
    });

  const fiList = document.getElementById("group-fi-list");
  fiList?.addEventListener("change", (e) => {
    const t = e.target as HTMLInputElement;
    if (!t.classList.contains("gfi-check")) return;
    const id = t.dataset.id!;
    if (t.checked) {
      for (const sub of subs) applyFeedbackItem(sub, id, p);
    } else {
      for (const sub of subs) removeFeedbackItem(sub, id, p);
    }
    renderGroupDetail();
  });

  // Group days-late
  (
    document.getElementById("group-days-late") as HTMLInputElement
  )?.addEventListener("input", (e) => {
    const val = parseInt((e.target as HTMLInputElement).value);
    if (isNaN(val)) return;
    for (const sub of subs) {
      clearPerfect(sub);
      sub.daysLate = val;
      sub.daysLateManual = true;
      updateRow(sub, p);
    }
    markDirty();
  });

  document
    .getElementById("btn-group-reset-days")
    ?.addEventListener("click", () => {
      for (const sub of subs) {
        if (sub.submissionDate && p.latePolicy.deadline) {
          const deadline = new Date(p.latePolicy.deadline);
          deadline.setHours(23, 59, 59, 999);
          const diffMs = sub.submissionDate.getTime() - deadline.getTime();
          sub.daysLate = diffMs > 0 ? Math.ceil(diffMs / 86400000) : 0;
        } else {
          sub.daysLate = 0;
        }
        sub.daysLateManual = false;
        updateRow(sub, p);
      }
      markDirty();
      renderGroupDetail();
    });
}

function renderSingleDetail() {
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
  const lp = p.latePolicy;
  const latePts = computeLatePenalty(sub, p);
  const hasDate = !!sub.submissionDate;

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
        <span class="grade-big ${isOverride ? "grade-overridden" : sub.markedPerfect ? "grade-perfect" : ""}" id="grade-display">${fmtGrade(grade)}</span>
        <span class="grade-max-detail"> / ${sub.maxGrade}</span>
        <input type="number" id="manual-grade" class="pts-input manual-grade-input"
          value="${isOverride ? sub.manualGradeOverride : fmtGrade(grade)}"
          min="0" step="0.5" ${!isOverride ? "disabled" : ""} />
        <button class="btn-icon grade-lock-btn" id="btn-grade-lock"
          title="${isOverride ? "Revert to computed grade" : "Set grade manually"}"
        >${isOverride ? "🔓" : "🔒"}</button>
        <button class="${sub.markedPerfect ? "btn-perfect-active" : "btn-perfect-action"}" id="btn-perfect"
          title="${sub.markedPerfect ? "Unmark perfect" : "Mark as full credit"}"
        >${sub.markedPerfect ? "★ Perfect" : "★ Mark perfect"}</button>
      </div>
      ${isOverride ? `<div class="override-note">Manual override active</div>` : ""}
      ${sub.markedPerfect && !isOverride ? `<div class="perfect-note">Marked perfect — ${esc(p.autoTexts.perfect)}</div>` : ""}
    </div>

    ${
      lp.type !== "none"
        ? `
    <div class="detail-section late-row-section">
      <div class="detail-sub">Late penalty</div>
      <div class="late-controls">
        ${
          hasDate
            ? `<span class="late-info">Submitted ${sub.submissionDate!.toLocaleDateString("en-US", { month: "short", day: "numeric" })}</span>`
            : ""
        }
        <label class="late-days-label">Days late:</label>
        <input type="number" id="days-late" class="pts-input" value="${sub.daysLate ?? 0}" min="0" step="1" />
        ${
          hasDate && sub.daysLateManual
            ? `<button class="btn-icon" id="btn-reset-days-late" title="Revert to computed value">↺</button>`
            : ""
        }
        ${latePts !== 0 ? `<span class="late-penalty-amt">${fmtPts(latePts)}</span>` : '<span class="late-penalty-amt late-none">on time</span>'}
        <label class="inline-check late-waive">
          <input type="checkbox" id="waive-late" ${sub.latePenaltyWaived ? "checked" : ""} /> waive
        </label>
      </div>
    </div>
    `
        : ""
    }

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
    state.selectedSubmissionIds.clear();
    state.lastClickedId = null;
    document
      .querySelectorAll(".sub-row")
      .forEach((r) => r.classList.remove("row-selected"));
    renderDetail();
  });

  // Days-late input — always editable; marks as manual override when changed
  (document.getElementById("days-late") as HTMLInputElement)?.addEventListener(
    "input",
    (e) => {
      clearPerfect(sub);
      sub.daysLate = parseInt((e.target as HTMLInputElement).value) || 0;
      sub.daysLateManual = true;
      markDirty();
      refreshGrade(sub, p);
      updateRow(sub, p);
      refreshPreview(sub, p);
      renderDetail(); // re-render to show ↺ button and update perfect button state
    },
  );

  // Reset days-late to auto-computed value from submission date vs deadline
  document
    .getElementById("btn-reset-days-late")
    ?.addEventListener("click", () => {
      if (sub.submissionDate && p.latePolicy.deadline) {
        const deadline = new Date(p.latePolicy.deadline);
        deadline.setHours(23, 59, 59, 999);
        const diffMs = sub.submissionDate.getTime() - deadline.getTime();
        sub.daysLate = diffMs > 0 ? Math.ceil(diffMs / 86400000) : 0;
      } else {
        sub.daysLate = 0;
      }
      sub.daysLateManual = false;
      markDirty();
      renderDetail();
      updateRow(sub, p);
    });

  (document.getElementById("waive-late") as HTMLInputElement)?.addEventListener(
    "change",
    (e) => {
      clearPerfect(sub);
      sub.latePenaltyWaived = (e.target as HTMLInputElement).checked;
      markDirty();
      renderSingleDetail();
      updateRow(sub, p);
    },
  );

  // Lock/unlock grade
  document.getElementById("btn-grade-lock")?.addEventListener("click", () => {
    if (sub.manualGradeOverride !== undefined) {
      delete sub.manualGradeOverride;
    } else {
      clearPerfect(sub);
      sub.manualGradeOverride = computeGrade(sub, p);
    }
    markDirty();
    renderDetail();
    updateRow(sub, p);
  });

  // Perfect score toggle
  document.getElementById("btn-perfect")?.addEventListener("click", () => {
    if (sub.markedPerfect) {
      sub.markedPerfect = false;
    } else {
      sub.appliedFeedback = [];
      sub.adHocFeedback = [];
      delete sub.manualGradeOverride;
      sub.daysLate = 0;
      sub.latePenaltyWaived = false;
      sub.markedPerfect = true;
    }
    markDirty();
    renderDetail();
    updateRow(sub, p);
  });

  (
    document.getElementById("manual-grade") as HTMLInputElement
  )?.addEventListener("input", (e) => {
    clearPerfect(sub);
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
      applyFeedbackItem(sub, id, p);
    } else {
      removeFeedbackItem(sub, id, p);
    }
    renderDetail();
    updateRow(sub, p);
  });

  fiList?.addEventListener("keydown", (e) => {
    const t = e.target as HTMLElement;
    if (e.key !== "Enter" || !t.classList.contains("dfi-lbl")) return;
    e.preventDefault();
    (t as HTMLTextAreaElement).blur();
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
    if (t.classList.contains("ah-lbl")) {
      ah.label = t.value;
      if (ah.label.trim()) clearPerfect(sub); // only clear perfect once something is typed
    }
    markDirty();
    refreshGrade(sub, p);
    updateRow(sub, p);
    refreshPreview(sub, p);
  });

  ahList?.addEventListener("focusout", (e) => {
    const target = e.target as HTMLInputElement;
    if (!target.classList.contains("ah-lbl")) return;
    const ahid = target.dataset.ahid;
    if (!ahid) return;
    const ah = sub.adHocFeedback.find((a) => a.id === ahid);
    if (!ah || ah.label.trim()) return;
    // Remove empty entry when focus leaves (unless staying within same item)
    const related = (e as FocusEvent).relatedTarget as HTMLElement | null;
    if (related?.dataset.ahid === ahid) return;
    sub.adHocFeedback = sub.adHocFeedback.filter((a) => a.id !== ahid);
    markDirty();
    renderDetail();
    updateRow(sub, p);
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
  const rounded = parseFloat(pts.toFixed(1));
  if (rounded > 0) return `+${rounded.toFixed(1)}`;
  if (rounded === 0) return "0";
  return rounded.toFixed(1);
}

function renderFeedbackSummaryList(sub: Submission, p: Project): string {
  const items: Array<{ pts: number; label: string; positive: boolean }> = [];

  if (
    sub.isMissing &&
    !sub.markedPerfect &&
    sub.appliedFeedback.length === 0 &&
    sub.adHocFeedback.length === 0 &&
    sub.manualGradeOverride === undefined &&
    computeLatePenalty(sub, p) === 0
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
    // Late penalty
    const latePts = computeLatePenalty(sub, p);
    if (latePts !== 0) {
      items.push({
        pts: latePts,
        label: lateFeedbackLabel(sub, p),
        positive: false,
      });
    }
    // Perfect auto-text: only shown when explicitly marked, not inferred from grade
    if (items.length === 0 && sub.markedPerfect) {
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
  if (summaryCell) summaryCell.textContent = buildFeedbackSummary(sub, p);
  row.classList.toggle("row-ungraded", !isGraded(sub, p));
  const isSelected =
    sub.email === state.selectedSubmissionId ||
    state.selectedSubmissionIds.has(sub.email);
  row.classList.toggle("row-selected", isSelected);
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
    <div class="resize-handle" id="resize-left" title="Drag to resize"></div>
    <main id="main">
      <div id="table-container"></div>
      <footer id="footer">
        <a href="https://github.com/adambyle/grader" target="_blank" rel="noopener">github.com/adambyle/grader</a>
        <span>© ${new Date().getFullYear()} Adam Byle</span>
      </footer>
    </main>
    <div class="resize-handle" id="resize-right" title="Drag to resize"></div>
    <section id="detail-panel"></section>
  </div>
`;

// ─── Reorder drag (persistent — outside renderSidebar) ───────────────────────

function initReorderDrag() {
  const sidebar = document.getElementById("sidebar")!;
  let draggingId: string | null = null;
  let dragOverId: string | null = null;

  sidebar.addEventListener("dragstart", (e) => {
    const li = (e.target as HTMLElement).closest<HTMLElement>(
      "li.feedback-item[data-id]",
    );
    if (!li) return;
    draggingId = li.dataset.id!;
    e.dataTransfer!.effectAllowed = "all";
    e.dataTransfer!.setData("text/plain", draggingId);
    setTimeout(() => li.classList.add("fi-dragging"), 0);
  });

  sidebar.addEventListener("dragend", () => {
    draggingId = null;
    dragOverId = null;
    sidebar
      .querySelectorAll(".fi-drag-over")
      .forEach((el) => el.classList.remove("fi-drag-over"));
    sidebar
      .querySelectorAll(".fi-dragging")
      .forEach((el) => el.classList.remove("fi-dragging"));
  });

  sidebar.addEventListener("dragover", (e) => {
    if (!draggingId) return;
    const li = (e.target as HTMLElement).closest<HTMLElement>(
      "li.feedback-item[data-id]",
    );
    if (!li || li.dataset.id === draggingId) return;
    e.preventDefault();
    e.dataTransfer!.dropEffect = "move";
    if (li.dataset.id !== dragOverId) {
      sidebar
        .querySelectorAll(".fi-drag-over")
        .forEach((el) => el.classList.remove("fi-drag-over"));
      li.classList.add("fi-drag-over");
      dragOverId = li.dataset.id!;
    }
  });

  sidebar.addEventListener("dragleave", (e) => {
    const li = (e.target as HTMLElement).closest<HTMLElement>(
      "li.feedback-item[data-id]",
    );
    if (li && !li.contains(e.relatedTarget as Node)) {
      li.classList.remove("fi-drag-over");
      if (dragOverId === li.dataset.id) dragOverId = null;
    }
  });

  sidebar.addEventListener("drop", (e) => {
    if (!draggingId) return;
    const li = (e.target as HTMLElement).closest<HTMLElement>(
      "li.feedback-item[data-id]",
    );
    if (!li || li.dataset.id === draggingId) return;
    e.preventDefault();
    const p = state.project;
    if (!p) return;
    const fromIdx = p.feedbackItems.findIndex((f) => f.id === draggingId);
    const toIdx = p.feedbackItems.findIndex((f) => f.id === li.dataset.id);
    if (fromIdx < 0 || toIdx < 0) return;
    const [moved] = p.feedbackItems.splice(fromIdx, 1);
    p.feedbackItems.splice(toIdx, 0, moved);
    draggingId = null;
    dragOverId = null;
    markDirty();
    renderSidebar();
    renderDetail();
  });
}

// ─── Resize handles ───────────────────────────────────────────────────────────

function initResizeHandles() {
  function makeDragger(
    handleId: string,
    cssVar: string,
    side: "left" | "right",
  ) {
    const handle = document.getElementById(handleId)!;
    let startX = 0;
    let startVal = 0;

    handle.addEventListener("mousedown", (e) => {
      startX = e.clientX;
      const raw = getComputedStyle(document.documentElement)
        .getPropertyValue(cssVar)
        .trim();
      startVal = parseInt(raw) || (side === "left" ? 282 : 318);
      document.body.style.cursor = "col-resize";
      document.body.style.userSelect = "none";

      function onMove(e: MouseEvent) {
        // Left handle: drag right → sidebar grows. Right handle: drag left → detail grows.
        const delta = side === "left" ? e.clientX - startX : startX - e.clientX;
        const next = Math.max(200, Math.min(520, startVal + delta));
        document.documentElement.style.setProperty(cssVar, next + "px");
      }

      function onUp() {
        document.body.style.cursor = "";
        document.body.style.userSelect = "";
        window.removeEventListener("mousemove", onMove);
        window.removeEventListener("mouseup", onUp);
        // Persist sizes
        localStorage.setItem(
          "grader_" + cssVar.slice(2),
          String(
            parseInt(
              getComputedStyle(document.documentElement).getPropertyValue(
                cssVar,
              ),
            ),
          ),
        );
      }

      window.addEventListener("mousemove", onMove);
      window.addEventListener("mouseup", onUp);
      e.preventDefault();
    });
  }

  makeDragger("resize-left", "--sidebar-w", "left");
  makeDragger("resize-right", "--detail-w", "right");

  // Restore persisted sizes
  const savedSidebar = localStorage.getItem("grader_sidebar-w");
  const savedDetail = localStorage.getItem("grader_detail-w");
  if (savedSidebar)
    document.documentElement.style.setProperty(
      "--sidebar-w",
      savedSidebar + "px",
    );
  if (savedDetail)
    document.documentElement.style.setProperty(
      "--detail-w",
      savedDetail + "px",
    );
}

// ─── Global keyboard shortcuts ────────────────────────────────────────────────

document.addEventListener("keydown", (e) => {
  // Ctrl/Cmd+S: save
  if ((e.ctrlKey || e.metaKey) && e.key === "s") {
    e.preventDefault();
    if (state.project) handleSaveJSON();
    return;
  }
  if (e.key !== "Escape") return;
  const active = document.activeElement as HTMLElement | null;
  if (
    active &&
    (active.tagName === "INPUT" ||
      active.tagName === "TEXTAREA" ||
      active.tagName === "SELECT")
  ) {
    active.blur();
  }
});

initReorderDrag();
initResizeHandles();
render();
