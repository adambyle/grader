import type { Project, AppState } from "./types";

const LS_KEY = "grader_autosave";

export function projectToJSON(project: Project): string {
  return JSON.stringify(project, null, 2);
}

export function projectFromJSON(text: string): Project {
  return JSON.parse(text) as Project;
}

export function saveToLocalStorage(project: Project) {
  try {
    localStorage.setItem(LS_KEY, projectToJSON(project));
  } catch {
    // quota exceeded — silently ignore
  }
}

export function loadFromLocalStorage(): Project | null {
  const raw = localStorage.getItem(LS_KEY);
  if (!raw) return null;
  try {
    return projectFromJSON(raw);
  } catch {
    return null;
  }
}

export function clearLocalStorage() {
  localStorage.removeItem(LS_KEY);
}

const hasFileSystemAccess = "showOpenFilePicker" in window;

export async function saveToFile(
  project: Project,
  state: AppState,
): Promise<FileSystemFileHandle | null> {
  if (!hasFileSystemAccess) {
    // Fallback: trigger a plain download
    downloadBlob(
      `${project.assignmentName || "grading"}.json`,
      projectToJSON(project),
      "application/json",
    );
    return null;
  }
  try {
    let handle = state.fileHandle;
    if (!handle) {
      handle = await (window as any).showSaveFilePicker({
        suggestedName: `${project.assignmentName || "grading"}.json`,
        types: [
          {
            description: "Grader JSON",
            accept: { "application/json": [".json"] },
          },
        ],
      });
    }
    const writable = await handle!.createWritable();
    await writable.write(projectToJSON(project));
    await writable.close();
    return handle!;
  } catch (e: any) {
    if (e?.name === "AbortError") return null;
    throw e;
  }
}

export async function loadFromFile(): Promise<{
  project: Project;
  handle: FileSystemFileHandle;
} | null> {
  if (!hasFileSystemAccess) {
    // Fallback: use a hidden <input type="file">
    return new Promise((resolve, reject) => {
      const input = document.createElement("input");
      input.type = "file";
      input.accept = ".json,application/json";
      input.onchange = async () => {
        const file = input.files?.[0];
        if (!file) {
          resolve(null);
          return;
        }
        try {
          const text = await file.text();
          resolve({ project: projectFromJSON(text), handle: null as any });
        } catch (e) {
          reject(e);
        }
      };
      input.oncancel = () => resolve(null);
      input.click();
    });
  }
  try {
    const [handle] = await (window as any).showOpenFilePicker({
      types: [
        {
          description: "Grader JSON",
          accept: { "application/json": [".json"] },
        },
      ],
    });
    const file = await handle.getFile();
    const text = await file.text();
    return { project: projectFromJSON(text), handle };
  } catch (e: any) {
    if (e?.name === "AbortError") return null;
    throw e;
  }
}

export function downloadBlob(
  filename: string,
  content: string,
  mime = "text/csv",
) {
  const blob = new Blob(["\ufeff" + content], { type: mime }); // BOM for Excel compat
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}
