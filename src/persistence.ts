import type { Project, AppState } from './types';

const LS_KEY = 'grader_autosave';

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

export async function saveToFile(project: Project, state: AppState): Promise<FileSystemFileHandle | null> {
  try {
    let handle = state.fileHandle;
    if (!handle) {
      handle = await (window as any).showSaveFilePicker({
        suggestedName: `${project.assignmentName || 'grading'}.json`,
        types: [{ description: 'Grader JSON', accept: { 'application/json': ['.json'] } }],
      });
    }
    const writable = await handle!.createWritable();
    await writable.write(projectToJSON(project));
    await writable.close();
    return handle!;
  } catch (e: any) {
    if (e?.name === 'AbortError') return null; // user cancelled
    throw e;
  }
}

export async function loadFromFile(): Promise<{ project: Project; handle: FileSystemFileHandle } | null> {
  try {
    const [handle] = await (window as any).showOpenFilePicker({
      types: [{ description: 'Grader JSON', accept: { 'application/json': ['.json'] } }],
    });
    const file = await handle.getFile();
    const text = await file.text();
    return { project: projectFromJSON(text), handle };
  } catch (e: any) {
    if (e?.name === 'AbortError') return null;
    throw e;
  }
}

export function downloadBlob(filename: string, content: string, mime = 'text/csv') {
  const blob = new Blob(['\ufeff' + content], { type: mime }); // BOM for Excel compat
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}
