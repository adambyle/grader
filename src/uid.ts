let counter = 0;
export function uid(): string {
  return `${Date.now()}-${++counter}-${Math.random().toString(36).slice(2, 7)}`;
}
