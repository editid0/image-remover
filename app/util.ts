export const DOCX_MIME =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

/** Trigger a browser download for raw bytes or a Blob. */
export function downloadBlob(
  data: Uint8Array | Blob,
  filename: string,
  type?: string
) {
  const blob =
    data instanceof Blob
      ? data
      : new Blob([data.slice().buffer], type ? { type } : undefined);
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}

/** Strip the trailing `.docx` extension from a file name. */
export function baseName(name: string): string {
  return name.replace(/\.docx$/i, "");
}

export function isDocx(file: File): boolean {
  return file.name.toLowerCase().endsWith(".docx");
}

/** Human-readable byte size, e.g. 1536 -> "1.5 KB". */
export function formatBytes(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  const units = ["KB", "MB", "GB"];
  let value = bytes / 1024;
  let i = 0;
  while (value >= 1024 && i < units.length - 1) {
    value /= 1024;
    i++;
  }
  return `${value.toFixed(value < 10 ? 1 : 0)} ${units[i]}`;
}
