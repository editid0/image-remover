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
