"use client";

import { useCallback, useRef, useState } from "react";

type Status =
  | { state: "idle" }
  | { state: "processing"; name: string }
  | { state: "done"; name: string; images: number; media: number }
  | { state: "error"; message: string };

export default function Home() {
  const [status, setStatus] = useState<Status>({ state: "idle" });
  const [dragging, setDragging] = useState(false);
  const inputRef = useRef<HTMLInputElement>(null);

  const handleFile = useCallback(async (file: File) => {
    if (!file.name.toLowerCase().endsWith(".docx")) {
      setStatus({ state: "error", message: "Only .docx files are supported." });
      return;
    }
    setStatus({ state: "processing", name: file.name });

    try {
      const body = new FormData();
      body.append("file", file);
      const res = await fetch("/api/remove-images", { method: "POST", body });

      if (!res.ok) {
        const data = await res.json().catch(() => null);
        throw new Error(data?.error ?? `Request failed (${res.status}).`);
      }

      const blob = await res.blob();
      const outName = file.name.replace(/\.docx$/i, "") + "-no-images.docx";
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = outName;
      a.click();
      URL.revokeObjectURL(url);

      setStatus({
        state: "done",
        name: outName,
        images: Number(res.headers.get("X-Images-Removed") ?? 0),
        media: Number(res.headers.get("X-Media-Removed") ?? 0),
      });
    } catch (err) {
      setStatus({
        state: "error",
        message: err instanceof Error ? err.message : "Something went wrong.",
      });
    }
  }, []);

  const onDrop = useCallback(
    (e: React.DragEvent) => {
      e.preventDefault();
      setDragging(false);
      const file = e.dataTransfer.files[0];
      if (file) handleFile(file);
    },
    [handleFile]
  );

  const processing = status.state === "processing";

  return (
    <main className="flex flex-1 flex-col items-center justify-center gap-8 p-8 font-sans">
      <div className="flex flex-col items-center gap-2 text-center">
        <h1 className="text-3xl font-bold tracking-tight">
          Word Image Remover
        </h1>
        <p className="max-w-md text-sm text-gray-500 dark:text-gray-400">
          Upload a .docx file and get back the same document with every image
          stripped out. Text, formatting, and layout stay untouched.
        </p>
      </div>

      <div
        role="button"
        tabIndex={0}
        aria-label="Upload a .docx file"
        onClick={() => !processing && inputRef.current?.click()}
        onKeyDown={(e) => {
          if ((e.key === "Enter" || e.key === " ") && !processing) {
            e.preventDefault();
            inputRef.current?.click();
          }
        }}
        onDragOver={(e) => {
          e.preventDefault();
          setDragging(true);
        }}
        onDragLeave={() => setDragging(false)}
        onDrop={onDrop}
        className={`flex w-full max-w-xl cursor-pointer flex-col items-center gap-3 rounded-2xl border-2 border-dashed p-12 transition-colors ${
          dragging
            ? "border-blue-500 bg-blue-50 dark:bg-blue-950/30"
            : "border-gray-300 hover:border-gray-400 dark:border-gray-700 dark:hover:border-gray-500"
        } ${processing ? "pointer-events-none opacity-60" : ""}`}
      >
        <svg
          className="h-10 w-10 text-gray-400"
          fill="none"
          viewBox="0 0 24 24"
          stroke="currentColor"
          strokeWidth={1.5}
          aria-hidden="true"
        >
          <path
            strokeLinecap="round"
            strokeLinejoin="round"
            d="M3 16.5v2.25A2.25 2.25 0 0 0 5.25 21h13.5A2.25 2.25 0 0 0 21 18.75V16.5m-13.5-9L12 3m0 0 4.5 4.5M12 3v13.5"
          />
        </svg>
        <p className="font-medium">
          {processing ? "Processing…" : "Drop a .docx file here"}
        </p>
        <p className="text-sm text-gray-500 dark:text-gray-400">
          or click to browse
        </p>
        <input
          ref={inputRef}
          type="file"
          accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
          className="hidden"
          onChange={(e) => {
            const file = e.target.files?.[0];
            if (file) handleFile(file);
            e.target.value = "";
          }}
        />
      </div>

      <div
        className="min-h-12 w-full max-w-xl text-center text-sm"
        aria-live="polite"
      >
        {status.state === "processing" && (
          <p className="text-gray-500 dark:text-gray-400">
            Removing images from{" "}
            <span className="font-medium">{status.name}</span>…
          </p>
        )}
        {status.state === "done" && (
          <div className="rounded-lg bg-green-50 p-4 text-green-800 dark:bg-green-950/40 dark:text-green-300">
            <p className="font-medium">Done — downloaded {status.name}</p>
            <p className="mt-1">
              Removed {status.images} image{status.images === 1 ? "" : "s"} and{" "}
              {status.media} media file{status.media === 1 ? "" : "s"}.
            </p>
          </div>
        )}
        {status.state === "error" && (
          <div className="rounded-lg bg-red-50 p-4 text-red-800 dark:bg-red-950/40 dark:text-red-300">
            {status.message}
          </div>
        )}
      </div>
    </main>
  );
}
