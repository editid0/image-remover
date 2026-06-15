"use client";

import { useCallback, useRef, useState } from "react";
import { DOCX_MIME } from "./util";

export default function Dropzone({
  onFile,
  disabled = false,
  prompt,
}: {
  onFile: (file: File) => void;
  disabled?: boolean;
  prompt: string;
}) {
  const [dragging, setDragging] = useState(false);
  const inputRef = useRef<HTMLInputElement>(null);

  const pick = useCallback(() => {
    if (!disabled) inputRef.current?.click();
  }, [disabled]);

  return (
    <div
      role="button"
      tabIndex={0}
      aria-label="Upload a .docx file"
      onClick={pick}
      onKeyDown={(e) => {
        if (e.key === "Enter" || e.key === " ") {
          e.preventDefault();
          pick();
        }
      }}
      onDragOver={(e) => {
        e.preventDefault();
        if (!disabled) setDragging(true);
      }}
      onDragLeave={() => setDragging(false)}
      onDrop={(e) => {
        e.preventDefault();
        setDragging(false);
        if (disabled) return;
        const file = e.dataTransfer.files[0];
        if (file) onFile(file);
      }}
      className={`flex w-full max-w-xl cursor-pointer flex-col items-center gap-3 rounded-2xl border-2 border-dashed p-12 transition-colors ${
        dragging
          ? "border-blue-500 bg-blue-50 dark:bg-blue-950/30"
          : "border-gray-300 hover:border-gray-400 dark:border-gray-700 dark:hover:border-gray-500"
      } ${disabled ? "pointer-events-none opacity-60" : ""}`}
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
      <p className="font-medium">{prompt}</p>
      <p className="text-sm text-gray-500 dark:text-gray-400">or click to browse</p>
      <input
        ref={inputRef}
        type="file"
        accept={`.docx,${DOCX_MIME}`}
        className="hidden"
        onChange={(e) => {
          const file = e.target.files?.[0];
          if (file) onFile(file);
          e.target.value = "";
        }}
      />
    </div>
  );
}
