"use client";

import { useState } from "react";
import RemovePanel from "./RemovePanel";
import ExtractPanel from "./ExtractPanel";
import ScrubPanel from "./ScrubPanel";

const MODES = [
  {
    id: "remove",
    label: "Remove images",
    blurb:
      "Get back the same document with every image stripped out. Text, formatting, and layout stay untouched.",
  },
  {
    id: "extract",
    label: "Extract images",
    blurb:
      "Pull every embedded image out of the document and download them individually or as a single zip.",
  },
  {
    id: "scrub",
    label: "Edit metadata",
    blurb:
      "See the author, dates, and other hidden properties baked into the file. Edit values, strip what you want gone, or add new Word properties.",
  },
] as const;

type Mode = (typeof MODES)[number]["id"];

export default function Home() {
  const [mode, setMode] = useState<Mode>("remove");
  const active = MODES.find((m) => m.id === mode)!;

  return (
    <main className="flex flex-1 flex-col items-center justify-center gap-8 p-8 font-sans">
      <div className="flex flex-col items-center gap-2 text-center">
        <h1 className="text-3xl font-bold tracking-tight">Word Doc Toolkit</h1>
        <p className="max-w-md text-sm text-gray-500 dark:text-gray-400">
          {active.blurb} Everything runs in your browser — files never leave your
          device.
        </p>
      </div>

      <div
        role="tablist"
        aria-label="Tool"
        className="inline-flex rounded-xl bg-gray-100 p-1 dark:bg-gray-800"
      >
        {MODES.map((m) => (
          <button
            key={m.id}
            role="tab"
            aria-selected={m.id === mode}
            onClick={() => setMode(m.id)}
            className={`rounded-lg px-4 py-2 text-sm font-medium transition-colors ${
              m.id === mode
                ? "bg-white text-gray-900 shadow-sm dark:bg-gray-700 dark:text-white"
                : "text-gray-500 hover:text-gray-800 dark:text-gray-400 dark:hover:text-gray-200"
            }`}
          >
            {m.label}
          </button>
        ))}
      </div>

      {mode === "remove" && <RemovePanel />}
      {mode === "extract" && <ExtractPanel />}
      {mode === "scrub" && <ScrubPanel />}
    </main>
  );
}
