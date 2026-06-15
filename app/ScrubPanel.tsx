"use client";

import { useCallback, useMemo, useState } from "react";
import {
  readDocxMetadata,
  scrubDocxMetadata,
  type MetadataField,
} from "@/lib/docx";
import Dropzone from "./Dropzone";
import { DOCX_MIME, baseName, downloadBlob, isDocx } from "./util";

type State =
  | { phase: "idle" }
  | { phase: "reading"; name: string }
  | {
      phase: "review";
      base: string;
      buffer: ArrayBuffer;
      fields: MetadataField[];
      selected: Set<string>;
    }
  | { phase: "none"; name: string }
  | { phase: "processing"; base: string; buffer: ArrayBuffer; selected: Set<string> }
  | { phase: "done"; name: string; removed: number }
  | { phase: "error"; message: string };

const GROUP_ORDER = ["People", "Document", "Dates", "Activity", "Custom"];

export default function ScrubPanel() {
  const [state, setState] = useState<State>({ phase: "idle" });

  const handleFile = useCallback(async (file: File) => {
    if (!isDocx(file)) {
      setState({ phase: "error", message: "Only .docx files are supported." });
      return;
    }
    setState({ phase: "reading", name: file.name });
    try {
      const buffer = await file.arrayBuffer();
      const fields = await readDocxMetadata(buffer);
      if (fields.length === 0) {
        setState({ phase: "none", name: file.name });
        return;
      }
      setState({
        phase: "review",
        base: baseName(file.name),
        buffer,
        fields,
        selected: new Set(fields.filter((f) => f.sensitive).map((f) => f.id)),
      });
    } catch (err) {
      setState({
        phase: "error",
        message: err instanceof Error ? err.message : "Something went wrong.",
      });
    }
  }, []);

  const scrub = useCallback(async () => {
    if (state.phase !== "review") return;
    const { base, buffer, selected } = state;
    setState({ phase: "processing", base, buffer, selected });
    try {
      const result = await scrubDocxMetadata(buffer, [...selected]);
      const outName = `${base}-clean.docx`;
      downloadBlob(result.data, outName, DOCX_MIME);
      setState({ phase: "done", name: outName, removed: result.removed });
    } catch (err) {
      setState({
        phase: "error",
        message: err instanceof Error ? err.message : "Something went wrong.",
      });
    }
  }, [state]);

  const setSelected = useCallback(
    (next: Set<string>) =>
      setState((s) => (s.phase === "review" ? { ...s, selected: next } : s)),
    []
  );

  const processing = state.phase === "reading" || state.phase === "processing";

  return (
    <>
      {state.phase !== "review" && (
        <Dropzone
          onFile={handleFile}
          disabled={processing}
          prompt={
            state.phase === "reading"
              ? "Reading metadata…"
              : "Drop a .docx file here"
          }
        />
      )}

      <div className="w-full max-w-xl text-center text-sm" aria-live="polite">
        {state.phase === "processing" && (
          <p className="text-gray-500 dark:text-gray-400">Scrubbing metadata…</p>
        )}
        {state.phase === "none" && (
          <div className="rounded-lg bg-amber-50 p-4 text-amber-800 dark:bg-amber-950/40 dark:text-amber-300">
            No metadata found in <span className="font-medium">{state.name}</span>.
          </div>
        )}
        {state.phase === "done" && (
          <div className="rounded-lg bg-green-50 p-4 text-green-800 dark:bg-green-950/40 dark:text-green-300">
            <p className="font-medium">Done — downloaded {state.name}</p>
            <p className="mt-1">
              Removed {state.removed} metadata field
              {state.removed === 1 ? "" : "s"}.
            </p>
          </div>
        )}
        {state.phase === "error" && (
          <div className="rounded-lg bg-red-50 p-4 text-red-800 dark:bg-red-950/40 dark:text-red-300">
            {state.message}
          </div>
        )}
      </div>

      {state.phase === "review" && (
        <Review
          fields={state.fields}
          selected={state.selected}
          onSelectedChange={setSelected}
          onScrub={scrub}
          onCancel={() => setState({ phase: "idle" })}
        />
      )}
    </>
  );
}

function Review({
  fields,
  selected,
  onSelectedChange,
  onScrub,
  onCancel,
}: {
  fields: MetadataField[];
  selected: Set<string>;
  onSelectedChange: (next: Set<string>) => void;
  onScrub: () => void;
  onCancel: () => void;
}) {
  const groups = useMemo(() => {
    const byGroup = new Map<string, MetadataField[]>();
    for (const f of fields) {
      const list = byGroup.get(f.group) ?? [];
      list.push(f);
      byGroup.set(f.group, list);
    }
    return [...byGroup.entries()].sort(
      (a, b) => GROUP_ORDER.indexOf(a[0]) - GROUP_ORDER.indexOf(b[0])
    );
  }, [fields]);

  const toggle = (id: string) => {
    const next = new Set(selected);
    if (next.has(id)) next.delete(id);
    else next.add(id);
    onSelectedChange(next);
  };

  return (
    <div className="flex w-full max-w-xl flex-col gap-4">
      <div className="flex flex-wrap items-center gap-x-4 gap-y-2 text-sm">
        <span className="text-gray-500 dark:text-gray-400">
          {selected.size} of {fields.length} selected
        </span>
        <div className="ml-auto flex gap-3">
          <button
            type="button"
            onClick={() => onSelectedChange(new Set(fields.map((f) => f.id)))}
            className="text-blue-600 hover:underline dark:text-blue-400"
          >
            Select all
          </button>
          <button
            type="button"
            onClick={() => onSelectedChange(new Set())}
            className="text-blue-600 hover:underline dark:text-blue-400"
          >
            Select none
          </button>
          <button
            type="button"
            onClick={() =>
              onSelectedChange(
                new Set(fields.filter((f) => f.sensitive).map((f) => f.id))
              )
            }
            className="text-blue-600 hover:underline dark:text-blue-400"
          >
            Recommended
          </button>
        </div>
      </div>

      <div className="flex flex-col gap-4">
        {groups.map(([group, items]) => (
          <fieldset
            key={group}
            className="rounded-lg border border-gray-200 dark:border-gray-700"
          >
            <legend className="ml-3 px-1 text-xs font-semibold uppercase tracking-wide text-gray-500 dark:text-gray-400">
              {group}
            </legend>
            <ul className="divide-y divide-gray-100 dark:divide-gray-800">
              {items.map((f) => (
                <li key={f.id}>
                  <label className="flex cursor-pointer items-start gap-3 px-3 py-2.5">
                    <input
                      type="checkbox"
                      checked={selected.has(f.id)}
                      onChange={() => toggle(f.id)}
                      className="mt-0.5 h-4 w-4 shrink-0"
                    />
                    <span className="min-w-0 flex-1">
                      <span className="block text-sm font-medium">{f.label}</span>
                      <span
                        className="block truncate text-xs text-gray-500 dark:text-gray-400"
                        title={f.value}
                      >
                        {f.value}
                      </span>
                    </span>
                  </label>
                </li>
              ))}
            </ul>
          </fieldset>
        ))}
      </div>

      <div className="flex items-center justify-end gap-3">
        <button
          type="button"
          onClick={onCancel}
          className="rounded-lg px-4 py-2 text-sm font-medium text-gray-600 transition-colors hover:bg-gray-100 dark:text-gray-300 dark:hover:bg-gray-800"
        >
          Use a different file
        </button>
        <button
          type="button"
          onClick={onScrub}
          disabled={selected.size === 0}
          className="rounded-lg bg-blue-600 px-4 py-2 text-sm font-medium text-white transition-colors hover:bg-blue-700 disabled:cursor-not-allowed disabled:opacity-50"
        >
          Remove selected &amp; download
        </button>
      </div>
    </div>
  );
}
