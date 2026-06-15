"use client";

import { useCallback, useMemo, useRef, useState } from "react";
import {
  applyDocxMetadata,
  readDocxMetadata,
  METADATA_CATALOG,
  type MetadataField,
  type MetadataKind,
} from "@/lib/docx";
import Dropzone from "./Dropzone";
import { DOCX_MIME, baseName, downloadBlob, isDocx } from "./util";

interface Row {
  key: string;
  id: string;
  label: string;
  group: string;
  kind: MetadataKind;
  custom: boolean;
  sensitive: boolean;
  /** Present in the file when it was opened (vs. added in this session). */
  original: boolean;
  value: string;
  removed: boolean;
}

type State =
  | { phase: "idle" }
  | { phase: "reading"; name: string }
  | {
      phase: "editor";
      base: string;
      buffer: ArrayBuffer;
      originalIds: Set<string>;
      rows: Row[];
    }
  | { phase: "none"; name: string }
  | { phase: "processing" }
  | { phase: "done"; name: string; changed: number; removed: number }
  | { phase: "error"; message: string };

const GROUP_ORDER = ["People", "Document", "Dates", "Activity", "Custom"];

const PLACEHOLDER: Record<MetadataKind, string> = {
  text: "value",
  number: "0",
  date: "YYYY-MM-DDThh:mm:ssZ",
};

function rowFromField(f: MetadataField): Row {
  return {
    key: f.id,
    id: f.id,
    label: f.label,
    group: f.group,
    kind: f.kind,
    custom: f.custom,
    sensitive: f.sensitive,
    original: true,
    value: f.value,
    removed: false,
  };
}

export default function ScrubPanel() {
  const [state, setState] = useState<State>({ phase: "idle" });
  const seq = useRef(0);

  const updateRows = useCallback(
    (f: (rows: Row[]) => Row[]) =>
      setState((s) => (s.phase === "editor" ? { ...s, rows: f(s.rows) } : s)),
    []
  );

  const handleFile = useCallback(async (file: File) => {
    if (!isDocx(file)) {
      setState({ phase: "error", message: "Only .docx files are supported." });
      return;
    }
    setState({ phase: "reading", name: file.name });
    try {
      const buffer = await file.arrayBuffer();
      const fields = await readDocxMetadata(buffer);
      setState({
        phase: "editor",
        base: baseName(file.name),
        buffer,
        originalIds: new Set(fields.map((f) => f.id)),
        rows: fields.map(rowFromField),
      });
    } catch (err) {
      setState({
        phase: "error",
        message: err instanceof Error ? err.message : "Something went wrong.",
      });
    }
  }, []);

  const addField = useCallback(
    (choice: string) => {
      if (!choice) return;
      if (choice === "custom") {
        updateRows((rows) => [
          ...rows,
          {
            key: `new-${seq.current++}`,
            id: "",
            label: "",
            group: "Custom",
            kind: "text",
            custom: true,
            sensitive: false,
            original: false,
            value: "",
            removed: false,
          },
        ]);
        return;
      }
      const def = METADATA_CATALOG.find((d) => d.id === choice);
      if (!def) return;
      updateRows((rows) => {
        // Re-adding a field that was deleted just clears the deletion.
        const existing = rows.find((r) => r.id === def.id);
        if (existing) {
          return rows.map((r) =>
            r.id === def.id ? { ...r, removed: false } : r
          );
        }
        return [
          ...rows,
          {
            key: `new-${seq.current++}`,
            id: def.id,
            label: def.label,
            group: def.group,
            kind: def.kind,
            custom: false,
            sensitive: false,
            original: false,
            value: "",
            removed: false,
          },
        ];
      });
    },
    [updateRows]
  );

  const apply = useCallback(async () => {
    if (state.phase !== "editor") return;
    const { base, buffer, originalIds, rows } = state;

    const set: Record<string, string> = {};
    for (const r of rows) {
      if (r.removed) continue;
      const name = r.custom ? r.label.trim() : "";
      if (r.custom && !name) continue;
      if (r.value.trim() === "") continue;
      set[r.custom ? `custom:${name}` : r.id] = r.value;
    }
    const finalIds = new Set(Object.keys(set));
    const remove = [...originalIds].filter((id) => !finalIds.has(id));

    setState({ phase: "processing" });
    try {
      const result = await applyDocxMetadata(buffer, { set, remove });
      const outName = `${base}-edited.docx`;
      downloadBlob(result.data, outName, DOCX_MIME);
      setState({
        phase: "done",
        name: outName,
        changed: result.changed,
        removed: result.removed,
      });
    } catch (err) {
      setState({
        phase: "error",
        message: err instanceof Error ? err.message : "Something went wrong.",
      });
    }
  }, [state]);

  const processing = state.phase === "reading" || state.phase === "processing";

  return (
    <>
      {state.phase !== "editor" && (
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
          <p className="text-gray-500 dark:text-gray-400">Saving metadata…</p>
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
              Wrote {state.changed} field{state.changed === 1 ? "" : "s"} and
              removed {state.removed} field{state.removed === 1 ? "" : "s"}.
            </p>
          </div>
        )}
        {state.phase === "error" && (
          <div className="rounded-lg bg-red-50 p-4 text-red-800 dark:bg-red-950/40 dark:text-red-300">
            {state.message}
          </div>
        )}
      </div>

      {state.phase === "editor" && (
        <Editor
          rows={state.rows}
          onChange={updateRows}
          onAddField={addField}
          onApply={apply}
          onCancel={() => setState({ phase: "idle" })}
        />
      )}
    </>
  );
}

function Editor({
  rows,
  onChange,
  onAddField,
  onApply,
  onCancel,
}: {
  rows: Row[];
  onChange: (f: (rows: Row[]) => Row[]) => void;
  onAddField: (choice: string) => void;
  onApply: () => void;
  onCancel: () => void;
}) {
  const present = useMemo(
    () => new Set(rows.filter((r) => !r.removed && !r.custom).map((r) => r.id)),
    [rows]
  );
  const addable = useMemo(
    () => METADATA_CATALOG.filter((d) => !present.has(d.id)),
    [present]
  );

  const groups = useMemo(() => {
    const byGroup = new Map<string, Row[]>();
    for (const r of rows) {
      const list = byGroup.get(r.group) ?? [];
      list.push(r);
      byGroup.set(r.group, list);
    }
    return [...byGroup.entries()].sort(
      (a, b) => GROUP_ORDER.indexOf(a[0]) - GROUP_ORDER.indexOf(b[0])
    );
  }, [rows]);

  const patch = (key: string, change: Partial<Row>) =>
    onChange((rs) => rs.map((r) => (r.key === key ? { ...r, ...change } : r)));

  const removeRow = (key: string) =>
    onChange((rs) =>
      rs
        // Drop session-added rows entirely; just flag original ones.
        .filter((r) => r.key !== key || r.original)
        .map((r) => (r.key === key ? { ...r, removed: true } : r))
    );

  const hasSensitive = rows.some((r) => r.sensitive && !r.removed);

  return (
    <div className="flex w-full max-w-xl flex-col gap-4">
      <div className="flex flex-wrap items-center gap-3 text-sm">
        <label className="flex items-center gap-2">
          <span className="text-gray-500 dark:text-gray-400">Add field</span>
          <select
            value=""
            onChange={(e) => onAddField(e.target.value)}
            className="rounded-lg border border-gray-300 bg-transparent px-2 py-1.5 dark:border-gray-700"
          >
            <option value="" disabled>
              Choose…
            </option>
            {GROUP_ORDER.filter((g) =>
              addable.some((d) => d.group === g)
            ).map((g) => (
              <optgroup key={g} label={g}>
                {addable
                  .filter((d) => d.group === g)
                  .map((d) => (
                    <option key={d.id} value={d.id}>
                      {d.label}
                    </option>
                  ))}
              </optgroup>
            ))}
            <option value="custom">Custom property…</option>
          </select>
        </label>
        {hasSensitive && (
          <button
            type="button"
            onClick={() =>
              onChange((rs) =>
                rs
                  .filter((r) => !r.sensitive || r.original)
                  .map((r) => (r.sensitive ? { ...r, removed: true } : r))
              )
            }
            className="ml-auto text-blue-600 hover:underline dark:text-blue-400"
          >
            Remove all sensitive
          </button>
        )}
      </div>

      {rows.length === 0 && (
        <p className="rounded-lg border border-dashed border-gray-300 p-6 text-center text-sm text-gray-500 dark:border-gray-700 dark:text-gray-400">
          No metadata. Add a field to write one in.
        </p>
      )}

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
              {items.map((r) => (
                <li key={r.key} className="px-3 py-2.5">
                  {r.removed ? (
                    <div className="flex items-center justify-between text-sm">
                      <span className="text-gray-400 line-through">
                        {r.label || "Custom property"}
                      </span>
                      <button
                        type="button"
                        onClick={() => patch(r.key, { removed: false })}
                        className="text-blue-600 hover:underline dark:text-blue-400"
                      >
                        Undo
                      </button>
                    </div>
                  ) : (
                    <div className="flex items-center gap-2">
                      {r.custom ? (
                        <input
                          value={r.label}
                          onChange={(e) => patch(r.key, { label: e.target.value })}
                          placeholder="Property name"
                          aria-label="Property name"
                          className="w-32 shrink-0 rounded-md border border-gray-300 bg-transparent px-2 py-1.5 text-sm font-medium dark:border-gray-700"
                        />
                      ) : (
                        <span className="w-32 shrink-0 text-sm font-medium">
                          {r.label}
                        </span>
                      )}
                      <input
                        value={r.value}
                        onChange={(e) => patch(r.key, { value: e.target.value })}
                        placeholder={PLACEHOLDER[r.kind]}
                        inputMode={r.kind === "number" ? "numeric" : "text"}
                        aria-label={`${r.label || "Custom property"} value`}
                        className="min-w-0 flex-1 rounded-md border border-gray-300 bg-transparent px-2 py-1.5 text-sm dark:border-gray-700"
                      />
                      <button
                        type="button"
                        onClick={() => removeRow(r.key)}
                        aria-label={`Remove ${r.label || "field"}`}
                        className="shrink-0 rounded-md px-2 py-1 text-gray-400 transition-colors hover:bg-gray-100 hover:text-red-600 dark:hover:bg-gray-800"
                      >
                        ✕
                      </button>
                    </div>
                  )}
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
          onClick={onApply}
          className="rounded-lg bg-blue-600 px-4 py-2 text-sm font-medium text-white transition-colors hover:bg-blue-700"
        >
          Save &amp; download
        </button>
      </div>
    </div>
  );
}
