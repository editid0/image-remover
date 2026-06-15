"use client";

import { useCallback, useEffect, useRef, useState } from "react";
import {
  extractImagesFromDocx,
  zipImages,
  type ExtractedImage,
} from "@/lib/docx";
import Dropzone from "./Dropzone";
import { baseName, downloadBlob, isDocx } from "./util";

interface Preview extends ExtractedImage {
  url: string;
}

type State =
  | { phase: "idle" }
  | { phase: "processing"; name: string }
  | { phase: "done"; base: string; images: Preview[] }
  | { phase: "empty"; name: string }
  | { phase: "error"; message: string };

// Types most browsers can render in an <img>; others get a placeholder tile.
const RENDERABLE = new Set([
  "image/png",
  "image/jpeg",
  "image/gif",
  "image/bmp",
  "image/webp",
  "image/svg+xml",
  "image/x-icon",
]);

export default function ExtractPanel() {
  const [state, setState] = useState<State>({ phase: "idle" });
  const urlsRef = useRef<string[]>([]);

  const revoke = useCallback(() => {
    urlsRef.current.forEach((u) => URL.revokeObjectURL(u));
    urlsRef.current = [];
  }, []);
  // Release object URLs when the panel unmounts (e.g. switching modes).
  useEffect(() => revoke, [revoke]);

  const handleFile = useCallback(
    async (file: File) => {
      if (!isDocx(file)) {
        setState({ phase: "error", message: "Only .docx files are supported." });
        return;
      }
      revoke();
      setState({ phase: "processing", name: file.name });
      try {
        const images = await extractImagesFromDocx(await file.arrayBuffer());
        if (images.length === 0) {
          setState({ phase: "empty", name: file.name });
          return;
        }
        const previews: Preview[] = images.map((img) => {
          const url = URL.createObjectURL(
            new Blob([img.data.slice().buffer], { type: img.contentType })
          );
          urlsRef.current.push(url);
          return { ...img, url };
        });
        setState({ phase: "done", base: baseName(file.name), images: previews });
      } catch (err) {
        setState({
          phase: "error",
          message: err instanceof Error ? err.message : "Something went wrong.",
        });
      }
    },
    [revoke]
  );

  const downloadAll = useCallback(async (base: string, images: Preview[]) => {
    const zip = await zipImages(images);
    downloadBlob(zip, `${base}-images.zip`, "application/zip");
  }, []);

  const processing = state.phase === "processing";

  return (
    <>
      <Dropzone
        onFile={handleFile}
        disabled={processing}
        prompt={processing ? "Extracting…" : "Drop a .docx file here"}
      />

      <div className="w-full max-w-xl text-center text-sm" aria-live="polite">
        {state.phase === "processing" && (
          <p className="text-gray-500 dark:text-gray-400">
            Extracting images from <span className="font-medium">{state.name}</span>…
          </p>
        )}
        {state.phase === "empty" && (
          <div className="rounded-lg bg-amber-50 p-4 text-amber-800 dark:bg-amber-950/40 dark:text-amber-300">
            No images found in <span className="font-medium">{state.name}</span>.
          </div>
        )}
        {state.phase === "error" && (
          <div className="rounded-lg bg-red-50 p-4 text-red-800 dark:bg-red-950/40 dark:text-red-300">
            {state.message}
          </div>
        )}
      </div>

      {state.phase === "done" && (
        <div className="flex w-full max-w-xl flex-col gap-4">
          <div className="flex items-center justify-between">
            <p className="text-sm text-gray-500 dark:text-gray-400">
              Found {state.images.length} image
              {state.images.length === 1 ? "" : "s"}.
            </p>
            <button
              type="button"
              onClick={() => downloadAll(state.base, state.images)}
              className="rounded-lg bg-blue-600 px-4 py-2 text-sm font-medium text-white transition-colors hover:bg-blue-700"
            >
              Download all (.zip)
            </button>
          </div>

          <ul className="grid grid-cols-2 gap-3 sm:grid-cols-3">
            {state.images.map((img) => (
              <li
                key={img.name}
                className="flex flex-col overflow-hidden rounded-lg border border-gray-200 dark:border-gray-700"
              >
                <div className="flex h-28 items-center justify-center bg-gray-50 p-2 dark:bg-gray-900">
                  {RENDERABLE.has(img.contentType) ? (
                    // eslint-disable-next-line @next/next/no-img-element
                    <img
                      src={img.url}
                      alt={img.name}
                      className="max-h-full max-w-full object-contain"
                    />
                  ) : (
                    <span className="text-xs font-medium uppercase text-gray-400">
                      {img.name.split(".").pop()}
                    </span>
                  )}
                </div>
                <button
                  type="button"
                  onClick={() => downloadBlob(img.data, img.name, img.contentType)}
                  className="truncate border-t border-gray-200 px-2 py-1.5 text-xs text-gray-600 transition-colors hover:bg-gray-50 dark:border-gray-700 dark:text-gray-300 dark:hover:bg-gray-900"
                  title={`Download ${img.name}`}
                >
                  {img.name}
                </button>
              </li>
            ))}
          </ul>
        </div>
      )}
    </>
  );
}
