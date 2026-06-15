"use client";

import { useCallback, useState } from "react";
import { removeImagesFromDocx } from "@/lib/docx";
import Dropzone from "./Dropzone";
import { DOCX_MIME, baseName, downloadBlob, isDocx } from "./util";

type Status =
  | { state: "idle" }
  | { state: "processing"; name: string }
  | { state: "done"; name: string; images: number; media: number }
  | { state: "error"; message: string };

export default function RemovePanel() {
  const [status, setStatus] = useState<Status>({ state: "idle" });

  const handleFile = useCallback(async (file: File) => {
    if (!isDocx(file)) {
      setStatus({ state: "error", message: "Only .docx files are supported." });
      return;
    }
    setStatus({ state: "processing", name: file.name });
    try {
      const result = await removeImagesFromDocx(await file.arrayBuffer());
      const outName = baseName(file.name) + "-no-images.docx";
      downloadBlob(result.data, outName, DOCX_MIME);
      setStatus({
        state: "done",
        name: outName,
        images: result.removedElements,
        media: result.removedMediaFiles,
      });
    } catch (err) {
      setStatus({
        state: "error",
        message: err instanceof Error ? err.message : "Something went wrong.",
      });
    }
  }, []);

  const processing = status.state === "processing";

  return (
    <>
      <Dropzone
        onFile={handleFile}
        disabled={processing}
        prompt={processing ? "Processing…" : "Drop a .docx file here"}
      />
      <div className="min-h-12 w-full max-w-xl text-center text-sm" aria-live="polite">
        {status.state === "processing" && (
          <p className="text-gray-500 dark:text-gray-400">
            Removing images from <span className="font-medium">{status.name}</span>…
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
    </>
  );
}
