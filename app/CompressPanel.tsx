"use client";

import { useCallback, useState } from "react";
import { compressImagesInDocx } from "@/lib/docx";
import { makeEncoder, type CompressOptions } from "./encodeImage";
import Dropzone from "./Dropzone";
import { DOCX_MIME, baseName, downloadBlob, formatBytes, isDocx } from "./util";

type Status =
  | { state: "idle" }
  | { state: "processing"; name: string }
  | {
      state: "done";
      name: string;
      totalImages: number;
      recompressed: number;
      originalBytes: number;
      compressedBytes: number;
    }
  | { state: "error"; message: string };

const RESOLUTIONS = [
  { label: "Original", value: 0 },
  { label: "2560 px", value: 2560 },
  { label: "1920 px", value: 1920 },
  { label: "1280 px", value: 1280 },
];

function savedPercent(original: number, compressed: number): number {
  if (original <= 0) return 0;
  return Math.round((1 - compressed / original) * 100);
}

export default function CompressPanel() {
  const [status, setStatus] = useState<Status>({ state: "idle" });
  const [quality, setQuality] = useState(80);
  const [format, setFormat] = useState<CompressOptions["format"]>("auto");
  const [maxDimension, setMaxDimension] = useState(0);

  const handleFile = useCallback(
    async (file: File) => {
      if (!isDocx(file)) {
        setStatus({ state: "error", message: "Only .docx files are supported." });
        return;
      }
      setStatus({ state: "processing", name: file.name });
      try {
        const encode = makeEncoder({
          format,
          quality: quality / 100,
          maxDimension,
        });
        const result = await compressImagesInDocx(await file.arrayBuffer(), encode);
        const outName = baseName(file.name) + "-compressed.docx";
        downloadBlob(result.data, outName, DOCX_MIME);
        setStatus({
          state: "done",
          name: outName,
          totalImages: result.totalImages,
          recompressed: result.recompressed,
          originalBytes: result.originalBytes,
          compressedBytes: result.compressedBytes,
        });
      } catch (err) {
        setStatus({
          state: "error",
          message: err instanceof Error ? err.message : "Something went wrong.",
        });
      }
    },
    [format, quality, maxDimension]
  );

  const processing = status.state === "processing";

  return (
    <>
      <div className="flex w-full max-w-xl flex-col gap-4 rounded-xl border border-gray-200 p-4 dark:border-gray-700">
        <label className="flex flex-col gap-1 text-sm">
          <span className="flex justify-between font-medium">
            <span>Quality</span>
            <span className="text-gray-500 dark:text-gray-400">{quality}%</span>
          </span>
          <input
            type="range"
            min={40}
            max={95}
            step={5}
            value={quality}
            disabled={processing}
            onChange={(e) => setQuality(Number(e.target.value))}
            className="accent-blue-600"
          />
        </label>

        <div className="grid grid-cols-2 gap-4">
          <label className="flex flex-col gap-1 text-sm">
            <span className="font-medium">Format</span>
            <select
              value={format}
              disabled={processing}
              onChange={(e) =>
                setFormat(e.target.value as CompressOptions["format"])
              }
              className="rounded-lg border border-gray-300 bg-white px-2 py-1.5 dark:border-gray-700 dark:bg-gray-800"
            >
              <option value="auto">Auto (recommended)</option>
              <option value="jpeg">JPEG (most compatible)</option>
              <option value="webp">WebP (smallest)</option>
            </select>
          </label>

          <label className="flex flex-col gap-1 text-sm">
            <span className="font-medium">Max resolution</span>
            <select
              value={maxDimension}
              disabled={processing}
              onChange={(e) => setMaxDimension(Number(e.target.value))}
              className="rounded-lg border border-gray-300 bg-white px-2 py-1.5 dark:border-gray-700 dark:bg-gray-800"
            >
              {RESOLUTIONS.map((r) => (
                <option key={r.value} value={r.value}>
                  {r.label}
                </option>
              ))}
            </select>
          </label>
        </div>

        <p className="text-xs text-gray-500 dark:text-gray-400">
          {format === "webp"
            ? "WebP makes the smallest files but only opens in Word 2021/365 and later."
            : "Auto keeps transparent images crisp (WebP) and shrinks photos hard (JPEG)."}
        </p>
      </div>

      <Dropzone
        onFile={handleFile}
        disabled={processing}
        prompt={processing ? "Compressing…" : "Drop a .docx file here"}
      />

      <div
        className="min-h-12 w-full max-w-xl text-center text-sm"
        aria-live="polite"
      >
        {status.state === "processing" && (
          <p className="text-gray-500 dark:text-gray-400">
            Compressing images in{" "}
            <span className="font-medium">{status.name}</span>…
          </p>
        )}
        {status.state === "done" && (
          <div className="rounded-lg bg-green-50 p-4 text-green-800 dark:bg-green-950/40 dark:text-green-300">
            {status.recompressed === 0 ? (
              <p className="font-medium">
                {status.totalImages === 0
                  ? "No compressible images found — downloaded an unchanged copy."
                  : "Images were already optimally compressed — nothing to shrink."}
              </p>
            ) : (
              <>
                <p className="font-medium">Done — downloaded {status.name}</p>
                <p className="mt-1">
                  Compressed {status.recompressed} of {status.totalImages} image
                  {status.totalImages === 1 ? "" : "s"} —{" "}
                  {formatBytes(status.originalBytes)} →{" "}
                  {formatBytes(status.compressedBytes)} (
                  {savedPercent(status.originalBytes, status.compressedBytes)}%
                  smaller).
                </p>
              </>
            )}
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
