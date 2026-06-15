import type { ImageEncoder } from "@/lib/docx";

export interface CompressOptions {
  /** "auto" picks JPEG for opaque images and WebP when transparency is present. */
  format: "auto" | "jpeg" | "webp";
  /** Encoder quality, 0..1. */
  quality: number;
  /** Longest-edge cap in pixels; 0 disables downscaling. */
  maxDimension: number;
}

const MIME: Record<"jpeg" | "webp", string> = {
  jpeg: "image/jpeg",
  webp: "image/webp",
};

/** Scan the canvas for any non-opaque pixel. */
function hasAlpha(ctx: CanvasRenderingContext2D, w: number, h: number): boolean {
  const { data } = ctx.getImageData(0, 0, w, h);
  for (let i = 3; i < data.length; i += 4) {
    if (data[i] < 255) return true;
  }
  return false;
}

/**
 * Builds a browser {@link ImageEncoder} that re-encodes images with the canvas
 * API. Optionally downscales to a max longest edge, then encodes to JPEG or
 * WebP at the given quality. JPEG output flattens transparency onto white
 * (matching Word); WebP preserves the alpha channel.
 *
 * Returns `null` for images the browser can't decode or encode, so the docx
 * layer keeps the original bytes untouched.
 */
export function makeEncoder(opts: CompressOptions): ImageEncoder {
  return async (data, contentType) => {
    let bitmap: ImageBitmap;
    try {
      bitmap = await createImageBitmap(
        new Blob([data.slice().buffer], { type: contentType })
      );
    } catch {
      return null;
    }

    const { width, height } = bitmap;
    let tw = width;
    let th = height;
    if (opts.maxDimension > 0 && Math.max(width, height) > opts.maxDimension) {
      const scale = opts.maxDimension / Math.max(width, height);
      tw = Math.max(1, Math.round(width * scale));
      th = Math.max(1, Math.round(height * scale));
    }

    const canvas = document.createElement("canvas");
    canvas.width = tw;
    canvas.height = th;
    const ctx = canvas.getContext("2d");
    if (!ctx) {
      bitmap.close();
      return null;
    }

    let format: "jpeg" | "webp";
    if (opts.format === "auto") {
      // Draw once to inspect alpha; the drawn pixels are reused either way.
      ctx.drawImage(bitmap, 0, 0, tw, th);
      format = hasAlpha(ctx, tw, th) ? "webp" : "jpeg";
      if (format === "jpeg") {
        // JPEG has no alpha — paint white behind the (opaque) image.
        ctx.globalCompositeOperation = "destination-over";
        ctx.fillStyle = "#ffffff";
        ctx.fillRect(0, 0, tw, th);
        ctx.globalCompositeOperation = "source-over";
      }
    } else {
      format = opts.format;
      if (format === "jpeg") {
        ctx.fillStyle = "#ffffff";
        ctx.fillRect(0, 0, tw, th);
      }
      ctx.drawImage(bitmap, 0, 0, tw, th);
    }
    bitmap.close();

    const blob = await new Promise<Blob | null>((resolve) =>
      canvas.toBlob((b) => resolve(b), MIME[format], opts.quality)
    );
    // If the browser ignored the requested type (e.g. no WebP support), bail.
    if (!blob || blob.type !== MIME[format]) return null;

    return {
      data: new Uint8Array(await blob.arrayBuffer()),
      ext: format,
      contentType: MIME[format],
    };
  };
}
