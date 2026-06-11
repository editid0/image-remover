import { NextRequest } from "next/server";
import { removeImagesFromDocx } from "@/lib/docx";

export const runtime = "nodejs";

const DOCX_MIME =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

export async function POST(req: NextRequest) {
  let file: FormDataEntryValue | null;
  try {
    file = (await req.formData()).get("file");
  } catch {
    return Response.json({ error: "Invalid form data." }, { status: 400 });
  }

  if (!(file instanceof File)) {
    return Response.json({ error: "No file uploaded." }, { status: 400 });
  }
  if (!file.name.toLowerCase().endsWith(".docx")) {
    return Response.json(
      { error: "Only .docx files are supported." },
      { status: 400 }
    );
  }

  try {
    const result = await removeImagesFromDocx(await file.arrayBuffer());
    const base = file.name.replace(/\.docx$/i, "");
    const safeName = `${base}-no-images.docx`.replace(/["\\\r\n]/g, "_");

    return new Response(new Blob([new Uint8Array(result.buffer)]), {
      headers: {
        "Content-Type": DOCX_MIME,
        "Content-Disposition": `attachment; filename="${safeName}"`,
        "X-Images-Removed": String(result.removedElements),
        "X-Media-Removed": String(result.removedMediaFiles),
      },
    });
  } catch (err) {
    const message =
      err instanceof Error ? err.message : "Failed to process the file.";
    return Response.json({ error: message }, { status: 422 });
  }
}
