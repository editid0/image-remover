import JSZip from "jszip";
import { DOMParser, XMLSerializer } from "@xmldom/xmldom";
import type {
  Document as XmlDocument,
  Element as XmlElement,
  Node as XmlNode,
} from "@xmldom/xmldom";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const V_NS = "urn:schemas-microsoft-com:vml";
const REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
const IMAGE_REL_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

export interface RemoveImagesResult {
  data: Uint8Array;
  /** Number of image elements stripped from the document XML. */
  removedElements: number;
  /** Number of media files deleted from the package. */
  removedMediaFiles: number;
}

interface ArrayLikeList<T> {
  length: number;
  item(index: number): T | null;
}

function toArray<T>(list: ArrayLikeList<T>): T[] {
  const out: T[] = [];
  for (let i = 0; i < list.length; i++) {
    const node = list.item(i);
    if (node) out.push(node);
  }
  return out;
}

function isAttached(el: XmlNode, doc: XmlDocument): boolean {
  let node: XmlNode | null = el;
  while (node.parentNode) node = node.parentNode;
  return node === doc;
}

/** Resolve a relationship Target against the directory its .rels file describes. */
function resolveTarget(baseDir: string, target: string): string {
  const raw = target.startsWith("/") ? target.slice(1) : baseDir + target;
  const out: string[] = [];
  for (const seg of raw.split("/")) {
    if (seg === "" || seg === ".") continue;
    if (seg === "..") out.pop();
    else out.push(seg);
  }
  return out.join("/");
}

/**
 * Removes all images from a .docx file.
 *
 * - Strips `w:drawing` (DrawingML) elements and `w:pict` (VML) elements that
 *   contain image data from the document body, headers, footers, footnotes
 *   and endnotes. VML text boxes without images are left intact.
 * - Drops image relationships that are no longer referenced.
 * - Deletes media files in `word/media/` that nothing references anymore.
 */
export async function removeImagesFromDocx(
  input: ArrayBuffer | Uint8Array
): Promise<RemoveImagesResult> {
  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(input);
  } catch {
    throw new Error("Not a valid .docx file (could not be read as a zip archive).");
  }
  if (!zip.file("word/document.xml")) {
    throw new Error("Not a valid .docx file (missing word/document.xml).");
  }

  const parser = new DOMParser();
  const serializer = new XMLSerializer();
  let removedElements = 0;

  const storyParts = Object.keys(zip.files).filter((name) =>
    /^word\/(document|header\d*|footer\d*|footnotes|endnotes)\.xml$/.test(name)
  );

  for (const partName of storyParts) {
    const xml = await zip.file(partName)!.async("string");
    const doc = parser.parseFromString(xml, "application/xml");

    const targets: XmlElement[] = toArray(
      doc.getElementsByTagNameNS(W_NS, "drawing")
    );
    for (const pict of toArray(doc.getElementsByTagNameNS(W_NS, "pict"))) {
      if (pict.getElementsByTagNameNS(V_NS, "imagedata").length > 0) {
        targets.push(pict);
      }
    }
    if (targets.length === 0) continue;

    for (const el of targets) {
      // an element nested inside an already-removed one is detached by now
      if (el.parentNode && isAttached(el, doc)) {
        el.parentNode.removeChild(el);
        removedElements++;
      }
    }

    const newXml = serializer.serializeToString(doc);
    zip.file(partName, newXml);

    // Drop image relationships this part no longer references.
    const relsName = partName.replace(/^word\//, "word/_rels/") + ".rels";
    const relsFile = zip.file(relsName);
    if (!relsFile) continue;
    const relsDoc = parser.parseFromString(
      await relsFile.async("string"),
      "application/xml"
    );
    let relsChanged = false;
    for (const rel of toArray(
      relsDoc.getElementsByTagNameNS(REL_NS, "Relationship")
    )) {
      const id = rel.getAttribute("Id");
      if (
        rel.getAttribute("Type") === IMAGE_REL_TYPE &&
        rel.getAttribute("TargetMode") !== "External" &&
        id &&
        !newXml.includes(`"${id}"`)
      ) {
        rel.parentNode!.removeChild(rel);
        relsChanged = true;
      }
    }
    if (relsChanged) {
      zip.file(relsName, serializer.serializeToString(relsDoc));
    }
  }

  // Delete media files nothing references anymore. Targets still referenced
  // elsewhere (e.g. picture bullets in numbering.xml) are kept.
  const referenced = new Set<string>();
  for (const name of Object.keys(zip.files).filter((n) => n.endsWith(".rels"))) {
    const relsDoc = parser.parseFromString(
      await zip.file(name)!.async("string"),
      "application/xml"
    );
    const baseDir = name.replace(/_rels\/[^/]*$/, "");
    for (const rel of toArray(
      relsDoc.getElementsByTagNameNS(REL_NS, "Relationship")
    )) {
      if (rel.getAttribute("TargetMode") === "External") continue;
      const target = rel.getAttribute("Target");
      if (target) referenced.add(resolveTarget(baseDir, target));
    }
  }

  let removedMediaFiles = 0;
  for (const name of Object.keys(zip.files)) {
    if (
      name.startsWith("word/media/") &&
      !zip.files[name].dir &&
      !referenced.has(name)
    ) {
      zip.remove(name);
      removedMediaFiles++;
    }
  }

  const data = await zip.generateAsync({
    type: "uint8array",
    compression: "DEFLATE",
  });

  return { data, removedElements, removedMediaFiles };
}

/* -------------------------------------------------------------------------- */
/* Image extraction                                                           */
/* -------------------------------------------------------------------------- */

export interface ExtractedImage {
  /** File name without directory, e.g. "image1.png". */
  name: string;
  data: Uint8Array;
  /** MIME type inferred from the file extension. */
  contentType: string;
}

const IMAGE_CONTENT_TYPES: Record<string, string> = {
  png: "image/png",
  jpg: "image/jpeg",
  jpeg: "image/jpeg",
  jpe: "image/jpeg",
  gif: "image/gif",
  bmp: "image/bmp",
  tif: "image/tiff",
  tiff: "image/tiff",
  webp: "image/webp",
  svg: "image/svg+xml",
  ico: "image/x-icon",
  emf: "image/emf",
  wmf: "image/wmf",
};

/**
 * Extracts every embedded image from a .docx file. Images live in
 * `word/media/`; entries are returned sorted by name. The document itself is
 * not modified — this is the inverse of {@link removeImagesFromDocx}.
 */
export async function extractImagesFromDocx(
  input: ArrayBuffer | Uint8Array
): Promise<ExtractedImage[]> {
  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(input);
  } catch {
    throw new Error("Not a valid .docx file (could not be read as a zip archive).");
  }
  if (!zip.file("word/document.xml")) {
    throw new Error("Not a valid .docx file (missing word/document.xml).");
  }

  const images: ExtractedImage[] = [];
  for (const path of Object.keys(zip.files).sort()) {
    const entry = zip.files[path];
    if (entry.dir || !path.startsWith("word/media/")) continue;
    const name = path.slice("word/media/".length);
    const ext = name.includes(".") ? name.split(".").pop()!.toLowerCase() : "";
    const contentType = IMAGE_CONTENT_TYPES[ext];
    if (!contentType) continue;
    images.push({ name, data: await entry.async("uint8array"), contentType });
  }
  return images;
}

/** Bundle extracted images into a single zip archive for download. */
export async function zipImages(images: ExtractedImage[]): Promise<Uint8Array> {
  const zip = new JSZip();
  for (const img of images) zip.file(img.name, img.data);
  return zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
}

/* -------------------------------------------------------------------------- */
/* Image compression                                                          */
/* -------------------------------------------------------------------------- */

/**
 * Raster formats the browser canvas can decode and therefore re-encode. Vector
 * and exotic formats (emf, wmf, tiff, svg, ico) and animated gifs are left
 * untouched so nothing is corrupted.
 */
const COMPRESSIBLE = new Set(["png", "jpg", "jpeg", "jpe", "bmp", "webp"]);

/** A re-encoded image returned by an {@link ImageEncoder}. */
export interface EncodedImage {
  data: Uint8Array;
  /** Lower-case file extension for the new bytes, e.g. "jpeg" or "webp". */
  ext: string;
  contentType: string;
}

/**
 * Re-encodes a single image. Receives the original bytes, the source MIME type
 * and the media file name; returns the re-encoded bytes (with their new
 * extension/type) or `null` to leave the image unchanged. The browser
 * implementation uses the canvas API ({@link "../app/encodeImage".makeEncoder});
 * Node callers (tests) inject their own. Returning larger bytes is harmless —
 * {@link compressImagesInDocx} keeps the original whenever the result is not
 * strictly smaller.
 */
export type ImageEncoder = (
  data: Uint8Array,
  contentType: string,
  name: string
) => Promise<EncodedImage | null>;

export interface CompressImagesResult {
  data: Uint8Array;
  /** Compressible images found in `word/media/`. */
  totalImages: number;
  /** Images actually replaced with smaller bytes. */
  recompressed: number;
  /** Combined size of the compressible images before compression. */
  originalBytes: number;
  /** Combined size after compression (originals kept where not smaller). */
  compressedBytes: number;
}

/** Express an absolute package path relative to a .rels file's base directory. */
function relativeTarget(baseDir: string, fullPath: string): string {
  const base = baseDir.split("/").filter(Boolean);
  const target = fullPath.split("/").filter(Boolean);
  let i = 0;
  while (i < base.length && i < target.length && base[i] === target[i]) i++;
  const up = base.slice(i).map(() => "..");
  return [...up, ...target.slice(i)].join("/");
}

/** Ensure `[Content_Types].xml` has a Default mapping for each new extension. */
async function ensureDefaultContentTypes(
  zip: JSZip,
  parser: DOMParser,
  serializer: XMLSerializer,
  exts: Set<string>
): Promise<void> {
  if (exts.size === 0) return;
  const ctFile = zip.file("[Content_Types].xml");
  if (!ctFile) return;
  const ctDoc = parser.parseFromString(await ctFile.async("string"), "application/xml");
  const have = new Set(
    toArray(ctDoc.getElementsByTagNameNS(CT_NS, "Default")).map((d) =>
      (d.getAttribute("Extension") ?? "").toLowerCase()
    )
  );
  let changed = false;
  for (const ext of exts) {
    if (have.has(ext)) continue;
    const ct = IMAGE_CONTENT_TYPES[ext];
    if (!ct) continue;
    const def = ctDoc.createElementNS(CT_NS, "Default");
    def.setAttribute("Extension", ext);
    def.setAttribute("ContentType", ct);
    ctDoc.documentElement!.appendChild(def);
    have.add(ext);
    changed = true;
  }
  if (changed) {
    zip.file("[Content_Types].xml", serializer.serializeToString(ctDoc));
  }
}

/**
 * Re-encodes every image in a .docx through the supplied {@link ImageEncoder}
 * to shrink the file. Beats Word's "Compress Pictures" because the encoder can
 * use modern codecs (WebP), a tunable quality and an optional resolution cap.
 *
 * - Only raster formats the encoder can decode are touched ({@link COMPRESSIBLE});
 *   everything else is left as-is.
 * - An image is replaced only when the new bytes are strictly smaller, so the
 *   output is never larger than the input.
 * - When the format changes (e.g. PNG → JPEG), the media file is renamed and the
 *   matching relationship `Target`s and `[Content_Types].xml` defaults are
 *   updated. The document XML refers to images by relationship id, so it is left
 *   untouched.
 */
export async function compressImagesInDocx(
  input: ArrayBuffer | Uint8Array,
  encode: ImageEncoder
): Promise<CompressImagesResult> {
  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(input);
  } catch {
    throw new Error("Not a valid .docx file (could not be read as a zip archive).");
  }
  if (!zip.file("word/document.xml")) {
    throw new Error("Not a valid .docx file (missing word/document.xml).");
  }

  let totalImages = 0;
  let recompressed = 0;
  let originalBytes = 0;
  let compressedBytes = 0;
  // old media path -> new media path, for entries whose extension changed.
  const renames = new Map<string, string>();
  const newExts = new Set<string>();

  const mediaPaths = Object.keys(zip.files).filter(
    (p) => p.startsWith("word/media/") && !zip.files[p].dir
  );

  for (const path of mediaPaths) {
    const name = path.slice("word/media/".length);
    const ext = name.includes(".") ? name.split(".").pop()!.toLowerCase() : "";
    if (!COMPRESSIBLE.has(ext)) continue;
    const contentType = IMAGE_CONTENT_TYPES[ext];
    if (!contentType) continue;

    totalImages++;
    const original = await zip.files[path].async("uint8array");
    originalBytes += original.length;

    let encoded: EncodedImage | null = null;
    try {
      encoded = await encode(original, contentType, name);
    } catch {
      encoded = null;
    }

    // Never grow a file: keep the original unless the result is strictly smaller.
    if (!encoded || encoded.data.length >= original.length) {
      compressedBytes += original.length;
      continue;
    }

    recompressed++;
    compressedBytes += encoded.data.length;

    const newExt = encoded.ext.toLowerCase();
    if (newExt === ext) {
      zip.file(path, encoded.data);
    } else {
      const stem = name.includes(".") ? name.slice(0, name.lastIndexOf(".")) : name;
      let newName = `${stem}.${newExt}`;
      let i = 1;
      // Avoid colliding with an existing, different media entry.
      while (zip.file(`word/media/${newName}`) && `word/media/${newName}` !== path) {
        newName = `${stem}-${i++}.${newExt}`;
      }
      const newPath = `word/media/${newName}`;
      zip.remove(path);
      zip.file(newPath, encoded.data);
      renames.set(path, newPath);
      newExts.add(newExt);
    }
  }

  // Repoint relationships at renamed media and register any new extensions.
  if (renames.size > 0) {
    const parser = new DOMParser();
    const serializer = new XMLSerializer();

    for (const relsName of Object.keys(zip.files).filter((n) => n.endsWith(".rels"))) {
      const relsDoc = parser.parseFromString(
        await zip.file(relsName)!.async("string"),
        "application/xml"
      );
      const baseDir = relsName.replace(/_rels\/[^/]*$/, "");
      let relsChanged = false;
      for (const rel of toArray(
        relsDoc.getElementsByTagNameNS(REL_NS, "Relationship")
      )) {
        if (rel.getAttribute("TargetMode") === "External") continue;
        const target = rel.getAttribute("Target");
        if (!target) continue;
        const renamed = renames.get(resolveTarget(baseDir, target));
        if (renamed) {
          rel.setAttribute("Target", relativeTarget(baseDir, renamed));
          relsChanged = true;
        }
      }
      if (relsChanged) {
        zip.file(relsName, serializer.serializeToString(relsDoc));
      }
    }

    await ensureDefaultContentTypes(zip, parser, serializer, newExts);
  }

  const data = await zip.generateAsync({
    type: "uint8array",
    compression: "DEFLATE",
  });

  return { data, totalImages, recompressed, originalBytes, compressedBytes };
}

/* -------------------------------------------------------------------------- */
/* Metadata scrubbing                                                         */
/* -------------------------------------------------------------------------- */

const CP_NS =
  "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
const DC_NS = "http://purl.org/dc/elements/1.1/";
const DCTERMS_NS = "http://purl.org/dc/terms/";
const EXT_NS =
  "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
const CUSTOM_NS =
  "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties";

const CORE_PART = "docProps/core.xml";
const APP_PART = "docProps/app.xml";
const CUSTOM_PART = "docProps/custom.xml";

const VT_NS = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";
const XSI_NS = "http://www.w3.org/2001/XMLSchema-instance";
const CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types";
const OFFICE_REL =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const CORE_REL_TYPE =
  "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
/** Standard format id Word uses for the custom-properties store. */
const CUSTOM_FMTID = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}";

/** Input hint for a metadata value, so the UI can pick a sensible control. */
export type MetadataKind = "text" | "date" | "number";

interface FieldDescriptor {
  id: string;
  label: string;
  group: string;
  /** Privacy-relevant fields that are pre-selected for removal. */
  sensitive: boolean;
  kind: MetadataKind;
  part: string;
  ns: string;
  local: string;
  /** Qualified name used when creating the element (prefix matches the part). */
  qualified: string;
  /** Core date fields carry an `xsi:type="dcterms:W3CDTF"` attribute. */
  w3cdtf?: boolean;
}

/** Known metadata fields Word writes into the core and extended property parts. */
const FIELD_DESCRIPTORS: FieldDescriptor[] = [
  { id: "core:creator", label: "Author", group: "People", sensitive: true, kind: "text", part: CORE_PART, ns: DC_NS, local: "creator", qualified: "dc:creator" },
  { id: "core:lastModifiedBy", label: "Last modified by", group: "People", sensitive: true, kind: "text", part: CORE_PART, ns: CP_NS, local: "lastModifiedBy", qualified: "cp:lastModifiedBy" },
  { id: "app:Company", label: "Company", group: "People", sensitive: true, kind: "text", part: APP_PART, ns: EXT_NS, local: "Company", qualified: "Company" },
  { id: "app:Manager", label: "Manager", group: "People", sensitive: true, kind: "text", part: APP_PART, ns: EXT_NS, local: "Manager", qualified: "Manager" },
  { id: "core:title", label: "Title", group: "Document", sensitive: false, kind: "text", part: CORE_PART, ns: DC_NS, local: "title", qualified: "dc:title" },
  { id: "core:subject", label: "Subject", group: "Document", sensitive: false, kind: "text", part: CORE_PART, ns: DC_NS, local: "subject", qualified: "dc:subject" },
  { id: "core:description", label: "Comments", group: "Document", sensitive: false, kind: "text", part: CORE_PART, ns: DC_NS, local: "description", qualified: "dc:description" },
  { id: "core:keywords", label: "Keywords / tags", group: "Document", sensitive: false, kind: "text", part: CORE_PART, ns: CP_NS, local: "keywords", qualified: "cp:keywords" },
  { id: "core:category", label: "Category", group: "Document", sensitive: false, kind: "text", part: CORE_PART, ns: CP_NS, local: "category", qualified: "cp:category" },
  { id: "core:contentStatus", label: "Content status", group: "Document", sensitive: false, kind: "text", part: CORE_PART, ns: CP_NS, local: "contentStatus", qualified: "cp:contentStatus" },
  { id: "app:Template", label: "Template", group: "Document", sensitive: false, kind: "text", part: APP_PART, ns: EXT_NS, local: "Template", qualified: "Template" },
  { id: "core:created", label: "Created", group: "Dates", sensitive: true, kind: "date", part: CORE_PART, ns: DCTERMS_NS, local: "created", qualified: "dcterms:created", w3cdtf: true },
  { id: "core:modified", label: "Modified", group: "Dates", sensitive: true, kind: "date", part: CORE_PART, ns: DCTERMS_NS, local: "modified", qualified: "dcterms:modified", w3cdtf: true },
  { id: "core:lastPrinted", label: "Last printed", group: "Dates", sensitive: true, kind: "date", part: CORE_PART, ns: CP_NS, local: "lastPrinted", qualified: "cp:lastPrinted" },
  { id: "core:revision", label: "Revision number", group: "Activity", sensitive: true, kind: "number", part: CORE_PART, ns: CP_NS, local: "revision", qualified: "cp:revision" },
  { id: "app:TotalTime", label: "Total editing time", group: "Activity", sensitive: true, kind: "number", part: APP_PART, ns: EXT_NS, local: "TotalTime", qualified: "TotalTime" },
  { id: "app:Application", label: "Application", group: "Activity", sensitive: false, kind: "text", part: APP_PART, ns: EXT_NS, local: "Application", qualified: "Application" },
  { id: "app:AppVersion", label: "App version", group: "Activity", sensitive: false, kind: "text", part: APP_PART, ns: EXT_NS, local: "AppVersion", qualified: "AppVersion" },
];

const DESCRIPTOR_BY_ID = new Map(FIELD_DESCRIPTORS.map((d) => [d.id, d]));

/** A known metadata field the UI can offer to add to a document. */
export interface MetadataFieldDef {
  id: string;
  label: string;
  group: string;
  kind: MetadataKind;
}

/** Catalogue of the standard Word metadata fields, for an "add field" menu. */
export const METADATA_CATALOG: MetadataFieldDef[] = FIELD_DESCRIPTORS.map(
  (d) => ({ id: d.id, label: d.label, group: d.group, kind: d.kind })
);

export interface MetadataField {
  id: string;
  label: string;
  value: string;
  group: string;
  sensitive: boolean;
  kind: MetadataKind;
  /** True for user-defined custom properties (the label is the property name). */
  custom: boolean;
}

export interface ScrubMetadataResult {
  data: Uint8Array;
  /** Number of metadata entries removed. */
  removed: number;
}

/** A set of metadata changes to apply: values to set/create and ids to remove. */
export interface MetadataEdit {
  /** Maps a field id to its new value. Creates the field if it does not exist. */
  set?: Record<string, string>;
  /** Field ids to delete. Removals are applied before sets. */
  remove?: string[];
}

export interface ApplyMetadataResult {
  data: Uint8Array;
  /** Number of fields removed. */
  removed: number;
  /** Number of fields created or updated. */
  changed: number;
}

/** Skeleton document for a property part that has to be created from scratch. */
const PART_SKELETON: Record<string, string> = {
  [CORE_PART]: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<cp:coreProperties xmlns:cp="${CP_NS}" xmlns:dc="${DC_NS}" xmlns:dcterms="${DCTERMS_NS}" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="${XSI_NS}"></cp:coreProperties>`,
  [APP_PART]: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Properties xmlns="${EXT_NS}" xmlns:vt="${VT_NS}"></Properties>`,
  [CUSTOM_PART]: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Properties xmlns="${CUSTOM_NS}" xmlns:vt="${VT_NS}"></Properties>`,
};

/** Content type and package relationship type for each property part. */
const PART_META: Record<string, { ct: string; relType: string }> = {
  [CORE_PART]: {
    ct: "application/vnd.openxmlformats-package.core-properties+xml",
    relType: CORE_REL_TYPE,
  },
  [APP_PART]: {
    ct: "application/vnd.openxmlformats-officedocument.extended-properties+xml",
    relType: `${OFFICE_REL}/extended-properties`,
  },
  [CUSTOM_PART]: {
    ct: "application/vnd.openxmlformats-officedocument.custom-properties+xml",
    relType: `${OFFICE_REL}/custom-properties`,
  },
};

function textOf(el: XmlElement): string {
  return (el.textContent ?? "").trim();
}

function setText(el: XmlElement, value: string): void {
  while (el.firstChild) el.removeChild(el.firstChild);
  el.appendChild(el.ownerDocument!.createTextNode(value));
}

function firstElement(
  doc: XmlDocument,
  ns: string,
  local: string
): XmlElement | null {
  const list = doc.getElementsByTagNameNS(ns, local);
  return list.length > 0 ? list.item(0) : null;
}

function firstChildElement(el: XmlElement): XmlElement | null {
  for (let n = el.firstChild; n; n = n.nextSibling) {
    if (n.nodeType === 1) return n as XmlElement;
  }
  return null;
}

async function loadParts(
  zip: JSZip,
  parser: DOMParser
): Promise<Map<string, XmlDocument>> {
  const docs = new Map<string, XmlDocument>();
  for (const part of [CORE_PART, APP_PART, CUSTOM_PART]) {
    const file = zip.file(part);
    if (file) {
      docs.set(part, parser.parseFromString(await file.async("string"), "application/xml"));
    }
  }
  return docs;
}

/**
 * Reads the document, application and custom metadata present in a .docx file.
 * Only fields with a non-empty value are returned.
 */
export async function readDocxMetadata(
  input: ArrayBuffer | Uint8Array
): Promise<MetadataField[]> {
  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(input);
  } catch {
    throw new Error("Not a valid .docx file (could not be read as a zip archive).");
  }

  const parser = new DOMParser();
  const docs = await loadParts(zip, parser);
  const fields: MetadataField[] = [];

  for (const d of FIELD_DESCRIPTORS) {
    const doc = docs.get(d.part);
    if (!doc) continue;
    const el = firstElement(doc, d.ns, d.local);
    if (!el) continue;
    const value = textOf(el);
    if (!value) continue;
    fields.push({
      id: d.id,
      label: d.label,
      value,
      group: d.group,
      sensitive: d.sensitive,
      kind: d.kind,
      custom: false,
    });
  }

  const customDoc = docs.get(CUSTOM_PART);
  if (customDoc) {
    for (const prop of toArray(customDoc.getElementsByTagNameNS(CUSTOM_NS, "property"))) {
      const name = prop.getAttribute("name");
      if (!name) continue;
      const value = textOf(prop);
      if (!value) continue;
      fields.push({
        id: `custom:${name}`,
        label: name,
        value,
        group: "Custom",
        sensitive: true,
        kind: "text",
        custom: true,
      });
    }
  }

  return fields;
}

/** Add the content-type override and package relationship for a new part. */
async function registerPart(
  part: string,
  zip: JSZip,
  parser: DOMParser,
  serializer: XMLSerializer
): Promise<void> {
  const meta = PART_META[part];

  const ctFile = zip.file("[Content_Types].xml");
  if (ctFile) {
    const ctDoc = parser.parseFromString(await ctFile.async("string"), "application/xml");
    const exists = toArray(ctDoc.getElementsByTagNameNS(CT_NS, "Override")).some(
      (o) => o.getAttribute("PartName") === `/${part}`
    );
    if (!exists) {
      const ov = ctDoc.createElementNS(CT_NS, "Override");
      ov.setAttribute("PartName", `/${part}`);
      ov.setAttribute("ContentType", meta.ct);
      ctDoc.documentElement!.appendChild(ov);
      zip.file("[Content_Types].xml", serializer.serializeToString(ctDoc));
    }
  }

  const relsFile = zip.file("_rels/.rels");
  if (relsFile) {
    const relsDoc = parser.parseFromString(await relsFile.async("string"), "application/xml");
    const rels = toArray(relsDoc.getElementsByTagNameNS(REL_NS, "Relationship"));
    const linked = rels.some(
      (r) => r.getAttribute("Target")?.replace(/^\//, "") === part
    );
    if (!linked) {
      let max = 0;
      for (const r of rels) {
        const m = /^rId(\d+)$/.exec(r.getAttribute("Id") ?? "");
        if (m) max = Math.max(max, parseInt(m[1], 10));
      }
      const rel = relsDoc.createElementNS(REL_NS, "Relationship");
      rel.setAttribute("Id", `rId${max + 1}`);
      rel.setAttribute("Type", meta.relType);
      rel.setAttribute("Target", part);
      relsDoc.documentElement!.appendChild(rel);
      zip.file("_rels/.rels", serializer.serializeToString(relsDoc));
    }
  }
}

/** Return the parsed property part, creating and registering it if absent. */
async function ensurePart(
  part: string,
  zip: JSZip,
  parser: DOMParser,
  serializer: XMLSerializer,
  docs: Map<string, XmlDocument>,
  changed: Set<string>
): Promise<XmlDocument> {
  const existing = docs.get(part);
  if (existing) return existing;
  const doc = parser.parseFromString(PART_SKELETON[part], "application/xml");
  docs.set(part, doc);
  changed.add(part);
  await registerPart(part, zip, parser, serializer);
  return doc;
}

function setKnownField(doc: XmlDocument, d: FieldDescriptor, value: string): void {
  let el = firstElement(doc, d.ns, d.local);
  if (!el) {
    el = doc.createElementNS(d.ns, d.qualified);
    if (d.w3cdtf) el.setAttributeNS(XSI_NS, "xsi:type", "dcterms:W3CDTF");
    doc.documentElement!.appendChild(el);
  }
  setText(el, value);
}

function setCustomProp(doc: XmlDocument, name: string, value: string): void {
  let maxPid = 1;
  for (const prop of toArray(doc.getElementsByTagNameNS(CUSTOM_NS, "property"))) {
    const pid = parseInt(prop.getAttribute("pid") ?? "", 10);
    if (!Number.isNaN(pid)) maxPid = Math.max(maxPid, pid);
    if (prop.getAttribute("name") === name) {
      const child = firstChildElement(prop);
      if (child) {
        setText(child, value);
      } else {
        const v = doc.createElementNS(VT_NS, "vt:lpwstr");
        setText(v, value);
        prop.appendChild(v);
      }
      return;
    }
  }
  const prop = doc.createElementNS(CUSTOM_NS, "property");
  prop.setAttribute("fmtid", CUSTOM_FMTID);
  prop.setAttribute("pid", String(maxPid + 1));
  prop.setAttribute("name", name);
  const v = doc.createElementNS(VT_NS, "vt:lpwstr");
  setText(v, value);
  prop.appendChild(v);
  doc.documentElement!.appendChild(prop);
}

function removeKnownField(doc: XmlDocument, d: FieldDescriptor): boolean {
  const el = firstElement(doc, d.ns, d.local);
  if (el && el.parentNode) {
    el.parentNode.removeChild(el);
    return true;
  }
  return false;
}

function removeCustomProp(doc: XmlDocument, name: string): boolean {
  let removed = false;
  for (const prop of toArray(doc.getElementsByTagNameNS(CUSTOM_NS, "property"))) {
    if (prop.getAttribute("name") === name && prop.parentNode) {
      prop.parentNode.removeChild(prop);
      removed = true;
    }
  }
  return removed;
}

/**
 * Applies a set of metadata edits to a .docx file: updates or creates the
 * fields in `set` and deletes the ids in `remove`. Known fields are written to
 * the core/extended property parts; `custom:<name>` ids become custom document
 * properties. Missing property parts (and their content-type / relationship
 * wiring) are created as needed.
 */
export async function applyDocxMetadata(
  input: ArrayBuffer | Uint8Array,
  edit: MetadataEdit
): Promise<ApplyMetadataResult> {
  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(input);
  } catch {
    throw new Error("Not a valid .docx file (could not be read as a zip archive).");
  }

  const parser = new DOMParser();
  const serializer = new XMLSerializer();
  const docs = await loadParts(zip, parser);
  const changed = new Set<string>();
  let removed = 0;
  let updated = 0;

  for (const id of edit.remove ?? []) {
    if (id.startsWith("custom:")) {
      const doc = docs.get(CUSTOM_PART);
      if (doc && removeCustomProp(doc, id.slice(7))) {
        changed.add(CUSTOM_PART);
        removed++;
      }
    } else {
      const d = DESCRIPTOR_BY_ID.get(id);
      const doc = d && docs.get(d.part);
      if (d && doc && removeKnownField(doc, d)) {
        changed.add(d.part);
        removed++;
      }
    }
  }

  for (const [id, value] of Object.entries(edit.set ?? {})) {
    if (id.startsWith("custom:")) {
      const name = id.slice(7).trim();
      if (!name) continue;
      const doc = await ensurePart(CUSTOM_PART, zip, parser, serializer, docs, changed);
      setCustomProp(doc, name, value);
      changed.add(CUSTOM_PART);
      updated++;
    } else {
      const d = DESCRIPTOR_BY_ID.get(id);
      if (!d) continue;
      const doc = await ensurePart(d.part, zip, parser, serializer, docs, changed);
      setKnownField(doc, d, value);
      changed.add(d.part);
      updated++;
    }
  }

  for (const part of changed) {
    zip.file(part, serializer.serializeToString(docs.get(part)!));
  }

  const data = await zip.generateAsync({
    type: "uint8array",
    compression: "DEFLATE",
  });

  return { data, removed, changed: updated };
}

/**
 * Removes the selected metadata entries from a .docx file. `ids` are the
 * {@link MetadataField.id} values to strip; everything else is left untouched.
 * Thin wrapper over {@link applyDocxMetadata}.
 */
export async function scrubDocxMetadata(
  input: ArrayBuffer | Uint8Array,
  ids: string[]
): Promise<ScrubMetadataResult> {
  const { data, removed } = await applyDocxMetadata(input, { remove: ids });
  return { data, removed };
}
