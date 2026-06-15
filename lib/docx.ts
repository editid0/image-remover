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

interface FieldDescriptor {
  id: string;
  label: string;
  group: string;
  /** Privacy-relevant fields that are pre-selected for removal. */
  sensitive: boolean;
  part: string;
  ns: string;
  local: string;
}

/** Known metadata fields in the core and extended property parts. */
const FIELD_DESCRIPTORS: FieldDescriptor[] = [
  { id: "core:creator", label: "Author", group: "People", sensitive: true, part: CORE_PART, ns: DC_NS, local: "creator" },
  { id: "core:lastModifiedBy", label: "Last modified by", group: "People", sensitive: true, part: CORE_PART, ns: CP_NS, local: "lastModifiedBy" },
  { id: "app:Company", label: "Company", group: "People", sensitive: true, part: APP_PART, ns: EXT_NS, local: "Company" },
  { id: "app:Manager", label: "Manager", group: "People", sensitive: true, part: APP_PART, ns: EXT_NS, local: "Manager" },
  { id: "core:title", label: "Title", group: "Document", sensitive: false, part: CORE_PART, ns: DC_NS, local: "title" },
  { id: "core:subject", label: "Subject", group: "Document", sensitive: false, part: CORE_PART, ns: DC_NS, local: "subject" },
  { id: "core:description", label: "Comments", group: "Document", sensitive: false, part: CORE_PART, ns: DC_NS, local: "description" },
  { id: "core:keywords", label: "Keywords / tags", group: "Document", sensitive: false, part: CORE_PART, ns: CP_NS, local: "keywords" },
  { id: "core:category", label: "Category", group: "Document", sensitive: false, part: CORE_PART, ns: CP_NS, local: "category" },
  { id: "core:contentStatus", label: "Content status", group: "Document", sensitive: false, part: CORE_PART, ns: CP_NS, local: "contentStatus" },
  { id: "app:Template", label: "Template", group: "Document", sensitive: false, part: APP_PART, ns: EXT_NS, local: "Template" },
  { id: "core:created", label: "Created", group: "Dates", sensitive: true, part: CORE_PART, ns: DCTERMS_NS, local: "created" },
  { id: "core:modified", label: "Modified", group: "Dates", sensitive: true, part: CORE_PART, ns: DCTERMS_NS, local: "modified" },
  { id: "core:lastPrinted", label: "Last printed", group: "Dates", sensitive: true, part: CORE_PART, ns: CP_NS, local: "lastPrinted" },
  { id: "core:revision", label: "Revision number", group: "Activity", sensitive: true, part: CORE_PART, ns: CP_NS, local: "revision" },
  { id: "app:TotalTime", label: "Total editing time", group: "Activity", sensitive: true, part: APP_PART, ns: EXT_NS, local: "TotalTime" },
  { id: "app:Application", label: "Application", group: "Activity", sensitive: false, part: APP_PART, ns: EXT_NS, local: "Application" },
  { id: "app:AppVersion", label: "App version", group: "Activity", sensitive: false, part: APP_PART, ns: EXT_NS, local: "AppVersion" },
];

export interface MetadataField {
  id: string;
  label: string;
  value: string;
  group: string;
  sensitive: boolean;
}

export interface ScrubMetadataResult {
  data: Uint8Array;
  /** Number of metadata entries removed. */
  removed: number;
}

function textOf(el: XmlElement): string {
  return (el.textContent ?? "").trim();
}

function firstElement(
  doc: XmlDocument,
  ns: string,
  local: string
): XmlElement | null {
  const list = doc.getElementsByTagNameNS(ns, local);
  return list.length > 0 ? list.item(0) : null;
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
 * Only fields that are actually present (with a non-empty value) are returned,
 * so the result mirrors what a metadata scrub would offer to strip.
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
    fields.push({ id: d.id, label: d.label, value, group: d.group, sensitive: d.sensitive });
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
      });
    }
  }

  return fields;
}

/**
 * Removes the selected metadata entries from a .docx file. `ids` are the
 * {@link MetadataField.id} values to strip; everything else is left untouched.
 */
export async function scrubDocxMetadata(
  input: ArrayBuffer | Uint8Array,
  ids: string[]
): Promise<ScrubMetadataResult> {
  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(input);
  } catch {
    throw new Error("Not a valid .docx file (could not be read as a zip archive).");
  }

  const parser = new DOMParser();
  const serializer = new XMLSerializer();
  const docs = await loadParts(zip, parser);
  const selected = new Set(ids);
  const changed = new Set<string>();
  let removed = 0;

  for (const d of FIELD_DESCRIPTORS) {
    if (!selected.has(d.id)) continue;
    const doc = docs.get(d.part);
    if (!doc) continue;
    const el = firstElement(doc, d.ns, d.local);
    if (el && el.parentNode) {
      el.parentNode.removeChild(el);
      changed.add(d.part);
      removed++;
    }
  }

  const customDoc = docs.get(CUSTOM_PART);
  if (customDoc) {
    for (const prop of toArray(customDoc.getElementsByTagNameNS(CUSTOM_NS, "property"))) {
      const name = prop.getAttribute("name");
      if (name && selected.has(`custom:${name}`) && prop.parentNode) {
        prop.parentNode.removeChild(prop);
        changed.add(CUSTOM_PART);
        removed++;
      }
    }
  }

  for (const part of changed) {
    zip.file(part, serializer.serializeToString(docs.get(part)!));
  }

  const data = await zip.generateAsync({
    type: "uint8array",
    compression: "DEFLATE",
  });

  return { data, removed };
}
