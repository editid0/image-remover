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
