/**
 * End-to-end tests for the .docx tools. Builds a minimal document with text,
 * images (DrawingML in the body, VML in the header) and metadata (core, app
 * and custom properties), then exercises:
 *   - removeImagesFromDocx  (images/rels/media gone, text survives)
 *   - extractImagesFromDocx (every embedded image returned)
 *   - readDocxMetadata / scrubDocxMetadata (selected fields stripped only)
 */
import JSZip from "jszip";
import {
  removeImagesFromDocx,
  extractImagesFromDocx,
  zipImages,
  readDocxMetadata,
  scrubDocxMetadata,
  applyDocxMetadata,
} from "../lib/docx";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PKG_R = "http://schemas.openxmlformats.org/package/2006/relationships";

// 1x1 transparent PNG
const PNG = Buffer.from(
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==",
  "base64"
);

async function buildTestDocx(): Promise<Buffer> {
  const zip = new JSZip();

  zip.file(
    "[Content_Types].xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  <Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>
</Types>`
  );

  zip.file(
    "_rels/.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${PKG_R}">
  <Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="${PKG_R}/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="${R}/extended-properties" Target="docProps/app.xml"/>
  <Relationship Id="rId4" Type="${R}/custom-properties" Target="docProps/custom.xml"/>
</Relationships>`
  );

  zip.file(
    "word/document.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W}" xmlns:r="${R}"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
  <w:body>
    <w:p><w:r><w:t>Hello before image</w:t></w:r></w:p>
    <w:p><w:r><w:drawing>
      <wp:inline><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
        <pic:pic><pic:blipFill><a:blip r:embed="rId10"/></pic:blipFill></pic:pic>
      </a:graphicData></a:graphic></wp:inline>
    </w:drawing></w:r></w:p>
    <w:p><w:r><w:t>Hello after image</w:t></w:r></w:p>
    <w:sectPr><w:headerReference w:type="default" r:id="rId11"/></w:sectPr>
  </w:body>
</w:document>`
  );

  zip.file(
    "word/_rels/document.xml.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${PKG_R}">
  <Relationship Id="rId10" Type="${R}/image" Target="media/image1.png"/>
  <Relationship Id="rId11" Type="${R}/header" Target="header1.xml"/>
</Relationships>`
  );

  zip.file(
    "word/header1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="${W}" xmlns:r="${R}" xmlns:v="urn:schemas-microsoft-com:vml">
  <w:p><w:r><w:t>Header text</w:t></w:r></w:p>
  <w:p><w:r><w:pict>
    <v:shape style="width:10pt;height:10pt"><v:imagedata r:id="rId1"/></v:shape>
  </w:pict></w:r></w:p>
  <w:p><w:r><w:pict>
    <v:shape style="width:10pt;height:10pt"><v:textbox><w:txbxContent>
      <w:p><w:r><w:t>Text box must survive</w:t></w:r></w:p>
    </w:txbxContent></v:textbox></v:shape>
  </w:pict></w:r></w:p>
</w:hdr>`
  );

  zip.file(
    "word/_rels/header1.xml.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${PKG_R}">
  <Relationship Id="rId1" Type="${R}/image" Target="media/image2.png"/>
</Relationships>`
  );

  zip.file(
    "docProps/core.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
  xmlns:dc="http://purl.org/dc/elements/1.1/"
  xmlns:dcterms="http://purl.org/dc/terms/"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:creator>Jane Doe</dc:creator>
  <cp:lastModifiedBy>John Smith</cp:lastModifiedBy>
  <dc:title>Quarterly Report</dc:title>
  <cp:revision>7</cp:revision>
  <dcterms:created xsi:type="dcterms:W3CDTF">2024-01-02T03:04:05Z</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">2024-02-03T04:05:06Z</dcterms:modified>
</cp:coreProperties>`
  );

  zip.file(
    "docProps/app.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
  xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Microsoft Office Word</Application>
  <Company>Acme Corp</Company>
  <Manager>Big Boss</Manager>
  <TotalTime>42</TotalTime>
</Properties>`
  );

  zip.file(
    "docProps/custom.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
  xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="ClientCode"><vt:lpwstr>SECRET-123</vt:lpwstr></property>
</Properties>`
  );

  zip.file("word/media/image1.png", PNG);
  zip.file("word/media/image2.png", PNG);

  return zip.generateAsync({ type: "nodebuffer" });
}

function assert(cond: boolean, label: string) {
  if (cond) console.log(`  PASS  ${label}`);
  else {
    console.error(`  FAIL  ${label}`);
    process.exitCode = 1;
  }
}

async function testRemove(input: Buffer) {
  console.log("\nremoveImagesFromDocx");
  const result = await removeImagesFromDocx(input);
  const out = await JSZip.loadAsync(result.data);
  const docXml = await out.file("word/document.xml")!.async("string");
  const hdrXml = await out.file("word/header1.xml")!.async("string");
  const docRels = await out.file("word/_rels/document.xml.rels")!.async("string");
  const hdrRels = await out.file("word/_rels/header1.xml.rels")!.async("string");

  assert(result.removedElements === 2, "removed 2 image elements");
  assert(result.removedMediaFiles === 2, "removed 2 media files");
  assert(!docXml.includes("w:drawing"), "no w:drawing left in document.xml");
  assert(!hdrXml.includes("imagedata"), "no v:imagedata left in header1.xml");
  assert(
    docXml.includes("Hello before image") && docXml.includes("Hello after image"),
    "body text survives"
  );
  assert(hdrXml.includes("Text box must survive"), "image-free text box survives");
  assert(!docRels.includes("media/image1.png"), "image rel removed from document rels");
  assert(!hdrRels.includes("media/image2.png"), "image rel removed from header rels");
  assert(!out.file("word/media/image1.png"), "media/image1.png deleted");
  assert(!out.file("word/media/image2.png"), "media/image2.png deleted");
}

async function testExtract(input: Buffer) {
  console.log("\nextractImagesFromDocx");
  const images = await extractImagesFromDocx(input);
  assert(images.length === 2, "extracted 2 images");
  assert(
    images.map((i) => i.name).join(",") === "image1.png,image2.png",
    "images returned sorted by name"
  );
  assert(images.every((i) => i.contentType === "image/png"), "png content type inferred");
  assert(images.every((i) => i.data.length === PNG.length), "image bytes preserved");

  const zipped = await zipImages(images);
  const out = await JSZip.loadAsync(zipped);
  assert(
    !!out.file("image1.png") && !!out.file("image2.png"),
    "zipImages bundles every image"
  );
}

async function testMetadata(input: Buffer) {
  console.log("\nreadDocxMetadata / scrubDocxMetadata");
  const fields = await readDocxMetadata(input);
  const byId = new Map(fields.map((f) => [f.id, f.value]));

  assert(byId.get("core:creator") === "Jane Doe", "reads author");
  assert(byId.get("core:lastModifiedBy") === "John Smith", "reads last modified by");
  assert(byId.get("core:title") === "Quarterly Report", "reads title");
  assert(byId.get("core:revision") === "7", "reads revision");
  assert(byId.get("core:created") === "2024-01-02T03:04:05Z", "reads created date");
  assert(byId.get("app:Company") === "Acme Corp", "reads company");
  assert(byId.get("app:Manager") === "Big Boss", "reads manager");
  assert(byId.get("app:TotalTime") === "42", "reads total editing time");
  assert(byId.get("custom:ClientCode") === "SECRET-123", "reads custom property");
  assert(
    fields.find((f) => f.id === "core:title")!.sensitive === false,
    "title is not flagged sensitive"
  );

  const result = await scrubDocxMetadata(input, [
    "core:creator",
    "core:lastModifiedBy",
    "app:Company",
    "custom:ClientCode",
  ]);
  assert(result.removed === 4, "removed 4 selected fields");

  const out = await JSZip.loadAsync(result.data);
  const core = await out.file("docProps/core.xml")!.async("string");
  const app = await out.file("docProps/app.xml")!.async("string");
  const custom = await out.file("docProps/custom.xml")!.async("string");

  assert(!core.includes("Jane Doe"), "author removed");
  assert(!core.includes("John Smith"), "last modified by removed");
  assert(core.includes("Quarterly Report"), "unselected title kept");
  assert(core.includes("2024-01-02T03:04:05Z"), "unselected created date kept");
  assert(!app.includes("Acme Corp"), "company removed");
  assert(app.includes("Big Boss"), "unselected manager kept");
  assert(!custom.includes("SECRET-123"), "custom property removed");
  assert(!!out.file("word/media/image1.png"), "scrub leaves images intact");
}

async function testEdit(input: Buffer) {
  console.log("\napplyDocxMetadata (edit existing parts)");
  const result = await applyDocxMetadata(input, {
    set: {
      "core:title": "Edited Title",
      "core:keywords": "alpha, beta",
      "custom:Project": "Apollo",
    },
    remove: ["app:Company"],
  });
  assert(result.changed === 3, "set 3 fields");
  assert(result.removed === 1, "removed 1 field");

  const out = await JSZip.loadAsync(result.data);
  const core = await out.file("docProps/core.xml")!.async("string");
  const app = await out.file("docProps/app.xml")!.async("string");
  const custom = await out.file("docProps/custom.xml")!.async("string");
  assert(core.includes("Edited Title"), "title updated in place");
  assert(!core.includes("Quarterly Report"), "old title value gone");
  assert(core.includes("alpha, beta"), "new keywords field created");
  assert(!app.includes("Acme Corp"), "company removed");
  assert(custom.includes("Apollo") && custom.includes("SECRET-123"), "custom prop added alongside existing");

  const meta = new Map((await readDocxMetadata(result.data)).map((f) => [f.id, f.value]));
  assert(meta.get("core:title") === "Edited Title", "re-read sees edited title");
  assert(meta.get("core:keywords") === "alpha, beta", "re-read sees new keywords");
  assert(meta.get("custom:Project") === "Apollo", "re-read sees new custom prop");
  assert(!meta.has("app:Company"), "re-read no longer sees company");
}

async function buildBareDocx(): Promise<Buffer> {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`
  );
  zip.file(
    "_rels/.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${PKG_R}">
  <Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/>
</Relationships>`
  );
  zip.file(
    "word/document.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:t>Bare</w:t></w:r></w:p></w:body></w:document>`
  );
  return zip.generateAsync({ type: "nodebuffer" });
}

async function testCreateParts() {
  console.log("\napplyDocxMetadata (create missing parts)");
  const bare = await buildBareDocx();
  const result = await applyDocxMetadata(bare, {
    set: {
      "core:creator": "New Author",
      "core:created": "2026-06-15T00:00:00Z",
      "app:Company": "NewCo",
      "custom:Foo": "Bar",
    },
  });
  assert(result.changed === 4, "set 4 fields into a doc with no docProps");

  const out = await JSZip.loadAsync(result.data);
  assert(!!out.file("docProps/core.xml"), "core.xml created");
  assert(!!out.file("docProps/app.xml"), "app.xml created");
  assert(!!out.file("docProps/custom.xml"), "custom.xml created");

  const core = await out.file("docProps/core.xml")!.async("string");
  assert(core.includes("New Author"), "author written to new core.xml");
  assert(
    core.includes('xsi:type="dcterms:W3CDTF"') && core.includes("2026-06-15T00:00:00Z"),
    "created date written with W3CDTF type"
  );

  const ct = await out.file("[Content_Types].xml")!.async("string");
  assert(ct.includes("/docProps/core.xml"), "content-type override added for core");
  assert(ct.includes("/docProps/custom.xml"), "content-type override added for custom");

  const rels = await out.file("_rels/.rels")!.async("string");
  assert(rels.includes("docProps/app.xml"), "package relationship added for app");

  const meta = new Map((await readDocxMetadata(result.data)).map((f) => [f.id, f.value]));
  assert(meta.get("core:creator") === "New Author", "re-read author from created part");
  assert(meta.get("app:Company") === "NewCo", "re-read company from created part");
  assert(meta.get("custom:Foo") === "Bar", "re-read custom prop from created part");
}

async function main() {
  const input = await buildTestDocx();
  await testRemove(input);
  await testExtract(input);
  await testMetadata(input);
  await testEdit(input);
  await testCreateParts();
  console.log(process.exitCode ? "\nTEST FAILED" : "\nALL TESTS PASSED");
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
