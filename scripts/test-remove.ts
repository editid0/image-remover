/**
 * End-to-end test: builds a minimal .docx containing text + images
 * (DrawingML in the body, VML in the header), runs removeImagesFromDocx,
 * and asserts images, relationships and media files are gone while
 * text survives.
 */
import JSZip from "jszip";
import { removeImagesFromDocx } from "../lib/docx";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

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
</Types>`
  );

  zip.file(
    "_rels/.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/>
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
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
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
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${R}/image" Target="media/image2.png"/>
</Relationships>`
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

async function main() {
  const input = await buildTestDocx();
  const result = await removeImagesFromDocx(input);

  console.log(
    `removedElements=${result.removedElements} removedMediaFiles=${result.removedMediaFiles}`
  );

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
  assert(hdrXml.includes("Header text"), "header text survives");
  assert(hdrXml.includes("Text box must survive"), "image-free text box survives");
  assert(!docRels.includes("media/image1.png"), "image rel removed from document rels");
  assert(!hdrRels.includes("media/image2.png"), "image rel removed from header rels");
  assert(docRels.includes("header1.xml"), "header rel kept");
  assert(!out.file("word/media/image1.png"), "media/image1.png deleted");
  assert(!out.file("word/media/image2.png"), "media/image2.png deleted");

  console.log(process.exitCode ? "\nTEST FAILED" : "\nALL TESTS PASSED");
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
