# Word Image Remover

A Next.js app that takes a Word document (`.docx`), removes all images, and returns the document with text and formatting intact.

## How it works

A `.docx` file is a zip archive of XML parts. The app:

1. Strips `w:drawing` (modern DrawingML) elements and `w:pict` (legacy VML) elements containing image data from the document body, headers, footers, footnotes, and endnotes. VML text boxes without images are left alone.
2. Removes image relationships that are no longer referenced.
3. Deletes media files in `word/media/` that nothing references anymore (images still used elsewhere, e.g. picture bullets, are kept).

Core logic lives in [`lib/docx.ts`](lib/docx.ts); the upload endpoint is [`app/api/remove-images/route.ts`](app/api/remove-images/route.ts).

## Usage

```bash
npm install
npm run dev
```

Open [http://localhost:3000](http://localhost:3000), drop in a `.docx`, and the image-free copy downloads automatically as `<name>-no-images.docx`.

## Test

```bash
npx tsx scripts/test-remove.ts
```

Builds a sample document with images in the body and header, runs the removal, and asserts images, relationships, and media files are gone while text survives.
