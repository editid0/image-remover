# Word Doc Toolkit

A Next.js app for working with the contents of a Word document (`.docx`) — remove
images, pull images out, or view and edit the document's hidden metadata. Three
tools, one drop zone:

- **Remove images** — get back the same document with every image stripped out.
  Text, formatting, and layout stay untouched.
- **Extract images** — pull every embedded image out of the document and download
  them individually or as a single zip.
- **Edit metadata** — see the author, dates, company, custom properties, and other
  hidden fields baked into the file. Edit values, strip the ones you want gone, or
  add new Word-supported properties.

Everything runs entirely in the browser — files never leave your device, so the
app works as a fully static site.

## How it works

A `.docx` file is a zip archive of XML parts. Core logic lives in
[`lib/docx.ts`](lib/docx.ts):

- **`removeImagesFromDocx`** strips `w:drawing` (modern DrawingML) and `w:pict`
  (legacy VML) image elements from the document body, headers, footers, footnotes,
  and endnotes, drops the now-unused image relationships, and deletes orphaned
  files in `word/media/`. VML text boxes without images, and images still used
  elsewhere (e.g. picture bullets), are left alone.
- **`extractImagesFromDocx`** returns every image in `word/media/`; `zipImages`
  bundles them for download.
- **`readDocxMetadata`** reads the core, extended (app), and custom properties
  from `docProps/`. **`applyDocxMetadata`** writes them back: it updates or creates
  fields and deletes the ones you removed, creating missing property parts (and
  their content-type / relationship wiring) as needed. **`scrubDocxMetadata`** is a
  thin remove-only wrapper.

## Usage

```bash
npm install
npm run dev
```

Open [http://localhost:3000](http://localhost:3000), pick a tool, and drop in a
`.docx`. Removing images downloads `<name>-no-images.docx`; extracting downloads
the images (or `<name>-images.zip`); editing metadata downloads
`<name>-edited.docx`.

## Test

```bash
npx tsx scripts/test-docx.ts
```

Builds a sample document with text, images (in the body and header), and metadata
(core, app, and custom properties), then asserts: image removal clears images,
relationships, and media while text survives; extraction returns every image;
metadata reads, edits, removals, and field/part creation all behave.

## Deploying to GitHub Pages

The app is a static export (`output: "export"` in [`next.config.ts`](next.config.ts));
`next build` writes the site to `out/`. The workflow in
[`.github/workflows/deploy.yml`](.github/workflows/deploy.yml) builds and deploys it
on every push to `main`.

One-time setup:

1. Push this repository to GitHub.
2. In the repository, go to **Settings → Pages** and set **Source** to **GitHub Actions**.
3. Push to `main` (or run the workflow manually from the Actions tab).

The site will be served at `https://<user>.github.io/<repo>/`. The workflow passes
the correct `BASE_PATH` automatically via `actions/configure-pages`, so it works for
both project pages and user/org pages without config changes.
