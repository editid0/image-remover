# Word Image Remover

A Next.js app that takes a Word document (`.docx`), removes all images, and returns the document with text and formatting intact.

## How it works

A `.docx` file is a zip archive of XML parts. The app:

1. Strips `w:drawing` (modern DrawingML) elements and `w:pict` (legacy VML) elements containing image data from the document body, headers, footers, footnotes, and endnotes. VML text boxes without images are left alone.
2. Removes image relationships that are no longer referenced.
3. Deletes media files in `word/media/` that nothing references anymore (images still used elsewhere, e.g. picture bullets, are kept).

Core logic lives in [`lib/docx.ts`](lib/docx.ts) and runs entirely in the browser — files never leave your device, so the app works as a fully static site.

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

## Deploying to GitHub Pages

The app is a static export (`output: "export"` in [`next.config.ts`](next.config.ts)); `next build` writes the site to `out/`. The workflow in [`.github/workflows/deploy.yml`](.github/workflows/deploy.yml) builds and deploys it on every push to `main`.

One-time setup:

1. Push this repository to GitHub.
2. In the repository, go to **Settings → Pages** and set **Source** to **GitHub Actions**.
3. Push to `main` (or run the workflow manually from the Actions tab).

The site will be served at `https://<user>.github.io/<repo>/`. The workflow passes the correct `BASE_PATH` automatically via `actions/configure-pages`, so it works for both project pages and user/org pages without config changes.
