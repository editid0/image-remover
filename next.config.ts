import type { NextConfig } from "next";

const nextConfig: NextConfig = {
  output: "export",
  // Set by the GitHub Pages workflow (e.g. "/image-remover" for project
  // pages, "" for user/org pages). Empty for local dev and `next build`.
  basePath: process.env.BASE_PATH || "",
};

export default nextConfig;
