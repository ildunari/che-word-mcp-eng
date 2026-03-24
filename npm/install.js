#!/usr/bin/env node
"use strict";

const { execSync } = require("child_process");
const fs = require("fs");
const path = require("path");

const PACKAGE_JSON = require("./package.json");
const VERSION = PACKAGE_JSON.version;
const PACKAGE_NAME = PACKAGE_JSON.name;
const BINARY_NAME = "CheWordMCP";
const REPO = "ildunari/che-word-mcp-eng";
const BINARY_DIR = path.join(__dirname, "bin");
const BINARY_PATH = path.join(BINARY_DIR, BINARY_NAME);

if (process.platform !== "darwin") {
  console.error(
    `${PACKAGE_NAME} requires macOS (Swift binary). ` +
      "Your platform: " +
      process.platform
  );
  process.exit(1);
}

if (fs.existsSync(BINARY_PATH)) {
  console.log(`${PACKAGE_NAME} binary already exists, skipping download.`);
  process.exit(0);
}

const url = `https://github.com/${REPO}/releases/download/v${VERSION}/${BINARY_NAME}`;

console.log(`Downloading ${PACKAGE_NAME} v${VERSION}...`);

fs.mkdirSync(BINARY_DIR, { recursive: true });

try {
  execSync(`curl -fsSL "${url}" -o "${BINARY_PATH}"`, { stdio: "inherit" });
  fs.chmodSync(BINARY_PATH, 0o755);
  console.log(`${PACKAGE_NAME} installed successfully.`);
} catch (err) {
  console.error(`Failed to download ${PACKAGE_NAME} binary from:`, url);
  console.error(err.message);
  process.exit(1);
}
