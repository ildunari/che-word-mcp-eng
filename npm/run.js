#!/usr/bin/env node
"use strict";

const { spawn } = require("child_process");
const path = require("path");

const PACKAGE_JSON = require("./package.json");
const PACKAGE_NAME = PACKAGE_JSON.name;
const BINARY_PATH = path.join(__dirname, "bin", "CheWordMCP");

const child = spawn(BINARY_PATH, process.argv.slice(2), {
  stdio: "inherit",
});

child.on("error", (err) => {
  if (err.code === "ENOENT") {
    console.error(
      `${PACKAGE_NAME} binary not found. Run: npm rebuild ${PACKAGE_NAME}`
    );
  } else {
    console.error(`Failed to start ${PACKAGE_NAME}:`, err.message);
  }
  process.exit(1);
});

child.on("exit", (code) => {
  process.exit(code ?? 1);
});
