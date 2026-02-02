#!/usr/bin/env node

/**
 * 构建脚本 - 将 src 中的文件复制到 lib
 */
const fs = require("fs");
const path = require("path");

const srcDir = path.join(__dirname, "src");
const libDir = path.join(__dirname, "lib");

// 确保 lib 目录存在
if (!fs.existsSync(libDir)) {
  fs.mkdirSync(libDir, { recursive: true });
}

// 复制所有 JS 文件
fs.readdirSync(srcDir).forEach((file) => {
  if (file.endsWith(".js")) {
    const srcFile = path.join(srcDir, file);
    const libFile = path.join(libDir, file);
    fs.copyFileSync(srcFile, libFile);
    console.log(`✓ Copied ${file} to lib/`);
  }
});

console.log("Build completed!");
