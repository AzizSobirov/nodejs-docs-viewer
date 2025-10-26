import express from "express";
import axios from "axios";
import { fileTypeFromBuffer } from "file-type";
import mime from "mime-types";
import fs from "fs-extra";
import path from "path";
import crypto from "crypto";
import libre from "libreoffice-convert";
import { promisify } from "util";
import { parse } from "csv-parse/sync";
import LuckyExcel from "luckyexcel";

const __dirname = path.resolve();
const convertAsync = promisify(libre.convert);

const app = express();
const PORT = process.env.PORT || 3000;
const CACHE_DIR = path.join(__dirname, "cache");
const MAX_FILE_SIZE =
  parseInt(process.env.MAX_FILE_SIZE || "200") * 1024 * 1024; // in MB, default 200MB
const CACHE_TIME = parseInt(process.env.CACHE_TIME || "86400") * 1000; // in seconds, default 24 hours

await fs.ensureDir(CACHE_DIR);

function hash(text) {
  return crypto.createHash("sha1").update(text).digest("hex");
}

async function downloadToBuffer(url) {
  const res = await axios.get(url, {
    responseType: "arraybuffer",
    maxContentLength: MAX_FILE_SIZE,
  });
  return Buffer.from(res.data);
}

async function detect(buffer, fallbackName) {
  const ft = await fileTypeFromBuffer(buffer);
  if (ft) return { ext: `.${ft.ext}`, mime: ft.mime };
  const ext = path.extname(fallbackName || "") || "";
  return { ext, mime: mime.lookup(ext) || "application/octet-stream" };
}

function csvToHtml(buffer) {
  const text = buffer.toString("utf8");
  const records = parse(text, { columns: false, skip_empty_lines: true });
  const rows = records
    .map(
      (r) =>
        `<tr>${r.map((c) => `<td>${escapeHtml(String(c))}</td>`).join("")}</tr>`
    )
    .join("");
  return `<!doctype html><meta charset="utf-8"><table border="1" cellpadding="4">${rows}</table>`;
}

// XLSX -> HTML (with Luckysheet)
async function excelToLuckysheet(buffer) {
  return new Promise((resolve, reject) => {
    LuckyExcel.transformExcelToLucky(buffer, (exportJson, luckysheetfile) => {
      if (
        !exportJson ||
        exportJson.sheets == null ||
        exportJson.sheets.length === 0
      ) {
        reject(new Error("Failed to parse Excel file"));
        return;
      }

      const html = `<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Excel Preview</title>
  <link rel='stylesheet' href='https://cdn.jsdelivr.net/npm/luckysheet@latest/dist/plugins/css/pluginsCss.css' />
  <link rel='stylesheet' href='https://cdn.jsdelivr.net/npm/luckysheet@latest/dist/plugins/plugins.css' />
  <link rel='stylesheet' href='https://cdn.jsdelivr.net/npm/luckysheet@latest/dist/css/luckysheet.css' />
  <link rel='stylesheet' href='https://cdn.jsdelivr.net/npm/luckysheet@latest/dist/assets/iconfont/iconfont.css' />
  <script src="https://cdn.jsdelivr.net/npm/luckysheet@latest/dist/plugins/js/plugin.js"></script>
  <script src="https://cdn.jsdelivr.net/npm/luckysheet@latest/dist/luckysheet.umd.js"></script>
  <style>
    body, html {
      margin: 0;
      padding: 0;
      width: 100%;
      height: 100%;
      overflow: hidden;
    }
    #luckysheet {
      margin: 0;
      padding: 0;
      position: absolute;
      width: 100%;
      height: 100%;
      left: 0;
      top: 0;
    }
  </style>
</head>
<body>
  <div id="luckysheet"></div>
  <script>
    $(function () {
      luckysheet.create({
        container: 'luckysheet',
        showtoolbar: true,
        showinfobar: true,
        showsheetbar: true,
        showstatisticBar: true,
        sheetFormulaBar: true,
        enableAddRow: false,
        enableAddCol: false,
        userInfo: false,
        showConfigWindowResize: false,
        allowEdit: false,
        data: ${JSON.stringify(exportJson.sheets)},
        title: '${escapeHtml(exportJson.info?.name || "Excel Preview")}',
        lang: 'en'
      });
    });
  </script>
</body>
</html>`;
      resolve(html);
    });
  });
}

function escapeHtml(s) {
  return s.replace(
    /[&<>"']/g,
    (c) =>
      ({
        "&": "&amp;",
        "<": "&lt;",
        ">": "&gt;",
        '"': "&quot;",
        "'": "&#39;",
      }[c])
  );
}

async function isCacheValid(filePath) {
  try {
    const stats = await fs.stat(filePath);
    const age = Date.now() - stats.mtimeMs;
    return age < CACHE_TIME;
  } catch {
    return false;
  }
}

function createPdfViewerHtml(pdfUrl) {
  return `<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>PDF Preview</title>
  <style>
    * {
      margin: 0;
      padding: 0;
      box-sizing: border-box;
    }
    body, html {
      width: 100%;
      height: 100%;
      overflow: hidden;
      background: #525659;
    }
    .loader-container {
      position: fixed;
      top: 0;
      left: 0;
      width: 100%;
      height: 100%;
      display: flex;
      flex-direction: column;
      align-items: center;
      justify-content: center;
      background: #525659;
      z-index: 9999;
      transition: opacity 0.3s ease-out;
    }
    .loader-container.hidden {
      opacity: 0;
      pointer-events: none;
    }
    .spinner {
      width: 50px;
      height: 50px;
      border: 4px solid rgba(255, 255, 255, 0.3);
      border-top-color: #fff;
      border-radius: 50%;
      animation: spin 1s linear infinite;
    }
    @keyframes spin {
      to { transform: rotate(360deg); }
    }
    .loader-text {
      color: #fff;
      font-family: Arial, sans-serif;
      font-size: 16px;
      margin-top: 20px;
    }
    #pdf-frame {
      width: 100%;
      height: 100%;
      border: none;
      display: none;
    }
    #pdf-frame.loaded {
      display: block;
    }
  </style>
</head>
<body>
  <div class="loader-container" id="loader">
    <div class="spinner"></div>
    <div class="loader-text">Loading document...</div>
  </div>
  <iframe id="pdf-frame" src="${pdfUrl}"></iframe>
  <script>
    const iframe = document.getElementById('pdf-frame');
    const loader = document.getElementById('loader');
    
    iframe.onload = function() {
      iframe.classList.add('loaded');
      loader.classList.add('hidden');
      setTimeout(() => {
        loader.style.display = 'none';
      }, 300);
    };
    
    // Fallback timeout in case onload doesn't fire
    setTimeout(() => {
      if (!iframe.classList.contains('loaded')) {
        iframe.classList.add('loaded');
        loader.classList.add('hidden');
      }
    }, 10000);
  </script>
</body>
</html>`;
}

function createUnsupportedFileHtml(downloadUrl, filename, ext) {
  return `<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>File Not Supported</title>
  <style>
    * {
      margin: 0;
      padding: 0;
      box-sizing: border-box;
    }
    body {
      font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, Cantarell, sans-serif;
      background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
      min-height: 100vh;
      display: flex;
      align-items: center;
      justify-content: center;
      padding: 20px;
    }
    .container {
      background: white;
      border-radius: 20px;
      padding: 60px 40px;
      max-width: 500px;
      text-align: center;
      box-shadow: 0 20px 60px rgba(0,0,0,0.3);
    }
    .icon {
      font-size: 80px;
      margin-bottom: 20px;
    }
    h1 {
      color: #333;
      font-size: 28px;
      margin-bottom: 15px;
    }
    p {
      color: #666;
      font-size: 16px;
      line-height: 1.6;
      margin-bottom: 30px;
    }
    .file-info {
      background: #f5f5f5;
      padding: 15px;
      border-radius: 10px;
      margin-bottom: 30px;
    }
    .file-name {
      font-weight: 600;
      color: #333;
      word-break: break-all;
      margin-bottom: 5px;
    }
    .file-type {
      color: #888;
      font-size: 14px;
    }
    .download-btn {
      display: inline-block;
      background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
      color: white;
      padding: 15px 40px;
      border-radius: 50px;
      text-decoration: none;
      font-weight: 600;
      font-size: 16px;
      transition: transform 0.2s, box-shadow 0.2s;
      box-shadow: 0 4px 15px rgba(102, 126, 234, 0.4);
    }
    .download-btn:hover {
      transform: translateY(-2px);
      box-shadow: 0 6px 20px rgba(102, 126, 234, 0.6);
    }
    .download-btn:active {
      transform: translateY(0);
    }
    .note {
      margin-top: 20px;
      font-size: 14px;
      color: #999;
    }
  </style>
</head>
<body>
  <div class="container">
    <div class="icon">📄</div>
    <h1>Preview Not Available</h1>
    <p>This file type cannot be previewed in the browser. You can download it to view on your device.</p>
    
    <div class="file-info">
      <div class="file-name">${escapeHtml(filename)}</div>
      <div class="file-type">File type: ${escapeHtml(ext || "Unknown")}</div>
    </div>
    
    <a href="${downloadUrl}" class="download-btn" download>
      ⬇️ Download File
    </a>
    
    <div class="note">
      The file will be downloaded to your device
    </div>
  </div>
</body>
</html>`;
}

// Route to serve raw PDF files
app.get("/pdf/:key", async (req, res) => {
  try {
    const key = req.params.key;
    const cachedPdf = path.join(CACHE_DIR, `${key}.pdf`);

    if (await fs.pathExists(cachedPdf)) {
      res.setHeader("Content-Type", "application/pdf");
      res.setHeader("Content-Disposition", 'inline; filename="preview.pdf"');
      return fs.createReadStream(cachedPdf).pipe(res);
    }

    res.status(404).send("PDF not found");
  } catch (err) {
    console.error(err);
    res.status(500).send("Error loading PDF");
  }
});

// Route to download raw files
app.get("/download/:key", async (req, res) => {
  try {
    const key = req.params.key;
    const cachedRaw = path.join(CACHE_DIR, `${key}.raw`);

    if (await fs.pathExists(cachedRaw)) {
      const buffer = await fs.readFile(cachedRaw);
      const detection = await detect(buffer, "");
      res.setHeader(
        "Content-Type",
        detection.mime || "application/octet-stream"
      );
      res.setHeader(
        "Content-Disposition",
        `attachment; filename="file${detection.ext}"`
      );
      return res.send(buffer);
    }

    res.status(404).send("File not found");
  } catch (err) {
    console.error(err);
    res.status(500).send("Error downloading file");
  }
});

app.get("/preview", async (req, res) => {
  try {
    const src = req.query.src;
    if (!src) return res.status(400).send("Missing src parameter");
    if (!/^https?:\/\//i.test(src))
      return res.status(400).send("src must be an http(s) URL");

    const key = hash(src);
    const cachedPdf = path.join(CACHE_DIR, `${key}.pdf`);
    const cachedHtml = path.join(CACHE_DIR, `${key}.html`);
    const cachedRaw = path.join(CACHE_DIR, `${key}.raw`);

    if ((await fs.pathExists(cachedHtml)) && (await isCacheValid(cachedHtml))) {
      return res.type("html").send(await fs.readFile(cachedHtml, "utf8"));
    }
    if ((await fs.pathExists(cachedPdf)) && (await isCacheValid(cachedPdf))) {
      const viewerHtml = createPdfViewerHtml(`/pdf/${key}`);
      return res.type("html").send(viewerHtml);
    }

    const buffer = await downloadToBuffer(src);
    await fs.writeFile(cachedRaw, buffer);

    const detection = await detect(buffer, src);
    const ext = (detection.ext || "").toLowerCase();
    const mimeType = detection.mime || "application/octet-stream";

    // PDF
    if (mimeType === "application/pdf" || ext === ".pdf") {
      await fs.writeFile(cachedPdf, buffer);
      const viewerHtml = createPdfViewerHtml(`/pdf/${key}`);
      return res.type("html").send(viewerHtml);
    }

    // DOCX
    // if (
    //   ext === ".docx" ||
    //   mimeType ===
    //     "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    // ) {
    //   const html = await docxToHtml(buffer);
    //   await fs.writeFile(cachedHtml, html, "utf8");
    //   return res.type("html").send(html);
    // }

    // DOC / DOCX / ODT
    if (
      ext === ".doc" ||
      ext === ".docx" ||
      mimeType === "application/msword" ||
      mimeType ===
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document"[
          ".odt"
        ]?.includes(ext)
    ) {
      const pdfBuf = await convertAsync(buffer, ".pdf", undefined);
      await fs.writeFile(cachedPdf, pdfBuf);
      const viewerHtml = createPdfViewerHtml(`/pdf/${key}`);
      return res.type("html").send(viewerHtml);
    }

    // XLSX / XLS
    if (ext === ".xlsx" || ext === ".xls" || mimeType.includes("spreadsheet")) {
      try {
        const html = await excelToLuckysheet(buffer);
        await fs.writeFile(cachedHtml, html, "utf8");
        return res.type("html").send(html);
      } catch (e) {
        console.error("Excel conversion error:", e);
        const pdfBuf = await convertAsync(buffer, ".pdf", undefined);
        await fs.writeFile(cachedPdf, pdfBuf);
        const viewerHtml = createPdfViewerHtml(`/pdf/${key}`);
        return res.type("html").send(viewerHtml);
      }
    }

    // PPTX / PPT
    if (
      ext === ".pptx" ||
      ext === ".ppt" ||
      mimeType.includes("presentation")
    ) {
      const pdfBuf = await convertAsync(buffer, ".pdf", undefined);
      await fs.writeFile(cachedPdf, pdfBuf);
      const viewerHtml = createPdfViewerHtml(`/pdf/${key}`);
      return res.type("html").send(viewerHtml);
    }

    // CSV
    if (ext === ".csv" || mimeType === "text/csv") {
      const html = csvToHtml(buffer);
      await fs.writeFile(cachedHtml, html, "utf8");
      return res.type("html").send(html);
    }

    // Fallback
    try {
      const pdfBuf = await convertAsync(buffer, ".pdf", undefined);
      await fs.writeFile(cachedPdf, pdfBuf);
      const viewerHtml = createPdfViewerHtml(`/pdf/${key}`);
      return res.type("html").send(viewerHtml);
    } catch {
      // Show unsupported file page with download button
      const filename = path.basename(src) || `file${ext}`;
      const unsupportedHtml = createUnsupportedFileHtml(
        `/download/${key}`,
        filename,
        ext
      );
      return res.type("html").send(unsupportedHtml);
    }
  } catch (err) {
    console.error(err);
    res.status(500).send("Server error: " + (err.message || String(err)));
  }
});

app.listen(PORT, () =>
  console.log(
    `✅ Office viewer running: http://localhost:${PORT}/preview?src=...`
  )
);
