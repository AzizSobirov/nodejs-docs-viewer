import express from "express";
import axios from "axios";
import { fileTypeFromBuffer } from "file-type";
import mime from "mime-types";
import fs from "fs-extra";
import path from "path";
import crypto from "crypto";
import libre from "libreoffice-convert";
import { promisify } from "util";
import mammoth from "mammoth";
import XLSX from "xlsx";
import { parse } from "csv-parse/sync";
import ExcelJS from "exceljs";

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

async function docxToHtml(buffer) {
  const result = await mammoth.convertToHtml({ buffer });
  return `<!doctype html><meta charset="utf-8">${result.value}`;
}

function xlsxToHtml(buffer) {
  const workbook = XLSX.read(buffer, { type: "buffer" });
  const sheets = workbook.SheetNames.map((name) => {
    const html = XLSX.utils.sheet_to_html(workbook.Sheets[name], { id: name });
    return `<h2>${name}</h2>` + html;
  }).join("<hr/>");
  return `<!doctype html><meta charset="utf-8">${sheets}`;
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

// XLSX -> HTML (with styles)
async function excelToStyledHtml(buffer) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);

  const worksheet = workbook.worksheets[0];

  let html = `<table border="1" cellspacing="0" cellpadding="4" style="border-collapse: collapse;">`;

  worksheet.eachRow((row) => {
    html += "<tr>";
    row.eachCell((cell) => {
      const { font, alignment, fill } = cell.style;

      // Basic styles
      const styles = [];
      if (font?.bold) styles.push("font-weight:bold");
      if (font?.italic) styles.push("font-style:italic");
      if (font?.color?.argb)
        styles.push(
          `color:#${font.color.argb.slice(2)}` // remove alpha
        );
      if (alignment?.horizontal)
        styles.push(`text-align:${alignment.horizontal}`);
      if (fill?.fgColor?.argb)
        styles.push(`background-color:#${fill.fgColor.argb.slice(2)}`);

      html += `<td style="${styles.join(";")}">${cell.value ?? ""}</td>`;
    });
    html += "</tr>";
  });

  html += "</table>";
  return html;
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
      res.setHeader("Content-Type", "application/pdf");
      res.setHeader("Content-Disposition", 'inline; filename="preview.pdf"');
      return fs.createReadStream(cachedPdf).pipe(res);
    }

    const buffer = await downloadToBuffer(src);
    await fs.writeFile(cachedRaw, buffer);

    const detection = await detect(buffer, src);
    const ext = (detection.ext || "").toLowerCase();
    const mimeType = detection.mime || "application/octet-stream";

    // PDF
    if (mimeType === "application/pdf" || ext === ".pdf") {
      await fs.writeFile(cachedPdf, buffer);
      res.setHeader("Content-Type", "application/pdf");
      res.setHeader("Content-Disposition", 'inline; filename="preview.pdf"');
      return fs.createReadStream(cachedPdf).pipe(res);
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
      res.setHeader("Content-Type", "application/pdf");
      res.setHeader("Content-Disposition", 'inline; filename="preview.pdf"');
      return fs.createReadStream(cachedPdf).pipe(res);
    }

    // XLSX / XLS
    if (ext === ".xlsx" || ext === ".xls" || mimeType.includes("spreadsheet")) {
      try {
        const html = await excelToStyledHtml(buffer);
        res.setHeader("Content-Type", "text/html");
        return res.send(html);
        // const html = xlsxToHtml(buffer);
        // await fs.writeFile(cachedHtml, html, "utf8");
        // return res.type("html").send(html);
      } catch (e) {
        const pdfBuf = await convertAsync(buffer, ".pdf", undefined);
        await fs.writeFile(cachedPdf, pdfBuf);
        res.setHeader("Content-Type", "application/pdf");
        res.setHeader("Content-Disposition", 'inline; filename="preview.pdf"');
        return fs.createReadStream(cachedPdf).pipe(res);
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
      res.setHeader("Content-Type", "application/pdf");
      res.setHeader("Content-Disposition", 'inline; filename="preview.pdf"');
      return fs.createReadStream(cachedPdf).pipe(res);
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
      res.setHeader("Content-Type", "application/pdf");
      res.setHeader("Content-Disposition", 'inline; filename="preview.pdf"');
      return fs.createReadStream(cachedPdf).pipe(res);
    } catch {
      res.setHeader("Content-Type", mimeType);
      res.setHeader("Content-Disposition", `attachment; filename="file${ext}"`);
      return res.send(buffer);
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
