#!/usr/bin/env node

const fs = require("fs");
const path = require("path");
const https = require("https");
const XLSX = require("xlsx");
const archiver = require("archiver");

const args = process.argv.slice(2);
const filePath = args[0];

if (!filePath) {
  console.log("Usage: node ziphire.js <excel_file.xlsx>");
  console.log("Example: node ziphire.js students.xlsx");
  process.exit(1);
}

if (!fs.existsSync(filePath)) {
  console.error(`❌ File not found: ${filePath}`);
  process.exit(1);
}

// ── Helpers ──────────────────────────────────────────────────────────────────

function extractFileId(url) {
  const s = String(url).trim();
  // /file/d/ID/
  const m1 = s.match(/\/file\/d\/([a-zA-Z0-9_-]+)/);
  if (m1) return m1[1];
  // ?id=ID or &id=ID
  const m2 = s.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (m2) return m2[1];
  // /folders/ — not a file, skip
  // bare ID only (33 chars, alphanumeric+dash+underscore)
  const m3 = s.match(/^([a-zA-Z0-9_-]{25,})$/);
  if (m3) return m3[1];
  return null;
}

function sanitize(str) {
  return String(str).trim().replace(/[^a-zA-Z0-9_\- ]/g, "").replace(/\s+/g, "_");
}

// Extract confirm token from Google's virus-scan warning page
function extractConfirmToken(html) {
  const m = html.match(/confirm=([0-9A-Za-z_\-]+)/);
  return m ? m[1] : "t";
}

function mergeCookies(existing, incomingSetCookie) {
  const jar = new Map();

  const addCookieString = (cookieStr) => {
    if (!cookieStr) return;
    cookieStr
      .split(";")
      .map((s) => s.trim())
      .filter(Boolean)
      .forEach((part) => {
        const eq = part.indexOf("=");
        if (eq <= 0) return;
        const key = part.slice(0, eq).trim();
        const val = part.slice(eq + 1).trim();
        if (!key) return;
        jar.set(key, val);
      });
  };

  addCookieString(existing);
  (incomingSetCookie || []).forEach((raw) => {
    const pair = String(raw).split(";")[0];
    addCookieString(pair);
  });

  return Array.from(jar.entries())
    .map(([k, v]) => `${k}=${v}`)
    .join("; ");
}

function createError(code, message, details = {}) {
  const err = new Error(message);
  err.code = code;
  err.details = details;
  return err;
}

function downloadFromDrive(fileId) {
  return new Promise((resolve, reject) => {
    const baseUrl = `https://drive.google.com/uc?export=download&id=${fileId}`;
    const maxHops = 12;

    function doRequest(url, cookies = "", hop = 0) {
      if (hop > maxHops) {
        return reject(
          createError(
            "MAX_REDIRECTS",
            `Too many redirects while downloading file id ${fileId}`,
            { fileId, url, hop, maxHops }
          )
        );
      }

      const get = https;
      const opts = {
        timeout: 20000,
        headers: {
          "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
          ...(cookies ? { Cookie: cookies } : {}),
        },
      };

      const req = get.get(url, opts, (res) => {
        const { statusCode } = res;
        const contentType = res.headers["content-type"] || "";
        const location = res.headers["location"] || "";
        const setCookie = res.headers["set-cookie"] || [];

        // Preserve and update cookies across all hops
        const mergedCookies = mergeCookies(cookies, setCookie);

        // Standard redirect
        if (statusCode === 301 || statusCode === 302 || statusCode === 303) {
          if (!location) {
            return reject(
              createError("REDIRECT_NO_LOCATION", "Redirect with no location", {
                fileId,
                url,
                statusCode,
                hop,
              })
            );
          }
          const nextUrl = location.startsWith("http") ? location : `https://drive.google.com${location}`;
          return doRequest(nextUrl, mergedCookies, hop + 1);
        }

        if (statusCode !== 200) {
          res.resume();
          return reject(
            createError("HTTP_STATUS", `HTTP ${statusCode}`, {
              fileId,
              url,
              statusCode,
              hop,
              contentType,
            })
          );
        }

        // If we got HTML, it's the virus-scan confirmation page
        if (contentType.includes("text/html")) {
          const chunks = [];
          res.on("data", (c) => chunks.push(c));
          res.on("end", () => {
            const html = Buffer.concat(chunks).toString();
            const confirm = extractConfirmToken(html);
            const uuid = html.match(/uuid=([a-zA-Z0-9_-]+)/)?.[1];
            // Build confirmed download URL
            let confirmUrl = `https://drive.google.com/uc?export=download&id=${fileId}&confirm=${confirm}`;
            if (uuid) confirmUrl += `&uuid=${uuid}`;
            doRequest(confirmUrl, mergedCookies, hop + 1);
          });
          res.on("error", (err) => {
            reject(
              createError("HTML_READ_ERROR", err.message, {
                fileId,
                url,
                hop,
              })
            );
          });
          return;
        }

        // Actual file content
        const chunks = [];
        res.on("data", (c) => chunks.push(c));
        res.on("end", () => resolve(Buffer.concat(chunks)));
        res.on("error", (err) => {
          reject(
            createError("FILE_STREAM_ERROR", err.message, {
              fileId,
              url,
              hop,
            })
          );
        });
      });

      req.on("timeout", () => {
        req.destroy(
          createError("REQUEST_TIMEOUT", `Timeout while requesting Drive URL`, {
            fileId,
            url,
            hop,
            timeoutMs: 20000,
          })
        );
      });

      req.on("error", (err) => {
        if (err && err.code && err.details) return reject(err);
        reject(
          createError("NETWORK_ERROR", err.message, {
            fileId,
            url,
            hop,
          })
        );
      });
    }

    doRequest(baseUrl);
  });
}

// ── Detect columns automatically ─────────────────────────────────────────────

function detectColumns(columns) {
  const find = (keywords) =>
    columns.find((c) =>
      keywords.some((k) => c.toLowerCase().includes(k))
    );

  return {
    roll: find(["roll", "enrollment", "enroll", "reg"]),
    name: find(["name", "student", "candidate"]),
    resume: find(["resume", "cv", "drive", "link", "url", "gdrive"]),
  };
}

// ── Main ─────────────────────────────────────────────────────────────────────

async function main() {
  console.log("\n🗂  ZipHire — Resume Bundler\n");
  const startedAt = new Date();

  const diag = [];
  const addDiag = (level, code, message, context = {}, hint = "") => {
    diag.push({
      time: new Date().toISOString(),
      level,
      code,
      message,
      context,
      hint,
    });
  };

  // Parse Excel
  const wb = XLSX.readFile(filePath);
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });

  if (!rows.length) {
    console.error("❌ Excel file is empty.");
    process.exit(1);
  }

  const columns = Object.keys(rows[0]);
  console.log(`📋 Columns found: ${columns.join(", ")}`);

  const detected = detectColumns(columns);

  if (!detected.roll || !detected.name || !detected.resume) {
    console.error("\n❌ Could not auto-detect columns.");
    console.error("   Make sure your Excel has columns like:");
    console.error("   Roll Number / Name / Resume (Drive Link)");
    console.error(`\n   Found: ${columns.join(", ")}`);
    process.exit(1);
  }

  console.log(`✅ Mapped → Roll: "${detected.roll}" | Name: "${detected.name}" | Resume: "${detected.resume}"`);
  console.log(`📦 Total students: ${rows.length}\n`);

  // Output ZIP path
  const outZip = path.join(
    path.dirname(filePath),
    `resumes_${Date.now()}.zip`
  );
  const output = fs.createWriteStream(outZip);
  const archive = archiver("zip", { zlib: { level: 6 } });
  archive.pipe(output);

  archive.on("warning", (err) => {
    addDiag("WARN", "ARCHIVE_WARNING", err.message, {}, "Check ZIP permissions and disk space.");
  });
  archive.on("error", (err) => {
    addDiag("ERROR", "ARCHIVE_ERROR", err.message, {}, "Check ZIP stream and output path permissions.");
  });
  output.on("error", (err) => {
    addDiag("ERROR", "OUTPUT_STREAM_ERROR", err.message, { outZip }, "Ensure output file is writable.");
  });

  const log = [];
  let ok = 0;
  let failed = 0;
  let skipped = 0;

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const roll = sanitize(row[detected.roll]);
    const name = sanitize(row[detected.name]);
    const driveUrl = String(row[detected.resume] || "").trim();
    const filename = `${roll}_${name}.pdf`;
    const prefix = `[${i + 1}/${rows.length}]`;

    if (!roll || !name || !driveUrl) {
      console.log(`${prefix} ⏭  Skipped — missing data`);
      log.push(`SKIP  ${filename} — missing data`);
      addDiag(
        "WARN",
        "MISSING_DATA",
        "Row skipped due to missing roll/name/resume",
        { row: i + 1, roll, name, driveUrl: driveUrl ? "present" : "missing" },
        "Ensure roll, name, and resume URL columns are filled for this student."
      );
      skipped++;
      continue;
    }

    const fileId = extractFileId(driveUrl);
    if (!fileId) {
      console.log(`${prefix} ❌ ${filename} — invalid Drive link: "${driveUrl.substring(0, 60)}"`);
      log.push(`FAIL  ${filename} — invalid Drive link: ${driveUrl}`);
      addDiag(
        "ERROR",
        "INVALID_DRIVE_LINK",
        "Could not extract Drive file ID from URL",
        { row: i + 1, filename, driveUrl },
        "Use a file link like https://drive.google.com/file/d/<FILE_ID>/view or a direct ?id=<FILE_ID> URL."
      );
      failed++;
      continue;
    }

    try {
      process.stdout.write(`${prefix} ⬇  ${filename} ... `);
      const buffer = await downloadFromDrive(fileId);
      archive.append(buffer, { name: filename });
      console.log("✅");
      log.push(`OK    ${filename}`);
      ok++;
    } catch (err) {
      console.log(`❌ ${err.message}`);
      log.push(`FAIL  ${filename} — ${err.message}`);
      addDiag(
        "ERROR",
        err.code || "DOWNLOAD_ERROR",
        err.message,
        {
          row: i + 1,
          filename,
          fileId,
          ...(err.details || {}),
        },
        "Verify the Drive file is accessible, link is valid, and network is stable."
      );
      failed++;
    }
  }

  const runSummary = {
    startedAt: startedAt.toISOString(),
    endedAt: new Date().toISOString(),
    inputFile: path.resolve(filePath),
    outputZip: outZip,
    totalRows: rows.length,
    ok,
    failed,
    skipped,
    diagnosticsCount: diag.length,
  };

  // Append log file inside ZIP
  archive.append(log.join("\n"), { name: "_download_log.txt" });
  archive.append(JSON.stringify(runSummary, null, 2), { name: "_run_summary.json" });
  archive.append(JSON.stringify(diag, null, 2), { name: "_debug_log.json" });

  await new Promise((resolve, reject) => {
    output.once("close", resolve);
    output.once("error", reject);
    archive.once("error", reject);
    archive.finalize().catch(reject);
  });

  console.log(`\n✅ Done — ${ok} downloaded, ${failed} failed, ${skipped} skipped`);
  console.log(`📁 ZIP saved to: ${outZip}\n`);

  if (failed > 0) {
    console.log("ℹ️  Check _debug_log.json inside ZIP for failure reasons and fix hints.");
  }
}

main().catch((e) => {
  console.error("Fatal error:", e.message);
  process.exit(1);
});