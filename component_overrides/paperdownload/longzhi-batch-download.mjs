import path from "node:path";
import fs from "node:fs/promises";
import { chromium } from "playwright";

const TARGET_PAGE_URL =
  process.env.TARGET_PAGE_URL ??
  "http://longzhi.net.cn/examTest/subjective/queryGrade/15299-85147.mooc?schoolId=&teamId=&selectType=review-select&condition=2";
let USERNAME = process.env.LZ_USERNAME ?? "";
let PASSWORD = process.env.LZ_PASSWORD ?? "";
const CHECK_CODE = process.env.LZ_CHECK_CODE;
const OUTPUT_ROOT = process.env.OUTPUT_ROOT ?? path.join(process.cwd(), "longzhi_batch_output");
const OUTPUT_DIR = process.env.OUTPUT_DIR ?? path.join(OUTPUT_ROOT, "downloads");
const SNAPSHOT_DIR = process.env.SNAPSHOT_DIR ?? path.join(OUTPUT_ROOT, "artifacts");
const STATE_DIR = process.env.STATE_DIR ?? path.join(OUTPUT_ROOT, "state");
const INDEX_FILE = process.env.DOWNLOAD_INDEX_FILE ?? path.join(STATE_DIR, "downloaded_index.json");
const RUN_LOG_FILE = process.env.RUN_LOG_FILE ?? path.join(STATE_DIR, "run_log.jsonl");
const INDEX_CSV_FILE = process.env.DOWNLOAD_INDEX_CSV ?? path.join(STATE_DIR, "downloaded_students.csv");
const MAX_STUDENTS = Number(process.env.MAX_STUDENTS ?? "0");
const PAGE_SIZE = Number(process.env.PAGE_SIZE ?? "100");
const START_PAGE = Number(process.env.START_PAGE ?? "1");
const REVIEW_ENTER_WAIT_MS = Number(process.env.REVIEW_ENTER_WAIT_MS ?? "9000");
const POST_VISIBLE_WAIT_MS = Number(process.env.POST_VISIBLE_WAIT_MS ?? "3000");
const HUMAN_READY_WAIT_MIN_MS = Number(process.env.HUMAN_READY_WAIT_MIN_MS ?? "1000");
const HUMAN_READY_WAIT_MAX_MS = Number(process.env.HUMAN_READY_WAIT_MAX_MS ?? "1800");
const DOWNLOAD_STEP_WAIT_MIN_MS = Number(process.env.DOWNLOAD_STEP_WAIT_MIN_MS ?? "450");
const DOWNLOAD_STEP_WAIT_MAX_MS = Number(process.env.DOWNLOAD_STEP_WAIT_MAX_MS ?? "1100");
const POST_ALL_DOWNLOAD_WAIT_MIN_MS = Number(process.env.POST_ALL_DOWNLOAD_WAIT_MIN_MS ?? "500");
const POST_ALL_DOWNLOAD_WAIT_MAX_MS = Number(process.env.POST_ALL_DOWNLOAD_WAIT_MAX_MS ?? "1200");
const RETURN_LIST_WAIT_MIN_MS = Number(process.env.RETURN_LIST_WAIT_MIN_MS ?? "600");
const RETURN_LIST_WAIT_MAX_MS = Number(process.env.RETURN_LIST_WAIT_MAX_MS ?? "1400");
const DOWNLOAD_REQUEST_TIMEOUT_MS = Number(process.env.DOWNLOAD_REQUEST_TIMEOUT_MS ?? "120000");
const DOWNLOAD_RETRY_COUNT = Number(process.env.DOWNLOAD_RETRY_COUNT ?? "2");
const NAV_TIMEOUT_MS = Number(process.env.NAV_TIMEOUT_MS ?? "90000");
const NAV_RETRY_COUNT = Number(process.env.NAV_RETRY_COUNT ?? "2");
let credentialLoadPromise = null;

function clampRange(minValue, maxValue) {
  const min = Number.isFinite(minValue) ? Math.max(0, Math.floor(minValue)) : 0;
  const maxRaw = Number.isFinite(maxValue) ? Math.max(0, Math.floor(maxValue)) : min;
  const max = Math.max(min, maxRaw);
  return { min, max };
}

function randomBetween(minValue, maxValue) {
  const { min, max } = clampRange(minValue, maxValue);
  if (max <= min) return min;
  return min + Math.floor(Math.random() * (max - min + 1));
}

async function humanPause(page, minMs, maxMs) {
  const waitMs = randomBetween(minMs, maxMs);
  if (waitMs > 0) {
    await page.waitForTimeout(waitMs);
  }
}

function readCredentialFromStdin() {
  return new Promise((resolve, reject) => {
    const chunks = [];
    process.stdin.on("data", (chunk) => chunks.push(Buffer.from(chunk)));
    process.stdin.on("end", () => resolve(Buffer.concat(chunks).toString("utf8")));
    process.stdin.on("error", reject);
    process.stdin.resume();
  });
}

async function ensureCredentialsLoaded() {
  if (USERNAME && PASSWORD) {
    return { username: USERNAME, password: PASSWORD };
  }
  if (process.env.LZ_CREDENTIAL_STDIN !== "1" && process.env.LZ_PASSWORD_STDIN !== "1") {
    return { username: USERNAME, password: PASSWORD };
  }
  if (!credentialLoadPromise) {
    credentialLoadPromise = readCredentialFromStdin().then((rawValue) => {
      if (process.env.LZ_CREDENTIAL_STDIN === "1") {
        const payload = JSON.parse(rawValue);
        USERNAME = USERNAME || String(payload?.username ?? "");
        PASSWORD = PASSWORD || String(payload?.password ?? "");
      } else {
        PASSWORD = rawValue.replace(/[\r\n]+$/, "");
      }
      return { username: USERNAME, password: PASSWORD };
    });
  }
  return credentialLoadPromise;
}

function safePart(text) {
  return String(text ?? "")
    .replace(/[<>:"/\\|?*\u0000-\u001F]/g, "_")
    .replace(/\s+/g, "")
    .trim();
}

function extFromContentDisposition(cd) {
  if (!cd) return "";
  const utf8 = cd.match(/filename\*=UTF-8''([^;]+)/i);
  if (utf8?.[1]) return path.extname(decodeURIComponent(utf8[1]));
  const plain = cd.match(/filename="?([^\";]+)"?/i);
  if (plain?.[1]) return path.extname(plain[1]);
  return "";
}

function extFromMime(contentType) {
  const ct = String(contentType || "").toLowerCase();
  if (ct.includes("application/vnd.openxmlformats-officedocument.wordprocessingml.document")) return ".docx";
  if (ct.includes("application/msword")) return ".doc";
  if (ct.includes("application/pdf")) return ".pdf";
  if (ct.includes("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")) return ".xlsx";
  if (ct.includes("application/vnd.ms-excel")) return ".xls";
  return "";
}

function extFromBody(body) {
  if (!body || body.length < 8) return "";
  // Compound file binary: old Office formats (.doc/.xls/.ppt). In this flow we prefer Word.
  if (
    body[0] === 0xd0 &&
    body[1] === 0xcf &&
    body[2] === 0x11 &&
    body[3] === 0xe0 &&
    body[4] === 0xa1 &&
    body[5] === 0xb1 &&
    body[6] === 0x1a &&
    body[7] === 0xe1
  ) {
    return ".doc";
  }
  // ZIP container. Detect Word package by folder marker.
  if (body[0] === 0x50 && body[1] === 0x4b) {
    const probe = body.subarray(0, Math.min(body.length, 500000)).toString("latin1");
    if (probe.includes("word/")) return ".docx";
    if (probe.includes("xl/")) return ".xlsx";
  }
  // PDF magic.
  if (body[0] === 0x25 && body[1] === 0x50 && body[2] === 0x44 && body[3] === 0x46) return ".pdf";
  return "";
}

async function screenshot(page, name) {
  await fs.mkdir(SNAPSHOT_DIR, { recursive: true });
  const p = path.resolve(SNAPSHOT_DIR, name);
  await page.screenshot({ path: p, fullPage: true });
  return p;
}

async function readJsonFile(filePath, fallbackValue) {
  try {
    const raw = await fs.readFile(filePath, "utf8");
    return JSON.parse(raw);
  } catch {
    return fallbackValue;
  }
}

async function writeJsonAtomic(filePath, value) {
  const tempPath = `${filePath}.tmp`;
  await fs.writeFile(tempPath, JSON.stringify(value, null, 2), "utf8");
  await fs.rename(tempPath, filePath);
}

function parseDownloadedFilename(fileName) {
  const ext = path.extname(fileName);
  const stem = fileName.slice(0, -ext.length);
  const stemNoSuffix = stem.replace(/_\d+$/, "");
  const m = stemNoSuffix.match(/^(\d{5,20})_(.+)$/);
  if (!m) return null;
  return {
    sid: m[1],
    name: m[2],
    base: `${m[1]}_${m[2]}`,
    file: fileName,
  };
}

async function ensureState() {
  await fs.mkdir(STATE_DIR, { recursive: true });
  const defaultIndex = {
    version: 1,
    createdAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
    students: {},
  };
  const index = await readJsonFile(INDEX_FILE, defaultIndex);
  if (!index.students || typeof index.students !== "object") {
    index.students = {};
  }

  // Bootstrap index from existing downloaded files so old data is not redownloaded.
  const existingFiles = await fs.readdir(OUTPUT_DIR).catch(() => []);
  for (const f of existingFiles) {
    const parsed = parseDownloadedFilename(f);
    if (!parsed) continue;
    if (!index.students[parsed.base]) {
      const now = new Date().toISOString();
      index.students[parsed.base] = {
        sid: parsed.sid,
        name: parsed.name,
        files: [parsed.file],
        firstDownloadedAt: now,
        lastDownloadedAt: now,
        lastRunId: "bootstrap-from-filesystem",
        totalFiles: 1,
      };
    } else if (!index.students[parsed.base].files.includes(parsed.file)) {
      index.students[parsed.base].files.push(parsed.file);
      index.students[parsed.base].totalFiles = index.students[parsed.base].files.length;
    }
  }
  index.updatedAt = new Date().toISOString();
  await writeJsonAtomic(INDEX_FILE, index);
  return index;
}

async function appendRunLog(record) {
  const enriched = {
    recordedAt: new Date().toISOString(),
    ...record,
  };
  await fs.appendFile(RUN_LOG_FILE, `${JSON.stringify(enriched)}\n`, "utf8");
}

function csvEscape(value) {
  const text = String(value ?? "");
  if (text.includes(",") || text.includes("\"") || text.includes("\n")) {
    return `"${text.replace(/"/g, "\"\"")}"`;
  }
  return text;
}

async function writeIndexCsv(index) {
  const rows = [["sid", "name", "base", "file_count", "last_downloaded_at", "files"]];
  for (const [base, item] of Object.entries(index.students)) {
    rows.push([
      item.sid || "",
      item.name || "",
      base,
      String(item.totalFiles || item.files?.length || 0),
      item.lastDownloadedAt || "",
      (item.files || []).join("|"),
    ]);
  }
  const content = rows.map((r) => r.map(csvEscape).join(",")).join("\n");
  await fs.writeFile(INDEX_CSV_FILE, content, "utf8");
}

async function loginIfNeeded(page) {
  if (!page.url().includes("/home/login.mooc")) return;
  const credentials = await ensureCredentialsLoaded();
  const dismissBlockingPrompts = async () => {
    let handled = false;
    for (let i = 0; i < 4; i += 1) {
      const okBtn = page
        .locator("div.d-buttons input[type='button'][value='确定'],div.d-buttons input.d-state-highlight")
        .first();
      const cancelBtn = page
        .locator("div.d-buttons input[type='button'][value='取消'],a.d-close")
        .first();
      if (await okBtn.isVisible().catch(() => false)) {
        await okBtn.click().catch(() => {});
        await page.waitForTimeout(700);
        handled = true;
        continue;
      }
      if (await cancelBtn.isVisible().catch(() => false)) {
        await cancelBtn.click().catch(() => {});
        await page.waitForTimeout(700);
        handled = true;
        continue;
      }
      break;
    }
    return handled;
  };

  const fillAndSubmit = async () => {
    await dismissBlockingPrompts();
    await page.locator("#loginName").fill(credentials.username ?? "");
    await page.locator("#password").fill(credentials.password ?? "");
    const captchaInput = page.locator("#checkCode");
    if ((await captchaInput.count()) > 0 && (await captchaInput.isVisible())) {
      if (!CHECK_CODE) throw new Error("检测到验证码，请设置 LZ_CHECK_CODE 后再运行。");
      await captchaInput.fill(CHECK_CODE);
    }
    await Promise.all([
      page.waitForLoadState("networkidle").catch(() => {}),
      page.evaluate(() => {
        const btn = document.querySelector("#userLogin");
        if (btn) btn.click();
      }),
    ]);
    await page.waitForTimeout(1000);
  };

  const clickConfirmIfVisible = async () => {
    const okBtn = page
      .locator("div.d-buttons input[type='button'][value='确定'],div.d-buttons input.d-state-highlight")
      .first();
    if (await okBtn.isVisible().catch(() => false)) {
      await okBtn.click().catch(() => {});
      await page.waitForTimeout(1000);
      return true;
    }
    return false;
  };

  const clickCancelIfVisible = async () => {
    const cancelBtn = page
      .locator("div.d-buttons input[type='button'][value='取消'],a.d-close")
      .first();
    if (await cancelBtn.isVisible().catch(() => false)) {
      await cancelBtn.click().catch(() => {});
      await page.waitForTimeout(900);
      return true;
    }
    return false;
  };

  // Retry login attempts because the site may reset fields after confirm dialogs.
  for (let attempt = 1; attempt <= 4; attempt += 1) {
    await fillAndSubmit();
    await dismissBlockingPrompts();
    const stillOnLogin = await page.locator("#loginName").isVisible().catch(() => false);
    if (!stillOnLogin) {
      // Weak-password or generic dialogs can appear after successful auth; dismiss and continue.
      for (let i = 0; i < 4; i += 1) {
        const handledCancel = await clickCancelIfVisible();
        const handledConfirm = await clickConfirmIfVisible();
        if (!handledCancel && !handledConfirm) break;
      }
      return;
    }
    await page.waitForTimeout(900);
  }
  throw new Error("登录未成功：多次重试后仍停留在登录页。");
}

async function setPageSize50(page) {
  const result = await page.evaluate((desiredSize) => {
    const choose = (values, wanted) => {
      const nums = values.map((v) => Number(v)).filter((n) => Number.isFinite(n));
      if (!nums.length) return "";
      let target = Number(wanted);
      if (!nums.includes(target)) {
        target = Math.max(...nums);
      }
      return String(target);
    };

    // Preferred path: visible custom dropdown in pager.
    const pagerSelect = document.querySelector("#pageId .dk-select.page-select");
    if (pagerSelect) {
      const options = Array.from(pagerSelect.querySelectorAll("li.dk-option")).map(
        (li) => (li.getAttribute("data-value") || "").trim()
      );
      const target = choose(options, desiredSize);
      if (target) {
        const selected = pagerSelect.querySelector(".dk-selected");
        const opt = pagerSelect.querySelector(`li.dk-option[data-value="${target}"]`);
        if (selected && opt) {
          selected.dispatchEvent(new MouseEvent("click", { bubbles: true }));
          opt.dispatchEvent(new MouseEvent("click", { bubbles: true }));
        }
      }
      const selectedText =
        (pagerSelect.querySelector(".dk-selected")?.textContent || "").trim();
      return { ok: true, selected: selectedText, options };
    }

    // Fallback: native select.
    const select = document.querySelector("select.page-select");
    if (!select) return { ok: false, selected: "", options: [] };
    const options = Array.from(select.options).map((o) => (o.value || "").trim());
    const target = choose(options, desiredSize);
    if (target && select.value !== target) {
      select.value = target;
      select.dispatchEvent(new Event("change", { bubbles: true }));
    }
    return { ok: true, selected: String(select.value || ""), options };
  }, PAGE_SIZE);

  if (result.ok) {
    await humanPause(page, 350, 900);
  }
  return result;
}

async function collectRows(page) {
  return page.evaluate(() => {
    const tables = Array.from(document.querySelectorAll("table"));
    let rows = [];
    for (const table of tables) {
      const bodyRows = Array.from(table.querySelectorAll("tbody tr"));
      for (const tr of bodyRows) {
        const cells = Array.from(tr.querySelectorAll("td"));
        if (cells.length < 3) continue;
        const reviewLink = Array.from(tr.querySelectorAll("a,button")).find((el) =>
          (el.textContent || "").includes("批阅")
        );
        if (!reviewLink) continue;
        const sid = (cells[1]?.textContent || "").trim();
        const name = (cells[2]?.textContent || "").trim();
        const href = reviewLink.getAttribute("href") || "";
        rows.push({
          sid,
          name,
          href,
          onclick: reviewLink.getAttribute("onclick") || "",
          submitId: reviewLink.getAttribute("submit_id") || "",
          reviewClass: reviewLink.getAttribute("class") || "",
          reviewData: Object.fromEntries(
            Array.from(reviewLink.attributes)
              .filter((a) => a.name.startsWith("data-"))
              .map((a) => [a.name, a.value])
          ),
          reviewHtml: reviewLink.outerHTML.slice(0, 500),
          text: (tr.textContent || "").replace(/\s+/g, " ").trim(),
        });
      }
      if (rows.length) break;
    }
    return rows;
  });
}

async function waitListReady(page) {
  await page.waitForLoadState("domcontentloaded").catch(() => {});
  await page.waitForLoadState("networkidle").catch(() => {});
  await page
    .waitForFunction(() => {
      const loadingSelectors = [
        ".loading",
        ".loading-mask",
        ".loading-msg",
        ".layui-layer-loading",
        ".spinner",
      ];
      for (const sel of loadingSelectors) {
        const el = document.querySelector(sel);
        if (el && el.offsetParent !== null) return false;
      }
      const loadingTextNode = Array.from(document.querySelectorAll("div,span,p")).find((el) =>
        /正在加载|加载中|请稍候/.test((el.textContent || "").trim())
      );
      if (loadingTextNode && loadingTextNode.offsetParent !== null) return false;
      return true;
    }, { timeout: 15000 })
    .catch(() => {});

  await page
    .waitForFunction(() => {
      const rows = document.querySelectorAll("table tbody tr");
      return rows.length > 0;
    }, { timeout: 15000 })
    .catch(() => {});

  await humanPause(page, 300, 800);
}

async function smartWaitForVisible(locator, maxMs) {
  await locator.waitFor({ state: "visible", timeout: maxMs }).catch(() => {});
}

async function waitReviewDownloadReady(page, links) {
  const startedAt = Date.now();
  await smartWaitForVisible(links.first(), REVIEW_ENTER_WAIT_MS);
  await page.waitForLoadState("domcontentloaded").catch(() => {});
  await page.waitForLoadState("networkidle").catch(() => {});
  await page
    .waitForFunction(() => {
      const loadingSelectors = [
        ".loading",
        ".loading-mask",
        ".loading-msg",
        ".layui-layer-loading",
        ".spinner",
      ];
      for (const sel of loadingSelectors) {
        const el = document.querySelector(sel);
        if (el && el.offsetParent !== null) return false;
      }
      const anchors = Array.from(document.querySelectorAll("a"));
      return anchors.some((a) => {
        if (a.offsetParent === null) return false;
        const text = (a.textContent || "").trim();
        if (!/下载|导出/.test(text)) return false;
        const href = (a.getAttribute("href") || "").trim();
        return href && !href.toLowerCase().startsWith("javascript:");
      });
    }, { timeout: POST_VISIBLE_WAIT_MS })
    .catch(() => {});
  await humanPause(page, HUMAN_READY_WAIT_MIN_MS, HUMAN_READY_WAIT_MAX_MS);
  return Date.now() - startedAt;
}

async function fetchFileWithRetry(context, url) {
  let lastError;
  for (let i = 0; i < DOWNLOAD_RETRY_COUNT; i += 1) {
    try {
      const res = await context.request.get(url, { timeout: DOWNLOAD_REQUEST_TIMEOUT_MS });
      if (!res.ok()) {
        throw new Error(`HTTP ${res.status()} ${url}`);
      }
      const body = await res.body();
      return { res, body };
    } catch (err) {
      lastError = err;
      if (i < DOWNLOAD_RETRY_COUNT - 1) {
        await new Promise((r) => setTimeout(r, 1200));
      }
    }
  }
  throw lastError;
}

async function gotoWithRetry(page, url, waitUntil = "domcontentloaded") {
  let lastError;
  for (let i = 0; i < NAV_RETRY_COUNT; i += 1) {
    try {
      await page.goto(url, { waitUntil, timeout: NAV_TIMEOUT_MS });
      return;
    } catch (err) {
      lastError = err;
      if (i < NAV_RETRY_COUNT - 1) {
        await page.waitForTimeout(1500);
      }
    }
  }
  throw lastError;
}

async function gotoNextPage(page) {
  const candidates = [
    "a:has-text('下一页')",
    "button:has-text('下一页')",
    "a.next",
    ".next a",
    "a[title*='下一']",
    "a:has-text('>')",
  ];
  for (const sel of candidates) {
    const locator = page.locator(sel).first();
    if ((await locator.count()) === 0) continue;
    if (!(await locator.isVisible().catch(() => false))) continue;
    const cls = (await locator.getAttribute("class")) || "";
    const disabled = /disabled|ban|forbid/i.test(cls);
    if (disabled) continue;
    await Promise.all([
      page.waitForLoadState("domcontentloaded").catch(() => {}),
      locator.click(),
    ]);
    await waitListReady(page);
    await humanPause(page, 300, 700);
    return true;
  }
  return false;
}

async function getCurrentAndTotalPage(page) {
  return page.evaluate(() => {
    const current =
      (document.querySelector("li.page-num.current")?.textContent || "1").trim() || "1";
    const totalText = (document.querySelector(".page-total")?.textContent || "").trim();
    const m = totalText.match(/(\d+)/);
    const total = m ? Number(m[1]) : 1;
    return {
      current: Number(current) || 1,
      total: Number.isFinite(total) && total > 0 ? total : 1,
    };
  });
}

async function gotoPage(page, pageNo) {
  if (pageNo <= 1) return true;
  const clicked = await page.evaluate((targetPage) => {
    const container = document.querySelector("#pageId") || document;
    const nums = Array.from(container.querySelectorAll("li.page-num"));
    for (const li of nums) {
      const t = (li.textContent || "").trim();
      if (t === String(targetPage)) {
        const a = li.querySelector("a");
        (a || li).dispatchEvent(new MouseEvent("click", { bubbles: true }));
        return true;
      }
    }
    return false;
  }, pageNo);
  if (!clicked) return false;
  await waitListReady(page);
  await humanPause(page, 250, 700);
  return true;
}

async function saveFromCurrentReviewPage(page, context, sid, name, indexBase) {
  const startedAtMs = Date.now();
  const links = page.locator("a:has-text('下载'):visible,a:has-text('导出'):visible");
  // Wait until review page controls are truly ready, then pause briefly like a human.
  const waitReviewControlsMs = await waitReviewDownloadReady(page, links);
  const count = await links.count();
  if (!count) {
    return {
      saved: 0,
      files: [],
      timing: {
        wait_review_controls_ms: waitReviewControlsMs,
        fetch_and_write_total_ms: 0,
        post_download_wait_ms: 0,
        total_ms: Date.now() - startedAtMs,
        file_timings_ms: [],
      },
    };
  }

  const files = [];
  const fileTimingsMs = [];
  const fetchAndWriteStartedAtMs = Date.now();
  for (let i = 0; i < count; i += 1) {
    const fileStartedAtMs = Date.now();
    const link = links.nth(i);
    const href = await link.getAttribute("href");
    if (!href || href.toLowerCase().startsWith("javascript:")) continue;
    const fileUrl = new URL(href, page.url()).toString();
    const { res, body } = await fetchFileWithRetry(context, fileUrl);
    const cdExt = extFromContentDisposition(res.headers()["content-disposition"]);
    const mimeExt = extFromMime(res.headers()["content-type"]);
    const bodyExt = extFromBody(body);
    const urlExt = path.extname(fileUrl);
    let ext = cdExt || mimeExt || bodyExt || urlExt || ".bin";
    if (ext.toLowerCase() === ".mooc") {
      ext = bodyExt || mimeExt || ".doc";
    }
    const base = `${safePart(sid)}_${safePart(name)}`;
    const suffix = i === 0 ? "" : `_${i + 1}`;
    const filename = `${base}${suffix}${ext}`;
    const savePath = path.resolve(OUTPUT_DIR, filename);
    await fs.mkdir(OUTPUT_DIR, { recursive: true });
    await fs.writeFile(savePath, body);
    files.push(savePath);
    fileTimingsMs.push({
      file: path.basename(savePath),
      elapsed_ms: Date.now() - fileStartedAtMs,
    });
    await humanPause(page, DOWNLOAD_STEP_WAIT_MIN_MS, DOWNLOAD_STEP_WAIT_MAX_MS);
  }
  const fetchAndWriteTotalMs = Date.now() - fetchAndWriteStartedAtMs;
  const postDownloadWaitStartedAtMs = Date.now();
  await humanPause(page, POST_ALL_DOWNLOAD_WAIT_MIN_MS, POST_ALL_DOWNLOAD_WAIT_MAX_MS);
  const postDownloadWaitMs = Date.now() - postDownloadWaitStartedAtMs;
  if (indexBase <= 3) {
    await screenshot(page, `review-sample-${indexBase}.png`);
  }
  return {
    saved: files.length,
    files,
    timing: {
      wait_review_controls_ms: waitReviewControlsMs,
      fetch_and_write_total_ms: fetchAndWriteTotalMs,
      post_download_wait_ms: postDownloadWaitMs,
      total_ms: Date.now() - startedAtMs,
      file_timings_ms: fileTimingsMs,
    },
  };
}

async function main() {
  const credentials = await ensureCredentialsLoaded();
  if (!credentials.username || !credentials.password) {
    throw new Error("Missing LZ_USERNAME/LZ_PASSWORD.");
  }

  await fs.mkdir(OUTPUT_DIR, { recursive: true });
  await fs.mkdir(SNAPSHOT_DIR, { recursive: true });
  await fs.mkdir(STATE_DIR, { recursive: true });

  const browser = await chromium.launch({ headless: false });
  const context = await browser.newContext({ acceptDownloads: true });
  const listPage = await context.newPage();

  const runId = `run-${new Date().toISOString().replace(/[:.]/g, "-")}`;
  const runStartedAt = new Date().toISOString();
  const runStartedAtMs = Date.now();
  let runStatus = "success";
  let runError = "";
  const runPhaseTimings = {};

  let processed = 0;
  let downloaded = 0;
  let skippedExisting = 0;
  let failed = 0;
  let noAttachment = 0;
  const seenStudents = new Set();
  const index = await ensureState();
  const existingBases = new Set(Object.keys(index.students));

  await appendRunLog({
    type: "run_start",
    runId,
    startedAt: runStartedAt,
    targetPage: TARGET_PAGE_URL,
    pageSize: PAGE_SIZE,
    startPage: START_PAGE,
  });

  try {
    const openTargetStartedAtMs = Date.now();
    await gotoWithRetry(listPage, TARGET_PAGE_URL, "domcontentloaded");
    runPhaseTimings.open_target_page_ms = Date.now() - openTargetStartedAtMs;
    await appendRunLog({ type: "phase_timing", runId, phase: "open_target_page", elapsed_ms: runPhaseTimings.open_target_page_ms });
    await screenshot(listPage, "list-01-open.png");
    const loginStartedAtMs = Date.now();
    await loginIfNeeded(listPage);
    runPhaseTimings.login_if_needed_ms = Date.now() - loginStartedAtMs;
    await appendRunLog({ type: "phase_timing", runId, phase: "login_if_needed", elapsed_ms: runPhaseTimings.login_if_needed_ms });
    const backToTargetStartedAtMs = Date.now();
    await gotoWithRetry(listPage, TARGET_PAGE_URL, "domcontentloaded");
    await waitListReady(listPage);
    runPhaseTimings.return_to_target_after_login_ms = Date.now() - backToTargetStartedAtMs;
    await appendRunLog({ type: "phase_timing", runId, phase: "return_to_target_after_login", elapsed_ms: runPhaseTimings.return_to_target_after_login_ms });
    await screenshot(listPage, "list-02-after-login.png");

    const setPageSizeStartedAtMs = Date.now();
    const pageSizeState = await setPageSize50(listPage);
    await waitListReady(listPage);
    runPhaseTimings.set_page_size_ms = Date.now() - setPageSizeStartedAtMs;
    await appendRunLog({ type: "phase_timing", runId, phase: "set_page_size", elapsed_ms: runPhaseTimings.set_page_size_ms });
    console.log(
      `Page size set: ${pageSizeState.ok ? "ok" : "failed"}, current=${pageSizeState.selected || "unknown"}, options=${(pageSizeState.options || []).join("/")}`
    );
    await screenshot(listPage, "list-03-page-size.png");

    let pageNo = START_PAGE > 1 ? START_PAGE : 1;
    if (pageNo > 1) {
      const gotoStartPageStartedAtMs = Date.now();
      const moved = await gotoPage(listPage, pageNo);
      if (!moved) {
        console.log(`Cannot jump to page ${pageNo}; fallback to page 1.`);
        pageNo = 1;
      }
      runPhaseTimings.goto_start_page_ms = Date.now() - gotoStartPageStartedAtMs;
      await appendRunLog({ type: "phase_timing", runId, phase: "goto_start_page", elapsed_ms: runPhaseTimings.goto_start_page_ms });
    }
    let stopEarly = false;
    const scanAndProcessStartedAtMs = Date.now();

    outerLoop:
    while (true) {
      const rows = await collectRows(listPage);
      console.log(`Page ${pageNo}: ${rows.length} rows detected.`);
      if (!rows.length) break;

      for (const row of rows) {
        const sid = safePart(row.sid);
        const name = safePart(row.name);
        if (!sid || !name) continue;
        const key = `${sid}_${name}`;
        const base = `${sid}_${name}`;
        if (existingBases.has(base)) {
          skippedExisting += 1;
          await appendRunLog({
            type: "student_skip_existing",
            runId,
            page: pageNo,
            sid,
            name,
            base,
          });
          continue;
        }
        if (seenStudents.has(key)) continue;
        seenStudents.add(key);

        if (!row.submitId) {
          failed += 1;
          await appendRunLog({
            type: "student_skip_missing_submit_id",
            runId,
            page: pageNo,
            sid,
            name,
            base,
          });
          console.log(`Skip missing submit_id: ${row.sid}_${row.name}`);
          continue;
        }

        processed += 1;
        const studentStartedAtMs = Date.now();
        let enterReviewMs = 0;
        let saveFromReviewMs = 0;
        let returnToListMs = 0;
        let returnToListMethod = "unknown";
        let studentOutcome = "unknown";
        let studentSavedCount = 0;
        let saveFromReviewDetailMs = {};
        try {
          const enterReviewStartedAtMs = Date.now();
          const clicked = await listPage.evaluate((submitId) => {
            const a = document.querySelector(`a.review_a[submit_id="${submitId}"]`);
            if (!a) return false;
            a.click();
            return true;
          }, row.submitId);
          enterReviewMs = Date.now() - enterReviewStartedAtMs;
          if (!clicked) {
            studentOutcome = "click_failed";
            failed += 1;
            await appendRunLog({
              type: "student_click_failed",
              runId,
              page: pageNo,
              sid,
              name,
              base,
              submitId: row.submitId,
            });
            console.log(`Click failed: submit_id=${row.submitId}`);
            continue;
          }
          await listPage.waitForLoadState("domcontentloaded").catch(() => {});
          await listPage.waitForLoadState("networkidle").catch(() => {});
          await humanPause(listPage, 600, 1200);
          const saveFromReviewStartedAtMs = Date.now();
          const result = await saveFromCurrentReviewPage(listPage, context, sid, name, processed);
          saveFromReviewMs = Date.now() - saveFromReviewStartedAtMs;
          saveFromReviewDetailMs = result.timing || {};
          studentSavedCount = Number(result.saved || 0);
          downloaded += result.saved;
          if (result.saved > 0) {
            studentOutcome = "downloaded";
            existingBases.add(base);
            const fileNames = result.files.map((f) => path.basename(f));
            const now = new Date().toISOString();
            const prev = index.students[base] ?? {
              sid,
              name,
              files: [],
              firstDownloadedAt: now,
              lastDownloadedAt: now,
              lastRunId: runId,
              totalFiles: 0,
            };
            prev.sid = sid;
            prev.name = name;
            prev.lastDownloadedAt = now;
            prev.lastRunId = runId;
            prev.files = Array.from(new Set([...(prev.files || []), ...fileNames]));
            prev.totalFiles = prev.files.length;
            index.students[base] = prev;

            await appendRunLog({
              type: "student_downloaded",
              runId,
              page: pageNo,
              sid,
              name,
              base,
              fileCount: result.saved,
              files: fileNames,
              timing: {
                student_total_ms: Date.now() - studentStartedAtMs,
                enter_review_ms: enterReviewMs,
                save_from_review_ms: saveFromReviewMs,
                save_from_review_detail_ms: result.timing || {},
                return_to_list_ms: returnToListMs,
              },
            });
          } else {
            studentOutcome = "no_attachment";
            noAttachment += 1;
            await appendRunLog({
              type: "student_no_attachment",
              runId,
              page: pageNo,
              sid,
              name,
              base,
              timing: {
                student_total_ms: Date.now() - studentStartedAtMs,
                enter_review_ms: enterReviewMs,
                save_from_review_ms: saveFromReviewMs,
                save_from_review_detail_ms: result.timing || {},
                return_to_list_ms: returnToListMs,
              },
            });
          }
          console.log(`Processed ${processed}: ${sid}_${name} -> downloaded ${result.saved}`);
        } catch (err) {
          studentOutcome = "failed";
          failed += 1;
          await appendRunLog({
            type: "student_failed",
            runId,
            page: pageNo,
            sid,
            name,
            base,
            error: err.message,
            timing: {
              student_total_ms: Date.now() - studentStartedAtMs,
              enter_review_ms: enterReviewMs,
              save_from_review_ms: saveFromReviewMs,
              return_to_list_ms: returnToListMs,
            },
          });
          console.log(`Processed ${processed}: ${sid}_${name} failed: ${err.message}`);
        } finally {
          const returnToListStartedAtMs = Date.now();
          returnToListMethod = "go_back";
          let listRecovered = false;
          try {
            await listPage.goBack({ waitUntil: "domcontentloaded", timeout: 20000 });
            await waitListReady(listPage);
            listRecovered = await listPage
              .evaluate(() => document.querySelectorAll("table tbody tr").length > 0)
              .catch(() => false);
          } catch {
            listRecovered = false;
          }
          if (!listRecovered) {
            returnToListMethod = "goto_target";
            await gotoWithRetry(listPage, TARGET_PAGE_URL, "domcontentloaded");
            await waitListReady(listPage);
          }
          await setPageSize50(listPage);
          await waitListReady(listPage);
          if (pageNo > 1) {
            await gotoPage(listPage, pageNo);
            await waitListReady(listPage);
          }
          await humanPause(listPage, RETURN_LIST_WAIT_MIN_MS, RETURN_LIST_WAIT_MAX_MS);
          returnToListMs = Date.now() - returnToListStartedAtMs;
          await appendRunLog({
            type: "student_timing",
            runId,
            page: pageNo,
            sid,
            name,
            base,
            outcome: studentOutcome,
            fileCount: studentSavedCount,
            timing: {
              student_total_ms: Date.now() - studentStartedAtMs,
              enter_review_ms: enterReviewMs,
              save_from_review_ms: saveFromReviewMs,
              save_from_review_detail_ms: saveFromReviewDetailMs,
              return_to_list_ms: returnToListMs,
              return_to_list_method: returnToListMethod,
            },
          });
        }

        if (MAX_STUDENTS > 0 && processed >= MAX_STUDENTS) {
          stopEarly = true;
          console.log(`Reached MAX_STUDENTS=${MAX_STUDENTS}, stop early.`);
          break outerLoop;
        }
      }

      const moved = await gotoNextPage(listPage);
      if (!moved) break;
      pageNo += 1;
      await screenshot(listPage, `list-page-${pageNo}.png`);
    }
    runPhaseTimings.scan_and_process_ms = Date.now() - scanAndProcessStartedAtMs;
    await appendRunLog({ type: "phase_timing", runId, phase: "scan_and_process", elapsed_ms: runPhaseTimings.scan_and_process_ms });

    console.log(
      `Batch finished: processed=${processed}, downloaded=${downloaded}, skipped_existing=${skippedExisting}, no_attachment=${noAttachment}, failed=${failed}${stopEarly ? ", stopped_early=true" : ""}`
    );
  } catch (err) {
    runStatus = "failed";
    runError = err.message;
    throw err;
  } finally {
    index.updatedAt = new Date().toISOString();
    await writeJsonAtomic(INDEX_FILE, index);
    await writeIndexCsv(index);
    await appendRunLog({
      type: "run_end",
      runId,
      status: runStatus,
      error: runError,
      startedAt: runStartedAt,
      endedAt: new Date().toISOString(),
      timing: {
        run_total_ms: Date.now() - runStartedAtMs,
        phases_ms: runPhaseTimings,
      },
      summary: {
        processed,
        downloaded,
        skippedExisting,
        noAttachment,
        failed,
        indexedStudents: Object.keys(index.students || {}).length,
      },
    });
    await context.close();
    await browser.close();
  }
}

main().catch((err) => {
  console.error("Run failed:", err.message);
  process.exit(1);
});
