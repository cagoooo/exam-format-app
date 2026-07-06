const form = document.querySelector("#convertForm");
const fileInput = document.querySelector("#fileInput");
const fileLabel = document.querySelector("#fileLabel");
const statusBox = document.querySelector("#status");
const preview = document.querySelector("#preview");
const downloadLink = document.querySelector("#downloadLink");
const pdfLink = document.querySelector("#pdfLink");
const pdfPreview = document.querySelector("#pdfPreview");
const pageBadge = document.querySelector("#pageBadge");
const report = document.querySelector("#report");
const analyzeBtn = document.querySelector("#analyzeBtn");
const clearBtn = document.querySelector("#clearBtn");
const API_BASE = (window.EXAM_FORMAT_API_BASE || "").replace(/\/$/, "");
const REQUEST_TIMEOUT_MS = 240000;

function apiUrl(path) {
  if (!API_BASE) return path;
  return `${API_BASE}${path}`;
}

function assetUrl(path) {
  if (!path || !API_BASE || !path.startsWith("/")) return path;
  return `${API_BASE}${path}`;
}

function setStatus(message, mode = "idle") {
  statusBox.textContent = message;
  statusBox.className = `status ${mode}`;
}

function currentFormData() {
  const data = new FormData(form);
  if (fileInput.files[0]) data.set("file", fileInput.files[0]);
  return data;
}

async function fetchWithTimeout(url, options = {}) {
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), REQUEST_TIMEOUT_MS);
  try {
    return await fetch(url, { ...options, signal: controller.signal });
  } catch (error) {
    if (error.name === "AbortError") {
      throw new Error("線上轉換等待逾時。檔案可能太大或圖片太多，請稍後再試，或先壓縮圖片後重新上傳。");
    }
    throw error;
  } finally {
    clearTimeout(timeout);
  }
}

async function readJsonResponse(res) {
  try {
    return await res.json();
  } catch {
    return { error: "伺服器回應格式不正確，請稍後再試。" };
  }
}

function formatExpiry(seconds) {
  if (!seconds) return "";
  const minutes = Math.round(seconds / 60);
  return `下載連結約 ${minutes} 分鐘內有效。`;
}

fileInput.addEventListener("change", () => {
  const file = fileInput.files[0];
  fileLabel.textContent = file ? file.name : "拖曳或選擇考卷 Word 檔";
  resetResult();
  setStatus(file ? "已選擇檔案，可以分析或直接轉換" : "尚未選擇檔案");
});

analyzeBtn.addEventListener("click", async () => {
  if (!fileInput.files[0]) {
    setStatus("請先選擇 Word 檔。", "error");
    return;
  }
  setStatus("正在分析原始檔內容...", "busy");
  preview.innerHTML = "";
  let res, json;
  try {
    res = await fetchWithTimeout(apiUrl("/api/analyze"), { method: "POST", body: currentFormData() });
    json = await readJsonResponse(res);
  } catch (error) {
    setStatus(error.message || "分析失敗。", "error");
    return;
  }
  if (!res.ok) {
    setStatus(json.error || "分析失敗。", "error");
    showDetails(json.details);
    return;
  }
  setStatus(`分析完成：共讀到 ${json.blocks} 個內容區塊。${formatExpiry(json.expires_in_seconds)}`, "ok");
  preview.innerHTML = json.preview.map((line) => `<p>${escapeHtml(line)}</p>`).join("");
});

form.addEventListener("submit", async (event) => {
  event.preventDefault();
  if (!fileInput.files[0]) {
    setStatus("請先選擇 Word 檔。", "error");
    return;
  }
  setStatus("正在套用固定格式並產生 Word...", "busy");
  resetResult({ keepPreview: true });
  let res, json;
  try {
    res = await fetchWithTimeout(apiUrl("/api/convert"), { method: "POST", body: currentFormData() });
    json = await readJsonResponse(res);
  } catch (error) {
    setStatus(error.message || "轉換失敗。", "error");
    return;
  }
  if (!res.ok) {
    setStatus(json.error || "轉換失敗。", "error");
    showDetails(json.details);
    return;
  }
  const pageText = json.pdf_pages ? `，PDF ${json.pdf_pages} 頁` : "";
  const compactText = json.compact_level ? `，已套用第 ${json.compact_level} 級壓縮` : "";
  setStatus(`轉換完成：段落 ${json.stats.paragraphs}、表格 ${json.stats.tables}${pageText}${compactText}。${formatExpiry(json.expires_in_seconds)}`, "ok");
  downloadLink.href = assetUrl(json.download);
  downloadLink.hidden = false;
  if (json.pdf) {
    pdfLink.href = assetUrl(json.pdf);
    pdfLink.hidden = false;
  }
  if (json.preview) {
    pdfPreview.src = assetUrl(json.preview);
    pdfPreview.hidden = false;
  }
  if (json.pdf_pages) {
    pageBadge.textContent = `PDF 頁數：${json.pdf_pages}`;
    pageBadge.className = json.pdf_pages > Number(form.target_pages.value || 0) ? "page-badge warn" : "page-badge ok";
    pageBadge.hidden = false;
  }
  if (json.report?.length || json.warnings?.length) {
    const items = [...(json.report || []), ...(json.warnings || []).map((item) => `提醒：${item}`)];
    report.innerHTML = items.map((item) => `<p>${escapeHtml(item)}</p>`).join("");
    report.hidden = false;
  }
});

clearBtn.addEventListener("click", async () => {
  await fetchWithTimeout(apiUrl("/api/clear"), { method: "POST" });
  resetResult();
  setStatus("暫存檔已清除。", "ok");
});

function resetResult(options = {}) {
  downloadLink.hidden = true;
  pdfLink.hidden = true;
  pdfPreview.hidden = true;
  pageBadge.hidden = true;
  report.hidden = true;
  report.innerHTML = "";
  if (!options.keepPreview) preview.innerHTML = "";
}

function showDetails(details = []) {
  if (!details?.length) return;
  report.innerHTML = details.map((item) => `<p>${escapeHtml(item)}</p>`).join("");
  report.hidden = false;
}

function escapeHtml(value) {
  return value.replace(/[&<>"']/g, (char) => ({
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    '"': "&quot;",
    "'": "&#039;",
  })[char]);
}
