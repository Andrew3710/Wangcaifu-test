const fileInput = document.querySelector("#file-input");
const pickFileBtn = document.querySelector("#pick-file");
const dropZone = document.querySelector("#drop-zone");
const fileName = document.querySelector("#file-name");
const anchorDateInput = document.querySelector("#anchor-date");
const startDateInput = document.querySelector("#start-date");
const endDateInput = document.querySelector("#end-date");
const generateBtn = document.querySelector("#generate-btn");
const resetBtn = document.querySelector("#reset-btn");
const copyBtn = document.querySelector("#copy-btn");
const downloadBtn = document.querySelector("#download-btn");
const downloadCirculationBtn = document.querySelector("#download-circulation-btn");
const statusEl = document.querySelector("#status");
const resultOutput = document.querySelector("#result-output");
const resultMeta = document.querySelector("#result-meta");
const summaryGrid = document.querySelector("#summary-grid");
const modeOptions = [...document.querySelectorAll(".mode-option")];

let selectedFile = null;
let lastOutput = "";
let lastCirculationRows = [];

initDates();
bindEvents();

function bindEvents() {
  pickFileBtn.addEventListener("click", () => fileInput.click());
  fileInput.addEventListener("change", () => {
    if (fileInput.files?.[0]) {
      setSelectedFile(fileInput.files[0]);
    }
  });

  dropZone.addEventListener("dragover", (event) => {
    event.preventDefault();
    dropZone.classList.add("dragover");
  });
  dropZone.addEventListener("dragleave", () => {
    dropZone.classList.remove("dragover");
  });
  dropZone.addEventListener("drop", (event) => {
    event.preventDefault();
    dropZone.classList.remove("dragover");
    const file = event.dataTransfer?.files?.[0];
    if (file) {
      setSelectedFile(file);
    }
  });

  document.querySelectorAll('input[name="mode"]').forEach((radio) => {
    radio.addEventListener("change", updateModeState);
  });

  document.querySelector("#fill-this-week").addEventListener("click", fillThisWeek);
  document.querySelector("#fill-last-week").addEventListener("click", fillLastWeek);
  generateBtn.addEventListener("click", submitReport);
  resetBtn.addEventListener("click", resetView);
  copyBtn.addEventListener("click", copyResult);
  downloadBtn.addEventListener("click", downloadResult);
  downloadCirculationBtn.addEventListener("click", downloadCirculation);
}

function initDates() {
  const today = new Date();
  anchorDateInput.value = formatDate(today);
  fillThisWeek();
  updateModeState();
}

function fillThisWeek() {
  const today = new Date();
  const day = today.getDay() || 7;
  const start = new Date(today);
  start.setDate(today.getDate() - day + 1);
  const end = new Date(start);
  end.setDate(start.getDate() + 6);

  startDateInput.value = formatDate(start);
  endDateInput.value = formatDate(end);
  setMode("custom");
}

function fillLastWeek() {
  const today = new Date();
  const day = today.getDay() || 7;
  const start = new Date(today);
  start.setDate(today.getDate() - day - 6);
  const end = new Date(start);
  end.setDate(start.getDate() + 6);

  startDateInput.value = formatDate(start);
  endDateInput.value = formatDate(end);
  setMode("custom");
}

function setMode(mode) {
  const radio = document.querySelector(`input[name="mode"][value="${mode}"]`);
  if (radio) {
    radio.checked = true;
  }
  updateModeState();
}

function updateModeState() {
  const mode = getMode();
  modeOptions.forEach((label) => {
    const input = label.querySelector("input");
    label.classList.toggle("active", input.checked);
  });

  anchorDateInput.closest(".field-block").style.opacity = mode === "week" ? "1" : "0.7";
  startDateInput.closest(".field-block").style.opacity = mode === "custom" ? "1" : "0.7";
  endDateInput.closest(".field-block").style.opacity = mode === "custom" ? "1" : "0.7";
}

function setSelectedFile(file) {
  selectedFile = file;
  fileName.textContent = `${file.name} · ${(file.size / 1024).toFixed(1)} KB`;
  statusEl.textContent = "文件已选择，可以直接生成统计结果。";
}

function getMode() {
  const checked = document.querySelector('input[name="mode"]:checked');
  return checked ? checked.value : "week";
}

async function submitReport() {
  if (!selectedFile) {
    statusEl.textContent = "请先选择 Excel 文件。";
    return;
  }

  const mode = getMode();
  const formData = new FormData();
  formData.append("file", selectedFile);
  formData.append("mode", mode);
  formData.append("anchor_date", anchorDateInput.value);
  formData.append("start_date", startDateInput.value);
  formData.append("end_date", endDateInput.value);

  setBusy(true, "正在生成统计结果...");
  try {
    const response = await fetch("/api/report", {
      method: "POST",
      body: formData,
    });
    const payload = await response.json();
    if (!response.ok || !payload.ok) {
      throw new Error(payload.error || "生成失败");
    }
    renderResult(payload.data);
    setBusy(false, "统计完成，可以复制或下载结果。");
  } catch (error) {
    setBusy(false, error.message || "生成失败");
  }
}

function renderResult(data) {
  lastOutput = data.output || "";
  resultOutput.textContent = lastOutput || "没有生成结果。";

  const summary = data.summary || {};
  const districtDepts = (summary.new_departments?.["区级"] || [])
    .map((item) => `${item.department}${item.count}张`)
    .join("、") || "无";
  const cityDepts = (summary.new_departments?.["市级"] || [])
    .map((item) => `${item.department}${item.count}张`)
    .join("、") || "无";
  const provinceDepts = (summary.new_departments?.["省级"] || [])
    .map((item) => `${item.department}${item.count}张`)
    .join("、") || "无";
  const fillCounts = summary.fill_counts || {};
  const districtFill = fillCounts["区级"] || {};
  const cityFill = fillCounts["市级"] || {};
  const provinceFill = fillCounts["省级"] || {};
  const newCounts = summary.new_counts || {};

  summaryGrid.innerHTML = [
    summaryCard("区级新增", `${newCounts["区级"] || 0}`, districtDepts),
    summaryCard("市级新增", `${newCounts["市级"] || 0}`, cityDepts),
    summaryCard("省级新增", `${newCounts["省级"] || 0}`, provinceDepts),
    summaryCard("区级填报率", districtFill.rate || "0.00%", `${districtFill.completed_tasks || 0}/${districtFill.total_tasks || 0}`),
    summaryCard("市级填报率", cityFill.rate || "0.00%", `${cityFill.completed_tasks || 0}/${cityFill.total_tasks || 0}`),
    summaryCard("省级填报率", provinceFill.rate || "0.00%", `${provinceFill.completed_tasks || 0}/${provinceFill.total_tasks || 0}`),
  ].join("");
  summaryGrid.classList.remove("hidden");

  resultMeta.innerHTML = `
    <div>统计文件：<strong>${escapeHtml(data.filename || "未命名文件")}</strong></div>
    <div>统计周期：<strong>${escapeHtml(data.start_date)} 至 ${escapeHtml(data.end_date)}</strong></div>
  `;
  resultMeta.classList.remove("hidden");

  lastCirculationRows = summary.circulation_rows || [];
  downloadCirculationBtn.textContent = lastCirculationRows.length
    ? `导出流转明细（${lastCirculationRows.length}）`
    : "导出流转明细";
}

function resetView() {
  selectedFile = null;
  fileInput.value = "";
  fileName.textContent = "当前未选择文件";
  lastOutput = "";
  lastCirculationRows = [];
  resultOutput.textContent = "统计结果会显示在这里。";
  resultMeta.classList.add("hidden");
  resultMeta.innerHTML = "";
  summaryGrid.classList.add("hidden");
  summaryGrid.innerHTML = "";
  downloadCirculationBtn.textContent = "导出流转明细";
  statusEl.textContent = "已清空结果。";
}

async function copyResult() {
  if (!lastOutput) {
    statusEl.textContent = "当前没有可复制的结果。";
    return;
  }

  try {
    await navigator.clipboard.writeText(lastOutput);
    statusEl.textContent = "结果已复制到剪贴板。";
  } catch {
    statusEl.textContent = "复制失败，请手动复制。";
  }
}

function downloadResult() {
  if (!lastOutput) {
    statusEl.textContent = "当前没有可下载的结果。";
    return;
  }

  const blob = new Blob([lastOutput], { type: "text/plain;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = `报表统计结果_${new Date().toISOString().slice(0, 10)}.txt`;
  anchor.click();
  URL.revokeObjectURL(url);
  statusEl.textContent = "结果已下载为文本文件。";
}

function downloadCirculation() {
  if (!selectedFile) {
    statusEl.textContent = "请先选择 Excel 文件。";
    window.alert("请先选择 Excel 文件。");
    return;
  }

  const formData = new FormData();
  formData.append("file", selectedFile);
  formData.append("mode", getMode());
  formData.append("anchor_date", anchorDateInput.value);
  formData.append("start_date", startDateInput.value);
  formData.append("end_date", endDateInput.value);

  statusEl.textContent = "正在导出流转明细...";
  fetch("/api/export-circulation", {
    method: "POST",
    body: formData,
  })
    .then(async (response) => {
      if (!response.ok) {
        let errorMessage = "导出失败";
        try {
          const payload = await response.json();
          errorMessage = payload.error || errorMessage;
        } catch {
          errorMessage = "导出失败";
        }
        throw new Error(errorMessage);
      }

      const blob = await response.blob();
      const disposition = response.headers.get("Content-Disposition") || "";
      const utf8Match = disposition.match(/filename\*=UTF-8''([^;]+)/);
      const asciiMatch = disposition.match(/filename="([^"]+)"/);
      const filename = utf8Match
        ? decodeURIComponent(utf8Match[1])
        : asciiMatch
          ? asciiMatch[1]
          : `报表流转明细_${new Date().toISOString().slice(0, 10)}.xlsx`;
      const url = URL.createObjectURL(blob);
      const anchor = document.createElement("a");
      anchor.href = url;
      anchor.download = filename;
      anchor.style.display = "none";
      document.body.appendChild(anchor);
      anchor.click();
      anchor.remove();
      URL.revokeObjectURL(url);
      statusEl.textContent = "流转明细已导出为 Excel。";
    })
    .catch((error) => {
      statusEl.textContent = error.message || "导出失败";
      window.alert(error.message || "导出失败");
    });
}

function setBusy(busy, message) {
  generateBtn.disabled = busy;
  generateBtn.textContent = busy ? "生成中..." : "生成统计结果";
  statusEl.textContent = message;
}

function summaryCard(label, value, sub) {
  return `
    <article class="summary-card">
      <span class="label">${escapeHtml(label)}</span>
      <strong>${escapeHtml(value)}</strong>
      <div class="sub">${escapeHtml(sub)}</div>
    </article>
  `;
}

function formatDate(date) {
  return date.toISOString().slice(0, 10);
}

function escapeHtml(text) {
  return String(text)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function csvCell(value) {
  const text = String(value ?? "");
  return `"${text.replaceAll('"', '""')}"`;
}
