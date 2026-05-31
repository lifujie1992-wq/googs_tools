const state = {
  excelFile: null,
  preview: null,
  downloadUrl: null,
};

const $ = (id) => document.getElementById(id);

const excelInput = $("excelInput");
const excelDrop = $("excelDrop");
const excelPickButton = $("excelPickButton");
const generateButton = $("generateButton");
const downloadButton = $("downloadButton");
const previewRows = $("previewRows");
const previewDialog = $("previewDialog");

function setStatus(text, kind = "idle") {
  const status = $("parseStatus");
  status.textContent = text;
  status.className = `status ${kind}`;
}

function setStep(step) {
  $("stepPreview").classList.toggle("active", step >= 2);
  $("stepDownload").classList.toggle("active", step >= 3);
}

function formatBytes(bytes) {
  if (!bytes) return "0 B";
  const units = ["B", "KB", "MB", "GB"];
  let value = bytes;
  let index = 0;
  while (value >= 1024 && index < units.length - 1) {
    value /= 1024;
    index += 1;
  }
  return `${value.toFixed(index === 0 ? 0 : 1)} ${units[index]}`;
}

function updateStats(summary = {}) {
  $("productCount").textContent = summary.product_count || 0;
  $("skuRows").textContent = summary.sku_rows || 0;
  $("errorRows").textContent = summary.error_rows || 0;
  $("folderCount").textContent = summary.folder_count || 0;
}

function renderPreview(products = []) {
  if (!products.length) {
    previewRows.innerHTML = '<tr><td colspan="4" class="empty">没有识别到商品数据</td></tr>';
    return;
  }

  previewRows.innerHTML = products
    .map((item) => {
      return `
        <tr>
          <td>${item.index}</td>
          <td>${escapeHtml(item.name)}</td>
          <td>${item.rows} 行</td>
          <td><button class="preview-link" type="button" data-index="${item.index - 1}">查看</button></td>
        </tr>
      `;
    })
    .join("");
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

async function parseExcel(file) {
  if (!file) return;
  state.excelFile = file;
  state.preview = null;
  clearDownload();

  $("excelName").textContent = file.name;
  $("currentFileName").textContent = file.name;
  $("currentFileMeta").textContent = `大小：${formatBytes(file.size)}，正在解析`;
  setStatus("解析中", "idle");
  generateButton.disabled = true;
  previewRows.innerHTML = '<tr><td colspan="4" class="empty">正在解析 Excel...</td></tr>';

  const formData = new FormData();
  formData.append("excel", file);

  try {
    const response = await fetch("/api/preview", {
      method: "POST",
      body: formData,
    });
    const data = await response.json();
    if (!response.ok) {
      throw new Error(data.detail || "解析失败");
    }
    state.preview = data;
    updateStats(data.summary);
    renderPreview(data.products);
    $("currentFileMeta").textContent = `大小：${formatBytes(file.size)}，识别到 ${data.summary.product_count} 个商品`;
    setStatus("解析成功", "ok");
    setStep(2);
    generateButton.disabled = data.summary.product_count < 1;
  } catch (error) {
    updateStats();
    renderPreview([]);
    $("currentFileMeta").textContent = error.message;
    setStatus("解析失败", "error");
    generateButton.disabled = true;
  }
}

function clearDownload() {
  if (state.downloadUrl) URL.revokeObjectURL(state.downloadUrl);
  state.downloadUrl = null;
  downloadButton.disabled = true;
  $("downloadTitle").textContent = "等待生成数据包";
  $("downloadMeta").textContent = "解析成功后可生成 ZIP 文件";
  setStep(state.preview ? 2 : 1);
}

async function generateZip() {
  if (!state.excelFile) return;

  generateButton.disabled = true;
  generateButton.textContent = "正在生成...";
  $("downloadTitle").textContent = "正在生成数据包";
  $("downloadMeta").textContent = "请稍等，系统正在拆分 Excel 并打包";

  const formData = new FormData();
  formData.append("excel", state.excelFile);
  formData.append("options", JSON.stringify({ preserveHeaders: true }));

  try {
    const response = await fetch("/api/generate", {
      method: "POST",
      body: formData,
    });
    if (!response.ok) {
      let message = "生成失败";
      try {
        const data = await response.json();
        message = data.detail || message;
      } catch {
        message = await response.text();
      }
      throw new Error(message);
    }

    const blob = await response.blob();
    clearDownload();
    state.downloadUrl = URL.createObjectURL(blob);
    downloadButton.disabled = false;
    $("downloadTitle").textContent = "数据包生成完成";
    $("downloadMeta").textContent = `总大小：${formatBytes(blob.size)}，可直接下载`;
    setStep(3);
  } catch (error) {
    $("downloadTitle").textContent = "生成失败";
    $("downloadMeta").textContent = error.message;
  } finally {
    generateButton.disabled = false;
    generateButton.textContent = "确认无误，开始生成";
  }
}

function downloadZip() {
  if (!state.downloadUrl) return;
  const link = document.createElement("a");
  link.href = state.downloadUrl;
  link.download = "商品数据包.zip";
  document.body.appendChild(link);
  link.click();
  link.remove();
}

function resetAll() {
  state.excelFile = null;
  state.preview = null;
  clearDownload();
  excelInput.value = "";
  $("excelName").textContent = "点击或拖拽 Excel 文件到此处上传";
  $("currentFileName").textContent = "尚未上传 Excel";
  $("currentFileMeta").textContent = "上传后会在这里显示解析状态";
  setStatus("等待上传", "idle");
  updateStats();
  previewRows.innerHTML = '<tr><td colspan="4" class="empty">上传 Excel 后显示预览结果</td></tr>';
  generateButton.disabled = true;
  setStep(1);
}

function showProductPreview(index) {
  const product = state.preview?.products?.[index];
  if (!product) return;
  $("dialogTitle").textContent = product.name;
  $("dialogBody").textContent = JSON.stringify(product.preview, null, 2);
  previewDialog.showModal();
}

function downloadSampleXlsx() {
  const link = document.createElement("a");
  link.href = "/samples/商品数据示例.xlsx";
  link.download = "商品数据示例.xlsx";
  link.click();
}

function openExcelPicker() {
  excelInput.click();
}

excelInput.addEventListener("change", () => parseExcel(excelInput.files[0]));
excelDrop.addEventListener("click", openExcelPicker);
excelDrop.addEventListener("keydown", (event) => {
  if (event.key !== "Enter" && event.key !== " ") return;
  event.preventDefault();
  openExcelPicker();
});
excelPickButton.addEventListener("click", (event) => {
  event.stopPropagation();
  openExcelPicker();
});
["dragenter", "dragover"].forEach((eventName) => {
  excelDrop.addEventListener(eventName, (event) => {
    event.preventDefault();
    excelDrop.classList.add("dragging");
  });
});

["dragleave", "drop"].forEach((eventName) => {
  excelDrop.addEventListener(eventName, (event) => {
    event.preventDefault();
    excelDrop.classList.remove("dragging");
  });
});

excelDrop.addEventListener("drop", (event) => {
  const file = [...event.dataTransfer.files].find((item) => item.name.toLowerCase().endsWith(".xlsx"));
  if (file) parseExcel(file);
});

previewRows.addEventListener("click", (event) => {
  const button = event.target.closest(".preview-link");
  if (!button) return;
  showProductPreview(Number(button.dataset.index));
});

generateButton.addEventListener("click", generateZip);
downloadButton.addEventListener("click", downloadZip);
$("resetButton").addEventListener("click", resetAll);
$("closeDialog").addEventListener("click", () => previewDialog.close());
$("sampleButton").addEventListener("click", downloadSampleXlsx);
$("helpButton").addEventListener("click", () => {
  alert("先上传前两行为表头、第二行包含“商品名称”列的 .xlsx 文件。系统会按商品名称分组预览，确认后生成每个商品独立文件夹、商品数据-资料编码.xlsx 和四个空图片文件夹。");
});
