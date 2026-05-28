const state = {
  excelFile: null,
  zipFile: null,
  preview: null,
  downloadUrl: null,
};

const $ = (id) => document.getElementById(id);

const excelInput = $("excelInput");
const zipInput = $("zipInput");
const excelDrop = $("excelDrop");
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
  if (state.zipFile) formData.append("images_zip", state.zipFile);
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
  const rows = [
    ["商品名称", "商品编码", "SKU编码", "规格1", "规格值1", "价格", "库存", "主图", "详情图", "颜色图"],
    ["移动WiFi A款", "A001", "A001-BLACK", "颜色", "黑色", "99", "120", "a-main.jpg", "a-detail-1.jpg", "a-black.jpg"],
    ["移动WiFi A款", "A001", "A001-WHITE", "颜色", "白色", "99", "80", "a-main.jpg", "a-detail-1.jpg", "a-white.jpg"],
    ["移动WiFi B款", "B001", "B001-64G", "容量", "64G", "129", "60", "b-main.jpg", "b-detail-1.jpg", "b-64g.jpg"],
  ];
  const xmlRows = rows
    .map((row, rowIndex) => {
      const cells = row
        .map((value, colIndex) => {
          const ref = `${columnName(colIndex)}${rowIndex + 1}`;
          return `<c r="${ref}" t="inlineStr"><is><t>${escapeXml(value)}</t></is></c>`;
        })
        .join("");
      return `<row r="${rowIndex + 1}">${cells}</row>`;
    })
    .join("");

  const sheet = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData>${xmlRows}</sheetData></worksheet>`;
  const files = {
    "[Content_Types].xml": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>',
    "_rels/.rels": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
    "xl/workbook.xml": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="商品数据" sheetId="1" r:id="rId1"/></sheets></workbook>',
    "xl/_rels/workbook.xml.rels": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>',
    "xl/worksheets/sheet1.xml": sheet,
  };
  buildZip(files).then((blob) => {
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = "商品数据示例.xlsx";
    link.click();
    URL.revokeObjectURL(url);
  });
}

function escapeXml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

function columnName(index) {
  let name = "";
  index += 1;
  while (index > 0) {
    const remainder = (index - 1) % 26;
    name = String.fromCharCode(65 + remainder) + name;
    index = Math.floor((index - 1) / 26);
  }
  return name;
}

async function buildZip(files) {
  const encoder = new TextEncoder();
  const chunks = [];
  const central = [];
  let offset = 0;

  for (const [name, content] of Object.entries(files)) {
    const nameBytes = encoder.encode(name);
    const data = encoder.encode(content);
    const crc = crc32(data);
    const local = new Uint8Array(30 + nameBytes.length);
    const view = new DataView(local.buffer);
    view.setUint32(0, 0x04034b50, true);
    view.setUint16(4, 20, true);
    view.setUint16(8, 0, true);
    view.setUint32(14, crc, true);
    view.setUint32(18, data.length, true);
    view.setUint32(22, data.length, true);
    view.setUint16(26, nameBytes.length, true);
    local.set(nameBytes, 30);
    chunks.push(local, data);

    const header = new Uint8Array(46 + nameBytes.length);
    const centralView = new DataView(header.buffer);
    centralView.setUint32(0, 0x02014b50, true);
    centralView.setUint16(4, 20, true);
    centralView.setUint16(6, 20, true);
    centralView.setUint32(16, crc, true);
    centralView.setUint32(20, data.length, true);
    centralView.setUint32(24, data.length, true);
    centralView.setUint16(28, nameBytes.length, true);
    centralView.setUint32(42, offset, true);
    header.set(nameBytes, 46);
    central.push(header);
    offset += local.length + data.length;
  }

  const centralSize = central.reduce((sum, item) => sum + item.length, 0);
  const end = new Uint8Array(22);
  const endView = new DataView(end.buffer);
  endView.setUint32(0, 0x06054b50, true);
  endView.setUint16(8, central.length, true);
  endView.setUint16(10, central.length, true);
  endView.setUint32(12, centralSize, true);
  endView.setUint32(16, offset, true);
  return new Blob([...chunks, ...central, end], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
}

const crcTable = (() => {
  const table = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) {
      c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    }
    table[n] = c >>> 0;
  }
  return table;
})();

function crc32(data) {
  let crc = 0xffffffff;
  for (const byte of data) {
    crc = crcTable[(crc ^ byte) & 0xff] ^ (crc >>> 8);
  }
  return (crc ^ 0xffffffff) >>> 0;
}

excelInput.addEventListener("change", () => parseExcel(excelInput.files[0]));
zipInput.addEventListener("change", () => {
  state.zipFile = zipInput.files[0] || null;
  $("zipName").textContent = state.zipFile ? state.zipFile.name : "把主图、详情图、颜色图等图片压缩成 ZIP 后上传";
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
  alert("先上传包含“商品名称”列的 .xlsx 文件。系统会按商品名称分组预览，确认后生成每个商品独立文件夹和商品数据.xlsx。图片 ZIP 为可选，图片文件名需和 Excel 中填写的一致。");
});
