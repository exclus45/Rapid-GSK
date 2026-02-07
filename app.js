const appEl = document.getElementById("app");

appEl.innerHTML = `
  <div class="tabs">
    <button class="tab-btn active" data-tab="import">Импорт</button>
    <button class="tab-btn" data-tab="optimize">Оптимизация</button>
    <button class="tab-btn" data-tab="export">Экспорт</button>
  </div>

  <div class="tab-panel active" id="tab-import">
    <div class="card">
      <div class="row center-row">
        <input id="csvInput" type="file" accept=".csv,.txt" multiple />
        <label class="import-btn" for="csvInput">Импорт</label>
        <button id="clearBtn" type="button" class="btn-secondary">Очистить</button>
      </div>
      <div id="dropZone">Перетащите файлы сюда или выберите через кнопку выше</div>
      <div id="fileList"></div>
    </div>

    <div class="card">
      <h2 class="center-heading">Всего элементов: <span id="totalCount">0</span></h2>
      <div id="tableWrap"></div>
    </div>
  </div>

  <div class="tab-panel" id="tab-optimize">
    <div class="card">
      <div class="opt-controls">
        <button id="optimizeBtn" class="import-btn" type="button">Рассчитать</button>
        <button id="printBtn" class="btn-secondary" type="button">Печать</button>
        <button id="saveBtn" class="btn-secondary" type="button">Сохранить</button>
      </div>
      <div id="optBlocks"></div>
    </div>
    <div class="card print-area">
      <div class="opt-result-header">
        <h2>Результат раскроя</h2>
        <div class="opt-date">
          <label for="cutDate">Дата:</label>
          <input id="cutDate" type="date" />
        </div>
      </div>
      <div id="optResults"><small>Нет расчетов</small></div>
    </div>
  </div>

  <div class="tab-panel" id="tab-export">
    <div class="card">
      <h2 id="exportTitle">Всего в автоматическом режиме: 0</h2>
      <div id="exportBlocks"><small>Нет данных</small></div>
    </div>
  </div>
`;

const tableWrap = document.getElementById("tableWrap");
const input = document.getElementById("csvInput");
const clearBtn = document.getElementById("clearBtn");
const dropZone = document.getElementById("dropZone");
const fileList = document.getElementById("fileList");
const totalCountEl = document.getElementById("totalCount");
const optBlocks = document.getElementById("optBlocks");
const optResults = document.getElementById("optResults");
const optimizeBtn = document.getElementById("optimizeBtn");
const printBtn = document.getElementById("printBtn");
const saveBtn = document.getElementById("saveBtn");
const cutDateInput = document.getElementById("cutDate");
const exportBlocks = document.getElementById("exportBlocks");
const exportTitle = document.getElementById("exportTitle");

const DEFAULT_STOCK = 6500;
const DEFAULT_KERF = 15;
const DEFAULT_MODE = "double";
const EXPORT_ORDER_NO = 1;
const EXPORT_CUT_NO = 1;
const EXPORT_COLOR = 0;
const EXPORT_PAIR_MODE = 0;

let allParts = [];
let optimizeCache = { groups: [], skewGroups: [], baseSkewCount: 0 };
let deletedParts = [];
const filesMap = new Map();
const contentHashes = new Set();
let exportCache = [];
let exportReady = false;

function sortParts(parts) {
  return [...parts].sort((a, b) => {
    const keyA = `${a.profileCode} ${a.name}`.toLowerCase();
    const keyB = `${b.profileCode} ${b.name}`.toLowerCase();
    const cmp = keyA.localeCompare(keyB, "ru", { sensitivity: "base" });
    if (cmp !== 0) return cmp;
    const ordA = `${a.orderId}`;
    const ordB = `${b.orderId}`;
    const oc = ordA.localeCompare(ordB, "ru", { sensitivity: "base" });
    if (oc !== 0) return oc;
    return `${a.productId}`.localeCompare(`${b.productId}`, "ru", { sensitivity: "base" });
  });
}

function normalizeGroupText(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/\s+/g, " ")
    .trim();
}

function normalizeGroupKey(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/\s+/g, "")
    .replace(/[^a-z0-9а-яё]/g, "");
}

function canonicalGroupName(part) {
  const nameRaw = String(part.name || "").trim();
  const nameKey = normalizeGroupKey(nameRaw);
  const codeKey = normalizeGroupKey(part.profileCode || "");
  const isXs35801 = codeKey.startsWith("xs35801");
  const isProwinDoor =
    nameKey.includes("рамаокондверн63ммprowin") ||
    nameKey.includes("рамаокондвер63ммprowin");
  if (isXs35801 && isProwinDoor) {
    return "Рама оконная 63 мм (ProWin)";
  }
  return nameRaw;
}

function groupByNomenclature(parts) {
  const groups = new Map();
  for (const p of parts) {
    const name = canonicalGroupName(p);
    const key = normalizeGroupText(name);
    if (!groups.has(key)) {
      groups.set(key, { key, name, items: [] });
    }
    groups.get(key).items.push(p);
  }
  return Array.from(groups.values()).sort((a, b) => {
    return String(a.name).localeCompare(String(b.name), "ru", { sensitivity: "base" });
  });
}

function normalizeAngles(value) {
  return String(value || "")
    .replace(/\s+/g, "")
    .replace(",", ".")
    .toLowerCase();
}

function sanitizeDigits(value) {
  const digits = String(value || "").replace(/\D/g, "");
  return digits || "0";
}

function extractProfileBase(value) {
  const match = String(value || "").match(/[A-Za-z0-9.]+/);
  return match ? match[0] : String(value || "").trim();
}

function parseAnglePair(angles) {
  const matches = String(angles || "").match(/(\d+(?:[.,]\d+)?)/g);
  const leftRaw = matches && matches[0] ? matches[0] : "90";
  const rightRaw = matches && matches[1] ? matches[1] : leftRaw;
  const left = Number(String(leftRaw).replace(",", "."));
  const right = Number(String(rightRaw).replace(",", "."));
  return {
    left: Number.isFinite(left) ? left : 90,
    right: Number.isFinite(right) ? right : 90
  };
}

function encodeWin1251(str) {
  const bytes = new Uint8Array(str.length);
  for (let i = 0; i < str.length; i++) {
    const code = str.charCodeAt(i);

    if (code <= 0x7f) {
      bytes[i] = code;
      continue;
    }
    if (code >= 0x0410 && code <= 0x042f) {
      bytes[i] = code - 0x0410 + 0xc0;
      continue;
    }
    if (code >= 0x0430 && code <= 0x044f) {
      bytes[i] = code - 0x0430 + 0xe0;
      continue;
    }

    switch (code) {
      case 0x0401: bytes[i] = 0xa8; break;
      case 0x0451: bytes[i] = 0xb8; break;
      case 0x2116: bytes[i] = 0xb9; break;
      case 0x00b0: bytes[i] = 0xb0; break;
      case 0x00ab: bytes[i] = 0xab; break;
      case 0x00bb: bytes[i] = 0xbb; break;
      case 0x2013: bytes[i] = 0x96; break;
      case 0x2014: bytes[i] = 0x97; break;
      case 0x2018: bytes[i] = 0x91; break;
      case 0x2019: bytes[i] = 0x92; break;
      case 0x201c: bytes[i] = 0x93; break;
      case 0x201d: bytes[i] = 0x94; break;
      case 0x2022: bytes[i] = 0x95; break;
      case 0x2026: bytes[i] = 0x85; break;
      default: bytes[i] = 0x3f;
    }
  }
  return bytes;
}

function padLeft(value, length) {
  return String(value).padStart(length, "0");
}

function angleToSCode(angleLeft, angleRight) {
  const map = { 45: "1", 90: "2", 135: "3" };
  const left = map[Math.round(angleLeft)] || "2";
  const right = map[Math.round(angleRight)] || "2";
  return left + right;
}

function angleToW(angle) {
  const val = Math.round(Number(angle) * 10);
  return padLeft(val, 4);
}

function isSkewPart(part) {
  if (!part.angles) return false;
  const norm = normalizeAngles(part.angles);
  return !["45-45", "45-90", "90-45"].includes(norm);
}

function renderGroupTable(items) {
  const sorted = [...items].sort((a, b) => {
    const aLen = Number(a.lengthMm) || 0;
    const bLen = Number(b.lengthMm) || 0;
    return bLen - aLen;
  });
  const rows = sorted.map((p) => `
    <tr>
      <td>${p.profileCode}</td>
      <td>${p.name}</td>
      <td>
        <span class="size-cell">${p.lengthMm}</span>
        <button class="row-remove" type="button" data-id="${p.id}" title="Удалить">×</button>
      </td>
      <td>${p.angles || ""}</td>
      <td>${(p.orient || "").replace("Вертикальн.", "Вертикаль").replace("Горизонтальн.", "Горизонт")}</td>
      <td>${p.quantity}</td>
      <td>${p.orderId}${p.productId ? " / " + p.productId : ""}</td>
    </tr>
  `).join("");

  return `
    <table>
      <thead>
        <tr>
          <th>Артикул</th>
          <th>Название</th>
          <th>Размер</th>
          <th>Углы</th>
          <th>Ориентир</th>
          <th>Количество</th>
          <th>Заказ / изделие</th>
        </tr>
      </thead>
      <tbody>${rows}</tbody>
    </table>
  `;
}

function renderTable(parts) {
  if (!parts.length) {
    tableWrap.innerHTML = "<small>Нет данных</small>";
    totalCountEl.textContent = "0";
    renderOptimize(parts);
    return;
  }

  const sorted = sortParts(parts);
  const normalParts = sorted.filter((p) => !isSkewPart(p));
  const skewParts = sorted.filter((p) => isSkewPart(p));
  const groups = groupByNomenclature(normalParts);
  const skewGroups = groupByNomenclature(skewParts);
  const totalCount = parts.reduce((s, it) => s + (Number(it.quantity) || 0), 0);
  totalCountEl.textContent = String(totalCount);

  const html = groups.map((g) => {
    const count = g.items.reduce((s, it) => s + (Number(it.quantity) || 0), 0);
    return `
      <div class="group" data-group="${g.key}">
        <div class="group-header">
          <button class="group-toggle" type="button" aria-expanded="true">−</button>
          <div class="group-title">${g.name}</div>
          <div class="group-meta">${count} шт.</div>
        </div>
        <div class="group-body">
          ${renderGroupTable(g.items)}
        </div>
      </div>
    `;
  }).join("");

  const skewHtml = skewGroups.length
    ? `
      <div class="group skew-group">
        <div class="group-header">
          <button class="group-toggle" type="button" aria-expanded="true">−</button>
          <div class="group-title">Косые элементы</div>
          <div class="group-meta">${skewParts.reduce((s, it) => s + (Number(it.quantity) || 0), 0)} шт.</div>
        </div>
        <div class="group-body">
          ${skewGroups.map((g) => renderGroupTable(g.items)).join("")}
        </div>
      </div>
    `
    : "";

  tableWrap.innerHTML = html + skewHtml;

  tableWrap.querySelectorAll(".group-toggle").forEach((btn) => {
    btn.addEventListener("click", () => {
      const group = btn.closest(".group");
      const body = group.querySelector(".group-body");
      const expanded = btn.getAttribute("aria-expanded") === "true";
      btn.setAttribute("aria-expanded", String(!expanded));
      btn.textContent = expanded ? "+" : "−";
      body.style.display = expanded ? "none" : "block";
    });
  });

  tableWrap.querySelectorAll(".row-remove").forEach((btn) => {
    btn.addEventListener("click", () => {
      const id = btn.getAttribute("data-id");
      if (!id) return;
      removePartById(id);
    });
  });

  renderOptimize(parts);
}

function removePartById(id) {
  const removed = allParts.filter((p) => p.id === id);
  if (removed.length) {
    deletedParts = deletedParts.concat(removed);
  }
  allParts = allParts.filter((p) => p.id !== id);
  filesMap.forEach((fileInfo) => {
    fileInfo.parts = fileInfo.parts.filter((p) => p.id !== id);
  });
  exportReady = false;
  renderTable(allParts);
}

function renderFileList() {
  if (!filesMap.size) {
    fileList.innerHTML = "";
    return;
  }
  const rows = Array.from(filesMap.values())
    .sort((a, b) => a.name.localeCompare(b.name, "ru", { sensitivity: "base" }))
    .map((f) => `
    <div class="file-item" data-key="${f.key}">
      <span class="file-name">${f.name}</span>
      <button class="file-remove" type="button" title="Удалить">×</button>
    </div>
  `).join("");
  fileList.innerHTML = rows;

  fileList.querySelectorAll(".file-remove").forEach((btn) => {
    btn.addEventListener("click", () => {
      const parent = btn.closest(".file-item");
      const key = parent?.getAttribute("data-key");
      if (!key) return;
      const fileInfo = filesMap.get(key);
      if (fileInfo?.hash) contentHashes.delete(fileInfo.hash);
      filesMap.delete(key);
      exportReady = false;
      rebuildAllParts();
      renderFileList();
      renderTable(allParts);
    });
  });
}

function rebuildAllParts() {
  allParts = [];
  for (const f of filesMap.values()) {
    allParts = allParts.concat(f.parts);
  }
}

function decodeSmart(buffer) {
  const utf8 = new TextDecoder("utf-8", { fatal: false }).decode(buffer);
  const replCount = (utf8.match(/\uFFFD/g) || []).length;
  if (replCount > 0) {
    return new TextDecoder("windows-1251", { fatal: false }).decode(buffer);
  }
  return utf8;
}

async function readFileToParts(file) {
  const buffer = await file.arrayBuffer();
  const text = decodeSmart(buffer);
  return { parts: parsePartsFromCsv(text), buffer };
}

async function hashBuffer(buffer) {
  const hash = await crypto.subtle.digest("SHA-256", buffer);
  const bytes = new Uint8Array(hash);
  return Array.from(bytes).map((b) => b.toString(16).padStart(2, "0")).join("");
}

function fileKey(file) {
  return `${file.name}|${file.size}|${file.lastModified}`;
}

async function addFiles(files) {
  for (const file of files) {
    const key = fileKey(file);
    if (filesMap.has(key)) continue;

    const { parts, buffer } = await readFileToParts(file);
    const contentHash = await hashBuffer(buffer);
    if (contentHashes.has(contentHash)) continue;

    contentHashes.add(contentHash);
    filesMap.set(key, { key, name: file.name, parts, hash: contentHash });
  }
  rebuildAllParts();
  exportReady = false;
  renderFileList();
  renderTable(allParts);
}

input.addEventListener("change", async () => {
  const files = Array.from(input.files || []);
  if (!files.length) return;
  await addFiles(files);
  input.value = "";
});

dropZone.addEventListener("dragover", (e) => {
  e.preventDefault();
  dropZone.classList.add("is-dragover");
});
dropZone.addEventListener("dragleave", () => {
  dropZone.classList.remove("is-dragover");
});
dropZone.addEventListener("drop", async (e) => {
  e.preventDefault();
  dropZone.classList.remove("is-dragover");
  const files = Array.from(e.dataTransfer.files || []).filter((f) =>
    f.name.toLowerCase().endsWith(".csv") || f.name.toLowerCase().endsWith(".txt")
  );
  if (!files.length) return;
  await addFiles(files);
});

clearBtn.addEventListener("click", () => {
  allParts = [];
  deletedParts = [];
  filesMap.clear();
  contentHashes.clear();
  exportReady = false;
  renderFileList();
  renderTable(allParts);
});

const tabButtons = document.querySelectorAll(".tab-btn");
const panels = {
  import: document.getElementById("tab-import"),
  optimize: document.getElementById("tab-optimize"),
  export: document.getElementById("tab-export")
};

tabButtons.forEach((btn) => {
  btn.addEventListener("click", () => {
    tabButtons.forEach((b) => b.classList.remove("active"));
    btn.classList.add("active");
    const tab = btn.getAttribute("data-tab");
    Object.values(panels).forEach((p) => p.classList.remove("active"));
    panels[tab].classList.add("active");
  });
});

function expandPartsToPieces(parts) {
  const pieces = [];
  for (const p of parts) {
    const len = Number(p.lengthMm);
    const qty = Number(p.quantity) || 0;
    if (!Number.isFinite(len) || len <= 0) continue;
    if (!Number.isFinite(qty) || qty <= 0) continue;
    for (let i = 0; i < qty; i++) {
      pieces.push({
        lengthMm: len,
        angles: String(p.angles || "").trim(),
        orderId: String(p.orderId || "").trim(),
        productId: String(p.productId || "").trim(),
        title: String(p.name || "").trim(),
        orient: String(p.orient || "").trim()
      });
    }
  }
  return pieces;
}

function buildOptimizeData(parts) {
  const sorted = sortParts(parts);
  const normalParts = sorted.filter((p) => !isSkewPart(p));
  const skewParts = sorted.filter((p) => isSkewPart(p));
  const groups = groupByNomenclature(normalParts).map((g) => {
    const pieces = expandPartsToPieces(g.items);
    return {
      title: g.name,
      profileCode: g.items[0]?.profileCode || "",
      pieces,
      total: pieces.length,
      mode: DEFAULT_MODE
    };
  });
  const skewGroups = groupByNomenclature(skewParts).map((g) => {
    const pieces = expandPartsToPieces(g.items);
    return {
      title: g.name,
      pieces,
      total: pieces.length
    };
  });
  const baseSkewCount = skewParts.reduce((s, it) => s + (Number(it.quantity) || 0), 0);
  return { groups, skewGroups, baseSkewCount };
}

function buildLeftoverSkewItems(groups) {
  const leftovers = [];
  groups.forEach((g) => {
    if (g.mode !== "double") return;
    const { leftovers: groupLeftovers } = buildDoublePairs(g.pieces);
    groupLeftovers.forEach((piece) => {
      leftovers.push({ ...piece, title: g.title });
    });
  });
  return leftovers;
}

function renderOptimize(parts) {
  if (!optBlocks) return;
  if (!parts.length) {
    optBlocks.innerHTML = "<small>Нет данных</small>";
    optResults.innerHTML = "<small>Нет расчетов</small>";
    optimizeCache = { groups: [], skewGroups: [], baseSkewCount: 0 };
    exportReady = false;
    renderExport();
    return;
  }

  optimizeCache = buildOptimizeData(parts);

  if (!optimizeCache.groups.length && !optimizeCache.baseSkewCount) {
    optBlocks.innerHTML = "<small>Нет данных</small>";
    optResults.innerHTML = "<small>Нет расчетов</small>";
    exportReady = false;
    renderExport();
    return;
  }

  const lines = optimizeCache.groups.map((g, index) => {
    const selectedDouble = g.mode === "double" ? "selected" : "";
    const selectedSingle = g.mode === "single" ? "selected" : "";
    const leftovers = g.mode === "double" ? buildDoublePairs(g.pieces).leftovers.length : 0;
    const shown = Math.max(0, g.total - leftovers);
    return `
      <div class="opt-line" data-index="${index}">
        <div class="opt-line-title">${g.title}</div>
        <div class="opt-line-count">${shown} шт.</div>
        <div class="opt-line-controls">
          <select class="opt-mode">
            <option value="double" ${selectedDouble}>Двойной</option>
            <option value="single" ${selectedSingle}>Одинарный</option>
          </select>
        </div>
      </div>
    `;
  }).join("");

  const buildSkewSummary = () => {
    const leftoverItems = buildLeftoverSkewItems(optimizeCache.groups);
    const leftoverCount = leftoverItems.length;
    const total = optimizeCache.baseSkewCount + leftoverCount;
    const skewByTitle = new Map();
    optimizeCache.skewGroups.forEach((g) => {
      const count = g.pieces.length;
      if (!skewByTitle.has(g.title)) skewByTitle.set(g.title, 0);
      skewByTitle.set(g.title, skewByTitle.get(g.title) + count);
    });
    leftoverItems.forEach((it) => {
      const title = it.title || "Косые элементы";
      if (!skewByTitle.has(title)) skewByTitle.set(title, 0);
      skewByTitle.set(title, skewByTitle.get(title) + 1);
    });
    const listHtml = Array.from(skewByTitle.entries())
      .map(([title, count]) => `<div class="opt-skew-item">${title} — ${count} шт.</div>`)
      .join("");
    return { total, listHtml };
  };

  const { total: totalSkew, listHtml: skewList } = buildSkewSummary();
  let skewLine = "";
  if (totalSkew) {
    skewLine = `
      <div class="opt-line opt-line-skew">
        <div class="opt-line-title">Ручной режим</div>
        <div class="opt-line-count"><span id="skewCount">${totalSkew}</span> шт.</div>
      </div>
      <div class="opt-skew-list" id="skewList">
        ${skewList}
      </div>
    `;
  }
  const deletedCount = deletedParts.reduce((s, p) => s + (Number(p.quantity) || 1), 0);
  const deletedLine = deletedCount
    ? `
      <div class="opt-line opt-line-deleted">
        <div class="opt-line-title">Запасы</div>
        <div class="opt-line-count">${deletedCount} шт.</div>
      </div>
    `
    : "";

  optBlocks.innerHTML = `
    <div class="opt-settings">
      <label>Длина хлыста, мм
        <input class="opt-stock" type="number" value="${DEFAULT_STOCK}" disabled />
      </label>
      <label>Пропил, мм
        <input class="opt-kerf" type="number" value="${DEFAULT_KERF}" disabled />
      </label>
    </div>
    <div class="opt-summary">
      ${lines}
      ${skewLine}
      ${deletedLine}
    </div>
  `;

  optBlocks.querySelectorAll(".opt-mode").forEach((select) => {
    select.addEventListener("change", () => {
      const line = select.closest(".opt-line");
      const idx = Number(line?.getAttribute("data-index"));
      if (!Number.isFinite(idx) || !optimizeCache.groups[idx]) return;
      optimizeCache.groups[idx].mode = select.value === "single" ? "single" : "double";
      const skewCountEl = document.getElementById("skewCount");
      const skewListEl = document.getElementById("skewList");
      if (skewCountEl || skewListEl) {
        const summary = buildSkewSummary();
        if (skewCountEl) skewCountEl.textContent = String(summary.total);
        if (skewListEl) skewListEl.innerHTML = summary.listHtml;
      }
      const total = optimizeCache.groups[idx].total;
      const leftover = optimizeCache.groups[idx].mode === "double"
        ? buildDoublePairs(optimizeCache.groups[idx].pieces).leftovers.length
        : 0;
        const shown = Math.max(0, total - leftover);
        const countEl = line.querySelector(".opt-line-count");
        if (countEl) countEl.textContent = `${shown} шт.`;
        exportReady = false;
        renderExport();
      });
    });

  renderExport();
}

function appendDeletedResults() {
  if (!deletedParts.length) return;
  const total = deletedParts.reduce((s, p) => s + (Number(p.quantity) || 1), 0);
  const notice = document.createElement("div");
  notice.className = "opt-deleted-notice";
  notice.textContent = `ЗАПАСЫ — ${total} шт.`;
  optResults.appendChild(notice);

  const list = document.createElement("div");
  list.className = "opt-deleted";
  const grouped = new Map();
  deletedParts.forEach((p) => {
    const name = canonicalGroupName(p) || p.name || "";
    if (!grouped.has(name)) grouped.set(name, []);
    grouped.get(name).push({ ...p, _displayName: name });
  });
  const sortedRows = [];
  Array.from(grouped.keys())
    .sort((a, b) => String(a).localeCompare(String(b), "ru", { sensitivity: "base" }))
    .forEach((name) => {
      const items = grouped.get(name) || [];
      items.sort((a, b) => (Number(b.lengthMm) || 0) - (Number(a.lengthMm) || 0));
      items.forEach((p) => sortedRows.push(p));
    });

  const rows = sortedRows.map((p) => {
    const orderText = p.orderId && p.productId ? `${p.orderId} / ${p.productId}` : "";
    const orient = p.orient || "";
    const qty = Number(p.quantity) || 1;
    return `
      <tr>
        <td>${p._displayName || p.name}</td>
        <td>${p.lengthMm}</td>
        <td>${qty}</td>
        <td>${orient}</td>
        <td>${orderText}</td>
      </tr>
    `;
  }).join("");
  list.innerHTML = `
    <table class="opt-deleted-table">
      <colgroup>
        <col style="width:35%">
        <col style="width:15%">
        <col style="width:10%">
        <col style="width:20%">
        <col style="width:20%">
      </colgroup>
      <thead>
        <tr>
          <th>Наименование</th>
          <th>Размер</th>
          <th>Кол-во</th>
          <th>Ориентир</th>
          <th>Заказ / изделие</th>
        </tr>
      </thead>
      <tbody>
        ${rows}
      </tbody>
    </table>
  `;
  optResults.appendChild(list);
}

function formatOptimizeResult(result, stock, kerf) {
  if (!result.bins.length && !result.unfit.length) {
    return "<div class=\"opt-result-sub\">Нет данных для раскроя.</div>";
  }

  const lines = [];
  lines.push(`<div class="opt-result-sub">Длина хлыста: ${stock} мм, пропил: ${kerf} мм</div>`);
  lines.push(`<div class="opt-result-count">Хлыстов: ${result.bins.length}</div>`);

  result.bins.forEach((bin, idx) => {
    const parts = bin.parts
      .map((p) => {
        const meta = p.meta || {};
        let orientLetter = "";
        const orientRaw = String(meta.orient || "").trim().toLowerCase();
        const orientKey = orientRaw.replace(/[^a-zа-яё]/g, "");
        if (orientKey.startsWith("вер") || orientKey === "в") orientLetter = "Вер";
        else if (orientKey.startsWith("гор") || orientKey === "г") orientLetter = "Гор";
        else if (orientKey.startsWith("нак") || orientKey === "н") orientLetter = "Нак";
        const baseText = meta.orderId && meta.productId ? `${meta.orderId} / ${meta.productId}` : "";
        const metaText = baseText && orientLetter ? `${baseText} (${orientLetter})` : baseText;
        const angleNorm = normalizeAngles(meta.angles || "");
        const chipClass = angleNorm && angleNorm !== "45-45" ? "cut-chip cut-chip--skew" : "cut-chip";
        return `
          <span class="${chipClass}">
            <span class="cut-chip-length">${p.length}</span>
            <span class="cut-chip-meta">${metaText}</span>
          </span>
        `;
      })
      .join("");
    const percent = stock > 0 ? (bin.leftover / stock) * 100 : 0;
    const percentText = `(${percent.toFixed(1)}%)`;
    const percentClass =
      percent <= 3 ? "percent-good" : percent <= 6 ? "percent-warn" : "percent-bad";
    lines.push(`
      <div class="opt-bin">
        <strong>Хлыст ${idx + 1}:</strong> ${parts}
        <span class="opt-left">Остаток: ${bin.leftover} мм</span>
        <span class="opt-left opt-left-percent ${percentClass}">${percentText}</span>
      </div>
    `);
  });

  if (result.unfit.length) {
    const list = result.unfit.map((p) => `${p.length}`).join(", ");
    lines.push(`<div class="opt-error">Не помещаются: ${list}</div>`);
  }

  return lines.join("");
}

function appendOptimizeResult(title, result, stock, kerf, modeLabel = "") {
  const modeHtml = modeLabel ? ` <span class="opt-mode-label">${modeLabel}</span>` : "";
  const section = document.createElement("div");
  section.className = "opt-result";
  section.innerHTML = `
    <div class="opt-result-title">${title}${modeHtml}</div>
    ${formatOptimizeResult(result, stock, kerf)}
  `;
  optResults.appendChild(section);
}

function appendOptimizeError(title, message, modeLabel = "") {
  const modeHtml = modeLabel ? ` <span class="opt-mode-label">${modeLabel}</span>` : "";
  const section = document.createElement("div");
  section.className = "opt-result";
  section.innerHTML = `
    <div class="opt-result-title">${title}${modeHtml}</div>
    <div class="opt-error">${message}</div>
  `;
  optResults.appendChild(section);
}

function buildDoublePairs(pieces) {
  const byKey = new Map();
  pieces.forEach((p) => {
    const key = `${p.lengthMm}||${p.angles || ""}`;
    if (!byKey.has(key)) byKey.set(key, []);
    byKey.get(key).push(p);
  });
  const paired = [];
  const leftovers = [];
  byKey.forEach((list) => {
    for (let i = 0; i + 1 < list.length; i += 2) {
      const a = list[i];
      const b = list[i + 1];
      paired.push({
        lengthMm: a.lengthMm,
        angles: a.angles,
        meta: [a, b]
      });
    }
    if (list.length % 2 === 1) {
      leftovers.push(list[list.length - 1]);
    }
  });
  return { paired, leftovers };
}

function buildExportGroups() {
  if (!optimizeCache.groups.length) return [];
  const groups = [];

  optimizeCache.groups.forEach((group) => {
    const mode = group.mode;
    let selectedPieces = [];
    if (group.mode === "double") {
      const { paired } = buildDoublePairs(group.pieces);
      paired.forEach((pair) => {
        const metaArr = Array.isArray(pair.meta) ? pair.meta : [pair.meta, pair.meta];
        metaArr.forEach((piece) => {
          if (piece) selectedPieces.push(piece);
        });
      });
    } else {
      selectedPieces = group.pieces;
    }

    if (!selectedPieces.length) return;

    const items = [];
    if (optimizeCache.lastExportResults && optimizeCache.lastExportResults.has(group.title)) {
      const result = optimizeCache.lastExportResults.get(group.title);
      result.bins.forEach((bin) => {
        bin.parts.forEach((part) => {
          const metaArr = Array.isArray(part.meta) ? part.meta : [part.meta];
          const meta = metaArr[0] || {};
          const lengthMm = Number(part.length || meta.lengthMm || 0);
          if (!Number.isFinite(lengthMm) || lengthMm <= 0) return;
          items.push({
            orderId: sanitizeDigits(meta.orderId),
            productId: sanitizeDigits(meta.productId),
            lengthMm,
            angles: String(meta.angles || "").trim(),
            quantity: 1,
            widthMm: meta.widthMm || "",
            heightMm: meta.heightMm || ""
          });
        });
      });
    } else {
      selectedPieces.forEach((piece) => {
        const lengthMm = Number(piece.lengthMm || piece.length || 0);
        if (!Number.isFinite(lengthMm) || lengthMm <= 0) return;
        items.push({
          orderId: sanitizeDigits(piece.orderId),
          productId: sanitizeDigits(piece.productId),
          lengthMm,
          angles: String(piece.angles || "").trim(),
          quantity: 1,
          widthMm: piece.widthMm || "",
          heightMm: piece.heightMm || ""
        });
      });
    }
    if (!items.length) return;

    const profileBase = extractProfileBase(group.profileCode);
    const totalQty = selectedPieces.length;

    groups.push({
      title: group.title,
      profileBase,
      totalQty,
      items,
      mode
    });
  });

  return groups;
}

function makeHeader(orderNo, cutNo, profileCode, color, pairMode, totalQty) {
  const order = padLeft(orderNo, 6);
  const cut = padLeft(cutNo, 4);
  const prof = padLeft(profileCode, 10);
  const colorStr = String(color);
  const pairStr = String(pairMode);
  const qty = padLeft(totalQty, 3);
  return "P" + order + cut + prof + colorStr + pairStr + qty;
}

function dateToOrderNo(dateValue) {
  const raw = String(dateValue || "").trim();
  if (!raw) return "000000";
  const parts = raw.split("-");
  if (parts.length !== 3) return "000000";
  const yyyy = parts[0];
  const mm = parts[1];
  const dd = parts[2];
  return `${dd}${mm}${yyyy.slice(-2)}`;
}

function makeNLines(items) {
  const lines = [];
  let index = 1;

  for (const item of items) {
    const orderPart = sanitizeDigits(item.orderId);
    const positionPart = sanitizeDigits(item.productId);
    const widthPart = item.widthMm ? `W${item.widthMm}` : "";
    const heightPart = item.heightMm ? `H${item.heightMm}` : "";
    const codeRaw = "Z" + orderPart + "P" + positionPart + widthPart + heightPart;
    const codeField = (codeRaw + " ".repeat(20)).slice(0, 20);

    const qty = padLeft(item.quantity, 3);
    const length010 = padLeft(Math.round(Number(item.lengthMm) * 10), 5);

    const { left, right } = parseAnglePair(item.angles);
    const sCode = angleToSCode(left, right);
    const wLeft = angleToW(left);
    const wRight = angleToW(right);

    const nNo = padLeft(index, 3);

    const line =
      "N" + nNo +
      "C" + codeField +
      "Z" + qty +
      "L" + length010 +
      "S" + sCode +
      "F001" +
      "H0" +
      "W" + wLeft +
      "W" + wRight +
      "W0000W0000";

    lines.push(line);
    index++;
  }
  return lines;
}

function makeProgramBlock(orderNo, cutNo, profileCode, colorCode, pairMode, items) {
  if (!items || !items.length) return "";
  const totalQty = items.reduce((s, it) => s + Number(it.quantity || 0), 0);
  const header = makeHeader(orderNo, cutNo, profileCode, colorCode, pairMode, totalQty);
  const nLines = makeNLines(items);
  return [header, ...nLines].join("\r\n") + "\r\n";
}

function buildJobTextForGroup(group, cutNoOverride) {
  if (!group || !group.items || !group.items.length) return "";
  const orderNo = dateToOrderNo(cutDateInput?.value);
  const items = group.items;
  return makeProgramBlock(
    orderNo,
    cutNoOverride ?? EXPORT_CUT_NO,
    group.profileBase,
    EXPORT_COLOR,
    EXPORT_PAIR_MODE,
    items
  );
}

function safeFileName(value) {
  const base = String(value || "program")
    .replace(/[^\wа-яё\-]+/gi, "_")
    .replace(/_+/g, "_")
    .replace(/^_+|_+$/g, "")
    .slice(0, 60);
  return base || "program";
}

function renderExport() {
  if (!exportBlocks) return;
  exportCache = buildExportGroups();
  if (exportTitle) {
    const totalAuto = exportCache.reduce((sum, g) => sum + (Number(g.totalQty) || 0), 0);
    exportTitle.textContent = `Всего в автоматическом режиме: ${totalAuto} шт.`;
  }
  if (!exportCache.length) {
    exportBlocks.innerHTML = "<small>Нет данных</small>";
    return;
  }

  exportBlocks.innerHTML = exportCache
    .map((g, idx) => {
      const modeLabel = g.mode === "double" ? "Двойной" : "Одинарный";
      const sample = g.items && g.items.length ? g.items[0] : null;
      const sampleOrder = sample ? sanitizeDigits(sample.orderId) : "0";
      const sampleProduct = sample ? sanitizeDigits(sample.productId) : "0";
      const sampleCode = `CZ${sampleOrder}P${sampleProduct}`;
      const orderNo = dateToOrderNo(cutDateInput?.value);
      const cutNo = padLeft(idx + 1, 4);
      const cutCount = g.items.reduce((sum, it) => sum + (Number(it.quantity) || 0), 0);
      const programCode = `P${orderNo}${cutNo}${padLeft(g.profileBase, 10)}${EXPORT_COLOR}${EXPORT_PAIR_MODE}${padLeft(cutCount, 3)}`;
      const disabledAttr = exportReady ? "" : "disabled";
      return `
        <div class="export-block" data-index="${idx}">
          <div class="export-header">
            <div class="export-title">${g.title}</div>
            <div class="export-count">${g.totalQty} шт.</div>
            <div class="export-lines">Резов: ${cutCount}</div>
            <div class="export-mode">${modeLabel}</div>
          </div>
          <div class="export-sub">Профиль: ${padLeft(g.profileBase, 10)}</div>
          <div class="export-hint-title">Подсказка по формату</div>
          <div class="export-hint">Пример заголовка программы: <code>${programCode}</code></div>
          <div class="export-hint export-hint-line">P + заказ(ДДММГГ) + отрез(0001) + профиль(10) + цвет(0) + режим(0) + кол-во(3)</div>
          <div class="export-hint">Пример C‑поля: <strong>${sampleCode}</strong> = заказ ${sampleOrder}, изделие ${sampleProduct}</div>
          <div class="export-hint export-hint-line">Пример N‑строки: <code>N001${sampleCode} ...</code></div>
          ${exportReady ? "" : "<div class=\"export-lock\">Сначала нажмите «Рассчитать» во вкладке Оптимизация.</div>"}
          <div class="export-actions">
            <button class="import-btn export-generate" type="button" data-index="${idx}" ${disabledAttr}>Сгенерировать JOB</button>
            <button class="btn-secondary export-download" type="button" data-index="${idx}" ${disabledAttr}>Скачать JOB</button>
          </div>
          <textarea class="export-text" id="exportText-${idx}" spellcheck="false"></textarea>
        </div>
      `;
    })
    .join("");

  exportBlocks.querySelectorAll(".export-generate").forEach((btn) => {
    btn.addEventListener("click", () => {
      if (!exportReady) return;
      const idx = Number(btn.getAttribute("data-index"));
      const group = exportCache[idx];
      if (!group) return;
      const output = buildJobTextForGroup(group, idx + 1);
      const area = document.getElementById(`exportText-${idx}`);
      if (area) area.value = output;
    });
  });

  exportBlocks.querySelectorAll(".export-download").forEach((btn) => {
    btn.addEventListener("click", async () => {
      if (!exportReady) return;
      const idx = Number(btn.getAttribute("data-index"));
      const group = exportCache[idx];
      if (!group) return;
      const area = document.getElementById(`exportText-${idx}`);
      let jobText = (area && area.value) || buildJobTextForGroup(group, idx + 1);
      if (!jobText) return;

      jobText = jobText.replace(/\r?\n/g, "\r\n");
      if (!jobText.endsWith("\r\n")) jobText += "\r\n";
      const bytes = encodeWin1251(jobText);
      const blob = new Blob([bytes], { type: "application/octet-stream" });

      const datePart = dateToOrderNo(cutDateInput?.value);
      const suggestedName = `${datePart}_${safeFileName(group.title)}.job`;
      if (window.showSaveFilePicker) {
        const handle = await window.showSaveFilePicker({
          suggestedName,
          types: [{ description: "JOB", accept: { "application/octet-stream": [".job"] } }]
        });
        let name = handle.name || suggestedName;
        if (!name.toLowerCase().endsWith(".job")) {
          name = name + ".job";
        }
        const writable = await handle.createWritable();
        await writable.write(blob);
        await writable.close();
        return;
      }

      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = suggestedName;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    });
  });
}

function splitPairedBins(result) {
  const bins = [];
  result.bins.forEach((bin) => {
    const binA = { used: bin.used, leftover: bin.leftover, parts: [] };
    const binB = { used: bin.used, leftover: bin.leftover, parts: [] };
    bin.parts.forEach((part) => {
      const metaArr = Array.isArray(part.meta) ? part.meta : [part.meta, part.meta];
      binA.parts.push({ length: part.length, meta: metaArr[0] || null });
      binB.parts.push({ length: part.length, meta: metaArr[1] || null });
    });
    bins.push(binA, binB);
  });

  const unfit = [];
  result.unfit.forEach((part) => {
    const metaArr = Array.isArray(part.meta) ? part.meta : [part.meta, part.meta];
    unfit.push({ length: part.length, meta: metaArr[0] || null });
    unfit.push({ length: part.length, meta: metaArr[1] || null });
  });
  return { bins, unfit };
}

optimizeBtn.addEventListener("click", () => {
  optResults.innerHTML = "";

  if (!optimizeCache.groups.length && !optimizeCache.skewGroups.length) {
    optResults.innerHTML = "<small>Нет данных</small>";
    return;
  }

  if (typeof window.optimizeCut !== "function") {
    optResults.innerHTML = "<small>Оптимайзер не подключен</small>";
    return;
  }

  if (optimizeCache.groups.length) {
    const autoCount = optimizeCache.groups.reduce((sum, g) => {
      if (g.mode === "double") {
        const leftovers = buildDoublePairs(g.pieces).leftovers.length;
        return sum + Math.max(0, g.total - leftovers);
      }
      return sum + g.total;
    }, 0);
    const notice = document.createElement("div");
    notice.className = "opt-auto-notice";
    notice.textContent = `АВТОМАТИЧЕСКИЙ РЕЖИМ — ${autoCount} шт.`;
    optResults.appendChild(notice);
  }

  const leftoverSkewItems = buildLeftoverSkewItems(optimizeCache.groups);

  optimizeCache.lastResults = new Map();
  optimizeCache.lastExportResults = new Map();

  optimizeCache.groups.forEach((group) => {
    const modeLabel = group.mode === "double" ? "Двойной" : "Одинарный";
    if (!group.pieces.length) {
      appendOptimizeError(group.title, "Нет данных для расчета.", modeLabel);
      return;
    }

    if (group.mode === "double") {
      const { paired } = buildDoublePairs(group.pieces);
      if (!paired.length) {
        appendOptimizeError(group.title, "Нет парных элементов для расчета.", modeLabel);
        return;
      }
      const pairItems = paired.map((p) => ({ lengthMm: p.lengthMm, qty: 1, meta: p.meta }));
      const baseResult = window.optimizeCut(pairItems, DEFAULT_STOCK, DEFAULT_KERF);
      const result = splitPairedBins(baseResult);
      optimizeCache.lastResults.set(group.title, result);
      optimizeCache.lastExportResults.set(group.title, baseResult);
      appendOptimizeResult(group.title, result, DEFAULT_STOCK, DEFAULT_KERF, modeLabel);
      return;
    }

    const singleItems = group.pieces.map((p) => ({ lengthMm: p.lengthMm, qty: 1, meta: p }));
    const result = window.optimizeCut(singleItems, DEFAULT_STOCK, DEFAULT_KERF);
    optimizeCache.lastResults.set(group.title, result);
    optimizeCache.lastExportResults.set(group.title, result);
    appendOptimizeResult(group.title, result, DEFAULT_STOCK, DEFAULT_KERF, modeLabel);
  });

  const byTitleMap = new Map();
  const addPieceToTitle = (title, piece) => {
    if (!byTitleMap.has(title)) byTitleMap.set(title, []);
    byTitleMap.get(title).push(piece);
  };

  optimizeCache.skewGroups.forEach((group) => {
    group.pieces.forEach((piece) => addPieceToTitle(group.title, piece));
  });
  leftoverSkewItems.forEach((piece) => addPieceToTitle(piece.title || "Косые элементы", piece));

  if (byTitleMap.size) {
    const manualCount = optimizeCache.skewGroups.reduce((s, g) => s + g.pieces.length, 0) + leftoverSkewItems.length;
    const notice = document.createElement("div");
    notice.className = "opt-manual-notice";
    notice.textContent = `ВНИМАНИЕ - РУЧНОЙ РЕЖИМ! — ${manualCount} шт.`;
    optResults.appendChild(notice);
  }

  byTitleMap.forEach((pieces, title) => {
    if (!pieces.length) return;
    const items = pieces.map((p) => ({ lengthMm: p.lengthMm, qty: 1, meta: p }));
    const result = window.optimizeCut(items, DEFAULT_STOCK, DEFAULT_KERF);
    appendOptimizeResult(`${title}`, result, DEFAULT_STOCK, DEFAULT_KERF, "Одинарный");
  });

  appendDeletedResults();
  exportReady = true;
  renderExport();
});

renderTable([]);

if (cutDateInput) {
  const today = new Date();
  const yyyy = today.getFullYear();
  const mm = String(today.getMonth() + 1).padStart(2, "0");
  const dd = String(today.getDate()).padStart(2, "0");
  cutDateInput.value = `${yyyy}-${mm}-${dd}`;
}

if (printBtn) {
  printBtn.addEventListener("click", () => {
    window.print();
  });
}

if (saveBtn) {
  saveBtn.addEventListener("click", () => {
    const dateValue = cutDateInput?.value || "";
    const safeDate = dateValue || new Date().toISOString().slice(0, 10);
    const title = "Результат раскроя";
    const content = optResults?.innerHTML || "";

    const html = `<!doctype html>
<html lang="ru">
<head>
  <meta charset="utf-8" />
  <title>${title} — ${safeDate}</title>
  <style>
    body { font-family: Arial, sans-serif; margin: 24px; color: #0f172a; }
    h1 { margin: 0 0 6px 0; }
    .date { color: #475569; margin-bottom: 16px; }
    .opt-result { border: 1px solid #d6dee8; border-radius: 10px; padding: 10px 12px; margin: 12px 0; background: #fff; }
    .opt-result-title { font-weight: 600; margin-bottom: 6px; }
    .opt-mode-label { margin-left: 8px; font-weight: 600; font-size: 13px; color: #475569; }
    .opt-result-sub, .opt-result-count { margin-bottom: 6px; }
    .opt-bin { padding: 4px 0; border-bottom: 1px dashed #e1e8f0; }
    .opt-bin:last-child { border-bottom: none; }
    .opt-left { margin-left: 8px; color: #475569; }
    .opt-left-percent { font-weight: 700; }
    .percent-good { color: #15803d; }
    .percent-warn { color: #d97706; }
    .percent-bad { color: #b42318; }
    .cut-chip { display: inline-flex; flex-direction: column; align-items: center; justify-content: center; padding: 4px 8px; margin-right: 6px; border: 1px solid #cbd5e1; border-radius: 8px; background: #f8fafc; font-weight: 700; text-align: center; min-width: 72px; }
    .cut-chip--skew { background: #111; border-color: #111; color: #fff; }
    .cut-chip-length { display: block; font-size: 16px; }
    .cut-chip-meta { display: block; margin-top: 2px; padding-top: 2px; border-top: 1px solid #cbd5e1; font-weight: 600; font-size: 13px; color: #475569; white-space: nowrap; }
    .cut-chip--skew .cut-chip-meta { border-top-color: #333; color: #e2e8f0; }
    .opt-auto-notice { margin: 12px 0; padding: 10px 12px; border: 2px solid #1d4ed8; background: #eff6ff; color: #1d4ed8; font-weight: 700; text-align: center; border-radius: 10px; }
    .opt-manual-notice { margin: 12px 0; padding: 10px 12px; border: 2px solid #b42318; background: #fff2f2; color: #b42318; font-weight: 700; text-align: center; border-radius: 10px; }
  </style>
</head>
<body>
  <h1>${title}</h1>
  <div class="date">Дата: ${safeDate}</div>
  <div class="results">
    ${content || "<div>Нет данных</div>"}
  </div>
</body>
</html>`;

    const blob = new Blob([html], { type: "text/html;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `cut-result-${safeDate}.html`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  });
}
