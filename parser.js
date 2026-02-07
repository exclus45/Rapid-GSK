/**
 * CSV parsing utilities.
 * Produces Part[] (normalized) using domain.createPart.
 */

/**
 * @param {string} line
 * @returns {string[]}
 */
function splitCsvLine(line, delimiter) {
  const out = [];
  let cur = "";
  let inQuotes = false;

  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    const next = line[i + 1];

    if (ch === '"') {
      if (inQuotes && next === '"') {
        cur += '"';
        i++;
      } else {
        inQuotes = !inQuotes;
      }
      continue;
    }

    if (!inQuotes && ch === delimiter) {
      out.push(cur);
      cur = "";
      continue;
    }

    cur += ch;
  }
  out.push(cur);
  return out;
}

/**
 * Try to detect delimiter by counting occurrences in the header.
 * @param {string} headerLine
 * @returns {string}
 */
function detectDelimiter(headerLine) {
  const candidates = [";", ",", "\t"];
  let best = candidates[0];
  let bestCount = -1;
  for (const c of candidates) {
    const count = headerLine.split(c).length - 1;
    if (count > bestCount) {
      bestCount = count;
      best = c;
    }
  }
  return best;
}

function parseNumber(str) {
  const cleaned = String(str || "")
    .replace(/мм\.?/gi, "")
    .replace(/[^0-9,\.]/g, "")
    .replace(",", ".");
  const num = Number(cleaned);
  return Number.isFinite(num) ? num : 0;
}

function parseLengthFromSize(str) {
  const text = String(str || "");
  const matches = text.match(/(\d+[,.]?\d*)/g);
  if (!matches || !matches.length) return 0;
  const last = matches[matches.length - 1].replace(",", ".");
  const num = Number(last);
  return Number.isFinite(num) ? Math.round(num) : 0;
}

function extractAngles(text) {
  const matches = String(text || "").match(/(\d{1,3}(?:[.,]\d+)?)\s*°/g);
  if (!matches || matches.length < 2) return "";
  const a = matches[0].match(/\d{1,3}(?:[.,]\d+)?/)[0];
  const b = matches[1].match(/\d{1,3}(?:[.,]\d+)?/)[0];
  return `${a} - ${b}`;
}

function parseOrderAndProduct(str) {
  const raw = String(str || "").trim();
  const parts = raw.split("/").map((s) => s.trim());
  return {
    orderId: parts[0] || "",
    productId: parts[1] || ""
  };
}

function detectFixedLayout(lines, delimiter) {
  for (const line of lines) {
    const cols = splitCsvLine(line, delimiter);
    if (cols.length < 8) continue;
    const article = String(cols[1] || "").trim();
    if (!article) continue;

    let sizeIdx = -1;
    let orderIdx = -1;
    for (let i = 0; i < cols.length; i++) {
      const v = String(cols[i] || "").trim();
      if (sizeIdx === -1 && /мм/i.test(v)) sizeIdx = i;
      if (orderIdx === -1 && v.includes("/")) orderIdx = i;
    }
    // If size was picked from the name column, search only after the name.
    if (sizeIdx <= 2) {
      sizeIdx = -1;
      for (let i = 3; i < cols.length; i++) {
        const v = String(cols[i] || "").trim();
        if (/мм/i.test(v)) {
          sizeIdx = i;
          break;
        }
      }
    }
    if (sizeIdx === -1 || orderIdx === -1) continue;

    let qtyIdx = orderIdx - 1;
    if (qtyIdx < 0) qtyIdx = -1;
    const qtyVal = qtyIdx >= 0 ? String(cols[qtyIdx] || "").trim() : "";
    if (qtyIdx >= 0 && !/^\d+$/.test(qtyVal)) {
      qtyIdx = -1;
    }

    let orientIdx = -1;
    for (let i = sizeIdx + 1; i < orderIdx; i++) {
      const v = String(cols[i] || "").trim();
      if (v && !/^\d+$/.test(v)) {
        orientIdx = i;
        break;
      }
    }

    if (qtyIdx === -1) continue;

    return {
      article: 1,
      name: 2,
      size: sizeIdx,
      qty: qtyIdx,
      orderPos: orderIdx,
      orient: orientIdx
    };
  }
  return null;
}

/**
 * Parse CSV text into Part[] (fixed 5 columns).
 * Columns: 1) Артикул 2) Название 3) Размер 4) Кол-во 5) Заказ/Изделие
 * @param {string} text
 * @returns {Part[]}
 */
function parsePartsFromCsv(text) {
  const lines = String(text || "")
    .split(/\r?\n/)
    .filter((l) => l.trim() !== "");

  if (!lines.length) return [];

  const delimiter = detectDelimiter(lines[0]);
  const layout = detectFixedLayout(lines, delimiter) || {
    article: 0,
    name: 1,
    size: 2,
    qty: 3,
    orderPos: 4,
    orient: -1
  };

  const parts = [];

  for (const line of lines) {
    const cols = splitCsvLine(line, delimiter);
    if (cols.length <= layout.orderPos) continue;

    const article = String(cols[layout.article] || "").trim();
    const name = String(cols[layout.name] || "").trim();
    let size = String(cols[layout.size] || "").trim();
    const qty = String(cols[layout.qty] || "").trim();
    const orderPos = String(cols[layout.orderPos] || "").trim();
    const orient = layout.orient >= 0 ? String(cols[layout.orient] || "").trim() : "";

    if (!article || /^счет/i.test(article) || /^стр\./i.test(article)) continue;
    const articleLower = article.toLowerCase();
    const nameLower = name.toLowerCase();
    const orderLower = orderPos.toLowerCase();
    const orientLower = orient.toLowerCase();
    if (
      articleLower === "артикул" ||
      nameLower === "название" ||
      orderLower.includes("заказ") ||
      orientLower.startsWith("ориент")
    ) {
      continue;
    }
    
    if (!isAllowedItem(article) && !isAllowedItem(name)) continue;

    const { orderId, productId } = parseOrderAndProduct(orderPos);

    let lengthMm = parseLengthFromSize(size);
    if (!lengthMm) {
      // Fallback: find last column with "мм" between name and order.
      let fallback = "";
      for (let i = 3; i < cols.length; i++) {
        const v = String(cols[i] || "").trim();
        if (/мм/i.test(v)) fallback = v;
      }
      if (fallback) {
        lengthMm = parseLengthFromSize(fallback);
      }
    }

    const angleText = extractAngles(cols.slice(3, layout.orderPos).join(" "));

    const raw = {
      orderId,
      productId,
      name,
      lengthMm,
      angles: angleText,
      quantity: parseNumber(qty),
      profileCode: article,
      position: "",
      widthMm: "",
      heightMm: "",
      orient
    };

    const part = createPart(raw);
    part.id = `p_${Date.now().toString(36)}_${Math.random().toString(36).slice(2)}`;
    parts.push(part);
  }

  return parts;
}

