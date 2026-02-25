/**
 * Simple cutting optimizer (best-fit decreasing).
 * items: [{ lengthMm, qty }]
 * rules: { scrapMax, minUsefulLeftover } optional
 */
function optimizeCut(items, stock, kerf, rules) {
  const parts = [];
  const hasRules =
    rules &&
    Number.isFinite(Number(rules.scrapMax)) &&
    Number.isFinite(Number(rules.minUsefulLeftover));
  const scrapMax = hasRules ? Number(rules.scrapMax) : 0;
  const minUsefulLeftover = hasRules ? Number(rules.minUsefulLeftover) : 0;

  function isBadLeftover(leftover) {
    if (!hasRules) return false;
    return leftover > scrapMax && leftover < minUsefulLeftover;
  }

  for (const it of items || []) {
    const length = Number(it.lengthMm || it.length || 0);
    const qty = Math.floor(Number(it.qty || it.quantity || 0));
    const meta = it.meta;
    if (!Number.isFinite(length) || length <= 0) continue;
    if (!Number.isFinite(qty) || qty <= 0) continue;
    for (let i = 0; i < qty; i++) {
      parts.push({ length, size: length + kerf, meta });
    }
  }

  parts.sort((a, b) => b.size - a.size);

  const bins = [];
  const unfit = [];

  for (const part of parts) {
    if (part.size > stock) {
      unfit.push(part);
      continue;
    }

    let bestIdx = -1;
    let bestLeftover = Infinity;
    let bestIsBad = true;

    for (let i = 0; i < bins.length; i++) {
      const remain = stock - bins[i].used;
      if (part.size > remain) continue;

      const leftover = remain - part.size;
      const bad = isBadLeftover(leftover);

      if (bestIdx === -1) {
        bestIdx = i;
        bestLeftover = leftover;
        bestIsBad = bad;
        continue;
      }

      if (bestIsBad && !bad) {
        bestIdx = i;
        bestLeftover = leftover;
        bestIsBad = false;
        continue;
      }

      if (bestIsBad === bad && leftover < bestLeftover) {
        bestIdx = i;
        bestLeftover = leftover;
      }
    }

    if (bestIdx === -1) {
      bins.push({ used: part.size, parts: [part] });
      continue;
    }

    const newBinLeftover = stock - part.size;
    const newBinIsBad = isBadLeftover(newBinLeftover);
    if (hasRules && bestIsBad && !newBinIsBad) {
      bins.push({ used: part.size, parts: [part] });
      continue;
    }

    bins[bestIdx].used += part.size;
    bins[bestIdx].parts.push(part);
  }

  const binsWithLeft = bins.map((bin) => ({
    ...bin,
    leftover: stock - bin.used
  }));

  return { bins: binsWithLeft, unfit };
}

window.optimizeCut = optimizeCut;
