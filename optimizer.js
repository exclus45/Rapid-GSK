/**
 * Simple cutting optimizer (best-fit decreasing).
 * items: [{ lengthMm, qty }]
 */
function optimizeCut(items, stock, kerf) {
  const parts = [];

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
    let bestRemain = Infinity;

    for (let i = 0; i < bins.length; i++) {
      const remain = stock - bins[i].used;
      if (part.size <= remain && remain - part.size < bestRemain) {
        bestRemain = remain - part.size;
        bestIdx = i;
      }
    }

    if (bestIdx === -1) {
      bins.push({ used: part.size, parts: [part] });
    } else {
      bins[bestIdx].used += part.size;
      bins[bestIdx].parts.push(part);
    }
  }

  const binsWithLeft = bins.map((bin) => ({
    ...bin,
    leftover: stock - bin.used
  }));

  return { bins: binsWithLeft, unfit };
}

window.optimizeCut = optimizeCut;
