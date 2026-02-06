// Auto-generated from 'Профиля отбор.xlsx'
const ALLOWED_ITEMS = new Set([
  's358.01',
  's358.02',
  's358.10',
  's358.16',
  's358.26',
  's571.11',
  's571.22',
  's670.02',
  's670.10',
  's670.11',
  's670.26',
  'xs358.01',
  'xs358.01 d',
  'xs358.02',
  'рама дверная 72 мм (58 серия)',
  'рама оконная 63 мм (experta) new',
  'рама оконная 63 мм (practica)',
  'рама оконная 63 мм (profecta)',
  'рама оконная 63 мм (prowin)',
  'створка дверная z 106 мм (70 серия)',
  'створка дверная z образная 98мм (58 серия)',
  'створка дверная т 106 мм (70 серия)',
  'створка дверная т 107 мм (58 серия)',
  'створка оконная 77 мм (experta)',
  'створка оконная 77 мм (practica)',
  'створка оконная 77 мм (profecta)',
  'створка оконная 77 мм (prowin)',
]);

function normalizeAllowed(value) {
  return String(value || '')
    .replace(/\u00a0/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

function isAllowedItem(value) {
  const norm = normalizeAllowed(value);
  if (ALLOWED_ITEMS.has(norm)) return true;
  for (const item of ALLOWED_ITEMS) {
    if (norm.startsWith(item)) return true;
  }
  return false;
}
