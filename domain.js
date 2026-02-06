/**
 * Domain models for Rapid DGL.
 * Keep this file free of UI/IO logic.
 */

/**
 * @typedef {Object} Part
 * @property {string} orderId
 * @property {string} productId
 * @property {string} name
 * @property {number} lengthMm
 * @property {string} angles
 * @property {number} quantity
 * @property {string} profileCode
 * @property {string} position
 * @property {number|null} widthMm
 * @property {number|null} heightMm
 * @property {string} orient
 */

/**
 * @typedef {Object} Product
 * @property {string} productId
 * @property {string} orderId
 * @property {Part[]} parts
 */

/**
 * @typedef {Object} Order
 * @property {string} orderId
 * @property {Product[]} products
 */

/**
 * @typedef {Object} CutItem
 * @property {string} partId
 * @property {number} lengthMm
 * @property {number} quantity
 * @property {number} leftAngle
 * @property {number} rightAngle
 * @property {Object} meta
 */

/**
 * @typedef {Object} CutBar
 * @property {string} barId
 * @property {number} barLengthMm
 * @property {CutItem[]} items
 * @property {number} wasteMm
 */

/**
 * @typedef {Object} CutPlan
 * @property {CutBar[]} bars
 * @property {Object} settings
 */

/**
 * @typedef {Object} MachineJob
 * @property {string} text
 * @property {string} encoding
 */

/**
 * Normalize a raw part object into a Part.
 * @param {Partial<Part>} raw
 * @returns {Part}
 */
function createPart(raw = {}) {
  return {
    orderId: String(raw.orderId || "").trim(),
    productId: String(raw.productId || "").trim(),
    name: String(raw.name || "").trim(),
    lengthMm: Number(raw.lengthMm || 0),
    angles: String(raw.angles || "").trim(),
    quantity: Number(raw.quantity || 0),
    profileCode: String(raw.profileCode || "").trim(),
    position: String(raw.position || "").trim(),
    widthMm: raw.widthMm === null || raw.widthMm === undefined || raw.widthMm === "" ? null : Number(raw.widthMm),
    heightMm: raw.heightMm === null || raw.heightMm === undefined || raw.heightMm === "" ? null : Number(raw.heightMm),
    orient: String(raw.orient || "").trim()
  };
}

/**
 * Create an Order wrapper.
 * @param {string} orderId
 * @param {Product[]} products
 * @returns {Order}
 */
function createOrder(orderId, products = []) {
  return {
    orderId: String(orderId || "").trim(),
    products
  };
}

/**
 * Create a Product wrapper.
 * @param {string} productId
 * @param {string} orderId
 * @param {Part[]} parts
 * @returns {Product}
 */
function createProduct(productId, orderId, parts = []) {
  return {
    productId: String(productId || "").trim(),
    orderId: String(orderId || "").trim(),
    parts
  };
}
