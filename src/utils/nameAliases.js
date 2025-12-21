import nameAliases from '../../nameAliases.json';

/**
 * 別名映射表
 * 將各種人名變體映射到標準名稱
 */
let aliasMap = null;

/**
 * 載入別名映射表
 * @returns {Object} 別名映射物件 { 原始名稱: 標準名稱 }
 */
export function getAliasMap() {
  if (aliasMap === null) {
    aliasMap = nameAliases || {};
  }
  return aliasMap;
}

/**
 * 將人名轉換為標準名稱（使用別名映射）
 * @param {string} name - 原始人名
 * @returns {string} 標準名稱
 */
export function normalizeName(name) {
  if (!name) return name;
  
  const map = getAliasMap();
  // 如果映射表中存在，返回標準名稱；否則返回原始名稱
  return map[name] || name;
}

/**
 * 將人名陣列轉換為標準名稱陣列
 * @param {string[]} names - 原始人名陣列
 * @returns {string[]} 標準名稱陣列
 */
export function normalizeNames(names) {
  if (!names || !Array.isArray(names)) return [];
  return names.map(name => normalizeName(name));
}

