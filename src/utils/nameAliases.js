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

/**
 * 提取姓名的最後兩個字元
 * @param {string} name - 原始姓名
 * @returns {string} 最後兩個字元
 */
export function getLastNameTwoChars(name) {
  if (!name || typeof name !== 'string') return '';
  const trimmed = name.trim();
  if (trimmed.length === 0) return '';
  // 如果姓名長度小於等於2，直接返回
  if (trimmed.length <= 2) return trimmed;
  // 返回最後兩個字元
  return trimmed.slice(-2);
}

/**
 * 根據最後兩個字元在 nameAliases 中查找對應的標準名稱
 * @param {string} lastTwoChars - 姓名的最後兩個字元
 * @returns {string|null} 標準名稱，如果找不到則返回 null
 */
export function findNameByLastTwoChars(lastTwoChars) {
  if (!lastTwoChars || lastTwoChars.length === 0) return null;
  
  const map = getAliasMap();
  
  // 遍歷映射表，查找值（標準名稱）的最後兩個字元匹配的項目
  for (const [key, value] of Object.entries(map)) {
    const valueLastTwo = getLastNameTwoChars(value);
    if (valueLastTwo === lastTwoChars) {
      return value; // 返回標準名稱
    }
  }
  
  return null;
}

