import { getLastNameTwoChars, findNameByLastTwoChars } from './nameAliases.js';
import reportType from '../../reportType.js';
import { getYear } from './dateParser.js';

/**
 * 根據參與內容匹配 reportType
 * @param {string} content - 參與內容
 * @returns {string|null} 匹配到的 reportType，如果沒有匹配則返回 null
 */
function matchReportType(content) {
  if (!content || typeof content !== 'string') return null;
  
  // 檢查參與內容是否包含 reportType 中的任何一個字串
  for (const type of reportType) {
    if (content.includes(type)) {
      return type;
    }
  }
  
  return null;
}

/**
 * 處理時數登錄表數據
 * 提取志工姓名最後兩個字元，使用 nameAliases 進行比對，統計時數
 * 同時檢查參與內容是否匹配 reportType
 * @param {Array} data - 時數登錄表原始數據
 * @returns {Array} 處理後的數據，包含標準化姓名和分類後的參與內容
 */
export function processHourLogData(data) {
  const processed = [];
  const errors = [];
  
  data.forEach(record => {
    const originalName = record.name;
    if (!originalName) return;
    
    // 提取姓名最後兩個字元
    const lastTwoChars = getLastNameTwoChars(originalName);
    
    // 在 nameAliases 中查找對應的標準名稱
    const standardName = findNameByLastTwoChars(lastTwoChars);
    
    if (!standardName) {
      // 找不到對應的別名，記錄錯誤
      const errorMsg = `找不到志工姓名對應: ${originalName} (最後兩字: ${lastTwoChars})`;
      errors.push(errorMsg);
      console.error(errorMsg);
    }
    
    // 檢查參與內容是否匹配 reportType
    const originalContent = record.content || '';
    const matchedType = matchReportType(originalContent);
    
    if (!matchedType) {
      // 沒有匹配到任何 reportType，記錄錯誤
      const errorMsg = `找不到參與內容分類: "${originalContent}" (第 ${record.rowNumber} 行)`;
      errors.push(errorMsg);
      console.error(errorMsg);
    }
    
    // 使用匹配到的 reportType 或原始內容
    processed.push({
      ...record,
      standardName: standardName || lastTwoChars || originalName,
      matchedContentType: matchedType || '未分類',
    });
  });
  
  // 如果有錯誤，打印錯誤訊息
  if (errors.length > 0) {
    console.error('時數登錄表處理錯誤：');
    errors.forEach(error => console.error('  -', error));
  }
  
  return processed;
}

/**
 * 計算各志工依參與內容分類的時數
 * 使用 reportType 作為分類標準
 * @param {Array} processedData - 處理後的時數登錄表數據
 * @returns {Array} 統計資料 [{ name: string, [參與內容1]: number, [參與內容2]: number, ... }]
 */
export function calculateVolunteerHoursByContent(processedData) {
  // 預設僅統計 2025 年（保留既有圖表 21 的行為）
  return calculateVolunteerHoursByContentWithOptions(processedData, { year: 2025 });
}

/**
 * 計算各志工依參與內容分類的時數（可指定年份篩選）
 * - options.year: 單一年份
 * - options.startYear / options.endYear: 年份區間（含端點）
 * - options.years: 指定年份清單（優先於 year / startYear-endYear）
 * @param {Array} processedData
 * @param {Object} options
 * @returns {{data: Array, contentTypes: Array}}
 */
export function calculateVolunteerHoursByContentWithOptions(processedData, options = {}) {
  const stats = {};
  // 使用 reportType 作為內容類型集合（加上「未分類」）
  const contentTypes = new Set([...reportType, '未分類']);
  const {
    years,
    year,
    startYear,
    endYear,
  } = options;

  const yearsSet = Array.isArray(years) && years.length > 0 ? new Set(years) : null;
  const useYearRange = typeof startYear === 'number' && typeof endYear === 'number';
  const useSingleYear = typeof year === 'number';
  
  // 統計每個志工和參與內容的時數
  processedData.forEach(record => {
    const recordYear = record?.date ? getYear(record.date) : 0;
    if (yearsSet && !yearsSet.has(recordYear)) return;
    if (!yearsSet && useYearRange && (recordYear < startYear || recordYear > endYear)) return;
    if (!yearsSet && !useYearRange && useSingleYear && recordYear !== year) return;

    const name = record.standardName;
    // 使用匹配到的 reportType 或「未分類」
    const content = record.matchedContentType || '未分類';
    const hours = record.hours || 0;
    
    if (!name || hours <= 0) return;
    
    // 初始化志工統計
    if (!stats[name]) {
      stats[name] = {};
    }
    
    // 累加時數
    if (!stats[name][content]) {
      stats[name][content] = 0;
    }
    
    stats[name][content] += hours;
  });
  
  // 轉換為陣列格式，適合圖表顯示
  const result = Object.entries(stats).map(([name, contents]) => {
    const item = { name };
    // 為每個 reportType 添加數值（如果該志工沒有該類型，則為0）
    contentTypes.forEach(content => {
      item[content] = contents[content] || 0;
    });
    return item;
  });
  
  // 按總時數排序（降序）
  result.sort((a, b) => {
    const totalA = Object.entries(a).reduce((sum, [key, val]) => {
      // 跳過 name 欄位，只計算數值
      if (key === 'name') return sum;
      return typeof val === 'number' ? sum + val : sum;
    }, 0);
    const totalB = Object.entries(b).reduce((sum, [key, val]) => {
      // 跳過 name 欄位，只計算數值
      if (key === 'name') return sum;
      return typeof val === 'number' ? sum + val : sum;
    }, 0);
    return totalB - totalA;
  });
  
  // 過濾掉沒有數據的 reportType（只保留有數據的類型）
  const usedContentTypes = new Set();
  result.forEach(item => {
    Object.keys(item).forEach(key => {
      if (key !== 'name' && item[key] > 0) {
        usedContentTypes.add(key);
      }
    });
  });
  
  return {
    data: result,
    contentTypes: Array.from(usedContentTypes).sort(),
  };
}

/**
 * 計算指定參與內容分類的「總時數」（單一系列），可指定年份篩選
 * 主要用於像「2023–2025 回流訓練時數」這類單柱狀圖。
 * @param {Array} processedData - 處理後的時數登錄表數據
 * @param {Object} options
 * @param {string} options.contentType - 例如 "回流訓練"
 * @param {number} [options.year]
 * @param {number} [options.startYear]
 * @param {number} [options.endYear]
 * @param {number[]} [options.years]
 * @returns {{data: Array<{name: string, hours: number}>, contentType?: string, year?: number, startYear?: number, endYear?: number, years?: number[]}}
 */
export function calculateVolunteerTotalHoursForContentType(processedData, options = {}) {
  const { contentType, years, year, startYear, endYear } = options;
  const yearsSet = Array.isArray(years) && years.length > 0 ? new Set(years) : null;
  const useYearRange = typeof startYear === 'number' && typeof endYear === 'number';
  const useSingleYear = typeof year === 'number';

  const stats = {};

  processedData.forEach(record => {
    const recordYear = record?.date ? getYear(record.date) : 0;
    if (yearsSet && !yearsSet.has(recordYear)) return;
    if (!yearsSet && useYearRange && (recordYear < startYear || recordYear > endYear)) return;
    if (!yearsSet && !useYearRange && useSingleYear && recordYear !== year) return;

    const name = record.standardName;
    const hours = record.hours || 0;
    const recordContent = record.matchedContentType || '未分類';

    if (!name || hours <= 0) return;
    if (contentType && recordContent !== contentType) return;

    if (!stats[name]) stats[name] = 0;
    stats[name] += hours;
  });

  const data = Object.entries(stats)
    .map(([name, hours]) => ({ name, hours }))
    .sort((a, b) => b.hours - a.hours);

  return { data, contentType, years, year, startYear, endYear };
}
