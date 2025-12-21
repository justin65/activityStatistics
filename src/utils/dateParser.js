/**
 * 解析各種日期格式（2025年表格專用）
 * 支援格式：
 * - "1/3-5" -> 2025/1/3 至 2025/1/5（返回開始日期）
 * - "11/21-11/25" -> 2025/11/21 至 2025/11/25（返回開始日期）
 * - "1/3" -> 2025/1/3
 * - 括號中的內容（星期）會被忽略
 * @param {string|number|Date} dateValue - Excel 中的日期值
 * @returns {Date|null} 解析後的 Date 物件（開始日期），失敗則返回 null
 */
export function parseDate(dateValue) {
  if (!dateValue) return null;
  
  // 如果已經是 Date 物件
  if (dateValue instanceof Date) {
    return isNaN(dateValue.getTime()) ? null : dateValue;
  }
  
  // 如果是數字（Excel 日期序列號）
  if (typeof dateValue === 'number') {
    // Excel 日期從 1900-01-01 開始計算（但 Excel 錯誤地認為 1900 是閏年）
    // 所以需要減去 1 天
    const excelEpoch = new Date(1899, 11, 30);
    const date = new Date(excelEpoch.getTime() + dateValue * 24 * 60 * 60 * 1000);
    return isNaN(date.getTime()) ? null : date;
  }
  
  // 轉換為字串
  let dateStr = String(dateValue).trim();
  if (!dateStr) return null;
  
  // 移除括號及其內容（星期資訊）
  // 同時處理中文括號（全形）和英文括號（半形）
  dateStr = dateStr
    .replace(/\([^)]*\)/g, '')  // 移除英文括號及其內容
    .replace(/（[^）]*）/g, '')  // 移除中文括號及其內容
    .trim();
  
  // 預設年份為 2025
  const defaultYear = 2025;
  
  // 處理日期範圍格式：如 "1/3-5" 或 "11/21-11/25"
  // 先檢查是否有範圍符號 "-"
  if (dateStr.includes('-')) {
    const parts = dateStr.split('-');
    if (parts.length === 2) {
      const startPart = parts[0].trim();
      const endPart = parts[1].trim();
      
      // 解析開始日期
      // 格式可能是 "1/3" 或 "11/21"
      const startMatch = startPart.match(/^(\d{1,2})\/(\d{1,2})$/);
      if (startMatch) {
        const month = parseInt(startMatch[1], 10);
        const day = parseInt(startMatch[2], 10);
        
        // 驗證並建立開始日期
        if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
          const startDate = new Date(defaultYear, month - 1, day);
          if (startDate.getFullYear() === defaultYear && 
              startDate.getMonth() === month - 1 && 
              startDate.getDate() === day) {
            // 返回開始日期（用於統計月份）
            return startDate;
          }
        }
      }
    }
  }
  
  // 處理單一日期格式：如 "1/3" 或 "11/21"
  const singleDateMatch = dateStr.match(/^(\d{1,2})\/(\d{1,2})$/);
  if (singleDateMatch) {
    const month = parseInt(singleDateMatch[1], 10);
    const day = parseInt(singleDateMatch[2], 10);
    
    if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
      const date = new Date(defaultYear, month - 1, day);
      if (date.getFullYear() === defaultYear && 
          date.getMonth() === month - 1 && 
          date.getDate() === day) {
        return date;
      }
    }
  }
  
  // 嘗試其他標準格式（保留原有邏輯作為備用）
  const patterns = [
    // YYYY/MM/DD 或 YYYY/M/D
    /^(\d{4})\/(\d{1,2})\/(\d{1,2})$/,
    // YYYY-MM-DD 或 YYYY-M-D
    /^(\d{4})-(\d{1,2})-(\d{1,2})$/,
    // YYYY年MM月DD日
    /^(\d{4})年(\d{1,2})月(\d{1,2})日$/,
    // MM/DD/YYYY 或 M/D/YYYY
    /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/,
    // YYYY.MM.DD
    /^(\d{4})\.(\d{1,2})\.(\d{1,2})$/,
  ];
  
  for (const pattern of patterns) {
    const match = dateStr.match(pattern);
    if (match) {
      let year, month, day;
      
      if (pattern.source.includes('年')) {
        // YYYY年MM月DD日
        year = parseInt(match[1], 10);
        month = parseInt(match[2], 10);
        day = parseInt(match[3], 10);
      } else if (pattern.source.startsWith('^\\d{4}')) {
        // YYYY/MM/DD 或 YYYY-MM-DD
        year = parseInt(match[1], 10);
        month = parseInt(match[2], 10);
        day = parseInt(match[3], 10);
      } else {
        // MM/DD/YYYY
        month = parseInt(match[1], 10);
        day = parseInt(match[2], 10);
        year = parseInt(match[3], 10);
      }
      
      // 驗證日期有效性
      if (year >= 1900 && year <= 2100 && month >= 1 && month <= 12 && day >= 1 && day <= 31) {
        const date = new Date(year, month - 1, day);
        // 驗證日期是否正確
        if (date.getFullYear() === year && date.getMonth() === month - 1 && date.getDate() === day) {
          return date;
        }
      }
    }
  }
  
  // 嘗試使用 JavaScript 內建的 Date 解析
  const parsed = new Date(dateStr);
  if (!isNaN(parsed.getTime())) {
    return parsed;
  }
  
  // 所有方法都失敗
  console.warn('無法解析日期:', dateValue);
  return null;
}

/**
 * 取得日期的月份（1-12）
 * @param {Date} date 
 * @returns {number} 月份
 */
export function getMonth(date) {
  if (!date) return 0;
  return date.getMonth() + 1;
}

/**
 * 取得日期的年份
 * @param {Date} date 
 * @returns {number} 年份
 */
export function getYear(date) {
  if (!date) return 0;
  return date.getFullYear();
}

/**
 * 格式化月份顯示（例如：2025年1月）
 * @param {number} year 
 * @param {number} month 
 * @returns {string}
 */
export function formatMonth(year, month) {
  return `${year}年${month}月`;
}

