import { getMonth, getYear, formatMonth } from './dateParser.js';
import { normalizeName } from './nameAliases.js';

/**
 * 解析名字中的日期格式，提取真實名字和計算天數
 * 支援格式：
 * - 建宇(8/23) -> { name: '建宇', days: 1 }
 * - 建宇(21-22) -> { name: '建宇', days: 2 }（需要活動日期來確定月份）
 * - 建宇(7/22-23) -> { name: '建宇', days: 2 }
 * - 建宇(6/30-7/1) -> { name: '建宇', days: 2 }
 * @param {string} nameWithDate - 包含日期的名字
 * @param {Date} activityDate - 活動日期（用於處理只有日期沒有月份的情況）
 * @returns {Object} { name: string, days: number, error?: string }
 */
function parseNameWithDate(nameWithDate, activityDate = null) {
  if (!nameWithDate || typeof nameWithDate !== 'string') {
    return { name: nameWithDate || '', days: 0, error: '名字為空或格式錯誤' };
  }

  const trimmed = nameWithDate.trim();
  
  // 檢查是否有括號（支援中文和英文括號）
  const bracketMatch = trimmed.match(/^(.+?)[（(](.+?)[）)]$/);
  
  if (!bracketMatch) {
    // 沒有括號，直接返回原名字
    return { name: trimmed, days: 0 };
  }

  const name = bracketMatch[1].trim();
  const dateStr = bracketMatch[2].trim();

  if (!name) {
    return { name: '', days: 0, error: `無法提取名字: ${trimmed}` };
  }

  // 預設年份為 2025（與 dateParser 一致）
  const defaultYear = 2025;
  
  // 從活動日期獲取月份（如果有的話）
  let defaultMonth = null;
  if (activityDate && activityDate instanceof Date && !isNaN(activityDate.getTime())) {
    defaultMonth = activityDate.getMonth() + 1; // getMonth() 返回 0-11
  }

  try {
    // 處理日期範圍格式
    if (dateStr.includes('-')) {
      const parts = dateStr.split('-');
      if (parts.length === 2) {
        const startPart = parts[0].trim();
        const endPart = parts[1].trim();

        // 解析開始日期
        let startMonth, startDay;
        if (startPart.includes('/')) {
          const startMatch = startPart.match(/^(\d{1,2})\/(\d{1,2})$/);
          if (!startMatch) {
            return { name, days: 0, error: `無法解析開始日期: ${startPart}` };
          }
          startMonth = parseInt(startMatch[1], 10);
          startDay = parseInt(startMatch[2], 10);
        } else {
          // 只有日期，沒有月份（例如：21-22）
          // 從活動日期獲取月份
          if (defaultMonth === null) {
            return { name, days: 0, error: `日期格式需要包含月份或活動日期: ${dateStr}` };
          }
          startMonth = defaultMonth;
          const day = parseInt(startPart, 10);
          if (isNaN(day) || day < 1 || day > 31) {
            return { name, days: 0, error: `無法解析開始日期: ${startPart}` };
          }
          startDay = day;
        }

        // 解析結束日期
        let endMonth, endDay;
        if (endPart.includes('/')) {
          const endMatch = endPart.match(/^(\d{1,2})\/(\d{1,2})$/);
          if (!endMatch) {
            return { name, days: 0, error: `無法解析結束日期: ${endPart}` };
          }
          endMonth = parseInt(endMatch[1], 10);
          endDay = parseInt(endMatch[2], 10);
        } else {
          // 只有日期，沒有月份（例如：22-23）
          // 假設與開始日期同一個月（如果開始日期也沒有月份，則使用活動日期的月份）
          if (startPart.includes('/')) {
            endMonth = startMonth;
          } else {
            endMonth = defaultMonth || startMonth;
          }
          const day = parseInt(endPart, 10);
          if (isNaN(day) || day < 1 || day > 31) {
            return { name, days: 0, error: `無法解析結束日期: ${endPart}` };
          }
          endDay = day;
        }

        // 驗證日期有效性
        if (startMonth < 1 || startMonth > 12 || startDay < 1 || startDay > 31 ||
            endMonth < 1 || endMonth > 12 || endDay < 1 || endDay > 31) {
          return { name, days: 0, error: `日期無效: ${dateStr}` };
        }

        // 計算天數
        const startDate = new Date(defaultYear, startMonth - 1, startDay);
        const endDate = new Date(defaultYear, endMonth - 1, endDay);

        // 驗證日期是否正確
        if (startDate.getMonth() !== startMonth - 1 || startDate.getDate() !== startDay ||
            endDate.getMonth() !== endMonth - 1 || endDate.getDate() !== endDay) {
          return { name, days: 0, error: `日期無效: ${dateStr}` };
        }

        // 計算天數差（包含開始和結束日期）
        const diffTime = endDate.getTime() - startDate.getTime();
        const diffDays = Math.floor(diffTime / (1000 * 60 * 60 * 24));
        const days = diffDays + 1; // 包含開始和結束日期

        if (days < 1) {
          return { name, days: 0, error: `日期範圍無效: ${dateStr}` };
        }

        return { name, days };
      }
    } else {
      // 單一日期格式（例如：8/23）
      const singleMatch = dateStr.match(/^(\d{1,2})\/(\d{1,2})$/);
      if (!singleMatch) {
        return { name, days: 0, error: `無法解析日期格式: ${dateStr}` };
      }

      const month = parseInt(singleMatch[1], 10);
      const day = parseInt(singleMatch[2], 10);

      if (month < 1 || month > 12 || day < 1 || day > 31) {
        return { name, days: 0, error: `日期無效: ${dateStr}` };
      }

      const date = new Date(defaultYear, month - 1, day);
      if (date.getMonth() !== month - 1 || date.getDate() !== day) {
        return { name, days: 0, error: `日期無效: ${dateStr}` };
      }

      return { name, days: 1 };
    }
  } catch (error) {
    return { name, days: 0, error: `解析日期時發生錯誤: ${error.message}` };
  }
}

/**
 * 過濾掉取消的活動
 * @param {Array} data - 原始資料
 * @returns {Array} 過濾後的資料
 */
export function filterCancelled(data) {
  return data.filter(record => {
    const status = record.status.toLowerCase();
    return !status.includes('取消');
  });
}

/**
 * 計算按月份和活動類型的統計（次數）
 * @param {Array} data - 過濾後的資料
 * @returns {Array} 統計資料
 */
export function calculateMonthlyActivityTypeCount(data) {
  const stats = {};
  
  data.forEach(record => {
    const year = getYear(record.date);
    const month = getMonth(record.date);
    const key = `${year}-${month}`;
    const monthLabel = formatMonth(year, month);
    const activityType = record.activityType || '未分類';
    
    if (!stats[key]) {
      stats[key] = {
        month: monthLabel,
        year,
        monthNum: month,
      };
    }
    
    if (!stats[key][activityType]) {
      stats[key][activityType] = 0;
    }
    
    stats[key][activityType]++;
  });
  
  return Object.values(stats).sort((a, b) => {
    if (a.year !== b.year) return a.year - b.year;
    return a.monthNum - b.monthNum;
  });
}

/**
 * 計算按月份和活動類型的統計（天數）
 * @param {Array} data - 過濾後的資料
 * @returns {Array} 統計資料
 */
export function calculateMonthlyActivityTypeDays(data) {
  const stats = {};
  
  data.forEach(record => {
    const year = getYear(record.date);
    const month = getMonth(record.date);
    const key = `${year}-${month}`;
    const monthLabel = formatMonth(year, month);
    const activityType = record.activityType || '未分類';
    const days = record.days || 0;
    
    if (!stats[key]) {
      stats[key] = {
        month: monthLabel,
        year,
        monthNum: month,
      };
    }
    
    if (!stats[key][activityType]) {
      stats[key][activityType] = 0;
    }
    
    stats[key][activityType] += days;
  });
  
  return Object.values(stats).sort((a, b) => {
    if (a.year !== b.year) return a.year - b.year;
    return a.monthNum - b.monthNum;
  });
}

/**
 * 計算按月份和縣市的統計（次數）
 * @param {Array} data - 過濾後的資料
 * @returns {Array} 統計資料
 */
export function calculateMonthlyCityCount(data) {
  const stats = {};
  
  data.forEach(record => {
    const year = getYear(record.date);
    const month = getMonth(record.date);
    const key = `${year}-${month}`;
    const monthLabel = formatMonth(year, month);
    const city = record.city || '未分類';
    
    if (!stats[key]) {
      stats[key] = {
        month: monthLabel,
        year,
        monthNum: month,
      };
    }
    
    if (!stats[key][city]) {
      stats[key][city] = 0;
    }
    
    stats[key][city]++;
  });
  
  return Object.values(stats).sort((a, b) => {
    if (a.year !== b.year) return a.year - b.year;
    return a.monthNum - b.monthNum;
  });
}

/**
 * 計算按月份和縣市的統計（天數）
 * @param {Array} data - 過濾後的資料
 * @returns {Array} 統計資料
 */
export function calculateMonthlyCityDays(data) {
  const stats = {};
  
  data.forEach(record => {
    const year = getYear(record.date);
    const month = getMonth(record.date);
    const key = `${year}-${month}`;
    const monthLabel = formatMonth(year, month);
    const city = record.city || '未分類';
    const days = record.days || 0;
    
    if (!stats[key]) {
      stats[key] = {
        month: monthLabel,
        year,
        monthNum: month,
      };
    }
    
    if (!stats[key][city]) {
      stats[key][city] = 0;
    }
    
    stats[key][city] += days;
  });
  
  return Object.values(stats).sort((a, b) => {
    if (a.year !== b.year) return a.year - b.year;
    return a.monthNum - b.monthNum;
  });
}

/**
 * 計算按人名的活動參與次數（使用別名映射合併）
 * 處理名字後面有日期的情況（例如：建宇(8/23)）
 * 按活動類型分組統計
 * @param {Array} data - 過濾後的資料
 * @returns {Array} 統計資料 [{ name: string, [活動類型1]: number, [活動類型2]: number, ... }]
 */
export function calculateParticipantCount(data) {
  const stats = {};
  const errors = [];
  const activityTypesSet = new Set();
  
  // 先收集所有活動類型
  data.forEach(record => {
    const activityType = record.activityType || '未分類';
    activityTypesSet.add(activityType);
  });
  
  data.forEach(record => {
    const activityType = record.activityType || '未分類';
    record.participants.forEach(participant => {
      // 處理新格式：participant 可能是 { name: string, days: number } 或 string
      let name, daysFromPrefix = 0;
      
      if (typeof participant === 'object' && participant.name) {
        // 新格式：日期前綴格式
        name = participant.name;
        daysFromPrefix = participant.days || 0;
      } else {
        // 舊格式：純字串
        name = participant;
      }
      
      // 解析名字中的日期（傳入活動日期以處理只有日期沒有月份的情況）
      const parsed = parseNameWithDate(name, record.date);
      
      // 如果有錯誤，記錄錯誤訊息
      if (parsed.error) {
        errors.push(`第 ${record.rowNumber} 行，參與人員 "${name}": ${parsed.error}`);
      }
      
      // 提取真實名字
      const realName = parsed.name;
      if (!realName) {
        return; // 跳過無法提取名字的情況
      }
      
      // 使用別名映射將人名轉換為標準名稱
      const normalizedName = normalizeName(realName);
      
      // 次數統計：每個活動記錄中的每個參與人員都只算1次
      // 日期信息只用於計算時數，不影響次數統計
      if (!stats[normalizedName]) {
        stats[normalizedName] = {};
        // 初始化所有活動類型為 0
        activityTypesSet.forEach(type => {
          stats[normalizedName][type] = 0;
        });
      }
      stats[normalizedName][activityType] = (stats[normalizedName][activityType] || 0) + 1;
    });
  });
  
  // 如果有錯誤，打印錯誤訊息
  if (errors.length > 0) {
    console.error('參與人員次數統計錯誤：');
    errors.forEach(error => console.error('  -', error));
  }
  
  // 轉換為陣列格式，並計算總次數用於排序
  const result = Object.entries(stats).map(([name, typeCounts]) => {
    const total = Object.values(typeCounts).reduce((sum, count) => sum + count, 0);
    return { name, ...typeCounts, total };
  }).sort((a, b) => b.total - a.total);
  
  // 移除 total 欄位（只保留 name 和活動類型）
  return result.map(({ total, ...rest }) => rest);
}

/**
 * 計算按人名的活動參與時數（使用別名映射合併）
 * 處理名字後面有日期的情況（例如：建宇(8/23)）
 * 如果名字中有日期，時數 = 天數 × 8小時；否則使用記錄中的時數
 * 按活動類型分組統計
 * @param {Array} data - 過濾後的資料
 * @returns {Array} 統計資料 [{ name: string, [活動類型1]: number, [活動類型2]: number, ... }]
 */
export function calculateParticipantHours(data) {
  const stats = {};
  const errors = [];
  const activityTypesSet = new Set();
  
  // 先收集所有活動類型
  data.forEach(record => {
    const activityType = record.activityType || '未分類';
    activityTypesSet.add(activityType);
  });
  
  data.forEach(record => {
    const recordHours = record.hours || 0;
    const activityType = record.activityType || '未分類';
    record.participants.forEach(participant => {
      // 處理新格式：participant 可能是 { name: string, days: number } 或 string
      let name, daysFromPrefix = 0;
      
      if (typeof participant === 'object' && participant.name) {
        // 新格式：日期前綴格式（例如：11/29-30：盈瑩）
        name = participant.name;
        daysFromPrefix = participant.days || 0;
      } else {
        // 舊格式：純字串（可能包含名字後面的日期，例如：盈瑩(11/29)）
        name = participant;
      }
      
      // 解析名字中的日期（傳入活動日期以處理只有日期沒有月份的情況）
      const parsed = parseNameWithDate(name, record.date);
      
      // 如果有錯誤，記錄錯誤訊息
      if (parsed.error) {
        errors.push(`第 ${record.rowNumber} 行，參與人員 "${name}": ${parsed.error}`);
      }
      
      // 提取真實名字
      const realName = parsed.name;
      if (!realName) {
        return; // 跳過無法提取名字的情況
      }
      
      // 使用別名映射將人名轉換為標準名稱
      const normalizedName = normalizeName(realName);
      
      // 計算時數
      let hours = 0;
      
      // 優先使用日期前綴的天數（新格式）
      if (daysFromPrefix > 0) {
        hours = daysFromPrefix * 8;
      } else if (parsed.days > 0) {
        // 其次使用名字中解析出的日期（舊格式：名字後面的日期）
        hours = parsed.days * 8;
      } else {
        // 否則使用記錄中的時數
        hours = recordHours;
      }
      
      if (!stats[normalizedName]) {
        stats[normalizedName] = {};
        // 初始化所有活動類型為 0
        activityTypesSet.forEach(type => {
          stats[normalizedName][type] = 0;
        });
      }
      stats[normalizedName][activityType] = (stats[normalizedName][activityType] || 0) + hours;
    });
  });
  
  // 如果有錯誤，打印錯誤訊息
  if (errors.length > 0) {
    console.error('參與人員時數統計錯誤：');
    errors.forEach(error => console.error('  -', error));
  }
  
  // 轉換為陣列格式，並計算總時數用於排序
  const result = Object.entries(stats).map(([name, typeHours]) => {
    const total = Object.values(typeHours).reduce((sum, hours) => sum + hours, 0);
    return { name, ...typeHours, total };
  }).sort((a, b) => b.total - a.total);
  
  // 移除 total 欄位（只保留 name 和活動類型）
  return result.map(({ total, ...rest }) => rest);
}

/**
 * 取得所有活動類型（用於圖表）
 * @param {Array} statsData - 統計資料
 * @returns {Array} 活動類型陣列
 */
export function getActivityTypes(statsData) {
  const types = new Set();
  statsData.forEach(item => {
    Object.keys(item).forEach(key => {
      if (key !== 'month' && key !== 'year' && key !== 'monthNum') {
        types.add(key);
      }
    });
  });
  return Array.from(types).sort();
}

/**
 * 計算依活動類型的總次數統計（用於圓餅圖）
 * @param {Array} data - 過濾後的資料
 * @returns {Array} 統計資料 [{ name: string, value: number }]
 */
export function calculateActivityTypeTotalCount(data) {
  const stats = {};
  
  data.forEach(record => {
    const activityType = record.activityType || '未分類';
    if (!stats[activityType]) {
      stats[activityType] = 0;
    }
    stats[activityType]++;
  });
  
  return Object.entries(stats)
    .map(([name, value]) => ({ name, value }))
    .sort((a, b) => b.value - a.value);
}

/**
 * 計算依活動類型的總天數統計（用於圓餅圖）
 * @param {Array} data - 過濾後的資料
 * @returns {Array} 統計資料 [{ name: string, value: number }]
 */
export function calculateActivityTypeTotalDays(data) {
  const stats = {};
  
  data.forEach(record => {
    const activityType = record.activityType || '未分類';
    const days = record.days || 0;
    if (!stats[activityType]) {
      stats[activityType] = 0;
    }
    stats[activityType] += days;
  });
  
  return Object.entries(stats)
    .map(([name, value]) => ({ name, value }))
    .sort((a, b) => b.value - a.value);
}

/**
 * 縣市區域分類（北中南東）- 用於排序
 */
const cityRegionsForSort = {
  // 北部
  '台北市': 1, '新北市': 1, '桃園市': 1, '新竹市': 1, '新竹縣': 1,
  '基隆市': 1, '宜蘭縣': 1,
  // 中部
  '台中市': 2, '苗栗縣': 2, '彰化縣': 2, '南投縣': 2, '雲林縣': 2,
  // 南部
  '高雄市': 3, '台南市': 3, '嘉義市': 3, '嘉義縣': 3, '屏東縣': 3, '澎湖縣': 3,
  // 東部
  '花蓮縣': 4, '台東縣': 4,
};

/**
 * 區域內排序順序（如果同區域有多個縣市）
 */
const cityOrderForSort = {
  // 北部
  '台北市': 1, '新北市': 2, '桃園市': 3, '新竹市': 4, '新竹縣': 5,
  '基隆市': 6, '宜蘭縣': 7,
  // 中部
  '台中市': 1, '苗栗縣': 2, '彰化縣': 3, '南投縣': 4, '雲林縣': 5,
  // 南部
  '高雄市': 1, '台南市': 2, '嘉義市': 3, '嘉義縣': 4, '屏東縣': 5, '澎湖縣': 6,
  // 東部
  '花蓮縣': 1, '台東縣': 2,
};

/**
 * 計算依縣市的總次數統計（用於圓餅圖）
 * @param {Array} data - 過濾後的資料
 * @returns {Array} 統計資料 [{ name: string, value: number }]
 */
export function calculateCityTotalCount(data) {
  const stats = {};
  
  data.forEach(record => {
    const city = record.city || '未分類';
    if (!stats[city]) {
      stats[city] = 0;
    }
    stats[city]++;
  });
  
  // 按照北中南東的順序排序，未分類放在最後
  return Object.entries(stats)
    .map(([name, value]) => ({ name, value }))
    .sort((a, b) => {
      // 明確處理「未分類」，讓它始終排在最後
      if (a.name === '未分類' && b.name !== '未分類') return 1;
      if (a.name !== '未分類' && b.name === '未分類') return -1;
      if (a.name === '未分類' && b.name === '未分類') return 0;
      
      const regionA = cityRegionsForSort[a.name] || 99; // 未定義的縣市放在最後（但未分類已經處理過了）
      const regionB = cityRegionsForSort[b.name] || 99;
      
      if (regionA !== regionB) {
        return regionA - regionB;
      }
      
      // 同區域內按順序排序
      const orderA = cityOrderForSort[a.name] || 999;
      const orderB = cityOrderForSort[b.name] || 999;
      if (orderA !== orderB) {
        return orderA - orderB;
      }
      
      // 如果都是未定義的縣市（且不是未分類），按名稱排序
      return a.name.localeCompare(b.name, 'zh-TW');
    });
}

/**
 * 計算依縣市的總天數統計（用於圓餅圖）
 * @param {Array} data - 過濾後的資料
 * @returns {Array} 統計資料 [{ name: string, value: number }]
 */
export function calculateCityTotalDays(data) {
  const stats = {};
  
  data.forEach(record => {
    const city = record.city || '未分類';
    const days = record.days || 0;
    if (!stats[city]) {
      stats[city] = 0;
    }
    stats[city] += days;
  });
  
  // 按照北中南東的順序排序，未分類放在最後
  return Object.entries(stats)
    .map(([name, value]) => ({ name, value }))
    .sort((a, b) => {
      // 明確處理「未分類」，讓它始終排在最後
      if (a.name === '未分類' && b.name !== '未分類') return 1;
      if (a.name !== '未分類' && b.name === '未分類') return -1;
      if (a.name === '未分類' && b.name === '未分類') return 0;
      
      const regionA = cityRegionsForSort[a.name] || 99; // 未定義的縣市放在最後（但未分類已經處理過了）
      const regionB = cityRegionsForSort[b.name] || 99;
      
      if (regionA !== regionB) {
        return regionA - regionB;
      }
      
      // 同區域內按順序排序
      const orderA = cityOrderForSort[a.name] || 999;
      const orderB = cityOrderForSort[b.name] || 999;
      if (orderA !== orderB) {
        return orderA - orderB;
      }
      
      // 如果都是未定義的縣市（且不是未分類），按名稱排序
      return a.name.localeCompare(b.name, 'zh-TW');
    });
}

/**
 * 計算依地區的總次數統計（用於圓餅圖）
 * @param {Array} data - 過濾後的資料
 * @returns {Array} 統計資料 [{ name: string, value: number }]
 */
export function calculateRegionTotalCount(data) {
  const stats = {};
  
  data.forEach(record => {
    const city = record.city || '未分類';
    const region = getRegion(city);
    if (!stats[region]) {
      stats[region] = 0;
    }
    stats[region]++;
  });
  
  return Object.entries(stats)
    .map(([name, value]) => ({ name, value }))
    .sort((a, b) => {
      // 按照北中南東的順序排序
      const order = { '北部': 1, '中部': 2, '南部': 3, '東部': 4, '未分類': 5 };
      return (order[a.name] || 99) - (order[b.name] || 99);
    });
}

/**
 * 計算依地區的總天數統計（用於圓餅圖）
 * @param {Array} data - 過濾後的資料
 * @returns {Array} 統計資料 [{ name: string, value: number }]
 */
export function calculateRegionTotalDays(data) {
  const stats = {};
  
  data.forEach(record => {
    const city = record.city || '未分類';
    const region = getRegion(city);
    const days = record.days || 0;
    if (!stats[region]) {
      stats[region] = 0;
    }
    stats[region] += days;
  });
  
  return Object.entries(stats)
    .map(([name, value]) => ({ name, value }))
    .sort((a, b) => {
      // 按照北中南東的順序排序
      const order = { '北部': 1, '中部': 2, '南部': 3, '東部': 4, '未分類': 5 };
      return (order[a.name] || 99) - (order[b.name] || 99);
    });
}

/**
 * 取得所有縣市（用於圖表）
 * @param {Array} statsData - 統計資料
 * @returns {Array} 縣市陣列
 */
export function getCities(statsData) {
  const cities = new Set();
  statsData.forEach(item => {
    Object.keys(item).forEach(key => {
      if (key !== 'month' && key !== 'year' && key !== 'monthNum') {
        cities.add(key);
      }
    });
  });
  return Array.from(cities).sort();
}

/**
 * 縣市區域分類（北中南東）
 */
const cityRegions = {
  // 北部
  '台北市': '北部', '新北市': '北部', '桃園市': '北部', '新竹市': '北部', '新竹縣': '北部',
  '基隆市': '北部', '宜蘭縣': '北部',
  // 中部
  '台中市': '中部', '苗栗縣': '中部', '彰化縣': '中部', '南投縣': '中部', '雲林縣': '中部',
  // 南部
  '高雄市': '南部', '台南市': '南部', '嘉義市': '南部', '嘉義縣': '南部', '屏東縣': '南部', '澎湖縣': '南部',
  // 東部
  '花蓮縣': '東部', '台東縣': '東部',
};

/**
 * 將縣市轉換為地區
 * @param {string} city - 縣市名稱
 * @returns {string} 地區名稱（北部、中部、南部、東部、未分類）
 */
function getRegion(city) {
  return cityRegions[city] || '未分類';
}

/**
 * 計算按月份和地區的統計（次數）
 * @param {Array} data - 過濾後的資料
 * @returns {Array} 統計資料
 */
export function calculateMonthlyRegionCount(data) {
  const stats = {};
  
  data.forEach(record => {
    const year = getYear(record.date);
    const month = getMonth(record.date);
    const key = `${year}-${month}`;
    const monthLabel = formatMonth(year, month);
    const city = record.city || '未分類';
    const region = getRegion(city);
    
    if (!stats[key]) {
      stats[key] = {
        month: monthLabel,
        year,
        monthNum: month,
      };
    }
    
    if (!stats[key][region]) {
      stats[key][region] = 0;
    }
    
    stats[key][region]++;
  });
  
  return Object.values(stats).sort((a, b) => {
    if (a.year !== b.year) return a.year - b.year;
    return a.monthNum - b.monthNum;
  });
}

/**
 * 計算按月份和地區的統計（天數）
 * @param {Array} data - 過濾後的資料
 * @returns {Array} 統計資料
 */
export function calculateMonthlyRegionDays(data) {
  const stats = {};
  
  data.forEach(record => {
    const year = getYear(record.date);
    const month = getMonth(record.date);
    const key = `${year}-${month}`;
    const monthLabel = formatMonth(year, month);
    const city = record.city || '未分類';
    const region = getRegion(city);
    const days = record.days || 0;
    
    if (!stats[key]) {
      stats[key] = {
        month: monthLabel,
        year,
        monthNum: month,
      };
    }
    
    if (!stats[key][region]) {
      stats[key][region] = 0;
    }
    
    stats[key][region] += days;
  });
  
  return Object.values(stats).sort((a, b) => {
    if (a.year !== b.year) return a.year - b.year;
    return a.monthNum - b.monthNum;
  });
}

