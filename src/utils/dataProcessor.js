import { getMonth, getYear, formatMonth } from './dateParser.js';
import { normalizeName } from './nameAliases.js';

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
 * @param {Array} data - 過濾後的資料
 * @returns {Array} 統計資料 [{ name: string, count: number }]
 */
export function calculateParticipantCount(data) {
  const stats = {};
  
  data.forEach(record => {
    record.participants.forEach(name => {
      // 使用別名映射將人名轉換為標準名稱
      const normalizedName = normalizeName(name);
      if (!stats[normalizedName]) {
        stats[normalizedName] = 0;
      }
      stats[normalizedName]++;
    });
  });
  
  return Object.entries(stats)
    .map(([name, count]) => ({ name, count }))
    .sort((a, b) => b.count - a.count);
}

/**
 * 計算按人名的活動參與時數（使用別名映射合併）
 * @param {Array} data - 過濾後的資料
 * @returns {Array} 統計資料 [{ name: string, hours: number }]
 */
export function calculateParticipantHours(data) {
  const stats = {};
  
  data.forEach(record => {
    const hours = record.hours || 0;
    record.participants.forEach(name => {
      // 使用別名映射將人名轉換為標準名稱
      const normalizedName = normalizeName(name);
      if (!stats[normalizedName]) {
        stats[normalizedName] = 0;
      }
      stats[normalizedName] += hours;
    });
  });
  
  return Object.entries(stats)
    .map(([name, hours]) => ({ name, hours }))
    .sort((a, b) => b.hours - a.hours);
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
  
  return Object.entries(stats)
    .map(([name, value]) => ({ name, value }))
    .sort((a, b) => b.value - a.value);
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
  
  return Object.entries(stats)
    .map(([name, value]) => ({ name, value }))
    .sort((a, b) => b.value - a.value);
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

