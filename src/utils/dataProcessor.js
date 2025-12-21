import { getMonth, getYear, formatMonth } from './dateParser.js';

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
 * 計算按人名的活動參與次數
 * @param {Array} data - 過濾後的資料
 * @returns {Array} 統計資料 [{ name: string, count: number }]
 */
export function calculateParticipantCount(data) {
  const stats = {};
  
  data.forEach(record => {
    record.participants.forEach(name => {
      if (!stats[name]) {
        stats[name] = 0;
      }
      stats[name]++;
    });
  });
  
  return Object.entries(stats)
    .map(([name, count]) => ({ name, count }))
    .sort((a, b) => b.count - a.count);
}

/**
 * 計算按人名的活動參與時數
 * @param {Array} data - 過濾後的資料
 * @returns {Array} 統計資料 [{ name: string, hours: number }]
 */
export function calculateParticipantHours(data) {
  const stats = {};
  
  data.forEach(record => {
    const hours = record.hours || 0;
    record.participants.forEach(name => {
      if (!stats[name]) {
        stats[name] = 0;
      }
      stats[name] += hours;
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

