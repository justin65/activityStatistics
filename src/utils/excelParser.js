import ExcelJS from 'exceljs';
import { parseDate } from './dateParser.js';

/**
 * 解析日期範圍，計算天數
 * 支援格式：
 * - "11/29-30" -> 2天（11/29和11/30）
 * - "11/29" -> 1天
 * - "6/30-7/1" -> 2天（跨月）
 * @param {string} dateStr - 日期字串
 * @param {Date} activityDate - 活動日期（用於確定年份和處理只有日期沒有月份的情況）
 * @returns {number} 天數
 */
function parseDateRange(dateStr, activityDate = null) {
  if (!dateStr) return 0;
  
  const trimmed = dateStr.trim();
  const defaultYear = 2025;
  
  // 從活動日期獲取月份（如果有的話）
  let defaultMonth = null;
  if (activityDate && activityDate instanceof Date && !isNaN(activityDate.getTime())) {
    defaultMonth = activityDate.getMonth() + 1;
  }
  
  // 處理日期範圍格式
  if (trimmed.includes('-')) {
    const parts = trimmed.split('-');
    if (parts.length === 2) {
      const startPart = parts[0].trim();
      const endPart = parts[1].trim();
      
      let startMonth, startDay, endMonth, endDay;
      
      // 解析開始日期
      if (startPart.includes('/')) {
        const startMatch = startPart.match(/^(\d{1,2})\/(\d{1,2})$/);
        if (!startMatch) return 0;
        startMonth = parseInt(startMatch[1], 10);
        startDay = parseInt(startMatch[2], 10);
      } else {
        if (defaultMonth === null) return 0;
        startMonth = defaultMonth;
        const day = parseInt(startPart, 10);
        if (isNaN(day)) return 0;
        startDay = day;
      }
      
      // 解析結束日期
      if (endPart.includes('/')) {
        const endMatch = endPart.match(/^(\d{1,2})\/(\d{1,2})$/);
        if (!endMatch) return 0;
        endMonth = parseInt(endMatch[1], 10);
        endDay = parseInt(endMatch[2], 10);
      } else {
        if (startPart.includes('/')) {
          endMonth = startMonth;
        } else {
          endMonth = defaultMonth || startMonth;
        }
        const day = parseInt(endPart, 10);
        if (isNaN(day)) return 0;
        endDay = day;
      }
      
      // 計算天數
      const startDate = new Date(defaultYear, startMonth - 1, startDay);
      const endDate = new Date(defaultYear, endMonth - 1, endDay);
      
      if (startDate.getMonth() !== startMonth - 1 || startDate.getDate() !== startDay ||
          endDate.getMonth() !== endMonth - 1 || endDate.getDate() !== endDay) {
        return 0;
      }
      
      const diffTime = endDate.getTime() - startDate.getTime();
      const diffDays = Math.floor(diffTime / (1000 * 60 * 60 * 24));
      return diffDays + 1; // 包含開始和結束日期
    }
  } else {
    // 單一日期格式
    const singleMatch = trimmed.match(/^(\d{1,2})\/(\d{1,2})$/);
    if (!singleMatch) return 0;
    
    const month = parseInt(singleMatch[1], 10);
    const day = parseInt(singleMatch[2], 10);
    
    if (month < 1 || month > 12 || day < 1 || day > 31) return 0;
    
    const date = new Date(defaultYear, month - 1, day);
    if (date.getMonth() !== month - 1 || date.getDate() !== day) return 0;
    
    return 1;
  }
}

/**
 * 解析服勤區欄位中的人名
 * 支援多種分隔符：換行符、頓號（、）、英文逗號（,）、中文逗號（，）
 * 支援日期前綴格式：
 * - "11/29-30：盈瑩、藝婷" -> 盈瑩和藝婷各算2天
 * - "11/29：武治、岱華" -> 武治和岱華各算1天
 * @param {string|number} value - 服勤區欄位的值
 * @param {Date} activityDate - 活動日期（用於解析日期範圍）
 * @returns {Array<{name: string, days: number}>} 人名和天數陣列，days為0表示沒有日期信息
 */
function parseParticipants(value, activityDate = null) {
  if (!value) return [];
  
  let str = String(value).trim();
  if (!str) return [];
  
  // 先處理換行符（可能是 \n 或 \r\n）
  str = str.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
  
  const result = [];
  
  // 按換行符分割
  const lines = str.split('\n');
  
  lines.forEach(line => {
    line = line.trim();
    if (!line) return;
    
    // 檢查是否有日期前綴格式：日期：或日期範圍：
    // 支援中文冒號（：）和英文冒號（:）
    const datePrefixMatch = line.match(/^(\d{1,2}\/\d{1,2}(?:-\d{1,2}(?:\/\d{1,2})?)?)[：:]\s*(.+)$/);
    
    if (datePrefixMatch) {
      // 有日期前綴
      const dateStr = datePrefixMatch[1];
      const namesStr = datePrefixMatch[2];
      
      // 解析日期範圍，計算天數
      const days = parseDateRange(dateStr, activityDate);
      
      // 解析人名（支援頓號、中文逗號、英文逗號）
      const nameSeparators = /[、，,]+/;
      const names = namesStr
        .split(nameSeparators)
        .map(name => name.trim())
        .filter(name => name.length > 0);
      
      // 為每個人名添加天數信息
      names.forEach(name => {
        result.push({ name, days });
      });
    } else {
      // 沒有日期前綴，使用原有邏輯
      const nameSeparators = /[、，,]+/;
      const names = line
        .split(nameSeparators)
        .map(name => name.trim())
        .filter(name => name.length > 0);
      
      names.forEach(name => {
        result.push({ name, days: 0 }); // days為0表示沒有日期信息
      });
    }
  });
  
  return result;
}

/**
 * 讀取 Excel 檔案並解析資料
 * @param {File} file - Excel 檔案
 * @returns {Promise<Array>} 解析後的資料陣列
 */
export async function parseExcelFile(file) {
  const workbook = new ExcelJS.Workbook();
  const buffer = await file.arrayBuffer();
  await workbook.xlsx.load(buffer);
  
  // 尋找名為 "2025" 的 sheet
  const sheet = workbook.getWorksheet('2025');
  if (!sheet) {
    throw new Error('找不到名為 "2025" 的工作表');
  }
  
  const data = [];
  
  // 從第 2 行開始讀取（假設第 1 行是標題）
  sheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return; // 跳過標題行
    
    try {
      // 讀取各欄位
      // A: 日期, B: 活動名稱, C: 辦理情形, D: 天數, E: 活動類型, G: 縣市, L: 志工人數, N: 可登錄時數, Q: 服勤區
      const dateValue = row.getCell('A').value;
      const activityName = row.getCell('B').value;
      const status = row.getCell('C').value;
      const daysValue = row.getCell('D').value;
      const activityType = row.getCell('E').value;
      const city = row.getCell('G').value;
      const volunteerCountValue = row.getCell('L').value;
      const hoursValue = row.getCell('N').value;
      const participantsValue = row.getCell('Q').value;
      
      // 解析日期
      const date = parseDate(dateValue);
      if (!date) {
        console.warn(`第 ${rowNumber} 行：無法解析日期`, dateValue);
        return; // 跳過無法解析日期的行
      }
      
      // 解析天數
      let days = 0;
      if (daysValue !== null && daysValue !== undefined) {
        const daysNum = typeof daysValue === 'number' ? daysValue : parseFloat(daysValue);
        if (!isNaN(daysNum) && daysNum > 0) {
          days = daysNum;
        }
      }
      
      // 解析志工人數
      let volunteerCount = 0;
      if (volunteerCountValue !== null && volunteerCountValue !== undefined) {
        const countNum = typeof volunteerCountValue === 'number' ? volunteerCountValue : parseFloat(volunteerCountValue);
        if (!isNaN(countNum) && countNum > 0) {
          volunteerCount = countNum;
        }
      }
      
      // 解析時數
      let hours = 0;
      if (hoursValue !== null && hoursValue !== undefined) {
        const hoursNum = typeof hoursValue === 'number' ? hoursValue : parseFloat(hoursValue);
        if (!isNaN(hoursNum) && hoursNum > 0) {
          hours = hoursNum;
        }
      }
      
      // 解析服勤區人名（傳入活動日期以處理日期範圍）
      const participants = parseParticipants(participantsValue, date);
      
      // 建立資料物件
      const record = {
        date,
        activityName: String(activityName || '').trim(),
        status: String(status || '').trim(),
        days,
        activityType: String(activityType || '').trim(),
        city: String(city || '').trim(),
        volunteerCount,
        hours,
        participants, // 現在是 { name: string, days: number }[] 格式
        rowNumber,
      };
      
      data.push(record);
    } catch (error) {
      console.warn(`第 ${rowNumber} 行解析錯誤:`, error);
    }
  });
  
  return data;
}

