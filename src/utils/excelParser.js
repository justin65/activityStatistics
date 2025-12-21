import ExcelJS from 'exceljs';
import { parseDate } from './dateParser.js';

/**
 * 解析服勤區欄位中的人名
 * 支援多種分隔符：換行符、頓號（、）、英文逗號（,）、中文逗號（，）
 * @param {string|number} value - 服勤區欄位的值
 * @returns {string[]} 人名陣列
 */
function parseParticipants(value) {
  if (!value) return [];
  
  let str = String(value).trim();
  if (!str) return [];
  
  // 先處理換行符（可能是 \n 或 \r\n）
  str = str.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
  
  let names = [];
  
  // 使用正則表達式同時分割所有可能的分隔符
  // 分隔符包括：換行符、頓號（、）、中文逗號（，）、英文逗號（,）
  // 使用正則表達式來匹配這些分隔符
  const separators = /[\n、，,]+/;
  
  // 先按所有分隔符分割
  names = str.split(separators);
  
  // 清理每個名字（去除前後空格）
  names = names
    .map(name => name.trim())
    .filter(name => name.length > 0);
  
  return names;
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
      // A: 日期, B: 活動名稱, C: 辦理情形, D: 天數, E: 活動類型, G: 縣市, N: 可登錄時數, Q: 服勤區
      const dateValue = row.getCell('A').value;
      const activityName = row.getCell('B').value;
      const status = row.getCell('C').value;
      const daysValue = row.getCell('D').value;
      const activityType = row.getCell('E').value;
      const city = row.getCell('G').value;
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
      
      // 解析時數
      let hours = 0;
      if (hoursValue !== null && hoursValue !== undefined) {
        const hoursNum = typeof hoursValue === 'number' ? hoursValue : parseFloat(hoursValue);
        if (!isNaN(hoursNum) && hoursNum > 0) {
          hours = hoursNum;
        }
      }
      
      // 解析服勤區人名
      const participants = parseParticipants(participantsValue);
      
      // 建立資料物件
      const record = {
        date,
        activityName: String(activityName || '').trim(),
        status: String(status || '').trim(),
        days,
        activityType: String(activityType || '').trim(),
        city: String(city || '').trim(),
        hours,
        participants,
        rowNumber,
      };
      
      data.push(record);
    } catch (error) {
      console.warn(`第 ${rowNumber} 行解析錯誤:`, error);
    }
  });
  
  return data;
}

