import ExcelJS from 'exceljs';
import { parseDate } from './dateParser.js';

/**
 * 讀取時數登錄表 Excel 檔案並解析資料
 * @param {File} file - Excel 檔案
 * @returns {Promise<Array>} 解析後的資料陣列
 */
export async function parseHourLogFile(file) {
  const workbook = new ExcelJS.Workbook();
  const buffer = await file.arrayBuffer();
  await workbook.xlsx.load(buffer);
  
  // 尋找名為 "表單回應" 的 sheet
  const sheet = workbook.getWorksheet('表單回應');
  if (!sheet) {
    throw new Error('找不到名為 "表單回應" 的工作表');
  }
  
  const data = [];
  
  // 從第 2 行開始讀取（假設第 1 行是標題）
  sheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return; // 跳過標題行
    
    try {
      // 讀取各欄位
      // B: 志工姓名, C: 日期, D: 參與內容, F: 參與時數
      const nameValue = row.getCell('B').value;
      const dateValue = row.getCell('C').value;
      const contentValue = row.getCell('D').value;
      const hoursValue = row.getCell('F').value;
      
      // 跳過空行
      if (!nameValue && !dateValue && !contentValue && !hoursValue) {
        return;
      }
      
      // 解析日期
      const date = parseDate(dateValue);
      if (!date) {
        return; // 跳過無法解析日期的行
      }
      
      // 解析時數
      let hours = 0;
      if (hoursValue !== null && hoursValue !== undefined) {
        const hoursNum = typeof hoursValue === 'number' ? hoursValue : parseFloat(hoursValue);
        if (!isNaN(hoursNum) && hoursNum > 0) {
          hours = hoursNum;
        }
      }
      
      // 建立資料物件
      const record = {
        name: String(nameValue || '').trim(),
        date,
        content: String(contentValue || '').trim(),
        hours,
        rowNumber,
      };
      
      // 只添加有效的記錄（至少要有姓名和時數）
      if (record.name && record.hours > 0) {
        data.push(record);
      }
    } catch (error) {
      console.warn(`第 ${rowNumber} 行解析錯誤:`, error);
    }
  });
  
  return data;
}
