import ExcelJS from 'exceljs';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

/**
 * 解析服勤區欄位中的人名
 * @param {string|number} value - 服勤區欄位的值
 * @returns {string[]} 人名陣列
 */
function parseParticipants(value) {
  if (!value) return [];
  
  let str = String(value).trim();
  if (!str) return [];
  
  // 先處理換行符（可能是 \n 或 \r\n）
  str = str.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
  
  // 使用多種分隔符來分割：換行、頓號、逗號
  let names = [];
  
  // 先按換行符分割
  if (str.includes('\n')) {
    names = str.split('\n').flatMap(line => {
      line = line.trim();
      if (!line) return [];
      // 每行可能還包含頓號或逗號
      if (line.includes('、')) {
        return line.split('、');
      } else if (line.includes(',')) {
        return line.split(',');
      } else if (line.includes('，')) {
        return line.split('，');
      }
      return [line];
    });
  } else {
    // 沒有換行符，直接按分隔符分割
    // 優先順序：頓號 > 中文逗號 > 英文逗號
    if (str.includes('、')) {
      names = str.split('、');
    } else if (str.includes('，')) {
      names = str.split('，');
    } else if (str.includes(',')) {
      names = str.split(',');
    } else {
      // 沒有分隔符，整個作為一個名字
      names = [str];
    }
  }
  
  // 清理每個名字（去除前後空格、移除空字串）
  return names
    .map(name => name.trim())
    .filter(name => name.length > 0);
}

/**
 * 從 Excel 檔案提取所有人名
 */
async function extractAllNames() {
  const excelPath = path.join(__dirname, '手作步道活動助教人力需求表.xlsx');
  
  if (!fs.existsSync(excelPath)) {
    console.error('找不到 Excel 檔案:', excelPath);
    process.exit(1);
  }
  
  const workbook = new ExcelJS.Workbook();
  const buffer = fs.readFileSync(excelPath);
  await workbook.xlsx.load(buffer);
  
  // 尋找名為 "2025" 的 sheet
  const sheet = workbook.getWorksheet('2025');
  if (!sheet) {
    console.error('找不到名為 "2025" 的工作表');
    process.exit(1);
  }
  
  const nameSet = new Set();
  let rowCount = 0;
  
  // 從第 2 行開始讀取（假設第 1 行是標題）
  sheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return; // 跳過標題行
    
    try {
      const participantsValue = row.getCell('Q').value;
      const participants = parseParticipants(participantsValue);
      
      participants.forEach(name => {
        if (name) {
          nameSet.add(name);
        }
      });
      
      rowCount++;
    } catch (error) {
      console.warn(`第 ${rowNumber} 行解析錯誤:`, error.message);
    }
  });
  
  // 轉換為陣列並排序
  const names = Array.from(nameSet).sort();
  
  console.log(`總共處理 ${rowCount} 行資料`);
  console.log(`找到 ${names.length} 個不重複的人名\n`);
  
  // 建立別名映射結構
  const aliasMap = {};
  names.forEach(name => {
    aliasMap[name] = name; // 預設別名就是自己
  });
  
  // 儲存為 JSON 檔案
  const outputPath = path.join(__dirname, 'nameAliases.json');
  fs.writeFileSync(
    outputPath,
    JSON.stringify(aliasMap, null, 2),
    'utf8'
  );
  
  console.log(`人名列表已儲存到: ${outputPath}`);
  console.log('\n人名列表：');
  names.forEach((name, index) => {
    console.log(`${index + 1}. ${name}`);
  });
  
  console.log('\n請編輯 nameAliases.json 檔案，將相同人員的別名設定為相同值。');
  console.log('例如：');
  console.log('  "張三": "張三",');
  console.log('  "張三三": "張三",  // 別名，會合併到張三');
  console.log('  "張三(小張)": "張三",  // 別名，會合併到張三');
}

// 執行提取
extractAllNames().catch(error => {
  console.error('發生錯誤:', error);
  process.exit(1);
});

