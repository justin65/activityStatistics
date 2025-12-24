import ExcelJS from 'exceljs';
import {
  calculateMonthlyActivityTypeCount,
  calculateMonthlyActivityTypeDays,
  calculateMonthlyCityCount,
  calculateMonthlyCityDays,
  calculateMonthlyRegionCount,
  calculateMonthlyRegionDays,
  calculateActivityTypeTotalCount,
  calculateActivityTypeTotalDays,
  calculateCityTotalCount,
  calculateCityTotalDays,
  calculateRegionTotalCount,
  calculateRegionTotalDays,
  calculateParticipantCount,
  calculateParticipantHours,
  getActivityTypes,
  getCities,
  calculateMonthlyVolunteerCountByType,
  calculateCityVolunteerCountByType,
  calculateRegionVolunteerCountByType,
  calculateVolunteerCountByActivityType,
  calculateVolunteerCountByRegion,
  calculateVolunteerCountByCity,
} from './dataProcessor.js';

function safeNumber(value) {
  const n = typeof value === 'number' ? value : Number(value);
  return Number.isFinite(n) ? n : 0;
}

function sortCitiesLikeCharts(citiesRaw) {
  // 與 StatisticsCharts.jsx 的排序規則一致：北中南東、區域內固定順序、未分類最後
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

  return [...citiesRaw].sort((a, b) => {
    if (a === '未分類' && b !== '未分類') return 1;
    if (a !== '未分類' && b === '未分類') return -1;
    if (a === '未分類' && b === '未分類') return 0;

    const regionA = cityRegionsForSort[a] || 99;
    const regionB = cityRegionsForSort[b] || 99;
    if (regionA !== regionB) return regionA - regionB;

    const orderA = cityOrderForSort[a] || 999;
    const orderB = cityOrderForSort[b] || 999;
    if (orderA !== orderB) return orderA - orderB;

    return a.localeCompare(b, 'zh-TW');
  });
}

function autoFitColumns(worksheet) {
  worksheet.columns.forEach((col) => {
    let maxLength = 10;
    col.eachCell({ includeEmpty: true }, (cell) => {
      const cellValue = cell.value;
      const text =
        cellValue === null || cellValue === undefined
          ? ''
          : typeof cellValue === 'string'
            ? cellValue
            : typeof cellValue === 'number'
              ? String(cellValue)
              : typeof cellValue === 'object' && cellValue.richText
                ? cellValue.richText.map(t => t.text).join('')
                : String(cellValue);
      maxLength = Math.max(maxLength, text.length);
    });
    col.width = Math.min(60, maxLength + 2);
  });
}

function addStackedBarSheet(workbook, sheetName, xHeader, xKey, rows, seriesKeys) {
  const ws = workbook.addWorksheet(String(sheetName));
  const header = [xHeader, ...seriesKeys, '總和'];
  ws.addRow(header);
  ws.getRow(1).font = { bold: true };

  rows.forEach((r) => {
    const values = seriesKeys.map((k) => safeNumber(r[k]));
    const total = values.reduce((sum, v) => sum + v, 0);
    ws.addRow([r[xKey], ...values, total]);
  });

  ws.views = [{ state: 'frozen', ySplit: 1 }];
  autoFitColumns(ws);
}

function addPieSheet(workbook, sheetName, categoryHeader, data) {
  const ws = workbook.addWorksheet(String(sheetName));
  ws.addRow([categoryHeader, '數量', '百分比']);
  ws.getRow(1).font = { bold: true };

  const total = data.reduce((sum, item) => sum + safeNumber(item.value), 0);
  data.forEach((item) => {
    const value = safeNumber(item.value);
    const pct = total > 0 ? value / total : 0;
    ws.addRow([item.name, value, pct]);
  });

  // 總和列
  ws.addRow(['總和', total, total > 0 ? 1 : 0]);

  // 百分比欄位格式
  ws.getColumn(3).numFmt = '0.00%';
  ws.views = [{ state: 'frozen', ySplit: 1 }];
  autoFitColumns(ws);
}

function normalizeStackedRows(rows, xKey, seriesKeys) {
  return rows.map((r) => {
    const out = { ...r };
    seriesKeys.forEach((k) => {
      if (out[k] === undefined) out[k] = 0;
    });
    // 確保 xKey 存在（避免 undefined）
    if (out[xKey] === undefined) out[xKey] = '';
    return out;
  });
}

function getRegionOrderKey(region) {
  const order = { '北部': 1, '中部': 2, '南部': 3, '東部': 4, '未分類': 5 };
  return order[region] || 99;
}

function getRegionSeriesKeysFromRows(rows) {
  const keys = new Set();
  rows.forEach((r) => {
    Object.keys(r).forEach((k) => {
      if (k !== 'month' && k !== 'year' && k !== 'monthNum') keys.add(k);
    });
  });
  const ordered = ['北部', '中部', '南部', '東部', '未分類'];
  const rest = Array.from(keys).filter(k => !ordered.includes(k)).sort();
  return [...ordered.filter(k => keys.has(k)), ...rest];
}

function getSeriesKeysFromRows(rows, excludeKeys) {
  const keys = new Set();
  rows.forEach((r) => {
    Object.keys(r).forEach((k) => {
      if (!excludeKeys.includes(k)) keys.add(k);
    });
  });
  return Array.from(keys).sort((a, b) => a.localeCompare(b, 'zh-TW'));
}

function formatDateForFilename(d) {
  const pad = (n) => String(n).padStart(2, '0');
  return `${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}_${pad(d.getHours())}${pad(d.getMinutes())}`;
}

function downloadBufferAsFile(buffer, filename) {
  const blob = new Blob([buffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

/**
 * 將目前畫面會出現的圖表資料匯出成 Excel
 * - 每個圖表一個 sheet，sheet name 為圖表編號
 * - 柱狀圖：Header = stack/系列分類，最後加「總和」欄；每個 x 軸項目一個 row
 * - 圓餅圖：Header =「數量」「百分比」；每個分類一個 row；最後一列「總和」
 */
export async function exportCurrentChartsToExcel({ data, hourLogData }) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'activityStatistics';
  workbook.created = new Date();

  const hasManpower = Array.isArray(data) && data.length > 0;
  const hasHourLog = !!(hourLogData && hourLogData.data && hourLogData.data.length > 0);
  let participantHoursForDifference = null;

  if (!hasManpower && !hasHourLog) {
    throw new Error('目前沒有可匯出的圖表資料，請先上傳檔案');
  }

  if (hasManpower) {
    const monthlyActivityTypeCount = calculateMonthlyActivityTypeCount(data);
    const monthlyActivityTypeDays = calculateMonthlyActivityTypeDays(data);
    const monthlyCityCount = calculateMonthlyCityCount(data);
    const monthlyCityDays = calculateMonthlyCityDays(data);
    const monthlyRegionCount = calculateMonthlyRegionCount(data);
    const monthlyRegionDays = calculateMonthlyRegionDays(data);

    const activityTypeTotalCount = calculateActivityTypeTotalCount(data);
    const activityTypeTotalDays = calculateActivityTypeTotalDays(data);
    const cityTotalCount = calculateCityTotalCount(data);
    const cityTotalDays = calculateCityTotalDays(data);
    const regionTotalCount = calculateRegionTotalCount(data);
    const regionTotalDays = calculateRegionTotalDays(data);

    const participantCount = calculateParticipantCount(data);
    const participantHours = calculateParticipantHours(data);
    participantHoursForDifference = participantHours;

    const monthlyVolunteerCountByType = calculateMonthlyVolunteerCountByType(data);
    const cityVolunteerCountByType = calculateCityVolunteerCountByType(data);
    const regionVolunteerCountByType = calculateRegionVolunteerCountByType(data);
    const volunteerCountByActivityType = calculateVolunteerCountByActivityType(data);
    const volunteerCountByRegion = calculateVolunteerCountByRegion(data);
    const volunteerCountByCity = calculateVolunteerCountByCity(data);

    const activityTypes = getActivityTypes(monthlyActivityTypeCount);
    const cities = sortCitiesLikeCharts(getCities(monthlyCityCount));

    // 1-2
    addStackedBarSheet(
      workbook,
      1,
      '月份',
      'month',
      normalizeStackedRows(monthlyActivityTypeCount, 'month', activityTypes),
      activityTypes
    );
    addStackedBarSheet(
      workbook,
      2,
      '月份',
      'month',
      normalizeStackedRows(monthlyActivityTypeDays, 'month', activityTypes),
      activityTypes
    );

    // 3-4
    addPieSheet(workbook, 3, '活動類型', activityTypeTotalCount);
    addPieSheet(workbook, 4, '活動類型', activityTypeTotalDays);

    // 5-6（依縣市）
    addStackedBarSheet(
      workbook,
      5,
      '月份',
      'month',
      normalizeStackedRows(monthlyCityCount, 'month', cities),
      cities
    );
    addStackedBarSheet(
      workbook,
      6,
      '月份',
      'month',
      normalizeStackedRows(monthlyCityDays, 'month', cities),
      cities
    );

    // 7-8
    addPieSheet(workbook, 7, '縣市', cityTotalCount);
    addPieSheet(workbook, 8, '縣市', cityTotalDays);

    // 9-10（依地區）
    {
      const regionSeries = getRegionSeriesKeysFromRows(monthlyRegionCount);
      const regionSeriesDays = getRegionSeriesKeysFromRows(monthlyRegionDays);
      addStackedBarSheet(
        workbook,
        9,
        '月份',
        'month',
        normalizeStackedRows(monthlyRegionCount, 'month', regionSeries),
        regionSeries
      );
      addStackedBarSheet(
        workbook,
        10,
        '月份',
        'month',
        normalizeStackedRows(monthlyRegionDays, 'month', regionSeriesDays),
        regionSeriesDays
      );
    }

    // 11-12
    addPieSheet(workbook, 11, '地區', regionTotalCount);
    addPieSheet(workbook, 12, '地區', regionTotalDays);

    // 13-15（志工人數堆疊）
    addStackedBarSheet(
      workbook,
      13,
      '月份',
      'month',
      normalizeStackedRows(monthlyVolunteerCountByType, 'month', activityTypes),
      activityTypes
    );

    // 14：依縣市（補齊 activityTypes）
    {
      const normalized = normalizeStackedRows(cityVolunteerCountByType, 'city', activityTypes);
      // 盡量讓列順序跟圖表一致（依 getCities 的排序）
      const map = new Map(normalized.map(r => [r.city, r]));
      const orderedRows = cities.map(city => map.get(city) || { city, ...Object.fromEntries(activityTypes.map(t => [t, 0])) });
      addStackedBarSheet(workbook, 14, '縣市', 'city', orderedRows, activityTypes);
    }

    // 15：依地區（補齊 activityTypes + 排序）
    {
      const normalized = normalizeStackedRows(regionVolunteerCountByType, 'region', activityTypes)
        .slice()
        .sort((a, b) => getRegionOrderKey(a.region) - getRegionOrderKey(b.region));
      addStackedBarSheet(workbook, 15, '地區', 'region', normalized, activityTypes);
    }

    // 16-18（圓餅）
    addPieSheet(workbook, 16, '活動類型', volunteerCountByActivityType);
    addPieSheet(workbook, 17, '地區', volunteerCountByRegion);
    addPieSheet(workbook, 18, '縣市', volunteerCountByCity);

    // 19-20（人名堆疊）
    {
      // legend 順序使用 activityTypes（與圖表 19 一致）
      addStackedBarSheet(workbook, 19, '人名', 'name', normalizeStackedRows(participantCount, 'name', activityTypes), activityTypes);
    }
    {
      // legend 順序使用 activityTypes（與圖表 20 一致）
      addStackedBarSheet(workbook, 20, '人名', 'name', normalizeStackedRows(participantHours, 'name', activityTypes), activityTypes);
    }
  }

  if (hasHourLog) {
    // 21：回報時數統計（堆疊）
    const contentTypes = Array.isArray(hourLogData.contentTypes) ? hourLogData.contentTypes : getSeriesKeysFromRows(hourLogData.data, ['name']);
    addStackedBarSheet(
      workbook,
      21,
      '人名',
      'name',
      normalizeStackedRows(hourLogData.data, 'name', contentTypes),
      contentTypes
    );

    // 22：回流訓練（單一系列）
    if (hourLogData?.retraining?.data && hourLogData.retraining.data.length > 0) {
      addStackedBarSheet(
        workbook,
        22,
        '人名',
        'name',
        hourLogData.retraining.data.map(r => ({ name: r.name, '回流訓練': safeNumber(r.hours) })),
        ['回流訓練']
      );
    }
  }

  // 23（需要同時有人力需求表 + 時數登錄表）—放在最後，確保 sheet 依編號順序（...21,22,23）
  if (hasManpower && hasHourLog && Array.isArray(participantHoursForDifference)) {
    const differenceRows = [];
    const participantHoursMap = {};
    participantHoursForDifference.forEach(item => {
      participantHoursMap[item.name] = safeNumber(item['手作']);
    });
    const hourLogMap = {};
    hourLogData.data.forEach(item => {
      hourLogMap[item.name] = safeNumber(item['步道實作帶領']);
    });
    const allNames = new Set([
      ...participantHoursForDifference.map(item => item.name),
      ...hourLogData.data.map(item => item.name),
    ]);
    allNames.forEach(name => {
      const handcraftHours = participantHoursMap[name] || 0;
      const trailLeadingHours = hourLogMap[name] || 0;
      const difference = handcraftHours - trailLeadingHours;
      if (difference !== 0) {
        differenceRows.push({ name, 差值: difference });
      }
    });
    differenceRows.sort((a, b) => safeNumber(b.差值) - safeNumber(a.差值));
    if (differenceRows.length > 0) {
      addStackedBarSheet(workbook, 23, '人名', 'name', differenceRows, ['差值']);
    }
  }

  const filename = `圖表資料匯出_${formatDateForFilename(new Date())}.xlsx`;
  const buffer = await workbook.xlsx.writeBuffer();
  downloadBufferAsFile(buffer, filename);
}


