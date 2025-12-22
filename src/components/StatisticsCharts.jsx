import React, { useMemo } from 'react';
import { Paper, Typography, Box, Grid } from '@mui/material';
import {
  BarChart,
  Bar,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  Legend,
  ResponsiveContainer,
  PieChart,
  Pie,
  Cell,
} from 'recharts';
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
} from '../utils/dataProcessor.js';

export default function StatisticsCharts({ data, hourLogData }) {
  if ((!data || data.length === 0) && (!hourLogData || !hourLogData.data || hourLogData.data.length === 0)) {
    return (
      <Typography variant="body1" color="text.secondary" align="center" sx={{ py: 4 }}>
        請先上傳 Excel 檔案以顯示統計圖表
      </Typography>
    );
  }

  // 使用 useMemo 緩存計算結果，只在 data 改變時重新計算
  const statistics = useMemo(() => {
    if (!data || data.length === 0) {
      return {
        monthlyActivityTypeCount: [],
        monthlyActivityTypeDays: [],
        monthlyCityCount: [],
        monthlyCityDays: [],
        monthlyRegionCount: [],
        monthlyRegionDays: [],
        activityTypeTotalCount: [],
        activityTypeTotalDays: [],
        cityTotalCount: [],
        cityTotalDays: [],
        regionTotalCount: [],
        regionTotalDays: [],
        participantCount: [],
        participantHours: [],
      };
    }

    return {
      monthlyActivityTypeCount: calculateMonthlyActivityTypeCount(data),
      monthlyActivityTypeDays: calculateMonthlyActivityTypeDays(data),
      monthlyCityCount: calculateMonthlyCityCount(data),
      monthlyCityDays: calculateMonthlyCityDays(data),
      monthlyRegionCount: calculateMonthlyRegionCount(data),
      monthlyRegionDays: calculateMonthlyRegionDays(data),
      activityTypeTotalCount: calculateActivityTypeTotalCount(data),
      activityTypeTotalDays: calculateActivityTypeTotalDays(data),
      cityTotalCount: calculateCityTotalCount(data),
      cityTotalDays: calculateCityTotalDays(data),
      regionTotalCount: calculateRegionTotalCount(data),
      regionTotalDays: calculateRegionTotalDays(data),
      participantCount: calculateParticipantCount(data),
      participantHours: calculateParticipantHours(data),
    };
  }, [data]);

  // 解構統計資料
  const {
    monthlyActivityTypeCount,
    monthlyActivityTypeDays,
    monthlyCityCount,
    monthlyCityDays,
    monthlyRegionCount,
    monthlyRegionDays,
    activityTypeTotalCount,
    activityTypeTotalDays,
    cityTotalCount,
    cityTotalDays,
    regionTotalCount,
    regionTotalDays,
    participantCount,
    participantHours,
  } = statistics;

  // 取得所有活動類型和縣市（用於圖表）
  const activityTypes = getActivityTypes(monthlyActivityTypeCount);
  const citiesRaw = getCities(monthlyCityCount);

  // 縣市區域分類（北中南東）
  const cityRegions = {
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

  // 區域內排序順序（如果同區域有多個縣市）
  const cityOrder = {
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

  // 對縣市進行排序：先按區域（北中南東），再按區域內順序，未分類放在最後
  const cities = [...citiesRaw].sort((a, b) => {
    // 明確處理「未分類」，讓它始終排在最後
    if (a === '未分類' && b !== '未分類') return 1;
    if (a !== '未分類' && b === '未分類') return -1;
    if (a === '未分類' && b === '未分類') return 0;
    
    const regionA = cityRegions[a] || 99; // 未定義的縣市放在最後（但未分類已經處理過了）
    const regionB = cityRegions[b] || 99;
    
    if (regionA !== regionB) {
      return regionA - regionB;
    }
    
    // 同區域內按順序排序
    const orderA = cityOrder[a] || 999;
    const orderB = cityOrder[b] || 999;
    if (orderA !== orderB) {
      return orderA - orderB;
    }
    
    // 如果都是未定義的縣市（且不是未分類），按名稱排序
    return a.localeCompare(b, 'zh-TW');
  });

  // 為 recharts 準備資料格式
  const formatStackedData = (statsData, categories) => {
    return statsData.map(item => {
      const result = { month: item.month };
      categories.forEach(cat => {
        result[cat] = item[cat] || 0;
      });
      return result;
    });
  };

  // 地區順序（北中南東）
  const regions = ['北部', '中部', '南部', '東部'];
  
  const activityTypeCountData = formatStackedData(monthlyActivityTypeCount, activityTypes);
  const activityTypeDaysData = formatStackedData(monthlyActivityTypeDays, activityTypes);
  const cityCountData = formatStackedData(monthlyCityCount, cities);
  const cityDaysData = formatStackedData(monthlyCityDays, cities);
  
  // 地區順序（包含未分類）
  const regionsWithUnclassified = [...regions];
  // 檢查是否有未分類的資料
  if (monthlyRegionCount.some(item => item['未分類'])) {
    if (!regionsWithUnclassified.includes('未分類')) {
      regionsWithUnclassified.push('未分類');
    }
  }
  
  const regionCountData = formatStackedData(monthlyRegionCount, regionsWithUnclassified);
  const regionDaysData = formatStackedData(monthlyRegionDays, regionsWithUnclassified);

  // 顏色配置
  const getColorPalette = (count) => {
    const colors = [
      '#8884d8', '#82ca9d', '#ffc658', '#ff7300', '#0088fe',
      '#00c49f', '#ffbb28', '#ff8042', '#8dd1e1', '#d084d0',
      '#ffb347', '#87ceeb', '#da70d6', '#20b2aa', '#ff6347',
      '#4682b4', '#32cd32', '#ff1493', '#00ced1', '#ff8c00',
      '#9370db', '#3cb371', '#dc143c', '#1e90ff', '#ffd700',
      '#ff69b4', '#00fa9a', '#ba55d3', '#f0e68c', '#98fb98',
    ];
    return colors.slice(0, count);
  };

  // 為活動類型建立顏色映射，確保每個活動類型都有唯一顏色
  const getActivityTypeColorMap = (typeList) => {
    const colorPalette = getColorPalette(typeList.length);
    const colorMap = {};
    typeList.forEach((type, index) => {
      colorMap[type] = colorPalette[index];
    });
    return colorMap;
  };

  // 為縣市建立顏色映射，確保每個縣市都有唯一顏色
  const getCityColorMap = (cityList) => {
    const colorPalette = getColorPalette(cityList.length);
    const colorMap = {};
    cityList.forEach((city, index) => {
      colorMap[city] = colorPalette[index];
    });
    return colorMap;
  };

  const activityTypeColorMap = getActivityTypeColorMap(activityTypes);
  const cityColorMap = getCityColorMap(cities);

  // 計算圓餅圖數據的總和
  const calculateTotal = (data) => {
    return data.reduce((sum, item) => sum + (item.value || 0), 0);
  };

  // 自訂 Tooltip
  const CustomTooltip = ({ active, payload, label }) => {
    if (active && payload && payload.length) {
      return (
        <Paper sx={{ p: 1.5, bgcolor: 'rgba(255, 255, 255, 0.95)' }}>
          <Typography variant="body2" sx={{ fontWeight: 'bold', mb: 0.5 }}>
            {label}
          </Typography>
          {payload.map((entry, index) => (
            <Typography
              key={index}
              variant="body2"
              sx={{ color: entry.color }}
            >
              {`${entry.name}: ${entry.value}`}
            </Typography>
          ))}
        </Paper>
      );
    }
    return null;
  };

  return (
    <Grid container spacing={3}>
      {/* 圖表 1: 月份活動次數（Stack：活動類型） */}
      <Grid item xs={12}>
        <Paper elevation={2} sx={{ p: 3 }}>
          <Typography variant="h6" gutterBottom>
            1. 各月份活動次數統計（依活動類型）
          </Typography>
          <Box sx={{ width: '100%', height: 400, mt: 2 }}>
            <ResponsiveContainer>
              <BarChart data={activityTypeCountData} margin={{ top: 20, right: 30, left: 20, bottom: 60 }}>
                <CartesianGrid strokeDasharray="3 3" />
                <XAxis
                  dataKey="month"
                  angle={-45}
                  textAnchor="end"
                  height={100}
                  interval={0}
                />
                <YAxis />
                <Tooltip content={<CustomTooltip />} />
                <Legend />
                {activityTypes.map((type) => (
                  <Bar
                    key={type}
                    dataKey={type}
                    stackId="a"
                    fill={activityTypeColorMap[type]}
                  />
                ))}
              </BarChart>
            </ResponsiveContainer>
          </Box>
        </Paper>
      </Grid>

      {/* 圖表 2: 月份活動天數（Stack：活動類型） */}
      <Grid item xs={12}>
        <Paper elevation={2} sx={{ p: 3 }}>
          <Typography variant="h6" gutterBottom>
            2. 各月份活動天數統計（依活動類型）
          </Typography>
          <Box sx={{ width: '100%', height: 400, mt: 2 }}>
            <ResponsiveContainer>
              <BarChart data={activityTypeDaysData} margin={{ top: 20, right: 30, left: 20, bottom: 60 }}>
                <CartesianGrid strokeDasharray="3 3" />
                <XAxis
                  dataKey="month"
                  angle={-45}
                  textAnchor="end"
                  height={100}
                  interval={0}
                />
                <YAxis />
                <Tooltip content={<CustomTooltip />} />
                <Legend />
                {activityTypes.map((type) => (
                  <Bar
                    key={type}
                    dataKey={type}
                    stackId="a"
                    fill={activityTypeColorMap[type]}
                  />
                ))}
              </BarChart>
            </ResponsiveContainer>
          </Box>
        </Paper>
      </Grid>

      {/* 圓餅圖：依活動類型統計次數 */}
      <Grid item xs={12} md={6}>
        <Paper elevation={2} sx={{ p: 3 }}>
          <Typography variant="h6" gutterBottom>
            3. 依活動類型統計次數（總次數：{calculateTotal(activityTypeTotalCount)}）
          </Typography>
          <Box sx={{ width: '100%', height: 400, mt: 2 }}>
            <ResponsiveContainer>
              <PieChart>
                <Pie
                  data={activityTypeTotalCount}
                  cx="50%"
                  cy="50%"
                  labelLine={false}
                  label={({ name, percent }) => `${name}: ${(percent * 100).toFixed(0)}%`}
                  outerRadius={120}
                  fill="#8884d8"
                  dataKey="value"
                >
                  {activityTypeTotalCount.map((entry) => {
                    const color = activityTypeColorMap[entry.name] || '#cccccc';
                    return <Cell key={`cell-${entry.name}`} fill={color} />;
                  })}
                </Pie>
                <Tooltip />
                <Legend />
              </PieChart>
            </ResponsiveContainer>
          </Box>
        </Paper>
      </Grid>

      {/* 圓餅圖：依活動類型統計天數 */}
      <Grid item xs={12} md={6}>
        <Paper elevation={2} sx={{ p: 3 }}>
          <Typography variant="h6" gutterBottom>
            4. 依活動類型統計天數（總天數：{calculateTotal(activityTypeTotalDays)}）
          </Typography>
          <Box sx={{ width: '100%', height: 400, mt: 2 }}>
            <ResponsiveContainer>
              <PieChart>
                <Pie
                  data={activityTypeTotalDays}
                  cx="50%"
                  cy="50%"
                  labelLine={false}
                  label={({ name, percent }) => `${name}: ${(percent * 100).toFixed(0)}%`}
                  outerRadius={120}
                  fill="#8884d8"
                  dataKey="value"
                >
                  {activityTypeTotalDays.map((entry) => {
                    const color = activityTypeColorMap[entry.name] || '#cccccc';
                    return <Cell key={`cell-${entry.name}`} fill={color} />;
                  })}
                </Pie>
                <Tooltip />
                <Legend />
              </PieChart>
            </ResponsiveContainer>
          </Box>
        </Paper>
      </Grid>

      {/* 圖表 5: 月份活動次數（Stack：縣市） */}
      <Grid item xs={12}>
        <Paper elevation={2} sx={{ p: 3 }}>
          <Typography variant="h6" gutterBottom>
            5. 各月份活動次數統計（依縣市）
          </Typography>
          <Box sx={{ width: '100%', height: 400, mt: 2 }}>
            <ResponsiveContainer>
              <BarChart data={cityCountData} margin={{ top: 20, right: 30, left: 20, bottom: 60 }}>
                <CartesianGrid strokeDasharray="3 3" />
                <XAxis
                  dataKey="month"
                  angle={-45}
                  textAnchor="end"
                  height={100}
                  interval={0}
                />
                <YAxis />
                <Tooltip content={<CustomTooltip />} />
                <Legend />
                {cities.map((city) => (
                  <Bar
                    key={city}
                    dataKey={city}
                    stackId="a"
                    fill={cityColorMap[city]}
                  />
                ))}
              </BarChart>
            </ResponsiveContainer>
          </Box>
        </Paper>
      </Grid>

      {/* 圖表 6: 月份活動天數（Stack：縣市） */}
      <Grid item xs={12}>
        <Paper elevation={2} sx={{ p: 3 }}>
          <Typography variant="h6" gutterBottom>
            6. 各月份活動天數統計（依縣市）
          </Typography>
          <Box sx={{ width: '100%', height: 400, mt: 2 }}>
            <ResponsiveContainer>
              <BarChart data={cityDaysData} margin={{ top: 20, right: 30, left: 20, bottom: 60 }}>
                <CartesianGrid strokeDasharray="3 3" />
                <XAxis
                  dataKey="month"
                  angle={-45}
                  textAnchor="end"
                  height={100}
                  interval={0}
                />
                <YAxis />
                <Tooltip content={<CustomTooltip />} />
                <Legend />
                {cities.map((city) => (
                  <Bar
                    key={city}
                    dataKey={city}
                    stackId="a"
                    fill={cityColorMap[city]}
                  />
                ))}
              </BarChart>
            </ResponsiveContainer>
          </Box>
        </Paper>
      </Grid>

      {/* 圓餅圖：依縣市統計次數 */}
      <Grid item xs={12} md={6}>
        <Paper elevation={2} sx={{ p: 3 }}>
          <Typography variant="h6" gutterBottom>
            7. 依縣市統計次數（總次數：{calculateTotal(cityTotalCount)}）
          </Typography>
          <Box sx={{ width: '100%', height: 400, mt: 2 }}>
            <ResponsiveContainer>
              <PieChart>
                <Pie
                  data={cityTotalCount}
                  cx="50%"
                  cy="50%"
                  labelLine={false}
                  label={({ name, percent }) => `${name}: ${(percent * 100).toFixed(0)}%`}
                  outerRadius={120}
                  fill="#8884d8"
                  dataKey="value"
                >
                  {cityTotalCount.map((entry) => {
                    // 使用與柱狀圖相同的顏色映射
                    const color = cityColorMap[entry.name] || '#cccccc';
                    return <Cell key={`cell-${entry.name}`} fill={color} />;
                  })}
                </Pie>
                <Tooltip />
                <Legend />
              </PieChart>
            </ResponsiveContainer>
          </Box>
        </Paper>
      </Grid>

      {/* 圓餅圖：依縣市統計天數 */}
      <Grid item xs={12} md={6}>
        <Paper elevation={2} sx={{ p: 3 }}>
          <Typography variant="h6" gutterBottom>
            8. 依縣市統計天數（總天數：{calculateTotal(cityTotalDays)}）
          </Typography>
          <Box sx={{ width: '100%', height: 400, mt: 2 }}>
            <ResponsiveContainer>
              <PieChart>
                <Pie
                  data={cityTotalDays}
                  cx="50%"
                  cy="50%"
                  labelLine={false}
                  label={({ name, percent }) => `${name}: ${(percent * 100).toFixed(0)}%`}
                  outerRadius={120}
                  fill="#8884d8"
                  dataKey="value"
                >
                  {cityTotalDays.map((entry) => {
                    // 使用與柱狀圖相同的顏色映射
                    const color = cityColorMap[entry.name] || '#cccccc';
                    return <Cell key={`cell-${entry.name}`} fill={color} />;
                  })}
                </Pie>
                <Tooltip />
                <Legend />
              </PieChart>
            </ResponsiveContainer>
          </Box>
        </Paper>
      </Grid>

      {/* 圖表 9: 月份活動次數（Stack：地區） */}
      <Grid item xs={12}>
        <Paper elevation={2} sx={{ p: 3 }}>
          <Typography variant="h6" gutterBottom>
            9. 各月份活動次數統計（依地區）
          </Typography>
          <Box sx={{ width: '100%', height: 400, mt: 2 }}>
            <ResponsiveContainer>
              <BarChart data={regionCountData} margin={{ top: 20, right: 30, left: 20, bottom: 60 }}>
                <CartesianGrid strokeDasharray="3 3" />
                <XAxis
                  dataKey="month"
                  angle={-45}
                  textAnchor="end"
                  height={100}
                  interval={0}
                />
                <YAxis />
                <Tooltip content={<CustomTooltip />} />
                <Legend />
                {regionsWithUnclassified.map((region, index) => {
                  const regionColors = ['#8884d8', '#82ca9d', '#ffc658', '#ff7300', '#cccccc'];
                  const colorIndex = region === '未分類' ? 4 : index;
                  return (
                    <Bar
                      key={region}
                      dataKey={region}
                      stackId="a"
                      fill={regionColors[colorIndex]}
                    />
                  );
                })}
              </BarChart>
            </ResponsiveContainer>
          </Box>
        </Paper>
      </Grid>

      {/* 圖表 10: 月份活動天數（Stack：地區） */}
      <Grid item xs={12}>
        <Paper elevation={2} sx={{ p: 3 }}>
          <Typography variant="h6" gutterBottom>
            10. 各月份活動天數統計（依地區）
          </Typography>
          <Box sx={{ width: '100%', height: 400, mt: 2 }}>
            <ResponsiveContainer>
              <BarChart data={regionDaysData} margin={{ top: 20, right: 30, left: 20, bottom: 60 }}>
                <CartesianGrid strokeDasharray="3 3" />
                <XAxis
                  dataKey="month"
                  angle={-45}
                  textAnchor="end"
                  height={100}
                  interval={0}
                />
                <YAxis />
                <Tooltip content={<CustomTooltip />} />
                <Legend />
                {regionsWithUnclassified.map((region, index) => {
                  const regionColors = ['#8884d8', '#82ca9d', '#ffc658', '#ff7300', '#cccccc'];
                  const colorIndex = region === '未分類' ? 4 : index;
                  return (
                    <Bar
                      key={region}
                      dataKey={region}
                      stackId="a"
                      fill={regionColors[colorIndex]}
                    />
                  );
                })}
              </BarChart>
            </ResponsiveContainer>
          </Box>
        </Paper>
      </Grid>

      {/* 圓餅圖：依地區統計次數 */}
      <Grid item xs={12} md={6}>
        <Paper elevation={2} sx={{ p: 3 }}>
          <Typography variant="h6" gutterBottom>
            11. 依地區統計次數（總次數：{calculateTotal(regionTotalCount)}）
          </Typography>
          <Box sx={{ width: '100%', height: 400, mt: 2 }}>
            <ResponsiveContainer>
              <PieChart>
                <Pie
                  data={regionTotalCount}
                  cx="50%"
                  cy="50%"
                  labelLine={false}
                  label={({ name, percent }) => `${name}: ${(percent * 100).toFixed(0)}%`}
                  outerRadius={120}
                  fill="#8884d8"
                  dataKey="value"
                >
                  {regionTotalCount.map((entry) => {
                    const regionColors = ['#8884d8', '#82ca9d', '#ffc658', '#ff7300', '#cccccc'];
                    const colorIndex = entry.name === '未分類' ? 4 : regionsWithUnclassified.indexOf(entry.name);
                    return <Cell key={`cell-${entry.name}`} fill={regionColors[colorIndex] || '#cccccc'} />;
                  })}
                </Pie>
                <Tooltip />
                <Legend />
              </PieChart>
            </ResponsiveContainer>
          </Box>
        </Paper>
      </Grid>

      {/* 圓餅圖：依地區統計天數 */}
      <Grid item xs={12} md={6}>
        <Paper elevation={2} sx={{ p: 3 }}>
          <Typography variant="h6" gutterBottom>
            12. 依地區統計天數（總天數：{calculateTotal(regionTotalDays)}）
          </Typography>
          <Box sx={{ width: '100%', height: 400, mt: 2 }}>
            <ResponsiveContainer>
              <PieChart>
                <Pie
                  data={regionTotalDays}
                  cx="50%"
                  cy="50%"
                  labelLine={false}
                  label={({ name, percent }) => `${name}: ${(percent * 100).toFixed(0)}%`}
                  outerRadius={120}
                  fill="#8884d8"
                  dataKey="value"
                >
                  {regionTotalDays.map((entry) => {
                    const regionColors = ['#8884d8', '#82ca9d', '#ffc658', '#ff7300', '#cccccc'];
                    const colorIndex = entry.name === '未分類' ? 4 : regionsWithUnclassified.indexOf(entry.name);
                    return <Cell key={`cell-${entry.name}`} fill={regionColors[colorIndex] || '#cccccc'} />;
                  })}
                </Pie>
                <Tooltip />
                <Legend />
              </PieChart>
            </ResponsiveContainer>
          </Box>
        </Paper>
      </Grid>

      {/* 圖表 13: 人名參與活動次數 */}
      <Grid item xs={12}>
        <Paper elevation={2} sx={{ p: 3 }}>
          <Typography variant="h6" gutterBottom>
            13. 參與人員活動次數統計
          </Typography>
          <Box sx={{ width: '100%', height: 400, mt: 2, mb: 4 }}>
            <ResponsiveContainer>
              <BarChart
                data={participantCount.slice(0, Math.ceil(participantCount.length / 2))}
                margin={{ top: 20, right: 30, left: 20, bottom: 100 }}
              >
                <CartesianGrid strokeDasharray="3 3" />
                <XAxis
                  dataKey="name"
                  angle={-45}
                  textAnchor="end"
                  height={120}
                  interval={0}
                />
                <YAxis />
                <Tooltip />
                <Bar dataKey="count" fill="#8884d8" />
              </BarChart>
            </ResponsiveContainer>
          </Box>
          <Box sx={{ width: '100%', height: 400, mt: 2 }}>
            <ResponsiveContainer>
              <BarChart
                data={participantCount.slice(Math.ceil(participantCount.length / 2))}
                margin={{ top: 20, right: 30, left: 20, bottom: 100 }}
              >
                <CartesianGrid strokeDasharray="3 3" />
                <XAxis
                  dataKey="name"
                  angle={-45}
                  textAnchor="end"
                  height={120}
                  interval={0}
                />
                <YAxis />
                <Tooltip />
                <Bar dataKey="count" fill="#8884d8" />
              </BarChart>
            </ResponsiveContainer>
          </Box>
        </Paper>
      </Grid>

      {/* 圖表 14: 人名參與活動時數 */}
      <Grid item xs={12}>
        <Paper elevation={2} sx={{ p: 3 }}>
          <Typography variant="h6" gutterBottom>
            14. 參與人員活動時數統計
          </Typography>
          <Box sx={{ width: '100%', height: 400, mt: 2, mb: 4 }}>
            <ResponsiveContainer>
              <BarChart
                data={participantHours.slice(0, Math.ceil(participantHours.length / 2))}
                margin={{ top: 20, right: 30, left: 20, bottom: 100 }}
              >
                <CartesianGrid strokeDasharray="3 3" />
                <XAxis
                  dataKey="name"
                  angle={-45}
                  textAnchor="end"
                  height={120}
                  interval={0}
                />
                <YAxis />
                <Tooltip />
                <Bar dataKey="hours" fill="#82ca9d" />
              </BarChart>
            </ResponsiveContainer>
          </Box>
          <Box sx={{ width: '100%', height: 400, mt: 2 }}>
            <ResponsiveContainer>
              <BarChart
                data={participantHours.slice(Math.ceil(participantHours.length / 2))}
                margin={{ top: 20, right: 30, left: 20, bottom: 100 }}
              >
                <CartesianGrid strokeDasharray="3 3" />
                <XAxis
                  dataKey="name"
                  angle={-45}
                  textAnchor="end"
                  height={120}
                  interval={0}
                />
                <YAxis />
                <Tooltip />
                <Bar dataKey="hours" fill="#82ca9d" />
              </BarChart>
            </ResponsiveContainer>
          </Box>
        </Paper>
      </Grid>

      {/* 圖表 15: 志工參與時數統計（依參與內容） */}
      {hourLogData && hourLogData.data && hourLogData.data.length > 0 && (() => {
        // 為參與內容建立顏色映射，確保每個內容類型都有唯一顏色
        const getContentColorMap = (contentList) => {
          const colorPalette = getColorPalette(contentList.length);
          const colorMap = {};
          contentList.forEach((content, index) => {
            colorMap[content] = colorPalette[index];
          });
          return colorMap;
        };
        
        const contentColorMap = getContentColorMap(hourLogData.contentTypes);
        
        // 計算每個志工的總時數並添加到數據中
        const dataWithTotal = hourLogData.data.map(item => {
          const total = Object.entries(item).reduce((sum, [key, val]) => {
            if (key === 'name') return sum;
            return typeof val === 'number' ? sum + val : sum;
          }, 0);
          return { ...item, total };
        });
        
        // 按總時數排序（降序）
        dataWithTotal.sort((a, b) => b.total - a.total);
        
        // 根據中位數分成兩部分
        const midIndex = Math.ceil(dataWithTotal.length / 2);
        const topHalf = dataWithTotal.slice(0, midIndex);
        const bottomHalf = dataWithTotal.slice(midIndex);
        
        // 自定義 Tooltip，顯示總時數在人名上面，過濾掉時數為0的項目
        const HourLogTooltip = ({ active, payload, label }) => {
          if (active && payload && payload.length) {
            // 計算總時數
            const total = payload.reduce((sum, entry) => {
              return sum + (typeof entry.value === 'number' ? entry.value : 0);
            }, 0);
            
            // 過濾掉時數為0的項目
            const filteredPayload = payload.filter(entry => 
              typeof entry.value === 'number' && entry.value > 0
            );
            
            return (
              <Paper sx={{ p: 1.5, bgcolor: 'rgba(255, 255, 255, 0.95)' }}>
                <Typography variant="body2" sx={{ fontWeight: 'bold', mb: 0.5, color: '#1976d2' }}>
                  總時數: {total} 小時
                </Typography>
                <Typography variant="body2" sx={{ fontWeight: 'bold', mb: 0.5 }}>
                  {label}
                </Typography>
                {filteredPayload.map((entry, index) => (
                  <Typography
                    key={index}
                    variant="body2"
                    sx={{ color: entry.color }}
                  >
                    {`${entry.name}: ${entry.value} 小時`}
                  </Typography>
                ))}
              </Paper>
            );
          }
          return null;
        };
        
        return (
          <Grid item xs={12}>
            <Paper elevation={2} sx={{ p: 3 }}>
              <Typography variant="h6" gutterBottom>
                15. 志工參與時數統計（依參與內容）
              </Typography>
              {/* 上半部分 */}
              <Box sx={{ width: '100%', height: 400, mt: 2, mb: 4 }}>
                <ResponsiveContainer>
                  <BarChart
                    data={topHalf}
                    margin={{ top: 20, right: 30, left: 20, bottom: 100 }}
                  >
                    <CartesianGrid strokeDasharray="3 3" />
                    <XAxis
                      dataKey="name"
                      angle={-45}
                      textAnchor="end"
                      height={120}
                      interval={0}
                    />
                    <YAxis />
                    <Tooltip content={<HourLogTooltip />} />
                    <Legend />
                    {hourLogData.contentTypes.map((content) => (
                      <Bar
                        key={content}
                        dataKey={content}
                        stackId="a"
                        fill={contentColorMap[content]}
                      />
                    ))}
                  </BarChart>
                </ResponsiveContainer>
              </Box>
              {/* 下半部分 */}
              {bottomHalf.length > 0 && (
                <Box sx={{ width: '100%', height: 400, mt: 2 }}>
                  <ResponsiveContainer>
                    <BarChart
                      data={bottomHalf}
                      margin={{ top: 20, right: 30, left: 20, bottom: 100 }}
                    >
                      <CartesianGrid strokeDasharray="3 3" />
                      <XAxis
                        dataKey="name"
                        angle={-45}
                        textAnchor="end"
                        height={120}
                        interval={0}
                      />
                      <YAxis />
                      <Tooltip content={<HourLogTooltip />} />
                      <Legend />
                      {hourLogData.contentTypes.map((content) => (
                        <Bar
                          key={content}
                          dataKey={content}
                          stackId="a"
                          fill={contentColorMap[content]}
                        />
                      ))}
                    </BarChart>
                  </ResponsiveContainer>
                </Box>
              )}
            </Paper>
          </Grid>
        );
      })()}
    </Grid>
  );
}

