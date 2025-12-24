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
  calculateMonthlyVolunteerCountByType,
  calculateCityVolunteerCountByType,
  calculateRegionVolunteerCountByType,
  calculateVolunteerCountByActivityType,
  calculateVolunteerCountByRegion,
  calculateVolunteerCountByCity,
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
        monthlyVolunteerCountByType: [],
        cityVolunteerCountByType: [],
        regionVolunteerCountByType: [],
        volunteerCountByActivityType: [],
        volunteerCountByRegion: [],
        volunteerCountByCity: [],
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
      monthlyVolunteerCountByType: calculateMonthlyVolunteerCountByType(data),
      cityVolunteerCountByType: calculateCityVolunteerCountByType(data),
      regionVolunteerCountByType: calculateRegionVolunteerCountByType(data),
      volunteerCountByActivityType: calculateVolunteerCountByActivityType(data),
      volunteerCountByRegion: calculateVolunteerCountByRegion(data),
      volunteerCountByCity: calculateVolunteerCountByCity(data),
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
    monthlyVolunteerCountByType,
    cityVolunteerCountByType,
    regionVolunteerCountByType,
    volunteerCountByActivityType,
    volunteerCountByRegion,
    volunteerCountByCity,
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
      {/* 圖表 1-20: 需要人力需求表 */}
      {data && data.length > 0 && (
        <>
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

      {/* 志工相關圖表 */}
      {data && data.length > 0 && (() => {
        // 準備數據格式
        const monthlyVolunteerData = formatStackedData(monthlyVolunteerCountByType, activityTypes);
        
        // 縣市志工數據，需要按 cities 陣列的順序排序（與圖表 5 一致）
        // 先建立一個映射，將城市數據轉換為對象
        const cityVolunteerMap = {};
        cityVolunteerCountByType.forEach(item => {
          cityVolunteerMap[item.city] = item;
        });
        
        // 按照 cities 陣列的順序來構建數據
        const cityVolunteerData = cities.map(city => {
          const item = cityVolunteerMap[city] || { city };
          const result = { city: item.city };
          activityTypes.forEach(type => {
            result[type] = item[type] || 0;
          });
          return result;
        });
        
        // 地區志工數據
        const regionVolunteerData = regionVolunteerCountByType.map(item => {
          const result = { region: item.region };
          activityTypes.forEach(type => {
            result[type] = item[type] || 0;
          });
          return result;
        }).sort((a, b) => {
          const order = { '北部': 1, '中部': 2, '南部': 3, '東部': 4, '未分類': 5 };
          return (order[a.region] || 99) - (order[b.region] || 99);
        });
        
        return (
          <>
            {/* 圖表 13: 志工人數統計（依月份） */}
            <Grid item xs={12}>
              <Paper elevation={2} sx={{ p: 3 }}>
                <Typography variant="h6" gutterBottom>
                  13. 志工人數統計（依月份）
                </Typography>
                <Box sx={{ width: '100%', height: 400, mt: 2 }}>
                  <ResponsiveContainer>
                    <BarChart data={monthlyVolunteerData} margin={{ top: 20, right: 30, left: 20, bottom: 60 }}>
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

            {/* 圖表 14: 志工人數統計（依縣市） */}
            <Grid item xs={12}>
              <Paper elevation={2} sx={{ p: 3 }}>
                <Typography variant="h6" gutterBottom>
                  14. 志工人數統計（依縣市）
                </Typography>
                <Box sx={{ width: '100%', height: 400, mt: 2 }}>
                  <ResponsiveContainer>
                    <BarChart data={cityVolunteerData} margin={{ top: 20, right: 30, left: 20, bottom: 60 }}>
                      <CartesianGrid strokeDasharray="3 3" />
                      <XAxis
                        dataKey="city"
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

            {/* 圖表 15: 志工人數統計（依地區） */}
            <Grid item xs={12}>
              <Paper elevation={2} sx={{ p: 3 }}>
                <Typography variant="h6" gutterBottom>
                  15. 志工人數統計（依地區）
                </Typography>
                <Box sx={{ width: '100%', height: 400, mt: 2 }}>
                  <ResponsiveContainer>
                    <BarChart data={regionVolunteerData} margin={{ top: 20, right: 30, left: 20, bottom: 60 }}>
                      <CartesianGrid strokeDasharray="3 3" />
                      <XAxis
                        dataKey="region"
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

            {/* 圓餅圖：依活動類型、依地區、依縣市（同一行） */}
            <Grid item xs={12} md={4}>
              <Paper elevation={2} sx={{ p: 3 }}>
                <Typography variant="h6" gutterBottom>
                  16. 志工人數統計（依活動類型）（總人數：{calculateTotal(volunteerCountByActivityType)}）
                </Typography>
                <Box sx={{ width: '100%', height: 400, mt: 2 }}>
                  <ResponsiveContainer>
                    <PieChart>
                      <Pie
                        data={volunteerCountByActivityType}
                        cx="50%"
                        cy="50%"
                        labelLine={false}
                        label={({ name, percent }) => `${name}: ${(percent * 100).toFixed(0)}%`}
                        outerRadius={120}
                        fill="#8884d8"
                        dataKey="value"
                      >
                        {volunteerCountByActivityType.map((entry) => {
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

            <Grid item xs={12} md={4}>
              <Paper elevation={2} sx={{ p: 3 }}>
                <Typography variant="h6" gutterBottom>
                  17. 志工人數統計（依地區）（總人數：{calculateTotal(volunteerCountByRegion)}）
                </Typography>
                <Box sx={{ width: '100%', height: 400, mt: 2 }}>
                  <ResponsiveContainer>
                    <PieChart>
                      <Pie
                        data={volunteerCountByRegion}
                        cx="50%"
                        cy="50%"
                        labelLine={false}
                        label={({ name, percent }) => `${name}: ${(percent * 100).toFixed(0)}%`}
                        outerRadius={120}
                        fill="#8884d8"
                        dataKey="value"
                      >
                        {volunteerCountByRegion.map((entry) => {
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

            <Grid item xs={12} md={4}>
              <Paper elevation={2} sx={{ p: 3 }}>
                <Typography variant="h6" gutterBottom>
                  18. 志工人數統計（依縣市）（總人數：{calculateTotal(volunteerCountByCity)}）
                </Typography>
                <Box sx={{ width: '100%', height: 400, mt: 2 }}>
                  <ResponsiveContainer>
                    <PieChart>
                      <Pie
                        data={volunteerCountByCity}
                        cx="50%"
                        cy="50%"
                        labelLine={false}
                        label={({ name, percent }) => `${name}: ${(percent * 100).toFixed(0)}%`}
                        outerRadius={120}
                        fill="#8884d8"
                        dataKey="value"
                      >
                        {volunteerCountByCity.map((entry) => {
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
          </>
        );
      })()}

      {/* 圖表 19: 人名參與活動次數 */}
      <Grid item xs={12}>
        <Paper elevation={2} sx={{ p: 3 }}>
          <Typography variant="h6" gutterBottom>
            19. 出勤活動次數統計（共 {participantCount.length} 位）
          </Typography>
          {(() => {
            // 計算每個人的總值（所有活動類型的總和）
            const calculateTotal = (item) => {
              return activityTypes.reduce((sum, type) => sum + (item[type] || 0), 0);
            };
            
            // 計算所有數據的總值
            const allTotals = participantCount.map(item => calculateTotal(item)).sort((a, b) => b - a);
            
            // 計算分位數
            const q75 = allTotals[Math.floor(allTotals.length * 0.25)] || 0; // 75分位數（前25%）
            const q50 = allTotals[Math.floor(allTotals.length * 0.5)] || 0; // 中位數
            const q25 = allTotals[Math.floor(allTotals.length * 0.75)] || 0; // 25分位數
            
            // 判斷是否需要分組：如果最大值是75分位數的2倍以上，則分組
            const maxValue = allTotals[0] || 0;
            const needsGrouping = maxValue > 0 && maxValue > q75 * 2;
            
            if (needsGrouping) {
              // 分組顯示：高值組（前25%）和其餘組
              const highValueGroup = participantCount.filter(item => calculateTotal(item) >= q75);
              const otherGroup = participantCount.filter(item => calculateTotal(item) < q75);
              
              const highMax = Math.max(...highValueGroup.map(calculateTotal), 0);
              const otherMax = Math.max(...otherGroup.map(calculateTotal), 0);
              
              const highDomain = [0, Math.ceil(highMax * 1.1)];
              const otherDomain = [0, Math.ceil(otherMax * 1.1)];
              
              return (
                <>
                  {highValueGroup.length > 0 && (
                    <>
                      <Typography variant="subtitle1" sx={{ mt: 2, mb: 1, fontWeight: 'bold', color: 'primary.main' }}>
                        高值組（前 {highValueGroup.length} 位）
                      </Typography>
                      {highValueGroup.length <= 30 ? (
                        // 人數不超過30，只顯示一個圖表
                        <Box sx={{ width: '100%', height: 400, mt: 2, mb: 4 }}>
                          <ResponsiveContainer>
                            <BarChart
                              data={highValueGroup}
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
                              <YAxis domain={highDomain} />
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
                      ) : (
                        // 人數超過30，分成上下兩部分
                        <>
                          {(() => {
                            const highTop = highValueGroup.slice(0, Math.ceil(highValueGroup.length / 2));
                            const highBottom = highValueGroup.slice(Math.ceil(highValueGroup.length / 2));
                            return (
                              <>
                                {highTop.length > 0 && (
                                  <Box sx={{ width: '100%', height: 400, mt: 2, mb: 4 }}>
                                    <ResponsiveContainer>
                                      <BarChart
                                        data={highTop}
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
                                        <YAxis domain={highDomain} />
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
                                )}
                                {highBottom.length > 0 && (
                                  <Box sx={{ width: '100%', height: 400, mt: 2, mb: 4 }}>
                                    <ResponsiveContainer>
                                      <BarChart
                                        data={highBottom}
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
                                        <YAxis domain={highDomain} />
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
                                )}
                              </>
                            );
                          })()}
                        </>
                      )}
                    </>
                  )}
                  {otherGroup.length > 0 && (
                    <>
                      <Typography variant="subtitle1" sx={{ mt: 2, mb: 1, fontWeight: 'bold', color: 'secondary.main' }}>
                        其餘組（{otherGroup.length} 位）
                      </Typography>
                      {otherGroup.length <= 30 ? (
                        // 人數不超過30，只顯示一個圖表
                        <Box sx={{ width: '100%', height: 400, mt: 2 }}>
                          <ResponsiveContainer>
                            <BarChart
                              data={otherGroup}
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
                              <YAxis domain={otherDomain} />
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
                      ) : (
                        // 人數超過30，分成上下兩部分
                        <>
                          {(() => {
                            const otherTop = otherGroup.slice(0, Math.ceil(otherGroup.length / 2));
                            const otherBottom = otherGroup.slice(Math.ceil(otherGroup.length / 2));
                            return (
                              <>
                                {otherTop.length > 0 && (
                                  <Box sx={{ width: '100%', height: 400, mt: 2, mb: 4 }}>
                                    <ResponsiveContainer>
                                      <BarChart
                                        data={otherTop}
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
                                        <YAxis domain={otherDomain} />
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
                                )}
                                {otherBottom.length > 0 && (
                                  <Box sx={{ width: '100%', height: 400, mt: 2 }}>
                                    <ResponsiveContainer>
                                      <BarChart
                                        data={otherBottom}
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
                                        <YAxis domain={otherDomain} />
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
                                )}
                              </>
                            );
                          })()}
                        </>
                      )}
                    </>
                  )}
                </>
              );
            } else {
              // 不需要分組，使用原來的顯示方式
              const topHalf = participantCount.slice(0, Math.ceil(participantCount.length / 2));
              const bottomHalf = participantCount.slice(Math.ceil(participantCount.length / 2));
              const maxValue = Math.max(...allTotals, 0);
              const yAxisDomain = [0, Math.ceil(maxValue * 1.1)];
              
              return (
                <>
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
                        <YAxis domain={yAxisDomain} />
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
                        <YAxis domain={yAxisDomain} />
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
                </>
              );
            }
          })()}
        </Paper>
      </Grid>

      {/* 圖表 20: 人名參與活動時數 */}
      <Grid item xs={12}>
        <Paper elevation={2} sx={{ p: 3 }}>
          <Typography variant="h6" gutterBottom>
            20. 出勤活動時數統計（共 {participantHours.length} 位）
          </Typography>
          {(() => {
            // 計算每個人的總值（所有活動類型的總和）
            const calculateTotal = (item) => {
              return activityTypes.reduce((sum, type) => sum + (item[type] || 0), 0);
            };
            
            // 計算所有數據的總值
            const allTotals = participantHours.map(item => calculateTotal(item)).sort((a, b) => b - a);
            
            // 計算分位數
            const q75 = allTotals[Math.floor(allTotals.length * 0.25)] || 0; // 75分位數（前25%）
            const maxValue = allTotals[0] || 0;
            const needsGrouping = maxValue > 0 && maxValue > q75 * 2;
            
            if (needsGrouping) {
              // 分組顯示：高值組（前25%）和其餘組
              const highValueGroup = participantHours.filter(item => calculateTotal(item) >= q75);
              const otherGroup = participantHours.filter(item => calculateTotal(item) < q75);
              
              const highMax = Math.max(...highValueGroup.map(calculateTotal), 0);
              const otherMax = Math.max(...otherGroup.map(calculateTotal), 0);
              
              const highDomain = [0, Math.ceil(highMax * 1.1)];
              const otherDomain = [0, Math.ceil(otherMax * 1.1)];
              
              return (
                <>
                  {highValueGroup.length > 0 && (
                    <>
                      <Typography variant="subtitle1" sx={{ mt: 2, mb: 1, fontWeight: 'bold', color: 'primary.main' }}>
                        高值組（前 {highValueGroup.length} 位）
                      </Typography>
                      {highValueGroup.length <= 30 ? (
                        // 人數不超過30，只顯示一個圖表
                        <Box sx={{ width: '100%', height: 400, mt: 2, mb: 4 }}>
                          <ResponsiveContainer>
                            <BarChart
                              data={highValueGroup}
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
                              <YAxis domain={highDomain} />
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
                      ) : (
                        // 人數超過30，分成上下兩部分
                        <>
                          {(() => {
                            const highTop = highValueGroup.slice(0, Math.ceil(highValueGroup.length / 2));
                            const highBottom = highValueGroup.slice(Math.ceil(highValueGroup.length / 2));
                            return (
                              <>
                                {highTop.length > 0 && (
                                  <Box sx={{ width: '100%', height: 400, mt: 2, mb: 4 }}>
                                    <ResponsiveContainer>
                                      <BarChart
                                        data={highTop}
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
                                        <YAxis domain={highDomain} />
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
                                )}
                                {highBottom.length > 0 && (
                                  <Box sx={{ width: '100%', height: 400, mt: 2, mb: 4 }}>
                                    <ResponsiveContainer>
                                      <BarChart
                                        data={highBottom}
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
                                        <YAxis domain={highDomain} />
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
                                )}
                              </>
                            );
                          })()}
                        </>
                      )}
                    </>
                  )}
                  {otherGroup.length > 0 && (
                    <>
                      <Typography variant="subtitle1" sx={{ mt: 2, mb: 1, fontWeight: 'bold', color: 'secondary.main' }}>
                        其餘組（{otherGroup.length} 位）
                      </Typography>
                      {otherGroup.length <= 30 ? (
                        // 人數不超過30，只顯示一個圖表
                        <Box sx={{ width: '100%', height: 400, mt: 2 }}>
                          <ResponsiveContainer>
                            <BarChart
                              data={otherGroup}
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
                              <YAxis domain={otherDomain} />
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
                      ) : (
                        // 人數超過30，分成上下兩部分
                        <>
                          {(() => {
                            const otherTop = otherGroup.slice(0, Math.ceil(otherGroup.length / 2));
                            const otherBottom = otherGroup.slice(Math.ceil(otherGroup.length / 2));
                            return (
                              <>
                                {otherTop.length > 0 && (
                                  <Box sx={{ width: '100%', height: 400, mt: 2, mb: 4 }}>
                                    <ResponsiveContainer>
                                      <BarChart
                                        data={otherTop}
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
                                        <YAxis domain={otherDomain} />
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
                                )}
                                {otherBottom.length > 0 && (
                                  <Box sx={{ width: '100%', height: 400, mt: 2 }}>
                                    <ResponsiveContainer>
                                      <BarChart
                                        data={otherBottom}
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
                                        <YAxis domain={otherDomain} />
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
                                )}
                              </>
                            );
                          })()}
                        </>
                      )}
                    </>
                  )}
                </>
              );
            } else {
              // 不需要分組，使用原來的顯示方式
              const topHalf = participantHours.slice(0, Math.ceil(participantHours.length / 2));
              const bottomHalf = participantHours.slice(Math.ceil(participantHours.length / 2));
              const maxValue = Math.max(...allTotals, 0);
              const yAxisDomain = [0, Math.ceil(maxValue * 1.1)];
              
              return (
                <>
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
                        <YAxis domain={yAxisDomain} />
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
                        <YAxis domain={yAxisDomain} />
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
                </>
              );
            }
          })()}
        </Paper>
      </Grid>
        </>
      )}

      {/* 圖表 21: 回報時數統計（需要時數登錄表） */}
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
        
        // 計算分位數
        const allTotals = dataWithTotal.map(item => item.total).sort((a, b) => b - a);
        const q75 = allTotals[Math.floor(allTotals.length * 0.25)] || 0; // 75分位數（前25%）
        const maxValue = allTotals[0] || 0;
        const needsGrouping = maxValue > 0 && maxValue > q75 * 2;
        
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
        
        if (needsGrouping) {
          // 分組顯示：高值組（前25%）和其餘組
          const highValueGroup = dataWithTotal.filter(item => item.total >= q75);
          const otherGroup = dataWithTotal.filter(item => item.total < q75);
          
          const highMax = Math.max(...highValueGroup.map(item => item.total), 0);
          const otherMax = Math.max(...otherGroup.map(item => item.total), 0);
          
          const highDomain = [0, Math.ceil(highMax * 1.1)];
          const otherDomain = [0, Math.ceil(otherMax * 1.1)];
          
          return (
            <Grid item xs={12}>
              <Paper elevation={2} sx={{ p: 3 }}>
                <Typography variant="h6" gutterBottom>
                  21. 回報時數統計（共 {hourLogData.data.length} 位）
                </Typography>
                {highValueGroup.length > 0 && (
                  <>
                    <Typography variant="subtitle1" sx={{ mt: 2, mb: 1, fontWeight: 'bold', color: 'primary.main' }}>
                      高值組（前 {highValueGroup.length} 位）
                    </Typography>
                    {highValueGroup.length <= 30 ? (
                      // 人數不超過30，只顯示一個圖表
                      <Box sx={{ width: '100%', height: 400, mt: 2, mb: 4 }}>
                        <ResponsiveContainer>
                          <BarChart
                            data={highValueGroup}
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
                            <YAxis domain={highDomain} />
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
                    ) : (
                      // 人數超過30，分成上下兩部分
                      <>
                        {(() => {
                          const highTop = highValueGroup.slice(0, Math.ceil(highValueGroup.length / 2));
                          const highBottom = highValueGroup.slice(Math.ceil(highValueGroup.length / 2));
                          return (
                            <>
                              {highTop.length > 0 && (
                                <Box sx={{ width: '100%', height: 400, mt: 2, mb: 4 }}>
                                  <ResponsiveContainer>
                                    <BarChart
                                      data={highTop}
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
                                      <YAxis domain={highDomain} />
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
                              {highBottom.length > 0 && (
                                <Box sx={{ width: '100%', height: 400, mt: 2, mb: 4 }}>
                                  <ResponsiveContainer>
                                    <BarChart
                                      data={highBottom}
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
                                      <YAxis domain={highDomain} />
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
                            </>
                          );
                        })()}
                      </>
                    )}
                  </>
                )}
                {otherGroup.length > 0 && (
                  <>
                    <Typography variant="subtitle1" sx={{ mt: 2, mb: 1, fontWeight: 'bold', color: 'secondary.main' }}>
                      其餘組（{otherGroup.length} 位）
                    </Typography>
                    {otherGroup.length <= 30 ? (
                      // 人數不超過30，只顯示一個圖表
                      <Box sx={{ width: '100%', height: 400, mt: 2 }}>
                        <ResponsiveContainer>
                          <BarChart
                            data={otherGroup}
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
                            <YAxis domain={otherDomain} />
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
                    ) : (
                      // 人數超過30，分成上下兩部分
                      <>
                        {(() => {
                          const otherTop = otherGroup.slice(0, Math.ceil(otherGroup.length / 2));
                          const otherBottom = otherGroup.slice(Math.ceil(otherGroup.length / 2));
                          return (
                            <>
                              {otherTop.length > 0 && (
                                <Box sx={{ width: '100%', height: 400, mt: 2, mb: 4 }}>
                                  <ResponsiveContainer>
                                    <BarChart
                                      data={otherTop}
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
                                      <YAxis domain={otherDomain} />
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
                              {otherBottom.length > 0 && (
                                <Box sx={{ width: '100%', height: 400, mt: 2 }}>
                                  <ResponsiveContainer>
                                    <BarChart
                                      data={otherBottom}
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
                                      <YAxis domain={otherDomain} />
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
                            </>
                          );
                        })()}
                      </>
                    )}
                  </>
                )}
              </Paper>
            </Grid>
          );
        } else {
          // 不需要分組，使用原來的顯示方式
          const midIndex = Math.ceil(dataWithTotal.length / 2);
          const topHalf = dataWithTotal.slice(0, midIndex);
          const bottomHalf = dataWithTotal.slice(midIndex);
          const yAxisDomain = [0, Math.ceil(maxValue * 1.1)];
          
          return (
            <Grid item xs={12}>
              <Paper elevation={2} sx={{ p: 3 }}>
                <Typography variant="h6" gutterBottom>
                  21. 回報時數統計（共 {hourLogData.data.length} 位）
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
                      <YAxis domain={yAxisDomain} />
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
                        <YAxis domain={yAxisDomain} />
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
        }
      })()}

      {/* 圖表 22: 2023-2025 回流訓練時數（需要時數登錄表） */}
      {hourLogData?.retraining?.data && hourLogData.retraining.data.length > 0 && (() => {
        const title = `22. 2023–2025 回流訓練時數（共 ${hourLogData.retraining.data.length} 位）`;

        const yearKeys = (hourLogData?.retraining?.years && hourLogData.retraining.years.length > 0)
          ? hourLogData.retraining.years.map(y => String(y))
          : ['2023', '2024', '2025'];

        // 計算每個人的總時數並排序（降序）
        const dataWithTotal = hourLogData.retraining.data
          .map(item => {
            const total = yearKeys.reduce((sum, y) => sum + (typeof item?.[y] === 'number' ? item[y] : 0), 0);
            return { ...item, total };
          })
          .sort((a, b) => (b.total || 0) - (a.total || 0));

        // 計算分位數以判斷是否需要分組（沿用圖表 21 的策略）
        const allTotals = dataWithTotal.map(item => item.total || 0).sort((a, b) => b - a);
        const q75 = allTotals[Math.floor(allTotals.length * 0.25)] || 0; // 75分位數（前25%）
        const maxValue = allTotals[0] || 0;
        const needsGrouping = maxValue > 0 && maxValue > q75 * 2;

        const RetrainingTooltip = ({ active, payload, label }) => {
          if (active && payload && payload.length) {
            const total = payload.reduce((sum, entry) => sum + (typeof entry.value === 'number' ? entry.value : 0), 0);
            return (
              <Paper sx={{ p: 1.5, bgcolor: 'rgba(255, 255, 255, 0.95)' }}>
                <Typography variant="body2" sx={{ fontWeight: 'bold', mb: 0.5, color: '#1976d2' }}>
                  總時數: {total} 小時
                </Typography>
                <Typography variant="body2" sx={{ fontWeight: 'bold' }}>
                  {label}
                </Typography>
                {payload
                  .filter(entry => typeof entry.value === 'number' && entry.value > 0)
                  .map((entry, index) => (
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

        const renderSingleChart = (chartData, yAxisDomain) => (
          <BarChart
            data={chartData}
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
            <YAxis domain={yAxisDomain} />
            <Tooltip content={<RetrainingTooltip />} />
            <Legend />
            {yearKeys.map((y, idx) => (
              <Bar
                key={y}
                dataKey={y}
                name={y}
                stackId="a"
                fill={['#1976d2', '#82ca9d', '#ffc658', '#8884d8', '#ff7300'][idx % 5]}
              />
            ))}
          </BarChart>
        );

        if (needsGrouping) {
          const highValueGroup = dataWithTotal.filter(item => (item.total || 0) >= q75);
          const otherGroup = dataWithTotal.filter(item => (item.total || 0) < q75);

          const calcTotal = (item) => yearKeys.reduce((sum, y) => sum + (typeof item?.[y] === 'number' ? item[y] : 0), 0);
          const highMax = Math.max(...highValueGroup.map(calcTotal), 0);
          const otherMax = Math.max(...otherGroup.map(calcTotal), 0);

          const highDomain = [0, Math.ceil(highMax * 1.1)];
          const otherDomain = [0, Math.ceil(otherMax * 1.1)];

          return (
            <Grid item xs={12}>
              <Paper elevation={2} sx={{ p: 3 }}>
                <Typography variant="h6" gutterBottom>
                  {title}
                </Typography>

                {highValueGroup.length > 0 && (
                  <>
                    <Typography variant="subtitle1" sx={{ mt: 2, mb: 1, fontWeight: 'bold', color: 'primary.main' }}>
                      高值組（前 {highValueGroup.length} 位）
                    </Typography>
                    {highValueGroup.length <= 30 ? (
                      <Box sx={{ width: '100%', height: 400, mt: 2, mb: 4 }}>
                        <ResponsiveContainer>
                          {renderSingleChart(highValueGroup, highDomain)}
                        </ResponsiveContainer>
                      </Box>
                    ) : (
                      <>
                        {(() => {
                          const highTop = highValueGroup.slice(0, Math.ceil(highValueGroup.length / 2));
                          const highBottom = highValueGroup.slice(Math.ceil(highValueGroup.length / 2));
                          return (
                            <>
                              {highTop.length > 0 && (
                                <Box sx={{ width: '100%', height: 400, mt: 2, mb: 4 }}>
                                  <ResponsiveContainer>
                                    {renderSingleChart(highTop, highDomain)}
                                  </ResponsiveContainer>
                                </Box>
                              )}
                              {highBottom.length > 0 && (
                                <Box sx={{ width: '100%', height: 400, mt: 2, mb: 4 }}>
                                  <ResponsiveContainer>
                                    {renderSingleChart(highBottom, highDomain)}
                                  </ResponsiveContainer>
                                </Box>
                              )}
                            </>
                          );
                        })()}
                      </>
                    )}
                  </>
                )}

                {otherGroup.length > 0 && (
                  <>
                    <Typography variant="subtitle1" sx={{ mt: 2, mb: 1, fontWeight: 'bold', color: 'secondary.main' }}>
                      其餘組（{otherGroup.length} 位）
                    </Typography>
                    {otherGroup.length <= 30 ? (
                      <Box sx={{ width: '100%', height: 400, mt: 2 }}>
                        <ResponsiveContainer>
                          {renderSingleChart(otherGroup, otherDomain)}
                        </ResponsiveContainer>
                      </Box>
                    ) : (
                      <>
                        {(() => {
                          const otherTop = otherGroup.slice(0, Math.ceil(otherGroup.length / 2));
                          const otherBottom = otherGroup.slice(Math.ceil(otherGroup.length / 2));
                          return (
                            <>
                              {otherTop.length > 0 && (
                                <Box sx={{ width: '100%', height: 400, mt: 2, mb: 4 }}>
                                  <ResponsiveContainer>
                                    {renderSingleChart(otherTop, otherDomain)}
                                  </ResponsiveContainer>
                                </Box>
                              )}
                              {otherBottom.length > 0 && (
                                <Box sx={{ width: '100%', height: 400, mt: 2 }}>
                                  <ResponsiveContainer>
                                    {renderSingleChart(otherBottom, otherDomain)}
                                  </ResponsiveContainer>
                                </Box>
                              )}
                            </>
                          );
                        })()}
                      </>
                    )}
                  </>
                )}
              </Paper>
            </Grid>
          );
        }

        // 不需要分組，分成上下兩半（沿用圖表 21 的方式）
        const midIndex = Math.ceil(dataWithTotal.length / 2);
        const topHalf = dataWithTotal.slice(0, midIndex);
        const bottomHalf = dataWithTotal.slice(midIndex);
        const yAxisDomain = [0, Math.ceil(maxValue * 1.1)];

        return (
          <Grid item xs={12}>
            <Paper elevation={2} sx={{ p: 3 }}>
              <Typography variant="h6" gutterBottom>
                {title}
              </Typography>
              <Box sx={{ width: '100%', height: 400, mt: 2, mb: 4 }}>
                <ResponsiveContainer>
                  {renderSingleChart(topHalf, yAxisDomain)}
                </ResponsiveContainer>
              </Box>
              {bottomHalf.length > 0 && (
                <Box sx={{ width: '100%', height: 400, mt: 2 }}>
                  <ResponsiveContainer>
                    {renderSingleChart(bottomHalf, yAxisDomain)}
                  </ResponsiveContainer>
                </Box>
              )}
            </Paper>
          </Grid>
        );
      })()}

      {/* 圖表 23: 出勤手作時數 - 回報步道實作帶領時數（需要兩個檔案） */}
      {data && data.length > 0 && hourLogData && hourLogData.data && hourLogData.data.length > 0 && (() => {
        // 計算差值：圖表14中的「手作」時數 - 圖表15中的「步道實作帶領」時數
        const differenceData = [];
        
        // 建立人名到時數的映射
        const participantHoursMap = {};
        participantHours.forEach(item => {
          participantHoursMap[item.name] = item['手作'] || 0;
        });
        
        const hourLogMap = {};
        hourLogData.data.forEach(item => {
          hourLogMap[item.name] = item['步道實作帶領'] || 0;
        });
        
        // 收集所有人名（兩個數據源的聯集）
        const allNames = new Set([
          ...participantHours.map(item => item.name),
          ...hourLogData.data.map(item => item.name)
        ]);
        
        // 計算每個人的差值
        allNames.forEach(name => {
          const handcraftHours = participantHoursMap[name] || 0;
          const trailLeadingHours = hourLogMap[name] || 0;
          const difference = handcraftHours - trailLeadingHours;
          
          // 只顯示有差值的人（差值不為0）
          if (difference !== 0) {
            differenceData.push({
              name,
              difference,
              handcraftHours,
              trailLeadingHours,
            });
          }
        });
        
        // 按差值排序（降序）
        differenceData.sort((a, b) => b.difference - a.difference);
        
        // 如果沒有數據，不顯示圖表
        if (differenceData.length === 0) {
          return null;
        }
        
        const totalCount = differenceData.length; // 保存總人數
        
        // 計算所有差值的絕對值
        const allAbsValues = differenceData.map(item => Math.abs(item.difference)).sort((a, b) => b - a);
        const q75 = allAbsValues[Math.floor(allAbsValues.length * 0.25)] || 0; // 75分位數（前25%）
        const maxAbsValue = allAbsValues[0] || 0;
        const needsGrouping = maxAbsValue > 0 && maxAbsValue > q75 * 2;
        
        // 自定義 Tooltip
        const DifferenceTooltip = ({ active, payload, label }) => {
          if (active && payload && payload.length) {
            const data = payload[0].payload;
            return (
              <Paper sx={{ p: 1.5, bgcolor: 'rgba(255, 255, 255, 0.95)' }}>
                <Typography variant="body2" sx={{ fontWeight: 'bold', mb: 0.5 }}>
                  {label}
                </Typography>
                <Typography variant="body2" sx={{ color: '#1976d2' }}>
                  差值: {data.difference} 小時
                </Typography>
                <Typography variant="body2" sx={{ color: '#8884d8' }}>
                  手作時數: {data.handcraftHours} 小時
                </Typography>
                <Typography variant="body2" sx={{ color: '#82ca9d' }}>
                  步道實作帶領時數: {data.trailLeadingHours} 小時
                </Typography>
              </Paper>
            );
          }
          return null;
        };
        
        if (needsGrouping) {
          // 分組顯示：高值組（前25%）和其餘組
          const highValueGroup = differenceData.filter(item => Math.abs(item.difference) >= q75);
          const otherGroup = differenceData.filter(item => Math.abs(item.difference) < q75);
          
          // 正數和負數的scale分開計算
          const highPositiveValues = highValueGroup.filter(item => item.difference > 0).map(item => item.difference);
          const highNegativeValues = highValueGroup.filter(item => item.difference < 0).map(item => item.difference);
          const highMaxPositive = highPositiveValues.length > 0 ? Math.max(...highPositiveValues) : 0;
          const highMinNegative = highNegativeValues.length > 0 ? Math.min(...highNegativeValues) : 0;
          const highDomain = [Math.ceil(highMinNegative * 1.1), Math.ceil(highMaxPositive * 1.1)];
          
          const otherPositiveValues = otherGroup.filter(item => item.difference > 0).map(item => item.difference);
          const otherNegativeValues = otherGroup.filter(item => item.difference < 0).map(item => item.difference);
          const otherMaxPositive = otherPositiveValues.length > 0 ? Math.max(...otherPositiveValues) : 0;
          const otherMinNegative = otherNegativeValues.length > 0 ? Math.min(...otherNegativeValues) : 0;
          const otherDomain = [Math.ceil(otherMinNegative * 1.1), Math.ceil(otherMaxPositive * 1.1)];
          
          return (
            <Grid item xs={12}>
              <Paper elevation={2} sx={{ p: 3 }}>
                <Typography variant="h6" gutterBottom>
                  23. 出勤手作時數 - 回報步道實作帶領時數（共 {totalCount} 位）
                </Typography>
                {highValueGroup.length > 0 && (
                  <>
                    <Typography variant="subtitle1" sx={{ mt: 2, mb: 1, fontWeight: 'bold', color: 'primary.main' }}>
                      高值組（前 {highValueGroup.length} 位）
                    </Typography>
                    {highValueGroup.length <= 30 ? (
                      // 人數不超過30，只顯示一個圖表
                      <Box sx={{ width: '100%', height: 400, mt: 2, mb: 4 }}>
                        <ResponsiveContainer>
                          <BarChart
                            data={highValueGroup}
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
                            <YAxis domain={highDomain} />
                            <Tooltip content={<DifferenceTooltip />} />
                            <Bar dataKey="difference">
                              {highValueGroup.map((entry, index) => (
                                <Cell 
                                  key={`cell-${index}`} 
                                  fill={entry.difference >= 0 ? '#1976d2' : '#dc004e'} 
                                />
                              ))}
                            </Bar>
                          </BarChart>
                        </ResponsiveContainer>
                      </Box>
                    ) : (
                      // 人數超過30，分成上下兩部分
                      <>
                        {(() => {
                          const highTop = highValueGroup.slice(0, Math.ceil(highValueGroup.length / 2));
                          const highBottom = highValueGroup.slice(Math.ceil(highValueGroup.length / 2));
                          return (
                            <>
                              {highTop.length > 0 && (
                                <Box sx={{ width: '100%', height: 400, mt: 2, mb: 4 }}>
                                  <ResponsiveContainer>
                                    <BarChart
                                      data={highTop}
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
                                      <YAxis domain={highDomain} />
                                      <Tooltip content={<DifferenceTooltip />} />
                                      <Bar dataKey="difference">
                                        {highTop.map((entry, index) => (
                                          <Cell 
                                            key={`cell-${index}`} 
                                            fill={entry.difference >= 0 ? '#1976d2' : '#dc004e'} 
                                          />
                                        ))}
                                      </Bar>
                                    </BarChart>
                                  </ResponsiveContainer>
                                </Box>
                              )}
                              {highBottom.length > 0 && (
                                <Box sx={{ width: '100%', height: 400, mt: 2, mb: 4 }}>
                                  <ResponsiveContainer>
                                    <BarChart
                                      data={highBottom}
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
                                      <YAxis domain={highDomain} />
                                      <Tooltip content={<DifferenceTooltip />} />
                                      <Bar dataKey="difference">
                                        {highBottom.map((entry, index) => (
                                          <Cell 
                                            key={`cell-${index}`} 
                                            fill={entry.difference >= 0 ? '#1976d2' : '#dc004e'} 
                                          />
                                        ))}
                                      </Bar>
                                    </BarChart>
                                  </ResponsiveContainer>
                                </Box>
                              )}
                            </>
                          );
                        })()}
                      </>
                    )}
                  </>
                )}
                {otherGroup.length > 0 && (
                  <>
                    <Typography variant="subtitle1" sx={{ mt: 2, mb: 1, fontWeight: 'bold', color: 'secondary.main' }}>
                      其餘組（{otherGroup.length} 位）
                    </Typography>
                    {otherGroup.length <= 30 ? (
                      // 人數不超過30，只顯示一個圖表
                      <Box sx={{ width: '100%', height: 400, mt: 2 }}>
                        <ResponsiveContainer>
                          <BarChart
                            data={otherGroup}
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
                            <YAxis domain={otherDomain} />
                            <Tooltip content={<DifferenceTooltip />} />
                            <Bar dataKey="difference">
                              {otherGroup.map((entry, index) => (
                                <Cell 
                                  key={`cell-${index}`} 
                                  fill={entry.difference >= 0 ? '#1976d2' : '#dc004e'} 
                                />
                              ))}
                            </Bar>
                          </BarChart>
                        </ResponsiveContainer>
                      </Box>
                    ) : (
                      // 人數超過30，分成上下兩部分
                      <>
                        {(() => {
                          const otherTop = otherGroup.slice(0, Math.ceil(otherGroup.length / 2));
                          const otherBottom = otherGroup.slice(Math.ceil(otherGroup.length / 2));
                          return (
                            <>
                              {otherTop.length > 0 && (
                                <Box sx={{ width: '100%', height: 400, mt: 2, mb: 4 }}>
                                  <ResponsiveContainer>
                                    <BarChart
                                      data={otherTop}
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
                                      <YAxis domain={otherDomain} />
                                      <Tooltip content={<DifferenceTooltip />} />
                                      <Bar dataKey="difference">
                                        {otherTop.map((entry, index) => (
                                          <Cell 
                                            key={`cell-${index}`} 
                                            fill={entry.difference >= 0 ? '#1976d2' : '#dc004e'} 
                                          />
                                        ))}
                                      </Bar>
                                    </BarChart>
                                  </ResponsiveContainer>
                                </Box>
                              )}
                              {otherBottom.length > 0 && (
                                <Box sx={{ width: '100%', height: 400, mt: 2 }}>
                                  <ResponsiveContainer>
                                    <BarChart
                                      data={otherBottom}
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
                                      <YAxis domain={otherDomain} />
                                      <Tooltip content={<DifferenceTooltip />} />
                                      <Bar dataKey="difference">
                                        {otherBottom.map((entry, index) => (
                                          <Cell 
                                            key={`cell-${index}`} 
                                            fill={entry.difference >= 0 ? '#1976d2' : '#dc004e'} 
                                          />
                                        ))}
                                      </Bar>
                                    </BarChart>
                                  </ResponsiveContainer>
                                </Box>
                              )}
                            </>
                          );
                        })()}
                      </>
                    )}
                  </>
                )}
              </Paper>
            </Grid>
          );
        } else {
          // 不需要分組，使用原來的顯示方式
          // 正數和負數的scale分開計算
          const positiveValues = differenceData.filter(item => item.difference > 0).map(item => item.difference);
          const negativeValues = differenceData.filter(item => item.difference < 0).map(item => item.difference);
          const maxPositive = positiveValues.length > 0 ? Math.max(...positiveValues) : 0;
          const minNegative = negativeValues.length > 0 ? Math.min(...negativeValues) : 0;
          const yAxisDomain = [Math.ceil(minNegative * 1.1), Math.ceil(maxPositive * 1.1)];
          
          // 如果人數小於30，只顯示一個圖表
          if (differenceData.length < 30) {
            return (
              <Grid item xs={12}>
                <Paper elevation={2} sx={{ p: 3 }}>
                  <Typography variant="h6" gutterBottom>
                    23. 出勤手作時數 - 回報步道實作帶領時數（共 {totalCount} 位）
                  </Typography>
                  <Box sx={{ width: '100%', height: 400, mt: 2 }}>
                    <ResponsiveContainer>
                      <BarChart
                        data={differenceData}
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
                        <YAxis domain={yAxisDomain} />
                        <Tooltip content={<DifferenceTooltip />} />
                        <Bar dataKey="difference">
                          {differenceData.map((entry, index) => (
                            <Cell 
                              key={`cell-${index}`} 
                              fill={entry.difference >= 0 ? '#1976d2' : '#dc004e'} 
                            />
                          ))}
                        </Bar>
                      </BarChart>
                    </ResponsiveContainer>
                  </Box>
                </Paper>
              </Grid>
            );
          }
          
          // 人數大於等於30，分成上下兩部分
          const midIndex = Math.ceil(differenceData.length / 2);
          const topHalf = differenceData.slice(0, midIndex);
          const bottomHalf = differenceData.slice(midIndex);
          
          return (
            <Grid item xs={12}>
              <Paper elevation={2} sx={{ p: 3 }}>
                <Typography variant="h6" gutterBottom>
                  23. 出勤手作時數 - 回報步道實作帶領時數（共 {totalCount} 位）
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
                      <YAxis domain={yAxisDomain} />
                      <Tooltip content={<DifferenceTooltip />} />
                      <Bar dataKey="difference">
                        {topHalf.map((entry, index) => (
                          <Cell 
                            key={`cell-${index}`} 
                            fill={entry.difference >= 0 ? '#1976d2' : '#dc004e'} 
                          />
                        ))}
                      </Bar>
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
                        <YAxis domain={yAxisDomain} />
                        <Tooltip content={<DifferenceTooltip />} />
                        <Bar dataKey="difference">
                          {bottomHalf.map((entry, index) => (
                            <Cell 
                              key={`cell-${index}`} 
                              fill={entry.difference >= 0 ? '#1976d2' : '#dc004e'} 
                            />
                          ))}
                        </Bar>
                      </BarChart>
                    </ResponsiveContainer>
                  </Box>
                )}
              </Paper>
            </Grid>
          );
        }
      })()}
    </Grid>
  );
}

