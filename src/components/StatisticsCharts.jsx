import React from 'react';
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
} from 'recharts';
import {
  calculateMonthlyActivityTypeCount,
  calculateMonthlyActivityTypeDays,
  calculateMonthlyCityCount,
  calculateMonthlyCityDays,
  calculateParticipantCount,
  calculateParticipantHours,
  getActivityTypes,
  getCities,
} from '../utils/dataProcessor.js';

export default function StatisticsCharts({ data }) {
  if (!data || data.length === 0) {
    return (
      <Typography variant="body1" color="text.secondary" align="center" sx={{ py: 4 }}>
        請先上傳 Excel 檔案以顯示統計圖表
      </Typography>
    );
  }

  // 計算各種統計資料
  const monthlyActivityTypeCount = calculateMonthlyActivityTypeCount(data);
  const monthlyActivityTypeDays = calculateMonthlyActivityTypeDays(data);
  const monthlyCityCount = calculateMonthlyCityCount(data);
  const monthlyCityDays = calculateMonthlyCityDays(data);
  const participantCount = calculateParticipantCount(data);
  const participantHours = calculateParticipantHours(data);

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

  // 對縣市進行排序：先按區域（北中南東），再按區域內順序
  const cities = [...citiesRaw].sort((a, b) => {
    const regionA = cityRegions[a] || 99; // 未定義的縣市放在最後
    const regionB = cityRegions[b] || 99;
    
    if (regionA !== regionB) {
      return regionA - regionB;
    }
    
    // 同區域內按順序排序
    const orderA = cityOrder[a] || 999;
    const orderB = cityOrder[b] || 999;
    return orderA - orderB;
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

  const activityTypeCountData = formatStackedData(monthlyActivityTypeCount, activityTypes);
  const activityTypeDaysData = formatStackedData(monthlyActivityTypeDays, activityTypes);
  const cityCountData = formatStackedData(monthlyCityCount, cities);
  const cityDaysData = formatStackedData(monthlyCityDays, cities);

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

  // 為縣市建立顏色映射，確保每個縣市都有唯一顏色
  const getCityColorMap = (cityList) => {
    const colorPalette = getColorPalette(cityList.length);
    const colorMap = {};
    cityList.forEach((city, index) => {
      colorMap[city] = colorPalette[index];
    });
    return colorMap;
  };

  const activityTypeColors = getColorPalette(activityTypes.length);
  const cityColorMap = getCityColorMap(cities);

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
                {activityTypes.map((type, index) => (
                  <Bar
                    key={type}
                    dataKey={type}
                    stackId="a"
                    fill={activityTypeColors[index % activityTypeColors.length]}
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
                {activityTypes.map((type, index) => (
                  <Bar
                    key={type}
                    dataKey={type}
                    stackId="a"
                    fill={activityTypeColors[index % activityTypeColors.length]}
                  />
                ))}
              </BarChart>
            </ResponsiveContainer>
          </Box>
        </Paper>
      </Grid>

      {/* 圖表 3: 月份活動次數（Stack：縣市） */}
      <Grid item xs={12}>
        <Paper elevation={2} sx={{ p: 3 }}>
          <Typography variant="h6" gutterBottom>
            3. 各月份活動次數統計（依縣市）
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

      {/* 圖表 4: 月份活動天數（Stack：縣市） */}
      <Grid item xs={12}>
        <Paper elevation={2} sx={{ p: 3 }}>
          <Typography variant="h6" gutterBottom>
            4. 各月份活動天數統計（依縣市）
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

      {/* 圖表 5: 人名參與活動次數 */}
      <Grid item xs={12}>
        <Paper elevation={2} sx={{ p: 3 }}>
          <Typography variant="h6" gutterBottom>
            5. 參與人員活動次數統計
          </Typography>
          <Box sx={{ width: '100%', height: 400, mt: 2 }}>
            <ResponsiveContainer>
              <BarChart
                data={participantCount}
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

      {/* 圖表 6: 人名參與活動時數 */}
      <Grid item xs={12}>
        <Paper elevation={2} sx={{ p: 3 }}>
          <Typography variant="h6" gutterBottom>
            6. 參與人員活動時數統計
          </Typography>
          <Box sx={{ width: '100%', height: 400, mt: 2 }}>
            <ResponsiveContainer>
              <BarChart
                data={participantHours}
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
    </Grid>
  );
}

