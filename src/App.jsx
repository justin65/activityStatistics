import React, { useState } from 'react';
import { Container, Typography, Box, Alert, CircularProgress } from '@mui/material';
import { ThemeProvider, createTheme } from '@mui/material/styles';
import CssBaseline from '@mui/material/CssBaseline';
import ExcelUploader from './components/ExcelUploader';
import StatisticsCharts from './components/StatisticsCharts';
import { parseExcelFile } from './utils/excelParser';
import { filterCancelled } from './utils/dataProcessor';
import { parseHourLogFile } from './utils/hourLogParser';
import { processHourLogData, calculateVolunteerHoursByContent, calculateVolunteerTotalHoursForContentType } from './utils/hourLogProcessor';

const theme = createTheme({
  palette: {
    mode: 'light',
    primary: {
      main: '#1976d2',
    },
    secondary: {
      main: '#dc004e',
    },
  },
});

function App() {
  const [data, setData] = useState([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);
  const [hourLogData, setHourLogData] = useState(null);
  const [loadingHourLog, setLoadingHourLog] = useState(false);
  const [errorHourLog, setErrorHourLog] = useState(null);

  const handleManpowerUpload = async (file) => {
    setLoading(true);
    setError(null);
    setData([]);

    try {
      console.log('開始解析人力需求表...');
      // 解析 Excel 檔案
      const rawData = await parseExcelFile(file);
      console.log(`解析完成，共 ${rawData.length} 筆原始資料`);
      
      // 過濾掉取消的活動
      const filteredData = filterCancelled(rawData);
      console.log(`過濾完成，共 ${filteredData.length} 筆有效資料`);
      
      if (filteredData.length === 0) {
        setError('沒有找到有效的活動資料，請確認 Excel 檔案格式是否正確');
      } else {
        setData(filteredData);
        console.log('資料載入完成');
      }
    } catch (err) {
      console.error('處理檔案時發生錯誤:', err);
      setError(err.message || '處理檔案時發生錯誤，請確認檔案格式是否正確');
    } finally {
      setLoading(false);
    }
  };

  const handleHourLogUpload = async (file) => {
    setLoadingHourLog(true);
    setErrorHourLog(null);
    setHourLogData(null);

    try {
      // 解析時數登錄表 Excel 檔案
      const rawData = await parseHourLogFile(file);
      
      if (rawData.length === 0) {
        setErrorHourLog('沒有找到有效的時數登錄資料，請確認 Excel 檔案格式是否正確');
      } else {
        // 處理數據（提取姓名最後兩字並比對）
        const processedData = processHourLogData(rawData);
        
        // 計算統計資料
        const stats = calculateVolunteerHoursByContent(processedData); // 預設維持 2025 年行為（圖表 21）
        const retrainingStats = calculateVolunteerTotalHoursForContentType(processedData, {
          contentType: '回流訓練',
          startYear: 2023,
          endYear: 2025,
        });

        setHourLogData({
          ...stats,
          retraining: retrainingStats,
        });
      }
    } catch (err) {
      console.error('處理時數登錄表時發生錯誤:', err);
      setErrorHourLog(err.message || '處理時數登錄表時發生錯誤，請確認檔案格式是否正確');
    } finally {
      setLoadingHourLog(false);
    }
  };

  return (
    <ThemeProvider theme={theme}>
      <CssBaseline />
      <Container maxWidth="xl" sx={{ py: 4 }}>
        <Typography variant="h4" component="h1" gutterBottom sx={{ mb: 3, fontWeight: 'bold' }}>
          活動統計儀表板
        </Typography>

        <ExcelUploader 
          onHourLogUpload={handleHourLogUpload}
          onManpowerUpload={handleManpowerUpload}
          isLoadingHourLog={loadingHourLog}
          isLoadingManpower={loading}
        />

        {(loading || loadingHourLog) && (
          <Box sx={{ display: 'flex', justifyContent: 'center', my: 4 }}>
            <CircularProgress />
          </Box>
        )}

        {error && (
          <Alert severity="error" sx={{ mb: 3 }}>
            人力需求表：{error}
          </Alert>
        )}

        {errorHourLog && (
          <Alert severity="error" sx={{ mb: 3 }}>
            時數登錄表：{errorHourLog}
          </Alert>
        )}

        {!loading && data.length > 0 && (
          <Box sx={{ mt: 2 }}>
            <Alert severity="success" sx={{ mb: 3 }}>
              成功載入 {data.length} 筆活動資料
            </Alert>
          </Box>
        )}

        {!loadingHourLog && hourLogData && (
          <Box sx={{ mt: 2 }}>
            <Alert severity="success" sx={{ mb: 3 }}>
              成功載入時數登錄資料（{hourLogData.data.length} 位志工）
            </Alert>
          </Box>
        )}

        {(!loading && data.length > 0) || (!loadingHourLog && hourLogData) ? (
          <Box sx={{ mt: 2 }}>
            <StatisticsCharts data={data} hourLogData={hourLogData} />
          </Box>
        ) : null}
      </Container>
    </ThemeProvider>
  );
}

export default App;

