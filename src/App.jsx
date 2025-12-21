import React, { useState } from 'react';
import { Container, Typography, Box, Alert, CircularProgress } from '@mui/material';
import { ThemeProvider, createTheme } from '@mui/material/styles';
import CssBaseline from '@mui/material/CssBaseline';
import ExcelUploader from './components/ExcelUploader';
import StatisticsCharts from './components/StatisticsCharts';
import { parseExcelFile } from './utils/excelParser';
import { filterCancelled } from './utils/dataProcessor';

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

  const handleFileUpload = async (file) => {
    setLoading(true);
    setError(null);
    setData([]);

    try {
      // 解析 Excel 檔案
      const rawData = await parseExcelFile(file);
      
      // 過濾掉取消的活動
      const filteredData = filterCancelled(rawData);
      
      if (filteredData.length === 0) {
        setError('沒有找到有效的活動資料，請確認 Excel 檔案格式是否正確');
      } else {
        setData(filteredData);
      }
    } catch (err) {
      console.error('處理檔案時發生錯誤:', err);
      setError(err.message || '處理檔案時發生錯誤，請確認檔案格式是否正確');
    } finally {
      setLoading(false);
    }
  };

  return (
    <ThemeProvider theme={theme}>
      <CssBaseline />
      <Container maxWidth="xl" sx={{ py: 4 }}>
        <Typography variant="h4" component="h1" gutterBottom sx={{ mb: 3, fontWeight: 'bold' }}>
          活動統計儀表板
        </Typography>

        <ExcelUploader onFileUpload={handleFileUpload} isLoading={loading} />

        {loading && (
          <Box sx={{ display: 'flex', justifyContent: 'center', my: 4 }}>
            <CircularProgress />
          </Box>
        )}

        {error && (
          <Alert severity="error" sx={{ mb: 3 }}>
            {error}
          </Alert>
        )}

        {!loading && data.length > 0 && (
          <Box sx={{ mt: 2 }}>
            <Alert severity="success" sx={{ mb: 3 }}>
              成功載入 {data.length} 筆活動資料
            </Alert>
            <StatisticsCharts data={data} />
          </Box>
        )}
      </Container>
    </ThemeProvider>
  );
}

export default App;

