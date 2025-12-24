import React, { useRef } from 'react';
import { Box, Button, Typography, Paper, Grid } from '@mui/material';
import CloudUploadIcon from '@mui/icons-material/CloudUpload';
import DownloadIcon from '@mui/icons-material/Download';

export default function ExcelUploader({ 
  onHourLogUpload, 
  onManpowerUpload, 
  isLoadingHourLog, 
  isLoadingManpower,
  onDownload,
  isDownloadDisabled,
  isDownloading,
}) {
  const hourLogInputRef = useRef(null);
  const manpowerInputRef = useRef(null);

  const handleHourLogChange = (event) => {
    const file = event.target.files?.[0];
    if (file) {
      onHourLogUpload(file);
    }
  };

  const handleManpowerChange = (event) => {
    const file = event.target.files?.[0];
    if (file) {
      onManpowerUpload(file);
    }
  };

  const handleHourLogButtonClick = () => {
    hourLogInputRef.current?.click();
  };

  const handleManpowerButtonClick = () => {
    manpowerInputRef.current?.click();
  };

  return (
    <Grid container spacing={3} sx={{ mb: 4 }}>
      {/* 人力需求表上傳（左側） */}
      <Grid item xs={12} md={4}>
        <Paper elevation={3} sx={{ p: 3 }}>
          <Typography variant="h6" gutterBottom>
            上傳人力需求表
          </Typography>
          <Box sx={{ display: 'flex', alignItems: 'center', gap: 2 }}>
            <input
              ref={manpowerInputRef}
              type="file"
              accept=".xlsx,.xls"
              onChange={handleManpowerChange}
              style={{ display: 'none' }}
              disabled={isLoadingManpower}
            />
            <Button
              variant="contained"
              component="span"
              startIcon={<CloudUploadIcon />}
              onClick={handleManpowerButtonClick}
              disabled={isLoadingManpower}
            >
              {isLoadingManpower ? '處理中...' : '選擇檔案'}
            </Button>
          </Box>
        </Paper>
      </Grid>

      {/* 時數登錄表上傳（右側） */}
      <Grid item xs={12} md={4}>
        <Paper elevation={3} sx={{ p: 3 }}>
          <Typography variant="h6" gutterBottom>
            上傳時數登錄表
          </Typography>
          <Box sx={{ display: 'flex', alignItems: 'center', gap: 2 }}>
            <input
              ref={hourLogInputRef}
              type="file"
              accept=".xlsx,.xls"
              onChange={handleHourLogChange}
              style={{ display: 'none' }}
              disabled={isLoadingHourLog}
            />
            <Button
              variant="contained"
              component="span"
              startIcon={<CloudUploadIcon />}
              onClick={handleHourLogButtonClick}
              disabled={isLoadingHourLog}
            >
              {isLoadingHourLog ? '處理中...' : '選擇檔案'}
            </Button>
          </Box>
        </Paper>
      </Grid>

      {/* 下載（右側） */}
      <Grid item xs={12} md={4}>
        <Paper elevation={3} sx={{ p: 3 }}>
          <Typography variant="h6" gutterBottom>
            下載
          </Typography>
          <Box sx={{ display: 'flex', alignItems: 'center', gap: 2 }}>
            <Button
              variant="contained"
              color="secondary"
              startIcon={<DownloadIcon />}
              onClick={onDownload}
              disabled={!!isDownloadDisabled || !!isDownloading}
            >
              {isDownloading ? '產生中...' : '下載圖表資料（Excel）'}
            </Button>
          </Box>
        </Paper>
      </Grid>
    </Grid>
  );
}

