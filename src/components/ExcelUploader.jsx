import React, { useRef } from 'react';
import { Box, Button, Typography, Paper } from '@mui/material';
import CloudUploadIcon from '@mui/icons-material/CloudUpload';

export default function ExcelUploader({ onFileUpload, isLoading }) {
  const fileInputRef = useRef(null);

  const handleFileChange = (event) => {
    const file = event.target.files?.[0];
    if (file) {
      onFileUpload(file);
    }
  };

  const handleButtonClick = () => {
    fileInputRef.current?.click();
  };

  return (
    <Paper elevation={3} sx={{ p: 3, mb: 4 }}>
      <Typography variant="h6" gutterBottom>
        上傳 Excel 檔案
      </Typography>
      <Box sx={{ display: 'flex', alignItems: 'center', gap: 2 }}>
        <input
          ref={fileInputRef}
          type="file"
          accept=".xlsx,.xls"
          onChange={handleFileChange}
          style={{ display: 'none' }}
          disabled={isLoading}
        />
        <Button
          variant="contained"
          component="span"
          startIcon={<CloudUploadIcon />}
          onClick={handleButtonClick}
          disabled={isLoading}
        >
          {isLoading ? '處理中...' : '選擇檔案'}
        </Button>
      </Box>
    </Paper>
  );
}

