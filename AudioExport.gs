/**
 * Module 2: Audio Export
 * Generates audio files from EK spreadsheets and creates WordPress CSV
 */

const AudioExport = {
  // Generate preview without creating audio
  generatePreview: function(spreadsheetId, options = {}) {
    try {
      if (!spreadsheetId) {
        return {success: false, error: 'Please select a spreadsheet'};
      }
      
      const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
      const sheet = spreadsheet.getActiveSheet();
      const data = sheet.getDataRange().getValues();
      const headers = data[0];
      const rows = data.slice(1);
      
      if (rows.length === 0) {
        return {success: false, error: 'No data found in spreadsheet'};
      }
      
      const rowsToProcess = options.testMode ? rows.slice(0, 3) : rows;
      const previewItems = [];
      
      rowsToProcess.forEach((row, rowIndex) => {
        const rowData = {};
        headers.forEach((header, colIndex) => {
          rowData[header] = row[colIndex];
        });
        
        const audioPreview = this.getAudioFieldsPreview(rowData, headers, options);
        
        audioPreview.forEach(item => {
          previewItems.push({
            row: rowIndex + 2,
            field: item.field,
            text: item.text,
            filename: item.filename,
            s3Url: item.s3Url,
            skipped: item.skipped || false
          });
        });
      });
      
      // Add drill audio preview
      const drillPreview = this.getDrillAudioPreview(sheet, data);
      if (drillPreview) {
        previewItems.push(drillPreview);
      }
      
      return {
        success: true,
        preview: {
          items: previewItems,
          totalRows: rowsToProcess.length,
          spreadsheetName: spreadsheet.getName()
        }
      };
      
    } catch (error) {
      console.error('Error generating preview:', error);
      return {success: false, error: error.toString()};
    }
  },
  
  // Process spreadsheet and generate audio
  processSpreadsheet: function(spreadsheetId, options = {}) {
    try {
      if (!spreadsheetId) {
        return {success: false, error: 'Please select a spreadsheet'};
      }
      
      const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
      const sheet = spreadsheet.getActiveSheet();
      const data = sheet.getDataRange().getValues();
      const headers = data[0];
      const rows = data.slice(1);
      
      if (rows.length === 0) {
        return {success: false, error: 'No data found in spreadsheet'};
      }
      
      const allProcessedData = {
        rows: [],
        headers: headers,
        audioResults: []
      };
      
      // Process each row
      rows.forEach((row, rowIndex) => {
        const rowData = {};
        headers.forEach((header, colIndex) => {
          rowData[header] = row[colIndex];
        });
        
        const audioResults = this.processAudioFields(rowData, headers, options);
        
        allProcessedData.rows.push({
          rowIndex: rowIndex + 1,
          data: rowData,
          audioResults: audioResults
        });
        
        allProcessedData.audioResults.push(...audioResults);
        
        Utilities.sleep(100); // Avoid rate limiting
      });
      
      // Process drill audio
      const drillResult = this.generateDrillAudio(sheet, data);
      if (drillResult && drillResult.success > 0) {
        allProcessedData.audioResults.push(drillResult.details[0]);
      }
      
      // Generate CSV
      const csvContent = this.generateCSV(allProcessedData, data, spreadsheet);
      const csvFile = this.saveCSVToDrive(csvContent, spreadsheet.getName());
      
      return {
        success: true,
        results: allProcessedData.rows.map(r => ({
          row: r.rowIndex + 1,
          code: r.data.code || r.data.Title || `Row ${r.rowIndex + 1}`,
          audioGenerated: r.audioResults.length,
          success: r.audioResults.filter(ar => ar.success).length,
          details: r.audioResults
        })),
        summary: {
          totalRows: rows.length,
          processedRows: allProcessedData.rows.length,
          totalAudioGenerated: allProcessedData.audioResults.length,
          totalAudioSuccess: allProcessedData.audioResults.filter(r => r.success).length,
          csvFile: {
            name: csvFile.getName(),
            url: csvFile.getUrl(),
            id: csvFile.getId()
          }
        }
      };
      
    } catch (error) {
      console.error('Error processing spreadsheet:', error);
      return {success: false, error: error.toString()};
    }
  },
  
  // Get audio fields preview
  getAudioFieldsPreview: function(rowData, headers, options = {}) {
    const audioPreview = [];
    const audioFieldPatterns = [/-file$/, /Filename$/];
    const skipExisting = options.skipExisting || false;
    const config = CONFIG.getAWSConfig();
    
    // Check if first column starts with Q1, Q2, or Q3
    const firstColumnValue = rowData[headers[0]] ? rowData[headers[0]].toString() : '';
    const shouldProcessRow = /^Q[1-3]/.test(firstColumnValue);
    
    if (!shouldProcessRow) {
      return audioPreview;
    }
    
    headers.forEach(header => {
      const isAudioField = audioFieldPatterns.some(pattern => pattern.test(header));
      
      if (isAudioField) {
        if (skipExisting && rowData[header]) {
          audioPreview.push({
            field: header,
            text: '[Already has audio]',
            filename: rowData[header].split('/').pop() || 'existing-file.mp3',
            s3Url: rowData[header],
            skipped: true
          });
          return;
        }
        
        let filename = null;
        let textToConvert = null;
        
        if (header === 'Filename' && rowData[header] && rowData[header].toString().trim()) {
          filename = rowData[header].toString().trim();
          textToConvert = rowData['English'];
        } else {
          const textFieldName = header.replace(/-file$/, '').replace(/Filename$/, '');
          const enField = `${textFieldName}-en`;
          
          if (rowData[enField] && rowData[enField].toString().trim()) {
            textToConvert = rowData[enField];
          } else if (rowData['English'] && header === 'Filename') {
            textToConvert = rowData['English'];
          } else if (rowData[textFieldName] && rowData[textFieldName].toString().trim()) {
            textToConvert = rowData[textFieldName];
          }
          
          if (textToConvert && textToConvert.toString().trim()) {
            filename = this.generateFilename(textToConvert);
          }
        }
        
        if (filename && textToConvert && textToConvert.toString().trim()) {
          const s3Url = `${config.baseUrl}${config.audioFolder}${filename}`;
          
          audioPreview.push({
            field: header,
            text: textToConvert.toString(),
            filename: filename,
            s3Url: s3Url,
            skipped: false
          });
        }
      }
    });
    
    return audioPreview;
  },
  
  // Process audio fields
  processAudioFields: function(rowData, headers, options = {}) {
    const audioResults = [];
    const audioFieldPatterns = [/-file$/, /Filename$/];
    const skipExisting = options.skipExisting || false;
    const config = CONFIG.getAWSConfig();
    
    // Check if first column starts with Q1, Q2, or Q3
    const firstColumnValue = rowData[headers[0]] ? rowData[headers[0]].toString() : '';
    const shouldProcessRow = /^Q[1-3]/.test(firstColumnValue);
    
    if (!shouldProcessRow) {
      return audioResults;
    }
    
    headers.forEach(header => {
      const isAudioField = audioFieldPatterns.some(pattern => pattern.test(header));
      
      if (isAudioField) {
        if (skipExisting && rowData[header]) {
          audioResults.push({
            field: header,
            filename: rowData[header],
            s3Url: rowData[header],
            success: true,
            skipped: true
          });
          return;
        }
        
        let filename = null;
        let textToConvert = null;
        
        if (header === 'Filename' && rowData[header] && rowData[header].toString().trim()) {
          filename = rowData[header].toString().trim();
          textToConvert = rowData['English'];
        } else {
          const textFieldName = header.replace(/-file$/, '').replace(/Filename$/, '');
          const enField = `${textFieldName}-en`;
          
          if (rowData[enField] && rowData[enField].toString().trim()) {
            textToConvert = rowData[enField];
          } else if (rowData['English'] && header === 'Filename') {
            textToConvert = rowData['English'];
          } else if (rowData[textFieldName] && rowData[textFieldName].toString().trim()) {
            textToConvert = rowData[textFieldName];
          }
          
          if (textToConvert && textToConvert.toString().trim()) {
            filename = this.generateFilename(textToConvert);
          }
        }
        
        if (filename && textToConvert && textToConvert.toString().trim()) {
          try {
            const audioData = AudioService.generateWithElevenLabs(textToConvert);
            
            if (audioData) {
              const result = S3Service.upload(audioData, filename, config.audioFolder);
              
              if (result.success) {
                audioResults.push({
                  field: header,
                  filename: filename,
                  s3Url: result.url,
                  success: true
                });
                
                rowData[header] = result.url;
              } else {
                throw new Error(result.error);
              }
            }
          } catch (error) {
            console.error(`Error generating audio for ${header}:`, error);
            audioResults.push({
              field: header,
              filename: filename,
              success: false,
              error: error.toString()
            });
          }
        }
      }
    });
    
    return audioResults;
  },
  
  // Get drill audio preview
  getDrillAudioPreview: function(sheet, data) {
    try {
      const config = CONFIG.getAWSConfig();
      const kigoValue = data[1] && data[1][1] ? data[1][1].toString() : 'drill';
      const drillFilename = `${kigoValue}-drill.mp3`;
      
      const startRow = 30;
      
      if (data.length <= startRow) {
        return null;
      }
      
      const sentences = [];
      const maxRows = Math.min(6, data.length - startRow);
      
      for (let i = 0; i < maxRows; i++) {
        const rowIndex = startRow + i;
        if (data[rowIndex] && data[rowIndex][2]) {
          const text = data[rowIndex][2].toString().trim();
          if (text) {
            const splitSentences = text.split(/\.\s+/).filter(s => s.trim());
            if (splitSentences.length > 1) {
              splitSentences.forEach(s => {
                const cleanSentence = s.trim();
                if (cleanSentence && !cleanSentence.endsWith('.')) {
                  sentences.push(cleanSentence + '.');
                } else if (cleanSentence) {
                  sentences.push(cleanSentence);
                }
              });
            } else if (text) {
              sentences.push(text);
            }
          }
        }
        
        if (sentences.length >= 6) {
          break;
        }
      }
      
      if (sentences.length === 0) {
        return null;
      }
      
      const combinedText = sentences.slice(0, 6).join(' ');
      const s3Url = `${config.baseUrl}${config.audioFolder}${drillFilename}`;
      
      return {
        row: 'Row 31+',
        field: 'drill-audio',
        text: combinedText,
        filename: drillFilename,
        s3Url: s3Url,
        skipped: false,
        isDrill: true
      };
      
    } catch (error) {
      console.error('Error in getDrillAudioPreview:', error);
      return null;
    }
  },
  
  // Generate drill audio
  generateDrillAudio: function(sheet, data) {
    try {
      const config = CONFIG.getAWSConfig();
      const kigoValue = data[1] && data[1][1] ? data[1][1].toString() : 'drill';
      const drillFilename = `${kigoValue}-drill.mp3`;
      
      const startRow = 30;
      
      if (data.length <= startRow) {
        console.log('Not enough rows for drill audio');
        return null;
      }
      
      const sentences = [];
      const maxRows = Math.min(6, data.length - startRow);
      
      for (let i = 0; i < maxRows; i++) {
        const rowIndex = startRow + i;
        if (data[rowIndex] && data[rowIndex][2]) {
          const text = data[rowIndex][2].toString().trim();
          if (text) {
            const splitSentences = text.split(/\.\s+/).filter(s => s.trim());
            if (splitSentences.length > 1) {
              splitSentences.forEach(s => {
                const cleanSentence = s.trim();
                if (cleanSentence && !cleanSentence.endsWith('.')) {
                  sentences.push(cleanSentence + '.');
                } else if (cleanSentence) {
                  sentences.push(cleanSentence);
                }
              });
            } else if (text) {
              sentences.push(text);
            }
          }
        }
        
        if (sentences.length >= 6) {
          break;
        }
      }
      
      if (sentences.length === 0) {
        console.log('No sentences found for drill audio');
        return null;
      }
      
      const combinedText = sentences.slice(0, 6).join(' ');
      
      try {
        const audioData = AudioService.generateWithElevenLabs(combinedText);
        
        if (audioData) {
          const result = S3Service.upload(audioData, drillFilename, config.audioFolder);
          
          if (result.success) {
            return {
              row: 'Drill Audio (Row 31+)',
              code: kigoValue,
              audioGenerated: 1,
              success: 1,
              details: [{
                field: 'drill-audio',
                filename: drillFilename,
                s3Url: result.url,
                success: true,
                sentences: sentences.length,
                text: combinedText.substring(0, 100) + (combinedText.length > 100 ? '...' : '')
              }]
            };
          } else {
            throw new Error(result.error);
          }
        }
      } catch (error) {
        console.error('Error generating drill audio:', error);
        return {
          row: 'Drill Audio (Row 31+)',
          code: kigoValue,
          audioGenerated: 0,
          success: 0,
          details: [{
            field: 'drill-audio',
            filename: drillFilename,
            success: false,
            error: error.toString()
          }]
        };
      }
      
    } catch (error) {
      console.error('Error in generateDrillAudio:', error);
      return null;
    }
  },
  
  // Generate filename from text
  generateFilename: function(text) {
    let filename = text.toString().toLowerCase();
    filename = filename.replace(/[.,\/#!$%\^&\*;:{}=\-_`~()?'"]/g, '');
    filename = filename.replace(/\s+/g, '-');
    filename = filename.replace(/-+/g, '-');
    
    if (filename.length > 50) {
      filename = filename.substring(0, 50);
    }
    
    filename = filename.replace(/-$/, '');
    return `${filename}.mp3`;
  },
  
// Generate CSV - FIXED VERSION
generateCSV: function(processedData, originalData, spreadsheet) {
  const csvRow = {};
  const config = CONFIG.getAWSConfig(); // Get config for URL construction
  
  // Extract basic info
  // Row 2, Column B for code (ek-0036)
  const codeRow = originalData[1];
  csvRow.code = codeRow && codeRow[1] ? codeRow[1].toString() : '';
  
  // Row 3, Column B for title (お互いの祖父母)
  const titleRow = originalData[2];
  csvRow.title = titleRow && titleRow[1] ? titleRow[1].toString() : '';
  
  // Badge from row 5, column B
  csvRow.badge = originalData[4] && originalData[4][1] ? originalData[4][1].toString() : '';
  
  // Description from row 4, column B
  csvRow.description = originalData[3] && originalData[3][1] ? originalData[3][1].toString() : '';
  
  // Initialize Q1, Q2, Q3 fields
  const qFields = ['Q1', 'Q2', 'Q3'];
  qFields.forEach(q => {
    csvRow[`${q}-jp`] = '';
    csvRow[`${q}-en`] = '';
    csvRow[`${q}-file`] = '';
    
    for (let i = 1; i <= 3; i++) {
      csvRow[`${q}-self-${i}-jp`] = '';
      csvRow[`${q}-self-${i}-en`] = '';
      csvRow[`${q}-self-${i}-file`] = '';
    }
    
    for (let i = 1; i <= 3; i++) {
      csvRow[`${q}-follow-${i}-jp`] = '';
      csvRow[`${q}-follow-${i}-en`] = '';
      csvRow[`${q}-follow-${i}-file`] = '';
    }
  });
  
  // Initialize Grammar and Practice fields
  csvRow['Grammar-example-jp'] = '';
  csvRow['Grammar-example-en'] = '';
  csvRow['Grammar-basic-en'] = '';
  csvRow['Grammar-reason-en'] = '';
  csvRow['Grammar-encouragement-en'] = '';
  
  csvRow['Practice 1 old-en'] = '';
  csvRow['practice 1 new-en'] = '';
  csvRow['Practice 2 old-en'] = '';
  csvRow['Practice 2 new-en'] = '';
  csvRow['Practice 3 old-en'] = '';
  csvRow['practice 3 new-en'] = '';
  
  // Map data from processed rows
  processedData.rows.forEach(processedRow => {
    const rowData = processedRow.data;
    const firstCol = rowData[processedData.headers[0]];
    
    if (!firstCol) return;
    
    const firstColStr = firstCol.toString();
    
    // Map Q1, Q2, Q3 rows
    if (firstColStr.startsWith('Q1') || firstColStr.startsWith('Q2') || firstColStr.startsWith('Q3')) {
      const qNum = firstColStr.substring(0, 2);
      
      let prefix = qNum;
      
      const selfMatch = firstColStr.match(/self[- ]?(\d)?/i);
      if (selfMatch) {
        const selfNum = selfMatch[1] || '1';
        prefix = `${qNum}-self-${selfNum}`;
      } else {
        const followMatch = firstColStr.match(/follow[- ]?(\d)?/i);
        if (followMatch) {
          const followNum = followMatch[1] || '1';
          prefix = `${qNum}-follow-${followNum}`;
        }
      }
      
      if (rowData[processedData.headers[1]]) {
        csvRow[`${prefix}-jp`] = rowData[processedData.headers[1]].toString();
      }
      
      if (rowData[processedData.headers[2]]) {
        csvRow[`${prefix}-en`] = rowData[processedData.headers[2]].toString();
      }
      
      // FIXED: Handle file field - check for Filename column (header index 3)
      if (rowData[processedData.headers[3]]) {
        const fileValue = rowData[processedData.headers[3]].toString();
        
        // Check if it's already a full URL or just a filename
        if (fileValue.startsWith('http://') || fileValue.startsWith('https://')) {
          // It's already a full URL (from S3 upload)
          csvRow[`${prefix}-file`] = fileValue;
        } else if (fileValue.trim() !== '') {
          // It's just a filename, construct the full URL
          csvRow[`${prefix}-file`] = `${config.baseUrl}${config.audioFolder}${fileValue}`;
        }
      }
      
      // Also check if there's a 'Filename' column that was updated with URL
      if (rowData['Filename']) {
        const filenameValue = rowData['Filename'].toString();
        if (filenameValue.startsWith('http://') || filenameValue.startsWith('https://')) {
          csvRow[`${prefix}-file`] = filenameValue;
        }
      }
    }
    
    // Map Grammar and Practice rows
    if (firstColStr.toLowerCase().includes('grammar')) {
      if (rowData[processedData.headers[1]]) {
        csvRow['Grammar-example-jp'] = rowData[processedData.headers[1]].toString();
      }
      if (rowData[processedData.headers[2]]) {
        csvRow['Grammar-example-en'] = rowData[processedData.headers[2]].toString();
      }
    }
    
    if (firstColStr.toLowerCase().includes('basic')) {
      if (rowData[processedData.headers[2]]) {
        csvRow['Grammar-basic-en'] = rowData[processedData.headers[2]].toString();
      }
    }
    
    if (firstColStr.toLowerCase().includes('reason')) {
      if (rowData[processedData.headers[2]]) {
        csvRow['Grammar-reason-en'] = rowData[processedData.headers[2]].toString();
      }
    }
    
    if (firstColStr.toLowerCase().includes('encouragement')) {
      if (rowData[processedData.headers[2]]) {
        csvRow['Grammar-encouragement-en'] = rowData[processedData.headers[2]].toString();
      }
    }
    
    if (firstColStr.toLowerCase().includes('practice')) {
      const practiceMatch = firstColStr.match(/practice\s*(\d)/i);
      if (practiceMatch) {
        const practiceNum = practiceMatch[1];
        
        if (firstColStr.toLowerCase().includes('old')) {
          if (rowData[processedData.headers[2]]) {
            csvRow[`Practice ${practiceNum} old-en`] = rowData[processedData.headers[2]].toString();
          }
        } else if (firstColStr.toLowerCase().includes('new')) {
          if (rowData[processedData.headers[2]]) {
            csvRow[`practice ${practiceNum} new-en`] = rowData[processedData.headers[2]].toString();
          }
        }
      }
    }
  });
  
  // Create CSV content
  const orderedFields = [
    'code', 'title', 'badge', 'description'
  ];
  
  ['Q1', 'Q2', 'Q3'].forEach(q => {
    orderedFields.push(`${q}-jp`, `${q}-en`, `${q}-file`);
    
    for (let i = 1; i <= 3; i++) {
      orderedFields.push(`${q}-self-${i}-jp`, `${q}-self-${i}-en`, `${q}-self-${i}-file`);
    }
    
    for (let i = 1; i <= 3; i++) {
      orderedFields.push(`${q}-follow-${i}-jp`, `${q}-follow-${i}-en`, `${q}-follow-${i}-file`);
    }
  });
  
  orderedFields.push(
    'Grammar-example-jp',
    'Grammar-example-en',
    'Grammar-basic-en',
    'Grammar-reason-en',
    'Grammar-encouragement-en',
    'Practice 1 old-en',
    'practice 1 new-en',
    'Practice 2 old-en',
    'Practice 2 new-en',
    'Practice 3 old-en',
    'practice 3 new-en'
  );
  
  const csvRows = [];
  
  // Add headers
  csvRows.push(orderedFields.map(h => `"${h}"`).join(','));
  
  // Add data row
  const values = orderedFields.map(header => {
    const value = csvRow[header] || '';
    const escaped = value.toString().replace(/"/g, '""');
    return `"${escaped}"`;
  });
  csvRows.push(values.join(','));
  
  return csvRows.join('\n');
},
  
  // Save CSV to Drive
  saveCSVToDrive: function(csvContent, baseName = 'export') {
    const timestamp = Utilities.formatDate(new Date(), 'GMT', 'yyyy-MM-dd-HHmmss');
    const filename = `${baseName}-wordpress-import-${timestamp}.csv`;
    
    const blob = Utilities.newBlob(csvContent, 'text/csv', filename);
    const file = DriveApp.createFile(blob);
    
    return file;
  },
  
  // Generate test audio
  generateTestAudio: function(text) {
    try {
      if (!text || text.trim() === '') {
        return {success: false, error: 'No text provided'};
      }
      
      const audioBlob = AudioService.generateWithElevenLabs(text);
      
      if (audioBlob) {
        const base64Audio = Utilities.base64Encode(audioBlob.getBytes());
        return {
          success: true,
          audioBase64: base64Audio,
          message: 'Test audio generated successfully'
        };
      } else {
        return {success: false, error: 'Failed to generate audio'};
      }
      
    } catch (error) {
      console.error('Error generating test audio:', error);
      return {success: false, error: error.toString()};
    }
  }
};

// Export functions for global access
function generatePreviewData(spreadsheetId, options) {
  return AudioExport.generatePreview(spreadsheetId, options);
}

function processSpreadsheetWithProgress(spreadsheetId, options) {
  return AudioExport.processSpreadsheet(spreadsheetId, options);
}

// IMPORTANT: Export the generateTestAudio function
function generateTestAudio(text) {
  return AudioExport.generateTestAudio(text);
}

function getRecentSpreadsheets() {
  return SharedServices.getEKSpreadsheets();
}

function searchSpreadsheets(searchTerm) {
  try {
    if (!searchTerm || searchTerm.trim() === '') {
      return SharedServices.getEKSpreadsheets();
    }
    
    const files = DriveApp.getFilesByType(MimeType.GOOGLE_SHEETS);
    const spreadsheets = [];
    const searchLower = searchTerm.toLowerCase();
    
    while (files.hasNext()) {
      const file = files.next();
      const fileName = file.getName();
      
      if (fileName.toLowerCase().includes(searchLower)) {
        spreadsheets.push({
          id: file.getId(),
          name: fileName,
          url: file.getUrl(),
          lastUpdated: file.getLastUpdated().toISOString()
        });
      }
      
      if (spreadsheets.length >= 50) break;
    }
    
    spreadsheets.sort((a, b) => {
      const aExact = a.name.toLowerCase() === searchLower;
      const bExact = b.name.toLowerCase() === searchLower;
      
      if (aExact && !bExact) return -1;
      if (!aExact && bExact) return 1;
      
      return new Date(b.lastUpdated) - new Date(a.lastUpdated);
    });
    
    return {success: true, spreadsheets: spreadsheets};
  } catch (error) {
    console.error('Error searching spreadsheets:', error);
    return {success: false, error: error.toString(), spreadsheets: []};
  }
}

function getVoices() {
  try {
    const config = CONFIG.getElevenLabsConfig();
    
    if (!config.apiKey) {
      return {success: false, message: 'Please set your ElevenLabs API key first'};
    }
    
    const response = UrlFetchApp.fetch('https://api.elevenlabs.io/v1/voices', {
      headers: {
        'xi-api-key': config.apiKey
      },
      muteHttpExceptions: true
    });
    
    if (response.getResponseCode() === 200) {
      const data = JSON.parse(response.getContentText());
      return {success: true, voices: data.voices};
    } else {
      return {success: false, message: 'Failed to fetch voices'};
    }
  } catch (error) {
    return {success: false, message: error.toString()};
  }
}
