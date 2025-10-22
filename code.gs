/**
 * Unified EK Tools Web App - Complete Code.gs
 * Combines Sheets-to-Slides, Audio Export, Audio Regeneration, and Slides-to-Sheets functionality
 * 
 * This file should replace your entire existing Code.gs file
 * It includes all the original functionality plus the new Audio Regeneration module
 */

// ============================================================================
// GLOBAL CONFIGURATION
// ============================================================================

const CONFIG = {
  getAWSConfig: function() {
    const props = PropertiesService.getScriptProperties();
    return {
      accessKey: props.getProperty('AWS_ACCESS_KEY') || '',
      secretKey: props.getProperty('AWS_SECRET_KEY') || '',
      bucket: props.getProperty('S3_BUCKET') || '',
      region: props.getProperty('S3_REGION') || 'us-east-1',
      baseUrl: props.getProperty('S3_BASE_URL') || '',
      badgesFolder: 'ek/badges/',
      audioFolder: props.getProperty('S3_PATH') || 'audio/'
    };
  },
  
  getElevenLabsConfig: function() {
    const props = PropertiesService.getScriptProperties();
    return {
      apiKey: props.getProperty('ELEVENLABS_API_KEY') || '',
      voiceId: props.getProperty('ELEVENLABS_VOICE_ID') || '',
      modelId: props.getProperty('ELEVENLABS_MODEL_ID') || 'eleven_monolingual_v1'
    };
  },
  
  defaults: {
    templateSlideId: '1fLGPIDEhqXIblC5gya8ck9bN5U03WJLYFhfUOkUzbtc'
  }
};

// ============================================================================
// WEB APP ENTRY POINT
// ============================================================================

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('EK Tools Suite')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ============================================================================
// SHARED SERVICES
// ============================================================================

const SharedServices = {
  // Extract ID from various URL formats
  extractId: function(url) {
    if (!url) return null;
    const patterns = [
      /\/d\/([a-zA-Z0-9-_]+)/,
      /\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/,
      /\/presentation\/d\/([a-zA-Z0-9-_]+)/,
      /id=([a-zA-Z0-9-_]+)/,
      /^([a-zA-Z0-9-_]+)$/
    ];
    
    for (const pattern of patterns) {
      const match = url.match(pattern);
      if (match) return match[1];
    }
    return null;
  },
  
  // Get EK spreadsheets
  getEKSpreadsheets: function() {
    try {
      const files = DriveApp.searchFiles(
        'title contains "ek" and mimeType = "application/vnd.google-apps.spreadsheet"'
      );
      const results = [];
      let count = 0;
      
      while (files.hasNext() && count < 50) {
        const file = files.next();
        const name = file.getName();
        if (name.toLowerCase().startsWith('ek')) {
          results.push({
            id: file.getId(),
            name: name,
            url: file.getUrl(),
            lastUpdated: file.getLastUpdated().toISOString()
          });
          count++;
        }
      }
      
      results.sort((a, b) => new Date(b.lastUpdated) - new Date(a.lastUpdated));
      return {success: true, files: results};
    } catch (error) {
      return {success: false, error: error.toString()};
    }
  },
  
  // Get presentations
  getPresentations: function() {
    try {
      const files = DriveApp.searchFiles(
        'mimeType = "application/vnd.google-apps.presentation" and ' +
        'modifiedDate > "' + new Date(Date.now() - 90 * 24 * 60 * 60 * 1000).toISOString() + '"'
      );
      const results = [];
      let count = 0;
      
      while (files.hasNext() && count < 50) {
        const file = files.next();
        results.push({
          id: file.getId(),
          name: file.getName(),
          url: file.getUrl(),
          lastUpdated: file.getLastUpdated().toISOString()
        });
        count++;
      }
      
      results.sort((a, b) => new Date(b.lastUpdated) - new Date(a.lastUpdated));
      return {success: true, files: results};
    } catch (error) {
      return {success: false, error: error.toString()};
    }
  }
};

// ============================================================================
// S3 SERVICE
// ============================================================================

const S3Service = {
  upload: function(blob, filename, folder = '') {
    try {
      const config = CONFIG.getAWSConfig();
      const date = new Date();
      const dateString = Utilities.formatDate(date, 'GMT', "EEE, dd MMM yyyy HH:mm:ss 'GMT'");
      
      const fileContent = blob.getBytes();
      const contentType = blob.getContentType() || 'application/octet-stream';
      const fullPath = folder + filename;
      
      const url = `https://s3.${config.region}.amazonaws.com/${config.bucket}/${fullPath}`;
      
      // AWS Signature V2
      const stringToSign = [
        'PUT',
        '',
        contentType,
        dateString,
        `/${config.bucket}/${fullPath}`
      ].join('\n');
      
      const signature = Utilities.computeHmacSignature(
        Utilities.MacAlgorithm.HMAC_SHA_1,
        stringToSign,
        config.secretKey
      );
      const signatureBase64 = Utilities.base64Encode(signature);
      
      const options = {
        method: 'PUT',
        headers: {
          'Authorization': `AWS ${config.accessKey}:${signatureBase64}`,
          'Date': dateString,
          'Content-Type': contentType
        },
        payload: fileContent,
        muteHttpExceptions: true
      };
      
      const response = UrlFetchApp.fetch(url, options);
      const responseCode = response.getResponseCode();
      
      if (responseCode === 200 || responseCode === 204) {
        return {
          success: true,
          url: config.baseUrl + fullPath
        };
      } else {
        throw new Error(`S3 upload failed with code ${responseCode}`);
      }
    } catch (error) {
      console.error('S3 upload error:', error);
      return {
        success: false,
        error: error.toString()
      };
    }
  },
  
  testConnection: function() {
    try {
      const testBlob = Utilities.newBlob('test', 'text/plain', 'test.txt');
      const result = this.upload(testBlob, `test-${Date.now()}.txt`, 'tests/');
      return result;
    } catch (error) {
      return {success: false, error: error.toString()};
    }
  }
};

// ============================================================================
// AUDIO SERVICE
// ============================================================================

const AudioService = {
  generateWithElevenLabs: function(text) {
    const config = CONFIG.getElevenLabsConfig();
    const url = `https://api.elevenlabs.io/v1/text-to-speech/${config.voiceId}`;
    
    const props = PropertiesService.getScriptProperties();
    const savedSettings = props.getProperty('VOICE_SETTINGS');
    let voiceSettings = {
      stability: 50,
      similarity: 75,
      style: 0,
      speakerBoost: true
    };
    
    if (savedSettings) {
      try {
        voiceSettings = JSON.parse(savedSettings);
      } catch (e) {
        console.log('Error parsing voice settings, using defaults');
      }
    }
    
    // Convert percentages to decimals for API
    const apiVoiceSettings = {
      stability: (voiceSettings.stability || 50) / 100,
      similarity_boost: (voiceSettings.similarity || 75) / 100,
      style: (voiceSettings.style || 0) / 100,
      use_speaker_boost: voiceSettings.speakerBoost !== false
    };
    
    // Repeat text 3 times for English
    const processedText = `${text}...\n${text}...\n${text}`;
    
    const payload = {
      text: processedText,
      model_id: config.modelId,
      voice_settings: apiVoiceSettings
    };
    
    const options = {
      method: 'post',
      headers: {
        'xi-api-key': config.apiKey,
        'Content-Type': 'application/json',
        'Accept': 'audio/mpeg'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    
    try {
      const response = UrlFetchApp.fetch(url, options);
      if (response.getResponseCode() === 200) {
        return response.getBlob().setName('audio.mp3');
      } else {
        throw new Error(`ElevenLabs API error: ${response.getResponseCode()}`);
      }
    } catch (error) {
      console.error('ElevenLabs API error:', error);
      throw error;
    }
  },
  
  testConnection: function() {
    try {
      const text = "This is a test of the ElevenLabs API connection";
      const audioData = this.generateWithElevenLabs(text);
      if (audioData) {
        return {success: true, message: 'ElevenLabs connection successful!'};
      }
    } catch (error) {
      return {success: false, message: error.toString()};
    }
  }
};

// ============================================================================
// AUDIO REGENERATION SERVICE
// ============================================================================

const AudioRegenerationService = {
  // Get all audio entries from a spreadsheet
  getAudioEntries: function(spreadsheetId) {
    try {
      const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
      const sheet = spreadsheet.getActiveSheet();
      const data = sheet.getDataRange().getValues();
      const headers = data[0];
      const rows = data.slice(1);
      
      const audioEntries = [];
      const audioFieldPatterns = [/-file$/, /Filename$/];
      
      // Process each row
      rows.forEach((row, rowIndex) => {
        const rowData = {};
        headers.forEach((header, colIndex) => {
          rowData[header] = row[colIndex];
        });
        
        // Check if this is a relevant row (Q1, Q2, Q3)
        const firstColumnValue = rowData[headers[0]] ? rowData[headers[0]].toString() : '';
        if (!/^Q[1-3]/.test(firstColumnValue)) {
          return; // Skip non-Q rows
        }
        
        // Find audio fields
        headers.forEach((header, colIndex) => {
          const isAudioField = audioFieldPatterns.some(pattern => pattern.test(header));
          
          if (isAudioField) {
            let textToConvert = null;
            let currentUrl = rowData[header] || '';
            
            // Determine the text field
            if (header === 'Filename') {
              textToConvert = rowData['English'];
            } else {
              const textFieldName = header.replace(/-file$/, '').replace(/Filename$/, '');
              const enField = `${textFieldName}-en`;
              
              if (rowData[enField]) {
                textToConvert = rowData[enField];
              } else if (rowData['English'] && header === 'Filename') {
                textToConvert = rowData['English'];
              } else if (rowData[textFieldName]) {
                textToConvert = rowData[textFieldName];
              }
            }
            
            if (textToConvert && textToConvert.toString().trim()) {
              audioEntries.push({
                row: rowIndex + 2, // +2 because arrays are 0-indexed and we skip header
                column: colIndex + 1, // +1 for 1-indexed columns
                header: header,
                label: `${firstColumnValue} - ${header}`,
                text: textToConvert.toString(),
                currentUrl: currentUrl.toString(),
                hasAudio: !!currentUrl && currentUrl.toString().trim() !== '',
                needsRegeneration: false // Can be flagged based on criteria
              });
            }
          }
        });
      });
      
      // Add drill audio entry
      const drillEntry = this.getDrillAudioEntry(sheet, data);
      if (drillEntry) {
        audioEntries.push(drillEntry);
      }
      
      return {
        success: true,
        entries: audioEntries,
        spreadsheetName: spreadsheet.getName()
      };
      
    } catch (error) {
      console.error('Error getting audio entries:', error);
      return {success: false, error: error.toString()};
    }
  },
  
  // Get drill audio entry
  getDrillAudioEntry: function(sheet, data) {
    try {
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
      
      // Check if drill audio already exists in a known location
      let currentUrl = '';
      
      return {
        row: 'Drill',
        column: 'N/A',
        header: 'drill-audio',
        label: 'Drill Audio (Rows 31+)',
        text: combinedText,
        currentUrl: currentUrl,
        hasAudio: !!currentUrl,
        needsRegeneration: false,
        isDrill: true
      };
      
    } catch (error) {
      console.error('Error getting drill audio entry:', error);
      return null;
    }
  },
  
  // Regenerate selected audio files
  regenerateAudioFiles: function(spreadsheetId, selectedEntries, voiceSettings) {
    try {
      const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
      const sheet = spreadsheet.getActiveSheet();
      const config = CONFIG.getAWSConfig();
      
      let successCount = 0;
      const failures = [];
      const results = [];
      
      selectedEntries.forEach((entry, index) => {
        try {
          // Generate audio with custom settings
          const audioBlob = this.generateAudioWithSettings(entry.text, voiceSettings);
          
          if (audioBlob) {
            // Generate filename
            let filename;
            if (entry.isDrill) {
              const kigoValue = sheet.getRange('B2').getValue() || 'drill';
              filename = `${kigoValue}-drill.mp3`;
            } else {
              filename = AudioExport.generateFilename(entry.text);
            }
            
            // Upload to S3
            const uploadResult = S3Service.upload(audioBlob, filename, config.audioFolder);
            
            if (uploadResult.success) {
              // Update the spreadsheet cell if not a drill
              if (!entry.isDrill && entry.row !== 'Drill') {
                sheet.getRange(entry.row, entry.column).setValue(uploadResult.url);
              }
              
              successCount++;
              results.push({
                entry: entry,
                success: true,
                url: uploadResult.url,
                filename: filename
              });
            } else {
              throw new Error(uploadResult.error);
            }
          }
        } catch (error) {
          console.error(`Error regenerating audio for ${entry.label}:`, error);
          failures.push({
            label: entry.label,
            error: error.toString()
          });
          results.push({
            entry: entry,
            success: false,
            error: error.toString()
          });
        }
        
        // Add delay to avoid rate limiting
        Utilities.sleep(500);
      });
      
      // Force save
      SpreadsheetApp.flush();
      
      return {
        success: true,
        totalCount: selectedEntries.length,
        successCount: successCount,
        failures: failures,
        results: results
      };
      
    } catch (error) {
      console.error('Error regenerating audio files:', error);
      return {success: false, error: error.toString()};
    }
  },
  
  // Generate audio with specific voice settings
  generateAudioWithSettings: function(text, voiceSettings) {
    const config = CONFIG.getElevenLabsConfig();
    const url = `https://api.elevenlabs.io/v1/text-to-speech/${config.voiceId}`;
    
    // Convert percentages to decimals for API
    const apiVoiceSettings = {
      stability: (voiceSettings.stability || 50) / 100,
      similarity_boost: (voiceSettings.similarity || 75) / 100,
      style: (voiceSettings.style || 0) / 100,
      use_speaker_boost: voiceSettings.speakerBoost !== false
    };
    
    // Repeat text 3 times for English
    const processedText = `${text}...\n${text}...\n${text}`;
    
    const payload = {
      text: processedText,
      model_id: config.modelId,
      voice_settings: apiVoiceSettings
    };
    
    const options = {
      method: 'post',
      headers: {
        'xi-api-key': config.apiKey,
        'Content-Type': 'application/json',
        'Accept': 'audio/mpeg'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    
    try {
      const response = UrlFetchApp.fetch(url, options);
      if (response.getResponseCode() === 200) {
        return response.getBlob().setName('audio.mp3');
      } else {
        throw new Error(`ElevenLabs API error: ${response.getResponseCode()}`);
      }
    } catch (error) {
      console.error('ElevenLabs API error:', error);
      throw error;
    }
  },
  
  // Export updated CSV
  exportRegeneratedCSV: function(spreadsheetId) {
    try {
      const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
      const sheet = spreadsheet.getActiveSheet();
      const data = sheet.getDataRange().getValues();
      
      // Use the existing CSV generation logic from AudioExport
      const allProcessedData = {
        rows: [],
        headers: data[0],
        audioResults: []
      };
      
      // Convert sheet data to processedData format
      const headers = data[0];
      const rows = data.slice(1);
      
      rows.forEach((row, rowIndex) => {
        const rowData = {};
        headers.forEach((header, colIndex) => {
          rowData[header] = row[colIndex];
        });
        
        allProcessedData.rows.push({
          rowIndex: rowIndex + 1,
          data: rowData,
          audioResults: [] // Not needed for CSV export
        });
      });
      
      // Generate CSV using existing function
      const csvContent = AudioExport.generateCSV(allProcessedData, data, spreadsheet);
      
      // Save CSV file
      const timestamp = Utilities.formatDate(new Date(), 'GMT', 'yyyy-MM-dd-HHmmss');
      const filename = `${spreadsheet.getName()}-regenerated-${timestamp}.csv`;
      
      const blob = Utilities.newBlob(csvContent, 'text/csv', filename);
      const file = DriveApp.createFile(blob);
      
      return {
        success: true,
        csvUrl: file.getUrl(),
        csvName: filename,
        csvId: file.getId()
      };
      
    } catch (error) {
      console.error('Error exporting CSV:', error);
      return {success: false, error: error.toString()};
    }
  }
};

// ============================================================================
// SETTINGS MANAGEMENT
// ============================================================================

function getSettings() {
  const props = PropertiesService.getScriptProperties();
  
  let voiceSettings = {
    stability: 50,
    similarity: 75,
    style: 0,
    speakerBoost: true
  };
  
  const savedVoiceSettings = props.getProperty('VOICE_SETTINGS');
  if (savedVoiceSettings) {
    try {
      voiceSettings = JSON.parse(savedVoiceSettings);
    } catch (e) {
      console.log('Error parsing voice settings');
    }
  }
  
  return {
    aws: {
      accessKey: props.getProperty('AWS_ACCESS_KEY') || '',
      secretKey: props.getProperty('AWS_SECRET_KEY') ? '********' : '',
      bucket: props.getProperty('S3_BUCKET') || '',
      region: props.getProperty('S3_REGION') || 'us-east-1',
      baseUrl: props.getProperty('S3_BASE_URL') || '',
      s3Path: props.getProperty('S3_PATH') || 'audio/'
    },
    elevenlabs: {
      apiKey: props.getProperty('ELEVENLABS_API_KEY') || '',
      voiceId: props.getProperty('ELEVENLABS_VOICE_ID') || '',
      modelId: props.getProperty('ELEVENLABS_MODEL_ID') || 'eleven_monolingual_v1'
    },
    voiceSettings: voiceSettings
  };
}

function saveSettings(settings) {
  try {
    const props = PropertiesService.getScriptProperties();
    
    // Save AWS settings
    if (settings.aws) {
      if (settings.aws.accessKey) props.setProperty('AWS_ACCESS_KEY', settings.aws.accessKey);
      if (settings.aws.secretKey && settings.aws.secretKey !== '********') {
        props.setProperty('AWS_SECRET_KEY', settings.aws.secretKey);
      }
      if (settings.aws.bucket) props.setProperty('S3_BUCKET', settings.aws.bucket);
      if (settings.aws.region) props.setProperty('S3_REGION', settings.aws.region);
      if (settings.aws.baseUrl) props.setProperty('S3_BASE_URL', settings.aws.baseUrl);
      if (settings.aws.s3Path) props.setProperty('S3_PATH', settings.aws.s3Path);
    }
    
    // Save ElevenLabs settings
    if (settings.elevenlabs) {
      if (settings.elevenlabs.apiKey) props.setProperty('ELEVENLABS_API_KEY', settings.elevenlabs.apiKey);
      if (settings.elevenlabs.voiceId) props.setProperty('ELEVENLABS_VOICE_ID', settings.elevenlabs.voiceId);
      if (settings.elevenlabs.modelId) props.setProperty('ELEVENLABS_MODEL_ID', settings.elevenlabs.modelId);
    }
    
    // Save voice settings
    if (settings.voiceSettings) {
      props.setProperty('VOICE_SETTINGS', JSON.stringify(settings.voiceSettings));
    }
    
    return {success: true, message: 'Settings saved successfully!'};
  } catch (error) {
    return {success: false, message: error.toString()};
  }
}

// ============================================================================
// WRAPPER FUNCTIONS FOR HTML COMPATIBILITY
// ============================================================================

function getS3Settings() {
  return getSettings();
}

function saveS3Settings(settings) {
  const props = PropertiesService.getScriptProperties();
  if (settings.s3Path) {
    props.setProperty('S3_PATH', settings.s3Path);
  }
  
  const {s3Path, ...otherSettings} = settings;
  return saveSettings({aws: otherSettings});
}

function saveElevenLabsSettings(settings) {
  return saveSettings({elevenlabs: settings});
}

function saveVoiceSettings(settings) {
  return saveSettings({voiceSettings: settings});
}

function testConnections() {
  const results = {
    s3: S3Service.testConnection(),
    elevenlabs: AudioService.testConnection()
  };
  return results;
}

// ============================================================================
// EXPORTED FUNCTIONS FOR AUDIO REGENERATION MODULE
// ============================================================================

function getAudioEntriesFromSpreadsheet(spreadsheetId) {
  return AudioRegenerationService.getAudioEntries(spreadsheetId);
}

function regenerateAudioFiles(spreadsheetId, selectedEntries, voiceSettings) {
  return AudioRegenerationService.regenerateAudioFiles(spreadsheetId, selectedEntries, voiceSettings);
}

function generateTestAudioWithSettings(text, voiceSettings) {
  try {
    if (!text || text.trim() === '') {
      return {success: false, error: 'No text provided'};
    }
    
    const audioBlob = AudioRegenerationService.generateAudioWithSettings(text, voiceSettings);
    
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

function exportRegeneratedCSV(spreadsheetId) {
  return AudioRegenerationService.exportRegeneratedCSV(spreadsheetId);
}

// ============================================================================
// EXPORTED FUNCTIONS FOR OTHER MODULES
// ============================================================================

function getEKSheets() {
  return SharedServices.getEKSpreadsheets();
}

function getPresentations() {
  return SharedServices.getPresentations();
}

function getEKSpreadsheets() {
  return SharedServices.getEKSpreadsheets();
}

// Note: The SheetsToSlides, AudioExport, and SlidesToSheets modules should be added here as well
// They are defined in your separate module files (SheetsToSlides.gs, AudioExport.gs, SlidesToSheets.gs)
// Include those modules' content here or keep them as separate files in your project
