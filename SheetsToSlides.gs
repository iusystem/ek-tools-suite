/**
 * Module 1: Sheets to Slides Generator
 * Generates presentations from EK spreadsheets with badge upload
 */

const SheetsToSlides = {
  // Process sheet and generate slides
  processSheet: function(spreadsheetUrl, templateId, badgeImageData) {
    try {
      const spreadsheetId = SharedServices.extractId(spreadsheetUrl);
      if (!spreadsheetId) {
        return {success: false, error: 'Invalid spreadsheet URL'};
      }
      
      const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
      const sheet = spreadsheet.getActiveSheet();
      
      // Check if it's an "ek" file
      const fileName = spreadsheet.getName();
      if (!fileName.toLowerCase().startsWith('ek')) {
        return {success: false, error: 'This tool only works with files starting with "ek"'};
      }
      
      // Use provided template ID or default
      const actualTemplateId = templateId || CONFIG.defaults.templateSlideId;
      
      // Get data mappings
      const mappings = this.createDataMappings(sheet);
      if (Object.keys(mappings).length === 0) {
        return {success: false, error: 'No data mappings found in the sheet'};
      }
      
      // Upload badge to S3 if provided
      let badgeS3Url = null;
      if (badgeImageData && badgeImageData.data) {
        const badgeUploadResult = this.uploadBadgeToS3(badgeImageData, spreadsheet.getName());
        if (badgeUploadResult.success) {
          badgeS3Url = badgeUploadResult.url;
          console.log('Badge uploaded to S3:', badgeS3Url);
          
          // Write the badge URL to cell B5
          try {
            sheet.getRange('B5').setValue(badgeS3Url);
            SpreadsheetApp.flush();
            console.log('Badge URL written to cell B5');
          } catch (cellError) {
            console.error('Failed to write badge URL to B5:', cellError);
          }
        } else {
          console.error('Failed to upload badge to S3:', badgeUploadResult.error);
        }
      }
      
      // Create the presentation
      const result = this.generatePresentation(spreadsheet, actualTemplateId, mappings, badgeImageData);
      
      return {
        success: true,
        url: result.url,
        name: result.name,
        replacements: result.replacements,
        imageReplacements: result.imageReplacements,
        mappingsCount: Object.keys(mappings).length,
        badgeS3Url: badgeS3Url,
        badgeUrlUpdated: badgeS3Url ? true : false
      };
      
    } catch (error) {
      console.error('Error processing sheet:', error);
      return {success: false, error: error.message};
    }
  },
  
  // Upload badge image to S3
  uploadBadgeToS3: function(badgeImageData, spreadsheetName) {
    try {
      if (!badgeImageData || !badgeImageData.data) {
        return {success: false, error: 'No badge data provided'};
      }
      
      const binaryData = Utilities.base64Decode(badgeImageData.data);
      const blob = Utilities.newBlob(binaryData, badgeImageData.type, badgeImageData.name);
      
      // Extract ek-xxxx from spreadsheet name
      let ekCode = '';
      const nameMatch = spreadsheetName.match(/ek[\s-]?([a-zA-Z0-9]+)/i);
      if (nameMatch && nameMatch[1]) {
        ekCode = nameMatch[1].toLowerCase();
      } else {
        const ekIndex = spreadsheetName.toLowerCase().indexOf('ek');
        if (ekIndex !== -1) {
          ekCode = spreadsheetName.substring(ekIndex + 2, ekIndex + 6).replace(/[^a-zA-Z0-9]/g, '');
        }
      }
      
      if (!ekCode) {
        ekCode = 'badge';
      }
      
      const config = CONFIG.getAWSConfig();
      const filename = `ek-${ekCode}-badge.png`;
      
      const result = S3Service.upload(blob, filename, config.badgesFolder);
      
      return {
        success: result.success,
        url: result.url,
        filename: filename,
        error: result.error
      };
      
    } catch (error) {
      console.error('Error uploading badge to S3:', error);
      return {
        success: false,
        error: error.toString()
      };
    }
  },
  
  // Generate the presentation
  generatePresentation: function(spreadsheet, templateId, mappings, badgeImageData) {
    const templateFile = DriveApp.getFileById(templateId);
    const newPresentationName = `${spreadsheet.getName()} - Generated Slides`;
    const copiedFile = templateFile.makeCopy(newPresentationName);
    
    const newPresentation = SlidesApp.openById(copiedFile.getId());
    const slides = newPresentation.getSlides();
    
    let totalReplacements = 0;
    let totalImageReplacements = 0;
    
    // Process badge image if provided
    let badgeImageBlob = null;
    let badgeUrlPattern = null;
    
    if (badgeImageData && badgeImageData.data) {
      const binaryData = Utilities.base64Decode(badgeImageData.data);
      badgeImageBlob = Utilities.newBlob(binaryData, badgeImageData.type, badgeImageData.name);
      badgeUrlPattern = this.identifyBadgeUrlPattern(slides);
    }
    
    // Process each slide
    slides.forEach((slide, slideIndex) => {
      if (this.isPracticeSlide(slide)) {
        const practiceReplacements = this.replacePracticeSlideContent(slide, mappings);
        totalReplacements += practiceReplacements;
      } else {
        const replacements = this.replaceTextInSlide(slide, mappings);
        totalReplacements += replacements;
      }
      
      if (badgeImageBlob) {
        const imageReplacements = this.replaceBadgeImage(slide, badgeImageBlob, badgeUrlPattern);
        totalImageReplacements += imageReplacements;
      }
    });
    
    return {
      url: newPresentation.getUrl(),
      name: newPresentationName,
      replacements: totalReplacements,
      imageReplacements: totalImageReplacements
    };
  },
  
  // Create data mappings from sheet
  createDataMappings: function(sheet) {
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    const mappings = {};
    
    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      const tag = row[0];
      const japaneseValue = row[1];
      const englishValue = row[2];
      
      if (!tag || tag.trim() === '') continue;
      if (tag === 'Header' || tag === 'A') continue;
      
      // Handle Title in Japanese
      if (tag === 'Title') {
        if (japaneseValue && japaneseValue.trim() !== '') {
          mappings[tag] = japaneseValue;
        }
      }
      // Handle Practice sentences
      else if (tag.toLowerCase().includes('practice')) {
        if (englishValue && englishValue.trim() !== '') {
          mappings[tag] = englishValue;
        }
      }
      // Handle all other content in English
      else {
        if (englishValue && englishValue.trim() !== '') {
          mappings[tag] = englishValue;
        }
      }
    }
    
    console.log('Created mappings:', mappings);
    return mappings;
  },
  
  // Check if a slide is a practice slide
  isPracticeSlide: function(slide) {
    const pageElements = slide.getPageElements();
    
    for (let element of pageElements) {
      try {
        if (element.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
          const shape = element.asShape();
          if (shape.getText) {
            const text = shape.getText().asString().toLowerCase();
            if (text.includes('practice') && (text.includes('old') || text.includes('new'))) {
              return true;
            }
          }
        }
      } catch (error) {
        // Skip if can't access text
      }
    }
    
    return false;
  },
  
  // Replace practice slide content
  replacePracticeSlideContent: function(slide, mappings) {
    const pageElements = slide.getPageElements();
    let replacementCount = 0;
    
    const textShapes = [];
    pageElements.forEach(element => {
      try {
        if (element.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
          const shape = element.asShape();
          if (shape.getText) {
            textShapes.push(shape);
          }
        }
      } catch (error) {
        // Skip if can't access shape
      }
    });
    
    // Sort shapes by vertical position
    textShapes.sort((a, b) => {
      try {
        const aTop = a.getTop();
        const bTop = b.getTop();
        return aTop - bTop;
      } catch (error) {
        return 0;
      }
    });
    
    // Determine practice number
    let practiceNumber = null;
    for (let shape of textShapes) {
      try {
        const text = shape.getText().asString().toLowerCase();
        if (text.includes('practice 1')) practiceNumber = 1;
        else if (text.includes('practice 2')) practiceNumber = 2;
        else if (text.includes('practice 3')) practiceNumber = 3;
        
        if (practiceNumber) break;
      } catch (error) {
        // Continue searching
      }
    }
    
    if (practiceNumber && textShapes.length >= 3) {
      try {
        let oldKey, newKey;
        if (practiceNumber === 1) {
          oldKey = 'Practice 1 old';
          newKey = 'practice 1 new';
        } else {
          oldKey = `Practice ${practiceNumber} old`;
          newKey = `Practice ${practiceNumber} new`;
        }
        
        // First text block: old content
        if (mappings[oldKey]) {
          textShapes[0].getText().setText(mappings[oldKey]);
          replacementCount++;
        }
        
        // Third text block: new content
        if (mappings[newKey]) {
          textShapes[2].getText().setText(mappings[newKey]);
          replacementCount++;
        }
        
      } catch (error) {
        console.log(`Error updating practice slide content: ${error.message}`);
        replacementCount += this.replaceTextInSlide(slide, mappings);
      }
    } else {
      replacementCount += this.replaceTextInSlide(slide, mappings);
    }
    
    return replacementCount;
  },
  
  // Replace text in slide
  replaceTextInSlide: function(slide, mappings) {
    let replacementCount = 0;
    const pageElements = slide.getPageElements();
    
    pageElements.forEach(element => {
      try {
        if (element.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
          const shape = element.asShape();
          if (shape.getText) {
            const textRange = shape.getText();
            replacementCount += this.replaceTextInTextRange(textRange, mappings);
          }
        } else if (element.getPageElementType() === SlidesApp.PageElementType.TABLE) {
          const table = element.asTable();
          const numRows = table.getNumRows();
          const numCols = table.getNumColumns();
          
          for (let row = 0; row < numRows; row++) {
            for (let col = 0; col < numCols; col++) {
              const cell = table.getCell(row, col);
              const textRange = cell.getText();
              replacementCount += this.replaceTextInTextRange(textRange, mappings);
            }
          }
        }
      } catch (error) {
        console.log(`Could not process element: ${error.message}`);
      }
    });
    
    return replacementCount;
  },
  
  // Replace text in text range
  replaceTextInTextRange: function(textRange, mappings) {
    let replacementCount = 0;
    let currentText = textRange.asString();
    
    Object.entries(mappings).forEach(([tag, value]) => {
      const placeholder = `<${tag}>`;
      if (currentText.includes(placeholder)) {
        currentText = currentText.replace(new RegExp(this.escapeRegExp(placeholder), 'g'), value);
        replacementCount++;
      }
    });
    
    if (replacementCount > 0) {
      textRange.setText(currentText);
    }
    
    return replacementCount;
  },
  
  // Helper to escape regex
  escapeRegExp: function(string) {
    return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  },
  
  // Identify badge URL pattern
  identifyBadgeUrlPattern: function(slides) {
    const imageUrlInfo = new Map();
    
    slides.forEach(slide => {
      const pageElements = slide.getPageElements();
      pageElements.forEach(element => {
        if (element.getPageElementType() === SlidesApp.PageElementType.IMAGE) {
          try {
            const image = element.asImage();
            const url = image.getContentUrl();
            const width = image.getWidth();
            const height = image.getHeight();
            
            if (!imageUrlInfo.has(url)) {
              imageUrlInfo.set(url, {
                count: 0,
                width: width,
                height: height,
                isSquare: width === height,
                size: width
              });
            }
            imageUrlInfo.get(url).count++;
          } catch (e) {
            // Skip if can't get URL
          }
        }
      });
    });
    
    let badgeUrl = null;
    let maxScore = 0;
    
    imageUrlInfo.forEach((info, url) => {
      let score = 0;
      
      if (info.width === 225 && info.height === 225) {
        score += 10;
      }
      
      if (info.isSquare) {
        score += 3;
      }
      
      if (info.count > 1) {
        score += info.count * 2;
      }
      
      if (url.includes('AGV_vUdpdReIOhhJCvPyXR0pXjZGCEVzOe4ZwQ6w')) {
        score += 5;
      }
      
      if (score > maxScore) {
        maxScore = score;
        badgeUrl = url;
      }
    });
    
    return badgeUrl;
  },
  
  // Replace badge image
  replaceBadgeImage: function(slide, newImageBlob, badgeUrlPattern) {
    let replacements = 0;
    const pageElements = slide.getPageElements();
    
    pageElements.forEach(element => {
      try {
        if (element.getPageElementType() === SlidesApp.PageElementType.IMAGE) {
          const image = element.asImage();
          let isBadgeImage = false;
          
          if (badgeUrlPattern) {
            try {
              const imageUrl = image.getContentUrl();
              if (imageUrl === badgeUrlPattern) {
                isBadgeImage = true;
              }
            } catch (e) {
              // URL might not be available
            }
          }
          
          if (!isBadgeImage) {
            const width = image.getWidth();
            const height = image.getHeight();
            
            if (width === 225 && height === 225) {
              isBadgeImage = true;
            }
          }
          
          if (isBadgeImage) {
            const transform = image.getTransform();
            const width = image.getWidth();
            const height = image.getHeight();
            const x = transform.getTranslateX();
            const y = transform.getTranslateY();
            
            element.remove();
            
            const newImage = slide.insertImage(newImageBlob);
            newImage.setLeft(x);
            newImage.setTop(y);
            newImage.setWidth(width);
            newImage.setHeight(height);
            
            replacements++;
          }
        }
      } catch (error) {
        console.log(`Could not process image element: ${error.message}`);
      }
    });
    
    return replacements;
  }
};

// Export functions for global access
function processSheet(spreadsheetUrl, templateId, badgeImageData) {
  return SheetsToSlides.processSheet(spreadsheetUrl, templateId, badgeImageData);
}

function getEKSheets() {
  return SharedServices.getEKSpreadsheets();
}
