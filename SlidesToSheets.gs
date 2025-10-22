/**
 * Module 3: Slides to Sheets Extractor
 * Extracts text from slides to predetermined spreadsheet cells
 */

const SlidesToSheets = {
  // Position mapping for extraction
  POSITION_MAPPING: {
    "7-4": "B8",   // Slide 7, 4th text element → B8
    "7-5": "B9",   // Slide 7, 5th text element → B9
    "8-4": "C8",   // Slide 8, 4th text element → C8
    "8-5": "C9",   // Slide 8, 5th text element → C9
    "9-4": "B11",  // Slide 9, 4th text element → B11
    "9-5": "B12",  // Slide 9, 5th text element → B12
    "10-4": "C11", // Slide 10, 4th text element → C11
    "10-5": "C12", // Slide 10, 5th text element → C12
    "11-4": "B15", // Slide 11, 4th text element → B15
    "11-5": "B16", // Slide 11, 5th text element → B16
    "12-4": "C15", // Slide 12, 4th text element → C15
    "12-5": "C16", // Slide 12, 5th text element → C16
    "13-4": "B18", // Slide 13, 4th text element → B18
    "13-5": "B19", // Slide 13, 5th text element → B19
    "14-4": "C18", // Slide 14, 4th text element → C18
    "14-5": "C19", // Slide 14, 5th text element → C19
    "15-4": "B22", // Slide 15, 4th text element → B22
    "15-5": "B23", // Slide 15, 5th text element → B23
    "16-4": "C22", // Slide 16, 4th text element → C22
    "16-5": "C23", // Slide 16, 5th text element → C23
    "17-4": "B25", // Slide 17, 4th text element → B25
    "17-5": "B26", // Slide 17, 5th text element → B26
    "18-4": "C25", // Slide 18, 4th text element → C25
    "18-5": "C26", // Slide 18, 5th text element → C26
  },
  
  // Main extraction function
extractWithMapping: function(presentationUrl, spreadsheetUrl) {
  try {
    const presentationId = SharedServices.extractId(presentationUrl);
    const presentation = SlidesApp.openById(presentationId);
    
    // Get or find the spreadsheet
    let spreadsheetId;
    if (spreadsheetUrl) {
      spreadsheetId = SharedServices.extractId(spreadsheetUrl);
    } else {
      // Try to find matching spreadsheet by name
      const presentationName = presentation.getName();
      const ekMatch = presentationName.match(/ek-\d+/i);
      if (ekMatch) {
        const ekNumber = ekMatch[0];
        const files = DriveApp.searchFiles(
          `title contains "${ekNumber}" and mimeType = "application/vnd.google-apps.spreadsheet"`
        );
        if (files.hasNext()) {
          spreadsheetId = files.next().getId();
        }
      }
    }
    
    if (!spreadsheetId) {
      return {success: false, error: 'Could not find spreadsheet'};
    }
    
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getActiveSheet();
    
    // Get texts from slides organized by slide number and position
    const slideTexts = this.getTextsFromSlidesWithPositions(presentation);
    
    // ADD DIAGNOSTIC LOGGING HERE
    console.log('Total slides with text:', Object.keys(slideTexts).length);
    console.log('Slide numbers with text:', Object.keys(slideTexts));
    console.log('Type of slide keys:', typeof Object.keys(slideTexts)[0]);
    
    // Apply the position-based mapping
    let updatedCells = 0;
    const updates = [];
    
    Object.entries(this.POSITION_MAPPING).forEach(([position, cellLocation]) => {
      const [slideNum, textPos] = position.split('-').map(Number);
      
      console.log(`\nProcessing mapping: ${position} -> ${cellLocation}`);
      console.log(`  Looking for: Slide ${slideNum}, Position ${textPos}`);
      
      // Try both number and string keys
      const slideContent = slideTexts[slideNum] || slideTexts[String(slideNum)];
      
      if (slideContent) {
        console.log(`  Slide ${slideNum} has ${slideContent.length} text elements`);
        if (slideContent[textPos - 1]) {
          const text = slideContent[textPos - 1];
          console.log(`  Found text: "${text.substring(0, 50)}..."`);
          
          try {
            // Extract column and row from cell location
            const col = cellLocation.charCodeAt(0) - 64; // A=1, B=2, C=3
            const row = parseInt(cellLocation.substring(1));
            
            console.log(`  Writing to cell ${cellLocation} (row ${row}, col ${col})`);
            
            // Remove bullet point if present
            let valueToSet = text;
            if (valueToSet.startsWith('・')) {
              valueToSet = valueToSet.substring(1);
            }
            
            sheet.getRange(row, col).setValue(valueToSet);
            updatedCells++;
            console.log(`  ✓ Successfully updated cell ${cellLocation}`);
            
            updates.push({
              slide: slideNum,
              position: textPos,
              text: text,
              cell: cellLocation
            });
          } catch (cellError) {
            console.error(`  ✗ Error updating cell ${cellLocation}:`, cellError);
          }
        } else {
          console.log(`  No text at position ${textPos} (array index ${textPos - 1})`);
          if (slideContent.length > 0) {
            console.log(`  Available positions: 1-${slideContent.length}`);
          }
        }
      } else {
        console.log(`  Slide ${slideNum} not found in extracted texts`);
      }
    });
    
    console.log(`\nFinal result: Updated ${updatedCells} cells`);

    SpreadsheetApp.flush();
    
    return {
      success: true,
      updatedCells: updatedCells,
      updates: updates,
      spreadsheetUrl: spreadsheet.getUrl()
    };
    
  } catch (error) {
    console.error('Error:', error);
    return {success: false, error: error.message};
  }
},
  
  // Get texts from all slides with positions
  getTextsFromSlidesWithPositions: function(presentation) {
    const slideTexts = {};
    const slides = presentation.getSlides();
    
    slides.forEach((slide, index) => {
      const slideNumber = index + 1;
      const texts = this.getTextsFromSlide(slide);
      if (texts.length > 0) {
        slideTexts[slideNumber] = texts;
      }
    });
    
    return slideTexts;
  },
  
  // Get all texts from a slide in order
  getTextsFromSlide: function(slide) {
    const texts = [];
    const elements = slide.getPageElements();
    
    // Convert to array and sort by position
    const sortedElements = [];
    for (let i = 0; i < elements.length; i++) {
      sortedElements.push(elements[i]);
    }
    
    sortedElements.sort((a, b) => {
      try {
        const aTop = a.getTop();
        const bTop = b.getTop();
        const aLeft = a.getLeft();
        const bLeft = b.getLeft();
        
        // First sort by vertical position (with tolerance)
        if (Math.abs(aTop - bTop) > 10) {
          return aTop - bTop;
        }
        // If roughly same height, sort by horizontal position
        return aLeft - bLeft;
      } catch (e) {
        return 0;
      }
    });
    
    sortedElements.forEach(element => {
      if (element.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
        try {
          const shape = element.asShape();
          const text = shape.getText().asString().trim();
          if (text && text.length > 0) {
            texts.push(text);
          }
        } catch (e) {
          // Skip if can't get text
        }
      }
    });
    
    return texts;
  },
  
  // View slide contents with positions
  viewSlideContents: function(presentationUrl) {
    try {
      const presentationId = SharedServices.extractId(presentationUrl);
      const presentation = SlidesApp.openById(presentationId);
      const slideTexts = this.getTextsFromSlidesWithPositions(presentation);
      
      const result = [];
      
      Object.entries(slideTexts).forEach(([slideNum, texts]) => {
        texts.forEach((text, index) => {
          const position = `${slideNum}-${index + 1}`;
          result.push({
            slide: parseInt(slideNum),
            position: index + 1,
            text: text,
            mapped: this.POSITION_MAPPING[position] || null
          });
        });
      });
      
      return {success: true, data: result};
    } catch (error) {
      return {success: false, error: error.message};
    }
  },
  
  // Get the current mapping
  getMapping: function() {
    return Object.entries(this.POSITION_MAPPING).map(([position, cell]) => {
      const [slide, pos] = position.split('-');
      return {
        text: `Slide ${slide}, Position ${pos}`,
        cell: cell
      };
    });
  }
};

// Export functions for global access
function extractWithMapping(presentationUrl, spreadsheetUrl) {
  return SlidesToSheets.extractWithMapping(presentationUrl, spreadsheetUrl);
}

function viewSlideContents(presentationUrl) {
  return SlidesToSheets.viewSlideContents(presentationUrl);
}

function getMapping() {
  return SlidesToSheets.getMapping();
}

function getPresentations() {
  return SharedServices.getPresentations();
}

function getEKSpreadsheets() {
  return SharedServices.getEKSpreadsheets();
}
