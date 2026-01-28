/* MIT License
 *
 * Copyright (c) 2025 Adobe MCP Contributors
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in all
 * copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
 * SOFTWARE.
 */

const { app } = require("indesign");

/**
 * Layout Word document content into InDesign
 * This is the main orchestration function for Word import
 */
const layoutWordContent = async (command) => {
    const doc = app.activeDocument;
    const options = command.options;
    
    const content = options.content || [];
    const styleMapping = options.styleMapping || {};
    const autoFlow = options.autoFlow !== false;
    const startPage = options.startPage || 0;
    
    let currentPageIndex = startPage;
    let pagesCreated = 0;
    let paragraphsPlaced = 0;
    let currentTextFrame = null;
    let linkedFrames = [];
    
    // Get page dimensions for text frame creation
    const getPageContentBounds = (pageIndex) => {
        const page = doc.pages[pageIndex];
        const pageBounds = page.bounds;
        const margins = page.marginPreferences;
        
        return {
            x: margins.left,
            y: margins.top,
            width: (pageBounds[3] - pageBounds[1]) - margins.left - margins.right,
            height: (pageBounds[2] - pageBounds[0]) - margins.top - margins.bottom
        };
    };
    
    // Ensure we have enough pages
    const ensurePage = (pageIndex) => {
        while (doc.pages.length <= pageIndex) {
            doc.pages.add();
            pagesCreated++;
        }
        return doc.pages[pageIndex];
    };
    
    // Create or get a text frame for content
    const getOrCreateTextFrame = (pageIndex) => {
        ensurePage(pageIndex);
        const bounds = getPageContentBounds(pageIndex);
        const page = doc.pages[pageIndex];
        
        const textFrame = page.textFrames.add({
            geometricBounds: [
                bounds.y,
                bounds.x,
                bounds.y + bounds.height,
                bounds.x + bounds.width
            ]
        });
        
        return textFrame;
    };
    
    // Get or create InDesign paragraph style
    const getOrCreateStyle = (wordStyleName) => {
        const indesignStyleName = styleMapping[wordStyleName] || wordStyleName;
        
        // Try to find existing style
        try {
            const existingStyle = doc.paragraphStyles.itemByName(indesignStyleName);
            if (existingStyle.isValid) {
                return existingStyle;
            }
        } catch (e) {
            // Style doesn't exist
        }
        
        // Create a basic style with the name
        try {
            const newStyle = doc.paragraphStyles.add({
                name: indesignStyleName
            });
            return newStyle;
        } catch (e) {
            // Return Basic Paragraph as fallback
            return doc.paragraphStyles.itemByName("[Basic Paragraph]");
        }
    };
    
    // Create initial text frame
    currentTextFrame = getOrCreateTextFrame(currentPageIndex);
    linkedFrames.push(currentTextFrame);
    
    // Process each content block
    for (let i = 0; i < content.length; i++) {
        const block = content[i];
        
        if (block.type === "paragraph") {
            // Build paragraph text from runs
            let paragraphText = "";
            if (block.runs && block.runs.length > 0) {
                for (const run of block.runs) {
                    paragraphText += run.text || "";
                }
            }
            
            // Add paragraph separator
            if (paragraphText.length > 0) {
                paragraphText += "\r"; // InDesign paragraph break
            } else if (block.style) {
                // Empty paragraph with style (intentional spacing)
                paragraphText = "\r";
            } else {
                continue; // Skip completely empty blocks
            }
            
            // Insert text at end of current frame's story
            const insertionPoint = currentTextFrame.parentStory.insertionPoints[-1];
            const startIndex = currentTextFrame.parentStory.characters.length;
            insertionPoint.contents = paragraphText;
            
            // Apply paragraph style
            if (block.style) {
                const style = getOrCreateStyle(block.style);
                try {
                    // Get the paragraph we just inserted
                    const story = currentTextFrame.parentStory;
                    const lastPara = story.paragraphs[-1];
                    if (lastPara) {
                        lastPara.appliedParagraphStyle = style;
                    }
                } catch (e) {
                    console.log(`Failed to apply style: ${e}`);
                }
            }
            
            // Apply inline formatting from runs
            if (block.runs && block.runs.length > 0) {
                let charIndex = startIndex;
                for (const run of block.runs) {
                    if (run.text && run.formatting) {
                        const runLength = run.text.length;
                        try {
                            const story = currentTextFrame.parentStory;
                            for (let c = charIndex; c < charIndex + runLength && c < story.characters.length; c++) {
                                const char = story.characters[c];
                                
                                if (run.formatting.bold) {
                                    char.fontStyle = "Bold";
                                }
                                if (run.formatting.italic) {
                                    char.fontStyle = run.formatting.bold ? "Bold Italic" : "Italic";
                                }
                                if (run.formatting.underline) {
                                    char.underline = true;
                                }
                            }
                        } catch (e) {
                            // Formatting application failed, continue
                        }
                        charIndex += runLength;
                    }
                }
            }
            
            paragraphsPlaced++;
            
            // Check for overflow and create new linked frame if needed
            if (autoFlow && currentTextFrame.overflows) {
                currentPageIndex++;
                const newFrame = getOrCreateTextFrame(currentPageIndex);
                currentTextFrame.nextTextFrame = newFrame;
                linkedFrames.push(newFrame);
                currentTextFrame = newFrame;
            }
        } else if (block.type === "table") {
            // Basic table support - convert to text for now
            // Full table support would require more complex implementation
            let tableText = "";
            for (const row of block.rows || []) {
                const cellTexts = [];
                for (const cell of row) {
                    let cellContent = "";
                    for (const para of cell) {
                        if (para.runs) {
                            for (const run of para.runs) {
                                cellContent += run.text || "";
                            }
                        }
                    }
                    cellTexts.push(cellContent);
                }
                tableText += cellTexts.join("\t") + "\r";
            }
            
            if (tableText) {
                const insertionPoint = currentTextFrame.parentStory.insertionPoints[-1];
                insertionPoint.contents = tableText;
                paragraphsPlaced++;
            }
        }
    }
    
    return {
        pagesCreated: pagesCreated,
        paragraphsPlaced: paragraphsPlaced,
        framesCreated: linkedFrames.length,
        startPage: startPage,
        endPage: currentPageIndex
    };
};

module.exports = {
    layoutWordContent
};
