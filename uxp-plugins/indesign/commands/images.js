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
 * Place an image on a page
 */
const placeImage = async (command) => {
    const doc = app.activeDocument;
    const options = command.options;
    
    // Validate page index
    if (options.pageIndex >= doc.pages.length) {
        throw new Error(`Page index ${options.pageIndex} out of range. Document has ${doc.pages.length} pages.`);
    }
    
    const page = doc.pages.item(options.pageIndex);
    
    // Create a rectangle frame for the image
    let width = options.width || 200;
    let height = options.height || 200;
    
    const bounds = [
        options.y,
        options.x,
        options.y + height,
        options.x + width
    ];
    
    const imageFrame = page.rectangles.add({
        geometricBounds: bounds
    });
    
    // Place the image file
    try {
        const placedImage = imageFrame.place(new File(options.imagePath));
        
        // Apply fit option
        const fitOptions = {
            "PROPORTIONALLY": () => imageFrame.fit(1718185072), // FitOptions.PROPORTIONALLY
            "FILL_PROPORTIONALLY": () => imageFrame.fit(1718185070), // FitOptions.FILL_PROPORTIONALLY
            "FIT_CONTENT": () => imageFrame.fit(1718187107), // FitOptions.CONTENT_TO_FRAME
            "CENTER_CONTENT": () => imageFrame.fit(1718186339), // FitOptions.CENTER_CONTENT
            "FRAME_TO_CONTENT": () => imageFrame.fit(1718187107) // FitOptions.FRAME_TO_CONTENT
        };
        
        if (options.fitOption && fitOptions[options.fitOption]) {
            fitOptions[options.fitOption]();
        } else {
            // Default to proportional fit
            imageFrame.fit(1718185072);
        }
        
        return {
            frameId: imageFrame.id.toString(),
            imagePath: options.imagePath,
            bounds: {
                x: options.x,
                y: options.y,
                width: width,
                height: height
            },
            pageIndex: options.pageIndex
        };
    } catch (e) {
        // Clean up the frame if image placement failed
        imageFrame.remove();
        throw new Error(`Failed to place image: ${e.message}`);
    }
};

/**
 * Get information about all placed images in the document
 */
const getImages = async (command) => {
    const doc = app.activeDocument;
    const images = [];
    
    // Iterate through all pages
    for (let p = 0; p < doc.pages.length; p++) {
        const page = doc.pages.item(p);
        
        // Check rectangles for placed images
        for (let i = 0; i < page.rectangles.length; i++) {
            const rect = page.rectangles.item(i);
            
            // Check if rectangle contains an image
            if (rect.images && rect.images.length > 0) {
                for (let j = 0; j < rect.images.length; j++) {
                    const img = rect.images.item(j);
                    const bounds = rect.geometricBounds;
                    
                    images.push({
                        frameId: rect.id.toString(),
                        imageId: img.id.toString(),
                        filePath: img.itemLink ? img.itemLink.filePath : null,
                        pageIndex: p,
                        bounds: {
                            y1: bounds[0],
                            x1: bounds[1],
                            y2: bounds[2],
                            x2: bounds[3],
                            width: bounds[3] - bounds[1],
                            height: bounds[2] - bounds[0]
                        }
                    });
                }
            }
        }
        
        // Also check graphics frames
        for (let i = 0; i < page.allGraphics.length; i++) {
            const graphic = page.allGraphics.item(i);
            const parent = graphic.parent;
            
            if (parent && parent.geometricBounds) {
                const bounds = parent.geometricBounds;
                
                // Avoid duplicates
                const exists = images.some(img => img.imageId === graphic.id.toString());
                if (!exists) {
                    images.push({
                        frameId: parent.id ? parent.id.toString() : null,
                        imageId: graphic.id.toString(),
                        filePath: graphic.itemLink ? graphic.itemLink.filePath : null,
                        pageIndex: p,
                        bounds: {
                            y1: bounds[0],
                            x1: bounds[1],
                            y2: bounds[2],
                            x2: bounds[3],
                            width: bounds[3] - bounds[1],
                            height: bounds[2] - bounds[0]
                        }
                    });
                }
            }
        }
    }
    
    return images;
};

/**
 * Place an image from Word document import with auto-positioning
 */
const placeImageFromWord = async (command) => {
    const doc = app.activeDocument;
    const options = command.options;
    
    // For auto-positioning, find a suitable location
    // This is a simplified implementation - could be enhanced with more sophisticated layout logic
    
    let pageIndex = 0;
    let x = 72; // 1 inch margin
    let y = 72;
    
    if (options.autoPosition) {
        // Find the last page with content
        pageIndex = Math.max(0, doc.pages.length - 1);
        
        // Position below existing content or at a default location
        const page = doc.pages.item(pageIndex);
        const pageHeight = page.bounds[2] - page.bounds[0];
        const pageWidth = page.bounds[3] - page.bounds[1];
        
        // Default to center of page
        x = pageWidth / 4;
        y = pageHeight / 2;
    }
    
    // Use the standard placeImage function
    return placeImage({
        options: {
            imagePath: options.imagePath,
            pageIndex: pageIndex,
            x: x,
            y: y,
            width: options.width || 300,
            height: options.height || 200,
            fitOption: "PROPORTIONALLY"
        }
    });
};

module.exports = {
    placeImage,
    getImages,
    placeImageFromWord
};
