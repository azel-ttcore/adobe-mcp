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
 * Create a new text frame on a page
 */
const createTextFrame = async (command) => {
    const doc = app.activeDocument;
    const options = command.options;
    
    // Validate page index
    if (options.pageIndex >= doc.pages.length) {
        throw new Error(`Page index ${options.pageIndex} out of range. Document has ${doc.pages.length} pages.`);
    }
    
    const page = doc.pages[options.pageIndex];
    
    // Create the text frame with geometric bounds [y1, x1, y2, x2]
    const bounds = [
        options.y,
        options.x,
        options.y + options.height,
        options.x + options.width
    ];
    
    const textFrame = page.textFrames.add({
        geometricBounds: bounds
    });
    
    // Add content if provided
    if (options.content) {
        textFrame.contents = options.content;
    }
    
    return {
        frameId: textFrame.id.toString(),
        bounds: {
            x: options.x,
            y: options.y,
            width: options.width,
            height: options.height
        },
        pageIndex: options.pageIndex
    };
};

/**
 * Get information about text frames in the document
 */
const getTextFrames = async (command) => {
    const doc = app.activeDocument;
    const options = command.options;
    const frames = [];
    
    const processPage = (page, pageIndex) => {
        for (let i = 0; i < page.textFrames.length; i++) {
            const frame = page.textFrames[i];
            const bounds = frame.geometricBounds;
            
            frames.push({
                id: frame.id.toString(),
                pageIndex: pageIndex,
                bounds: {
                    y1: bounds[0],
                    x1: bounds[1],
                    y2: bounds[2],
                    x2: bounds[3],
                    width: bounds[3] - bounds[1],
                    height: bounds[2] - bounds[0]
                },
                contentLength: frame.contents ? frame.contents.length : 0,
                overflows: frame.overflows,
                hasNextFrame: frame.nextTextFrame !== null,
                hasPreviousFrame: frame.previousTextFrame !== null
            });
        }
    };
    
    if (options.pageIndex !== undefined) {
        if (options.pageIndex >= doc.pages.length) {
            throw new Error(`Page index ${options.pageIndex} out of range`);
        }
        processPage(doc.pages[options.pageIndex], options.pageIndex);
    } else {
        for (let p = 0; p < doc.pages.length; p++) {
            processPage(doc.pages[p], p);
        }
    }
    
    return frames;
};

/**
 * Link two text frames for text flow
 */
const linkTextFrames = async (command) => {
    const doc = app.activeDocument;
    const options = command.options;
    
    // Find source frame
    let sourceFrame = null;
    let targetFrame = null;
    
    for (let p = 0; p < doc.pages.length; p++) {
        const page = doc.pages[p];
        for (let i = 0; i < page.textFrames.length; i++) {
            const frame = page.textFrames[i];
            if (frame.id.toString() === options.sourceFrameId) {
                sourceFrame = frame;
            }
            if (frame.id.toString() === options.targetFrameId) {
                targetFrame = frame;
            }
        }
    }
    
    if (!sourceFrame) {
        throw new Error(`Source frame with ID ${options.sourceFrameId} not found`);
    }
    if (!targetFrame) {
        throw new Error(`Target frame with ID ${options.targetFrameId} not found`);
    }
    
    // Link the frames
    sourceFrame.nextTextFrame = targetFrame;
    
    return {
        linked: true,
        sourceFrameId: options.sourceFrameId,
        targetFrameId: options.targetFrameId
    };
};

/**
 * Insert text into a text frame
 */
const insertText = async (command) => {
    const doc = app.activeDocument;
    const options = command.options;
    
    // Find the frame
    let textFrame = null;
    for (let p = 0; p < doc.pages.length; p++) {
        const page = doc.pages[p];
        for (let i = 0; i < page.textFrames.length; i++) {
            const frame = page.textFrames[i];
            if (frame.id.toString() === options.frameId) {
                textFrame = frame;
                break;
            }
        }
        if (textFrame) break;
    }
    
    if (!textFrame) {
        throw new Error(`Text frame with ID ${options.frameId} not found`);
    }
    
    const position = options.position || "END";
    let insertionPoint;
    
    switch (position) {
        case "START":
            insertionPoint = textFrame.insertionPoints[0];
            insertionPoint.contents = options.text;
            break;
        case "END":
            insertionPoint = textFrame.insertionPoints[-1];
            insertionPoint.contents = options.text;
            break;
        case "REPLACE":
            textFrame.contents = options.text;
            break;
        default:
            throw new Error(`Invalid position: ${position}`);
    }
    
    // Apply style if specified
    if (options.styleName) {
        try {
            const style = doc.paragraphStyles.itemByName(options.styleName);
            if (style.isValid) {
                // Apply to the inserted text
                const story = textFrame.parentStory;
                if (story && story.paragraphs.length > 0) {
                    const lastPara = story.paragraphs[-1];
                    lastPara.appliedParagraphStyle = style;
                }
            }
        } catch (e) {
            console.log(`Style '${options.styleName}' not found`);
        }
    }
    
    return {
        inserted: true,
        frameId: options.frameId,
        textLength: options.text.length,
        position: position
    };
};

module.exports = {
    createTextFrame,
    getTextFrames,
    linkTextFrames,
    insertText
};
