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
 * Get all paragraph styles from the active document
 */
const getParagraphStyles = async (command) => {
    const doc = app.activeDocument;
    const styles = [];
    
    for (let i = 0; i < doc.paragraphStyles.length; i++) {
        const style = doc.paragraphStyles[i];
        styles.push({
            name: style.name,
            id: style.id,
            basedOn: style.basedOn ? style.basedOn.name : null,
            fontFamily: style.appliedFont ? style.appliedFont.name : null,
            fontSize: style.pointSize,
            leading: style.leading,
            alignment: style.justification ? style.justification.toString() : null,
            spaceBefore: style.spaceBefore,
            spaceAfter: style.spaceAfter,
            firstLineIndent: style.firstLineIndent,
            leftIndent: style.leftIndent,
            rightIndent: style.rightIndent
        });
    }
    
    return styles;
};

/**
 * Create a new paragraph style
 */
const createParagraphStyle = async (command) => {
    const doc = app.activeDocument;
    const options = command.options;
    
    // Check if style already exists
    let existingStyle = null;
    try {
        existingStyle = doc.paragraphStyles.itemByName(options.styleName);
        if (existingStyle.isValid) {
            return {
                status: "exists",
                message: `Style '${options.styleName}' already exists`,
                styleId: existingStyle.id
            };
        }
    } catch (e) {
        // Style doesn't exist, continue to create
    }
    
    // Create the new style
    const newStyle = doc.paragraphStyles.add({
        name: options.styleName
    });
    
    // Apply basedOn if specified
    if (options.basedOn) {
        try {
            const baseStyle = doc.paragraphStyles.itemByName(options.basedOn);
            if (baseStyle.isValid) {
                newStyle.basedOn = baseStyle;
            }
        } catch (e) {
            console.log(`Base style '${options.basedOn}' not found`);
        }
    }
    
    // Apply font family
    if (options.fontFamily) {
        try {
            newStyle.appliedFont = options.fontFamily;
        } catch (e) {
            console.log(`Font '${options.fontFamily}' not found`);
        }
    }
    
    // Apply font size
    if (options.fontSize !== undefined) {
        newStyle.pointSize = options.fontSize;
    }
    
    // Apply leading
    if (options.leading !== undefined) {
        if (options.leading === 0) {
            newStyle.leading = "Auto";
        } else {
            newStyle.leading = options.leading;
        }
    }
    
    // Apply alignment
    if (options.alignment) {
        const alignmentMap = {
            "LEFT": 1818584692, // Justification.LEFT_ALIGN
            "CENTER": 1667591796, // Justification.CENTER_ALIGN
            "RIGHT": 1919379572, // Justification.RIGHT_ALIGN
            "JUSTIFY": 1818915700, // Justification.LEFT_JUSTIFIED
            "FULLY_JUSTIFIED": 1718971500 // Justification.FULLY_JUSTIFIED
        };
        if (alignmentMap[options.alignment]) {
            newStyle.justification = alignmentMap[options.alignment];
        }
    }
    
    // Apply spacing
    if (options.spaceBefore !== undefined) {
        newStyle.spaceBefore = options.spaceBefore;
    }
    if (options.spaceAfter !== undefined) {
        newStyle.spaceAfter = options.spaceAfter;
    }
    
    // Apply indents
    if (options.firstLineIndent !== undefined) {
        newStyle.firstLineIndent = options.firstLineIndent;
    }
    if (options.leftIndent !== undefined) {
        newStyle.leftIndent = options.leftIndent;
    }
    if (options.rightIndent !== undefined) {
        newStyle.rightIndent = options.rightIndent;
    }
    
    // Apply color
    if (options.color) {
        try {
            // Create or find RGB color
            const colorName = `RGB_${options.color.red}_${options.color.green}_${options.color.blue}`;
            let color;
            try {
                color = doc.colors.itemByName(colorName);
                if (!color.isValid) {
                    throw new Error("Color not found");
                }
            } catch (e) {
                color = doc.colors.add({
                    name: colorName,
                    model: 1380401491, // ColorModel.PROCESS
                    space: 1666336578, // ColorSpace.RGB
                    colorValue: [options.color.red, options.color.green, options.color.blue]
                });
            }
            newStyle.fillColor = color;
        } catch (e) {
            console.log(`Failed to apply color: ${e}`);
        }
    }
    
    return {
        status: "created",
        styleName: options.styleName,
        styleId: newStyle.id
    };
};

/**
 * Apply a paragraph style to text
 */
const applyParagraphStyle = async (command) => {
    const doc = app.activeDocument;
    const options = command.options;
    
    // Find the style
    let style;
    try {
        style = doc.paragraphStyles.itemByName(options.styleName);
        if (!style.isValid) {
            throw new Error(`Style '${options.styleName}' not found`);
        }
    } catch (e) {
        throw new Error(`Style '${options.styleName}' not found`);
    }
    
    // Get the story
    if (options.storyIndex >= doc.stories.length) {
        throw new Error(`Story index ${options.storyIndex} out of range`);
    }
    const story = doc.stories[options.storyIndex];
    
    // Apply to specific paragraph or all
    if (options.paragraphIndex !== undefined) {
        if (options.paragraphIndex >= story.paragraphs.length) {
            throw new Error(`Paragraph index ${options.paragraphIndex} out of range`);
        }
        story.paragraphs[options.paragraphIndex].appliedParagraphStyle = style;
        return {
            applied: true,
            paragraphIndex: options.paragraphIndex
        };
    } else {
        // Apply to all paragraphs
        let count = 0;
        for (let i = 0; i < story.paragraphs.length; i++) {
            story.paragraphs[i].appliedParagraphStyle = style;
            count++;
        }
        return {
            applied: true,
            paragraphCount: count
        };
    }
};

module.exports = {
    getParagraphStyles,
    createParagraphStyle,
    applyParagraphStyle
};
