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
 * Open an existing InDesign document
 */
const openDocument = async (command) => {
    const options = command.options;
    
    const file = new File(options.filePath);
    if (!file.exists) {
        throw new Error(`File not found: ${options.filePath}`);
    }
    
    const doc = app.open(file);
    
    return {
        documentId: doc.id.toString(),
        name: doc.name,
        filePath: doc.filePath ? doc.filePath.fsName : options.filePath,
        pageCount: doc.pages.length,
        documentPreferences: {
            pageWidth: doc.documentPreferences.pageWidth,
            pageHeight: doc.documentPreferences.pageHeight,
            facingPages: doc.documentPreferences.facingPages
        }
    };
};

/**
 * Save the active document
 */
const saveDocument = async (command) => {
    const doc = app.activeDocument;
    const options = command.options;
    
    if (options.filePath) {
        const file = new File(options.filePath);
        doc.save(file);
        return {
            saved: true,
            filePath: options.filePath
        };
    } else {
        if (doc.saved) {
            doc.save();
            return {
                saved: true,
                filePath: doc.filePath ? doc.filePath.fsName : null
            };
        } else {
            throw new Error("Document has never been saved. Please provide a file path.");
        }
    }
};

/**
 * Export the document as PDF
 */
const exportPdf = async (command) => {
    const doc = app.activeDocument;
    const options = command.options;
    
    const file = new File(options.filePath);
    
    // Set PDF export preferences
    if (options.preset) {
        try {
            const preset = app.pdfExportPresets.itemByName(options.preset);
            if (preset.isValid) {
                app.pdfExportPreferences.pdfExportPreset = preset;
            }
        } catch (e) {
            console.log(`PDF preset '${options.preset}' not found, using default`);
        }
    }
    
    // Export to PDF
    doc.exportFile(1952403524, file); // ExportFormat.PDF_TYPE
    
    return {
        exported: true,
        filePath: options.filePath,
        preset: options.preset || "default"
    };
};

module.exports = {
    openDocument,
    saveDocument,
    exportPdf
};
