/* MIT License
 *
 * Copyright (c) 2025 Mike Chambers
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

//const fs = require("uxp").storage.localFileSystem;
//const openfs = require('fs')
const {app, DocumentIntentOptions} = require("indesign");

// Import command modules
const { getParagraphStyles, createParagraphStyle, applyParagraphStyle } = require("./styles.js");
const { createTextFrame, getTextFrames, linkTextFrames, insertText } = require("./textframes.js");
const { placeImage, getImages, placeImageFromWord } = require("./images.js");
const { addPage, getPages, getMasterPages } = require("./pages.js");
const { openDocument, saveDocument, exportPdf } = require("./documents.js");
const { layoutWordContent } = require("./layout.js");


const createDocument = async (command) => {
    console.log("createDocument")

    const options = command.options

    let documents = app.documents
    let margins = options.margins

    let unit = getUnitForIntent(DocumentIntentOptions.WEB_INTENT)

    app.marginPreferences.bottom = `${margins.bottom}${unit}`
    app.marginPreferences.top = `${margins.top}${unit}`
    app.marginPreferences.left = `${margins.left}${unit}`
    app.marginPreferences.right = `${margins.right}${unit}`

    app.marginPreferences.columnCount = options.columns.count
    app.marginPreferences.columnGutter = `${options.columns.gutter}${unit}`
    

    let documentPreferences = {
        pageWidth: `${options.pageWidth}${unit}`,
        pageHeight: `${options.pageHeight}${unit}`,
        pagesPerDocument: options.pagesPerDocument,
        facingPages: options.facingPages,
        intent: DocumentIntentOptions.WEB_INTENT
    }

    const showingWindow = true
    //Boolean showingWindow, DocumentPreset documentPreset, Object withProperties 
    documents.add({showingWindow, documentPreferences})
}


const getUnitForIntent = (intent) => {

    if(intent && intent.toString() === DocumentIntentOptions.WEB_INTENT.toString()) {
        return "px"
    }

    return "pt"
}

const parseAndRouteCommand = async (command) => {
    let action = command.action;

    let f = commandHandlers[action];

    if (typeof f !== "function") {
        throw new Error(`Unknown Command: ${action}`);
    }
    
    console.log(f.name)
    return f(command);
};


const commandHandlers = {
    // Document creation
    createDocument,
    
    // Document management
    openDocument,
    saveDocument,
    exportPdf,
    
    // Paragraph styles
    getParagraphStyles,
    createParagraphStyle,
    applyParagraphStyle,
    
    // Text frames
    createTextFrame,
    getTextFrames,
    linkTextFrames,
    insertText,
    
    // Images
    placeImage,
    getImages,
    placeImageFromWord,
    
    // Pages
    addPage,
    getPages,
    getMasterPages,
    
    // Word import layout
    layoutWordContent
};


const getActiveDocumentSettings = (command) => {
    const document = app.activeDocument


    const d = document.documentPreferences
    const documentPreferences = {
        pageWidth:d.pageWidth,
        pageHeight:d.pageHeight,
        pagesPerDocument:d.pagesPerDocument,
        facingPages:d.facingPages,
        measurementUnit:getUnitForIntent(d.intent)
    }

    const marginPreferences = {
        top:document.marginPreferences.top,
        bottom:document.marginPreferences.bottom,
        left:document.marginPreferences.left,
        right:document.marginPreferences.right,
        columnCount : document.marginPreferences.columnCount,
        columnGutter : document.marginPreferences.columnGutter
    }
    return {documentPreferences, marginPreferences}
}

const checkRequiresActiveDocument = async (command) => {
    if (!requiresActiveDocument(command)) {
        return;
    }

    let document = app.activeDocument;
    if (!document) {
        throw new Error(
            `${command.action} : Requires an open InDesign document`
        );
    }
};

const requiresActiveDocument = (command) => {
    // Commands that don't require an active document
    const noDocRequired = ["createDocument", "openDocument"];
    return !noDocRequired.includes(command.action);
};


module.exports = {
    getActiveDocumentSettings,
    checkRequiresActiveDocument,
    parseAndRouteCommand
};
