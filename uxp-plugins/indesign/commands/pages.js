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
 * Add a new page to the document
 */
const addPage = async (command) => {
    const doc = app.activeDocument;
    const options = command.options;
    
    let newPage;
    
    if (options.atIndex !== undefined) {
        // Insert at specific position
        if (options.atIndex >= doc.pages.length) {
            // Add at end
            newPage = doc.pages.add();
        } else {
            // Insert before the page at index
            const refPage = doc.pages.item(options.atIndex);
            newPage = doc.pages.add(1818653800, refPage); // LocationOptions.BEFORE
        }
    } else {
        // Add at end
        newPage = doc.pages.add();
    }
    
    // Apply master page if specified
    if (options.masterPage) {
        try {
            const master = doc.masterSpreads.itemByName(options.masterPage);
            if (master.isValid) {
                newPage.appliedMaster = master;
            }
        } catch (e) {
            console.log(`Master page '${options.masterPage}' not found`);
        }
    }
    
    // Find the index of the new page
    let pageIndex = -1;
    for (let i = 0; i < doc.pages.length; i++) {
        const page = doc.pages.item(i);
        if (page && page.id === newPage.id) {
            pageIndex = i;
            break;
        }
    }
    
    return {
        pageId: newPage.id.toString(),
        pageIndex: pageIndex,
        pageName: newPage.name
    };
};

/**
 * Get information about all pages in the document
 */
const getPages = async (command) => {
    const doc = app.activeDocument;
    const pages = [];
    
    for (let i = 0; i < doc.pages.length; i++) {
        const page = doc.pages.item(i);
        const bounds = page.bounds;
        
        pages.push({
            index: i,
            id: page.id.toString(),
            name: page.name,
            bounds: {
                y1: bounds[0],
                x1: bounds[1],
                y2: bounds[2],
                x2: bounds[3],
                width: bounds[3] - bounds[1],
                height: bounds[2] - bounds[0]
            },
            appliedMaster: page.appliedMaster ? page.appliedMaster.name : null,
            marginPreferences: {
                top: page.marginPreferences.top,
                bottom: page.marginPreferences.bottom,
                left: page.marginPreferences.left,
                right: page.marginPreferences.right,
                columnCount: page.marginPreferences.columnCount,
                columnGutter: page.marginPreferences.columnGutter
            }
        });
    }
    
    return pages;
};

/**
 * Get information about master pages
 */
const getMasterPages = async (command) => {
    const doc = app.activeDocument;
    const masterPages = [];
    
    for (let i = 0; i < doc.masterSpreads.length; i++) {
        const master = doc.masterSpreads.item(i);
        
        masterPages.push({
            index: i,
            id: master.id.toString(),
            name: master.name,
            baseName: master.baseName,
            pageCount: master.pages.length
        });
    }
    
    return masterPages;
};

module.exports = {
    addPage,
    getPages,
    getMasterPages
};
