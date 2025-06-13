/**
 * GitHub  https://github.com/tanaikech/PDFApp<br>
 * Library name
 * @type {string}
 * @const {string}
 * @readonly
 */
var appName = "PDFApp";

var _PDFLib_Instance_ = null;
var _Fontkit_Instance_ = null;
var _PDFLib_Promise_ = null;
var _Fontkit_Promise_ = null;

async function _ensurePdfLibLoaded() {
  if (_PDFLib_Instance_) return _PDFLib_Instance_;
  if (!_PDFLib_Promise_) {
    _PDFLib_Promise_ = new Promise(async (resolve, reject) => {
      try {
        const cdnjs = "https://cdn.jsdelivr.net/npm/pdf-lib/dist/pdf-lib.min.js";
        if (typeof this.setTimeout === "undefined") {
          this.setTimeout = function (callback, delay, ...args) {
            if (delay && delay > 0) {
              Utilities.sleep(delay); // Synchronous sleep
            }
            if (typeof callback === "function") {
              return callback.apply(null, args);
            }
            return 0;
          };
        }

        if (typeof this.clearTimeout === "undefined") {
          this.clearTimeout = function (id) {};
        }
        const jsLibContent = UrlFetchApp.fetch(cdnjs).getContentText();
        eval(jsLibContent);
        if (!globalThis.PDFLib) throw new Error("Failed to load PDFLib from CDN: globalThis.PDFLib not defined.");
        _PDFLib_Instance_ = globalThis.PDFLib;
        delete globalThis.PDFLib;
        resolve(_PDFLib_Instance_);
      } catch (e) {
        _PDFLib_Promise_ = null;
        console.error("Error loading PDFLib:", e);
        reject(e);
      }
    });
  }
  return _PDFLib_Promise_;
}

async function _ensureFontkitLoaded() {
  if (_Fontkit_Instance_) return _Fontkit_Instance_;
  if (_Fontkit_Promise_) return await _Fontkit_Promise_;
  _Fontkit_Promise_ = new Promise(async (resolve, reject) => {
    try {
      const cdnFontkit = "https://unpkg.com/@pdf-lib/fontkit/dist/fontkit.umd.min.js";
      const jsLibContent = UrlFetchApp.fetch(cdnFontkit).getContentText();
      eval(jsLibContent);
      if (!globalThis.fontkit) throw new Error("Failed to load Fontkit from CDN: globalThis.fontkit not defined.");
      _Fontkit_Instance_ = globalThis.fontkit;
      delete globalThis.fontkit;
      resolve(_Fontkit_Instance_);
    } catch (e) {
      _Fontkit_Promise_ = null;
      console.error("Error loading Fontkit:", e);
      reject(e);
    }
  });
  return _Fontkit_Promise_;
}

/**
 * ### Description
 * Give the source PDF blob. This method is used with other methods.
 * 
 * @param {Object} blob PDF blob.
 * @return {PDFApp}
 */
function setPDFBlob(blob = null) {
  if (!blob || blob.toString() != "Blob" || blob.getContentType() != MimeType.PDF) {
    throw new Error("Please set the source PDF blob using the setPDFBlob method.");
  }
  this._gPDFA_Blob = blob;
  return this;
}

/**
 * ### Description
 * When you want to use the standard font, please use this method.
 * You can select one of them from the below page.
 * ref: https://pdf-lib.js.org/docs/api/enums/standardfonts
 * This method is used with other methods.
 * 
 * @param {String} name Font name of the built-in standard fonts.
 * @return {PDFApp}
 */
function useStandardFont(name = null) {
  this._gPDFA_Config = this._gPDFA_Config || {};
  this._gPDFA_Config.useStandardFont = name;
  return this;
}

/**
 * ### Description
 * When you want to use the custom font, please use this method. This method is used with other methods.
 * 
 * @param {Object} blob Blob of custom font. TTF and OTF files can be used.
 * @return {PDFApp}
 */
function useCustomFont(blob = null) {
  this._gPDFA_Config = this._gPDFA_Config || {};
  this._gPDFA_Config.useCustomFont = blob;
  return this;
}

/**
 * ### Description
 * Export specific pages from a PDF blob.
 * 
 * @param {number[]} pageNumbers Array including the page numbers you want to export.
 * @return {promise} PDF Blob including the exported pages.
 */
async function exportPages(pageNumbers) {
  if (!this._gPDFA_Blob) {
    throw new Error("Please set the source PDF blob using the setPDFBlob method.");
  }
  const pdfLib = await _ensurePdfLibLoaded();
  const PDFA = new PDFApp(this._gPDFA_Config || {}, pdfLib);
  return PDFA.exportPages(this._gPDFA_Blob, pageNumbers);
}

/**
 * ### Description
 * Get PDF metadata from a PDF blob.
 * 
 * @return {promise} PDF metadata.
 */
async function getMetadata() {
  if (!this._gPDFA_Blob) {
    throw new Error("Please set the source PDF blob using the setPDFBlob method.");
  }
  const pdfLib = await _ensurePdfLibLoaded();
  const PDFA = new PDFApp(this._gPDFA_Config || {}, pdfLib);
  return PDFA.getMetadata(this._gPDFA_Blob);
}

/**
 * ### Description
 * Update PDF metadata of a PDF blob.
 * 
 * @param {Object} object Object for updating PDF metadata.
 * @return {promise} Updated PDF blob.
 */
async function updateMetadata(object) {
  if (!this._gPDFA_Blob) {
    throw new Error("Please set the source PDF blob using the setPDFBlob method.");
  }
  const pdfLib = await _ensurePdfLibLoaded();
  const PDFA = new PDFApp(this._gPDFA_Config || {}, pdfLib);
  return PDFA.updateMetadata(this._gPDFA_Blob, object);
}

/**
 * ### Description
 * Reorder pages of a PDF blob.
 * 
 * @param {Object} object Object for reordering pages of PDF.
 * @return {promise} Updated PDF blob.
 */
async function reorderPages(object) {
  if (!this._gPDFA_Blob) {
    throw new Error("Please set the source PDF blob using the setPDFBlob method.");
  }
  const pdfLib = await _ensurePdfLibLoaded();
  const PDFA = new PDFApp(this._gPDFA_Config || {}, pdfLib);
  return PDFA.reorderPages(this._gPDFA_Blob, object);
}

/**
 * ### Description
 * Merge multiple PDF files in a single PDF.
 * 
 * @param {Object[]} pdfBlobs Array including PDF Blobs for merging in a single PDF.
 * @return {promise} Merged PDF Blob.
 */
async function mergePDFs(pdfBlobs) {
  if (!pdfBlobs.every(blob => blob && blob.toString() == "Blob" && blob.getContentType() == MimeType.PDF)) {
    throw new Error("Please set the source PDF blobs for merging in a single PDF.");
  }
  const pdfLib = await _ensurePdfLibLoaded();
  const PDFA = new PDFApp(this._gPDFA_Config || {}, pdfLib);
  return PDFA.mergePDFs(pdfBlobs);
}

/**
 * ### Description
 * Convert PDF pages to PNG images.
 * ref: https://tanaikech.github.io/2023/01/11/converting-all-pages-in-pdf-file-to-png-images-using-google-apps-script/
 *
 * @return {promise} PDF Blob.
 */
async function convertPDFToPng() {
  if (!this._gPDFA_Blob) {
    throw new Error("Please set the source PDF blob using the setPDFBlob method.");
  }
  const pdfLib = await _ensurePdfLibLoaded();
  const PDFA = new PDFApp(this._gPDFA_Config || {}, pdfLib);
  return PDFA.convertPDFToPng(this._gPDFA_Blob);
}

/**
 * ### Description
 * Get values from PDF Form.
 * ref: https://tanaikech.github.io/2023/08/02/retrieving-and-putting-values-for-pdf-forms-using-google-apps-script/
 *
 * @return {promise} Object including the values of PDF Form.
 */
async function getValuesFromPDFForm() {
  if (!this._gPDFA_Blob) {
    throw new Error("Please set the source PDF blob using the setPDFBlob method.");
  }
  const pdfLib = await _ensurePdfLibLoaded();
  const PDFA = new PDFApp(this._gPDFA_Config || {}, pdfLib);
  return PDFA.getValuesFromPDFForm(this._gPDFA_Blob);
}

/**
 * ### Description
 * Set values to PDF Form.
 * ref: https://tanaikech.github.io/2023/08/02/retrieving-and-putting-values-for-pdf-forms-using-google-apps-script/
 *
 * @param {Object} object Object for putting values to PDF Form.
 * @return {promise} Object including the values of PDF Form.
 */
async function setValuesToPDFForm(object) {
  if (!this._gPDFA_Blob) {
    throw new Error("Please set the source PDF blob using the setPDFBlob method.");
  }
  const pdfLib = await _ensurePdfLibLoaded();
  const fontkit = (this._gPDFA_Config && this._gPDFA_Config.useCustomFont) ? await _ensureFontkitLoaded() : null;
  const PDFA = new PDFApp(this._gPDFA_Config || {}, pdfLib, fontkit);
  return PDFA.setValuesToPDFForm(this._gPDFA_Blob, object);
}

/**
 * ### Description
 * Create PDF Form By Google Slide template.
 * ref: https://medium.com/google-cloud/creating-pdf-forms-from-google-slide-template-using-google-apps-script-cef35e7d9822
 *
 * @param {String} id File ID of Google Slide template.
 * @param {Object} object Object for setting.
 * @return {promise} Object including the values of PDF Form.
 */
async function createPDFFormBySlideTemplate(id, object) {
  const pdfLib = await _ensurePdfLibLoaded();
  const fontkit = (this._gPDFA_Config && this._gPDFA_Config.useCustomFont) ? await _ensureFontkitLoaded() : null;
  const PDFA = new PDFApp(this._gPDFA_Config || {}, pdfLib, fontkit);
  return PDFA.createPDFFormBySlideTemplate(id, object);
}

/**
 * ### Description
 * Embed objects into PDF blob.
 * ref: https://medium.com/google-cloud/embedding-objects-in-pdf-using-google-apps-script-ddbee857c642
 *
 * @param {Object} pdfBlob Blob of PDF data for embedding objects.
 * @param {Object} object Object including the values for embedding objects.
 * @return {promise} PDF Blob.
 */
async function embedObjects(object) {
  if (!this._gPDFA_Blob) {
    throw new Error("Please set the source PDF blob using the setPDFBlob method.");
  }
  const pdfLib = await _ensurePdfLibLoaded();
  const fontkit = await _ensureFontkitLoaded();
  const PDFA = new PDFApp(this._gPDFA_Config || {}, pdfLib, fontkit);
  return PDFA.embedObjects(this._gPDFA_Blob, object);
}

/**
 * ### Description
 * Insert header and/or footer into PDF blob.
 *
 * @param {Object} pdfBlob Blob of PDF data for embedding objects.
 * @param {Object} object Object including the values for inserting header and footer.
 * @return {promise} PDF Blob.
 */
async function insertHeaderFooter(object) {
  if (!this._gPDFA_Blob) {
    throw new Error("Please set the source PDF blob using the setPDFBlob method.");
  }
  const pdfLib = await _ensurePdfLibLoaded();
  const fontkit = (this._gPDFA_Config && this._gPDFA_Config.useCustomFont) ? await _ensureFontkitLoaded() : null;
  const PDFA = new PDFApp(this._gPDFA_Config || {}, pdfLib, fontkit);
  return PDFA.insertHeaderFooter(this._gPDFA_Blob, object);
}

/**
 * ### Description
 * Split each page of a PDF to an individual PDF file.
 *
 * @return {promise} PDF Blobs.
 */
async function splitPDF() {
  if (!this._gPDFA_Blob) {
    throw new Error("Please set the source PDF blob using the setPDFBlob method.");
  }
  const pdfLib = await _ensurePdfLibLoaded();
  const PDFA = new PDFApp(this._gPDFA_Config || {}, pdfLib);
  return PDFA.splitPDF(this._gPDFA_Blob);
}

/**
 * ### Description
 * Add page numbers to PDF.
 * 
 * @param {Object} object Object including the format of page number.
 * @return {promise} PDF Blobs.
 */
async function addPageNumbers(object) {
  if (!this._gPDFA_Blob) {
    throw new Error("Please set the source PDF blob using the setPDFBlob method.");
  }
  const pdfLib = await _ensurePdfLibLoaded();
  const PDFA = new PDFApp(this._gPDFA_Config || {}, pdfLib);
  return PDFA.addPageNumbers(this._gPDFA_Blob, object);
}


/**
 * ### Description
 * This is a Class PDFApp for managing PDF using Google Apps Script.
 *
 * Author: Tanaike ( https://tanaikech.github.io/ )
 */
class PDFApp {
  /**
   * ### Description
   * Constructor of this class.
   *
   * @return {void}
   */
  constructor(config, pdfLib, fontkitInstance) {
    this.PDFLib = pdfLib;
    this.fontkit = fontkitInstance;
    if (config && config.useCustomFont && config.useCustomFont.toString() == "Blob") {
      this.customFont = config.useCustomFont;
    } else if (config && config.useStandardFont && typeof config.useStandardFont == "string") {
      this.standardFont = config.useStandardFont;
    }
  }

  /**
   * ### Description
   * Export specific pages from a PDF blob.
   * ref: https://medium.com/google-cloud/exporting-specific-pages-from-a-pdf-as-a-new-pdf-using-google-apps-script-2f22d07b4618
   *
   * @param {Object} pdfBlob Blob of PDF data by retrieving with Google Apps Script.
   * @param {number[]} pageNumbers Array including the page numbers you want to export.
   * @return {promise} PDF Blob including the exported pages.
   */
  async exportPages(pdfBlob, pageNumbers) {
    if (!pageNumbers || !Array.isArray(pageNumbers) || pageNumbers.length == 0) {
      throw new Error("Please set the page numbers you want to export.");
    }
    try {
      const pdfData = await this.getPDFObjectFromBlob_(pdfBlob);
      const pdfDoc = await this.PDFLib.PDFDocument.create();
      const pages = await pdfDoc.copyPages(pdfData, pdfData.getPageIndices());
      pages.forEach((page, i) => {
        if (pageNumbers.includes(i + 1)) {
          pdfDoc.addPage(page);
        }
      });
      const bytes = await pdfDoc.save();
      return Utilities.newBlob([...new Int8Array(bytes)], MimeType.PDF, `new_${pdfBlob.getName()}`);
    } catch (err) {
      throw err;
    }
  }

  /**
   * ### Description
   * Get PDF metadata from a PDF blob.
   * ref: https://medium.com/google-cloud/management-of-pdf-metadata-using-google-apps-script-60fd41f4fc16
   *
   * @param {Object} pdfBlob Blob of PDF data by retrieving with Google Apps Script.
   * @return {promise} PDF metadata.
   */
  async getMetadata(pdfBlob) {
    const keys = ["title", "subject", "author", "creator", "creationDate", "modificationDate", "keywords", "producer"];
    const pdfData = await this.getPDFObjectFromBlob_(pdfBlob);
    try {
      const metadata = keys.reduce((o, k) => ((o[k] = pdfData[`get${k.charAt(0).toUpperCase() + k.slice(1)}`]() || null), o), {});
      metadata.numberOfPages = pdfData.getPageCount();
      const pdfDoc = await this.PDFLib.PDFDocument.create();
      const pages = await pdfDoc.copyPages(pdfData, pdfData.getPageIndices());
      metadata.pageInfo = pages.map((page, i) => {
        const { width, height } = page.getSize();
        const { x, y } = page.getPosition();
        return { page: i + 1, pageWidth: width, pageHeight: height, defaultPositionX: x, defaultPositionY: y };
      });
      return metadata;
    } catch (err) {
      throw err;
    }
  }

  /**
   * ### Description
   * Update PDF metadata of a PDF blob.
   * ref: https://medium.com/google-cloud/management-of-pdf-metadata-using-google-apps-script-60fd41f4fc16
   *
   * @param {Object} pdfBlob Blob of PDF data by retrieving with Google Apps Script.
   * @param {Object} object Object including the values for updating metadata.
   * @return {promise} PDF Blob.
   */
  async updateMetadata(pdfBlob, object) {
    if (typeof object != "object" || Object.keys(object).length == 0) {
      throw new Error("Please set valid object for updating PDF metadata.");
    }
    const pdfData = await this.getPDFObjectFromBlob_(pdfBlob);
    const keys = ["title", "subject", "author", "creator", "creationDate", "modificationDate", "keywords", "producer"];
    try {
      for (const k of keys) {
        if (object.hasOwnProperty(k)) {
          const f = `set${k.charAt(0).toUpperCase() + k.slice(1)}`;
          if (k == "title") {
            pdfData[f](...(Array.isArray(object[k]) ? object[k] : [object[k]]));
          } else {
            let valueToSet = object[k];
            if (["creationDate", "modificationDate"].includes(k)) {
              valueToSet = new Date(object[k]);
            } else if (k == "keywords") {
              valueToSet = Array.isArray(object[k]) ? object[k] : [object[k]];
            }
            pdfData[f](valueToSet);
          }
        }
      }
      const bytes = await pdfData.save();
      return Utilities.newBlob([...new Int8Array(bytes)], MimeType.PDF, `new_${pdfBlob.getName()}`);
    } catch (err) {
      throw err;
    }
  }

  /**
   * ### Description
   * Update PDF metadata of a PDF blob.
   * ref: https://medium.com/google-cloud/changing-order-of-pages-in-pdf-file-using-google-apps-script-f6b3de05d7df
   *
   * @param {Object} pdfBlob Blob of PDF data by retrieving with Google Apps Script.
   * @param {Object} object Object including the values for reordering pages.
   * @return {promise} PDF Blob.
   */
  async reorderPages(pdfBlob, object) {
    if (typeof object != "object" || Object.keys(object).length == 0) {
      throw new Error("Please set valid object for reordering PDF pages.");
    }
    try {
      const { newOrderOfpages, ignoreSkippedPages } = object;
      const pdfData = this.getPDFObjectFromBlob_(pdfBlob);
      const numberOfPages = pdfData.getPageCount();
      const maxPage = Math.max(...newOrderOfpages);
      if (numberOfPages < maxPage || numberOfPages < newOrderOfpages.length) {
        throw new Error("Maximum page in the order of pages is over than the maximum page of the original PDF file.");
      }
      let skippedPages = [];
      if (!ignoreSkippedPages && numberOfPages > newOrderOfpages.length) {
        skippedPages = [...Array(numberOfPages)].map((_, i) => i + 1).filter(e => !newOrderOfpages.includes(e));
      }
      const pdfDoc = await this.PDFLib.PDFDocument.create();
      const pages = await pdfDoc.copyPages(pdfData, pdfData.getPageIndices());
      [...newOrderOfpages, ...skippedPages].forEach(e => pdfDoc.addPage(pages[e - 1]));
      const bytes = await pdfDoc.save();
      return Utilities.newBlob([...new Int8Array(bytes)], MimeType.PDF, `new_${pdfBlob.getName()}`);
    } catch (err) { throw err; }
  }

  /**
   * ### Description
   * Merge multiple PDF files in a single PDF.
   * ref: https://tanaikech.github.io/2023/01/10/merging-multiple-pdf-files-as-a-single-pdf-file-using-google-apps-script/
   *
   * @param {Object[]} pdfBlobs Array including PDF Blobs for merging in a single PDF.
   * @return {promise} PDF Blob.
   */
  async mergePDFs(pdfBlobs) {
    try {
      const data = pdfBlobs.map(blob => new Uint8Array(blob.getBytes()));
      const pdfDoc = await this.PDFLib.PDFDocument.create();
      for (let i = 0; i < data.length; i++) {
        const pdfData = await this.PDFLib.PDFDocument.load(data[i]);
        const pages = await pdfDoc.copyPages(pdfData, pdfData.getPageIndices());
        pages.forEach(page => pdfDoc.addPage(page));
      }
      const bytes = await pdfDoc.save();
      return Utilities.newBlob([...new Int8Array(bytes)], MimeType.PDF, "new_PDFFile.pdf");
    } catch (err) {
      throw err;
    }
  }

  /**
   * ### Description
   * Convert PDF pages to PNG images.
   * ref: https://tanaikech.github.io/2023/01/11/converting-all-pages-in-pdf-file-to-png-images-using-google-apps-script/
   *
   * @param {Object} pdfBlob Blob of PDF data.
   * @return {promise} PDF Blob.
   */
  async convertPDFToPng(pdfBlob) {
    try {
      const pdfData = this.getPDFObjectFromBlob_(pdfBlob);
      const pageLength = pdfData.getPageCount();
      console.log(`Total pages: ${pageLength}`);
      const obj = { imageBlobs: [], fileIds: [] };
      const token = ScriptApp.getOAuthToken();
      for (let i = 0; i < pageLength; i++) {
        console.log(`Processing page: ${i + 1}`);
        const pdfDoc = await this.PDFLib.PDFDocument.create();
        const [page] = await pdfDoc.copyPages(pdfData, [i]);
        pdfDoc.addPage(page);
        const bytes = await pdfDoc.save();
        const blob = Utilities.newBlob([...new Int8Array(bytes)], MimeType.PDF, `temp_page${i + 1}.pdf`);
        const id = DriveApp.createFile(blob).getId();
        Utilities.sleep(3000); // This is used for preparing the thumbnail of the created file.
        const res = UrlFetchApp.fetch(
          `https://drive.google.com/thumbnail?id=${id}&sz=w1000`,
          {
            headers: { authorization: "Bearer " + token },
            muteHttpExceptions: true
          }
        );
        if (res.getResponseCode() != 200) {
          throw new Error(`Failed to get thumbnail: ${res.getContentText()}`);
        }
        const imageBlob = res.getBlob().setName(`page${i + 1}.png`);
        obj.imageBlobs.push(imageBlob);
        obj.fileIds.push(id);
      }
      obj.fileIds.forEach(id => DriveApp.getFileById(id).setTrashed(true));
      return obj.imageBlobs;
    } catch (err) {
      throw err;
    }
  }

  /**
   * ### Description
   * Get values from PDF Form.
   * ref: https://medium.com/google-cloud/retrieving-and-putting-values-for-pdf-forms-using-google-apps-script-92412a7cf0af
   *
   * @param {Object} pdfBlob Blob of PDF data.
   * @return {promise} Object including the values of PDF Form.
   */
  async getValuesFromPDFForm(pdfBlob) {
    try {
      const pdfData = this.getPDFObjectFromBlob_(pdfBlob);
      const form = pdfData.getForm();
      const { PDFTextField, PDFDropdown, PDFCheckBox, PDFRadioGroup } = this.PDFLib;
      const obj = form.getFields().map(function (f) {
        const retObj = { name: f.getName() };
        if (f instanceof PDFTextField) {
          retObj.value = f.getText();
          retObj.type = "Textbox";
        } else if (f instanceof PDFDropdown) {
          retObj.value = f.getSelected();
          retObj.options = f.getOptions();
          retObj.type = "Dropdown";
        } else if (f instanceof PDFCheckBox) {
          retObj.value = f.isChecked();
          retObj.type = "Checkbox";
        } else if (f instanceof PDFRadioGroup) {
          retObj.value = f.getSelected();
          retObj.options = f.getOptions();
          retObj.type = "Radiobutton";
        } else {
          retObj.type = "Unsupported type";
        }
        return retObj;
      });
      return obj;
    } catch (err) { throw err; }
  }

  /**
   * ### Description
   * Set values to PDF Form.
   * ref: https://medium.com/google-cloud/retrieving-and-putting-values-for-pdf-forms-using-google-apps-script-92412a7cf0af
   *
   * @param {Object} pdfBlob Blob of PDF data.
   * @param {Object} object Object for putting values to PDF Form.
   * @return {promise} Object including the values of PDF Form.
   */
  async setValuesToPDFForm(pdfBlob, object) {
    if (!object.values || !Array.isArray(object.values)) {
      throw new Error("Please set valid values.");
    }
    try {
      const pdfData = this.getPDFObjectFromBlob_(pdfBlob);
      const form = pdfData.getForm();
      if (this.standardFont || this.customFont) {
        this.setCustomFont_(pdfData, form, { standardFont: this.standardFont, customFont: this.customFont });
      }
      const { PDFTextField, PDFDropdown, PDFCheckBox, PDFRadioGroup } = this.PDFLib;
      for (let { name, value } of object.values) {
        const field = form.getField(name);
        if (field instanceof PDFTextField) {
          field.setText(value);
        } else if (field instanceof PDFDropdown) {
          if (field.isMultiselect()) {
            for (let v of value) {
              field.select(v);
            }
          } else {
            field.select(value);
          }
        } else if (field instanceof PDFCheckBox) {
          field[value ? "check" : "uncheck"]();
        } else if (field instanceof PDFRadioGroup) {
          field.select(value);
        }
      }
      const bytes = await pdfData.save();
      return Utilities.newBlob([...new Int8Array(bytes)], MimeType.PDF, `new_${pdfBlob.getName()}`);
    } catch (e) {
      throw e;
    }
  }

  /**
   * ### Description
   * Create PDF Form By Google Slide template.
   * ref: https://medium.com/google-cloud/creating-pdf-forms-from-google-slide-template-using-google-apps-script-cef35e7d9822
   *
   * @param {String} id File ID of Google Slide template.
   * @param {Object} object Object for setting.
   * @return {promise} Object including the values of PDF Form.
   */
  async createPDFFormBySlideTemplate(id, object) {
    if (!id || id == "") {
      throw new Error("Please set the file ID of Google Slide including the template for PDF Form.");
    }
    if (!object || !object.values || !Array.isArray(object.values) || object.values.length == 0) {
      throw new Error("Please set valid object for creating PDF Form from Google Slide template.");
    }
    try {
      const obj = this.getObjectFromSlide_(id, object.values);
      const newBlob = await this.createFields_(this, obj);
      return newBlob;
    } catch (e) { throw e; }
  }

  /**
   * ### Description
   * Embed objects into PDF blob.
   * ref: https://medium.com/google-cloud/embedding-objects-in-pdf-using-google-apps-script-ddbee857c642
   *
   * @param {Object} pdfBlob Blob of PDF data for embedding objects.
   * @param {Object} object Object including the values for embedding objects.
   * @return {promise} PDF Blob.
   */
  async embedObjects(pdfBlob, object) {
    if (!object || typeof object != "object") {
      throw new Error("Please an object for embeddig the objects.");
    }
    try {
      const { updatedObject, customFontCheck } = this.updateObject_(object);
      const pdfData = this.getPDFObjectFromBlob_(pdfBlob);
      const pdfDoc = await this.PDFLib.PDFDocument.create();
      if (customFontCheck && this.fontkit) {
        pdfDoc.registerFontkit(this.fontkit);
      }
      const pages = await pdfDoc.copyPages(pdfData, pdfData.getPageIndices());
      for (let i = 0; i < pages.length; i++) {
        const page = pages[i];
        const forPage = updatedObject[`page${i + 1}`];
        if (forPage) {
          for (let j = 0; j < forPage.length; j++) {
            const o = forPage[j];
            if (o.imageFileId) {
              const image = await pdfDoc[o.method](o.imageBytes);
              if (o.scale) {
                const updatedImage = image.scale(o.scale);
                o.width = o.width || updatedImage.width;
                o.height = o.height || updatedImage.height;
              } else {
                o.width = o.width || image.width;
                o.height = o.height || image.height;
              }
              delete o.imageBytes;
              delete o.method;
              delete o.imageFileId;
              delete o.scale;
              page.drawImage(image, o);
            } else if (o.text) {
              if (o.standardFont || o.customFont) {
                o.font = await pdfDoc.embedFont(o.standardFont ? this.PDFLib.StandardFonts[o.standardFont] : o.customFont);
              }
              page.drawText(o.text, o);
            }
          }
        }
        pdfDoc.addPage(page);
      }
      const bytes = await pdfDoc.save();
      return Utilities.newBlob([...new Int8Array(bytes)], MimeType.PDF, `new_${pdfBlob.getName()}`);
    } catch (e) { throw e; }
  }


  /**
   * ### Description
   * Insert header and/or footer into PDF blob.
   *
   * @param {Object} pdfBlob Blob of PDF data.
   * @param {Object} object Object including the values for inserting header and footer.
   * @return {promise} PDF Blob.
   */
  async insertHeaderFooter(pdfBlob, object) {
    if (!object || typeof object != "object") {
      throw new Error("Please an object for embeddig the objects.");
    }
    try {
      const pdfDoc = await this.PDFLib.PDFDocument.create();
      const form = pdfDoc.getForm();
      let embeddedFont = null;
      if (this.standardFont || this.customFont) {
        this.setCustomFont_(pdfDoc, form, { standardFont: this.standardFont, customFont: this.customFont });
        embeddedFont = await pdfDoc.embedFont(this.standardFont ? this.PDFLib.StandardFonts[this.standardFont] : new Uint8Array(this.customFont.getBytes()));
      }
      const pdfData = this.getPDFObjectFromBlob_(pdfBlob);
      const numberOfPages = pdfData.getPageCount();
      const pages = await pdfDoc.copyPages(pdfData, pdfData.getPageIndices());
      const { header, footer } = object;
      const headers = header ? Object.entries(header).map(([k, v]) => [`header.${k}`, v]) : [];
      const footers = footer ? Object.entries(footer).map(([k, v]) => [`footer.${k}`, v]) : [];
      const sortOrder = ["LEFT", "CENTER", "RIGHT"];
      [footers, headers].forEach((f, _, x) => f.sort((a, b) => {
        const i1 = sortOrder.findIndex(e => a[0].includes(e.toLowerCase()));
        const i2 = sortOrder.findIndex(e => b[0].includes(e.toLowerCase()));
        const vlen = x.length;
        return (i1 > -1 ? i1 : vlen) - (i2 > -1 ? i2 : vlen);
      }));
      const alignObj = { "center": "Center", "left": "Left", "right": "Right" };
      for (let i = 0; i < numberOfPages; i++) {
        const pageNumber = i + 1;
        const page = pdfDoc.addPage(pages[i]);
        const pageHeight = page.getHeight();
        const pageWidth = page.getWidth();
        if (headers.length > 0) {
          const sizeWidthHead = pageWidth / (headers.length);
          for (let j = 0; j < headers.length; j++) {
            const [k, v] = headers[j];
            const o = {
              borderWidth: v.borderWidth || 0,
              x: j * sizeWidthHead,
              y: pageHeight - ((v.yOffset || 0) + (v.height || 20)),
              width: sizeWidthHead,
              height: v.height || 30,
              ...v,
              font: embeddedFont
            };
            await this.addHeaderFooterFields_(this, { page, form, pageNumber, k, v, o, alignObj, font: embeddedFont });
          }
        }
        if (footers.length > 0) {
          const sizeWidthFoot = pageWidth / (footers.length);
          for (let j = 0; j < footers.length; j++) {
            const [k, v] = footers[j];
            const o = {
              borderWidth: v.borderWidth || 0,
              x: j * sizeWidthFoot,
              y: v.yOffset || 0,
              width: sizeWidthFoot,
              height: v.height || 30,
              ...v,
              font: embeddedFont
            };
            await this.addHeaderFooterFields_(this, { page, form, pageNumber, k, v, o, alignObj, font: embeddedFont });
          }
        }
      }
      const bytes = await pdfDoc.save();
      return Utilities.newBlob([...new Int8Array(bytes)], MimeType.PDF, `new_${pdfBlob.getName()}`);
    } catch (e) {
      throw e;
    }
  }

  /**
   * ### Description
   * Split each page of a PDF to an individual PDF file.
   *
   * @param {Object} pdfBlob Blob of PDF data.
   * @return {promise} Array of PDF Blobs.
   */
  async splitPDF(pdfBlob) {
    try {
      const pdfData = this.getPDFObjectFromBlob_(pdfBlob);
      const pageLength = pdfData.getPageCount();
      console.log(`Total pages: ${pageLength}`);
      const pdfBlobs = [];
      for (let i = 0; i < pageLength; i++) {
        console.log(`Processing page: ${i + 1}`);
        const pdfDoc = await this.PDFLib.PDFDocument.create();
        const [page] = await pdfDoc.copyPages(pdfData, [i]);
        pdfDoc.addPage(page);
        const bytes = await pdfDoc.save();
        const blob = Utilities.newBlob([...new Int8Array(bytes)], MimeType.PDF, `page${i + 1}.pdf`);
        pdfBlobs.push(blob);
      }
      return pdfBlobs;
    } catch (err) { throw err; }
  }

  /**
   * ### Description
   * Add page numbers to PDF.
   * 
   * @param {Object} pdfBlob Blob of PDF data.
   * @param {Object} object Object including the format of page number.
   * @return {promise} PDF Blobs.
   */
  async addPageNumbers(pdfBlob, object) {
    if (!object || typeof object != "object" || !["size", "x", "y"].every(e => e in object)) {
      throw new Error("Please an object for adding page numbers.");
    }
    try {
      const pdfData = await this.getPDFObjectFromBlob_(pdfBlob);
      const pdfDoc = await this.PDFLib.PDFDocument.create();
      const pages = await pdfDoc.copyPages(pdfData, pdfData.getPageIndices());
      pages.forEach((page, i) => {
        let drawOptions = { ...object };
        if (isNaN(drawOptions.x)) {
          const { width } = page.getSize();
          const xPositions = { center: width / 2, left: (drawOptions.size || 10), right: width - (drawOptions.size || 10) * ((i + 1).toString().length) - (drawOptions.size || 10) };
          drawOptions.x = xPositions[object.x.toLowerCase()] || width / 2;
        }
        page.drawText(`${i + 1}`, drawOptions);
        pdfDoc.addPage(page);
      });
      const bytes = await pdfDoc.save();
      return Utilities.newBlob([...new Int8Array(bytes)], MimeType.PDF, `new_${pdfBlob.getName()}`);
    } catch (err) {
      throw err;
    }
  }

  /**
   * ### Description
   * Create fields of PDF Form.
   *
   * @param {Object} self this of this Class object.
   * @param {Object} object Object for creating fields.
   * @return {Object} Blob of new PDF.
   */
  addHeaderFooterFields_(self, object) {
    const { page, form, pageNumber, k, v, o, alignObj, font } = object;
    const fieldName = `${k}.${pageNumber}`;
    const textBox = form.createTextField(fieldName);
    if (v.text) {
      textBox.setText(v.text.toString().replace("{{pageNumber}}", pageNumber).replace("{{totalPages}}", form.doc.getPageCount()));
    }
    if (v.alignment) {
      textBox.setAlignment(self.PDFLib.TextAlignment[alignObj[v.alignment.toLowerCase()]]);
    }
    textBox.disableScrolling();
    textBox.disableMultiline();
    textBox.enableReadOnly();
    ["x", "y", "width", "text"].forEach(e => delete v[e]);
    textBox.addToPage(page, o);
  }

  /**
   * ### Description
   * Create fields of PDF Form.
   *
   * @param {Object} self this of this Class object.
   * @param {Object} object Object for creating fields.
   * @return {Object} Blob of new PDF.
   */
  async createFields_(self, object) {
    const { obj, blob } = object;
    const pdfDoc = await this.PDFLib.PDFDocument.create();
    const form = pdfDoc.getForm();
    if (self.standardFont || self.customFont) {
      await self.setCustomFont_(pdfDoc, form, { standardFont: self.standardFont, customFont: self.customFont });
    }
    const pdfData = this.getPDFObjectFromBlob_(blob);
    const numberOfPages = pdfData.getPageCount();
    const sourcePages = await pdfDoc.copyPages(pdfData, pdfData.getPageIndices());
    const xAxisOffset = 0.5;
    const yAxisOffset = 0.5;
    for (let i = 0; i < numberOfPages; i++) {
      const pageNumber = i + 1;
      const page = pdfDoc.addPage(sourcePages[i]);
      const pageHeight = page.getHeight();
      const yOffset = pageHeight;
      obj[i].forEach((v, k) => {
        if (k == "checkbox") {
          v.forEach(t => {
            t.forEach(u => {
              const checkbox = form.createCheckBox(u.title);
              checkbox.addToPage(page, { x: u.leftOffset - xAxisOffset, y: yOffset - u.topOffset - u.height + yAxisOffset, width: u.width, height: u.height });
              self.setStyles_(checkbox, u);
            });
          });
        } else if (k == "radiobutton") {
          v.forEach((t, kk) => {
            const radio = form.createRadioGroup(t[0].group ? `${t[0].group}.page${pageNumber}` : `radiobutton.${kk}.page${pageNumber}`);
            t.forEach(u => {
              radio.addOptionToPage(u.title, page, { x: u.leftOffset - xAxisOffset, y: yOffset - u.topOffset - u.height + yAxisOffset, width: u.width, height: u.height });
              self.setStyles_(radio, u);
            });
          });
        } else if (k == "textbox") {
          v.forEach(t => {
            t.forEach(u => {
              const textBox = form.createTextField(u.title);
              textBox.addToPage(page, {
                x: u.leftOffset - xAxisOffset,
                y: yOffset - u.topOffset - u.height + yAxisOffset,
                width: u.width,
                height: u.height,
              });
              self.setStyles_(textBox, u);
            });
          });
        } else if (k == "dropdownlist") {
          v.forEach(t => {
            t.forEach(u => {
              const drowdown = form.createDropdown(u.title);
              drowdown.addToPage(page, {
                x: u.leftOffset - xAxisOffset,
                y: yOffset - u.topOffset - u.height + yAxisOffset,
                width: u.width,
                height: u.height
              });
              self.setStyles_(drowdown, u);
            });
          });
        }
      });
    }
    const bytes = await pdfDoc.save();
    return Utilities.newBlob([...new Int8Array(bytes)], MimeType.PDF, `new_${blob.getName()}`);
  }

  /**
   * ### Description
   * Set custom font to PDF form.
   *
   * @param {Object} pdfDoc Object of PDF document.
   * @param {Object} form Object of PDF form.
   * @return {void}
   */
  async setCustomFont_(pdfDoc, form, { standardFont, customFont }) {
    let embeddedFontRef;
    if (standardFont) {
      embeddedFontRef = await pdfDoc.embedFont(this.PDFLib.StandardFonts[standardFont]);
    } else if (customFont) {
      if (!this.fontkit) throw new Error("Fontkit not loaded. Ensure useCustomFont is called or fontkit is provided if using custom fonts in setCustomFont_.");
      pdfDoc.registerFontkit(this.fontkit);
      embeddedFontRef = await pdfDoc.embedFont(new Uint8Array(customFont.getBytes()));
    }

    // Ref: https://github.com/Hopding/pdf-lib/issues/1152
    const rawUpdateFieldAppearances = form.updateFieldAppearances.bind(form);
    form.updateFieldAppearances = function () {
      return rawUpdateFieldAppearances(embeddedFontRef);
    };
  }

  /**
   * ### Description
   * Set styles to the field of PDF form.
   *
   * @param {Object} instance Instance of field.
   * @param {Object} u Methods and values for setting the styles.
   * @return {Object} Object for creating the fields of PDF Form.
   */
  setStyles_(instance, u) {
    if (u.description.methods && u.description.methods.length > 0) {
      u.description.methods.forEach(({ method, value }) => {
        if (value && Array.isArray(value)) {
          value = [...value];
        }
        instance[method](value || null)
      });
    }
  }

  /**
   * ### Description
   * Get an object for creating the fields of PDF Form from Google Slide.
   *
   * @param {String} id File ID of Google Slide template.
   * @return {Object} Object for creating the fields of PDF Form.
   */
  getObjectFromSlide_(id, object) {
    const inputObj = object.reduce((o, e) => (o[e.shapeTitle] = e, o), {});
    const slide = SlidesApp.openById(id);
    const slides = slide.getSlides();
    const ar = slides.map((s, i) =>
      s.getShapes().reduce((ar, e) => {
        const t = e.getTitle().trim();
        if (t && inputObj[t]) {
          const page = i + 1;
          const [type, group, name] = t.split(".").map(f => f.trim());
          let setMethods = inputObj[t];
          if (type == "radiobutton" && inputObj[t] && inputObj[t].methods && inputObj[t].methods.length > 0) {
            const temp = JSON.parse(JSON.stringify(inputObj[t]));
            temp.methods.forEach(f => {
              if (f.method == "select") {
                f.value = `${f.value}.page${page}`;
              }
            });
            setMethods = temp;
          }
          ar.push({
            title: `${t}.page${page}`,
            page,
            type,
            group,
            name,
            shape: e,
            topOffset: e.getTop(),
            leftOffset: e.getLeft(),
            width: e.getWidth(),
            height: e.getHeight(),
            description: setMethods,
          });
        }
        return ar;
      }, [])
    );

    const duplicateCheckObj = ar.reduce((m, e) => {
      e.forEach(({ title }) => m.set(title, m.has(title) ? m.get(title) + 1 : 1));
      return m;
    }, new Map());
    const duplicateCheck = [...duplicateCheckObj].filter(([_, v]) => v > 1);
    if (duplicateCheck.length > 0) {
      DriveApp.getFileById(id).setTrashed(true);
      throw new Error(`Duplicate titles were found. The duplicated titles are "${duplicateCheck.map(([k]) => k).join(",")}".`);
    }

    ar.forEach(page => page.forEach(({ shape }) => shape.remove()));
    slide.saveAndClose();

    const obj = ar.map(d => d.reduce((m, e) => m.set(e.type, m.has(e.type) ? [...m.get(e.type), e] : [e]), new Map()));
    obj.forEach(page =>
      page.forEach((v, k, m) => m.set(k, v.reduce((mm, e) => mm.set(e.group, mm.has(e.group) ? [...mm.get(e.group), e] : [e]), new Map())))
    );
    const blob = DriveApp.getFileById(id).getBlob();
    return { obj, blob };
  }

  /**
   * ### Description
   * Update object.
   *
   * @param {Object} object Object for embedding objects into PDF.
   * @return {void}
   */
  updateObject_(object) {
    let customFontCheck = false;
    const updatedObject = Object.fromEntries(Object.entries(object).map(([k, p]) => [k, p.map(e => {
      if (e.imageFileId) {
        const imageBlob = DriveApp.getFileById(e.imageFileId).getBlob();
        const imageBytes = new Uint8Array(imageBlob.getBytes());
        const mimeType = imageBlob.getContentType();
        let method;
        if (mimeType == MimeType.PNG) {
          method = "embedPng";
        } else if (mimeType == MimeType.JPEG) {
          method = "embedJpg";
        } else {
          throw new Error("This image file cannot be used.");
        }
        e.method = method;
        e.imageBytes = imageBytes;
      } else if (e.text) {
        if (e.customFont) {
          customFontCheck = true;
          e.customFont = new Uint8Array(e.customFont.getBytes());
        }
      }
      return e;
    })]));
    return { updatedObject, customFontCheck };
  }

  /**
   * ### Description
   * Get PDF object from PDF blob.
   *
   * @param {Object} blob PDF Blob retrieved by Google Apps Script.
   * @return {void}
   */
  async getPDFObjectFromBlob_(blob) {
    if (!this.PDFLib) throw new Error("PDFLib not initialized on PDFApp instance.");
    return await this.PDFLib.PDFDocument.load(new Uint8Array(blob.getBytes()));
  }
}
