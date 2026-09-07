/**
 * Excel Parser Module
 * Handles Excel and CSV file parsing, header signature detection, and normalization.
 */

const ExcelParser = (() => {
  /**
   * Reads a File object using FileReader as ArrayBuffer
   * @param {File} file
   * @returns {Promise<ArrayBuffer>}
   */
  const readFileAsArrayBuffer = (file) => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => resolve(e.target.result);
      reader.onerror = (err) => reject(err);
      reader.readAsArrayBuffer(file);
    });
  };

  /**
   * Normalizes header strings (trims whitespace, optional case)
   * @param {string} str
   * @param {boolean} shouldTrim
   * @returns {string}
   */
  const cleanHeader = (str, shouldTrim = true) => {
    if (str === null || str === undefined) return '';
    let val = String(str);
    if (shouldTrim) {
      val = val.trim().replace(/\s+/g, ' ');
    }
    return val;
  };

  /**
   * Generates a schema signature key from an array of headers
   * @param {string[]} headers
   * @param {boolean} isStrictOrder - whether column order matters
   * @param {boolean} shouldTrim
   * @returns {{ signature: string, normalizedHeaders: string[] }}
   */
  const generateSchemaSignature = (headers, isStrictOrder = false, shouldTrim = true) => {
    const cleaned = headers.map(h => cleanHeader(h, shouldTrim)).filter(h => h.length > 0);
    
    if (isStrictOrder) {
      // Order strictly preserved: header1 -> header2 -> header3
      const signature = 'STRICT::' + cleaned.map(h => h.toLowerCase()).join('==>');
      return { signature, normalizedHeaders: cleaned };
    } else {
      // Order independent: sort alphabetically for signature
      const sortedKeys = [...cleaned].map(h => h.toLowerCase()).sort();
      const signature = 'FLEX::' + sortedKeys.join('|||');
      return { signature, normalizedHeaders: cleaned };
    }
  };

  /**
   * Checks if a row is completely blank
   * @param {Object|Array} row
   * @returns {boolean}
   */
  const isRowBlank = (row) => {
    if (!row) return true;
    if (Array.isArray(row)) {
      return row.every(val => val === null || val === undefined || String(val).trim() === '');
    }
    const values = Object.values(row);
    return values.length === 0 || values.every(val => val === null || val === undefined || String(val).trim() === '');
  };

  /**
   * Parses an Excel or CSV file
   * @param {File} file
   * @param {Object} options
   * @returns {Promise<Object>}
   */
  const parseFile = async (file, options = {}) => {
    const {
      matchMode = 'flexible', // 'flexible' | 'strict'
      trimHeaders = true,
      ignoreBlankRows = true
    } = options;

    const buffer = await readFileAsArrayBuffer(file);
    
    // Read workbook via SheetJS
    const workbook = XLSX.read(buffer, {
      type: 'array',
      cellDates: true,
      cellNF: false,
      cellText: false
    });

    const parsedSheets = [];
    const isStrict = matchMode === 'strict';

    for (const sheetName of workbook.SheetNames) {
      const worksheet = workbook.Sheets[sheetName];
      if (!worksheet || !worksheet['!ref']) continue;

      // Convert sheet to array of arrays first to inspect header row
      const rawRows = XLSX.utils.sheet_to_json(worksheet, {
        header: 1,
        defval: '',
        blankrows: false
      });

      if (!rawRows || rawRows.length === 0) continue;

      // First non-empty row is treated as header
      let headerRowIndex = 0;
      while (headerRowIndex < rawRows.length && isRowBlank(rawRows[headerRowIndex])) {
        headerRowIndex++;
      }

      if (headerRowIndex >= rawRows.length) continue; // No valid header row

      const rawHeaders = rawRows[headerRowIndex].map(h => cleanHeader(h, trimHeaders));
      // Remove trailing completely empty header cells
      while (rawHeaders.length > 0 && rawHeaders[rawHeaders.length - 1] === '') {
        rawHeaders.pop();
      }

      if (rawHeaders.length === 0) continue;

      const { signature, normalizedHeaders } = generateSchemaSignature(rawHeaders, isStrict, trimHeaders);

      // Parse data rows
      const dataRows = [];
      for (let i = headerRowIndex + 1; i < rawRows.length; i++) {
        const rowArr = rawRows[i];
        if (ignoreBlankRows && isRowBlank(rowArr)) continue;

        const rowObj = {};
        for (let c = 0; c < normalizedHeaders.length; c++) {
          const colName = normalizedHeaders[c];
          let val = rowArr[c];
          if (val instanceof Date) {
            // Format dates nicely YYYY-MM-DD
            try {
              val = val.toISOString().split('T')[0];
            } catch (e) {
              // fallback to string
              val = String(val);
            }
          } else if (val === undefined || val === null) {
            val = '';
          }
          rowObj[colName] = val;
        }
        dataRows.push(rowObj);
      }

      parsedSheets.push({
        sheetName,
        headers: normalizedHeaders,
        signature,
        rows: dataRows,
        rowCount: dataRows.length
      });
    }

    return {
      fileId: `${file.name}_${file.size}_${file.lastModified}`,
      fileName: file.name,
      fileSize: file.size,
      fileType: file.name.split('.').pop().toLowerCase(),
      sheets: parsedSheets,
      totalRows: parsedSheets.reduce((sum, s) => sum + s.rowCount, 0)
    };
  };

  return {
    parseFile,
    cleanHeader,
    generateSchemaSignature
  };
})();

// Export to global scope
window.ExcelParser = ExcelParser;
