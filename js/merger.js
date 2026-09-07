/**
 * Excel Merger Module
 * Groups sheets/files with identical schemas and handles export to single/multi-sheet Excel or ZIP.
 */

const ExcelMerger = (() => {
  const SOURCE_COL_NAME = '__출처_파일명';

  /**
   * Groups parsed file data by schema signature
   * @param {Array} parsedFiles - array of parsed file objects from ExcelParser
   * @param {Object} options - { addSourceFileColumn: boolean }
   * @returns {Array} array of group objects
   */
  const groupFilesBySchema = (parsedFiles, options = {}) => {
    const { addSourceFileColumn = true } = options;
    const groupsMap = new Map();

    parsedFiles.forEach(file => {
      file.sheets.forEach(sheet => {
        const sig = sheet.signature;
        if (!groupsMap.has(sig)) {
          groupsMap.set(sig, {
            id: `group_${groupsMap.size + 1}`,
            index: groupsMap.size + 1,
            signature: sig,
            headers: [...sheet.headers],
            files: new Set(),
            fileDetails: [],
            mergedRows: [],
            totalRows: 0
          });
        }

        const group = groupsMap.get(sig);
        group.files.add(file.fileName);
        group.fileDetails.push({
          fileName: file.fileName,
          sheetName: sheet.sheetName,
          rowCount: sheet.rowCount
        });

        // Add rows to group, preserving headers and optionally prepending source column
        const processedRows = sheet.rows.map(row => {
          const rowCopy = {};
          if (addSourceFileColumn) {
            rowCopy[SOURCE_COL_NAME] = file.fileName;
          }
          // Ensure all canonical headers are present in order
          group.headers.forEach(h => {
            rowCopy[h] = row[h] !== undefined ? row[h] : '';
          });
          return rowCopy;
        });

        group.mergedRows.push(...processedRows);
        group.totalRows += sheet.rowCount;
      });
    });

    // Convert map to array and compute display metadata
    return Array.from(groupsMap.values()).map(g => {
      const fileList = Array.from(g.files);
      const displayHeaders = addSourceFileColumn 
        ? [SOURCE_COL_NAME, ...g.headers]
        : [...g.headers];

      // Provide a clean summary title based on first 3 columns or first filename
      const sampleCols = g.headers.slice(0, 3).join(', ');
      const title = `규격 #${g.index} (${g.headers.length}개 열: ${sampleCols}${g.headers.length > 3 ? '...' : ''})`;

      return {
        ...g,
        files: fileList,
        displayHeaders,
        title
      };
    });
  };

  /**
   * Creates a SheetJS worksheet with auto-width columns
   * @param {Array} rows
   * @param {Array} headers
   * @returns {Object} XLSX Worksheet
   */
  const createWorksheetWithStyle = (rows, headers) => {
    const ws = XLSX.utils.json_to_sheet(rows, { header: headers });

    // Auto-calculate column widths
    const colWidths = headers.map(header => {
      let maxLen = String(header).length * 2; // Korean chars take more visual width
      // Sample up to 50 rows for performance
      const sample = rows.slice(0, 50);
      for (const row of sample) {
        const val = row[header];
        if (val !== undefined && val !== null) {
          const strLen = String(val).length;
          if (strLen > maxLen) maxLen = Math.min(strLen * 1.5, 60);
        }
      }
      return { wch: Math.max(maxLen + 4, 12) };
    });

    ws['!cols'] = colWidths;
    return ws;
  };

  /**
   * Sanitizes a sheet name to adhere to Excel's 31-character limit and forbidden characters
   * @param {string} name
   * @returns {string}
   */
  const sanitizeSheetName = (name) => {
    let clean = name.replace(/[\\/*?:[\]]/g, '_');
    if (clean.length > 30) {
      clean = clean.substring(0, 30);
    }
    return clean || 'Sheet1';
  };

  /**
   * Triggers download of a Blob in the browser
   * @param {Blob} blob
   * @param {string} filename
   */
  const triggerDownload = (blob, filename) => {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => {
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    }, 100);
  };

  /**
   * Exports a single group as an Excel (.xlsx) file
   * @param {Object} group
   * @param {string} customName
   */
  const exportGroupAsExcel = (group, customName) => {
    const wb = XLSX.utils.book_new();
    const ws = createWorksheetWithStyle(group.mergedRows, group.displayHeaders);
    
    const sheetName = sanitizeSheetName(`규격_${group.index}_병합데이터`);
    XLSX.utils.book_append_sheet(wb, ws, sheetName);

    const safeCols = group.headers.slice(0, 2).join('_').replace(/[^a-zA-Z0-9가-힣_]/g, '');
    const filename = customName || `병합결과_규격${group.index}_${safeCols || '데이터'}.xlsx`;

    XLSX.writeFile(wb, filename);
  };

  /**
   * Exports all groups as a single Excel workbook, where each group has its own sheet
   * @param {Array} groups
   * @param {string} filename
   */
  const exportAllGroupsAsMultiSheetExcel = (groups, filename = '전체규격_통합_워크북.xlsx') => {
    if (!groups || groups.length === 0) return;

    const wb = XLSX.utils.book_new();

    groups.forEach(group => {
      const ws = createWorksheetWithStyle(group.mergedRows, group.displayHeaders);
      const safeCols = group.headers.slice(0, 2).join('_').replace(/[^a-zA-Z0-9가-힣_]/g, '');
      const sheetName = sanitizeSheetName(`규격${group.index}_${safeCols}`);
      XLSX.utils.book_append_sheet(wb, ws, sheetName);
    });

    XLSX.writeFile(wb, filename);
  };

  /**
   * Exports all groups as individual Excel files inside a single ZIP archive
   * @param {Array} groups
   * @param {string} zipFilename
   */
  const exportAllGroupsAsZip = async (groups, zipFilename = '규격별_엑셀_모음.zip') => {
    if (!groups || groups.length === 0) return;
    if (typeof JSZip === 'undefined') {
      throw new Error('JSZip 라이브러리가 로드되지 않았습니다.');
    }

    const zip = new JSZip();

    for (const group of groups) {
      const wb = XLSX.utils.book_new();
      const ws = createWorksheetWithStyle(group.mergedRows, group.displayHeaders);
      const sheetName = sanitizeSheetName(`규격${group.index}_병합데이터`);
      XLSX.utils.book_append_sheet(wb, ws, sheetName);

      const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
      const safeCols = group.headers.slice(0, 2).join('_').replace(/[^a-zA-Z0-9가-힣_]/g, '');
      const filename = `규격${group.index}_${safeCols}_병합.xlsx`;

      zip.file(filename, wbout);
    }

    const content = await zip.generateAsync({ type: 'blob' });
    triggerDownload(content, zipFilename);
  };

  return {
    SOURCE_COL_NAME,
    groupFilesBySchema,
    exportGroupAsExcel,
    exportAllGroupsAsMultiSheetExcel,
    exportAllGroupsAsZip
  };
})();

// Export to global scope
window.ExcelMerger = ExcelMerger;
