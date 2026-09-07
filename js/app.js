/**
 * ExcelGroup Merger - Main Application Controller
 */

document.addEventListener('DOMContentLoaded', () => {
  // ==========================================
  // 1. Application State
  // ==========================================
  const state = {
    rawFiles: [],        // Raw File objects
    parsedFiles: [],     // Parsed results from ExcelParser
    groups: [],          // Grouped results from ExcelMerger
    options: {
      matchMode: 'flexible',
      addSource: true,
      ignoreBlank: true,
      trimHeaders: true
    },
    modal: {
      activeGroup: null,
      currentPage: 1,
      pageSize: 50,
      filterText: ''
    }
  };

  // ==========================================
  // 2. DOM Elements Selection
  // ==========================================
  const dropZone = document.getElementById('drop-zone');
  const fileInput = document.getElementById('file-input');
  const btnSampleData = document.getElementById('btn-sample-data');
  const btnThemeToggle = document.getElementById('btn-theme-toggle');
  const btnClearAll = document.getElementById('btn-clear-all');

  // Option Controls
  const matchModeSelect = document.getElementById('match-mode');
  const optAddSource = document.getElementById('opt-add-source');
  const optIgnoreBlank = document.getElementById('opt-ignore-blank');
  const optTrimHeaders = document.getElementById('opt-trim-headers');

  // Sections & Indicators
  const uploadedFilesSection = document.getElementById('uploaded-files-section');
  const totalUploadedCount = document.getElementById('total-uploaded-count');
  const fileChipsContainer = document.getElementById('file-chips-container');
  const loadingIndicator = document.getElementById('loading-indicator');
  const loadingText = document.getElementById('loading-text');
  const resultsSection = document.getElementById('results-section');
  const emptyState = document.getElementById('empty-state');
  const groupsContainer = document.getElementById('groups-container');

  // Stats
  const statTotalFiles = document.getElementById('stat-total-files');
  const statTotalGroups = document.getElementById('stat-total-groups');
  const statTotalRows = document.getElementById('stat-total-rows');

  // Batch Export Buttons
  const btnDownloadMultiSheet = document.getElementById('btn-download-multisheet');
  const btnDownloadZip = document.getElementById('btn-download-zip');

  // Modal Elements
  const previewModal = document.getElementById('preview-modal');
  const modalCloseBtn = document.getElementById('btn-modal-close');
  const modalGroupBadge = document.getElementById('modal-group-badge');
  const modalTitle = document.getElementById('modal-title');
  const modalFileCount = document.getElementById('modal-file-count');
  const modalRowCount = document.getElementById('modal-row-count');
  const modalSearchInput = document.getElementById('modal-search-input');
  const previewThead = document.getElementById('preview-thead');
  const previewTbody = document.getElementById('preview-tbody');
  const paginationInfo = document.getElementById('pagination-info');
  const btnPagePrev = document.getElementById('btn-page-prev');
  const btnPageNext = document.getElementById('btn-page-next');
  const pageNumDisplay = document.getElementById('page-num-display');
  const btnModalExport = document.getElementById('btn-modal-export');

  // Toast Container
  const toastContainer = document.getElementById('toast-container');

  // ==========================================
  // 3. UI Helpers: Toasts, Icons, Theme
  // ==========================================
  const renderIcons = () => {
    if (window.lucide) {
      window.lucide.createIcons();
    }
  };

  const showToast = (message, type = 'info', duration = 3500) => {
    const toast = document.createElement('div');
    toast.className = `toast ${type}`;
    
    let iconName = 'info';
    if (type === 'success') iconName = 'check-circle-2';
    if (type === 'warning') iconName = 'alert-triangle';
    if (type === 'error') iconName = 'alert-circle';

    toast.innerHTML = `
      <div class="toast-icon"><i data-lucide="${iconName}"></i></div>
      <div class="toast-message">${message}</div>
    `;

    toastContainer.appendChild(toast);
    renderIcons();

    setTimeout(() => {
      toast.remove();
    }, duration);
  };

  const initTheme = () => {
    const savedTheme = localStorage.getItem('excel_merger_theme') || 'dark';
    document.body.setAttribute('data-theme', savedTheme);
  };

  const toggleTheme = () => {
    const current = document.body.getAttribute('data-theme') || 'dark';
    const next = current === 'dark' ? 'light' : 'dark';
    document.body.setAttribute('data-theme', next);
    localStorage.setItem('excel_merger_theme', next);
    renderIcons();
  };

  // ==========================================
  // 4. File Processing & Grouping Pipeline
  // ==========================================
  const formatFileSize = (bytes) => {
    if (bytes < 1024) return bytes + ' B';
    if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
    return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
  };

  const readSettingsFromUI = () => {
    state.options.matchMode = matchModeSelect.value;
    state.options.addSource = optAddSource.checked;
    state.options.ignoreBlank = optIgnoreBlank.checked;
    state.options.trimHeaders = optTrimHeaders.checked;
  };

  const processAndGroupFiles = async () => {
    if (state.rawFiles.length === 0) {
      state.parsedFiles = [];
      state.groups = [];
      renderUI();
      return;
    }

    loadingIndicator.classList.remove('hidden');
    resultsSection.classList.add('hidden');
    emptyState.classList.add('hidden');

    try {
      readSettingsFromUI();

      // Parse all raw files with current settings
      const parsePromises = state.rawFiles.map((file, idx) => {
        loadingText.textContent = `엑셀 파싱 중 (${idx + 1}/${state.rawFiles.length}): ${file.name}`;
        return ExcelParser.parseFile(file, state.options);
      });

      state.parsedFiles = await Promise.all(parsePromises);

      // Group parsed files by schema
      state.groups = ExcelMerger.groupFilesBySchema(state.parsedFiles, {
        addSourceFileColumn: state.options.addSource
      });

      renderUI();
      showToast(`${state.rawFiles.length}개 파일 분석 완료! ${state.groups.length}개 규격으로 분류되었습니다.`, 'success');
    } catch (err) {
      console.error('Processing error:', err);
      showToast(`파일 처리 중 오류가 발생했습니다: ${err.message}`, 'error', 5000);
    } finally {
      loadingIndicator.classList.add('hidden');
    }
  };

  // ==========================================
  // 5. Render Functions
  // ==========================================
  const renderUI = () => {
    renderFileChips();
    renderSummaryAndGroups();
    renderIcons();
  };

  const renderFileChips = () => {
    if (state.rawFiles.length === 0) {
      uploadedFilesSection.classList.add('hidden');
      return;
    }

    uploadedFilesSection.classList.remove('hidden');
    totalUploadedCount.textContent = state.rawFiles.length;
    fileChipsContainer.innerHTML = '';

    state.rawFiles.forEach((file, index) => {
      const chip = document.createElement('div');
      chip.className = 'file-chip';
      chip.innerHTML = `
        <i data-lucide="file-spreadsheet" class="chip-icon"></i>
        <span class="chip-name" title="${file.name}">${file.name}</span>
        <span class="chip-meta">(${formatFileSize(file.size)})</span>
        <span class="chip-remove" data-index="${index}" title="파일 제거">
          <i data-lucide="x"></i>
        </span>
      `;

      chip.querySelector('.chip-remove').addEventListener('click', (e) => {
        e.stopPropagation();
        removeFileAtIndex(index);
      });

      fileChipsContainer.appendChild(chip);
    });
  };

  const renderSummaryAndGroups = () => {
    if (state.groups.length === 0) {
      resultsSection.classList.add('hidden');
      emptyState.classList.remove('hidden');
      return;
    }

    emptyState.classList.add('hidden');
    resultsSection.classList.remove('hidden');

    // Calculate totals
    const totalFiles = state.rawFiles.length;
    const totalGroups = state.groups.length;
    const totalRows = state.groups.reduce((acc, g) => acc + g.totalRows, 0);

    statTotalFiles.textContent = totalFiles.toLocaleString();
    statTotalGroups.textContent = totalGroups.toLocaleString();
    statTotalRows.textContent = totalRows.toLocaleString();

    // Render Group Cards
    groupsContainer.innerHTML = '';

    state.groups.forEach((group) => {
      const card = document.createElement('div');
      card.className = 'group-card';
      card.setAttribute('data-group-id', group.id);

      // Columns list chips
      const columnChipsHtml = group.headers.map(col => `
        <span class="column-chip">
          <i data-lucide="columns"></i>
          ${col}
        </span>
      `).join('');

      // Files list in group
      const filesListHtml = group.files.map(filename => `
        <span class="group-file-item">
          <i data-lucide="file-check"></i>
          ${filename}
        </span>
      `).join('');

      card.innerHTML = `
        <div class="group-card-header">
          <div class="group-info-left">
            <span class="group-tag">규격 #${group.index}</span>
            <div class="group-title-text">
              <span class="group-name">${group.files.length}개 파일 병합</span>
              <span class="group-meta">총 ${group.totalRows.toLocaleString()}개 데이터 행 &bull; ${group.headers.length}개 컬럼 규격</span>
            </div>
          </div>
          <div class="group-actions">
            <button class="btn btn-sm btn-glass btn-preview-group" data-group-id="${group.id}">
              <i data-lucide="eye"></i>
              <span>미리보기</span>
            </button>
            <button class="btn btn-sm btn-primary btn-export-single" data-group-id="${group.id}">
              <i data-lucide="download"></i>
              <span>엑셀 다운로드</span>
            </button>
          </div>
        </div>

        <div class="schema-columns-section">
          <span class="schema-label">감지된 헤더 규격 (${group.headers.length}개 컬럼)</span>
          <div class="schema-columns-chips">
            ${columnChipsHtml}
          </div>
        </div>

        <div class="group-files-section">
          <div class="group-files-header">
            <span>포함된 파일 목록 (${group.files.length}개)</span>
          </div>
          <div class="group-file-items">
            ${filesListHtml}
          </div>
        </div>
      `;

      // Event: Preview button
      card.querySelector('.btn-preview-group').addEventListener('click', () => {
        openPreviewModal(group);
      });

      // Event: Export single group
      card.querySelector('.btn-export-single').addEventListener('click', () => {
        try {
          ExcelMerger.exportGroupAsExcel(group);
          showToast(`규격 #${group.index} 엑셀 파일이 다운로드되었습니다.`, 'success');
        } catch (e) {
          showToast(`다운로드 실패: ${e.message}`, 'error');
        }
      });

      groupsContainer.appendChild(card);
    });
  };

  // ==========================================
  // 6. Preview Modal Handling & Pagination
  // ==========================================
  const openPreviewModal = (group) => {
    state.modal.activeGroup = group;
    state.modal.currentPage = 1;
    state.modal.filterText = '';
    modalSearchInput.value = '';

    modalGroupBadge.textContent = `규격 #${group.index}`;
    modalTitle.textContent = `${group.files.length}개 파일 병합 결과 미리보기`;
    modalFileCount.textContent = group.files.length;
    modalRowCount.textContent = group.totalRows.toLocaleString();

    renderModalTable();
    previewModal.classList.remove('hidden');
    previewModal.setAttribute('aria-hidden', 'false');
    renderIcons();
  };

  const closePreviewModal = () => {
    previewModal.classList.add('hidden');
    previewModal.setAttribute('aria-hidden', 'true');
    state.modal.activeGroup = null;
  };

  const getFilteredModalRows = () => {
    const group = state.modal.activeGroup;
    if (!group) return [];

    const query = state.modal.filterText.trim().toLowerCase();
    if (!query) return group.mergedRows;

    return group.mergedRows.filter(row => {
      return Object.values(row).some(val => 
        String(val).toLowerCase().includes(query)
      );
    });
  };

  const renderModalTable = () => {
    const group = state.modal.activeGroup;
    if (!group) return;

    const filteredRows = getFilteredModalRows();
    const totalFiltered = filteredRows.length;
    const totalPages = Math.max(1, Math.ceil(totalFiltered / state.modal.pageSize));

    if (state.modal.currentPage > totalPages) {
      state.modal.currentPage = totalPages;
    }

    // Render Table Header
    previewThead.innerHTML = '';
    const trHead = document.createElement('tr');
    
    // Row number column
    const thIdx = document.createElement('th');
    thIdx.textContent = '#';
    thIdx.style.width = '50px';
    trHead.appendChild(thIdx);

    group.displayHeaders.forEach(header => {
      const th = document.createElement('th');
      th.textContent = header;
      if (header === ExcelMerger.SOURCE_COL_NAME) {
        th.classList.add('source-col');
      }
      trHead.appendChild(th);
    });
    previewThead.appendChild(trHead);

    // Slice rows for current page
    const startIdx = (state.modal.currentPage - 1) * state.modal.pageSize;
    const endIdx = Math.min(startIdx + state.modal.pageSize, totalFiltered);
    const pageRows = filteredRows.slice(startIdx, endIdx);

    // Render Table Body
    previewTbody.innerHTML = '';
    if (pageRows.length === 0) {
      const trEmpty = document.createElement('tr');
      trEmpty.innerHTML = `<td colspan="${group.displayHeaders.length + 1}" style="text-align: center; padding: 40px; color: var(--text-muted);">검색 결과와 일치하는 데이터가 없습니다.</td>`;
      previewTbody.appendChild(trEmpty);
    } else {
      pageRows.forEach((row, i) => {
        const tr = document.createElement('tr');
        
        // Row number cell
        const tdIdx = document.createElement('td');
        tdIdx.textContent = (startIdx + i + 1);
        tdIdx.style.color = 'var(--text-muted)';
        tr.appendChild(tdIdx);

        group.displayHeaders.forEach(col => {
          const td = document.createElement('td');
          const val = row[col] !== undefined ? row[col] : '';

          if (col === ExcelMerger.SOURCE_COL_NAME) {
            td.innerHTML = `<span class="source-cell-badge">${val}</span>`;
          } else {
            td.textContent = val;
          }
          tr.appendChild(td);
        });
        previewTbody.appendChild(tr);
      });
    }

    // Update Pagination UI
    paginationInfo.textContent = totalFiltered > 0
      ? `${(startIdx + 1).toLocaleString()} - ${endIdx.toLocaleString()} / 총 ${totalFiltered.toLocaleString()}개 행`
      : '0 / 0개 행';
    
    pageNumDisplay.textContent = `${state.modal.currentPage} / ${totalPages}`;
    btnPagePrev.disabled = state.modal.currentPage <= 1;
    btnPageNext.disabled = state.modal.currentPage >= totalPages;
  };

  // ==========================================
  // 7. File Ingestion Helpers
  // ==========================================
  const addFiles = (files) => {
    const validExtensions = ['.xlsx', '.xls', '.csv'];
    const newFiles = [];

    Array.from(files).forEach(file => {
      const ext = '.' + file.name.split('.').pop().toLowerCase();
      if (validExtensions.includes(ext)) {
        // Prevent exact duplicate by filename and size
        const isDuplicate = state.rawFiles.some(f => f.name === file.name && f.size === file.size);
        if (!isDuplicate) {
          newFiles.push(file);
        }
      }
    });

    if (newFiles.length === 0) {
      showToast('지원되는 새 엑셀/CSV 파일이 없거나 이미 등록된 파일입니다.', 'warning');
      return;
    }

    state.rawFiles.push(...newFiles);
    processAndGroupFiles();
  };

  const removeFileAtIndex = (index) => {
    if (index >= 0 && index < state.rawFiles.length) {
      const removed = state.rawFiles.splice(index, 1);
      showToast(`'${removed[0].name}' 파일이 제외되었습니다.`, 'info');
      processAndGroupFiles();
    }
  };

  const clearAllFiles = () => {
    if (state.rawFiles.length === 0) return;
    state.rawFiles = [];
    state.parsedFiles = [];
    state.groups = [];
    renderUI();
    showToast('모든 파일이 초기화되었습니다.', 'info');
  };

  // ==========================================
  // 8. Event Listeners Setup
  // ==========================================
  
  // Theme Toggle
  btnThemeToggle.addEventListener('click', toggleTheme);

  // Drag & Drop
  ['dragenter', 'dragover'].forEach(eventName => {
    dropZone.addEventListener(eventName, (e) => {
      e.preventDefault();
      e.stopPropagation();
      dropZone.classList.add('drag-active');
    }, false);
  });

  ['dragleave', 'drop'].forEach(eventName => {
    dropZone.addEventListener(eventName, (e) => {
      e.preventDefault();
      e.stopPropagation();
      dropZone.classList.remove('drag-active');
    }, false);
  });

  dropZone.addEventListener('drop', (e) => {
    const dt = e.dataTransfer;
    if (dt && dt.files && dt.files.length > 0) {
      addFiles(dt.files);
    }
  });

  dropZone.addEventListener('click', () => {
    fileInput.click();
  });

  fileInput.addEventListener('change', (e) => {
    if (e.target.files && e.target.files.length > 0) {
      addFiles(e.target.files);
      fileInput.value = ''; // Reset input to allow re-uploading same file if desired
    }
  });

  // Clear all button
  btnClearAll.addEventListener('click', clearAllFiles);

  // Settings change events - re-run grouping immediately
  [matchModeSelect, optAddSource, optIgnoreBlank, optTrimHeaders].forEach(el => {
    el.addEventListener('change', () => {
      if (state.rawFiles.length > 0) {
        showToast('설정이 변경되어 규격을 다시 분류합니다.', 'info');
        processAndGroupFiles();
      }
    });
  });

  // Sample data button
  btnSampleData.addEventListener('click', () => {
    try {
      const sampleFiles = SampleDataGenerator.generateSampleFiles();
      addFiles(sampleFiles);
      showToast('체험용 샘플 엑셀 파일 5개가 생성되어 자동 분류되었습니다!', 'success');
    } catch (e) {
      console.error(e);
      showToast('샘플 생성 중 오류: ' + e.message, 'error');
    }
  });

  // Batch Export: Multi-Sheet Excel
  btnDownloadMultiSheet.addEventListener('click', () => {
    try {
      if (state.groups.length === 0) return;
      ExcelMerger.exportAllGroupsAsMultiSheetExcel(state.groups, '전체규격_통합_엑셀.xlsx');
      showToast('모든 규격이 시트별로 포함된 통합 엑셀 파일이 다운로드되었습니다.', 'success');
    } catch (e) {
      showToast(`통합 다운로드 실패: ${e.message}`, 'error');
    }
  });

  // Batch Export: ZIP Archive
  btnDownloadZip.addEventListener('click', async () => {
    try {
      if (state.groups.length === 0) return;
      showToast('규격별 엑셀 파일을 압축하고 있습니다...', 'info');
      await ExcelMerger.exportAllGroupsAsZip(state.groups, '규격별_엑셀_모음.zip');
      showToast('규격별 엑셀 파일들이 ZIP으로 다운로드되었습니다.', 'success');
    } catch (e) {
      showToast(`ZIP 다운로드 실패: ${e.message}`, 'error');
    }
  });

  // Modal events
  modalCloseBtn.addEventListener('click', closePreviewModal);
  previewModal.addEventListener('click', (e) => {
    if (e.target === previewModal) {
      closePreviewModal();
    }
  });

  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape' && !previewModal.classList.contains('hidden')) {
      closePreviewModal();
    }
  });

  modalSearchInput.addEventListener('input', (e) => {
    state.modal.filterText = e.target.value;
    state.modal.currentPage = 1;
    renderModalTable();
  });

  btnPagePrev.addEventListener('click', () => {
    if (state.modal.currentPage > 1) {
      state.modal.currentPage--;
      renderModalTable();
    }
  });

  btnPageNext.addEventListener('click', () => {
    const totalFiltered = getFilteredModalRows().length;
    const totalPages = Math.ceil(totalFiltered / state.modal.pageSize);
    if (state.modal.currentPage < totalPages) {
      state.modal.currentPage++;
      renderModalTable();
    }
  });

  btnModalExport.addEventListener('click', () => {
    if (state.modal.activeGroup) {
      ExcelMerger.exportGroupAsExcel(state.modal.activeGroup);
      showToast(`규격 #${state.modal.activeGroup.index} 엑셀 파일이 다운로드되었습니다.`, 'success');
    }
  });

  // Initialize
  initTheme();
  renderIcons();
});
