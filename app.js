/**
 * Smile PriceTag Master - Application Engine
 * High-performance, zero-dependency retail price tag generator
 * Features:
 * 1. A4 Portrait Multi-Tag Rendering (4x4 = 16 tags, 3x4 = 12 tags, etc.)
 * 2. Thick Big Black Typography with crisp white stroke matching reference
 * 3. Excel (.xlsx/.xls) Parser:
 *    - Pure client-side ZIP Central Directory + DecompressionStream
 *    - Server API binary stream fallback
 *    - HTML Table / TSV fallback for POS exported .xls
 * 4. CSV / TSV streaming parser
 * 5. Google Sheets direct fetch & quick-paste table importer
 * 6. Folder Image Auto-Detection with fuzzy filename matching
 * 7. Google Drive Image Link Resolver
 * 8. In-Place Table Editor (edit any product after upload)
 * 9. High-DPI 300 DPI Canvas Rendering & Print Engine
 * 10. LocalStorage Auto-Persistence
 */

const DEFAULT_SAMPLE_PRODUCTS = [
  { id: '1', title: 'OFFER LITCHI', mrp: 309.0, offerPrice: 149.0, unit: 'KG', storeName: 'SMILE', image: '', imageScale: 1.0, imageX: 0, imageY: 0 },
  { id: '2', title: 'ROYAL APPLE GALA', mrp: 240.0, offerPrice: 169.0, unit: 'KG', storeName: 'SMILE', image: '', imageScale: 1.0, imageX: 0, imageY: 0 },
  { id: '3', title: 'ROBUSTA BANANA', mrp: 60.0, offerPrice: 39.0, unit: 'KG', storeName: 'SMILE', image: '', imageScale: 1.0, imageX: 0, imageY: 0 },
  { id: '4', title: 'NAGPUR ORANGE', mrp: 120.0, offerPrice: 79.0, unit: 'KG', storeName: 'SMILE', image: '', imageScale: 1.0, imageX: 0, imageY: 0 },
  { id: '5', title: 'ALPHONSO MANGO', mrp: 450.0, offerPrice: 299.0, unit: 'KG', storeName: 'SMILE', image: '', imageScale: 1.0, imageX: 0, imageY: 0 },
  { id: '6', title: 'NASHIK RED ONION', mrp: 45.0, offerPrice: 32.0, unit: 'KG', storeName: 'SMILE', image: '', imageScale: 1.0, imageX: 0, imageY: 0 },
  { id: '7', title: 'AGRA FRESH POTATO', mrp: 35.0, offerPrice: 24.0, unit: 'KG', storeName: 'SMILE', image: '', imageScale: 1.0, imageX: 0, imageY: 0 },
  { id: '8', title: 'HYBRID TOMATO', mrp: 40.0, offerPrice: 28.0, unit: 'KG', storeName: 'SMILE', image: '', imageScale: 1.0, imageX: 0, imageY: 0 },
  { id: '9', title: 'TONED FRESH MILK', mrp: 34.0, offerPrice: 29.0, unit: 'PKT', storeName: 'SMILE', image: '', imageScale: 1.0, imageX: 0, imageY: 0 },
  { id: '10', title: 'SONAKA GREEN GRAPES', mrp: 130.0, offerPrice: 89.0, unit: 'KG', storeName: 'SMILE', image: '', imageScale: 1.0, imageX: 0, imageY: 0 },
  { id: '11', title: 'KIRAN WATERMELON', mrp: 35.0, offerPrice: 25.0, unit: 'KG', storeName: 'SMILE', image: '', imageScale: 1.0, imageX: 0, imageY: 0 },
  { id: '12', title: 'SWEET STRAWBERRY', mrp: 99.0, offerPrice: 69.0, unit: 'BOX', storeName: 'SMILE', image: '', imageScale: 1.0, imageX: 0, imageY: 0 },
  { id: '13', title: 'RANI PINEAPPLE', mrp: 80.0, offerPrice: 59.0, unit: 'PC', storeName: 'SMILE', image: '', imageScale: 1.0, imageX: 0, imageY: 0 },
  { id: '14', title: 'ORANGE CARROT', mrp: 55.0, offerPrice: 39.0, unit: 'KG', storeName: 'SMILE', image: '', imageScale: 1.0, imageX: 0, imageY: 0 },
  { id: '15', title: 'SALAD CUCUMBER', mrp: 40.0, offerPrice: 29.0, unit: 'KG', storeName: 'SMILE', image: '', imageScale: 1.0, imageX: 0, imageY: 0 },
  { id: '16', title: 'GREEN CAPSICUM', mrp: 70.0, offerPrice: 48.0, unit: 'KG', storeName: 'SMILE', image: '', imageScale: 1.0, imageX: 0, imageY: 0 },
  { id: '17', title: 'RED CHILLI', mrp: 30.0, offerPrice: 19.0, unit: 'PKT', storeName: 'SMILE', image: '', imageScale: 1.0, imageX: 0, imageY: 0 },
  { id: '18', title: 'FRESH CORIANDER', mrp: 25.0, offerPrice: 15.0, unit: 'BUNCH', storeName: 'SMILE', image: '', imageScale: 1.0, imageX: 0, imageY: 0 },
  { id: '19', title: 'ORGANIC GINGER', mrp: 70.0, offerPrice: 45.0, unit: 'KG', storeName: 'SMILE', image: '', imageScale: 1.0, imageX: 0, imageY: 0 },
  { id: '20', title: 'GARLIC PREMIUM', mrp: 95.0, offerPrice: 65.0, unit: 'KG', storeName: 'SMILE', image: '', imageScale: 1.0, imageX: 0, imageY: 0 }
];

function colLetterToIndex(colStr) {
  let idx = 0;
  for (let i = 0; i < colStr.length; i++) {
    const code = colStr.charCodeAt(i);
    if (code >= 65 && code <= 90) {
      idx = idx * 26 + (code - 64);
    } else if (code >= 97 && code <= 122) {
      idx = idx * 26 + (code - 96);
    }
  }
  return Math.max(0, idx - 1);
}

class PriceTagApp {
  constructor() {
    this.tags = [];
    this.settings = {
      storeName: 'SMILE',
      currency: '₹',
      grid: '4x5',
      zoom: 65,
      showCutGuides: true,
      globalFontScale: 135,
      globalTitleSize: 100,
      globalPriceSize: 100,
      globalStrokeWidth: 3,
      mrpBadgeMode: 'auto'
    };
    this.selectedTagIndex = 0;
    this.pendingFolderMatches = [];

    this.initElements();
    this.attachEventListeners();
    this.loadInitialData();
    setTimeout(() => this.fitA4Zoom(), 150);
  }

  // =========================================================================
  // DOM ELEMENT REFERENCES
  // =========================================================================
  initElements() {
    // Header & Global Settings
    this.elGlobalStoreName = document.getElementById('globalStoreName');
    this.elGlobalCurrency = document.getElementById('globalCurrency');
    this.elGridPreset = document.getElementById('gridPreset');
    this.elGlobalFontScale = document.getElementById('globalFontScale');
    this.elMrpBadgeToggle = document.getElementById('mrpBadgeToggle');
    this.elBtnPrint = document.getElementById('btnPrint');
    this.elBtnDownloadA4Sheet = document.getElementById('btnDownloadA4Sheet');
    this.elBtnDownloadA4Text = document.getElementById('btnDownloadA4Text');
    this.elBtnSheetDownloadA4PNG = document.getElementById('btnSheetDownloadA4PNG');
    this.elBtnSheetDownloadA4Text = document.getElementById('btnSheetDownloadA4Text');
    this.elSheetDownloadAllWrapper = document.getElementById('sheetDownloadAllWrapper');
    this.elBtnSheetDownloadAllA4PNG = document.getElementById('btnSheetDownloadAllA4PNG');
    this.elBtnSheetDownloadAllText = document.getElementById('btnSheetDownloadAllText');
    this.elBtnExportPNG = document.getElementById('btnDownloadPNG') || document.getElementById('btnExportPNG');
    this.elBtnSheetDownloadPNG = document.getElementById('btnSheetDownloadPNG');

    // Ribbon & Actions
    this.elFileUploadInput = document.getElementById('fileUploadInput');
    this.elFolderUploadInput = document.getElementById('folderUploadInput');
    this.elBtnGoogleSheets = document.getElementById('btnGoogleSheets');
    this.elBtnDriveImages = document.getElementById('btnDriveImages');
    this.elBtnAddTag = document.getElementById('btnAddTag');
    this.elBtnClearAllMrp = document.getElementById('btnClearAllMrp');
    this.elBtnResetSample = document.getElementById('btnResetSample');
    this.elBtnClearAll = document.getElementById('btnClearAll');

    // Badges & Counters
    this.elTagCountBadge = document.getElementById('tagCountBadge');
    this.elTableCountBadge = document.getElementById('tableCountBadge');
    this.elVisibleRowCount = document.getElementById('visibleRowCount');
    this.elTotalSavingsSum = document.getElementById('totalSavingsSum');

    // View Controls
    this.elA4Stage = document.getElementById('a4Stage');
    this.elCutGuidesToggle = document.getElementById('cutGuidesToggle');
    this.elZoomSlider = document.getElementById('zoomSlider');
    this.elZoomVal = document.getElementById('zoomVal');
    this.elBtnZoomFit = document.getElementById('btnZoomFit');
    this.elCurrentPageNum = document.getElementById('currentPageNum');
    this.elTotalPagesNum = document.getElementById('totalPagesNum');

    // Table Editor
    this.elTableBody = document.getElementById('tableBody');
    this.elTableSearchInput = document.getElementById('tableSearchInput');
    this.elBtnQuickAddRow = document.getElementById('btnQuickAddRow');
    this.elBtnExportCSV = document.getElementById('btnExportCSV');

    // Designer Studio
    this.elDesignerTagSelector = document.getElementById('designerTagSelector');
    this.elDesignerTagContainer = document.getElementById('designerTagContainer');
    this.elDtTitle = document.getElementById('dtTitle');
    this.elDtMrp = document.getElementById('dtMrp');
    this.elDtOffer = document.getElementById('dtOffer');
    this.elDtUnit = document.getElementById('dtUnit');
    this.elDtSavings = document.getElementById('dtSavings');
    this.elDtStoreName = document.getElementById('dtStoreName');
    this.elDtImageUpload = document.getElementById('dtImageUpload');
    this.elDtZoomSlider = document.getElementById('dtZoomSlider');
    this.elDtZoomVal = document.getElementById('dtZoomVal');
    this.elDtPanXSlider = document.getElementById('dtPanXSlider');
    this.elDtPanXVal = document.getElementById('dtPanXVal');
    this.elDtPanYSlider = document.getElementById('dtPanYSlider');
    this.elDtPanYVal = document.getElementById('dtPanYVal');
    this.elBtnResetImagePos = document.getElementById('btnResetImagePos');
    this.elDtTitleSize = document.getElementById('dtTitleSize');
    this.elDtTitleSizeVal = document.getElementById('dtTitleSizeVal');
    this.elDtPriceSize = document.getElementById('dtPriceSize');
    this.elDtPriceSizeVal = document.getElementById('dtPriceSizeVal');
    this.elDtStrokeWidth = document.getElementById('dtStrokeWidth');
    this.elDtStrokeVal = document.getElementById('dtStrokeVal');
    this.elBtnDownloadSinglePNG = document.getElementById('btnDownloadSinglePNG');

    // Modals
    this.elModalGoogleSheets = document.getElementById('modalGoogleSheets');
    this.elBtnCloseGSheetModal = document.getElementById('btnCloseGSheetModal');
    this.elGSheetUrlInput = document.getElementById('gsheetUrlInput');
    this.elBtnFetchGSheet = document.getElementById('btnFetchGSheet');
    this.elGSheetPasteTextarea = document.getElementById('gsheetPasteTextarea');
    this.elBtnApplyPaste = document.getElementById('btnApplyPaste');

    this.elModalFolderMatch = document.getElementById('modalFolderMatch');
    this.elBtnCloseFolderModal = document.getElementById('btnCloseFolderModal');
    this.elFolderMatchSummary = document.getElementById('folderMatchSummary');
    this.elFolderMatchList = document.getElementById('folderMatchList');
    this.elBtnConfirmMatches = document.getElementById('btnConfirmMatches');
    this.elBtnCancelMatches = document.getElementById('btnCancelMatches');

    // Google Drive Folder Modal
    this.elModalDriveFolder = document.getElementById('modalDriveFolder');
    this.elBtnCloseDriveModal = document.getElementById('btnCloseDriveModal');
    this.elDriveFolderUrlInput = document.getElementById('driveFolderUrlInput');
    this.elDriveApiKeyInput = document.getElementById('driveApiKeyInput');
    this.elBtnScanDriveFolder = document.getElementById('btnScanDriveFolder');
    this.elBtnScanDriveText = document.getElementById('btnScanDriveText');

    this.elModalQuickEdit = document.getElementById('modalQuickEdit');
    this.elBtnCloseQuickEdit = document.getElementById('btnCloseQuickEdit');
    this.elQeTagId = document.getElementById('qeTagId');
    this.elQeTitle = document.getElementById('qeTitle');
    this.elQeMrp = document.getElementById('qeMrp');
    this.elQeOffer = document.getElementById('qeOffer');
    this.elQeUnit = document.getElementById('qeUnit');
    this.elQeSavings = document.getElementById('qeSavings');
    this.elQeImageUpload = document.getElementById('qeImageUpload');
    this.elBtnSaveQuickEdit = document.getElementById('btnSaveQuickEdit');
    this.elBtnOpenInDesigner = document.getElementById('btnOpenInDesigner');
    this.elBtnDownloadQePNG = document.getElementById('btnDownloadQePNG');

    // Alert Banner
    this.elStatusAlert = document.getElementById('statusAlert');
    this.elStatusIcon = document.getElementById('statusIcon');
    this.elStatusMessage = document.getElementById('statusMessage');
    this.elBtnCloseStatus = document.getElementById('btnCloseStatus');
  }

  // =========================================================================
  // EVENT LISTENERS
  // =========================================================================
  attachEventListeners() {
    // Tab switching
    document.querySelectorAll('.tab-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        const tabId = btn.getAttribute('data-tab');
        this.switchTab(tabId);
        if (tabId === 'tab-sheet') {
          setTimeout(() => this.fitA4Zoom(), 60);
        }
      });
    });

    // Window resize listener to keep A4 sheet fitted to window
    window.addEventListener('resize', () => {
      this.fitA4Zoom();
    });

    // Global Settings
    this.elGlobalStoreName.addEventListener('input', () => {
      this.settings.storeName = this.elGlobalStoreName.value.trim() || 'SMILE';
      this.tags.forEach(t => t.storeName = this.settings.storeName);
      this.saveToStorage();
      this.render();
    });

    this.elGlobalCurrency.addEventListener('change', () => {
      this.settings.currency = this.elGlobalCurrency.value;
      document.querySelectorAll('.curr-symbol').forEach(el => el.textContent = this.settings.currency || '');
      this.saveToStorage();
      this.render();
    });

    this.elGridPreset.addEventListener('change', () => {
      this.settings.grid = this.elGridPreset.value;
      this.saveToStorage();
      this.renderA4Sheets();
    });

    if (this.elGlobalFontScale) {
      this.elGlobalFontScale.addEventListener('change', () => {
        this.settings.globalFontScale = parseInt(this.elGlobalFontScale.value, 10) || 115;
        this.saveToStorage();
        this.renderA4Sheets();
        this.renderDesignerTag();
      });
    }

    if (this.elMrpBadgeToggle) {
      this.elMrpBadgeToggle.addEventListener('change', () => {
        this.settings.mrpBadgeMode = this.elMrpBadgeToggle.value;
        this.saveToStorage();
        this.renderA4Sheets();
        this.renderDesignerTag();
        this.renderTableRows();
        this.notify(
          this.settings.mrpBadgeMode === 'hide'
            ? 'MRP and Smile Save badges are now HIDDEN on all tags.'
            : 'MRP and Smile Save badges enabled (will only show for items with MRP).',
          'info'
        );
      });
    }

    this.elCutGuidesToggle.addEventListener('change', () => {
      this.settings.showCutGuides = this.elCutGuidesToggle.checked;
      this.renderA4Sheets();
    });

    this.elZoomSlider.addEventListener('input', () => {
      this.settings.zoom = parseInt(this.elZoomSlider.value, 10);
      this.elZoomVal.textContent = `${this.settings.zoom}%`;
      this.applyA4Zoom();
    });

    this.elBtnZoomFit.addEventListener('click', () => {
      this.fitA4Zoom();
    });

    // Action Buttons
    this.elBtnPrint.addEventListener('click', () => {
      window.print();
    });

    if (this.elBtnDownloadA4Sheet) {
      this.elBtnDownloadA4Sheet.addEventListener('click', () => {
        this.downloadAllA4Sheets();
      });
    }

    if (this.elBtnSheetDownloadA4PNG) {
      this.elBtnSheetDownloadA4PNG.addEventListener('click', () => {
        const config = this.getGridConfig();
        const totalPages = Math.max(1, Math.ceil(this.tags.length / config.perPage));
        if (totalPages > 1) {
          this.downloadAllA4Sheets();
        } else {
          this.exportA4SheetAsPNG(0, 1);
        }
      });
    }

    if (this.elBtnSheetDownloadAllA4PNG) {
      this.elBtnSheetDownloadAllA4PNG.addEventListener('click', () => {
        this.downloadAllA4Sheets();
      });
    }

    if (this.elBtnExportPNG) {
      this.elBtnExportPNG.addEventListener('click', () => {
        const tag = this.tags[this.selectedTagIndex] || this.tags[0];
        if (tag) {
          this.exportSingleTagAsPNG(tag);
        } else {
          this.notify('No tags available to download.', 'warning');
        }
      });
    }

    if (this.elBtnSheetDownloadPNG) {
      this.elBtnSheetDownloadPNG.addEventListener('click', () => {
        const tag = this.tags[this.selectedTagIndex] || this.tags[0];
        if (tag) {
          this.exportSingleTagAsPNG(tag);
        } else {
          this.notify('No tags available to download.', 'warning');
        }
      });
    }

    // Ribbon Import Actions
    this.elFileUploadInput.addEventListener('change', (e) => this.handleFileUpload(e));
    this.elFolderUploadInput.addEventListener('change', (e) => this.handleFolderUpload(e));
    this.elBtnGoogleSheets.addEventListener('click', () => this.openModal(this.elModalGoogleSheets));
    this.elBtnDriveImages.addEventListener('click', () => this.openModal(this.elModalDriveFolder));
    if (this.elBtnCloseDriveModal) {
      this.elBtnCloseDriveModal.addEventListener('click', () => this.closeModal(this.elModalDriveFolder));
    }
    if (this.elBtnScanDriveFolder) {
      this.elBtnScanDriveFolder.addEventListener('click', () => this.scanGoogleDriveFolder());
    }
    this.elBtnAddTag.addEventListener('click', () => this.addNewTag());
    if (this.elBtnClearAllMrp) {
      this.elBtnClearAllMrp.addEventListener('click', () => {
        if (confirm('Remove MRP from all tags? Both Yellow MRP and Red Smile Save badges will disappear.')) {
          this.tags.forEach(t => { t.mrp = 0; });
          this.settings.mrpBadgeMode = 'hide';
          if (this.elMrpBadgeToggle) this.elMrpBadgeToggle.value = 'hide';
          this.saveToStorage();
          this.render();
          this.notify('Removed MRP from all tags successfully!', 'success');
        }
      });
    }
    this.elBtnResetSample.addEventListener('click', () => this.loadSampleData(true));
    this.elBtnClearAll.addEventListener('click', () => {
      if (confirm('Clear all tags from the sheet?')) {
        this.tags = [];
        this.saveToStorage();
        this.render();
      }
    });

    // Modal Close buttons
    this.elBtnCloseGSheetModal.addEventListener('click', () => this.closeModal(this.elModalGoogleSheets));
    this.elBtnCloseFolderModal.addEventListener('click', () => this.closeModal(this.elModalFolderMatch));
    this.elBtnCancelMatches.addEventListener('click', () => this.closeModal(this.elModalFolderMatch));
    this.elBtnCloseQuickEdit.addEventListener('click', () => this.closeModal(this.elModalQuickEdit));
    this.elBtnCloseStatus.addEventListener('click', () => this.elStatusAlert.style.display = 'none');

    // Google Sheets Modal Tabs & Actions
    document.querySelectorAll('.modal-tab-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        document.querySelectorAll('.modal-tab-btn').forEach(b => b.classList.remove('active'));
        document.querySelectorAll('.modal-tab-content').forEach(c => c.classList.remove('active'));
        btn.classList.add('active');
        const target = document.getElementById(btn.dataset.modaltab);
        if (target) target.classList.add('active');
      });
    });

    this.elBtnFetchGSheet.addEventListener('click', () => this.fetchGoogleSheetData());
    this.elBtnApplyPaste.addEventListener('click', () => this.applyPastedSheetData());
    this.elBtnConfirmMatches.addEventListener('click', () => this.applyFolderMatches());

    // Table search & actions
    this.elTableSearchInput.addEventListener('input', () => this.renderTableRows());
    this.elBtnQuickAddRow.addEventListener('click', () => this.addNewTag());
    this.elBtnExportCSV.addEventListener('click', () => this.exportCSV());

    // Single Tag Designer Studio Controls
    this.elDesignerTagSelector.addEventListener('change', () => {
      this.selectedTagIndex = parseInt(this.elDesignerTagSelector.value, 10);
      this.populateDesignerForm();
      this.renderDesignerTag();
    });

    this.elDtTitle.addEventListener('input', () => this.updateSelectedTagField('title', this.elDtTitle.value));
    this.elDtMrp.addEventListener('input', () => this.updateSelectedTagField('mrp', parseFloat(this.elDtMrp.value) || 0));
    this.elDtOffer.addEventListener('input', () => this.updateSelectedTagField('offerPrice', parseFloat(this.elDtOffer.value) || 0));
    this.elDtUnit.addEventListener('input', () => this.updateSelectedTagField('unit', this.elDtUnit.value.toUpperCase()));
    this.elDtStoreName.addEventListener('input', () => this.updateSelectedTagField('storeName', this.elDtStoreName.value));

    this.elDtImageUpload.addEventListener('change', (e) => {
      const file = e.target.files[0];
      if (file) {
        const reader = new FileReader();
        reader.onload = (ev) => {
          this.updateSelectedTagField('image', ev.target.result);
          this.renderDesignerTag();
        };
        reader.readAsDataURL(file);
      }
    });

    this.elDtZoomSlider.addEventListener('input', () => {
      const val = parseInt(this.elDtZoomSlider.value, 10);
      this.elDtZoomVal.textContent = `${val}%`;
      this.updateSelectedTagField('imageScale', val / 100);
      this.renderDesignerTag();
    });

    this.elDtPanXSlider.addEventListener('input', () => {
      const val = parseInt(this.elDtPanXSlider.value, 10);
      this.elDtPanXVal.textContent = `${val}px`;
      this.updateSelectedTagField('imageX', val);
      this.renderDesignerTag();
    });

    this.elDtPanYSlider.addEventListener('input', () => {
      const val = parseInt(this.elDtPanYSlider.value, 10);
      this.elDtPanYVal.textContent = `${val}px`;
      this.updateSelectedTagField('imageY', val);
      this.renderDesignerTag();
    });

    this.elBtnResetImagePos.addEventListener('click', () => {
      this.elDtZoomSlider.value = 100;
      this.elDtZoomVal.textContent = '100%';
      this.elDtPanXSlider.value = 0;
      this.elDtPanXVal.textContent = '0px';
      this.elDtPanYSlider.value = 0;
      this.elDtPanYVal.textContent = '0px';
      const tag = this.tags[this.selectedTagIndex];
      if (tag) {
        tag.imageScale = 1.0;
        tag.imageX = 0;
        tag.imageY = 0;
        this.saveToStorage();
        this.renderDesignerTag();
        this.renderA4Sheets();
      }
    });

    this.elDtTitleSize.addEventListener('input', () => {
      const val = parseInt(this.elDtTitleSize.value, 10);
      this.elDtTitleSizeVal.textContent = `${val}%`;
      this.updateSelectedTagField('titleSize', val);
      this.renderDesignerTag();
    });

    this.elDtPriceSize.addEventListener('input', () => {
      const val = parseInt(this.elDtPriceSize.value, 10);
      this.elDtPriceSizeVal.textContent = `${val}%`;
      this.updateSelectedTagField('priceSize', val);
      this.renderDesignerTag();
    });

    this.elDtStrokeWidth.addEventListener('input', () => {
      const val = parseInt(this.elDtStrokeWidth.value, 10);
      this.elDtStrokeVal.textContent = `${val}px`;
      this.updateSelectedTagField('strokeWidth', val);
      this.renderDesignerTag();
    });

    this.elBtnDownloadSinglePNG.addEventListener('click', () => {
      const tag = this.tags[this.selectedTagIndex];
      if (tag) this.exportSingleTagAsPNG(tag);
    });

    // Quick Edit modal save
    this.elBtnSaveQuickEdit.addEventListener('click', () => {
      const idx = parseInt(this.elQeTagId.textContent, 10) - 1;
      const tag = this.tags[idx];
      if (tag) {
        tag.title = this.elQeTitle.value.trim().toUpperCase();
        tag.mrp = parseFloat(this.elQeMrp.value) || 0;
        tag.offerPrice = parseFloat(this.elQeOffer.value) || 0;
        tag.unit = this.elQeUnit.value.trim().toUpperCase();
        this.closeModal(this.elModalQuickEdit);
        this.saveToStorage();
        this.render();
      }
    });

    this.elBtnOpenInDesigner.addEventListener('click', () => {
      const idx = parseInt(this.elQeTagId.textContent, 10) - 1;
      this.closeModal(this.elModalQuickEdit);
      this.switchTab('tab-designer');
      this.elDesignerTagSelector.value = idx;
      this.selectedTagIndex = idx;
      this.populateDesignerForm();
      this.renderDesignerTag();
    });

    if (this.elBtnDownloadQePNG) {
      this.elBtnDownloadQePNG.addEventListener('click', () => {
        const idx = parseInt(this.elQeTagId.textContent, 10) - 1;
        const tag = this.tags[idx];
        if (tag) {
          tag.title = this.elQeTitle.value.trim().toUpperCase() || tag.title;
          tag.mrp = parseFloat(this.elQeMrp.value) || tag.mrp;
          tag.offerPrice = parseFloat(this.elQeOffer.value) || tag.offerPrice;
          tag.unit = this.elQeUnit.value.trim().toUpperCase() || tag.unit;
          this.exportSingleTagAsPNG(tag);
        }
      });
    }

    this.elQeImageUpload.addEventListener('change', (e) => {
      const file = e.target.files[0];
      if (file) {
        const reader = new FileReader();
        reader.onload = (ev) => {
          const idx = parseInt(this.elQeTagId.textContent, 10) - 1;
          if (this.tags[idx]) {
            this.tags[idx].image = ev.target.result;
            this.saveToStorage();
            this.render();
          }
        };
        reader.readAsDataURL(file);
      }
    });

    const updateQeSavings = () => {
      const mrp = parseFloat(this.elQeMrp.value) || 0;
      const offer = parseFloat(this.elQeOffer.value) || 0;
      if (mrp > 0) {
        const diff = Math.max(0, mrp - offer);
        this.elQeSavings.value = `${this.settings.currency}${diff.toFixed(2)}`;
      } else {
        this.elQeSavings.value = 'No MRP (No badge)';
      }
    };
    this.elQeMrp.addEventListener('input', updateQeSavings);
    this.elQeOffer.addEventListener('input', updateQeSavings);

    // Keyboard Shortcuts
    window.addEventListener('keydown', (e) => {
      if ((e.ctrlKey || e.metaKey) && e.key === 'p') {
        e.preventDefault();
        window.print();
      }
    });
  }

  // =========================================================================
  // STORAGE PERSISTENCE
  // =========================================================================
  saveToStorage() {
    try {
      localStorage.setItem('smile_pricetags_data', JSON.stringify({
        tags: this.tags,
        settings: this.settings
      }));
    } catch (e) {
      // Ignore quota errors if images are large base64
    }
  }

  loadInitialData() {
    try {
      const saved = localStorage.getItem('smile_pricetags_data');
      if (saved) {
        const parsed = JSON.parse(saved);
        if (parsed.tags && parsed.tags.length > 0) {
          this.tags = parsed.tags;
          if (parsed.settings) Object.assign(this.settings, parsed.settings);

          // Always enable auto badge mode so MRP and Saving badges are visible
          this.settings.mrpBadgeMode = 'auto';

          // If stored tags have 0 MRP from previous session, restore default MRPs so MRP and Saving work immediately
          if (!this.tags.some(t => Number(t.mrp) > 0)) {
            DEFAULT_SAMPLE_PRODUCTS.forEach((sample, idx) => {
              if (this.tags[idx] && this.tags[idx].title === sample.title) {
                this.tags[idx].mrp = sample.mrp;
              }
            });
            this.saveToStorage();
          }

          // Clear any old sample produce images from previous sessions so products start clean without images
          let cleanedSampleImgs = false;
          this.tags.forEach(t => {
            if (t.image && (t.image.includes('assets/sample_images') || t.image.includes('assets/litchi.jpg'))) {
              t.image = '';
              cleanedSampleImgs = true;
            }
          });
          if (cleanedSampleImgs) {
            this.saveToStorage();
          }

          this.syncSettingsUI();
          this.render();
          this.notify(`Loaded ${this.tags.length} price tags.`, 'info');
          return;
        }
      }
    } catch (e) {
      console.warn('Could not read from localStorage:', e);
    }

    this.loadSampleData(false);
  }

  syncSettingsUI() {
    this.elGlobalStoreName.value = this.settings.storeName || 'SMILE';
    this.elGlobalCurrency.value = this.settings.currency || '₹';
    this.elGridPreset.value = this.settings.grid || '4x5';
    if (this.elGlobalFontScale) {
      this.elGlobalFontScale.value = this.settings.globalFontScale || 135;
    }
    if (this.elMrpBadgeToggle) {
      this.elMrpBadgeToggle.value = this.settings.mrpBadgeMode || 'auto';
    }
  }

  loadSampleData(notifyUser = true) {
    this.tags = JSON.parse(JSON.stringify(DEFAULT_SAMPLE_PRODUCTS));
    this.settings.grid = '4x5';
    this.settings.mrpBadgeMode = 'auto';
    this.saveToStorage();
    this.render();
    if (notifyUser) {
      this.notify('Loaded 20 supermarket items. Ready to print or auto-detect photos!', 'success');
    }
  }

  switchTab(tabId) {
    document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
    document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));

    const btn = document.querySelector(`.tab-btn[data-tab="${tabId}"]`);
    const panel = document.getElementById(tabId);
    if (btn) btn.classList.add('active');
    if (panel) panel.classList.add('active');

    if (tabId === 'tab-designer') {
      this.populateDesignerSelector();
      this.populateDesignerForm();
      this.renderDesignerTag();
    } else if (tabId === 'tab-sheet') {
      this.renderA4Sheets();
    } else if (tabId === 'tab-editor') {
      this.renderTableRows();
    }
  }

  openModal(modal) {
    if (modal) modal.style.display = 'flex';
  }

  closeModal(modal) {
    if (modal) modal.style.display = 'none';
  }

  notify(message, type = 'info') {
    const icons = { info: 'ℹ️', success: '✅', error: '⚠️' };
    this.elStatusIcon.textContent = icons[type] || 'ℹ️';
    this.elStatusMessage.textContent = message;
    this.elStatusAlert.style.display = 'flex';
    setTimeout(() => {
      this.elStatusAlert.style.display = 'none';
    }, 6000);
  }

  addNewTag() {
    const newId = `tag-${Date.now()}`;
    const newTag = {
      id: newId,
      title: 'NEW PRODUCT OFFER',
      mrp: 0,
      offerPrice: 79,
      unit: 'KG',
      storeName: this.settings.storeName,
      image: '',
      imageScale: 1,
      imageX: 0,
      imageY: 0,
      titleSize: 100,
      priceSize: 100,
      strokeWidth: 3
    };
    this.tags.push(newTag);
    this.saveToStorage();
    this.render();
    this.notify(`Added new tag. Total: ${this.tags.length} tags`, 'success');
  }

  // =========================================================================
  // ROBUST MULTI-ENGINE EXCEL PARSER (.XLSX / .XLS / CSV)
  // =========================================================================
  async handleFileUpload(e) {
    const file = e.target.files[0];
    if (!file) return;

    const ext = file.name.split('.').pop().toLowerCase();
    this.notify(`Reading ${file.name}...`, 'info');

    if (ext === 'xlsx' || ext === 'xls') {
      await this.parseExcelFile(file);
    } else if (ext === 'csv' || ext === 'tsv') {
      const text = await file.text();
      this.parseDelimitedText(text, ext === 'tsv' ? '\t' : ',');
    } else {
      this.notify('Unsupported file format. Please upload Excel (.xlsx) or CSV.', 'error');
    }
    e.target.value = '';
  }

  async parseExcelFile(file) {
    this.notify(`Reading Excel file: ${file.name}...`, 'info');

    // Engine 1: Pure in-browser ZIP Central Directory DecompressionStream parser
    try {
      const arrayBuffer = await file.arrayBuffer();
      const rows = await this.clientSideUnzipXLSX(new Uint8Array(arrayBuffer));
      if (rows && rows.length > 1) {
        this.importRowsData(rows);
        return;
      }
    } catch (err) {
      console.warn('Client-side Central Directory XLSX parse notice:', err);
    }

    // Engine 2: Direct server binary octet-stream endpoint
    try {
      const resp = await fetch('/api/parse-excel', {
        method: 'POST',
        headers: { 'Content-Type': 'application/octet-stream' },
        body: file
      });
      if (resp.ok) {
        const data = await resp.json();
        if (data.success && data.rows && data.rows.length > 1) {
          this.importRowsData(data.rows);
          return;
        }
      }
    } catch (e) {
      console.warn('Server parse attempt notice:', e);
    }

    // Engine 3: Check if file is actually an HTML table or TSV saved as .xls
    try {
      const text = await file.text();
      if (text.includes('<table') || text.includes('<tr')) {
        const parser = new DOMParser();
        const doc = parser.parseFromString(text, 'text/html');
        const trs = doc.querySelectorAll('tr');
        const rows = [];
        trs.forEach(tr => {
          const cells = Array.from(tr.querySelectorAll('td, th')).map(c => c.textContent.trim());
          if (cells.some(c => c.length > 0)) rows.push(cells);
        });
        if (rows.length > 1) {
          this.importRowsData(rows);
          return;
        }
      } else if (text.includes('\t') && text.split('\n').length > 1) {
        this.parseDelimitedText(text, '\t');
        return;
      }
    } catch (e) {
      console.warn('Text/HTML fallback notice:', e);
    }

    // Friendly fallback alert with instant paste guidance
    alert(`Could not automatically parse "${file.name}".\n\nQuick Solution:\n1. Open your file in Excel or Google Sheets.\n2. Select your items and copy them (Ctrl + C).\n3. Click the "Google Sheets / Paste" button in the ribbon.\n4. Go to "Quick Paste Table" and press Ctrl + V!`);
  }

  async clientSideUnzipXLSX(buffer) {
    const view = new DataView(buffer.buffer, buffer.byteOffset, buffer.byteLength);
    
    // Find End of Central Directory Record (EOCD)
    let eocdOffset = -1;
    for (let i = buffer.byteLength - 22; i >= Math.max(0, buffer.byteLength - 65557); i--) {
      if (view.getUint32(i, true) === 0x06054b50) {
        eocdOffset = i;
        break;
      }
    }
    if (eocdOffset === -1) {
      throw new Error('Not a standard ZIP/XLSX file');
    }

    const totalEntries = view.getUint16(eocdOffset + 10, true);
    const cdOffset = view.getUint32(eocdOffset + 16, true);
    const entries = {};
    let cur = cdOffset;

    for (let i = 0; i < totalEntries; i++) {
      const sig = view.getUint32(cur, true);
      if (sig !== 0x02014b50) break;

      const compMethod = view.getUint16(cur + 10, true);
      const compSize = view.getUint32(cur + 20, true);
      const uncompSize = view.getUint32(cur + 24, true);
      const nameLen = view.getUint16(cur + 28, true);
      const extraLen = view.getUint16(cur + 30, true);
      const commentLen = view.getUint16(cur + 32, true);
      const localHeaderOffset = view.getUint32(cur + 42, true);

      const nameBytes = buffer.slice(cur + 46, cur + 46 + nameLen);
      const fileName = new TextDecoder().decode(nameBytes);

      const localNameLen = view.getUint16(localHeaderOffset + 26, true);
      const localExtraLen = view.getUint16(localHeaderOffset + 28, true);
      const dataStart = localHeaderOffset + 30 + localNameLen + localExtraLen;
      const dataBytes = buffer.slice(dataStart, dataStart + compSize);

      entries[fileName] = { compMethod, compSize, uncompSize, dataBytes };
      cur += 46 + nameLen + extraLen + commentLen;
    }

    const decompressEntry = async (entry) => {
      if (!entry) return '';
      if (entry.compMethod === 0) {
        return new TextDecoder().decode(entry.dataBytes);
      } else if (entry.compMethod === 8 && typeof DecompressionStream !== 'undefined') {
        const ds = new DecompressionStream('deflate-raw');
        const writer = ds.writable.getWriter();
        writer.write(entry.dataBytes);
        writer.close();
        const decomp = await new Response(ds.readable).arrayBuffer();
        return new TextDecoder().decode(decomp);
      }
      throw new Error(`Unsupported compression method: ${entry.compMethod}`);
    };

    // Shared strings extraction
    const sharedStrings = [];
    if (entries['xl/sharedStrings.xml']) {
      const ssXml = await decompressEntry(entries['xl/sharedStrings.xml']);
      const siMatches = ssXml.match(/<si[\s\S]*?<\/si>/gi) || [];
      for (const si of siMatches) {
        const tMatches = si.match(/<t[^>]*>([\s\S]*?)<\/t>/gi) || [];
        const parts = tMatches.map(t => t.replace(/<[^>]+>/g, ''));
        sharedStrings.push(parts.join(''));
      }
    }

    // Find first worksheet
    const sheetKey = Object.keys(entries).find(k => k.startsWith('xl/worksheets/sheet') && k.endsWith('.xml')) || 'xl/worksheets/sheet1.xml';
    if (!entries[sheetKey]) {
      throw new Error('No worksheets found in workbook');
    }

    const sheetXml = await decompressEntry(entries[sheetKey]);
    const rows = [];
    const rowMatches = sheetXml.match(/<row[\s\S]*?<\/row>/gi) || [];

    for (const rowXml of rowMatches) {
      const cMatches = rowXml.match(/<c[\s\S]*?<\/c>/gi) || [];
      let maxCol = 0;
      const rowMap = {};

      for (const cXml of cMatches) {
        const rMatch = cXml.match(/\sr=\"([A-Za-z]+)([0-9]+)\"/);
        const colLetter = rMatch ? rMatch[1] : null;
        const colIdx = colLetter ? colLetterToIndex(colLetter) : Object.keys(rowMap).length;
        maxCol = Math.max(maxCol, colIdx);

        const isShared = /\st=\"s\"/.test(cXml);
        const isInline = /\st=\"inlineStr\"/.test(cXml);
        let val = '';

        if (isShared) {
          const vMatch = cXml.match(/<v>([0-9]+)<\/v>/);
          if (vMatch) {
            const sIdx = parseInt(vMatch[1], 10);
            val = sharedStrings[sIdx] || '';
          }
        } else if (isInline) {
          const tMatch = cXml.match(/<t[^>]*>([\s\S]*?)<\/t>/);
          val = tMatch ? tMatch[1] : '';
        } else {
          const vMatch = cXml.match(/<v>([\s\S]*?)<\/v>/);
          val = vMatch ? vMatch[1] : '';
        }
        rowMap[colIdx] = (val || '').trim();
      }

      const rowArr = [];
      for (let c = 0; c <= maxCol; c++) {
        rowArr.push(rowMap[c] || '');
      }
      if (rowArr.some(cell => cell.length > 0)) {
        rows.push(rowArr);
      }
    }

    return rows;
  }

  parseDelimitedText(text, delimiter = ',') {
    const lines = text.split(/\r?\n/).map(l => l.trim()).filter(l => l.length > 0);
    if (lines.length < 2) {
      this.notify('File must have a header row and at least one data row.', 'error');
      return;
    }

    if (lines[0].includes('\t')) delimiter = '\t';
    else if (lines[0].includes(';') && !lines[0].includes(',')) delimiter = ';';

    const rows = lines.map(line => {
      const pattern = new RegExp(
        `(?:${delimiter}|^)(?:"([^"]*(?:""[^"]*)*)"|([^"${delimiter}]*))`,
        'gi'
      );
      const row = [];
      let match;
      while ((match = pattern.exec(line))) {
        let val = match[1] ? match[1].replace(/""/g, '"') : match[2];
        row.push((val || '').trim());
      }
      return row;
    });

    this.importRowsData(rows);
  }

  importRowsData(rows) {
    if (!rows || rows.length < 2) return;

    const headers = rows[0].map(h => (h || '').toLowerCase().trim());

    const getColScore = (header, keywords) => {
      let score = 0;
      for (const kw of keywords) {
        if (header === kw) return 100;
        const escaped = kw.replace(/[-/\\^$*+?.()|[\]{}]/g, '\\$&');
        if (new RegExp(`(^|\\s|_|-)${escaped}($|\\s|_|-)`, 'i').test(header)) {
          score = Math.max(score, 80);
        } else if (header.includes(kw)) {
          score = Math.max(score, 40);
        }
      }
      return score;
    };

    const findCol = (keywords, excludeIndices = []) => {
      let bestIdx = -1;
      let highestScore = 0;
      for (let i = 0; i < headers.length; i++) {
        if (excludeIndices.includes(i)) continue;
        const s = getColScore(headers[i], keywords);
        if (s > highestScore) {
          highestScore = s;
          bestIdx = i;
        }
      }
      return highestScore > 0 ? bestIdx : -1;
    };

    const titleIdx = findCol(['title', 'product', 'item name', 'item', 'name', 'desc', 'particulars']);
    const mrpIdx = findCol([
      'mrp', 'm.r.p', 'm.r.p.', 'maximum retail price', 'max retail price', 'retail price', 
      'market price', 'regular price', 'original price', 'list price', 'old price', 'was price', 
      'strike price', 'base price', 'actual price'
    ]);
    let offerIdx = findCol([
      'offer price', 'special price', 'sale price', 'selling price', 'deal price', 'net price', 
      'discounted price', 'discount price', 'our price', 'final price', 'offer', 'sale', 'selling', 
      'deal', 'discount', 'rate', 'price', 'amount', 'cost', 'net'
    ], mrpIdx !== -1 ? [mrpIdx] : []);

    // If only one single price column is found, treat it as the offer/selling price
    let finalMrpIdx = mrpIdx;
    let finalOfferIdx = offerIdx;
    if (finalMrpIdx !== -1 && finalOfferIdx === -1) {
      finalOfferIdx = finalMrpIdx;
      finalMrpIdx = -1;
    }

    const unitIdx = findCol(['unit', 'qty', 'measure', 'per', 'uom', 'pack'], [titleIdx, finalMrpIdx, finalOfferIdx]);
    const storeIdx = findCol(['store', 'brand', 'shop'], [titleIdx, finalMrpIdx, finalOfferIdx]);
    const imageIdx = findCol(['image', 'img', 'photo', 'pic', 'file', 'url'], [titleIdx, finalMrpIdx, finalOfferIdx]);

    const newTags = [];
    for (let r = 1; r < rows.length; r++) {
      const row = rows[r];
      if (!row || row.length === 0 || !row.some(cell => cell.length > 0)) continue;

      const title = (titleIdx !== -1 && row[titleIdx]) ? row[titleIdx] : (row[0] || `ITEM ${r}`);
      const mrpRaw = (finalMrpIdx !== -1 && row[finalMrpIdx]) ? row[finalMrpIdx].replace(/[^0-9.]/g, '') : '';
      const offerRaw = (finalOfferIdx !== -1 && row[finalOfferIdx]) ? row[finalOfferIdx].replace(/[^0-9.]/g, '') : '0';
      const unit = (unitIdx !== -1 && row[unitIdx]) ? row[unitIdx].toUpperCase() : 'KG';
      const store = (storeIdx !== -1 && row[storeIdx]) ? row[storeIdx] : this.settings.storeName;
      let img = (imageIdx !== -1 && row[imageIdx]) ? row[imageIdx] : '';

      img = this.convertDriveLink(img);

      newTags.push({
        id: `tag-${r}-${Date.now()}`,
        title: title.toUpperCase(),
        mrp: parseFloat(mrpRaw) || 0,
        offerPrice: parseFloat(offerRaw) || 0,
        unit: unit.replace(/^\//, '').trim(),
        storeName: store,
        image: img,
        imageScale: 1,
        imageX: 0,
        imageY: 0,
        titleSize: 100,
        priceSize: 100,
        strokeWidth: 3
      });
    }

    if (newTags.length > 0) {
      this.tags = newTags;
      if (newTags.some(t => t.mrp > 0)) {
        this.settings.mrpBadgeMode = 'auto';
        if (this.elMrpBadgeToggle) this.elMrpBadgeToggle.value = 'auto';
      }
      if (newTags.length === 20 && this.settings.grid === '4x4') {
        this.settings.grid = '4x5';
        if (this.elGridPreset) this.elGridPreset.value = '4x5';
      }
      this.saveToStorage();
      this.render();
      this.notify(`Successfully imported ${newTags.length} products!`, 'success');
      this.switchTab('tab-sheet');
    } else {
      this.notify('Could not find valid product rows in uploaded file.', 'error');
    }
  }

  // =========================================================================
  // GOOGLE SHEETS & QUICK PASTE
  // =========================================================================
  async fetchGoogleSheetData() {
    const url = this.elGSheetUrlInput.value.trim();
    if (!url) {
      alert('Please paste a Google Sheets URL');
      return;
    }

    const match = url.match(/\/d\/([a-zA-Z0-9-_]+)/);
    if (!match) {
      alert('Invalid Google Sheets URL format. It should contain /d/<sheet-id>/');
      return;
    }

    const sheetId = match[1];
    const directCsvUrl = `https://docs.google.com/spreadsheets/d/${sheetId}/export?format=csv`;

    this.notify('Fetching Google Sheet data...', 'info');

    try {
      const resp = await fetch(directCsvUrl);
      if (resp.ok) {
        const csvText = await resp.text();
        this.closeModal(this.elModalGoogleSheets);
        this.parseDelimitedText(csvText, ',');
        return;
      }
    } catch (e) {
      // CORS fallback
    }

    try {
      const proxyResp = await fetch('/api/fetch-gsheet', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ url })
      });
      const proxyData = await proxyResp.json();
      if (proxyData.success && proxyData.csv) {
        this.closeModal(this.elModalGoogleSheets);
        this.parseDelimitedText(proxyData.csv, ',');
        return;
      }
      throw new Error(proxyData.error || 'Failed to fetch via proxy');
    } catch (err) {
      alert(`Could not fetch Google Sheet directly due to permission/CORS. \n\nTip: In Google Sheets, make sure 'Anyone with the link can view' is enabled, OR copy the cells and use the 'Quick Paste Table' tab!`);
    }
  }

  applyPastedSheetData() {
    const rawText = this.elGSheetPasteTextarea.value.trim();
    if (!rawText) {
      alert('Please paste spreadsheet data first.');
      return;
    }
    this.closeModal(this.elModalGoogleSheets);
    this.parseDelimitedText(rawText, '\t');
  }

  // =========================================================================
  // FOLDER IMAGE AUTO-DETECTION & FUZZY MATCHING
  // =========================================================================
  // =========================================================================
  // FOLDER IMAGE AUTO-DETECTION & FUZZY MATCHING
  // =========================================================================
  async handleFolderUpload(e) {
    const files = Array.from(e.target.files);
    if (files.length === 0) return;

    const imageFiles = files.filter(f => {
      const ext = f.name.split('.').pop().toLowerCase();
      return ['jpg', 'jpeg', 'png', 'webp', 'gif', 'svg', 'avif'].includes(ext);
    });

    if (imageFiles.length === 0) {
      this.notify('No image files found in the selected folder.', 'error');
      return;
    }

    this.notify(`Scanning ${imageFiles.length} images from folder...`, 'info');

    const fileEntries = imageFiles.map(file => {
      const cleanName = file.name
        .replace(/\.[^/.]+$/, '')
        .toLowerCase()
        .replace(/[-_]/g, ' ')
        .replace(/[^a-z0-9\s]/g, '')
        .trim();

      return {
        file,
        name: file.name,
        cleanName,
        url: URL.createObjectURL(file)
      };
    });

    this.matchFilesToTags(fileEntries, 'Local Folder', '📁');
    e.target.value = '';
  }

  matchFilesToTags(fileEntries, sourceLabel = 'Local Folder', sourceIcon = '📁') {
    this.pendingFolderMatches = [];
    let matchCount = 0;

    const STOP_WORDS = new Set([
      'offer', 'offers', 'fresh', 'kg', 'kgs', 'gm', 'gms', 'g', 'pc', 'pcs',
      'piece', 'pieces', 'super', 'best', 'grade', 'a', 'the', 'of', 'per',
      'new', 'special', 'pack', 'pkt', 'pkts', 'packet', 'packets', 'ltr',
      'ltrs', 'l', 'ml', 'box', 'pouch', 'rs', 'no', 'nos', 'item', 'items',
      'product', 'products', 'smile', 'local'
    ]);

    const normalize = (str) => {
      return (str || '')
        .toLowerCase()
        .replace(/\.[^/.]+$/, '')
        .replace(/[-_.,/\\()]/g, ' ')
        .replace(/[^a-z0-9\s]/g, '')
        .trim();
    };

    const getKeywords = (str) => {
      const norm = normalize(str);
      return norm.split(/\s+/).filter(w => w.length >= 2 && !STOP_WORDS.has(w));
    };

    // Pre-calculate keywords for all images in the folder
    const processedFiles = fileEntries.map(entry => {
      const clean = normalize(entry.name);
      const words = getKeywords(entry.name);
      return {
        ...entry,
        clean,
        words
      };
    });

    const usedFileNames = new Set();

    this.tags.forEach((tag, idx) => {
      const cleanTagTitle = normalize(tag.title);
      const tagWords = getKeywords(tag.title);

      let bestMatch = null;
      let highestScore = 0;

      processedFiles.forEach(entry => {
        let score = 0;

        if (entry.clean === cleanTagTitle && cleanTagTitle.length > 0) {
          // 1. Exact string match
          score = 100;
        } else if (tagWords.length > 0 && entry.words.length > 0) {
          const matchingWords = tagWords.filter(w => entry.words.includes(w));

          if (matchingWords.length === tagWords.length && matchingWords.length === entry.words.length) {
            // 2. Exact word set match (order independent)
            score = 95;
          } else if (entry.words.length > 0 && entry.words.every(w => tagWords.includes(w))) {
            // 3. All file words are in the tag title (e.g. file: "onion", tag: "red onion")
            score = 85 + Math.min(10, matchingWords.length * 2);
          } else if (tagWords.length > 0 && tagWords.every(w => entry.words.includes(w))) {
            // 4. All tag words are in the file name (e.g. tag: "potato", file: "fresh_potato_agra")
            score = 80 + Math.min(10, matchingWords.length * 2);
          } else if (matchingWords.length > 0) {
            // 5. Partial word overlap
            const ratio = matchingWords.length / Math.max(tagWords.length, entry.words.length);
            score = Math.round(40 + ratio * 40);
          }
        }

        // Slight bonus if not already claimed by a previous tag
        if (score > 0 && !usedFileNames.has(entry.name)) {
          score += 2;
        }

        if (score > highestScore && score >= 45) {
          highestScore = Math.min(score, 100);
          bestMatch = entry;
        }
      });

      if (bestMatch) {
        matchCount++;
        usedFileNames.add(bestMatch.name);
      }

      this.pendingFolderMatches.push({
        tagIndex: idx,
        tagTitle: tag.title,
        matchedFile: bestMatch,
        score: highestScore
      });
    });

    this.elFolderMatchSummary.innerHTML = `Found <strong>${fileEntries.length}</strong> images in ${sourceLabel}. Matched images for <strong>${matchCount}</strong> of your <strong>${this.tags.length}</strong> uploaded products. (Unmatched files in the folder are ignored).`;

    this.elFolderMatchList.innerHTML = this.pendingFolderMatches.map(m => `
      <div class="match-row">
        ${m.matchedFile ? `<img src="${m.matchedFile.url}" class="match-thumb" alt="Preview" onerror="this.onerror=null;this.src='/api/proxy-image?url='+encodeURIComponent('${m.matchedFile.url}');">` : `<div class="match-thumb" style="display:flex;align-items:center;justify-content:center;background:#1e293b;color:#64748b;font-size:11px;">No Img</div>`}
        <div class="match-info">
          <div class="match-title">${m.tagTitle}</div>
          <div class="match-filename">${m.matchedFile ? `Matched: ${sourceIcon} ${m.matchedFile.name}` : '❌ No matching image in folder (kept as is)'}</div>
        </div>
        <div>
          <span class="match-badge ${m.matchedFile ? 'success' : 'unmatched'}">
            ${m.matchedFile ? `Match (${m.score}%)` : 'No Image'}
          </span>
        </div>
      </div>
    `).join('');

    this.openModal(this.elModalFolderMatch);
  }

  applyFolderMatches() {
    let appliedCount = 0;
    this.pendingFolderMatches.forEach(m => {
      if (m.matchedFile && this.tags[m.tagIndex]) {
        this.tags[m.tagIndex].image = m.matchedFile.url;
        appliedCount++;
      }
    });

    this.closeModal(this.elModalFolderMatch);
    this.saveToStorage();
    this.render();
    this.notify(`Applied images to ${appliedCount} products from your uploaded list!`, 'success');
  }

  // =========================================================================
  // GOOGLE DRIVE AUTO-DETECTION & LINK RESOLVER
  // =========================================================================
  async scanGoogleDriveFolder() {
    const folderInput = this.elDriveFolderUrlInput ? this.elDriveFolderUrlInput.value.trim() : '';
    const apiKey = this.elDriveApiKeyInput ? this.elDriveApiKeyInput.value.trim() : '';

    if (!folderInput) {
      this.notify('Please enter a Google Drive folder link or folder ID.', 'error');
      if (this.elDriveFolderUrlInput) this.elDriveFolderUrlInput.focus();
      return;
    }

    let folderId = null;
    const m1 = folderInput.match(/folders\/([a-zA-Z0-9_-]+)/);
    if (m1) {
      folderId = m1[1];
    } else {
      const m2 = folderInput.match(/[?&]id=([a-zA-Z0-9_-]+)/);
      if (m2) {
        folderId = m2[1];
      } else if (/^[a-zA-Z0-9_-]{15,60}$/.test(folderInput)) {
        folderId = folderInput;
      }
    }

    if (!folderId) {
      this.notify('Could not detect a valid Google Drive Folder ID from link.', 'error');
      return;
    }

    const btn = this.elBtnScanDriveFolder;
    const txt = this.elBtnScanDriveText;
    const origTxt = txt ? txt.textContent : 'Scan & Auto-Detect Images';

    if (btn) btn.disabled = true;
    if (txt) txt.textContent = 'Scanning Google Drive...';
    this.notify('Connecting to Google Drive folder and scanning images...', 'info');

    try {
      let fileEntries = [];

      // 1. Call server API endpoint
      try {
        const resp = await fetch('/api/fetch-gdrive-folder', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ url: folderInput, folderId, apiKey })
        });

        if (resp.ok) {
          const resData = await resp.json();
          if (resData.success && Array.isArray(resData.files) && resData.files.length > 0) {
            fileEntries = resData.files.map(f => {
              const cleanName = f.name
                .replace(/\.[^/.]+$/, '')
                .toLowerCase()
                .replace(/[-_]/g, ' ')
                .replace(/[^a-z0-9\s]/g, '')
                .trim();
              return {
                id: f.id,
                name: f.name,
                cleanName,
                url: f.url || `https://lh3.googleusercontent.com/d/${f.id}`
              };
            });
          } else if (resData.error) {
            throw new Error(resData.error);
          }
        } else {
          const errData = await resp.json().catch(() => ({}));
          throw new Error(errData.error || `Server returned status ${resp.status}`);
        }
      } catch (srvErr) {
        console.warn('Server fetch-gdrive-folder error:', srvErr);

        // 2. Client-side fallback if Google Drive API Key provided
        if (apiKey) {
          const q = encodeURIComponent(`'${folderId}' in parents and trashed=false`);
          const apiUrl = `https://www.googleapis.com/drive/v3/files?q=${q}&fields=files(id,name,mimeType)&pageSize=1000&key=${apiKey}`;
          const gResp = await fetch(apiUrl);
          if (gResp.ok) {
            const gData = await gResp.json();
            const imgs = (gData.files || []).filter(f => (f.mimeType && f.mimeType.startsWith('image/')) || /\.(jpe?g|png|webp|gif|avif|bmp)$/i.test(f.name));
            if (imgs.length > 0) {
              fileEntries = imgs.map(f => ({
                id: f.id,
                name: f.name,
                cleanName: f.name.replace(/\.[^/.]+$/, '').toLowerCase().replace(/[-_]/g, ' ').replace(/[^a-z0-9\s]/g, '').trim(),
                url: `https://lh3.googleusercontent.com/d/${f.id}`
              }));
            }
          }
        }

        if (fileEntries.length === 0) {
          throw srvErr;
        }
      }

      if (fileEntries.length === 0) {
        throw new Error("No images found in Google Drive folder. Please ensure the folder is shared as 'Anyone with the link can view'.");
      }

      this.closeModal(this.elModalDriveFolder);
      this.matchFilesToTags(fileEntries, 'Google Drive', '☁️');
      this.notify(`Found ${fileEntries.length} images from Google Drive! Review the matches below.`, 'success');

    } catch (err) {
      console.error(err);
      alert(`Google Drive Detection Notice:\n\n${err.message || err}\n\nTip: In Google Drive, right-click the folder -> Share -> change "General access" to "Anyone with the link can view".`);
    } finally {
      if (btn) btn.disabled = false;
      if (txt) txt.textContent = origTxt;
    }
  }

  convertDriveLink(url) {
    if (!url || typeof url !== 'string') return url;
    const match = url.match(/(?:drive\.google\.com\/(?:file\/d\/|(?:open|uc)\?id=)|(?:[?&]id=))([a-zA-Z0-9_-]{20,})/);
    if (match) {
      return `https://lh3.googleusercontent.com/d/${match[1]}`;
    }
    return url;
  }

  resolveGoogleDriveImages() {
    let count = 0;
    this.tags.forEach(tag => {
      const converted = this.convertDriveLink(tag.image);
      if (converted !== tag.image) {
        tag.image = converted;
        count++;
      }
    });

    if (count > 0) {
      this.saveToStorage();
      this.render();
      this.notify(`Converted ${count} Google Drive image links to high-speed direct URLs!`, 'success');
    } else {
      const customUrl = prompt('Enter a Google Drive image link to set as default:');
      if (customUrl) {
        const resolved = this.convertDriveLink(customUrl);
        this.tags.forEach(t => t.image = resolved);
        this.saveToStorage();
        this.render();
        this.notify('Applied Google Drive image to all tags!', 'success');
      }
    }
  }

  // =========================================================================
  // GENERAL RENDER CYCLE
  // =========================================================================
  render() {
    this.updateCounters();
    this.renderA4Sheets();
    this.renderTableRows();
    this.populateDesignerSelector();
    this.renderDesignerTag();
  }

  updateCounters() {
    const count = this.tags.length;
    this.elTagCountBadge.textContent = count;
    this.elTableCountBadge.textContent = count;
    this.elVisibleRowCount.textContent = count;

    let totalSavings = 0;
    this.tags.forEach(t => {
      const diff = Math.max(0, (t.mrp || 0) - (t.offerPrice || 0));
      totalSavings += diff;
    });
    this.elTotalSavingsSum.textContent = `${this.settings.currency}${totalSavings.toFixed(2)}`;
  }

  // =========================================================================
  // TAB 1: A4 PORTRAIT SHEET RENDERING (16 TAGS, 3x4, etc.)
  // =========================================================================
  getGridConfig() {
    switch (this.settings.grid) {
      case 'landscape-4x3': return { cols: 4, rows: 3, perPage: 12, cssClass: 'grid-landscape-4x3', orientation: 'landscape' };
      case 'landscape-4x4': return { cols: 4, rows: 4, perPage: 16, cssClass: 'grid-landscape-4x4', orientation: 'landscape' };
      case '4x5': return { cols: 4, rows: 5, perPage: 20, cssClass: 'grid-4x5', orientation: 'portrait' };
      case '5x4': return { cols: 5, rows: 4, perPage: 20, cssClass: 'grid-5x4', orientation: 'portrait' };
      case '4x4': return { cols: 4, rows: 4, perPage: 16, cssClass: 'grid-4x4', orientation: 'portrait' };
      case '4x3': return { cols: 4, rows: 3, perPage: 12, cssClass: 'grid-4x3', orientation: 'portrait' };
      case '3x4': return { cols: 3, rows: 4, perPage: 12, cssClass: 'grid-3x4', orientation: 'portrait' };
      case '3x3': return { cols: 3, rows: 3, perPage: 9, cssClass: 'grid-3x3', orientation: 'portrait' };
      case '2x4': return { cols: 2, rows: 4, perPage: 8, cssClass: 'grid-2x4', orientation: 'portrait' };
      case '2x2': return { cols: 2, rows: 2, perPage: 4, cssClass: 'grid-2x2', orientation: 'portrait' };
      case '1x4': return { cols: 1, rows: 4, perPage: 4, cssClass: 'grid-1x4', orientation: 'portrait' };
      case '1x3': return { cols: 1, rows: 3, perPage: 3, cssClass: 'grid-1x3', orientation: 'portrait' };
      case '1x2': return { cols: 1, rows: 2, perPage: 2, cssClass: 'grid-1x2', orientation: 'portrait' };
      case '1x1': return { cols: 1, rows: 1, perPage: 1, cssClass: 'grid-1x1', orientation: 'portrait' };
      default: return { cols: 4, rows: 5, perPage: 20, cssClass: 'grid-4x5', orientation: 'portrait' };
    }
  }

  renderA4Sheets() {
    const config = this.getGridConfig();
    const isLandscape = config.orientation === 'landscape';
    const perPage = config.perPage;
    const totalTags = this.tags.length;
    const numPages = Math.max(1, Math.ceil(totalTags / perPage));

    this.elCurrentPageNum.textContent = '1';
    this.elTotalPagesNum.textContent = numPages;
    this.totalPages = numPages;

    // Update dynamic print orientation stylesheet so browser print dialog matches sheet layout
    let printStyleEl = document.getElementById('printOrientationStyle');
    if (!printStyleEl) {
      printStyleEl = document.createElement('style');
      printStyleEl.id = 'printOrientationStyle';
      document.head.appendChild(printStyleEl);
    }
    printStyleEl.textContent = `@media print { @page { size: A4 ${isLandscape ? 'landscape' : 'portrait'}; margin: 4mm; } }`;

    if (this.elBtnDownloadA4Text) {
      this.elBtnDownloadA4Text.textContent = numPages > 1 
        ? `Download All A4 ${isLandscape ? 'Landscape ' : ''}Sheets (${numPages} Sheets)` 
        : `Download A4 ${isLandscape ? 'Landscape ' : ''}Sheet`;
    }

    if (this.elSheetDownloadAllWrapper) {
      this.elSheetDownloadAllWrapper.style.display = numPages > 1 ? 'inline-flex' : 'none';
      if (this.elBtnSheetDownloadAllText) {
        this.elBtnSheetDownloadAllText.textContent = `Download All Sheets (${numPages})`;
      }
    }

    if (this.elBtnSheetDownloadA4Text) {
      this.elBtnSheetDownloadA4Text.textContent = numPages > 1 
        ? 'Download Sheet 1 (PNG)' 
        : `Download A4 ${isLandscape ? 'Landscape ' : ''}Sheet (PNG)`;
    }

    this.elA4Stage.innerHTML = '';
    this.elA4Stage.className = `a4-stage ${this.settings.showCutGuides ? 'show-cut-guides' : ''}`;

    for (let pageIdx = 0; pageIdx < numPages; pageIdx++) {
      const pageDiv = document.createElement('div');
      pageDiv.className = `a4-page ${isLandscape ? 'landscape' : ''}`;

      const gridDiv = document.createElement('div');
      gridDiv.className = `a4-grid ${config.cssClass}`;

      const startIndex = pageIdx * perPage;
      const endIndex = Math.min(startIndex + perPage, totalTags);

      for (let i = startIndex; i < endIndex; i++) {
        const tag = this.tags[i];
        const tagCard = this.createTagDOMElement(tag, i);
        gridDiv.appendChild(tagCard);
      }

      const currentTagsOnPage = endIndex - startIndex;
      if (currentTagsOnPage < perPage) {
        for (let emptyIdx = currentTagsOnPage; emptyIdx < perPage; emptyIdx++) {
          const emptyCard = document.createElement('div');
          emptyCard.className = 'price-tag-card empty-slot';
          emptyCard.innerHTML = `
            <div style="display:flex;align-items:center;justify-content:center;height:100%;color:#94a3b8;font-size:11px;font-weight:600;border:1.5px dashed #cbd5e1;margin:4px;border-radius:4px;background:#f8fafc;">
              + Empty Slot ${emptyIdx + 1}
            </div>`;
          gridDiv.appendChild(emptyCard);
        }
      }

      pageDiv.appendChild(gridDiv);
      this.elA4Stage.appendChild(pageDiv);
    }

    this.applyA4Zoom();
    this.autoFitAllLongTitles(this.elA4Stage);
  }

  autoFitAllLongTitles(container = this.elA4Stage) {
    if (!container) return;
    requestAnimationFrame(() => {
      const titles = container.querySelectorAll('.tag-title-text');
      titles.forEach(titleEl => {
        const row = titleEl.closest('.tag-title-row');
        if (!row) return;

        let maxIterations = 8;
        // If text exceeds row height (i.e. more than 2 lines or overflows row), shrink progressively
        while (maxIterations > 0 && (titleEl.scrollHeight > row.clientHeight + 1 || titleEl.offsetHeight > row.clientHeight + 1)) {
          const card = titleEl.closest('.price-tag-card');
          const currentShrink = parseFloat(titleEl.style.getPropertyValue('--title-shrink') || (card ? card.style.getPropertyValue('--title-shrink') : '1') || '1');
          if (currentShrink <= 0.30) break;
          const nextShrink = (currentShrink * 0.90).toFixed(3);
          titleEl.style.setProperty('--title-shrink', nextShrink);
          if (card) card.style.setProperty('--title-shrink', nextShrink);
          maxIterations--;
        }
      });
    });
  }

  createTagDOMElement(tag, index) {
    const card = document.createElement('div');
    const badgeMode = this.settings.mrpBadgeMode || 'auto';
    const numMrp = Number(tag.mrp || 0);
    const numOffer = Number(tag.offerPrice || 0);
    const hasMrp = (badgeMode !== 'hide') && numMrp > 0 && numMrp > numOffer;
    const mrpVal = hasMrp ? numMrp.toFixed(1) : '';
    const offerVal = (tag.offerPrice !== undefined && tag.offerPrice !== null) ? numOffer.toFixed(2) : '0.00';
    const savings = hasMrp ? Math.max(0, numMrp - numOffer).toFixed(2) : '0.00';
    const store = tag.storeName || this.settings.storeName;
    const unitText = tag.unit ? ` / ${tag.unit}` : '';

    const hasImage = Boolean(tag.image && tag.image.trim().length > 0);
    card.className = `price-tag-card ${hasImage ? 'has-image' : 'no-image'} ${hasMrp ? 'has-badges' : 'no-badges'}`;
    card.dataset.index = index;

    const imgScale = tag.imageScale || 1.0;
    const imgX = tag.imageX || 0;
    const imgY = tag.imageY || 0;

    // Apply global font scaling (default 135% for Extra Heavy Black)
    const globalBoost = (this.settings.globalFontScale || 135) / 100;
    const titleMult = ((tag.titleSize || 100) / 100) * globalBoost;
    const priceMult = ((tag.priceSize || 100) / 100) * globalBoost;

    // Auto-shrink factor for longer titles to guarantee it fits into at most 2 lines without any part hidden
    const titleStr = (tag.title || 'SPECIAL OFFER').trim().toUpperCase();
    const words = titleStr.split(/\s+/).filter(w => w.length > 0);
    const longestWordLen = words.reduce((max, w) => Math.max(max, w.length), 0);

    let shrinkFactor = 1.0;
    const len = titleStr.length;

    // Granular progressive shrink factor
    if (len > 38) {
      shrinkFactor = 0.42;
    } else if (len > 32) {
      shrinkFactor = 0.48;
    } else if (len > 26) {
      shrinkFactor = 0.55;
    } else if (len > 21) {
      shrinkFactor = 0.63;
    } else if (len > 17) {
      shrinkFactor = 0.72;
    } else if (len > 13) {
      shrinkFactor = 0.82;
    } else if (len > 9) {
      shrinkFactor = 0.92;
    }

    // Check longest single word so individual long words fit horizontally
    if (longestWordLen > 14) {
      shrinkFactor = Math.min(shrinkFactor, 0.45);
    } else if (longestWordLen > 11) {
      shrinkFactor = Math.min(shrinkFactor, 0.58);
    } else if (longestWordLen > 9) {
      shrinkFactor = Math.min(shrinkFactor, 0.72);
    }

    // For 20-tag grids (4x5 and 5x4) where tags are narrower on A4, give extra room
    if ((this.settings.grid === '4x5' || this.settings.grid === '5x4') && len > 12) {
      shrinkFactor = Math.round(shrinkFactor * 0.88 * 100) / 100;
    }

    shrinkFactor = Math.max(0.30, shrinkFactor);

    // Auto-shrink factor for price & unit to guarantee it fits inside the column
    const fullPriceStr = `${offerVal}${unitText}`;
    let priceShrink = 1.0;
    if (fullPriceStr.length >= 14) {
      priceShrink = 0.68;
    } else if (fullPriceStr.length >= 12) {
      priceShrink = 0.76;
    } else if (fullPriceStr.length >= 10) {
      priceShrink = 0.85;
    } else if (fullPriceStr.length >= 8) {
      priceShrink = 0.92;
    }

    card.style.setProperty('--font-title-mult', titleMult);
    card.style.setProperty('--font-price-mult', priceMult);
    card.style.setProperty('--title-shrink', shrinkFactor);
    card.style.setProperty('--price-shrink', priceShrink);

    card.innerHTML = `
      ${hasImage ? `
      <div class="tag-bg-layer">
        <img src="${tag.image}" 
             class="tag-bg-img" 
             style="transform: translate(${imgX}px, ${imgY}px) scale(${imgScale});" 
             alt="${tag.title}">
      </div>` : ''}

      <div class="tag-content-layer ${hasMrp ? 'has-badges' : 'no-badges'}">
        <!-- Store Brand Logo (Permanently on top of each price tag) -->
        <div class="tag-logo-row">
          <img src="BIGSAVE.PNG" class="tag-top-logo" alt="BIG SAVE Smile">
        </div>

        <!-- Top Badges / MRP & Saving Row -->
        ${hasMrp ? `
        <div class="tag-badge-row">
          <div class="mrp-text-nobadge">
            <span>MRP </span>
            <span class="mrp-strike">${mrpVal}</span>
          </div>
          <div class="save-badge">
            <span>${store} SAVE ${this.settings.currency}${savings}</span>
          </div>
        </div>
        ` : ''}

        <!-- Product Title (Thick Big Black Font) -->
        <div class="tag-title-row">
          <div class="tag-title-text">
            ${tag.title || 'SPECIAL OFFER'}
          </div>
        </div>

        <!-- Price & Unit (Thick Big Black Font) -->
        <div class="tag-price-row">
          <div class="tag-price-text">
            ${offerVal}<span class="tag-unit-text">${unitText}</span>
          </div>
        </div>
      </div>
    `;

    card.addEventListener('click', () => {
      this.selectedTagIndex = index;
      this.openQuickEditModal(index);
    });

    return card;
  }

  applyA4Zoom() {
    const config = this.getGridConfig();
    const isLandscape = config.orientation === 'landscape';
    const scale = this.settings.zoom / 100;
    const a4MmWidth = isLandscape ? 297 : 210;
    const a4MmHeight = isLandscape ? 210 : 297;
    document.querySelectorAll('.a4-page').forEach(page => {
      page.style.transform = `scale(${scale})`;
      page.style.marginBottom = `${-(a4MmHeight * (1 - scale))}mm`;
      if (scale < 1) {
        page.style.marginLeft = `${-(a4MmWidth * (1 - scale) / 2)}mm`;
        page.style.marginRight = `${-(a4MmWidth * (1 - scale) / 2)}mm`;
      } else {
        page.style.marginLeft = 'auto';
        page.style.marginRight = 'auto';
      }
    });
  }

  fitA4Zoom() {
    if (!this.elA4Stage) return;
    const config = this.getGridConfig();
    const isLandscape = config.orientation === 'landscape';
    const stageRect = this.elA4Stage.getBoundingClientRect();
    const stageWidth = Math.max(180, (stageRect.width || this.elA4Stage.clientWidth || window.innerWidth) - 40);
    const stageHeight = Math.max(200, (stageRect.height || this.elA4Stage.clientHeight || (window.innerHeight - 200)) - 60);
    const a4PxWidth = isLandscape ? 1123 : 794;   // 297mm vs 210mm at 96 DPI
    const a4PxHeight = isLandscape ? 794 : 1123;  // 210mm vs 297mm at 96 DPI

    const zoomW = (stageWidth / a4PxWidth) * 100;
    const zoomH = (stageHeight / a4PxHeight) * 100;

    // Scale so the full sheet (all rows and columns) fits entirely on the user's screen
    let targetZoom = Math.floor(Math.min(zoomW, zoomH));
    targetZoom = Math.max(20, Math.min(100, targetZoom));

    this.settings.zoom = targetZoom;
    if (this.elZoomSlider) this.elZoomSlider.value = targetZoom;
    if (this.elZoomVal) this.elZoomVal.textContent = `${targetZoom}%`;
    this.applyA4Zoom();
  }

  openQuickEditModal(index) {
    const tag = this.tags[index];
    if (!tag) return;

    const badgeMode = this.settings.mrpBadgeMode || 'auto';
    const numMrp = Number(tag.mrp || 0);
    const numOffer = Number(tag.offerPrice || 0);
    const hasMrp = (badgeMode !== 'hide') && numMrp > 0 && numMrp > numOffer;
    this.elQeTagId.textContent = index + 1;
    this.elQeTitle.value = tag.title || '';
    this.elQeMrp.value = numMrp > 0 ? tag.mrp : '';
    this.elQeOffer.value = tag.offerPrice || 0;
    this.elQeUnit.value = tag.unit || 'KG';

    if (hasMrp) {
      const diff = Math.max(0, numMrp - numOffer);
      this.elQeSavings.value = `${this.settings.currency}${diff.toFixed(2)}`;
    } else {
      this.elQeSavings.value = 'No MRP (No badge)';
    }
    this.elQeImageUpload.value = '';

    this.openModal(this.elModalQuickEdit);
  }

  // =========================================================================
  // TAB 2: DATA TABLE / BATCH EDITOR
  // =========================================================================
  renderTableRows() {
    const search = this.elTableSearchInput.value.toLowerCase().trim();
    const rows = [];

    const badgeMode = this.settings.mrpBadgeMode || 'auto';

    this.tags.forEach((tag, idx) => {
      if (search && !tag.title.toLowerCase().includes(search)) return;

      const numMrp = Number(tag.mrp || 0);
      const numOffer = Number(tag.offerPrice || 0);
      const hasMrp = (badgeMode !== 'hide') && numMrp > 0 && numMrp > numOffer;
      const savingsBadgeHtml = hasMrp
        ? `<span class="table-savings-badge">${tag.storeName || 'SMILE'} SAVE ${this.settings.currency}${Math.max(0, numMrp - numOffer).toFixed(2)}</span>`
        : `<span class="table-savings-badge" style="opacity: 0.5; background: #e2e8f0; color: #475569;">No MRP (No Badge)</span>`;

      rows.push(`
        <tr data-index="${idx}">
          <td>${idx + 1}</td>
          <td>
            <label class="thumb-cell-wrapper" title="${tag.image ? 'Click to replace photo' : 'Click to upload photo'}">
              ${tag.image 
                ? `<img src="${tag.image}" class="table-thumb" alt="Product">`
                : `<div class="table-thumb-empty" title="Upload Photo"><svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="18" height="18" rx="2"/><circle cx="8.5" cy="8.5" r="1.5"/><path d="M21 15l-5-5L5 21"/></svg><span>+ Photo</span></div>`
              }
              <input type="file" accept="image/*" class="table-row-img-input" data-index="${idx}">
            </label>
          </td>
          <td>
            <input type="text" class="table-input tbl-title" data-index="${idx}" value="${tag.title}">
          </td>
          <td>
            <input type="number" step="0.5" class="table-input tbl-mrp" data-index="${idx}" value="${tag.mrp > 0 ? tag.mrp : ''}" placeholder="None">
          </td>
          <td>
            <input type="number" step="0.5" class="table-input tbl-offer" data-index="${idx}" value="${tag.offerPrice}">
          </td>
          <td>
            <input type="text" class="table-input tbl-unit" data-index="${idx}" value="${tag.unit}">
          </td>
          <td>
            ${savingsBadgeHtml}
          </td>
          <td>
            <div class="table-row-actions">
              <button class="btn-icon tbl-btn-png" data-index="${idx}" title="Download this tag as PNG" style="color: #22c55e;">
                <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4M7 10l5 5 5-5M12 15V3"/></svg>
              </button>
              <button class="btn-icon tbl-btn-edit" data-index="${idx}" title="Open in Designer Studio">
                <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 20h9M16.5 3.5a2.121 2.121 0 013 3L7 19l-4 1 1-4L16.5 3.5z"/></svg>
              </button>
              <button class="btn-icon tbl-btn-dup" data-index="${idx}" title="Duplicate tag">
                <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="9" y="9" width="13" height="13" rx="2" ry="2"/><path d="M5 15H4a2 2 0 01-2-2V4a2 2 0 012-2h9a2 2 0 012 2v1"/></svg>
              </button>
              <button class="btn-icon danger tbl-btn-del" data-index="${idx}" title="Delete tag">
                <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 01-2 2H7a2 2 0 01-2-2V6m3 0V4a2 2 0 012-2h4a2 2 0 012 2v2"/></svg>
              </button>
            </div>
          </td>
        </tr>
      `);
    });

    this.elTableBody.innerHTML = rows.join('');
    this.attachTableInputListeners();
  }

  attachTableInputListeners() {
    this.elTableBody.querySelectorAll('.tbl-title').forEach(inp => {
      inp.addEventListener('input', (e) => {
        const idx = parseInt(e.target.dataset.index, 10);
        this.tags[idx].title = e.target.value.toUpperCase();
        this.saveToStorage();
        this.renderA4Sheets();
      });
    });

    this.elTableBody.querySelectorAll('.tbl-mrp').forEach(inp => {
      inp.addEventListener('input', (e) => {
        const idx = parseInt(e.target.dataset.index, 10);
        this.tags[idx].mrp = parseFloat(e.target.value) || 0;
        if (this.tags[idx].mrp > 0 && this.settings.mrpBadgeMode === 'hide') {
          this.settings.mrpBadgeMode = 'auto';
          if (this.elMrpBadgeToggle) this.elMrpBadgeToggle.value = 'auto';
        }
        this.updateTableRowSavings(idx);
        this.saveToStorage();
        this.renderA4Sheets();
      });
    });

    this.elTableBody.querySelectorAll('.tbl-offer').forEach(inp => {
      inp.addEventListener('input', (e) => {
        const idx = parseInt(e.target.dataset.index, 10);
        this.tags[idx].offerPrice = parseFloat(e.target.value) || 0;
        this.updateTableRowSavings(idx);
        this.saveToStorage();
        this.renderA4Sheets();
      });
    });

    this.elTableBody.querySelectorAll('.tbl-unit').forEach(inp => {
      inp.addEventListener('input', (e) => {
        const idx = parseInt(e.target.dataset.index, 10);
        this.tags[idx].unit = e.target.value.toUpperCase();
        this.saveToStorage();
        this.renderA4Sheets();
      });
    });

    this.elTableBody.querySelectorAll('.table-row-img-input').forEach(inp => {
      inp.addEventListener('change', (e) => {
        const idx = parseInt(e.target.dataset.index, 10);
        const file = e.target.files[0];
        if (file) {
          const reader = new FileReader();
          reader.onload = (ev) => {
            this.tags[idx].image = ev.target.result;
            this.saveToStorage();
            this.render();
          };
          reader.readAsDataURL(file);
        }
      });
    });

    this.elTableBody.querySelectorAll('.tbl-btn-edit').forEach(btn => {
      btn.addEventListener('click', () => {
        const idx = parseInt(btn.dataset.index, 10);
        this.switchTab('tab-designer');
        this.elDesignerTagSelector.value = idx;
        this.selectedTagIndex = idx;
        this.populateDesignerForm();
        this.renderDesignerTag();
      });
    });

    this.elTableBody.querySelectorAll('.tbl-btn-dup').forEach(btn => {
      btn.addEventListener('click', () => {
        const idx = parseInt(btn.dataset.index, 10);
        const copy = JSON.parse(JSON.stringify(this.tags[idx]));
        copy.id = `tag-${Date.now()}`;
        this.tags.splice(idx + 1, 0, copy);
        this.saveToStorage();
        this.render();
      });
    });

    this.elTableBody.querySelectorAll('.tbl-btn-del').forEach(btn => {
      btn.addEventListener('click', () => {
        const idx = parseInt(btn.dataset.index, 10);
        if (confirm(`Delete tag "${this.tags[idx].title}"?`)) {
          this.tags.splice(idx, 1);
          this.saveToStorage();
          this.render();
        }
      });
    });

    this.elTableBody.querySelectorAll('.tbl-btn-png').forEach(btn => {
      btn.addEventListener('click', () => {
        const idx = parseInt(btn.dataset.index, 10);
        if (this.tags[idx]) {
          this.exportSingleTagAsPNG(this.tags[idx]);
        }
      });
    });
  }

  updateTableRowSavings(idx) {
    const tr = this.elTableBody.querySelector(`tr[data-index="${idx}"]`);
    if (!tr) return;
    const tag = this.tags[idx];
    const badgeMode = this.settings.mrpBadgeMode || 'auto';
    const numMrp = Number(tag.mrp || 0);
    const numOffer = Number(tag.offerPrice || 0);
    const hasMrp = (badgeMode !== 'hide') && numMrp > 0 && numMrp > numOffer;
    const badge = tr.querySelector('.table-savings-badge');
    if (badge) {
      if (hasMrp) {
        const diff = Math.max(0, numMrp - numOffer).toFixed(2);
        badge.textContent = `${tag.storeName || 'SMILE'} SAVE ${this.settings.currency}${diff}`;
        badge.style.opacity = '1';
        badge.style.background = '';
        badge.style.color = '';
      } else {
        badge.textContent = 'No MRP (No Badge)';
        badge.style.opacity = '0.5';
        badge.style.background = '#e2e8f0';
        badge.style.color = '#475569';
      }
    }
    this.updateCounters();
  }

  exportCSV() {
    const rows = [['Item Name', 'MRP', 'Offer Price', 'Unit', 'Store Name']];
    this.tags.forEach(t => {
      rows.push([`"${t.title.replace(/"/g, '""')}"`, t.mrp, t.offerPrice, t.unit, t.storeName]);
    });
    const csvContent = "data:text/csv;charset=utf-8," + rows.map(e => e.join(",")).join("\n");
    const encodedUri = encodeURI(csvContent);
    const link = document.createElement("a");
    link.setAttribute("href", encodedUri);
    link.setAttribute("download", `smile_price_tags_${Date.now()}.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  }

  // =========================================================================
  // TAB 3: DESIGNER STUDIO & SINGLE TAG CUSTOMIZATION
  // =========================================================================
  populateDesignerSelector() {
    this.elDesignerTagSelector.innerHTML = this.tags.map((t, idx) => `
      <option value="${idx}" ${idx === this.selectedTagIndex ? 'selected' : ''}>
        #${idx + 1}: ${t.title || 'Untitled'} (${this.settings.currency}${t.offerPrice || 0})
      </option>
    `).join('');
  }

  populateDesignerForm() {
    const tag = this.tags[this.selectedTagIndex];
    if (!tag) return;

    const badgeMode = this.settings.mrpBadgeMode || 'auto';
    const numMrp = Number(tag.mrp || 0);
    const numOffer = Number(tag.offerPrice || 0);
    const hasMrp = (badgeMode !== 'hide') && numMrp > 0 && numMrp > numOffer;
    this.elDtTitle.value = tag.title || '';
    this.elDtMrp.value = numMrp > 0 ? tag.mrp : '';
    this.elDtOffer.value = tag.offerPrice || 0;
    this.elDtUnit.value = tag.unit || 'KG';
    this.elDtStoreName.value = tag.storeName || this.settings.storeName;

    if (hasMrp) {
      const diff = Math.max(0, numMrp - numOffer);
      this.elDtSavings.value = `${this.settings.currency}${diff.toFixed(2)}`;
    } else {
      this.elDtSavings.value = 'No MRP (No badge)';
    }

    const zoom = Math.round((tag.imageScale || 1.0) * 100);
    this.elDtZoomSlider.value = zoom;
    this.elDtZoomVal.textContent = `${zoom}%`;

    const panX = tag.imageX || 0;
    this.elDtPanXSlider.value = panX;
    this.elDtPanXVal.textContent = `${panX}px`;

    const panY = tag.imageY || 0;
    this.elDtPanYSlider.value = panY;
    this.elDtPanYVal.textContent = `${panY}px`;

    const titleSize = tag.titleSize || 100;
    this.elDtTitleSize.value = titleSize;
    this.elDtTitleSizeVal.textContent = `${titleSize}%`;

    const priceSize = tag.priceSize || 100;
    this.elDtPriceSize.value = priceSize;
    this.elDtPriceSizeVal.textContent = `${priceSize}%`;

    const stroke = tag.strokeWidth || 3;
    this.elDtStrokeWidth.value = stroke;
    this.elDtStrokeVal.textContent = `${stroke}px`;
  }

  updateSelectedTagField(field, value) {
    const tag = this.tags[this.selectedTagIndex];
    if (!tag) return;

    tag[field] = value;
    if (field === 'mrp' || field === 'offerPrice') {
      const badgeMode = this.settings.mrpBadgeMode || 'auto';
      const numMrp = Number(tag.mrp || 0);
      const numOffer = Number(tag.offerPrice || 0);
      const hasMrp = (badgeMode !== 'hide') && numMrp > 0 && numMrp > numOffer;
      if (hasMrp) {
        const diff = Math.max(0, numMrp - numOffer);
        this.elDtSavings.value = `${this.settings.currency}${diff.toFixed(2)}`;
      } else {
        this.elDtSavings.value = 'No MRP (No badge)';
      }
    }

    this.saveToStorage();
    this.renderDesignerTag();
    this.renderA4Sheets();
  }

  renderDesignerTag() {
    const tag = this.tags[this.selectedTagIndex];
    if (!tag) {
      this.elDesignerTagContainer.innerHTML = '<div style="padding: 20px; color:#fff;">No tag selected</div>';
      return;
    }

    this.elDesignerTagContainer.innerHTML = '';
    const tagElement = this.createTagDOMElement(tag, this.selectedTagIndex);
    tagElement.style.width = '100%';
    tagElement.style.height = '100%';
    this.elDesignerTagContainer.appendChild(tagElement);
    this.autoFitAllLongTitles(this.elDesignerTagContainer);
  }

  // =========================================================================
  // HIGH-DPI CANVAS RENDERING & EXPORT ENGINE (A4 Sheets & Single Tags)
  // =========================================================================
  async drawTagOnCanvas(ctx, tag, x, y, width, height, imgCache = null) {
    if (!tag) {
      // Draw empty slot with dashed border
      ctx.save();
      ctx.strokeStyle = '#cbd5e1';
      ctx.lineWidth = 2;
      ctx.setLineDash([8, 8]);
      ctx.strokeRect(x + 2, y + 2, width - 4, height - 4);
      ctx.restore();
      return;
    }

    ctx.save();

    // 1. Draw solid white background card
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(x, y, width, height);

    // 2. Product image background (negative z-index backdrop)
    if (tag.image) {
      try {
        const img = await this.loadImageAsync(tag.image, imgCache);
        ctx.save();
        ctx.beginPath();
        ctx.rect(x, y, width, height);
        ctx.clip();

        const scale = (tag.imageScale || 1.0);
        const panX = (tag.imageX || 0) * (width / 560);
        const panY = (tag.imageY || 0) * (height / 393);

        const dw = width;
        const dh = height;

        const dx = x + (width - dw * scale) / 2 + panX;
        const dy = y + (height - dh * scale) / 2 + panY;

        ctx.drawImage(img, dx, dy, dw * scale, dh * scale);
        ctx.restore();
      } catch (e) {
        console.warn('Could not load tag image for canvas export:', e);
      }
    }

    // 3. Outer black border
    ctx.strokeStyle = '#000000';
    ctx.lineWidth = (width > 800) ? 6 : 3.5;
    ctx.strokeRect(x + 1.5, y + 1.5, width - 3, height - 3);

    // 4. Badges (Only rendered when MRP is enabled, present, and greater than offerPrice)
    const badgeMode = this.settings.mrpBadgeMode || 'auto';
    const numMrp = Number(tag.mrp || 0);
    const numOffer = Number(tag.offerPrice || 0);
    const hasMrp = (badgeMode !== 'hide') &&
      numMrp > 0 &&
      numMrp > numOffer;
    const mrpVal = hasMrp ? numMrp.toFixed(1) : '';
    const savings = hasMrp ? Math.max(0, numMrp - numOffer).toFixed(2) : '0.00';
    const store = tag.storeName || this.settings.storeName;
    // 4. Store Brand Logo (Permanently on top of each price tag)
    const logoY = y + Math.round(height * 0.016);
    const logoMaxH = hasMrp ? Math.round(height * 0.11) : Math.round(height * 0.15);
    try {
      const logoImg = await this.loadImageAsync('BIGSAVE.PNG', imgCache);
      const aspect = (logoImg.naturalWidth || logoImg.width) / (logoImg.naturalHeight || logoImg.height);
      const logoH = logoMaxH;
      const logoW = Math.min(Math.round(logoH * aspect), Math.round(width * 0.52));
      const actualH = Math.round(logoW / aspect);
      const logoX = x + Math.round((width - logoW) / 2);

      // White pill background for crisp visibility against any backdrop
      ctx.save();
      ctx.fillStyle = '#ffffff';
      const padX = Math.round(width * 0.015);
      const padY = Math.max(1, Math.round(height * 0.005));
      ctx.fillRect(logoX - padX, logoY - padY, logoW + padX * 2, actualH + padY * 2);
      ctx.drawImage(logoImg, logoX, logoY, logoW, actualH);
      ctx.restore();
    } catch (e) {
      console.warn('Could not draw store logo on canvas:', e);
    }

    const badgeY = y + Math.round(height * 0.14);
    const badgeH = Math.round(height * 0.082);
    const badgeX = x + Math.round(width * 0.035);

    if (hasMrp) {
      // 1. MRP Amount (No Badge: Clean text with strikethrough)
      const mrpFontSize = Math.round(badgeH * 0.54);
      ctx.font = `900 ${mrpFontSize}px "Impact", "LocalHeavyCondensed", "Arial Black", sans-serif`;
      const mrpTextCenterY = badgeY + badgeH / 2;

      const mrpPrefixW = ctx.measureText('MRP ').width;
      const mrpNumW = ctx.measureText(mrpVal).width;
      const mrpTotalW = mrpPrefixW + mrpNumW;

      // Subtle backing for clean legibility against background photo
      ctx.save();
      ctx.fillStyle = 'rgba(255, 255, 255, 0.88)';
      ctx.fillRect(badgeX - 2, badgeY, mrpTotalW + 6, badgeH);
      ctx.restore();

      ctx.fillStyle = '#000000';
      ctx.textAlign = 'left';
      ctx.textBaseline = 'middle';
      ctx.fillText('MRP ', badgeX + 2, mrpTextCenterY);

      const mrpNumX = badgeX + 2 + mrpPrefixW;
      ctx.fillText(mrpVal, mrpNumX, mrpTextCenterY);

      // Strike-through line on MRP amount
      ctx.strokeStyle = '#000000';
      ctx.lineWidth = Math.max(2, Math.round(badgeH * 0.09));
      ctx.beginPath();
      ctx.moveTo(mrpNumX - 2, mrpTextCenterY);
      ctx.lineTo(mrpNumX + mrpNumW + 2, mrpTextCenterY);
      ctx.stroke();

      // 2. Red SAVE Badge (Prominent colored badge)
      const saveBadgeX = badgeX + mrpTotalW + Math.round(width * 0.03);
      const saveBadgeW = (x + width - Math.round(width * 0.035)) - saveBadgeX;

      ctx.fillStyle = '#d80000';
      ctx.fillRect(saveBadgeX, badgeY, saveBadgeW, badgeH);

      const saveFontSize = Math.round(badgeH * 0.48);
      ctx.font = `800 ${saveFontSize}px "Impact", "LocalHeavyCondensed", "Arial Black", sans-serif`;
      ctx.fillStyle = '#ffffff';
      ctx.textAlign = 'center';
      ctx.fillText(`${store} SAVE ${this.settings.currency}${savings}`, saveBadgeX + saveBadgeW / 2, mrpTextCenterY);
    }

    // 5. Product Title (Consumes upper section with Thin White Outline)
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    const globalBoost = (this.settings.globalFontScale || 135) / 100;
    const titleStr = (tag.title || 'SPECIAL OFFER').trim().toUpperCase();

    const words = titleStr.split(/\s+/).filter(w => w.length > 0);
    let lines = [titleStr];
    let baseTitleRatio = (width > 800) ? 0.095 : 0.125;

    if (words.length > 1 && titleStr.length > 10) {
      // Find the most balanced two-line split
      let bestSplitIdx = 1;
      let minDiff = Infinity;
      for (let i = 1; i < words.length; i++) {
        const line1 = words.slice(0, i).join(' ');
        const line2 = words.slice(i).join(' ');
        const diff = Math.abs(line1.length - line2.length);
        if (diff < minDiff) {
          minDiff = diff;
          bestSplitIdx = i;
        }
      }
      lines = [words.slice(0, bestSplitIdx).join(' '), words.slice(bestSplitIdx).join(' ')];
      baseTitleRatio = (width > 800) ? 0.080 : 0.100;
    } else if (titleStr.length > 16) {
      baseTitleRatio = (width > 800) ? 0.070 : 0.088;
    }

    const maxTitleByHeight = hasMrp ? Math.round(height * 0.20) : Math.round(height * 0.26);
    let titleFontSize = Math.min(
      Math.round(width * baseTitleRatio * ((tag.titleSize || 100) / 100) * (globalBoost / 1.15)),
      maxTitleByHeight
    );
    if (titleFontSize < 14) titleFontSize = 14;
    ctx.font = `900 ${titleFontSize}px "Impact", "LocalHeavyCondensed", "Arial Black", "Anton", sans-serif`;

    const maxTitleWidth = width - 16;
    let maxLineW = 0;
    lines.forEach(line => {
      const lw = ctx.measureText(line).width;
      if (lw > maxLineW) maxLineW = lw;
    });

    if (maxLineW > maxTitleWidth) {
      titleFontSize = Math.max(10, Math.floor(titleFontSize * (maxTitleWidth / maxLineW)));
      ctx.font = `900 ${titleFontSize}px "Impact", "LocalHeavyCondensed", "Arial Black", "Anton", sans-serif`;
    }

    const titleCenterY = hasMrp 
      ? y + Math.round(height * 0.49) 
      : y + Math.round(height * 0.42);
    const lineHeight = titleFontSize * 0.92;

    ctx.strokeStyle = '#ffffff';
    ctx.lineWidth = Math.max(3, Math.round(titleFontSize * 0.075));
    ctx.lineJoin = 'miter';
    ctx.miterLimit = 2;

    const centerX = x + width / 2;
    if (lines.length === 1) {
      ctx.strokeText(lines[0], centerX, titleCenterY);
      ctx.fillStyle = '#000000';
      ctx.fillText(lines[0], centerX, titleCenterY);
    } else {
      const startY = titleCenterY - (lineHeight / 2);
      lines.forEach((line, i) => {
        const lineY = startY + (i * lineHeight);
        ctx.strokeText(line, centerX, lineY);
        ctx.fillStyle = '#000000';
        ctx.fillText(line, centerX, lineY);
      });
    }

    // 6. Price & Unit (Massive Thick Black Font, Auto-fit into the Column Width and consuming lower section)
    const maxPriceWidth = width - 20; // 10px margin on each side of column
    const basePriceRatio = (width > 800) ? 0.15 : 0.195;
    const maxPriceByHeight = hasMrp ? Math.round(height * 0.26) : Math.round(height * 0.34);
    let priceFontSize = Math.min(
      Math.round(width * basePriceRatio * ((tag.priceSize || 100) / 100) * (globalBoost / 1.15)),
      maxPriceByHeight
    );
    ctx.font = `900 ${priceFontSize}px "Impact", "LocalHeavyCondensed", "Arial Black", "Anton", sans-serif`;
    ctx.textBaseline = 'middle';

    const offerVal = Number(tag.offerPrice || 0).toFixed(2);
    const unitText = tag.unit ? ` / ${tag.unit}` : '';
    const fullPriceStr = `${offerVal}${unitText}`;

    // Measure exact rendered text width on canvas and auto-fit to column
    let measuredPriceWidth = ctx.measureText(fullPriceStr).width;
    if (measuredPriceWidth > maxPriceWidth) {
      priceFontSize = Math.max(18, Math.floor(priceFontSize * (maxPriceWidth / measuredPriceWidth)));
      ctx.font = `900 ${priceFontSize}px "Impact", "LocalHeavyCondensed", "Arial Black", "Anton", sans-serif`;
    }

    const priceY = hasMrp 
      ? y + Math.round(height * 0.82) 
      : y + Math.round(height * 0.78);
    ctx.strokeStyle = '#ffffff';
    ctx.lineWidth = Math.max(3, Math.round(priceFontSize * 0.065));
    ctx.strokeText(fullPriceStr, centerX, priceY);

    ctx.fillStyle = '#000000';
    ctx.fillText(fullPriceStr, centerX, priceY);

    ctx.restore();
  }

  async exportSingleTagAsPNG(tag) {
    if (document.fonts) {
      try {
        await document.fonts.ready;
      } catch (_) {}
    }

    const canvas = document.createElement('canvas');
    const width = 1120;
    const height = 786;
    canvas.width = width;
    canvas.height = height;
    const ctx = canvas.getContext('2d');
    const imgCache = new Map();

    await this.drawTagOnCanvas(ctx, tag, 0, 0, width, height, imgCache);

    const link = document.createElement('a');
    link.download = `${(tag.title || 'tag').toLowerCase().replace(/\s+/g, '_')}_pricetag.png`;
    link.href = canvas.toDataURL('image/png');
    link.click();
    this.notify(`Downloaded ${link.download}!`, 'success');
  }

  async exportA4SheetAsPNG(pageIdx = 0, totalPages = 1) {
    if (document.fonts) {
      try {
        await document.fonts.ready;
      } catch (_) {}
    }

    const config = this.getGridConfig();
    const isLandscape = config.orientation === 'landscape';
    const cols = config.cols;
    const rows = config.rows;
    const perPage = config.perPage;

    // 300 DPI: Portrait 2480x3508 (210x297mm), Landscape 3508x2480 (297x210mm)
    const a4W = isLandscape ? 3508 : 2480;
    const a4H = isLandscape ? 2480 : 3508;

    const canvas = document.createElement('canvas');
    canvas.width = a4W;
    canvas.height = a4H;
    const ctx = canvas.getContext('2d');

    // Solid white A4 sheet background
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, a4W, a4H);

    // Margins (6mm at 300 DPI = 71px)
    const marginX = 71;
    const marginY = 71;
    const gapX = 24; // 2mm gap between tags
    const gapY = 24; // 2mm gap between tags

    const printW = a4W - (marginX * 2);
    const printH = a4H - (marginY * 2);

    const tagW = Math.floor((printW - (cols - 1) * gapX) / cols);
    const tagH = Math.floor((printH - (rows - 1) * gapY) / rows);

    const startIndex = pageIdx * perPage;
    const endIndex = Math.min(startIndex + perPage, this.tags.length);
    const imgCache = new Map();

    // Draw all tags for this page
    for (let slot = 0; slot < perPage; slot++) {
      const tagIndex = startIndex + slot;
      const tag = tagIndex < endIndex ? this.tags[tagIndex] : null;

      const c = slot % cols;
      const r = Math.floor(slot / cols);

      const tagX = marginX + c * (tagW + gapX);
      const tagY = marginY + r * (tagH + gapY);

      await this.drawTagOnCanvas(ctx, tag, tagX, tagY, tagW, tagH, imgCache);

      // Cut guides if enabled
      if (this.settings.showCutGuides) {
        ctx.save();
        ctx.strokeStyle = '#94a3b8';
        ctx.lineWidth = 1.5;
        ctx.setLineDash([8, 8]);
        ctx.strokeRect(tagX, tagY, tagW, tagH);
        ctx.restore();
      }
    }

    // Corner crop ticks for cutting
    if (this.settings.showCutGuides) {
      ctx.save();
      ctx.strokeStyle = '#64748b';
      ctx.lineWidth = 2;
      const tick = 28;
      // Top-Left
      ctx.beginPath();
      ctx.moveTo(marginX - tick, marginY); ctx.lineTo(marginX, marginY);
      ctx.moveTo(marginX, marginY - tick); ctx.lineTo(marginX, marginY);
      // Top-Right
      ctx.moveTo(a4W - marginX + tick, marginY); ctx.lineTo(a4W - marginX, marginY);
      ctx.moveTo(a4W - marginX, marginY - tick); ctx.lineTo(a4W - marginX, marginY);
      // Bottom-Left
      ctx.beginPath();
      ctx.moveTo(marginX - tick, a4H - marginY); ctx.lineTo(marginX, a4H - marginY);
      ctx.moveTo(marginX, a4H - marginY + tick); ctx.lineTo(marginX, a4H - marginY);
      // Bottom-Right
      ctx.moveTo(a4W - marginX + tick, a4H - marginY); ctx.lineTo(a4W - marginX, a4H - marginY);
      ctx.moveTo(a4W - marginX, a4H - marginY + tick); ctx.lineTo(a4W - marginX, a4H - marginY);
      ctx.stroke();
      ctx.restore();
    }

    // Sheet Footer info line
    ctx.save();
    ctx.font = '600 24px "Inter", sans-serif';
    ctx.fillStyle = '#64748b';
    ctx.textAlign = 'right';
    ctx.fillText(`${this.settings.storeName} Supermarket | A4 ${isLandscape ? 'Landscape' : 'Portrait'} Sheet (${pageIdx + 1} of ${totalPages}) | ${config.cols}x${config.rows} Grid (${perPage} Tags) | 300 DPI Print Quality`, a4W - marginX, a4H - 24);
    ctx.restore();

    const orientationSlug = isLandscape ? 'landscape_' : '';
    const fileName = totalPages > 1 
      ? `pricetags_a4_${orientationSlug}sheet_${pageIdx + 1}_of_${totalPages}.png`
      : `pricetags_a4_${orientationSlug}sheet_${pageIdx + 1}.png`;

    return new Promise((resolve) => {
      canvas.toBlob((blob) => {
        if (!blob) {
          resolve(false);
          return;
        }
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.download = fileName;
        link.href = url;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        setTimeout(() => {
          URL.revokeObjectURL(url);
          resolve(true);
        }, 400);
      }, 'image/png');
    });
  }

  async downloadAllA4Sheets() {
    const config = this.getGridConfig();
    const isLandscape = config.orientation === 'landscape';
    const perPage = config.perPage;
    const totalTags = this.tags.length;
    const numPages = Math.max(1, Math.ceil(totalTags / perPage));

    if (numPages === 1) {
      this.notify(`Generating 300 DPI A4 ${isLandscape ? 'Landscape ' : ''}Sheet with all ${totalTags} tags to print together...`, 'info');
      await this.exportA4SheetAsPNG(0, 1);
      this.notify(`Downloaded A4 ${isLandscape ? 'Landscape ' : ''}Sheet (Ready to print together)!`, 'success');
    } else {
      this.notify(`Generating ${numPages} A4 ${isLandscape ? 'Landscape ' : ''}Sheets to download...`, 'info');
      for (let p = 0; p < numPages; p++) {
        this.notify(`Generating A4 ${isLandscape ? 'Landscape ' : ''}Sheet ${p + 1} of ${numPages}...`, 'info');
        await this.exportA4SheetAsPNG(p, numPages);
        if (p < numPages - 1) {
          await new Promise(r => setTimeout(r, 700));
        }
      }
      this.notify(`Successfully downloaded all ${numPages} A4 ${isLandscape ? 'Landscape ' : ''}Sheets!`, 'success');
    }
  }

  loadImageAsync(src, cache = null) {
    if (cache && cache.has(src)) {
      return Promise.resolve(cache.get(src));
    }
    return new Promise((resolve, reject) => {
      const img = new Image();
      img.crossOrigin = 'anonymous';
      img.onload = () => {
        if (cache) cache.set(src, img);
        resolve(img);
      };
      img.onerror = () => {
        // If external image fails due to CORS or direct connection, fallback to server proxy
        if (src && !src.startsWith('/api/proxy-image') && !src.startsWith('data:') && !src.startsWith('blob:')) {
          const proxyUrl = `/api/proxy-image?url=${encodeURIComponent(src)}`;
          const proxyImg = new Image();
          proxyImg.crossOrigin = 'anonymous';
          proxyImg.onload = () => {
            if (cache) cache.set(src, proxyImg);
            resolve(proxyImg);
          };
          proxyImg.onerror = () => reject(new Error(`Failed to load image: ${src}`));
          proxyImg.src = proxyUrl;
        } else {
          reject(new Error(`Failed to load image: ${src}`));
        }
      };
      img.src = src;
    });
  }
}

// Initialize on DOM load
window.addEventListener('DOMContentLoaded', () => {
  window.priceTagApp = new PriceTagApp();
});
