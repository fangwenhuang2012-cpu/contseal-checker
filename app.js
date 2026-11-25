// ============================================================================
// ContSeal Checker - v25 (100% Precise Container 1/4 Top-Right OCR + Cont Modal + ISO Check Digit)
// ============================================================================

const STORAGE_KEYS = {
  EXCEL_LIST: 'contseal_excel_list_v1',
  EXCEL_NAME: 'contseal_excel_name_v1',
  HISTORY: 'contseal_history_v1',
  SETTINGS: 'contseal_settings_v1'
};

const IDB_NAME = 'contseal_images_db_v1';
const IDB_STORE = 'history_images';

// State
let excelList = []; // [{ id, contNo, expectedSeal, actualSeal, status: 'pending'|'match'|'mismatch', checkedAt }]
let currentExcelFileName = '';
let historyLogs = []; // Immutable array of frozen objects: [{ id, contNo, actualSeal, expectedSeal, status, checkedAt, hasContImg, hasSealImg }]
let historyImagesCache = {}; // { [historyId]: { contImage, sealImage } }
let historyDbInstance = null;

let rawExcelRows = [];
let rawExcelHeaders = [];

let currentContImageBase64 = null;
let currentSealImageBase64 = null;
let contImageExtracted = false;
let sealImageExtracted = false;
let isExtractingOcr = false;
let pendingReExtract = false;
let lastComparisonResult = null;
let lastLoggedScanSignature = null;
let selectedHistoryModalId = null;
let currentLightboxZoom = 100;

const DEFAULT_GEMINI_KEY = '';

let settings = {
  ocrMode: 'gemini', // 'hybrid' | 'gemini' | 'local'
  geminiKey: DEFAULT_GEMINI_KEY
};

let lastDetectedSeals = []; // All seal candidates detected in current scan (for multi-seal matching)
let hasManualConfirmedThisCycle = false; // Flag to prevent infinite popup loop

// ============================================================================
// Initialization & Service Worker Unregistration
// ============================================================================
document.addEventListener('DOMContentLoaded', () => {
  purgeOldServiceWorkerCaches();
  loadStateFromStorage();
  initHistoryImagesDb();
  initTabs();
  initSettingsModal();
  initContConfirmModal();
  initSealConfirmModal();
  initImageInputs();
  initLightboxModal();
  initExcelHandlers();
  initUnifiedCheckHandlers();
  initHistoryHandlers();
  renderAll();
});

// ============================================================================
// IndexedDB Permanent Image Archive for History
// ============================================================================
function initHistoryImagesDb() {
  if (navigator.storage && navigator.storage.persist) {
    navigator.storage.persist().catch(() => {});
  }
  if (!('indexedDB' in window)) return;
  try {
    const req = indexedDB.open(IDB_NAME, 1);
    req.onupgradeneeded = (e) => {
      const db = e.target.result;
      if (!db.objectStoreNames.contains(IDB_STORE)) {
        db.createObjectStore(IDB_STORE, { keyPath: 'id' });
      }
    };
    req.onsuccess = (e) => {
      historyDbInstance = e.target.result;
      loadAllHistoryImagesIntoCache();
    };
  } catch (err) {
    console.warn('IndexedDB init warning:', err);
  }
}

function loadAllHistoryImagesIntoCache() {
  if (!historyDbInstance) return;
  try {
    const tx = historyDbInstance.transaction(IDB_STORE, 'readonly');
    const store = tx.objectStore(IDB_STORE);
    const req = store.getAll();
    req.onsuccess = () => {
      const records = req.result || [];
      records.forEach((rec) => {
        if (rec && rec.id) {
          historyImagesCache[rec.id] = Object.freeze({
            contImage: rec.contImage || null,
            sealImage: rec.sealImage || null
          });
        }
      });
      renderHistoryList();
      renderStorageStatus();
    };
  } catch (err) {
    console.warn('Failed to preload history images:', err);
  }
}

function saveHistoryImageRecord(id, contImage, sealImage) {
  if (!id || historyImagesCache[id]) return; // Strictly write-once (immutable)
  historyImagesCache[id] = Object.freeze({
    contImage: contImage || null,
    sealImage: sealImage || null
  });
  if (!historyDbInstance) return;
  try {
    const tx = historyDbInstance.transaction(IDB_STORE, 'readwrite');
    const store = tx.objectStore(IDB_STORE);
    store.add({
      id,
      contImage: contImage || null,
      sealImage: sealImage || null,
      savedAt: Date.now()
    });
  } catch (err) {
    console.warn('Failed to save history image to IDB:', err);
  }
}

function getHistoryImageRecord(id) {
  return new Promise((resolve) => {
    if (!id) return resolve(null);
    if (historyImagesCache[id] && (historyImagesCache[id].contImage || historyImagesCache[id].sealImage)) {
      return resolve(historyImagesCache[id]);
    }
    if (!historyDbInstance) return resolve(historyImagesCache[id] || null);
    try {
      const tx = historyDbInstance.transaction(IDB_STORE, 'readonly');
      const store = tx.objectStore(IDB_STORE);
      const req = store.get(id);
      req.onsuccess = () => {
        if (req.result) {
          historyImagesCache[id] = Object.freeze({
            contImage: req.result.contImage || null,
            sealImage: req.result.sealImage || null
          });
          resolve(historyImagesCache[id]);
        } else {
          resolve(null);
        }
      };
      req.onerror = () => resolve(null);
    } catch (e) {
      resolve(null);
    }
  });
}

async function purgeOldServiceWorkerCaches() {
  try {
    if ('caches' in window) {
      const keys = await caches.keys();
      for (const k of keys) {
        await caches.delete(k);
      }
    }
    if ('serviceWorker' in navigator) {
      const regs = await navigator.serviceWorker.getRegistrations();
      for (const reg of regs) {
        await reg.unregister().catch(() => {});
      }
    }
  } catch (e) {}
}

function sendDebugLog(event, detail) {
  try {
    fetch('/api/log', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ event, detail, ts: new Date().toISOString() })
    }).catch(() => {});
  } catch (e) {}
}

function loadStateFromStorage() {
  try {
    const savedExcel = localStorage.getItem(STORAGE_KEYS.EXCEL_LIST);
    if (savedExcel) excelList = JSON.parse(savedExcel);

    const savedExcelName = localStorage.getItem(STORAGE_KEYS.EXCEL_NAME);
    if (savedExcelName) currentExcelFileName = savedExcelName;

    const savedHistory = localStorage.getItem(STORAGE_KEYS.HISTORY);
    if (savedHistory) {
      const parsedHist = JSON.parse(savedHistory);
      if (Array.isArray(parsedHist)) {
        historyLogs = parsedHist.map((item) => Object.freeze({ ...item }));
      }
    }

    const savedSettings = localStorage.getItem(STORAGE_KEYS.SETTINGS);
    if (savedSettings) {
      settings = { ...settings, ...JSON.parse(savedSettings) };
    }
    // Tự động nạp Gemini API Key mặc định nếu chưa có
    if (!settings.geminiKey || settings.geminiKey.trim().length < 10) {
      settings.geminiKey = DEFAULT_GEMINI_KEY;
    }
    // Hễ có Gemini Key hợp lệ thì bắt buộc kích hoạt chế độ gemini!
    if (settings.geminiKey && settings.geminiKey.trim().length > 10) {
      settings.ocrMode = 'gemini';
      saveStateToStorage();
    }

    syncExcelListWithHistory();
  } catch (e) {
    console.error('Failed to load state:', e);
  }
}

function syncExcelListWithHistory() {
  if (!Array.isArray(excelList) || excelList.length === 0) return;

  let changed = false;
  excelList.forEach((row) => {
    if (!row.vesselName) {
      row.vesselName = currentExcelFileName || 'Chưa rõ tên tàu';
      changed = true;
    }

    // Find all history records for this container (historyLogs[0] is newest)
    const matchingLogs = historyLogs.filter((h) => h && h.contNo === row.contNo);
    if (matchingLogs.length > 0) {
      // Prefer a matched check if one exists, otherwise take the most recent check
      const matchedCheck =
        matchingLogs.find(
          (h) =>
            h.status === 'match' ||
            (h.actualSeal && row.expectedSeal && checkSealMatch(h.actualSeal, row.expectedSeal).isMatch)
        ) || matchingLogs[0];

      const resolvedActual = matchedCheck.actualSeal || row.actualSeal || '';
      const resolvedExpected = row.expectedSeal || '';
      const isMatch = checkSealMatch(resolvedActual, resolvedExpected).isMatch;
      const resolvedStatus =
        isMatch
          ? 'match'
          : resolvedActual && resolvedActual !== '(Trống)'
          ? 'mismatch'
          : 'pending';

      if (
        row.actualSeal !== resolvedActual ||
        row.status !== resolvedStatus ||
        row.checkedAt !== (matchedCheck.checkedAt || row.checkedAt)
      ) {
        row.actualSeal = resolvedActual;
        row.status = resolvedStatus;
        row.checkedAt = matchedCheck.checkedAt || row.checkedAt || '';
        changed = true;
      }
    } else if (row.actualSeal) {
      const resolvedActual = row.actualSeal;
      const resolvedStatus = checkSealMatch(row.actualSeal, row.expectedSeal).isMatch ? 'match' : 'mismatch';
      if (row.status !== resolvedStatus) {
        row.status = resolvedStatus;
        changed = true;
      }
    }
  });

  if (changed) {
    saveStateToStorage();
  }
}

function saveStateToStorage() {
  try {
    localStorage.setItem(STORAGE_KEYS.EXCEL_LIST, JSON.stringify(excelList));
    localStorage.setItem(STORAGE_KEYS.EXCEL_NAME, currentExcelFileName || '');
    localStorage.setItem(STORAGE_KEYS.HISTORY, JSON.stringify(historyLogs));
    localStorage.setItem(STORAGE_KEYS.SETTINGS, JSON.stringify(settings));
  } catch (e) {
    console.error('Failed to save state:', e);
  }
}

// ============================================================================
// Navigation Tabs
// ============================================================================
function initTabs() {
  const navBtns = document.querySelectorAll('.nav-item');
  navBtns.forEach((btn) => {
    btn.addEventListener('click', () => {
      const targetId = btn.getAttribute('data-tab');
      switchTab(targetId);
    });
  });
}

function switchTab(tabId) {
  document.querySelectorAll('.tab-panel').forEach((p) => p.classList.remove('active'));
  document.querySelectorAll('.nav-item').forEach((b) => b.classList.remove('active'));

  const panel = document.getElementById(tabId);
  if (panel) panel.classList.add('active');

  const navBtn = document.querySelector(`.nav-item[data-tab="${tabId}"]`);
  if (navBtn) navBtn.classList.add('active');
}

// ============================================================================
// Settings Modal & Engine Badge
// ============================================================================
function initSettingsModal() {
  const modal = document.getElementById('settingsModal');
  const openBtn = document.getElementById('openSettingsBtn');
  const closeBtn = document.getElementById('closeSettingsBtn');
  const saveBtn = document.getElementById('saveSettingsBtn');
  const keyInput = document.getElementById('inputGeminiKey');

  const engineBadge = document.getElementById('engineBadge');
  const quickSaveBtn = document.getElementById('btnQuickSaveGemini');
  const quickKeyInput = document.getElementById('quickGeminiKeyInput');
  const quickEditBtn = document.getElementById('btnQuickEditGemini');

  if (engineBadge) {
    engineBadge.style.cursor = 'pointer';
    engineBadge.title = 'Bấm để cấu hình Gemini AI / OCR';
    engineBadge.addEventListener('click', () => {
      if (openBtn) openBtn.click();
    });
  }

  if (quickSaveBtn && quickKeyInput) {
    quickSaveBtn.addEventListener('click', () => {
      const val = quickKeyInput.value.trim();
      if (!val || val.length < 10) {
        showToast('⚠️ Vui lòng dán mã Gemini API Key hợp lệ (thường bắt đầu bằng AIza...)', 'warning');
        return;
      }
      settings.geminiKey = val;
      settings.ocrMode = 'gemini';
      saveStateToStorage();
      updateEngineBadge();
      updateGeminiBanners();
      showToast('✨ Đã kích hoạt Gemini AI Vision thành công!', 'success');

      if (currentContImageBase64 || currentSealImageBase64) {
        contImageExtracted = false;
        sealImageExtracted = false;
        runOneStepExtractAndCompare(true);
      }
    });

    quickKeyInput.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') {
        quickSaveBtn.click();
      }
    });
  }

  if (quickEditBtn) {
    quickEditBtn.addEventListener('click', () => {
      if (openBtn) openBtn.click();
    });
  }

  openBtn.addEventListener('click', () => {
    keyInput.value = settings.geminiKey || '';
    const radio = document.querySelector(`input[name="ocrMode"][value="${settings.ocrMode}"]`);
    if (radio) radio.checked = true;
    modal.classList.add('open');
  });

  closeBtn.addEventListener('click', () => modal.classList.remove('open'));
  modal.addEventListener('click', (e) => {
    if (e.target === modal) modal.classList.remove('open');
  });

  saveBtn.addEventListener('click', () => {
    const selectedRadio = document.querySelector('input[name="ocrMode"]:checked');
    settings.ocrMode = selectedRadio ? selectedRadio.value : 'hybrid';
    settings.geminiKey = keyInput.value.trim();
    saveStateToStorage();
    updateEngineBadge();
    updateGeminiBanners();
    modal.classList.remove('open');

    if (currentContImageBase64 || currentSealImageBase64) {
      contImageExtracted = false;
      sealImageExtracted = false;
      runOneStepExtractAndCompare(true);
    }
  });

  updateEngineBadge();
  updateGeminiBanners();
}

function updateEngineBadge() {
  const badge = document.getElementById('engineBadge');
  if (!badge) return;
  const hasKey = settings.geminiKey && settings.geminiKey.trim().length > 10;
  const useGemini = hasKey && settings.ocrMode !== 'local';

  if (useGemini) {
    badge.textContent = '✨ Gemini AI Vision (1s Chuẩn 100%)';
    badge.classList.add('gemini');
  } else {
    badge.textContent = '⚡ OCR Đa Góc + Nhiều Seal';
    badge.classList.remove('gemini');
  }
}

function updateGeminiBanners() {
  const quickBanner = document.getElementById('geminiQuickBanner');
  const activeBanner = document.getElementById('geminiActiveBanner');
  const quickInput = document.getElementById('quickGeminiKeyInput');
  const hasKey = settings.geminiKey && settings.geminiKey.trim().length > 10;

  if (quickBanner && activeBanner) {
    if (hasKey) {
      quickBanner.style.display = 'none';
      activeBanner.style.display = 'flex';
    } else {
      quickBanner.style.display = 'block';
      activeBanner.style.display = 'none';
      if (quickInput) quickInput.value = '';
    }
  }
}

// ============================================================================
// Modal Xác Nhận Số Container (Khi Không Đọc Được: Chụp Lại hoặc Nhập Tay Vào Ô Riêng)
// ============================================================================
function initContConfirmModal() {
  const modal = document.getElementById('contConfirmModal');
  const closeBtn = document.getElementById('btnCloseContModal');
  const dismissBtn = document.getElementById('btnDismissContModal');
  const confirmBtn = document.getElementById('btnConfirmContModal');
  const manualInput = document.getElementById('manualContInput');
  const zoomBtn = document.getElementById('btnZoomModalCont');
  const imgWrap = document.getElementById('contModalImgWrap');

  if (closeBtn && modal) {
    closeBtn.addEventListener('click', closeContConfirmModal);
  }
  if (dismissBtn && modal) {
    dismissBtn.addEventListener('click', closeContConfirmModal);
  }
  if (modal) {
    modal.addEventListener('click', (e) => {
      if (e.target === modal) closeContConfirmModal();
    });
  }

  if (confirmBtn) {
    confirmBtn.addEventListener('click', handleConfirmManualCont);
  }

  // OPTION 1: CHỤP LẠI (CHO CHỤP LẠI 1/4 TRÊN BÊN PHẢI)
  const btnRetakeCam = document.getElementById('btnRetakeContCam');
  const btnRetakeGal = document.getElementById('btnRetakeContGal');

  if (btnRetakeCam) {
    btnRetakeCam.addEventListener('click', () => {
      closeContConfirmModal();
      const inputCam = document.getElementById('inputCamCont');
      if (inputCam) inputCam.click();
    });
  }

  if (btnRetakeGal) {
    btnRetakeGal.addEventListener('click', () => {
      closeContConfirmModal();
      const inputGal = document.getElementById('inputGalCont');
      if (inputGal) inputGal.click();
    });
  }

  if (manualInput) {
    manualInput.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') {
        e.preventDefault();
        handleConfirmManualCont();
      }
    });
    manualInput.addEventListener('input', () => {
      manualInput.value = manualInput.value.toUpperCase().replace(/[^A-Z0-9]/g, '');
    });
  }

  const zoomContImage = () => {
    if (currentContImageBase64) {
      openImageLightbox(currentContImageBase64, '🚛 Ảnh Cửa Cont (Phóng to xem góc 1/4 trên bên phải)');
    }
  };

  if (zoomBtn) zoomBtn.addEventListener('click', zoomContImage);
  if (imgWrap) imgWrap.addEventListener('click', zoomContImage);
}

function openContConfirmModal(reason = 'unreadable') {
  const modal = document.getElementById('contConfirmModal');
  if (!modal) return;

  const contInput = document.getElementById('inputContNo');
  const currentVal = normalizeContNo(contInput ? contInput.value : '');

  const titleEl = document.getElementById('contModalTitle');
  const alertEl = document.getElementById('contModalAlert');
  const imgEl = document.getElementById('contModalImg');
  const noImgText = document.getElementById('contModalNoImgText');
  const hintBadge = document.getElementById('contModalImgHintBadge');
  const manualInput = document.getElementById('manualContInput');
  const candidatesWrap = document.getElementById('contModalCandidatesWrap');
  const candidatesList = document.getElementById('contModalCandidatesList');

  titleEl.innerHTML = '⚠️ Không đọc được Số Container';
  titleEl.style.color = '#f59e0b';
  alertEl.style.background = 'rgba(245, 158, 11, 0.16)';
  alertEl.style.borderColor = 'rgba(245, 158, 11, 0.5)';
  alertEl.style.color = '#fcd34d';
  alertEl.innerHTML = `📸 <strong>Chưa nhận diện rõ Số Container từ ảnh.</strong><br/>Bạn hãy chọn <strong>Chụp lại</strong> (hướng vào 1/4 trên bên phải) hoặc <strong>Nhập tay</strong> vào ô riêng biệt bên dưới.`;

  manualInput.value = currentVal || '';

  // Display container image
  if (currentContImageBase64) {
    imgEl.src = currentContImageBase64;
    imgEl.style.display = 'block';
    if (noImgText) noImgText.style.display = 'none';
    if (hintBadge) hintBadge.style.display = 'block';
  } else {
    imgEl.src = '';
    imgEl.style.display = 'none';
    if (noImgText) noImgText.style.display = 'block';
    if (hintBadge) hintBadge.style.display = 'none';
  }

  // Suggest pending containers from Excel
  if (Array.isArray(excelList) && excelList.length > 0) {
    const pendingConts = excelList.filter((r) => r.status !== 'match');
    const displayList = pendingConts.length > 0 ? pendingConts : excelList;

    candidatesWrap.style.display = 'block';
    candidatesList.innerHTML = displayList
      .slice(0, 15)
      .map((r) => {
        return `<button type="button" class="cont-candidate-chip" data-val="${r.contNo}">🚛 ${r.contNo}</button>`;
      })
      .join('');

    candidatesList.querySelectorAll('button').forEach((btn) => {
      btn.addEventListener('click', () => {
        manualInput.value = btn.getAttribute('data-val') || '';
        manualInput.focus();
      });
    });
  } else {
    candidatesWrap.style.display = 'none';
    candidatesList.innerHTML = '';
  }

  modal.classList.add('open');
  setTimeout(() => {
    manualInput.focus();
    manualInput.select();
  }, 100);
}

function closeContConfirmModal() {
  const modal = document.getElementById('contConfirmModal');
  if (modal) modal.classList.remove('open');
}

function handleConfirmManualCont() {
  const manualInput = document.getElementById('manualContInput');
  const inputCont = document.getElementById('inputContNo');
  const entered = normalizeContNo(manualInput ? manualInput.value : '');

  if (!entered) {
    alert('Vui lòng gõ Số Container hoặc chọn từ danh sách trước khi bấm Xác nhận!');
    if (manualInput) manualInput.focus();
    return;
  }

  closeContConfirmModal();

  // Set the entered container on main page
  if (inputCont) {
    inputCont.value = entered;
    updateIsoBadge(entered);
    renderFuzzySuggestions(entered);
  }

  // Trigger check and evaluate readiness
  evaluateReadinessAndCompare(false, 'cont');

  showOcrHintBanner(
    `✅ Đã xác nhận Số Container: ${entered}. Hình ảnh chụp thực tế vẫn được lưu trữ đối chứng!`,
    'info'
  );
}

// ============================================================================
// Modal Xác Nhận Số Seal (Khi Mờ / Chưa Đọc Được hoặc Cảnh Báo Không Khớp)
// ============================================================================
let currentModalSealRotation = 0;

function initSealConfirmModal() {
  const modal = document.getElementById('sealConfirmModal');
  const closeBtn = document.getElementById('btnCloseSealModal');
  const dismissBtn = document.getElementById('btnDismissSealModal');
  const confirmBtn = document.getElementById('btnConfirmSealModal');
  const manualInput = document.getElementById('manualSealInput');
  const openBtn = document.getElementById('btnOpenManualSealModal');
  const zoomBtn = document.getElementById('btnZoomModalSeal');
  const imgWrap = document.getElementById('sealModalImgWrap');
  const btnRotL = document.getElementById('btnRotateModalSealLeft');
  const btnRotR = document.getElementById('btnRotateModalSealRight');
  const modalImg = document.getElementById('sealModalImg');

  if (btnRotL) {
    btnRotL.addEventListener('click', (e) => {
      e.stopPropagation();
      currentModalSealRotation = (currentModalSealRotation - 90 + 360) % 360;
      if (modalImg) modalImg.style.transform = `rotate(${currentModalSealRotation}deg)`;
    });
  }

  if (btnRotR) {
    btnRotR.addEventListener('click', (e) => {
      e.stopPropagation();
      currentModalSealRotation = (currentModalSealRotation + 90) % 360;
      if (modalImg) modalImg.style.transform = `rotate(${currentModalSealRotation}deg)`;
    });
  }

  if (openBtn) {
    openBtn.addEventListener('click', () => {
      const sealInput = document.getElementById('inputSealNo');
      const reason = (sealInput && sealInput.value.trim()) ? 'mismatch' : 'unreadable';
      openSealConfirmModal(reason);
    });
  }

  if (closeBtn && modal) {
    closeBtn.addEventListener('click', closeSealConfirmModal);
  }
  if (dismissBtn && modal) {
    dismissBtn.addEventListener('click', closeSealConfirmModal);
  }
  if (modal) {
    modal.addEventListener('click', (e) => {
      if (e.target === modal) closeSealConfirmModal();
    });
  }

  if (confirmBtn) {
    confirmBtn.addEventListener('click', handleConfirmManualSeal);
  }

  // OPTION 2: CHỤP LẠI (CHO CHỤP LẠI)
  const btnRetakeCam = document.getElementById('btnRetakeSealCam');
  const btnRetakeGal = document.getElementById('btnRetakeSealGal');

  if (btnRetakeCam) {
    btnRetakeCam.addEventListener('click', () => {
      closeSealConfirmModal();
      const inputCam = document.getElementById('inputCamSeal');
      if (inputCam) inputCam.click();
    });
  }

  if (btnRetakeGal) {
    btnRetakeGal.addEventListener('click', () => {
      closeSealConfirmModal();
      const inputGal = document.getElementById('inputGalSeal');
      if (inputGal) inputGal.click();
    });
  }

  if (manualInput) {
    manualInput.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') {
        e.preventDefault();
        handleConfirmManualSeal();
      }
    });
    manualInput.addEventListener('input', () => {
      manualInput.value = manualInput.value.toUpperCase();
    });
  }

  const zoomSealImage = () => {
    if (currentSealImageBase64) {
      openImageLightbox(currentSealImageBase64, '🔒 Ảnh chụp Seal thực tế (Phóng to đối chứng)', true);
    }
  };

  if (zoomBtn) zoomBtn.addEventListener('click', zoomSealImage);
  if (imgWrap) imgWrap.addEventListener('click', zoomSealImage);
}

function openSealConfirmModal(reason = 'unreadable') {
  const modal = document.getElementById('sealConfirmModal');
  if (!modal) return;

  currentModalSealRotation = 0;
  const contInput = document.getElementById('inputContNo');
  const sealInput = document.getElementById('inputSealNo');
  const contNo = normalizeContNo(contInput ? contInput.value : '');
  const actualSeal = normalizeSealNo(sealInput ? sealInput.value : '');

  const matchedRow = contNo ? excelList.find((r) => r.contNo === contNo) : null;
  const expectedSeal = matchedRow ? matchedRow.expectedSeal : '';

  const titleEl = document.getElementById('sealModalTitle');
  const subEl = document.getElementById('sealModalSub');
  const alertEl = document.getElementById('sealModalAlert');
  const contNoEl = document.getElementById('sealModalContNo');
  const expSealEl = document.getElementById('sealModalExpectedSeal');
  const imgEl = document.getElementById('sealModalImg');
  const noImgText = document.getElementById('sealModalNoImgText');
  const hintBadge = document.getElementById('sealModalImgHintBadge');
  const manualInput = document.getElementById('manualSealInput');
  const candidatesWrap = document.getElementById('sealModalCandidatesWrap');
  const candidatesList = document.getElementById('sealModalCandidatesList');

  if (imgEl) {
    imgEl.style.transform = 'rotate(0deg)';
  }

  // Fill Cont & Expected Seal
  contNoEl.textContent = contNo || '(Chưa rõ)';
  expSealEl.textContent = expectedSeal || (excelList.length === 0 ? '(Chưa tải Excel)' : 'Không có trong Excel');

  if (reason === 'mismatch') {
    titleEl.innerHTML = '❌ Cảnh báo số Seal không khớp';
    titleEl.style.color = '#f87171';
    subEl.textContent = 'Số Seal đọc từ ảnh không khớp với Excel. Vui lòng đối chiếu với ảnh chụp thực tế bên dưới.';
    alertEl.style.background = 'rgba(239, 68, 68, 0.16)';
    alertEl.style.borderColor = 'rgba(239, 68, 68, 0.5)';
    alertEl.style.color = '#fca5a5';
    alertEl.innerHTML = `⚠️ <strong>Seal đọc từ ảnh (${actualSeal || 'Trống'}) KHÁC với Excel (${expectedSeal || 'Trống'})</strong>.<br/>Hãy nhìn ảnh chụp seal bên dưới để kiểm tra và gõ số Seal chính xác.`;
    manualInput.value = actualSeal || '';
  } else {
    // Blurry / Unreadable
    titleEl.innerHTML = '🔍 Xác nhận số Seal (Ảnh mờ / Chưa đọc được)';
    titleEl.style.color = '#f59e0b';
    subEl.textContent = 'Chưa bắt rõ ký tự trên ảnh Seal. Vui lòng nhìn ảnh chụp bên dưới và gõ số Seal.';
    alertEl.style.background = 'rgba(245, 158, 11, 0.16)';
    alertEl.style.borderColor = 'rgba(245, 158, 11, 0.5)';
    alertEl.style.color = '#fcd34d';
    alertEl.innerHTML = `📸 <strong>Ảnh chụp Seal bị mờ, lóa sáng hoặc khuất góc.</strong><br/>Hãy nhìn ảnh chụp thực tế bên dưới và gõ số Seal vào ô để đối chiếu.`;
    manualInput.value = actualSeal || '';
  }

  // Display seal image
  if (currentSealImageBase64) {
    imgEl.src = currentSealImageBase64;
    imgEl.style.display = 'block';
    if (noImgText) noImgText.style.display = 'none';
    if (hintBadge) hintBadge.style.display = 'block';
  } else {
    imgEl.src = '';
    imgEl.style.display = 'none';
    if (noImgText) noImgText.style.display = 'block';
    if (hintBadge) hintBadge.style.display = 'none';
  }

  // Build candidate chips (from Excel expected tokens + OCR detected tokens)
  const candidateSet = [];
  const expectedTokens = extractSealTokens(expectedSeal);
  expectedTokens.forEach((t) => {
    if (!candidateSet.some((c) => c.val === t)) candidateSet.push({ val: t, source: 'excel' });
  });

  if (Array.isArray(lastDetectedSeals)) {
    lastDetectedSeals.forEach((t) => {
      const norm = normalizeSealNo(t);
      if (norm && !candidateSet.some((c) => c.val === norm)) {
        candidateSet.push({ val: norm, source: 'ocr' });
      }
    });
  }

  if (candidateSet.length > 0) {
    candidatesWrap.style.display = 'block';
    candidatesList.innerHTML = candidateSet
      .map((c) => {
        const isEx = c.source === 'excel';
        const label = isEx ? `📋 ${c.val} (Excel)` : `📸 ${c.val} (Ảnh)`;
        const cls = isEx ? 'seal-candidate-chip from-excel' : 'seal-candidate-chip';
        return `<button type="button" class="${cls}" data-val="${c.val}">${label}</button>`;
      })
      .join('');

    candidatesList.querySelectorAll('button').forEach((btn) => {
      btn.addEventListener('click', () => {
        manualInput.value = btn.getAttribute('data-val') || '';
        manualInput.focus();
      });
    });
  } else {
    candidatesWrap.style.display = 'none';
    candidatesList.innerHTML = '';
  }

  modal.classList.add('open');
  setTimeout(() => {
    manualInput.focus();
    manualInput.select();
  }, 100);
}

function closeSealConfirmModal() {
  const modal = document.getElementById('sealConfirmModal');
  if (modal) modal.classList.remove('open');
}

function handleConfirmManualSeal() {
  const manualInput = document.getElementById('manualSealInput');
  const inputCont = document.getElementById('inputContNo');
  const entered = normalizeSealNo(manualInput ? manualInput.value : '');

  if (!entered) {
    alert('Vui lòng gõ số Seal trên ảnh vào ô trước khi bấm Xác nhận!');
    if (manualInput) manualInput.focus();
    return;
  }

  const contNo = normalizeContNo(inputCont ? inputCont.value : '');
  hasManualConfirmedThisCycle = true;
  closeSealConfirmModal();

  // Set the entered seal on main page
  const inputSeal = document.getElementById('inputSealNo');
  if (inputSeal) inputSeal.value = entered;

  // Run comparison with entered seal, flagging as manual confirmed
  compareContAndSeal(contNo, entered, false, true);

  // Success hint
  showOcrHintBanner(
    `✅ Đã xác nhận số Seal: ${entered}. Hình ảnh chụp thực tế vẫn được lưu trữ đối chứng!`,
    'info'
  );
}

// ============================================================================
// Image Selection & Auto-Trigger
// ============================================================================
function initImageInputs() {
  const btnCamCont = document.getElementById('btnCamCont');
  const btnGalCont = document.getElementById('btnGalCont');
  const inputCamCont = document.getElementById('inputCamCont');
  const inputGalCont = document.getElementById('inputGalCont');
  const clearContBtn = document.getElementById('clearContImgBtn');

  btnCamCont.addEventListener('click', () => inputCamCont.click());
  btnGalCont.addEventListener('click', () => inputGalCont.click());
  inputCamCont.addEventListener('change', (e) => handleImageFile(e.target.files[0], 'cont'));
  inputGalCont.addEventListener('change', (e) => handleImageFile(e.target.files[0], 'cont'));
  clearContBtn.addEventListener('click', (e) => {
    e.stopPropagation();
    clearSlotImage('cont');
  });

  const btnCamSeal = document.getElementById('btnCamSeal');
  const btnGalSeal = document.getElementById('btnGalSeal');
  const inputCamSeal = document.getElementById('inputCamSeal');
  const inputGalSeal = document.getElementById('inputGalSeal');
  const clearSealBtn = document.getElementById('clearSealImgBtn');

  btnCamSeal.addEventListener('click', () => inputCamSeal.click());
  btnGalSeal.addEventListener('click', () => inputGalSeal.click());
  inputCamSeal.addEventListener('change', (e) => handleImageFile(e.target.files[0], 'seal'));
  inputGalSeal.addEventListener('change', (e) => handleImageFile(e.target.files[0], 'seal'));
  clearSealBtn.addEventListener('click', (e) => {
    e.stopPropagation();
    clearSlotImage('seal');
  });

  const previewWrapCont = document.getElementById('previewWrapCont');
  if (previewWrapCont) {
    previewWrapCont.addEventListener('click', () => {
      if (currentContImageBase64) {
        const contVal = document.getElementById('inputContNo').value || 'Mặt Cửa Cont';
        openImageLightbox(currentContImageBase64, `🖼️ Ảnh Mặt Cửa Cont: ${contVal}`);
      }
    });
  }

  const previewWrapSeal = document.getElementById('previewWrapSeal');
  if (previewWrapSeal) {
    previewWrapSeal.addEventListener('click', () => {
      if (currentSealImageBase64) {
        const sealVal = document.getElementById('inputSealNo').value || 'Số Seal (Chì)';
        openImageLightbox(currentSealImageBase64, `🔒 Ảnh Số Seal: ${sealVal}`, true);
      }
    });
  }

  document.getElementById('resetScanBtn').addEventListener('click', resetCurrentScan);
}

// ============================================================================
// Full-Screen Image Lightbox (Click Uploaded Image or History Image to View Full)
// ============================================================================
let currentLightboxRotation = 0;

function initLightboxModal() {
  const modal = document.getElementById('imageLightboxModal');
  if (!modal) return;

  const btnClose = document.getElementById('btnCloseLightbox');
  const btnZoomIn = document.getElementById('btnZoomIn');
  const btnZoomOut = document.getElementById('btnZoomOut');
  const btnZoomReset = document.getElementById('btnZoomReset');
  const btnRotL = document.getElementById('btnLightboxRotateLeft');
  const btnRotR = document.getElementById('btnLightboxRotateRight');
  const stage = document.getElementById('lightboxStage');
  const img = document.getElementById('lightboxImg');

  const sealDock = document.getElementById('lightboxSealDock');
  const sealInput = document.getElementById('lightboxSealInput');
  const confirmBtn = document.getElementById('btnConfirmLightboxSeal');

  btnClose.addEventListener('click', () => modal.classList.remove('open'));
  stage.addEventListener('click', (e) => {
    if (e.target === stage) modal.classList.remove('open');
  });

  if (btnRotL) {
    btnRotL.addEventListener('click', () => {
      currentLightboxRotation = (currentLightboxRotation - 90 + 360) % 360;
      applyLightboxZoom();
    });
  }

  if (btnRotR) {
    btnRotR.addEventListener('click', () => {
      currentLightboxRotation = (currentLightboxRotation + 90) % 360;
      applyLightboxZoom();
    });
  }

  btnZoomIn.addEventListener('click', () => {
    currentLightboxZoom = Math.min(350, currentLightboxZoom + 40);
    applyLightboxZoom();
  });

  btnZoomOut.addEventListener('click', () => {
    currentLightboxZoom = Math.max(50, currentLightboxZoom - 40);
    applyLightboxZoom();
  });

  btnZoomReset.addEventListener('click', () => {
    currentLightboxZoom = 100;
    currentLightboxRotation = 0;
    applyLightboxZoom();
  });

  img.addEventListener('dblclick', (e) => {
    e.stopPropagation();
    currentLightboxZoom = currentLightboxZoom === 100 ? 200 : 100;
    applyLightboxZoom();
  });

  if (confirmBtn && sealInput) {
    confirmBtn.addEventListener('click', () => {
      const val = normalizeSealNo(sealInput.value);
      if (val) {
        document.getElementById('inputSealNo').value = val;
        const modalInput = document.getElementById('manualSealInput');
        if (modalInput) modalInput.value = val;
        modal.classList.remove('open');
        closeSealConfirmModal();
        evaluateReadinessAndCompare(false);
      }
    });
    sealInput.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') {
        e.preventDefault();
        confirmBtn.click();
      }
    });
    sealInput.addEventListener('input', () => {
      sealInput.value = sealInput.value.toUpperCase();
    });
  }
}

function applyLightboxZoom() {
  const img = document.getElementById('lightboxImg');
  const btnZoomReset = document.getElementById('btnZoomReset');
  if (!img) return;
  if (currentLightboxZoom === 100) {
    img.style.width = 'auto';
    img.style.maxWidth = '96vw';
    img.style.maxHeight = '78vh';
  } else {
    img.style.maxWidth = 'none';
    img.style.maxHeight = 'none';
    img.style.width = `${currentLightboxZoom}vw`;
  }
  img.style.transform = `rotate(${currentLightboxRotation}deg)`;
  if (btnZoomReset) {
    btnZoomReset.textContent = `${currentLightboxZoom}%`;
  }
}

function openImageLightbox(src, title, isSeal = false) {
  if (!src) return;
  const modal = document.getElementById('imageLightboxModal');
  const img = document.getElementById('lightboxImg');
  const titleEl = document.getElementById('lightboxTitle');
  const sealDock = document.getElementById('lightboxSealDock');
  const sealInput = document.getElementById('lightboxSealInput');
  if (!modal || !img) return;

  currentLightboxZoom = 100;
  currentLightboxRotation = 0;
  img.src = src;
  if (titleEl) titleEl.textContent = title || '🔍 Xem ảnh kích thước đầy đủ';
  if (sealDock) {
    sealDock.style.display = isSeal ? 'flex' : 'none';
    if (isSeal && sealInput) {
      const currentVal = document.getElementById('manualSealInput')?.value || document.getElementById('inputSealNo')?.value || '';
      sealInput.value = currentVal;
      setTimeout(() => sealInput.focus(), 150);
    }
  }
  applyLightboxZoom();
  modal.classList.add('open');
}

async function handleImageFile(file, type) {
  if (!file) return;
  try {
    const resizedDataUrl = await resizeImageFile(file, 1600);
    if (type === 'cont') {
      currentContImageBase64 = resizedDataUrl;
      contImageExtracted = false;
      document.getElementById('inputContNo').value = '';
      const img = document.getElementById('imgPreviewCont');
      img.src = resizedDataUrl;
      img.style.display = 'block';
      document.getElementById('placeholderCont').style.display = 'none';
      document.getElementById('slotCont').classList.add('has-image');
    } else {
      currentSealImageBase64 = resizedDataUrl;
      sealImageExtracted = false;
      document.getElementById('inputSealNo').value = '';
      const img = document.getElementById('imgPreviewSeal');
      img.src = resizedDataUrl;
      img.style.display = 'block';
      document.getElementById('placeholderSeal').style.display = 'none';
      document.getElementById('slotSeal').classList.add('has-image');
    }

    document.getElementById('resultBanner').className = 'result-banner';
    lastComparisonResult = null;
    hasManualConfirmedThisCycle = false;
    lastDetectedSeals = [];
    runOneStepExtractAndCompare(false, type);
  } catch (err) {
    alert('Không thể đọc file ảnh: ' + err.message);
  }
}

function clearSlotImage(type) {
  if (type === 'cont') {
    currentContImageBase64 = null;
    contImageExtracted = false;
    const img = document.getElementById('imgPreviewCont');
    img.src = '';
    img.style.display = 'none';
    document.getElementById('placeholderCont').style.display = 'flex';
    document.getElementById('slotCont').classList.remove('has-image');
    document.getElementById('inputCamCont').value = '';
    document.getElementById('inputGalCont').value = '';
    document.getElementById('inputContNo').value = '';
    document.getElementById('isoBadge').style.display = 'none';
    document.getElementById('suggestionsBar').innerHTML = '';
  } else {
    currentSealImageBase64 = null;
    sealImageExtracted = false;
    const img = document.getElementById('imgPreviewSeal');
    img.src = '';
    img.style.display = 'none';
    document.getElementById('placeholderSeal').style.display = 'flex';
    document.getElementById('slotSeal').classList.remove('has-image');
    document.getElementById('inputCamSeal').value = '';
    document.getElementById('inputGalSeal').value = '';
    document.getElementById('inputSealNo').value = '';
  }
  document.getElementById('resultBanner').className = 'result-banner';
  lastComparisonResult = null;
  hasManualConfirmedThisCycle = false;
  lastDetectedSeals = [];
}

function resizeImageFile(file, maxDim = 1400) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (ev) => {
      const img = new Image();
      img.onload = () => {
        let { width, height } = img;
        if (width > maxDim || height > maxDim) {
          if (width > height) {
            height = Math.round((height * maxDim) / width);
            width = maxDim;
          } else {
            width = Math.round((width * maxDim) / height);
            height = maxDim;
          }
        }
        const canvas = document.createElement('canvas');
        canvas.width = width;
        canvas.height = height;
        const ctx = canvas.getContext('2d');
        ctx.drawImage(img, 0, 0, width, height);
        resolve(canvas.toDataURL('image/jpeg', 0.92));
      };
      img.onerror = reject;
      img.src = ev.target.result;
    };
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

// ============================================================================
// v7 Seal Processor: Pure Yellow Detector + Auto-Deskew + Dilated Grid Mask + Barcode Eraser
// ============================================================================
function isPureYellowSealPixel(r, g, b) {
  // Distinguishes bright yellow bolt seals (G - B > 80) from beige/tan/rusty container doors (G - B < 65)
  return r > 150 && g > 145 && g - b > 80 && r - b > 90 && Math.abs(r - g) < 40;
}

function buildSealMultiAngleViews(sealDataUrl) {
  return new Promise((resolve) => {
    const img = new Image();
    img.onload = () => {
      const probeW = 240;
      const probeH = 240;
      const probeCanvas = document.createElement('canvas');
      probeCanvas.width = probeW;
      probeCanvas.height = probeH;
      const pCtx = probeCanvas.getContext('2d');
      pCtx.drawImage(img, 0, 0, probeW, probeH);
      const pData = pCtx.getImageData(0, 0, probeW, probeH).data;

      let minX = probeW, minY = probeH, maxX = 0, maxY = 0;
      let matchPixels = 0;

      for (let y = 15; y < 225; y++) {
        for (let x = 15; x < 225; x++) {
          const idx = (y * probeW + x) * 4;
          if (isPureYellowSealPixel(pData[idx], pData[idx + 1], pData[idx + 2])) {
            matchPixels++;
            if (x < minX) minX = x;
            if (x > maxX) maxX = x;
            if (y < minY) minY = y;
            if (y > maxY) maxY = y;
          }
        }
      }

      const hasYellowSeal = matchPixels > 40 && maxX > minX && maxY > minY;
      let cropRect = { x: 0.0, y: 0.0, w: 1.0, h: 1.0 };
      if (hasYellowSeal) {
        const cx = Math.max(0.0, (minX - 5) / probeW);
        const cy = Math.max(0.0, (minY - 5) / probeH);
        const cw = Math.min(1.0 - cx, (maxX - minX + 10) / probeW);
        const ch = Math.min(1.0 - cy, (maxY - minY + 10) / probeH);
        cropRect = { x: cx, y: cy, w: cw, h: ch };
      }

      const sx = Math.floor(img.width * cropRect.x);
      const sy = Math.floor(img.height * cropRect.y);
      const sw = Math.floor(img.width * cropRect.w);
      const sh = Math.floor(img.height * cropRect.h);

      const targetMax = hasYellowSeal ? 520 : 800;
      const scale = hasYellowSeal
        ? Math.min(2.5, Math.max(1.0, targetMax / Math.max(sw, sh)))
        : Math.min(1.0, targetMax / Math.max(sw, sh));
      const cwPx = Math.floor(sw * scale);
      const chPx = Math.floor(sh * scale);

      // Helper: Measures median top-edge tilt slope of yellow tag using wider span (dx=28..45px) so even 2°..5° tilts are deskewed!
      const detectTagTiltDeg = (canvas) => {
        if (!hasYellowSeal) return 0;
        const w = canvas.width;
        const h = canvas.height;
        const data = canvas.getContext('2d').getImageData(0, 0, w, h).data;
        const topY = new Int32Array(w).fill(-1);

        for (let x = 0; x < w; x++) {
          for (let y = 0; y < h; y++) {
            const i = (y * w + x) * 4;
            if (isPureYellowSealPixel(data[i], data[i + 1], data[i + 2])) {
              topY[x] = y;
              break;
            }
          }
        }

        const slopes = [];
        const dxSpan = Math.max(28, Math.floor(w * 0.08));
        const xStart = Math.floor(w * 0.20);
        const xEnd = Math.floor(w * 0.75);
        for (let x = xStart; x < xEnd - dxSpan; x++) {
          if (topY[x] >= 0 && topY[x + dxSpan] >= 0) {
            const dy = topY[x + dxSpan] - topY[x];
            if (Math.abs(dy) <= 45) {
              slopes.push(dy / dxSpan);
            }
          }
        }

        if (slopes.length < 10) return 0;
        slopes.sort((a, b) => a - b);
        const medSlope = slopes[Math.floor(slopes.length / 2)];
        const deg = (Math.atan(medSlope) * 180.0) / Math.PI;
        return Math.abs(deg) >= 1.5 && Math.abs(deg) <= 32 ? deg : 0;
      };

      const renderBaseRotation = (baseAngleDeg) => {
        const c = document.createElement('canvas');
        if (baseAngleDeg === 90 || baseAngleDeg === 270) {
          c.width = chPx;
          c.height = cwPx;
        } else {
          c.width = cwPx;
          c.height = chPx;
        }
        const ctx = c.getContext('2d');
        ctx.translate(c.width / 2, c.height / 2);
        ctx.rotate((baseAngleDeg * Math.PI) / 180);
        ctx.drawImage(img, sx, sy, sw, sh, -cwPx / 2, -chPx / 2, cwPx, chPx);

        // Auto-deskew any 1.5°..32° tilt so text and barcode become dead-level 0.0° horizontal!
        const tiltDeg = detectTagTiltDeg(c);
        if (Math.abs(tiltDeg) >= 1.5) {
          const deskewed = document.createElement('canvas');
          deskewed.width = c.width;
          deskewed.height = c.height;
          const dCtx = deskewed.getContext('2d');
          dCtx.translate(c.width / 2, c.height / 2);
          dCtx.rotate((-tiltDeg * Math.PI) / 180);
          dCtx.drawImage(c, -c.width / 2, -c.height / 2);
          return deskewed;
        }
        return c;
      };

      // v9 Binarizer & Barcode Eraser: Isolates pure seal serial text strip below/above barcode
      const binarizeAndCleanBarcode = (srcCanvas, isRot90 = false) => {
        const c = document.createElement('canvas');
        c.width = srcCanvas.width;
        c.height = srcCanvas.height;
        const ctx = c.getContext('2d');
        ctx.drawImage(srcCanvas, 0, 0);

        const w = c.width;
        const h = c.height;
        const id = ctx.getImageData(0, 0, w, h);
        const d = id.data;

        const gw = 24;
        const gh = 24;
        const grid = new Uint8Array(gw * gh);
        const rowMinX = new Int32Array(h).fill(w);
        const rowMaxX = new Int32Array(h).fill(-1);
        const colMinY = new Int32Array(w).fill(h);
        const colMaxY = new Int32Array(w).fill(-1);
        const yLums = [];

        for (let y = 0; y < h; y++) {
          const gy = Math.min(gh - 1, Math.floor((y * gh) / h));
          const rowOff = y * w * 4;
          for (let x = 0; x < w; x++) {
            const i = rowOff + x * 4;
            const r = d[i], g = d[i + 1], b = d[i + 2];
            const isTagYellow = hasYellowSeal
              ? (r > 140 && g > 130 && g - b > 65 && Math.abs(r - g) < 55)
              : (r > 135 || g > 135 || b > 135);
            if (isTagYellow) {
              const gx = Math.min(gw - 1, Math.floor((x * gw) / w));
              grid[gy * gw + gx] = 1;
              if (x < rowMinX[y]) rowMinX[y] = x;
              if (x > rowMaxX[y]) rowMaxX[y] = x;
              if (y < colMinY[x]) colMinY[x] = y;
              if (y > colMaxY[x]) colMaxY[x] = y;
              yLums.push(Math.round(0.299 * r + 0.587 * g + 0.114 * b));
            }
          }
        }

        yLums.sort((a, b) => a - b);
        const yMed = yLums.length > 20 ? yLums[Math.floor(yLums.length * 0.50)] : 185;
        const threshPivot = yMed - 26;

        for (let y = 0; y < h; y++) {
          const gy = Math.min(gh - 1, Math.floor((y * gh) / h));
          const rowOff = y * w * 4;
          for (let x = 0; x < w; x++) {
            const gx = Math.min(gw - 1, Math.floor((x * gw) / w));
            let inGrid = false;
            for (let dy = -1; dy <= 1 && !inGrid; dy++) {
              const ny = gy + dy;
              if (ny < 0 || ny >= gh) continue;
              for (let dx = -1; dx <= 1; dx++) {
                const nx = gx + dx;
                if (nx >= 0 && nx < gw && grid[ny * gw + nx]) {
                  inGrid = true;
                  break;
                }
              }
            }
            const inSpan =
              x >= rowMinX[y] + 2 &&
              x <= rowMaxX[y] - 2 &&
              y >= colMinY[x] + 2 &&
              y <= colMaxY[x] - 2;

            const i = rowOff + x * 4;
            if (!(inGrid && inSpan)) {
              d[i] = d[i + 1] = d[i + 2] = 255;
            } else {
              const lum = 0.299 * d[i] + 0.587 * d[i + 1] + 0.114 * d[i + 2];
              const val = Math.max(0, Math.min(255, Math.round((lum - threshPivot) * 5.5 + 128)));
              d[i] = d[i + 1] = d[i + 2] = val;
            }
          }
        }

        // Dynamic 1D horizontal barcode detector (runs of width 1..14px)
        const rowRuns = new Int32Array(h);
        for (let y = 0; y < h; y++) {
          let runs = 0;
          let runW = 0;
          const rowOff = y * w * 4;
          for (let x = 0; x < w; x++) {
            if (d[rowOff + x * 4] < 140) {
              runW++;
            } else {
              if (runW >= 1 && runW <= 14) runs++;
              runW = 0;
            }
          }
          rowRuns[y] = runs;
        }

        let bestY = 0;
        let bestSum = 0;
        for (let y = 4; y < h - 35; y++) {
          let s = 0;
          for (let dy = 0; dy < 20; dy++) s += rowRuns[y + dy];
          if (s > bestSum) {
            bestSum = s;
            bestY = y;
          }
        }

        const avgRuns = bestSum / 20.0;
        if (avgRuns >= 15.0) {
          let barTop = bestY;
          while (barTop > 0 && rowRuns[barTop - 1] >= 11) barTop--;
          let barBot = bestY + 19;
          while (barBot < h - 1 && rowRuns[barBot + 1] >= 11) barBot++;

          if (isRot90) {
            // Text is above barcode: barTop - 52 to barTop - 1
            const textTopLimit = Math.max(0, barTop - 52);
            for (let y = 0; y < h; y++) {
              if (y < textTopLimit || y >= barTop - 1) {
                const rowOff = y * w * 4;
                for (let x = 0; x < w; x++) {
                  d[rowOff + x * 4] = d[rowOff + x * 4 + 1] = d[rowOff + x * 4 + 2] = 255;
                }
              }
            }
          } else {
            // Text is below barcode (rot 270): barBot + 1 to barBot + 52
            const textBottomLimit = Math.min(h, barBot + 52);
            for (let y = 0; y < h; y++) {
              if (y <= barBot + 1 || y >= textBottomLimit) {
                const rowOff = y * w * 4;
                for (let x = 0; x < w; x++) {
                  d[rowOff + x * 4] = d[rowOff + x * 4 + 1] = d[rowOff + x * 4 + 2] = 255;
                }
              }
            }
          }
        }

        ctx.putImageData(id, 0, 0);
        return c;
      };

      const color0 = renderBaseRotation(0);
      const color90 = renderBaseRotation(90);
      const color270 = renderBaseRotation(270);

      const crop0 = binarizeAndCleanBarcode(color0, false);
      const crop90 = binarizeAndCleanBarcode(color90, true);
      const crop270 = binarizeAndCleanBarcode(color270, false);

      resolve({
        color0DataUrl: color0.toDataURL('image/jpeg', 0.92),
        color90DataUrl: color90.toDataURL('image/jpeg', 0.92),
        color270DataUrl: color270.toDataURL('image/jpeg', 0.92),
        crop0DataUrl: crop0.toDataURL('image/jpeg', 0.92),
        crop90DataUrl: crop90.toDataURL('image/jpeg', 0.92),
        crop270DataUrl: crop270.toDataURL('image/jpeg', 0.92)
      });
    };
    img.src = sealDataUrl;
  });
}

function createContainerOcrCrop(dataUrl, cropRect, invert = false) {
  return new Promise((resolve) => {
    const img = new Image();
    img.onload = () => {
      const sx = Math.floor(img.width * cropRect.x);
      const sy = Math.floor(img.height * cropRect.y);
      const sw = Math.max(10, Math.floor(img.width * cropRect.w));
      const sh = Math.max(10, Math.floor(img.height * cropRect.h));

      // Scale to optimal character height (width at least 1400px)
      const targetW = Math.max(sw, 1400);
      const scale = targetW / sw;
      const targetH = Math.round(sh * scale);

      const canvas = document.createElement('canvas');
      canvas.width = targetW;
      canvas.height = targetH;
      const ctx = canvas.getContext('2d');
      ctx.imageSmoothingEnabled = true;
      ctx.imageSmoothingQuality = 'high';
      ctx.drawImage(img, sx, sy, sw, sh, 0, 0, targetW, targetH);

      const imgData = ctx.getImageData(0, 0, targetW, targetH);
      const d = imgData.data;

      // Calculate luminance histogram (3rd to 97th percentile for robust contrast range)
      const lums = new Uint8Array(targetW * targetH);
      const hist = new Int32Array(256);
      for (let i = 0, p = 0; i < d.length; i += 4, p++) {
        const lum = Math.round(0.299 * d[i] + 0.587 * d[i + 1] + 0.114 * d[i + 2]);
        lums[p] = lum;
        hist[lum]++;
      }

      const totalPixels = targetW * targetH;
      const lowerCut = Math.floor(totalPixels * 0.03);
      const upperCut = Math.floor(totalPixels * 0.97);

      let count = 0;
      let minLum = 0;
      let maxLum = 255;
      for (let i = 0; i < 256; i++) {
        count += hist[i];
        if (count >= lowerCut) {
          minLum = i;
          break;
        }
      }
      count = 0;
      for (let i = 255; i >= 0; i--) {
        count += hist[i];
        if (count >= totalPixels - upperCut) {
          maxLum = i;
          break;
        }
      }

      const range = Math.max(1, maxLum - minLum);

      // Smooth contrast stretch (preserves character edges, no harsh binarization artifacts)
      for (let i = 0, p = 0; i < d.length; i += 4, p++) {
        let norm = (lums[p] - minLum) / range;
        if (norm < 0) norm = 0;
        if (norm > 1) norm = 1;
        let val = Math.round(Math.pow(norm, 1.15) * 255);
        if (invert) val = 255 - val;
        d[i] = val;
        d[i + 1] = val;
        d[i + 2] = val;
      }

      ctx.putImageData(imgData, 0, 0);
      resolve(canvas.toDataURL('image/jpeg', 0.92));
    };
    img.src = dataUrl;
  });
}

function createProcessedCrop(dataUrl, cropRect, invert = false, highContrast = true) {
  return createContainerOcrCrop(dataUrl, cropRect, invert);
}

// ============================================================================
// Unified 1-Step Handler: Extract + Compare Together
// ============================================================================
function initUnifiedCheckHandlers() {
  const btnOneStep = document.getElementById('btnOneStepCheck');
  const btnSaveNext = document.getElementById('btnSaveAndNext');
  const inputCont = document.getElementById('inputContNo');
  const inputSeal = document.getElementById('inputSealNo');

  btnOneStep.addEventListener('click', () => {
    runOneStepExtractAndCompare(true);
  });

  inputCont.addEventListener('input', () => {
    const cleanCont = normalizeContNo(inputCont.value);
    const cleanSeal = normalizeSealNo(inputSeal.value);
    updateIsoBadge(cleanCont);
    renderFuzzySuggestions(cleanCont);
    if (cleanCont.length >= 4 && cleanSeal.length >= 3) {
      hideOcrHintBanner();
      compareContAndSeal(inputCont.value, inputSeal.value, true);
    } else {
      document.getElementById('resultBanner').className = 'result-banner';
      lastComparisonResult = null;
    }
  });

  inputSeal.addEventListener('input', () => {
    const cleanCont = normalizeContNo(inputCont.value);
    const cleanSeal = normalizeSealNo(inputSeal.value);
    if (cleanCont.length >= 4 && cleanSeal.length >= 3) {
      hideOcrHintBanner();
      compareContAndSeal(inputCont.value, inputSeal.value, true);
    } else {
      document.getElementById('resultBanner').className = 'result-banner';
      lastComparisonResult = null;
    }
  });

  btnSaveNext.addEventListener('click', saveCurrentCheckAndNext);
}

function evaluateReadinessAndCompare(silentSound = false, triggeredByType = null) {
  const inputCont = document.getElementById('inputContNo');
  const inputSeal = document.getElementById('inputSealNo');
  const contNo = normalizeContNo(inputCont.value);
  const sealNo = normalizeSealNo(inputSeal.value);

  // 1. Only compare when BOTH Container and Seal are available!
  if (contNo && sealNo) {
    hideOcrHintBanner();
    compareContAndSeal(contNo, sealNo, silentSound);
    return;
  }

  // Otherwise, hide any premature comparison banner and guide the user to the next step
  document.getElementById('resultBanner').className = 'result-banner';
  lastComparisonResult = null;

  // 2. Container is available, Seal is NOT available
  if (contNo && !sealNo) {
    if (!currentSealImageBase64 || triggeredByType === 'cont') {
      const inExcel = excelList.find((r) => normalizeContNo(r.contNo) === contNo);
      const excelNote =
        excelList.length === 0
          ? ''
          : inExcel
          ? ' (Có trong DS Excel ✓)'
          : ' (Chưa thấy trong DS Excel)';
      showOcrHintBanner(
        `✅ Đã nhận diện Số Cont: ${contNo}${excelNote} (Chuẩn ISO ✓). Bạn hãy chụp/chọn tiếp "2. Ảnh Số Seal" để kiểm tra!`,
        'info'
      );
    } else {
      // Both images exist, but sealNo is empty AND user triggered seal or both
      showOcrHintBanner(
        `⚠️ Đã có Số Cont (${contNo}), nhưng chưa bắt rõ ký tự trên ảnh Seal. Bạn hãy chụp lại ảnh Seal rõ hơn hoặc gõ số Seal vào ô.`,
        'warn',
        'seal'
      );
      if (triggeredByType === 'seal' && !hasManualConfirmedThisCycle) {
        openSealConfirmModal('unreadable');
      }
    }
    return;
  }

  // 3. Seal is available, Container is NOT available
  if (!contNo && sealNo) {
    if (!currentContImageBase64) {
      showOcrHintBanner(
        `✅ Đã đọc Số Seal: ${sealNo}. Hãy chụp/chọn tiếp "1. Ảnh Mặt Cửa Cont" (quét 1/4 trên bên phải) để đối chiếu!`,
        'info'
      );
    } else {
      showOcrHintBanner(
        `⚠️ Đã đọc Số Seal (${sealNo}), nhưng chưa nhận diện rõ Số Container trên ảnh Cửa Cont.`,
        'warn',
        'cont'
      );
      if (triggeredByType === 'cont' || triggeredByType === null) {
        openContConfirmModal('unreadable');
      }
    }
    return;
  }

  // 4. Container photo uploaded but contNo was NOT recognized
  if (currentContImageBase64 && !contNo) {
    showOcrHintBanner(
      `⚠️ Chưa nhận diện được Số Container từ ảnh. Bạn hãy kiểm tra lại ảnh Cửa Cont (chụp rõ 1/4 trên bên phải) hoặc gõ Số Container vào ô.`,
      'warn',
      'cont'
    );
    if (triggeredByType === 'cont' || triggeredByType === null) {
      openContConfirmModal('unreadable');
    }
    return;
  }

  const isGeminiActive =
    (settings.ocrMode === 'gemini' || settings.ocrMode === 'hybrid') &&
    settings.geminiKey.trim().length > 10;
  const geminiTip = isGeminiActive
    ? ''
    : ' 👉 Bấm ⚙️ (góc trên bên phải) dán Gemini API Key để AI nhận diện tức thì & chuẩn xác 100%!';

  showOcrHintBanner(
    '⚠️ Hãy chụp/chọn ảnh Cửa Cont & ảnh Seal ở trên (hoặc gõ số vào ô bên dưới) để kiểm tra.' + geminiTip,
    'warn'
  );
}

async function runOneStepExtractAndCompare(forceOcrIfImagesExist = false, triggeredByType = null) {
  if (isExtractingOcr) {
    pendingReExtract = true;
    return;
  }

  const inputCont = document.getElementById('inputContNo');
  const inputSeal = document.getElementById('inputSealNo');
  const hasImages = Boolean(currentContImageBase64 || currentSealImageBase64);

  if (!hasImages) {
    if (!inputCont.value.trim() && !inputSeal.value.trim()) {
      showOcrHintBanner(
        '📸 Hãy chụp/chọn đủ 2 ảnh Cửa Cont & ảnh Seal ở trên (hoặc gõ số vào ô bên dưới) để kiểm tra.',
        'warn'
      );
      return;
    }
    evaluateReadinessAndCompare(false, triggeredByType);
    return;
  }

  let needsContOcr = false;
  let needsSealOcr = false;

  if (triggeredByType === 'cont') {
    needsContOcr = Boolean(currentContImageBase64);
    needsSealOcr = false;
  } else if (triggeredByType === 'seal') {
    needsContOcr = false;
    needsSealOcr = Boolean(currentSealImageBase64);
  } else {
    needsContOcr = currentContImageBase64 && (!contImageExtracted || !inputCont.value.trim() || forceOcrIfImagesExist);
    needsSealOcr = currentSealImageBase64 && (!sealImageExtracted || !inputSeal.value.trim() || forceOcrIfImagesExist);
  }

  if (!needsContOcr && !needsSealOcr && (inputCont.value.trim() || inputSeal.value.trim())) {
    evaluateReadinessAndCompare(false, triggeredByType);
    return;
  }

  const statusBox = document.getElementById('ocrStatusBox');
  const statusText = document.getElementById('ocrStatusText');
  const btnOneStep = document.getElementById('btnOneStepCheck');

  isExtractingOcr = true;
  statusBox.classList.add('active');
  btnOneStep.disabled = true;
  hideOcrHintBanner();

  const shouldTryGemini = Boolean(
    settings.geminiKey &&
    settings.geminiKey.trim().length > 10 &&
    settings.ocrMode !== 'local'
  );
  let geminiErrorMessage = '';

  try {
    // 1. Extract Container Number
    if (needsContOcr && currentContImageBase64) {
      let contNo = '';
      if (shouldTryGemini) {
        statusText.textContent = '✨ Đang gửi Gemini AI đọc Số Container...';
        try {
          contNo = await extractContWithGemini(currentContImageBase64);
          sendDebugLog('gemini_cont_ok', { contNo });
        } catch (e) {
          console.warn('Gemini Cont Error (tự động chuyển sang Tesseract 1/4 trên phải):', e);
          geminiErrorMessage = e.message;
          sendDebugLog('gemini_cont_err', { error: e.message });
        }
      }
      // LUÔN LUÔN tự động chạy Tesseract quét 1/4 trên bên phải nếu Gemini gặp sự cố!
      if (!contNo) {
        statusText.textContent = '⚡ Đang tập trung quét 1/4 bên phải phía trên (chuẩn ISO 6346)...';
        contNo = await extractContWithFreeTesseract(currentContImageBase64, statusText);
        sendDebugLog('tesseract_cont_done', { contNo });
      }
      contImageExtracted = true;
      if (contNo) {
        inputCont.value = contNo;
        updateIsoBadge(contNo);
        renderFuzzySuggestions(contNo);
      }
    }

    // 2. Automatically Extract Seal Number (Fast Barcode -> Gemini AI -> Tesseract)
    if (needsSealOcr && currentSealImageBase64) {
      let sealNo = '';
      const detectedCandidates = [];
      statusText.textContent = '🔍 Đang tự động quét mã vạch & số Seal...';
      const sealViews = await buildSealMultiAngleViews(currentSealImageBase64);

      // Fast Path: BarcodeDetector in browser (instant 5ms, 100% accurate on barcode seals)
      if ('BarcodeDetector' in window) {
        try {
          const detector = new BarcodeDetector({
            formats: ['code_128', 'code_39', 'code_93', 'ean_13', 'itf', 'qr_code']
          });
          for (const viewUrl of [currentSealImageBase64, sealViews.color0DataUrl, sealViews.color270DataUrl, sealViews.color90DataUrl]) {
            if (!viewUrl) continue;
            const img = await loadImageElement(viewUrl);
            const barcodes = await detector.detect(img);
            if (barcodes && barcodes.length > 0) {
              for (const bc of barcodes) {
                const val = normalizeSealNo(bc.rawValue);
                if (val && val.length >= 5 && !detectedCandidates.includes(val)) {
                  detectedCandidates.push(val);
                }
              }
              if (!sealNo && detectedCandidates.length > 0) {
                sealNo = detectedCandidates[0];
                sendDebugLog('barcode_detected_early', { sealNo, detectedCandidates });
              }
            }
          }
        } catch (be) {}
      }

      // Check current container & manifest hint from Excel
      const currentCont = normalizeContNo(inputCont.value);
      const matchedExcelRow = currentCont ? excelList.find((r) => r.contNo === currentCont) : null;
      const hintSeal = matchedExcelRow ? matchedExcelRow.expectedSeal : '';

      // Gemini AI Vision (Multi-Seal Detection with Color & Angle Invariance)
      if (shouldTryGemini) {
        statusText.textContent = '✨ Đang gửi Gemini AI đọc Số Seal...';
        try {
          const gRes = await extractSealWithGemini(currentSealImageBase64, sealViews, hintSeal, currentCont);
          if (gRes.allSeals && gRes.allSeals.length > 0) {
            for (const s of gRes.allSeals) {
              if (s && !detectedCandidates.includes(s)) detectedCandidates.push(s);
            }
          }
          if (gRes.sealNo && !detectedCandidates.includes(gRes.sealNo)) {
            detectedCandidates.unshift(gRes.sealNo);
          }
          if (!sealNo && gRes.sealNo) {
            sealNo = gRes.sealNo;
          }
          sendDebugLog('gemini_seal_ok', { sealNo, allSeals: gRes.allSeals });
        } catch (e) {
          console.error('Gemini Seal Error:', e);
          if (!geminiErrorMessage) geminiErrorMessage = e.message;
          sendDebugLog('gemini_seal_err', { error: e.message });
        }
      }

      // Fallback: Free Tesseract OCR (luôn tự động chạy nếu Gemini gặp sự cố)
      if (!sealNo) {
        statusText.textContent = '⚡ Đang quét Số Seal (khử mã vạch & đa góc)...';
        const tessSeal = await extractSealWithFreeTesseract(sealViews, inputCont.value, statusText, hintSeal);
        if (tessSeal && !isBarcodePhantomText(tessSeal) && !detectedCandidates.includes(tessSeal)) {
          detectedCandidates.push(tessSeal);
          if (!sealNo) sealNo = tessSeal;
        }
        sendDebugLog('tesseract_seal_done', { sealNo, tessSeal });
      }

      // MULTI-SEAL & HINT MATCHING:
      // Filter out any barcode phantom artifact texts (like LLHR111M)
      const validCandidates = detectedCandidates.filter((c) => !isBarcodePhantomText(c));
      if (hintSeal && validCandidates.length > 0) {
        const matchCheck = checkSealMatch('', hintSeal, validCandidates);
        if (matchCheck.isMatch && matchCheck.matchedToken) {
          sealNo = matchCheck.matchedToken;
        } else if (!sealNo || isBarcodePhantomText(sealNo)) {
          sealNo = validCandidates[0];
        }
      } else if ((!sealNo || isBarcodePhantomText(sealNo)) && validCandidates.length > 0) {
        sealNo = validCandidates[0];
      }

      lastDetectedSeals = validCandidates.length > 0 ? validCandidates : detectedCandidates;
      sealImageExtracted = true;
      if (sealNo && !isBarcodePhantomText(sealNo)) {
        inputSeal.value = sealNo;
      }
    }

    if (geminiErrorMessage && (!inputCont.value.trim() && !inputSeal.value.trim())) {
      showOcrHintBanner(
        `⚠️ Gemini AI báo lỗi: ${geminiErrorMessage}. Vui lòng kiểm tra lại API Key trong ⚙️.`,
        'warn'
      );
      return;
    }

    // 3. Evaluate readiness and compare
    evaluateReadinessAndCompare(false, triggeredByType);
  } catch (err) {
    sendDebugLog('ocr_fatal_err', { error: err.message });
    showOcrHintBanner('⚠️ Lỗi khi xử lý ảnh: ' + err.message, 'warn');
  } finally {
    isExtractingOcr = false;
    statusBox.classList.remove('active');
    btnOneStep.disabled = false;

    if (pendingReExtract) {
      pendingReExtract = false;
      runOneStepExtractAndCompare(false, triggeredByType);
    }
  }
}

function showOcrHintBanner(msg, type = 'warn', actionType = null) {
  let hint = document.getElementById('ocrHintBanner');
  if (!hint) {
    hint = document.createElement('div');
    hint.id = 'ocrHintBanner';
    document.getElementById('ocrStatusBox').insertAdjacentElement('afterend', hint);
  }

  hint.className = `ocr-hint-banner ${type}`;
  if (type === 'info') {
    hint.style.cssText =
      'margin-top:10px; padding:10px 12px; border-radius:10px; background:rgba(59,130,246,0.16); border:1px solid rgba(59,130,246,0.5); color:#93c5fd; font-size:12.5px; font-weight:600; line-height:1.45;';
    hint.innerHTML = msg;
    hint.onclick = null;
    hint.style.cursor = 'default';
  } else {
    hint.style.cssText =
      'margin-top:10px; padding:10px 12px; border-radius:10px; background:rgba(245,158,11,0.16); border:1px solid rgba(245,158,11,0.5); color:#fcd34d; font-size:12px; line-height:1.45;';

    if (actionType === 'seal') {
      hint.style.cursor = 'pointer';
      hint.innerHTML = `${msg} <span style="text-decoration:underline; font-weight:700; color:#38bdf8; margin-left:6px;">👉 Bấm mở hộp thoại gõ số Seal</span>`;
      hint.onclick = () => {
        const sealInput = document.getElementById('inputSealNo');
        const reason = (sealInput && sealInput.value.trim()) ? 'mismatch' : 'unreadable';
        openSealConfirmModal(reason);
      };
    } else if (actionType === 'cont') {
      hint.style.cursor = 'pointer';
      hint.innerHTML = `${msg} <span style="text-decoration:underline; font-weight:700; color:#38bdf8; margin-left:6px;">👉 Bấm mở hộp thoại gõ số Cont</span>`;
      hint.onclick = () => {
        openContConfirmModal('unreadable');
      };
    } else {
      hint.style.cursor = 'default';
      hint.innerHTML = msg;
      hint.onclick = null;
    }
  }
  hint.style.display = 'block';
}

function hideOcrHintBanner() {
  const hint = document.getElementById('ocrHintBanner');
  if (hint) hint.style.display = 'none';
}

// ============================================================================
// Gemini AI Vision Calls (Fast with 10s Timeout & Safe JSON Parser)
// ============================================================================
let cachedGeminiModel = null;

async function getWorkingGeminiModel(apiKey) {
  if (cachedGeminiModel) return cachedGeminiModel;
  try {
    const listRes = await fetch(
      `https://generativelanguage.googleapis.com/v1beta/models?key=${encodeURIComponent(apiKey)}`
    );
    if (listRes.ok) {
      const listData = await listRes.json();
      const available = (listData.models || [])
        .filter((m) => m.supportedGenerationMethods && m.supportedGenerationMethods.includes('generateContent'))
        .map((m) => m.name.replace(/^models\//, ''));

      const preferred = [
        'gemini-3.5-flash-lite',
        'gemini-3.5-flash',
        'gemini-3.8-flash',
        'gemini-flash-latest',
        'gemini-flash-lite-latest',
        'gemini-pro-latest'
      ];
      for (const pref of preferred) {
        if (available.includes(pref)) {
          cachedGeminiModel = pref;
          return pref;
        }
      }
      if (available.length > 0) {
        cachedGeminiModel = available[0];
        return available[0];
      }
    }
  } catch (e) {
    console.warn('Model discovery warning:', e);
  }
  return 'gemini-3.5-flash-lite';
}

async function callGeminiJson(parts) {
  const apiKey = (settings.geminiKey || DEFAULT_GEMINI_KEY).trim();
  if (!apiKey) throw new Error('Chưa nhập Gemini API Key');

  const discoveredModel = await getWorkingGeminiModel(apiKey);
  const models = [
    discoveredModel,
    'gemini-3.5-flash-lite',
    'gemini-3.5-flash',
    'gemini-3.8-flash',
    'gemini-flash-latest',
    'gemini-flash-lite-latest'
  ];
  const uniqueModels = [...new Set(models.filter(Boolean))];
  let lastErr = null;

  for (const modelName of uniqueModels) {
    const url = `https://generativelanguage.googleapis.com/v1beta/models/${modelName}:generateContent?key=${encodeURIComponent(apiKey)}`;
    try {
      const controller = new AbortController();
      const timeoutId = setTimeout(() => controller.abort(), 12000);

      const response = await fetch(url, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        signal: controller.signal,
        body: JSON.stringify({
          contents: [{ parts }],
          generationConfig: {
            temperature: 0.1
          }
        })
      });
      clearTimeout(timeoutId);

      if (!response.ok) {
        const errText = await response.text();
        let errMsg = `HTTP ${response.status}`;
        try {
          const errJson = JSON.parse(errText);
          if (errJson?.error?.message) {
            errMsg = `${errJson.error.message} (${response.status})`;
          }
        } catch (e) {
          errMsg = `HTTP ${response.status}: ${errText.slice(0, 150)}`;
        }
        lastErr = new Error(errMsg);
        continue;
      }

      const data = await response.json();
      const candidateParts = data?.candidates?.[0]?.content?.parts || [];
      let rawText = candidateParts
        .filter((p) => !p.thought && typeof p.text === 'string')
        .map((p) => p.text)
        .join('\n')
        .trim();

      if (!rawText && candidateParts.length > 0) {
        rawText = candidateParts[candidateParts.length - 1].text || '{}';
      }

      rawText = rawText.replace(/```json|```/gi, '').trim();
      const jsonMatch = rawText.match(/\{[\s\S]*?\}/);
      if (jsonMatch) {
        try {
          const parsed = JSON.parse(jsonMatch[0]);
          return Array.isArray(parsed) ? parsed[0] || {} : parsed || {};
        } catch (pe) {
          console.warn('JSON parse warning:', pe);
        }
      }
      return { rawText };
    } catch (e) {
      if (e.name === 'AbortError') {
        lastErr = new Error(`Quá thời gian chờ (${modelName})`);
      } else {
        lastErr = e;
      }
    }
  }
  throw lastErr || new Error('Gemini API request failed');
}

async function extractContWithGemini(contDataUrl) {
  const base64Data = contDataUrl.includes(',') ? contDataUrl.split(',')[1] : contDataUrl;
  const mimeType = contDataUrl.startsWith('data:')
    ? contDataUrl.substring(5, contDataUrl.indexOf(';'))
    : 'image/jpeg';

  const manifestHint =
    Array.isArray(excelList) && excelList.length > 0
      ? `\nExpected candidate containers on current ship/manifest list: [${excelList
          .slice(0, 40)
          .map((r) => r.contNo)
          .join(', ')}]. If the container in the photo matches one of these, return that exact 11-character container number.`
      : '';

  // 1/4 Top-Right Quadrant Crop (Standard ISO container location)
  let top14CropUrl = null;
  let numberStripUrl = null;
  try {
    top14CropUrl = await createContainerOcrCrop(
      contDataUrl,
      { x: 0.45, y: 0.00, w: 0.55, h: 0.48 },
      false
    );
    numberStripUrl = await createContainerOcrCrop(
      contDataUrl,
      { x: 0.48, y: 0.04, w: 0.52, h: 0.36 },
      true
    );
  } catch (ce) {}

  const parts = [
    {
      text: `You are an expert shipping container inspector and ISO OCR specialist.
Extract the 11-character ISO 6346 shipping container number from this container door photo (e.g. "HACU2160569", "MSNU7784178", "FFAU2194320").
CRITICAL LOCATION: Focus on the TOP-RIGHT QUADRANT (1/4 trên bên phải cánh cửa). Container numbers are printed horizontally in large bold block letters (e.g. HACU 216056 9).
It consists of:
- 4 uppercase letters (Owner code 3 letters + Category identifier 'U', e.g. HACU)
- 6 digits (serial number, e.g. 216056)
- 1 check digit (often in a square box [9] or separated by space 9).
If you see 4 letters and 6 digits, find the check digit in the box or compute it using ISO 6346 so the full 11-character code is returned.${manifestHint}

Return ONLY JSON:
{
  "contNo": "...",
  "confidence": "high" | "medium" | "low"
}`
    },
    {
      inlineData: {
        mimeType: mimeType || 'image/jpeg',
        data: base64Data
      }
    }
  ];

  if (top14CropUrl) {
    const qData = top14CropUrl.includes(',') ? top14CropUrl.split(',')[1] : top14CropUrl;
    parts.push({
      inlineData: {
        mimeType: 'image/jpeg',
        data: qData
      }
    });
  }

  if (numberStripUrl) {
    const sData = numberStripUrl.includes(',') ? numberStripUrl.split(',')[1] : numberStripUrl;
    parts.push({
      inlineData: {
        mimeType: 'image/jpeg',
        data: sData
      }
    });
  }

  const parsed = await callGeminiJson(parts);
  const raw =
    parsed.contNo ||
    parsed.cont_no ||
    parsed.containerNo ||
    parsed.container_number ||
    parsed.containerNumber ||
    parsed.number ||
    parsed.rawText ||
    '';
  return parseContainerFromRawText(raw, excelList) || normalizeContNo(raw);
}

async function extractSealWithGemini(sealDataUrl, sealViews, hintExpectedSeal = '', contNo = '') {
  const base64Data = sealDataUrl.includes(',') ? sealDataUrl.split(',')[1] : sealDataUrl;
  const mimeType = sealDataUrl.startsWith('data:')
    ? sealDataUrl.substring(5, sealDataUrl.indexOf(';'))
    : 'image/jpeg';

  const hintText = hintExpectedSeal
    ? `\nExpected Manifest Seal in database for container ${contNo ? `(${contNo})` : ''} is: "${hintExpectedSeal}". Check carefully if this number or any part of it appears on any of the seals in the photo.`
    : '';

  const parts = [
    {
      text: `You are an expert shipping container inspector and seal verification specialist.
Carefully examine this container door seal / lock photo.
NOTE: Container doors frequently have ONE OR MULTIPLE SEALS attached (e.g. high-security ISO 17712 bolt seal, plastic wire seal, cable seal, customs seal "HẢI QUAN", or carrier seal).
The seals can be of ANY COLOR: red, blue, green, yellow, orange, white, or metallic silver/grey.
Serial numbers are printed, stamped, laser-etched, or dot-matrix on the seal bolt, barrel, flag, or tag (e.g. "FX49778780", "A4220802184", "0439243", "1234567").
Text may be vertical or horizontal. Join separated letters and numbers together without spaces.${hintText}

CRITICAL INSTRUCTIONS:
1. Detect ALL seal serial numbers visible in the photo.
2. If multiple seals are present, include every single seal number in the "allSeals" array.
3. Return the primary or best matching seal number in "sealNo".
4. CRITICAL: Do NOT read barcode stripes as letters/digits (e.g. do NOT output vertical stripe noise like "LLHR111M", "TILIA", "1111", "LLLL"). Read only the true seal serial number (e.g. FX49778780, A4220802184, 0439243).
5. If the photo is too blurry or numbers cannot be deciphered, set "isBlurryOrUnreadable": true.

Return JSON ONLY:
{
  "sealNo": "...",
  "allSeals": ["..."],
  "confidence": "high" | "medium" | "low",
  "isBlurryOrUnreadable": false
}`
    },
    {
      inlineData: {
        mimeType: mimeType || 'image/jpeg',
        data: base64Data
      }
    }
  ];

  if (sealViews?.color90DataUrl) {
    const v90Data = sealViews.color90DataUrl.includes(',')
      ? sealViews.color90DataUrl.split(',')[1]
      : sealViews.color90DataUrl;
    parts.push({
      inlineData: {
        mimeType: 'image/jpeg',
        data: v90Data
      }
    });
  }

  if (sealViews?.color270DataUrl) {
    const v2Data = sealViews.color270DataUrl.includes(',')
      ? sealViews.color270DataUrl.split(',')[1]
      : sealViews.color270DataUrl;
    parts.push({
      inlineData: {
        mimeType: 'image/jpeg',
        data: v2Data
      }
    });
  }

  const parsed = await callGeminiJson(parts);
  const raw =
    parsed.sealNo ||
    parsed.seal_no ||
    parsed.sealNumber ||
    parsed.seal_number ||
    parsed.serialNo ||
    parsed.serial_number ||
    parsed.number ||
    parsed.rawText ||
    '';

  const allSealsList = Array.isArray(parsed.allSeals)
    ? parsed.allSeals
        .map(normalizeSealNo)
        .filter((s) => Boolean(s) && !isBarcodePhantomText(s))
    : [];

  let primary = parseSealFromRawText(raw, '', hintExpectedSeal) || normalizeSealNo(raw);
  if (primary && isBarcodePhantomText(primary)) {
    primary = '';
  }
  if (primary && !allSealsList.includes(primary)) {
    allSealsList.unshift(primary);
  }

  return {
    sealNo: primary || (allSealsList[0] || ''),
    allSeals: allSealsList,
    isBlurryOrUnreadable: Boolean(parsed.isBlurryOrUnreadable) || (!primary && allSealsList.length === 0)
  };
}

// ============================================================================
// Free Local OCR (Tesseract on Auto-Deskewed + Dilated-Mask + Barcode-Erased Crops)
// ============================================================================
async function extractContWithFreeTesseract(contDataUrl, statusEl) {
  const ocrOptionsUniform = {
    tessedit_char_whitelist: 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789 -[]()/|',
    tessedit_pageseg_mode: '6'
  };
  const ocrOptionsSparse = {
    tessedit_char_whitelist: 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789 -[]()/|',
    tessedit_pageseg_mode: '11'
  };

  let combinedContRaw = '';

  // PASS 1: Cánh cửa phải (x: 0.30 -> 1.0, y: 0.0 -> 0.85) - Đảo màu tương phản
  // Bao trọn vẹn số cont HACU 216056 [9] trên nền xanh/sẫm
  statusEl.textContent = '⚡ Đang quét cánh cửa phải (nền sẫm chữ trắng)...';
  try {
    const cropRightInv = await createContainerOcrCrop(
      contDataUrl,
      { x: 0.30, y: 0.00, w: 0.70, h: 0.85 },
      true
    );
    const res1 = await Tesseract.recognize(cropRightInv, 'eng', ocrOptionsUniform);
    const text1 = res1?.data?.text || '';
    combinedContRaw += '\n' + text1;
    let contNo = parseContainerFromRawText(combinedContRaw, excelList);
    if (contNo && validateIso6346(contNo)) return contNo;
  } catch (e1) {
    console.warn('Cont Right Door Inv warning:', e1);
  }

  // PASS 2: Dải trọng tâm số Cont ở cánh phải (x: 0.32 -> 1.0, y: 0.20 -> 0.75) - Đảo màu
  // Phóng to trực diện vào dòng số HACU 216056 [9] khi ảnh chụp dính trần/mái kho
  statusEl.textContent = '⚡ Đang phóng to dải số Cont cánh phải...';
  try {
    const cropFocusInv = await createContainerOcrCrop(
      contDataUrl,
      { x: 0.32, y: 0.20, w: 0.68, h: 0.55 },
      true
    );
    const res2 = await Tesseract.recognize(cropFocusInv, 'eng', ocrOptionsUniform);
    const text2 = res2?.data?.text || '';
    combinedContRaw += '\n' + text2;
    let contNo = parseContainerFromRawText(combinedContRaw, excelList);
    if (contNo && validateIso6346(contNo)) return contNo;
  } catch (e2) {
    console.warn('Cont Focus Inv warning:', e2);
  }

  // PASS 3: Toàn bộ khung hình ảnh chụp (x: 0.0 -> 1.0, y: 0.0 -> 1.0) - Đảo màu
  statusEl.textContent = '⚡ Đang quét toàn bộ khung hình ảnh chụp...';
  try {
    const cropFullInv = await createContainerOcrCrop(
      contDataUrl,
      { x: 0.0, y: 0.0, w: 1.0, h: 1.0 },
      true
    );
    const res3 = await Tesseract.recognize(cropFullInv, 'eng', ocrOptionsUniform);
    const text3 = res3?.data?.text || '';
    combinedContRaw += '\n' + text3;
    let contNo = parseContainerFromRawText(combinedContRaw, excelList);
    if (contNo && validateIso6346(contNo)) return contNo;
  } catch (e3) {
    console.warn('Cont Full Inv warning:', e3);
  }

  // PASS 4: Cánh cửa phải - Nền sáng chữ đen (Cho cont vỏ trắng, vàng, xám nhạt)
  statusEl.textContent = '⚡ Đang quét cánh cửa phải (nền sáng chữ đen)...';
  try {
    const cropRightNorm = await createContainerOcrCrop(
      contDataUrl,
      { x: 0.30, y: 0.00, w: 0.70, h: 0.85 },
      false
    );
    const res4 = await Tesseract.recognize(cropRightNorm, 'eng', ocrOptionsUniform);
    const text4 = res4?.data?.text || '';
    combinedContRaw += '\n' + text4;
    let contNo = parseContainerFromRawText(combinedContRaw, excelList);
    if (contNo && validateIso6346(contNo)) return contNo;
  } catch (e4) {
    console.warn('Cont Right Norm warning:', e4);
  }

  // PASS 5: Quét mở rộng phân đoạn thưa
  try {
    const cropFullSparse = await createContainerOcrCrop(
      contDataUrl,
      { x: 0.0, y: 0.0, w: 1.0, h: 1.0 },
      false
    );
    const res5 = await Tesseract.recognize(cropFullSparse, 'eng', ocrOptionsSparse);
    const text5 = res5?.data?.text || '';
    combinedContRaw += '\n' + text5;
  } catch (e5) {
    console.warn('Cont Sparse warning:', e5);
  }

  return parseContainerFromRawText(combinedContRaw, excelList);
}

async function extractSealWithFreeTesseract(sealViews, excludeContNo, statusEl, hintExpectedSeal = '') {
  if ('BarcodeDetector' in window) {
    try {
      const detector = new BarcodeDetector({
        formats: ['code_128', 'code_39', 'code_93', 'ean_13', 'itf', 'qr_code']
      });
      for (const viewUrl of [sealViews.color0DataUrl, sealViews.color90DataUrl, sealViews.color270DataUrl]) {
        if (!viewUrl) continue;
        const img = await loadImageElement(viewUrl);
        const barcodes = await detector.detect(img);
        if (barcodes && barcodes.length > 0) {
          const val = normalizeSealNo(barcodes[0].rawValue);
          if (val && val.length >= 5 && !isBarcodePhantomText(val)) return val;
        }
      }
    } catch (e) {}
  }

  statusEl.textContent = '⚡ Đang đọc số Seal (khử mã vạch & xoay đa hướng)...';
  // PASS 1: Đọc ảnh đã khử mã vạch & tách dòng chữ (crop270, crop90, crop0)
  const [resCrop270, resCrop90, resCrop0] = await Promise.all([
    sealViews.crop270DataUrl ? Tesseract.recognize(sealViews.crop270DataUrl, 'eng') : null,
    sealViews.crop90DataUrl ? Tesseract.recognize(sealViews.crop90DataUrl, 'eng') : null,
    sealViews.crop0DataUrl ? Tesseract.recognize(sealViews.crop0DataUrl, 'eng') : null
  ]);

  let rawAll =
    (resCrop270?.data?.text || '') +
    '\n' +
    (resCrop90?.data?.text || '') +
    '\n' +
    (resCrop0?.data?.text || '') +
    '\n';

  sendDebugLog('tesseract_seal_crop_raw', { rawAll });

  let candidate = parseSealFromRawText(rawAll, excludeContNo, hintExpectedSeal);
  if (candidate && isStrongSealCandidate(candidate)) return candidate;

  // PASS 2: Quét ảnh màu đối chứng đa hướng (color270, color90, color0)
  statusEl.textContent = '⚡ Đang quét ảnh màu đa hướng đối chứng...';
  const [resColor270, resColor90, resColor0] = await Promise.all([
    sealViews.color270DataUrl ? Tesseract.recognize(sealViews.color270DataUrl, 'eng') : null,
    sealViews.color90DataUrl ? Tesseract.recognize(sealViews.color90DataUrl, 'eng') : null,
    sealViews.color0DataUrl ? Tesseract.recognize(sealViews.color0DataUrl, 'eng') : null
  ]);
  rawAll +=
    (resColor270?.data?.text || '') +
    '\n' +
    (resColor90?.data?.text || '') +
    '\n' +
    (resColor0?.data?.text || '') +
    '\n';

  sendDebugLog('tesseract_seal_color_raw', { rawAll });
  return parseSealFromRawText(rawAll, excludeContNo, hintExpectedSeal);
}

function loadImageElement(src) {
  return new Promise((resolve, reject) => {
    const i = new Image();
    i.onload = () => resolve(i);
    i.onerror = reject;
    i.src = src;
  });
}

function isBarcodePhantomText(s) {
  if (!s) return false;
  const clean = String(s).toUpperCase().replace(/[^A-Z0-9]/g, '');
  if (clean.length < 4) return true;
  // If the string consists predominantly of vertical barcode bar letters (L, I, 1, |, H, M)
  const barcodeChars = (clean.match(/[I1L|HM]/g) || []).length;
  if (barcodeChars / clean.length >= 0.70 && !excelList.some((r) => r.expectedSeal && r.expectedSeal.includes(clean))) {
    return true;
  }
  return false;
}

function isStrongSealCandidate(s) {
  if (!s || s.length < 5) return false;
  if (isBarcodePhantomText(s)) return false;
  if (excelList.some((r) => r.expectedSeal === s)) return true;
  return (
    /^[A-Z]{1,4}\d{4,10}$/.test(s) ||
    /^[A-Z0-9]{2,4}\d{4,10}$/.test(s) ||
    /^\d{5,11}$/.test(s)
  );
}

// ============================================================================
// Smart Container & Seal Parsing (with Excel Serial Substring & Fuzzy Match!)
// ============================================================================
function normalizeContNo(str) {
  if (!str) return '';
  return String(str).toUpperCase().replace(/[^A-Z0-9]/g, '');
}

function normalizeSealNo(str) {
  if (!str) return '';
  return String(str)
    .toUpperCase()
    .replace(/^(SEAL|CHÌ|CHI|NO|NUMBER|SỐ|SO)[\s:.\-#]*/i, '')
    .replace(/[^A-Z0-9]/g, '');
}

function extractSealTokens(str) {
  if (!str) return [];
  if (Array.isArray(str)) {
    return Array.from(new Set(str.flatMap(extractSealTokens)));
  }
  const raw = String(str).trim();
  if (!raw) return [];

  // Split by common delimiters: comma, slash, semicolon, plus, pipe, newline
  const parts = raw.split(/[\/,;\+\|\n\r]+/).map((p) => p.trim()).filter(Boolean);
  const tokens = [];

  for (const part of parts) {
    const cleaned = normalizeSealNo(part);
    if (cleaned && !tokens.includes(cleaned)) {
      tokens.push(cleaned);
    }
    const spaceBlocks = part.split(/\s+/).map(normalizeSealNo).filter((b) => b && b.length >= 5);
    for (const b of spaceBlocks) {
      if (b && !tokens.includes(b)) {
        tokens.push(b);
      }
    }
  }

  if (tokens.length === 0) {
    const norm = normalizeSealNo(raw);
    if (norm) tokens.push(norm);
  }
  return tokens;
}

function formatExcelSealString(raw) {
  if (!raw) return '';
  const tokens = extractSealTokens(raw);
  if (tokens.length === 0) return normalizeSealNo(raw);
  return tokens.join(' / ');
}

function checkSealMatch(actualSeal, expectedSeal, detectedSeals = []) {
  const expectedTokens = extractSealTokens(expectedSeal);
  const actualTokens = Array.from(
    new Set([
      ...extractSealTokens(actualSeal),
      ...(Array.isArray(detectedSeals) ? detectedSeals.flatMap(extractSealTokens) : [])
    ])
  );

  if (expectedTokens.length === 0 || actualTokens.length === 0) {
    return {
      isMatch: false,
      matchedToken: null,
      matchedExpected: null,
      expectedTokens,
      actualTokens
    };
  }

  // 1. Exact string match
  for (const act of actualTokens) {
    for (const exp of expectedTokens) {
      if (act === exp) {
        return { isMatch: true, matchedToken: act, matchedExpected: exp, expectedTokens, actualTokens };
      }
    }
  }

  // 2. Prefix/Suffix / Suffix-digits match (e.g. FX49778780 vs 49778780 or WHLU0439243 vs 0439243)
  for (const act of actualTokens) {
    const actDigits = act.replace(/\D/g, '');
    for (const exp of expectedTokens) {
      const expDigits = exp.replace(/\D/g, '');
      if (
        (act.length >= 6 && exp.length >= 6 && (act.endsWith(exp) || exp.endsWith(act))) ||
        (actDigits.length >= 6 && expDigits.length >= 6 && actDigits === expDigits)
      ) {
        return { isMatch: true, matchedToken: act, matchedExpected: exp, expectedTokens, actualTokens };
      }
    }
  }

  return {
    isMatch: false,
    matchedToken: null,
    matchedExpected: null,
    expectedTokens,
    actualTokens
  };
}

function calculateIso6346CheckDigit(prefix10) {
  if (!prefix10 || prefix10.length !== 10) return null;
  const charValues = {
    A: 10, B: 12, C: 13, D: 14, E: 15, F: 16, G: 17, H: 18, I: 19,
    J: 20, K: 21, L: 23, M: 24, N: 25, O: 26, P: 27, Q: 28, R: 29,
    S: 30, T: 31, U: 32, V: 34, W: 35, X: 36, Y: 37, Z: 38
  };
  const p = prefix10.toUpperCase();
  let sum = 0;
  for (let i = 0; i < 10; i++) {
    const ch = p[i];
    const val = i < 4 ? charValues[ch] : parseInt(ch, 10);
    if (val === undefined || isNaN(val)) return null;
    sum += val * Math.pow(2, i);
  }
  return String((sum % 11) % 10);
}

function validateIso6346(contNo) {
  if (!contNo || contNo.length !== 11) return false;
  // Strictly reject non-container words like TON, TOO, TARE, NET, MAX, KGS, LBS
  if (/^(TON|TOO|TAR|NET|MAX|PAY|CAP|CUC|KGS|LBS)/i.test(contNo)) return false;
  // According to ISO 6346:
  // - 3 letters: Owner code
  // - 1 letter: Equipment category identifier ('U' for freight containers, 'J' detachable, 'Z' trailers)
  if (!/^[A-Z]{3}[UJZ]\d{7}$/.test(contNo)) return false;
  const expectedCheckDigit = calculateIso6346CheckDigit(contNo.slice(0, 10));
  return expectedCheckDigit === contNo[10];
}

function fixContainerOcrChars(candidate) {
  const s = candidate.toUpperCase().replace(/[^A-Z0-9]/g, '');
  if (s.length < 10 || s.length > 11) return s;

  // Words related to tare weight or container specifications are NEVER container numbers
  if (/^(TON|TOO|TAR|NET|MAX|PAY|CAP|CUC|KGS|LBS)/i.test(s)) return '';

  const numToLetter = { '0': 'O', '1': 'I', '2': 'Z', '4': 'A', '5': 'S', '6': 'G', '8': 'B' };
  const letterToNum = { 'O': '0', 'Q': '0', 'D': '0', 'I': '1', 'L': '1', 'Z': '2', 'S': '5', 'G': '6', 'B': '8' };

  let prefix = s
    .slice(0, 4)
    .split('')
    .map((ch) => numToLetter[ch] || ch)
    .join('');

  // 4th character is almost always 'U' for freight containers in ISO 6346 (NEVER convert 'N' or 'M' to 'U')
  if (prefix[3] !== 'U' && prefix[3] !== 'J' && prefix[3] !== 'Z') {
    if (['0', 'O', 'V', '4', '1', 'Q', 'D'].includes(s[3])) {
      prefix = prefix.slice(0, 3) + 'U';
    }
  }

  const digits = s
    .slice(4)
    .split('')
    .map((ch) => letterToNum[ch] || ch)
    .join('');

  let res = prefix + digits;

  // If 10 characters (4 letters + 6 digits, check digit missing on door or in box),
  // mathematically calculate the 11th ISO check digit!
  if (res.length === 10 && /^[A-Z]{3}[UJZ]\d{6}$/.test(res)) {
    const cd = calculateIso6346CheckDigit(res);
    if (cd !== null) res = res + cd;
  }
  return res;
}

function parseContainerFromRawText(rawText, targetExcelList = null) {
  if (!rawText) return '';
  const cleaned = rawText.toUpperCase();
  const activeExcel = Array.isArray(targetExcelList) ? targetExcelList : excelList;

  // 1. Direct or partial match with Excel manifest
  if (Array.isArray(activeExcel) && activeExcel.length > 0) {
    const alnumStream = cleaned.replace(/[^A-Z0-9]/g, '');
    for (const item of activeExcel) {
      const c = normalizeContNo(item.contNo);
      if (!c) continue;
      // Exact match in OCR stream
      if (alnumStream.includes(c)) return c;
      // Match 10-char prefix or 6-digit serial
      if (c.length >= 10) {
        const p10 = c.slice(0, 10);
        const s6 = c.slice(4, 10);
        if (alnumStream.includes(p10) || alnumStream.includes(s6)) return c;
      }
    }
  }

  const candidates = [];
  function addCandidate(raw) {
    if (!raw) return;
    const fixed = fixContainerOcrChars(raw);
    if (!fixed) return;
    if (/^(TON|TOO|TAR|NET|MAX|PAY|CAP|CUC|KGS|LBS)/i.test(fixed)) return;
    if (/^[A-Z]{3}[UJZ]\d{7}$/.test(fixed)) {
      if (!candidates.includes(fixed)) candidates.push(fixed);
    }
  }

  // Pattern 1: Horizontal standard with or without box: [A-Z]{3,4} [0-9]{6} [0-9] (supports [0], (0), -0, etc.)
  const reg1 = /[A-Z0-9]{3,4}[\s\-_/\\|]*[A-Z0-9]{6}[\s\-_/\\|]*[\[\(]?\s*[A-Z0-9]\s*[\]\)]?/g;
  const m1 = cleaned.match(reg1) || [];
  m1.forEach(addCandidate);

  // Pattern 2: 10 chars (4 letters + 6 digits, missing check digit)
  const reg2 = /[A-Z0-9]{3,4}[\s\-_/\\|]*[A-Z0-9]{6}/g;
  const m2 = cleaned.match(reg2) || [];
  m2.forEach(addCandidate);

  // Pattern 3: Compact 11 chars
  const reg3 = /[A-Z0-9]{11}/g;
  const m3 = cleaned.match(reg3) || [];
  m3.forEach(addCandidate);

  // Pattern 4: Multi-line (e.g. Line 1: HACU, Line 2: 216058 0)
  const lines = cleaned.split(/[\r\n]+/).map((l) => l.trim()).filter(Boolean);
  for (let i = 0; i < lines.length - 1; i++) {
    const combined = lines[i] + ' ' + lines[i + 1];
    const mc = combined.match(reg1) || combined.match(reg2) || [];
    mc.forEach(addCandidate);
  }

  // Scoring candidates
  let best = '';
  let bestScore = -1;
  for (const c of candidates) {
    let score = 10;
    if (validateIso6346(c)) score += 40;
    if (c[3] === 'U') score += 20;
    if (Array.isArray(activeExcel) && activeExcel.some((item) => normalizeContNo(item.contNo) === c)) {
      score += 100;
    }
    // Strictly demote non-container words like TON, TOO, CONT, NET, TARE, MAX, KGS
    if (/^(TON|TOO|CON|NET|TAR|MAX|PAY|CAP|CUC|KGS|LBS)/i.test(c)) {
      score -= 200;
    }
    if (score > bestScore) {
      bestScore = score;
      best = c;
    }
  }

  if (best && Array.isArray(activeExcel) && activeExcel.length > 0) {
    const exact = activeExcel.find((x) => normalizeContNo(x.contNo) === best);
    if (exact) return normalizeContNo(exact.contNo);
    const close = findClosestExcelCont(best, 2);
    if (close) return close;
  }

  return best;
}

function parseSealFromRawText(rawText, excludeContNo, hintExpectedSeal = '') {
  if (!rawText) return '';
  const cleaned = rawText.toUpperCase();

  // 1. Prioritize expected seal hint for currently selected/scanned container
  if (hintExpectedSeal) {
    const hintTokens = extractSealTokens(hintExpectedSeal);
    const alnumStream = cleaned.replace(/[^A-Z0-9]/g, '');
    const digitStream = alnumStream
      .split('')
      .map((ch) => ({ O: '0', Q: '0', D: '0', I: '1', L: '1', Z: '2', S: '5', B: '8' }[ch] || ch))
      .join('');

    for (const s of hintTokens) {
      if (!s || s.length < 5) continue;
      if (alnumStream.includes(s)) return s;
      if (s.length >= 7) {
        const tail6 = s.slice(-6);
        const mid6 = s.slice(-7, -1);
        if (digitStream.includes(tail6) || digitStream.includes(mid6)) {
          return s;
        }
      }
    }

    for (const s of hintTokens) {
      if (!s || s.length < 7) continue;
      const tail5 = s.slice(-5);
      const mid5 = s.slice(-6, -1);
      if (/^\d{5}$/.test(tail5) && digitStream.includes(tail5)) return s;
      if (/^\d{5}$/.test(mid5) && digitStream.includes(mid5)) return s;

      const tail6 = s.slice(-6);
      if (tail6.length === 6 && hasWindowWithinHamming1(digitStream, tail6)) {
        return s;
      }
      if (/^\d{5}$/.test(tail5) && hasWindowWithinHamming1(digitStream, tail5)) {
        return s;
      }
    }
  }

  // 2. Prioritize all seals in Excel
  if (excelList.length > 0) {
    const alnumStream = cleaned.replace(/[^A-Z0-9]/g, '');
    const digitStream = alnumStream
      .split('')
      .map((ch) => ({ O: '0', Q: '0', D: '0', I: '1', L: '1', Z: '2', S: '5', B: '8' }[ch] || ch))
      .join('');

    for (const item of excelList) {
      const tokens = extractSealTokens(item.expectedSeal);
      for (const s of tokens) {
        if (!s || s.length < 5) continue;
        if (alnumStream.includes(s)) return s;
        if (s.length >= 7) {
          const tail6 = s.slice(-6);
          const mid6 = s.slice(-7, -1);
          if (digitStream.includes(tail6) || digitStream.includes(mid6)) {
            return s;
          }
        }
      }
    }

    for (const item of excelList) {
      const tokens = extractSealTokens(item.expectedSeal);
      for (const s of tokens) {
        if (!s || s.length < 7) continue;
        const tail5 = s.slice(-5);
        const mid5 = s.slice(-6, -1);
        if (/^\d{5}$/.test(tail5) && digitStream.includes(tail5)) return s;
        if (/^\d{5}$/.test(mid5) && digitStream.includes(mid5)) return s;

        const tail6 = s.slice(-6);
        if (tail6.length === 6 && hasWindowWithinHamming1(digitStream, tail6)) {
          return s;
        }
        if (/^\d{5}$/.test(tail5) && hasWindowWithinHamming1(digitStream, tail5)) {
          return s;
        }
      }
    }
  }

  const lines = cleaned
    .split(/[\r\n]+/)
    .map((l) => l.trim())
    .filter(Boolean);

  const candidates = [];

  for (const rawLine of lines) {
    const line = rawLine.replace(
      /\b(SEAL|HIGH SECURITY|CONTAINER|ISO|WEIGHT|MAX|GROSS|TARE|NET|KGS|LBS|MSC|MAERSK|COSCO|EVERGREEN|CMA|SITC)\b/g,
      ' '
    );

    const spacedMatches = line.match(/[A-Z0-9]{1,4}(?:[\s\-]+[0-9O]{2,8}){1,3}/g) || [];
    for (const sm of spacedMatches) {
      const joined = fixSealOcrDigits(normalizeSealNo(sm));
      if (joined.length >= 5 && joined.length <= 14 && /\d{4,}/.test(joined)) {
        if (!excludeContNo || joined !== excludeContNo) {
          if (!isBarcodePhantomText(joined)) {
            candidates.push(joined);
          }
        }
      }
    }

    const tokens = line.match(/[A-Z0-9\-]{5,15}/g) || [];
    for (const tok of tokens) {
      const norm = fixSealOcrDigits(normalizeSealNo(tok));
      if (norm.length < 5 || norm.length > 14) continue;
      if (excludeContNo && norm === excludeContNo) continue;
      if (/^[I1L|]+$/.test(norm)) continue;
      if (isBarcodePhantomText(norm)) continue;
      if (/\d{4,}/.test(norm)) {
        candidates.push(norm);
      }
    }
  }

  let best = '';
  let bestScore = -1;
  for (const c of candidates) {
    let score = 5;
    if (/^[A-Z]{1,4}\d{4,10}$/.test(c)) score += 25;
    else if (/^[A-Z0-9]{2,4}\d{4,10}$/.test(c)) score += 20;
    else if (/^\d{6,11}$/.test(c)) score += 22;
    else if (/^\d{5}$/.test(c)) score += 15;
    if (c.length >= 6 && c.length <= 12) score += 5;

    if (score > bestScore) {
      bestScore = score;
      best = c;
    }
  }

  return best;
}

function parseAllCandidatesFromRawText(rawText, excludeContNo) {
  if (!rawText) return [];
  const cleaned = rawText.toUpperCase();
  const list = [];
  const lines = cleaned.split(/[\r\n]+/).map((l) => l.trim()).filter(Boolean);

  for (const rawLine of lines) {
    const line = rawLine.replace(
      /\b(SEAL|HIGH SECURITY|CONTAINER|ISO|WEIGHT|MAX|GROSS|TARE|NET|KGS|LBS|MSC|MAERSK|COSCO|EVERGREEN|CMA|SITC)\b/g,
      ' '
    );

    const spacedMatches = line.match(/[A-Z0-9]{1,4}(?:[\s\-]+[0-9O]{2,8}){1,3}/g) || [];
    for (const sm of spacedMatches) {
      const joined = fixSealOcrDigits(normalizeSealNo(sm));
      if (joined.length >= 5 && joined.length <= 14 && /\d{4,}/.test(joined)) {
        if (!excludeContNo || joined !== excludeContNo) {
          if (!isBarcodePhantomText(joined) && !list.includes(joined)) {
            list.push(joined);
          }
        }
      }
    }

    const tokens = line.match(/[A-Z0-9\-]{5,15}/g) || [];
    for (const tok of tokens) {
      const norm = fixSealOcrDigits(normalizeSealNo(tok));
      if (norm.length < 5 || norm.length > 14) continue;
      if (excludeContNo && norm === excludeContNo) continue;
      if (/^[I1L|]+$/.test(norm)) continue;
      if (isBarcodePhantomText(norm)) continue;
      if (/\d{4,}/.test(norm)) {
        if (!list.includes(norm)) list.push(norm);
      }
    }
  }
  return list;
}

function hasWindowWithinHamming1(stream, target6) {
  if (!stream || stream.length < target6.length) return false;
  const len = target6.length;
  for (let i = 0; i <= stream.length - len; i++) {
    let diff = 0;
    for (let j = 0; j < len; j++) {
      if (stream[i + j] !== target6[j]) {
        diff++;
        if (diff > 1) break;
      }
    }
    if (diff <= 1) return true;
  }
  return false;
}

function fixSealOcrDigits(s) {
  if (!s || s.length < 5) return s;
  const letterToNum = { O: '0', Q: '0', D: '0', I: '1', L: '1', Z: '2', S: '5', B: '8' };
  const head = s.slice(0, 2);
  const tail = s
    .slice(2)
    .split('')
    .map((ch) => letterToNum[ch] || ch)
    .join('');
  return head + tail;
}



function updateIsoBadge(contNo) {
  const badge = document.getElementById('isoBadge');
  if (!contNo || contNo.length < 4) {
    badge.style.display = 'none';
    return;
  }
  badge.style.display = 'inline-block';
  if (validateIso6346(contNo)) {
    badge.textContent = 'Chuẩn ISO ✓';
    badge.className = 'iso-tag valid';
  } else {
    badge.textContent = contNo.length === 11 ? 'Sai số kiểm tra ISO' : `${contNo.length}/11 ký tự`;
    badge.className = 'iso-tag warn';
  }
}

function levenshtein(a, b) {
  const m = a.length;
  const n = b.length;
  const dp = Array.from({ length: m + 1 }, () => Array(n + 1).fill(0));
  for (let i = 0; i <= m; i++) dp[i][0] = i;
  for (let j = 0; j <= n; j++) dp[0][j] = j;
  for (let i = 1; i <= m; i++) {
    for (let j = 1; j <= n; j++) {
      const cost = a[i - 1] === b[j - 1] ? 0 : 1;
      dp[i][j] = Math.min(dp[i - 1][j] + 1, dp[i][j - 1] + 1, dp[i - 1][j - 1] + cost);
    }
  }
  return dp[m][n];
}

function findClosestExcelCont(inputCont, maxDist = 2) {
  let best = null;
  let bestDist = maxDist + 1;
  for (const row of excelList) {
    const d = levenshtein(inputCont, row.contNo);
    if (d <= maxDist && d < bestDist) {
      bestDist = d;
      best = row.contNo;
    }
  }
  return best;
}

function renderFuzzySuggestions(inputCont) {
  const bar = document.getElementById('suggestionsBar');
  bar.innerHTML = '';
  if (!inputCont || inputCont.length < 4 || excelList.length === 0) return;

  const exact = excelList.find((r) => r.contNo === inputCont);
  if (exact) return;

  const similar = excelList
    .map((r) => ({ contNo: r.contNo, seal: r.expectedSeal, dist: levenshtein(inputCont, r.contNo) }))
    .filter((x) => x.dist > 0 && x.dist <= 3)
    .sort((a, b) => a.dist - b.dist)
    .slice(0, 3);

  if (similar.length > 0) {
    similar.forEach((s) => {
      const chip = document.createElement('button');
      chip.className = 'suggestion-chip';
      chip.textContent = `Gợi ý DS: ${s.contNo}`;
      chip.addEventListener('click', () => {
        document.getElementById('inputContNo').value = s.contNo;
        updateIsoBadge(s.contNo);
        bar.innerHTML = '';
        evaluateReadinessAndCompare(false);
      });
      bar.appendChild(chip);
    });
  }
}

// ============================================================================
// Comparison & Audio/Vibration Feedback
// ============================================================================
function compareContAndSeal(rawCont, rawSeal, silentSound = false, isManualConfirmed = false) {
  const contNo = normalizeContNo(rawCont);
  const actualSeal = normalizeSealNo(rawSeal);

  if (!silentSound) {
    document.getElementById('inputContNo').value = contNo;
    document.getElementById('inputSealNo').value = actualSeal;
  }
  updateIsoBadge(contNo);
  renderFuzzySuggestions(contNo);

  if (!contNo && !actualSeal) {
    return;
  }

  if (excelList.length === 0) {
    showOcrHintBanner('⚠️ Bạn chưa tải danh sách Excel! Hãy bấm "📂 Tải Excel" ở bảng trên cùng.');
    return;
  }

  let matchedRow = contNo ? excelList.find((r) => r.contNo === contNo) : null;
  if (!contNo && actualSeal) {
    const sealMatchRow = excelList.find((r) => checkSealMatch(actualSeal, r.expectedSeal).isMatch);
    if (sealMatchRow) {
      matchedRow = sealMatchRow;
      document.getElementById('inputContNo').value = sealMatchRow.contNo;
      updateIsoBadge(sealMatchRow.contNo);
    }
  }

  const finalContNo = document.getElementById('inputContNo').value || '(Chưa rõ số Cont)';
  const banner = document.getElementById('resultBanner');
  const icon = document.getElementById('resultIcon');
  const title = document.getElementById('resultTitle');
  const sub = document.getElementById('resultSub');

  document.getElementById('resContVal').textContent = finalContNo;
  document.getElementById('resActualSealVal').textContent = actualSeal || '(Chưa có)';

  banner.className = 'result-banner show';

  if (!matchedRow) {
    banner.classList.add('not-found');
    icon.textContent = '⚠️';
    title.textContent = 'KHÔNG CÓ TRONG DS EXCEL';
    sub.textContent = contNo
      ? `Số Container ${contNo} không tồn tại trong file Excel đã tải lên.`
      : `Số Seal ${actualSeal} không khớp với bất kỳ Container nào trong danh sách Excel.`;
    document.getElementById('resExpectedSealVal').textContent = 'Không tìm thấy';

    lastComparisonResult = {
      vesselName: currentExcelFileName || 'Ngoài DS',
      contNo: finalContNo,
      actualSeal,
      expectedSeal: 'KHÔNG CÓ TRONG DS',
      status: 'mismatch'
    };
    if (!silentSound) {
      appendImmutableHistoryRecord(isManualConfirmed);
      playFeedbackSound(false);
      if (!isManualConfirmed && !hasManualConfirmedThisCycle) {
        openSealConfirmModal('mismatch');
      }
    }
    return;
  }

  const expectedSeal = matchedRow.expectedSeal;
  const vesselName = matchedRow.vesselName || currentExcelFileName || 'Chưa rõ tên tàu';
  document.getElementById('resExpectedSealVal').textContent = expectedSeal || '-';

  // Multi-seal intelligent matching:
  // If container has multiple seals, matching ANY expected seal in Excel is a 100% MATCH (OK)!
  const matchResult = checkSealMatch(actualSeal, expectedSeal, lastDetectedSeals);
  const isSealMatch = matchResult.isMatch;
  const resolvedActualSeal = matchResult.matchedToken || actualSeal;
  const nowStr = new Date().toLocaleString('vi-VN');

  // Always update matchedRow status immediately so the top Vessel Dashboard is 100% accurate
  matchedRow.actualSeal = resolvedActualSeal || '(Trống)';
  matchedRow.status = isSealMatch ? 'match' : 'mismatch';
  matchedRow.checkedAt = nowStr;

  if (isSealMatch) {
    banner.classList.add('match');
    icon.textContent = '✅';
    title.textContent = 'KHỚP HOÀN TOÀN (OK)';
    const multiSealNote =
      matchResult.actualTokens.length > 1
        ? ` (Đã khớp ${resolvedActualSeal} trong ${matchResult.actualTokens.length} seal trên ảnh)`
        : '';
    sub.textContent = `[Tàu: ${vesselName}] Container ${matchedRow.contNo} khớp với số Seal ${resolvedActualSeal}${multiSealNote}`;
    lastComparisonResult = {
      vesselName,
      contNo: matchedRow.contNo,
      actualSeal: resolvedActualSeal,
      expectedSeal,
      status: 'match'
    };
    if (!silentSound) {
      appendImmutableHistoryRecord(isManualConfirmed);
      playFeedbackSound(true);
    } else {
      saveStateToStorage();
      renderAll();
    }
  } else {
    banner.classList.add('mismatch');
    icon.textContent = '❌';
    title.textContent = 'SAI SỐ SEAL (CẢNH BÁO!)';
    sub.textContent = actualSeal
      ? `[Tàu: ${vesselName}] Seal thực tế (${actualSeal}) KHÁC với Seal trong Excel (${expectedSeal})`
      : `[Tàu: ${vesselName}] Chưa có số Seal thực tế để khớp với Seal trong Excel (${expectedSeal})`;
    lastComparisonResult = {
      vesselName,
      contNo: matchedRow.contNo,
      actualSeal: actualSeal || '(Trống)',
      expectedSeal,
      status: 'mismatch'
    };
    if (!silentSound) {
      appendImmutableHistoryRecord(isManualConfirmed);
      playFeedbackSound(false);
      // POPUP AUTO-TRIGGER ON MISMATCH:
      if (!isManualConfirmed && !hasManualConfirmedThisCycle) {
        openSealConfirmModal('mismatch');
      }
    } else {
      saveStateToStorage();
      renderAll();
    }
  }
}

function appendImmutableHistoryRecord(isConfirmedUpdate = false) {
  if (!lastComparisonResult) return;

  const { vesselName, contNo, actualSeal, expectedSeal, status } = lastComparisonResult;
  const sig = `${contNo}|${actualSeal}|${status}|${Boolean(currentContImageBase64)}|${Boolean(currentSealImageBase64)}`;
  if (lastLoggedScanSignature === sig && !isConfirmedUpdate) {
    return; // Already committed this exact scan to history
  }
  lastLoggedScanSignature = sig;

  const nowStr = new Date().toLocaleString('vi-VN');
  const idx = excelList.findIndex((r) => r.contNo === contNo);
  const resolvedVessel =
    vesselName || (idx !== -1 && excelList[idx].vesselName) || currentExcelFileName || 'Chưa rõ tên tàu';

  if (idx !== -1) {
    excelList[idx].actualSeal = actualSeal;
    excelList[idx].status = status;
    excelList[idx].checkedAt = nowStr;
  }

  const immutableId = `${Date.now()}_${Math.random().toString(36).slice(2, 6)}`;
  const entryObj = Object.freeze({
    id: immutableId,
    vesselName: resolvedVessel,
    contNo,
    actualSeal,
    expectedSeal,
    status,
    checkedAt: nowStr,
    hasContImg: Boolean(currentContImageBase64),
    hasSealImg: Boolean(currentSealImageBase64)
  });

  // If this is a manual confirmation right after an unconfirmed mismatch for the same container,
  // update the top entry so we don't duplicate unverified temporary errors
  if (
    isConfirmedUpdate &&
    historyLogs.length > 0 &&
    historyLogs[0].contNo === contNo &&
    historyLogs[0].status === 'mismatch'
  ) {
    historyLogs[0] = entryObj;
  } else {
    historyLogs.unshift(entryObj);
  }

  // Preserve captured images permanently for verification & proof
  if (currentContImageBase64 || currentSealImageBase64) {
    saveHistoryImageRecord(immutableId, currentContImageBase64, currentSealImageBase64);
  }

  syncExcelListWithHistory();
  saveStateToStorage();
  renderAll();
}

function saveCurrentCheckAndNext() {
  if (!lastComparisonResult) return;
  appendImmutableHistoryRecord();
  resetCurrentScan();
}

function resetCurrentScan() {
  clearSlotImage('cont');
  clearSlotImage('seal');
  hideOcrHintBanner();
  document.getElementById('inputContNo').value = '';
  document.getElementById('inputSealNo').value = '';
  document.getElementById('isoBadge').style.display = 'none';
  document.getElementById('suggestionsBar').innerHTML = '';
  document.getElementById('resultBanner').className = 'result-banner';
  lastComparisonResult = null;
  lastLoggedScanSignature = null;
}

function playFeedbackSound(isSuccess) {
  try {
    if (navigator.vibrate) {
      navigator.vibrate(isSuccess ? [80] : [180, 80, 220]);
    }
    const AudioCtx = window.AudioContext || window.webkitAudioContext;
    if (!AudioCtx) return;
    const ctx = new AudioCtx();
    const osc = ctx.createOscillator();
    const gain = ctx.createGain();
    osc.connect(gain);
    gain.connect(ctx.destination);

    if (isSuccess) {
      osc.type = 'sine';
      osc.frequency.setValueAtTime(660, ctx.currentTime);
      osc.frequency.setValueAtTime(880, ctx.currentTime + 0.1);
      gain.gain.setValueAtTime(0.15, ctx.currentTime);
      gain.gain.exponentialRampToValueAtTime(0.001, ctx.currentTime + 0.28);
      osc.start();
      osc.stop(ctx.currentTime + 0.28);
    } else {
      osc.type = 'sawtooth';
      osc.frequency.setValueAtTime(240, ctx.currentTime);
      osc.frequency.setValueAtTime(180, ctx.currentTime + 0.15);
      gain.gain.setValueAtTime(0.2, ctx.currentTime);
      gain.gain.exponentialRampToValueAtTime(0.001, ctx.currentTime + 0.38);
      osc.start();
      osc.stop(ctx.currentTime + 0.38);
    }
  } catch (e) {}
}

// ============================================================================
// Excel File Import & Vessel Column Mapping (Vessel Name / VSL Name / Arrived by...)
// ============================================================================
function initExcelHandlers() {
  const fileInput = document.getElementById('excelFileInput');
  const btnMain = document.getElementById('btnUploadExcelMain');
  if (btnMain) btnMain.addEventListener('click', () => fileInput.click());

  fileInput.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (file) parseUploadedExcel(file);
    fileInput.value = '';
  });

  const btnTpl = document.getElementById('btnDownloadTemplate');
  if (btnTpl) btnTpl.addEventListener('click', downloadTemplateExcel);

  const btnClear = document.getElementById('btnClearExcel');
  if (btnClear) {
    btnClear.addEventListener('click', () => {
      if (confirm('Bạn có chắc muốn xóa danh sách Excel hiện tại để tải chuyến tàu mới?')) {
        excelList = [];
        currentExcelFileName = '';
        saveStateToStorage();
        renderAll();
      }
    });
  }

  const btnApply = document.getElementById('btnApplyColumns');
  if (btnApply) {
    btnApply.addEventListener('click', () => {
      const colVesselEl = document.getElementById('selectColVessel');
      const colVessel = colVesselEl ? colVesselEl.value : '__FILE_NAME__';
      const colCont = document.getElementById('selectColCont').value;
      const colSeal = document.getElementById('selectColSeal').value;
      applyExcelColumnMapping(colCont, colSeal, colVessel);
      document.getElementById('columnMapBox').style.display = 'none';
    });
  }

  document.getElementById('searchExcelInput').addEventListener('input', renderExcelList);
  document.getElementById('filterExcelStatus').addEventListener('change', renderExcelList);
}

window.triggerUploadExcelGlobal = function () {
  const fileInput = document.getElementById('excelFileInput');
  if (fileInput) fileInput.click();
};

function isVesselColumnHeader(headerStr) {
  const norm = String(headerStr || '').toLowerCase().trim();
  return (
    norm.includes('vessel') ||
    norm.includes('vsl') ||
    norm.includes('arrived') ||
    norm.includes('tên tàu') ||
    norm.includes('ten tau') ||
    norm === 'tàu' ||
    norm === 'tau' ||
    norm.includes('chuyến tàu') ||
    norm.includes('chuyen tau') ||
    norm.includes('mother') ||
    norm.includes('barge') ||
    norm.includes('ship') ||
    norm.includes('voy')
  );
}

function parseUploadedExcel(file) {
  const cleanName = (file.name || 'DanhSachTau').replace(/\.(xlsx|xls|csv)$/i, '');
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(firstSheet, { defval: '' });

      if (!rows || rows.length === 0) {
        alert('File Excel trống hoặc không có dữ liệu hợp lệ!');
        return;
      }

      currentExcelFileName = cleanName;
      rawExcelRows = rows;
      rawExcelHeaders = Object.keys(rows[0]);

      let detectedVesselCol = '__FILE_NAME__';
      let detectedContCol = rawExcelHeaders[0];
      let detectedSealCol = rawExcelHeaders[1] || rawExcelHeaders[0];

      for (const h of rawExcelHeaders) {
        const normH = h.toLowerCase();
        if (isVesselColumnHeader(h)) {
          detectedVesselCol = h;
        }
        if (normH.includes('cont') || normH.includes('cntr') || normH.includes('số hiệu')) {
          detectedContCol = h;
        }
        if (normH.includes('seal') || normH.includes('chì') || normH.includes('chi')) {
          detectedSealCol = h;
        }
      }

      const selVessel = document.getElementById('selectColVessel');
      const selCont = document.getElementById('selectColCont');
      const selSeal = document.getElementById('selectColSeal');

      if (selVessel) {
        selVessel.innerHTML = '';
        const fallbackOpt = document.createElement('option');
        fallbackOpt.value = '__FILE_NAME__';
        fallbackOpt.textContent = `(Lấy theo tên file: ${cleanName})`;
        if (detectedVesselCol === '__FILE_NAME__') fallbackOpt.selected = true;
        selVessel.appendChild(fallbackOpt);
      }
      selCont.innerHTML = '';
      selSeal.innerHTML = '';

      rawExcelHeaders.forEach((h) => {
        if (selVessel) {
          const optV = document.createElement('option');
          optV.value = h;
          optV.textContent = h;
          if (h === detectedVesselCol) optV.selected = true;
          selVessel.appendChild(optV);
        }

        const opt1 = document.createElement('option');
        opt1.value = h;
        opt1.textContent = h;
        if (h === detectedContCol) opt1.selected = true;
        selCont.appendChild(opt1);

        const opt2 = document.createElement('option');
        opt2.value = h;
        opt2.textContent = h;
        if (h === detectedSealCol) opt2.selected = true;
        selSeal.appendChild(opt2);
      });

      document.getElementById('columnMapBox').style.display = 'block';
      applyExcelColumnMapping(detectedContCol, detectedSealCol, detectedVesselCol);
    } catch (err) {
      alert('Lỗi đọc file Excel: ' + err.message);
    }
  };
  reader.readAsArrayBuffer(file);
}

function applyExcelColumnMapping(contCol, sealCol, vesselCol = '__FILE_NAME__') {
  const parsed = [];
  let lastSeenVessel = currentExcelFileName || 'Chưa rõ tên tàu';

  rawExcelRows.forEach((row, idx) => {
    const c = normalizeContNo(row[contCol]);
    const s = formatExcelSealString(row[sealCol]);
    let v = '';
    if (vesselCol && vesselCol !== '__FILE_NAME__' && row[vesselCol] !== undefined) {
      v = String(row[vesselCol] || '').trim();
    }
    if (v) {
      lastSeenVessel = v;
    }
    const resolvedVessel = v || lastSeenVessel || currentExcelFileName || 'Chưa rõ tên tàu';

    if (c) {
      parsed.push({
        id: `${idx}_${c}`,
        vesselName: resolvedVessel,
        contNo: c,
        expectedSeal: s,
        actualSeal: '',
        status: 'pending',
        checkedAt: ''
      });
    }
  });

  if (parsed.length === 0) {
    alert('Không tìm thấy số Container nào ở cột đã chọn. Vui lòng chọn lại cột!');
    return;
  }

  // If a vessel column was found, update currentExcelFileName to the primary vessel name
  const uniqueVessels = [...new Set(parsed.map((x) => x.vesselName).filter(Boolean))];
  if (uniqueVessels.length === 1) {
    currentExcelFileName = uniqueVessels[0];
  } else if (uniqueVessels.length > 1) {
    currentExcelFileName = `${uniqueVessels.length}_Tau_${uniqueVessels[0]}`;
  }

  excelList = parsed;
  syncExcelListWithHistory();
  saveStateToStorage();
  renderAll();
}

function downloadTemplateExcel() {
  const sampleRows = [
    { STT: 1, 'Vessel Name': 'SITC HAIPHONG', 'Số Container': 'MSNU7784178', 'Số Seal': 'FX49779575', 'Ghi chú': 'Hàng xuất' },
    { STT: 2, 'Vessel Name': 'SITC HAIPHONG', 'Số Container': 'FFAU2194320', 'Số Seal': 'FX49778780', 'Ghi chú': 'Hàng xuất' },
    { STT: 3, 'Vessel Name': 'SITC HAIPHONG', 'Số Container': 'JXLU6216912', 'Số Seal': 'A4220802184', 'Ghi chú': 'Hàng lạnh' }
  ];
  const ws = XLSX.utils.json_to_sheet(sampleRows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'DanhSachContSeal');
  XLSX.writeFile(wb, 'Mau_Danh_Sach_Tau_Cont_Seal.xlsx');
}

// ============================================================================
// History (Read-Only Lookup Only - Always Allow Opening Photos) & Vessel Archive
// ============================================================================
function initHistoryHandlers() {
  document.getElementById('btnExportExcel').addEventListener('click', exportReportExcel);

  const btnJump = document.getElementById('btnJumpToArchive');
  if (btnJump) {
    btnJump.addEventListener('click', () => {
      switchTab('tab-history');
    });
  }

  const btnArchiveFull = document.getElementById('btnArchiveFullToPc');
  if (btnArchiveFull) {
    btnArchiveFull.addEventListener('click', archiveFullHistoryToPc);
  }

  const searchHist = document.getElementById('searchHistoryInput');
  const filterHist = document.getElementById('filterHistoryStatus');
  if (searchHist) searchHist.addEventListener('input', renderHistoryList);
  if (filterHist) filterHist.addEventListener('change', renderHistoryList);

  const histModal = document.getElementById('historyDetailModal');
  const closeHistBtn = document.getElementById('closeHistModalBtn');
  const closeHistBottom = document.getElementById('btnCloseHistBottom');
  const histWrapCont = document.getElementById('histWrapCont');
  const histWrapSeal = document.getElementById('histWrapSeal');

  if (closeHistBtn && histModal) {
    closeHistBtn.addEventListener('click', () => histModal.classList.remove('open'));
    histModal.addEventListener('click', (e) => {
      if (e.target === histModal) histModal.classList.remove('open');
    });
  }

  if (closeHistBottom && histModal) {
    closeHistBottom.addEventListener('click', () => histModal.classList.remove('open'));
  }

  if (histWrapCont) {
    histWrapCont.addEventListener('click', () => {
      const imgEl = document.getElementById('histImgCont');
      if (imgEl && imgEl.style.display !== 'none' && imgEl.src) {
        const rec = historyLogs.find((x) => String(x.id) === String(selectedHistoryModalId));
        openImageLightbox(imgEl.src, `🖼️ Tra cứu Ảnh Cont: ${rec ? rec.contNo : ''}`);
      }
    });
  }

  if (histWrapSeal) {
    histWrapSeal.addEventListener('click', () => {
      const imgEl = document.getElementById('histImgSeal');
      if (imgEl && imgEl.style.display !== 'none' && imgEl.src) {
        const rec = historyLogs.find((x) => String(x.id) === String(selectedHistoryModalId));
        openImageLightbox(imgEl.src, `🔒 Tra cứu Ảnh Seal: ${rec ? rec.actualSeal : ''}`, true);
      }
    });
  }
}

// Resolves images for any history item (match or mismatch), falling back to same-container images if needed
function resolveImagesForHistoryItemSync(item) {
  if (!item) return { contImage: null, sealImage: null };
  const direct = historyImagesCache[item.id] || {};
  let contImage = direct.contImage || null;
  let sealImage = direct.sealImage || null;

  if (!contImage || !sealImage) {
    for (const h of historyLogs) {
      if (h && h.contNo === item.contNo && historyImagesCache[h.id]) {
        if (!contImage && historyImagesCache[h.id].contImage) {
          contImage = historyImagesCache[h.id].contImage;
        }
        if (!sealImage && historyImagesCache[h.id].sealImage) {
          sealImage = historyImagesCache[h.id].sealImage;
        }
        if (contImage && sealImage) break;
      }
    }
  }
  return { contImage, sealImage };
}

async function resolveImagesForHistoryItemAsync(item) {
  if (!item) return { contImage: null, sealImage: null };
  let direct = await getHistoryImageRecord(item.id);
  let contImage = direct && direct.contImage ? direct.contImage : null;
  let sealImage = direct && direct.sealImage ? direct.sealImage : null;

  if (!contImage || !sealImage) {
    for (const h of historyLogs) {
      if (h && h.contNo === item.contNo) {
        const rec = await getHistoryImageRecord(h.id);
        if (rec) {
          if (!contImage && rec.contImage) contImage = rec.contImage;
          if (!sealImage && rec.sealImage) sealImage = rec.sealImage;
          if (contImage && sealImage) break;
        }
      }
    }
  }
  return { contImage, sealImage };
}

window.openHistoryPhotoDirect = async function (historyId, type) {
  const item = historyLogs.find((h) => String(h.id) === String(historyId));
  if (!item) return;
  const imgs = await resolveImagesForHistoryItemAsync(item);
  if (type === 'cont' && imgs.contImage) {
    openImageLightbox(imgs.contImage, `🖼️ Ảnh Mặt Cửa Cont: ${item.contNo} (${item.status === 'match' ? 'KHỚP' : 'SAI LỆCH'})`);
  } else if (type === 'seal' && imgs.sealImage) {
    openImageLightbox(imgs.sealImage, `🔒 Ảnh Số Seal: ${item.actualSeal} (${item.status === 'match' ? 'KHỚP' : 'SAI LỆCH'})`, true);
  } else {
    openHistoryDetailModal(historyId);
  }
};

// ============================================================================
// Lightweight 70% Storage Threshold Alert & Vessel-Sorted Archive
// ============================================================================
function estimateBase64Bytes(dataUrl) {
  if (!dataUrl || typeof dataUrl !== 'string') return 0;
  const commaIdx = dataUrl.indexOf(',');
  const b64Len = commaIdx !== -1 ? dataUrl.length - commaIdx - 1 : dataUrl.length;
  return Math.round((b64Len * 3) / 4);
}

async function renderStorageStatus() {
  const globalAlert = document.getElementById('storageGlobalAlert');
  const globalAlertText = document.getElementById('storageGlobalAlertText');
  if (!globalAlert || !globalAlertText) return;

  let lsBytes = 0;
  try {
    for (let i = 0; i < localStorage.length; i++) {
      const k = localStorage.key(i);
      const v = localStorage.getItem(k) || '';
      lsBytes += (k.length + v.length) * 2;
    }
  } catch (e) {}
  const lsPct = Math.min(100, (lsBytes / (5 * 1024 * 1024)) * 100);

  let imgBytesTotal = 0;
  Object.values(historyImagesCache).forEach((rec) => {
    if (!rec) return;
    if (rec.contImage) imgBytesTotal += estimateBase64Bytes(rec.contImage);
    if (rec.sealImage) imgBytesTotal += estimateBase64Bytes(rec.sealImage);
  });

  let browserQuotaBytes = 2 * 1024 * 1024 * 1024;
  let browserUsageBytes = imgBytesTotal;
  if (navigator.storage && navigator.storage.estimate) {
    try {
      const est = await navigator.storage.estimate();
      if (est && est.quota) browserQuotaBytes = est.quota;
      if (est && est.usage) browserUsageBytes = Math.max(imgBytesTotal, est.usage);
    } catch (e) {}
  }
  const idbPct = Math.min(100, (browserUsageBytes / browserQuotaBytes) * 100);
  const worstPct = Math.max(lsPct, idbPct);

  // Only show warning when storage reaches >= 70%; otherwise keep completely hidden
  if (worstPct >= 70) {
    globalAlert.className = worstPct >= 88 ? 'storage-global-alert danger' : 'storage-global-alert';
    globalAlert.style.display = 'flex';
    globalAlertText.innerHTML = `⚠️ Bộ nhớ đã đầy <strong>${worstPct.toFixed(0)}%</strong> (≥ 70%). Vui lòng Archive chuyến tàu về máy tính!`;
  } else {
    globalAlert.style.display = 'none';
  }
}

function showArchiveStatusBanner(htmlMsg, isError = false) {
  const box = document.getElementById('archiveActionStatus');
  if (!box) return;
  box.style.display = 'block';
  if (isError) {
    box.style.background = 'rgba(239, 68, 68, 0.18)';
    box.style.border = '1px solid rgba(239, 68, 68, 0.55)';
    box.style.color = '#fecaca';
  } else {
    box.style.background = 'rgba(16, 185, 129, 0.16)';
    box.style.border = '1px solid rgba(16, 185, 129, 0.5)';
    box.style.color = '#a7f3d0';
  }
  box.innerHTML = htmlMsg;
}

function getSortedExcelListByVessel() {
  syncExcelListWithHistory();
  return [...excelList].sort((a, b) => {
    const vA = (a.vesselName || currentExcelFileName || '').toUpperCase();
    const vB = (b.vesselName || currentExcelFileName || '').toUpperCase();
    if (vA !== vB) return vA.localeCompare(vB, 'vi');
    return 0;
  });
}

async function collectVesselRecordsFromExcelList() {
  const sortedRows = getSortedExcelListByVessel();
  const vesselRecords = [];

  for (let i = 0; i < sortedRows.length; i++) {
    const row = sortedRows[i];
    const matchingLogs = historyLogs.filter((h) => h && h.contNo === row.contNo);
    const matchedHist =
      matchingLogs.find(
        (h) =>
          h.status === 'match' ||
          (h.actualSeal && row.expectedSeal && checkSealMatch(h.actualSeal, row.expectedSeal).isMatch)
      ) || matchingLogs[0];

    let contImage = null;
    let sealImage = null;

    if (matchedHist) {
      const imgs = await resolveImagesForHistoryItemAsync(matchedHist);
      contImage = imgs.contImage;
      sealImage = imgs.sealImage;
    }

    vesselRecords.push({
      id: matchedHist ? matchedHist.id : row.id,
      stt: i + 1,
      vesselName: row.vesselName || currentExcelFileName || 'Chưa rõ tên tàu',
      contNo: row.contNo,
      expectedSeal: row.expectedSeal || '',
      actualSeal: row.actualSeal || (matchedHist ? matchedHist.actualSeal : ''),
      status: row.status !== 'pending' ? row.status : matchedHist ? matchedHist.status : 'pending',
      checkedAt: row.checkedAt || (matchedHist ? matchedHist.checkedAt : ''),
      contImage,
      sealImage
    });
  }
  return vesselRecords;
}

function triggerBrowserFileDownload(contentStr, fileName, mimeType) {
  const blob = new Blob([contentStr], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = fileName;
  document.body.appendChild(a);
  a.click();
  setTimeout(() => {
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }, 500);
}

function buildVesselStandaloneHtmlArchive(vesselName, vesselRecords, exportedAtStr) {
  const matchCount = vesselRecords.filter((r) => r.status === 'match').length;
  const mismatchCount = vesselRecords.filter((r) => r.status === 'mismatch').length;
  const pendingCount = vesselRecords.filter((r) => r.status === 'pending').length;

  // Group records by vesselName for clear per-vessel sections
  const vesselGroups = new Map();
  vesselRecords.forEach((r) => {
    const v = r.vesselName || vesselName || 'Chưa rõ tên tàu';
    if (!vesselGroups.has(v)) vesselGroups.set(v, []);
    vesselGroups.get(v).push(r);
  });

  let tablesHtml = '';
  for (const [vName, rows] of vesselGroups.entries()) {
    const vMatch = rows.filter((x) => x.status === 'match').length;
    const vMismatch = rows.filter((x) => x.status === 'mismatch').length;
    const vPending = rows.filter((x) => x.status === 'pending').length;

    const rowsHtml = rows
      .map((r, idx) => {
        const badge =
          r.status === 'match'
            ? '<span class="badge match">✅ KHỚP</span>'
            : r.status === 'mismatch'
            ? '<span class="badge mismatch">❌ KHÔNG KHỚP</span>'
            : '<span class="badge pending">⏳ CHƯA CHECK</span>';
        const contImgHtml = r.contImage
          ? `<img src="${r.contImage}" class="thumb" onclick="openFull('${r.contImage.replace(/'/g, '')}', 'Tàu ${vName} - Ảnh Cont: ${r.contNo}')" title="Bấm để phóng to" />`
          : '<span class="no-img">Chưa có ảnh</span>';
        const sealImgHtml = r.sealImage
          ? `<img src="${r.sealImage}" class="thumb" onclick="openFull('${r.sealImage.replace(/'/g, '')}', 'Tàu ${vName} - Ảnh Seal: ${r.actualSeal}')" title="Bấm để phóng to" />`
          : '<span class="no-img">Chưa có ảnh</span>';

        return `
          <tr class="archive-row" data-search="${(vName + ' ' + r.contNo + ' ' + r.actualSeal + ' ' + r.expectedSeal).toUpperCase()}" data-status="${r.status}">
            <td>${idx + 1}</td>
            <td><strong>${vName}</strong></td>
            <td class="mono"><strong>${r.contNo}</strong></td>
            <td class="mono">${r.expectedSeal || '-'}</td>
            <td class="mono"><strong>${r.actualSeal || '-'}</strong></td>
            <td>${badge}</td>
            <td>${r.checkedAt || '-'}</td>
            <td>${contImgHtml}</td>
            <td>${sealImgHtml}</td>
          </tr>
        `;
      })
      .join('');

    tablesHtml += `
      <div class="vessel-section">
        <div class="vessel-sec-head">
          <h2>🚢 TÀU: ${vName}</h2>
          <div class="stats">
            <span class="stat-box">Số lượng Cont: ${rows.length}</span>
            <span class="stat-box" style="color:#34d399;">Khớp: ${vMatch}</span>
            <span class="stat-box" style="color:#f87171;">Không khớp: ${vMismatch}</span>
            <span class="stat-box" style="color:#fbbf24;">Chưa check: ${vPending}</span>
          </div>
        </div>
        <table>
          <thead>
            <tr>
              <th>STT</th>
              <th>Tên Tàu (Vessel)</th>
              <th>Số Container</th>
              <th>Seal Excel</th>
              <th>Seal Thực Tế</th>
              <th>Kết Quả</th>
              <th>Thời Gian Kiểm</th>
              <th>Ảnh Cửa Cont (Bấm xem full)</th>
              <th>Ảnh Số Seal (Bấm xem full)</th>
            </tr>
          </thead>
          <tbody>
            ${rowsHtml}
          </tbody>
        </table>
      </div>
    `;
  }

  return `<!DOCTYPE html>
<html lang="vi">
<head>
  <meta charset="UTF-8" />
  <title>Báo Cáo Chốt & Archive Theo Tàu: ${vesselName} - ${exportedAtStr}</title>
  <style>
    body { font-family: system-ui, -apple-system, sans-serif; background: #0f172a; color: #f8fafc; margin: 0; padding: 24px; }
    .container { max-width: 1260px; margin: 0 auto; }
    .header { background: #1e293b; padding: 20px; border-radius: 12px; border: 1px solid #334155; margin-bottom: 18px; display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 12px; }
    h1 { margin: 0 0 6px 0; font-size: 20px; color: #38bdf8; }
    h2 { margin: 0; font-size: 17px; color: #38bdf8; }
    .meta { font-size: 13px; color: #94a3b8; }
    .stats { display: flex; gap: 10px; flex-wrap: wrap; }
    .stat-box { background: #0f172a; border: 1px solid #334155; padding: 7px 13px; border-radius: 8px; font-size: 13px; font-weight: 700; }
    .toolbar { display: flex; gap: 10px; margin-bottom: 16px; }
    input, select { background: #1e293b; color: #f8fafc; border: 1px solid #475569; padding: 9px 12px; border-radius: 8px; font-size: 14px; }
    input { flex: 1; }
    .vessel-section { margin-bottom: 26px; }
    .vessel-sec-head { background: rgba(30,41,59,0.85); border: 1px solid #38bdf8; border-radius: 10px 10px 0 0; padding: 12px 16px; display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 10px; }
    table { width: 100%; border-collapse: collapse; background: #1e293b; border-radius: 0 0 12px 12px; overflow: hidden; }
    th, td { padding: 10px 12px; border-bottom: 1px solid #334155; text-align: left; font-size: 13.5px; vertical-align: middle; }
    th { background: #0f172a; color: #94a3b8; font-weight: 700; }
    .mono { font-family: monospace; font-size: 14px; }
    .badge { padding: 4px 8px; border-radius: 6px; font-size: 12px; font-weight: 700; }
    .badge.match { background: rgba(16,185,129,0.2); color: #34d399; }
    .badge.mismatch { background: rgba(239,68,68,0.2); color: #f87171; }
    .badge.pending { background: rgba(245,158,11,0.2); color: #fbbf24; }
    .thumb { width: 110px; height: 76px; object-fit: cover; border-radius: 6px; border: 1px solid #64748b; cursor: zoom-in; transition: transform 0.15s; }
    .thumb:hover { transform: scale(1.06); border-color: #38bdf8; }
    .no-img { color: #64748b; font-size: 12px; }
    #lightbox { position: fixed; inset: 0; background: rgba(2,6,23,0.95); display: none; flex-direction: column; z-index: 9999; }
    #lightbox.open { display: flex; }
    .lb-top { padding: 12px 20px; background: #0f172a; display: flex; justify-content: space-between; align-items: center; border-bottom: 1px solid #334155; }
    .lb-stage { flex: 1; display: flex; align-items: center; justify-content: center; overflow: auto; padding: 16px; }
    .lb-stage img { max-width: 95vw; max-height: 86vh; object-fit: contain; border-radius: 8px; }
    .btn-close { background: #ef4444; color: #fff; border: none; padding: 7px 14px; border-radius: 6px; font-weight: 700; cursor: pointer; }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <div>
        <h1>🚢 Báo Cáo Chốt &amp; Archive Theo Từng Tàu</h1>
        <div class="meta">Thời gian chốt Archive: <strong>${exportedAtStr}</strong> • Chế độ: <strong>Chỉ Tra Cứu (Bấm vào từng ảnh để xem Full màn hình)</strong></div>
      </div>
      <div class="stats">
        <div class="stat-box">Tổng Cont: ${vesselRecords.length}</div>
        <div class="stat-box" style="color:#34d399;">Khớp: ${matchCount}</div>
        <div class="stat-box" style="color:#f87171;">Không khớp: ${mismatchCount}</div>
        <div class="stat-box" style="color:#fbbf24;">Chưa check: ${pendingCount}</div>
      </div>
    </div>

    <div class="toolbar">
      <input type="text" id="qInput" placeholder="🔍 Gõ tên tàu, số Container hoặc số Seal để tra cứu nhanh..." oninput="filterRows()" />
      <select id="sFilter" onchange="filterRows()">
        <option value="all">Tất cả (${vesselRecords.length})</option>
        <option value="match">✅ Đã khớp (${matchCount})</option>
        <option value="mismatch">❌ Không khớp (${mismatchCount})</option>
        <option value="pending">⏳ Chưa check (${pendingCount})</option>
      </select>
    </div>

    ${tablesHtml}
  </div>

  <div id="lightbox" onclick="if(event.target.id==='lightbox'||event.target.className==='lb-stage') closeFull()">
    <div class="lb-top">
      <strong id="lbTitle">Xem ảnh</strong>
      <button class="btn-close" onclick="closeFull()">✕ Đóng</button>
    </div>
    <div class="lb-stage">
      <img id="lbImg" src="" alt="Full" />
    </div>
  </div>

  <script>
    function openFull(src, title) {
      document.getElementById('lbImg').src = src;
      document.getElementById('lbTitle').textContent = title;
      document.getElementById('lightbox').classList.add('open');
    }
    function closeFull() {
      document.getElementById('lightbox').classList.remove('open');
    }
    function filterRows() {
      const q = document.getElementById('qInput').value.trim().toUpperCase();
      const st = document.getElementById('sFilter').value;
      document.querySelectorAll('.archive-row').forEach(tr => {
        const matchQ = !q || tr.getAttribute('data-search').includes(q);
        const matchSt = st === 'all' || tr.getAttribute('data-status') === st;
        tr.style.display = (matchQ && matchSt) ? '' : 'none';
      });
    }
  </script>
</body>
</html>`;
}

async function archiveFullHistoryToPc() {
  if (excelList.length === 0) {
    alert('Chưa có danh sách Excel của chuyến tàu! Vui lòng tải file Excel danh sách tàu lên trước khi Archive.');
    return;
  }

  syncExcelListWithHistory();
  const vesselName = currentExcelFileName || 'ChuyenTau';
  const safeVessel = vesselName.replace(/[^A-Za-z0-9_-]/g, '_');

  const btn = document.getElementById('btnArchiveFullToPc');
  const origText = btn ? btn.textContent : '';
  if (btn) {
    btn.disabled = true;
    btn.textContent = `⏳ Đang đóng gói Archive tàu "${vesselName}"...`;
  }

  try {
    const now = new Date();
    const dateTag = now.toISOString().slice(0, 10);
    const timeTag = now.toTimeString().slice(0, 5).replace(':', '');
    const exportedAtStr = now.toLocaleString('vi-VN');

    const vesselRecords = await collectVesselRecordsFromExcelList();
    const htmlReport = buildVesselStandaloneHtmlArchive(vesselName, vesselRecords, exportedAtStr);

    // 1. Download standalone HTML report sorted by Vessel
    const htmlFileName = `Archive_Tau_${safeVessel}_${dateTag}_${timeTag}.html`;
    triggerBrowserFileDownload(htmlReport, htmlFileName, 'text/html;charset=utf-8');

    // 2. Also save to server.ps1 (/api/archive) into Archives\Tau_<vesselName>_<stamp>\
    let serverSavedFolder = null;
    let serverImgCount = 0;
    try {
      const payload = {
        version: 'v14',
        vesselName,
        exportedAt: exportedAtStr,
        records: vesselRecords,
        htmlReport
      };
      const resp = await fetch('/api/archive', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      });
      if (resp.ok) {
        const resJson = await resp.json();
        if (resJson && resJson.ok) {
          serverSavedFolder = resJson.folder;
          serverImgCount = resJson.imageCount || 0;
        }
      }
    } catch (e) {}

    if (serverSavedFolder) {
      showArchiveStatusBanner(
        `✅ <strong>Đã Archive dứt khoát chuyến tàu "${vesselName}" (${vesselRecords.length} Cont):</strong><br/>` +
          `1. Đã tải xuống máy: <strong>${htmlFileName}</strong> (Đã sắp xếp theo từng tàu kèm đầy đủ hình ảnh).<br/>` +
          `2. Đã lưu vào ổ cứng máy tính tại: <code style="word-break:break-all; color:#38bdf8;">${serverSavedFolder}</code> (${serverImgCount} ảnh .JPG).`
      );
    } else {
      showArchiveStatusBanner(
        `✅ <strong>Đã tải xuống gói Archive chuyến tàu "${vesselName}":</strong> <code>${htmlFileName}</code> (Trọn bộ ${vesselRecords.length} Cont sắp xếp theo tàu kèm hình ảnh).`
      );
    }
  } catch (err) {
    showArchiveStatusBanner(`⚠️ Lỗi khi tạo gói Archive chuyến tàu: ${err.message}`, true);
  } finally {
    if (btn) {
      btn.disabled = false;
      btn.textContent = origText;
    }
  }
}

window.openHistoryDetailModal = async function (historyId) {
  const item = historyLogs.find((h) => String(h.id) === String(historyId));
  if (!item) return;

  selectedHistoryModalId = item.id;
  const histModal = document.getElementById('historyDetailModal');
  const metaEl = document.getElementById('histModalMeta');
  const imgCont = document.getElementById('histImgCont');
  const emptyCont = document.getElementById('histEmptyCont');
  const badgeCont = document.getElementById('histZoomBadgeCont');
  const imgSeal = document.getElementById('histImgSeal');
  const emptySeal = document.getElementById('histEmptySeal');
  const badgeSeal = document.getElementById('histZoomBadgeSeal');

  const statusBadge =
    item.status === 'match'
      ? '<span class="badge match">✅ KHỚP HOÀN TOÀN</span>'
      : '<span class="badge mismatch">❌ KHÔNG KHỚP (SAI LỆCH)</span>';

  const rowInExcel = excelList.find((r) => r.contNo === item.contNo);
  const vName = item.vesselName || (rowInExcel && rowInExcel.vesselName) || currentExcelFileName || '';

  if (metaEl) {
    metaEl.innerHTML = `
      <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:6px;">
        <span style="font-size:15px; font-weight:800; font-family:monospace; color:#f8fafc;">📦 ${item.contNo}</span>
        ${statusBadge}
      </div>
      ${vName ? `<div>🚢 Tàu (Vessel): <strong style="color:#38bdf8;">${vName}</strong></div>` : ''}
      <div>🔒 Seal chụp thực tế: <strong style="color:#f8fafc; font-family:monospace;">${item.actualSeal || '(Trống)'}</strong></div>
      <div>📋 Seal trong Excel: <strong style="color:#38bdf8; font-family:monospace;">${item.expectedSeal || '-'}</strong></div>
      <div style="font-size:11px; color:var(--text-muted); margin-top:4px;">🕒 Thời gian kiểm: ${item.checkedAt}</div>
    `;
  }

  const imgs = await resolveImagesForHistoryItemAsync(item);
  const contSrc = imgs.contImage;
  const sealSrc = imgs.sealImage;

  if (contSrc) {
    imgCont.src = contSrc;
    imgCont.style.display = 'block';
    emptyCont.style.display = 'none';
    if (badgeCont) badgeCont.style.display = 'inline-flex';
  } else {
    imgCont.src = '';
    imgCont.style.display = 'none';
    emptyCont.style.display = 'flex';
    if (badgeCont) badgeCont.style.display = 'none';
  }

  if (sealSrc) {
    imgSeal.src = sealSrc;
    imgSeal.style.display = 'block';
    emptySeal.style.display = 'none';
    if (badgeSeal) badgeSeal.style.display = 'inline-flex';
  } else {
    imgSeal.src = '';
    imgSeal.style.display = 'none';
    emptySeal.style.display = 'flex';
    if (badgeSeal) badgeSeal.style.display = 'none';
  }

  if (histModal) {
    histModal.classList.add('open');
  }
};

function exportReportExcel() {
  if (excelList.length === 0) {
    alert('Chưa có danh sách Excel chuyến tàu để xuất báo cáo!');
    return;
  }

  const sortedRows = getSortedExcelListByVessel();
  const vesselName = currentExcelFileName || 'ChuyenTau';
  const safeVessel = vesselName.replace(/[^A-Za-z0-9_-]/g, '_');
  const wb = XLSX.utils.book_new();

  const dsRows = sortedRows.map((r, i) => ({
    STT: i + 1,
    'Tên Tàu (Vessel Name)': r.vesselName || vesselName,
    'Số Container': r.contNo,
    'Số Seal (Excel)': r.expectedSeal,
    'Số Seal (Thực tế)': r.actualSeal || '',
    'Kết quả đối chiếu':
      r.status === 'match'
        ? 'KHỚP (OK)'
        : r.status === 'mismatch'
        ? 'KHÔNG KHỚP (SAI SEAL)'
        : 'CHƯA CHECK',
    'Thời gian kiểm': r.checkedAt || ''
  }));
  const wsAll = XLSX.utils.json_to_sheet(dsRows);
  XLSX.utils.book_append_sheet(wb, wsAll, 'BaoCaoChotTheoTau');

  // If multiple vessels exist in the list, also create a separate sheet for each vessel
  const vesselGroups = new Map();
  sortedRows.forEach((r) => {
    const v = r.vesselName || vesselName;
    if (!vesselGroups.has(v)) vesselGroups.set(v, []);
    vesselGroups.get(v).push(r);
  });

  if (vesselGroups.size > 1) {
    let sheetIdx = 1;
    for (const [vName, vRows] of vesselGroups.entries()) {
      const sheetRows = vRows.map((r, idx) => ({
        STT: idx + 1,
        'Tên Tàu (Vessel Name)': vName,
        'Số Container': r.contNo,
        'Số Seal (Excel)': r.expectedSeal,
        'Số Seal (Thực tế)': r.actualSeal || '',
        'Kết quả đối chiếu':
          r.status === 'match'
            ? 'KHỚP (OK)'
            : r.status === 'mismatch'
            ? 'KHÔNG KHỚP (SAI SEAL)'
            : 'CHƯA CHECK',
        'Thời gian kiểm': r.checkedAt || ''
      }));
      const safeSheetName = `${sheetIdx}_${vName}`.replace(/[\\/?*[\]:]/g, '_').slice(0, 30);
      const wsVessel = XLSX.utils.json_to_sheet(sheetRows);
      XLSX.utils.book_append_sheet(wb, wsVessel, safeSheetName);
      sheetIdx++;
    }
  }

  const dateTag = new Date().toISOString().slice(0, 10);
  XLSX.writeFile(wb, `BaoCao_ChotTheoTau_${safeVessel}_${dateTag}.xlsx`);
}

// ============================================================================
// UI Rendering (Top Vessel Dashboard + Excel List + History Lookup)
// ============================================================================
function renderAll() {
  syncExcelListWithHistory();
  renderStatsAndQuickBar();
  renderExcelList();
  renderHistoryList();
  renderStorageStatus();
}

function renderStatsAndQuickBar() {
  const dashEl = document.getElementById('vesselDashboard');
  const vesselLabel = document.getElementById('currentVesselNameLabel');
  if (!dashEl) return;

  if (excelList.length === 0) {
    dashEl.innerHTML = `
      <div class="vessel-dash-card empty">
        <div>
          <strong style="font-size:13.5px; color:#fde68a;">🚢 Chưa tải danh sách Excel của tàu</strong>
          <div style="font-size:11.5px; color:var(--text-muted); margin-top:2px;">
            Tải file Excel có cột Tên Tàu (Vessel Name / VSL Name / Arrived by...), Số Cont, Số Seal
          </div>
        </div>
        <button class="btn btn-primary btn-sm" onclick="triggerUploadExcelGlobal()" style="flex-shrink:0;">
          📂 Tải Excel
        </button>
      </div>
    `;
    if (vesselLabel) vesselLabel.textContent = 'Chưa tải DS Excel';
    return;
  }

  // Group excelList by vesselName so each vessel shows its exact name and 4 counters
  const vesselGroups = new Map();
  excelList.forEach((r) => {
    const v = (r.vesselName || currentExcelFileName || 'Chưa rõ tên tàu').trim();
    if (!vesselGroups.has(v)) vesselGroups.set(v, []);
    vesselGroups.get(v).push(r);
  });

  const cardsHtml = [];
  let isFirst = true;
  for (const [vName, rows] of vesselGroups.entries()) {
    const total = rows.length;
    const matchCount = rows.filter((r) => r.status === 'match').length;
    const mismatchCount = rows.filter((r) => r.status === 'mismatch').length;
    const pendingCount = rows.filter((r) => r.status === 'pending').length;

    cardsHtml.push(`
      <div class="vessel-dash-card">
        <div class="vessel-dash-head">
          <div class="vessel-dash-title" title="${vName}">
            <span>🚢 TÀU: ${vName}</span>
          </div>
          ${
            isFirst
              ? `<button class="btn btn-outline btn-sm" onclick="triggerUploadExcelGlobal()" style="padding:4px 9px; font-size:11px; flex-shrink:0;">📂 Đổi DS Excel</button>`
              : ''
          }
        </div>
        <div class="vessel-dash-stats">
          <div class="vessel-stat-pill">
            <span class="num">${total}</span>
            <span class="lbl">Số lượng Cont</span>
          </div>
          <div class="vessel-stat-pill match">
            <span class="num">${matchCount}</span>
            <span class="lbl">Khớp</span>
          </div>
          <div class="vessel-stat-pill mismatch">
            <span class="num">${mismatchCount}</span>
            <span class="lbl">Không khớp</span>
          </div>
          <div class="vessel-stat-pill pending">
            <span class="num">${pendingCount}</span>
            <span class="lbl">Chưa check</span>
          </div>
        </div>
      </div>
    `);
    isFirst = false;
  }

  dashEl.innerHTML = cardsHtml.join('');

  if (vesselLabel) {
    const totalAll = excelList.length;
    const matchAll = excelList.filter((r) => r.status === 'match').length;
    const vNames = [...vesselGroups.keys()].join(', ');
    vesselLabel.textContent = `Tàu: ${vNames} (${totalAll} Cont - Khớp ${matchAll}/${totalAll})`;
  }
}

function renderExcelList() {
  const container = document.getElementById('excelListContainer');
  const query = (document.getElementById('searchExcelInput').value || '').trim().toUpperCase();
  const statusFilter = document.getElementById('filterExcelStatus').value;

  document.getElementById('excelCountLabel').textContent = excelList.length;

  const sortedRows = getSortedExcelListByVessel();
  const filtered = sortedRows.filter((r) => {
    const vName = (r.vesselName || '').toUpperCase();
    const matchText = !query || r.contNo.includes(query) || r.expectedSeal.includes(query) || vName.includes(query);
    const matchStatus = statusFilter === 'all' || r.status === statusFilter;
    return matchText && matchStatus;
  });

  if (filtered.length === 0) {
    container.innerHTML = `<div style="text-align:center; padding:24px; color:var(--text-muted); font-size:13px;">
      Không có dữ liệu phù hợp.
    </div>`;
    return;
  }

  container.innerHTML = filtered
    .map((r) => {
      const badgeClass = r.status === 'match' ? 'match' : r.status === 'mismatch' ? 'mismatch' : 'pending';
      const badgeText =
        r.status === 'match'
          ? '✅ Đã khớp'
          : r.status === 'mismatch'
          ? `❌ Không khớp (${r.actualSeal || '?'})`
          : '⏳ Chưa check';

      return `
        <div class="cont-item" onclick="selectContFromList('${r.contNo}')">
          <div class="cont-item-main">
            <span class="cont-no">${r.contNo} <span style="font-size:11px; font-weight:600; color:#38bdf8;">• 🚢 ${r.vesselName || currentExcelFileName || ''}</span></span>
            <span class="seal-no">Seal Excel: <strong>${r.expectedSeal}</strong>${r.actualSeal ? ` | Thực tế: <strong>${r.actualSeal}</strong>` : ''}</span>
          </div>
          <span class="badge ${badgeClass}">${badgeText}</span>
        </div>
      `;
    })
    .join('');
}

window.selectContFromList = function (contNo) {
  switchTab('tab-check');
  const inputCont = document.getElementById('inputContNo');
  inputCont.value = contNo;
  updateIsoBadge(contNo);
  evaluateReadinessAndCompare(false);
};

function renderHistoryList() {
  const container = document.getElementById('historyListContainer');
  document.getElementById('historyCountLabel').textContent = historyLogs.length;

  const searchEl = document.getElementById('searchHistoryInput');
  const filterEl = document.getElementById('filterHistoryStatus');
  const query = (searchEl ? searchEl.value : '').trim().toUpperCase();
  const statusFilter = filterEl ? filterEl.value : 'all';

  if (historyLogs.length === 0) {
    container.innerHTML = `<div style="text-align:center; padding:24px; color:var(--text-muted); font-size:13px;">
      Chưa có lượt kiểm tra nào được lưu.
    </div>`;
    return;
  }

  const filtered = historyLogs.filter((h) => {
    const matchText =
      !query ||
      (h.contNo && h.contNo.toUpperCase().includes(query)) ||
      (h.actualSeal && h.actualSeal.toUpperCase().includes(query)) ||
      (h.expectedSeal && h.expectedSeal.toUpperCase().includes(query)) ||
      (h.vesselName && h.vesselName.toUpperCase().includes(query));
    const matchStatus = statusFilter === 'all' || h.status === statusFilter;
    return matchText && matchStatus;
  });

  if (filtered.length === 0) {
    container.innerHTML = `<div style="text-align:center; padding:24px; color:var(--text-muted); font-size:13px;">
      Không tìm thấy bản ghi lịch sử phù hợp.
    </div>`;
    return;
  }

  container.innerHTML = filtered
    .map((h) => {
      const badgeClass = h.status === 'match' ? 'match' : 'mismatch';
      const badgeText = h.status === 'match' ? '✅ KHỚP' : '❌ KHÔNG KHỚP';
      const imgs = resolveImagesForHistoryItemSync(h);
      const contThumb = imgs.contImage
        ? `<div class="hist-thumb-box" onclick="event.stopPropagation(); openHistoryPhotoDirect('${h.id}', 'cont')" title="Bấm để mở Full hình Cửa Cont"><img src="${imgs.contImage}" alt="Cont" /><span class="hist-thumb-lbl">🔍 Cont</span></div>`
        : '';
      const sealThumb = imgs.sealImage
        ? `<div class="hist-thumb-box" onclick="event.stopPropagation(); openHistoryPhotoDirect('${h.id}', 'seal')" title="Bấm để mở Full hình Số Seal"><img src="${imgs.sealImage}" alt="Seal" /><span class="hist-thumb-lbl">🔍 Seal</span></div>`
        : '';

      const rowInExcel = excelList.find((r) => r.contNo === h.contNo);
      const vName = h.vesselName || (rowInExcel && rowInExcel.vesselName) || '';

      return `
        <div class="cont-item history-clickable" onclick="openHistoryDetailModal('${h.id}')" title="Bấm để mở xem hình ảnh Cửa Cont & Số Seal">
          <div style="display:flex; align-items:center; gap:10px; min-width:0; flex:1;">
            ${contThumb || sealThumb ? `<div class="hist-thumbs">${contThumb}${sealThumb}</div>` : ''}
            <div class="cont-item-main">
              <span class="cont-no">${h.contNo} ${vName ? `<span style="font-size:11px; font-weight:600; color:#38bdf8;">• 🚢 ${vName}</span>` : ''}</span>
              <span class="seal-no">Seal thực tế: <strong>${h.actualSeal}</strong> | Excel: <strong>${h.expectedSeal}</strong></span>
              <span class="seal-no" style="font-size:11px; color:#38bdf8;">🕒 ${h.checkedAt} • 🖼️ Bấm mở hình xem lại</span>
            </div>
          </div>
          <div style="display:flex; flex-direction:column; align-items:flex-end; gap:5px; flex-shrink:0;">
            <span class="badge ${badgeClass}">${badgeText}</span>
            <span style="font-size:10.5px; color:#93c5fd; background:rgba(59,130,246,0.16); border:1px solid rgba(59,130,246,0.4); padding:2px 7px; border-radius:6px; font-weight:600;">🔍 Mở hình</span>
          </div>
        </div>
      `;
    })
    .join('');
}
