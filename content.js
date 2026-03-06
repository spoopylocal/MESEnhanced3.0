// Browser API compatibility (Chrome/Firefox)
if (typeof browser === 'undefined') {
  var browser = chrome;
}

const COPY_BUTTON_CLASS = "mes-copy-btn";
const COPY_ALL_BUTTON_CLASS = "mes-copy-all-btn";
const COPY_ALL_PN_QTY_BUTTON_CLASS = "mes-copy-all-pn-qty-btn";
const LPN_COPY_BUTTON_CLASS = "mes-copy-lpn-btn";
const LPN_LOC_COPY_BUTTON_CLASS = "mes-copy-lpn-loc-btn";
const COPY_ALL_LPN_BUTTON_CLASS = "mes-copy-all-lpn-btn";
const COPY_ALL_LPN_LOC_BUTTON_CLASS = "mes-copy-all-lpn-loc-btn";
const COPY_WO_BUTTON_CLASS = "mes-copy-wo-btn";
const COPY_ASSET_BUTTON_CLASS = "mes-copy-asset-btn";
const COPY_SR_BUTTON_CLASS = "mes-copy-sr-btn";
const WRAPPER_CLASS = "mes-pn-wrapper";
const BUTTON_SETTINGS_STORAGE_KEY = "mes-button-settings-v1";
const UPDATE_DISMISSED_KEY = "mes-update-dismissed-version";
const UPDATE_LAST_PROMPTED_KEY = "mes-update-last-prompted-day";
const DEFAULT_BUTTON_SETTINGS = {
  primaryColor: "#4b5563",
  hoverColor: "#374151",
  textColor: "#ffffff"
};

let currentButtonSettings = { ...DEFAULT_BUTTON_SETTINGS };

// Helper function to check if extension context is still valid
function isExtensionContextValid() {
  try {
    // Try to access the extension runtime
    return !!(browser && browser.runtime && browser.runtime.id);
  } catch (e) {
    return false;
  }
}

// Helper function to safely send message to background script
async function safeSendMessage(message) {
  if (!isExtensionContextValid()) {
    throw new Error('Extension has been reloaded. Please refresh this page to continue.');
  }
  
  return new Promise((resolve, reject) => {
    try {
      chrome.runtime.sendMessage(message, (response) => {
        if (chrome.runtime.lastError) {
          reject(new Error(chrome.runtime.lastError.message));
        } else {
          resolve(response);
        }
      });
    } catch (error) {
      if (error.message && error.message.includes('Extension context invalidated')) {
        reject(new Error('Extension has been reloaded. Please refresh this page to continue.'));
      } else {
        reject(error);
      }
    }
  });
}

function isValidHexColor(value) {
  return typeof value === "string" && /^#[0-9a-fA-F]{6}$/.test(value);
}

function sanitizeButtonSettings(raw) {
  const primaryColor = isValidHexColor(raw?.primaryColor)
    ? raw.primaryColor
    : DEFAULT_BUTTON_SETTINGS.primaryColor;
  const hoverColor = isValidHexColor(raw?.hoverColor)
    ? raw.hoverColor
    : DEFAULT_BUTTON_SETTINGS.hoverColor;
  const textColor = isValidHexColor(raw?.textColor)
    ? raw.textColor
    : DEFAULT_BUTTON_SETTINGS.textColor;

  return {
    primaryColor,
    hoverColor,
    textColor
  };
}

function loadButtonSettings() {
  try {
    const stored = localStorage.getItem(BUTTON_SETTINGS_STORAGE_KEY);
    if (!stored) {
      return { ...DEFAULT_BUTTON_SETTINGS };
    }

    return sanitizeButtonSettings(JSON.parse(stored));
  } catch (error) {
    console.warn("Failed to load button settings", error);
    return { ...DEFAULT_BUTTON_SETTINGS };
  }
}

function saveButtonSettings(settings) {
  try {
    localStorage.setItem(BUTTON_SETTINGS_STORAGE_KEY, JSON.stringify(settings));
  } catch (error) {
    console.warn("Failed to save button settings", error);
  }
}

function applyButtonSettings(settings) {
  const finalSettings = sanitizeButtonSettings(settings);
  const root = document.documentElement;

  root.style.setProperty("--mes-btn-bg", finalSettings.primaryColor);
  root.style.setProperty("--mes-btn-border", "transparent");
  root.style.setProperty("--mes-btn-hover", finalSettings.hoverColor);
  root.style.setProperty("--mes-btn-text", finalSettings.textColor);

  currentButtonSettings = finalSettings;
}

function getButtonSettingsFromControls() {
  const primaryInput = document.getElementById("mes-btn-primary-color");
  const hoverInput = document.getElementById("mes-btn-hover-color");
  const textInput = document.getElementById("mes-btn-text-color");

  return sanitizeButtonSettings({
    primaryColor: primaryInput?.value,
    hoverColor: hoverInput?.value,
    textColor: textInput?.value
  });
}

function syncButtonControls(settings) {
  const primaryInput = document.getElementById("mes-btn-primary-color");
  const hoverInput = document.getElementById("mes-btn-hover-color");
  const textInput = document.getElementById("mes-btn-text-color");

  if (primaryInput) {
    primaryInput.value = settings.primaryColor;
  }
  if (hoverInput) {
    hoverInput.value = settings.hoverColor;
  }
  if (textInput) {
    textInput.value = settings.textColor;
  }
}

function getDayKey(timestamp = Date.now()) {
  const d = new Date(timestamp);
  const year = d.getFullYear();
  const month = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function openUrlInNewTab(url) {
  if (!url) return;
  window.open(url, "_blank", "noopener,noreferrer");
}

function closeUpdatePopup() {
  document.getElementById("mes-update-popup")?.remove();
}

function showUpdatePopup(status) {
  if (!status?.updateAvailable || !status?.latestVersion) {
    return;
  }

  closeUpdatePopup();

  const popup = document.createElement("div");
  popup.id = "mes-update-popup";
  popup.className = "mes-update-popup";

  const notes = Array.isArray(status.notes) ? status.notes.slice(0, 5) : [];
  const notesHtml = notes.length
    ? `<ul>${notes.map((item) => `<li>${item}</li>`).join("")}</ul>`
    : `<ul><li>Performance improvements and fixes</li></ul>`;

  popup.innerHTML = `
    <div class="mes-update-title">Update Available: v${status.latestVersion}</div>
    <div class="mes-update-sub">Current version: v${status.currentVersion || "unknown"}</div>
    <div class="mes-update-notes">${notesHtml}</div>
    <div class="mes-update-actions">
      <button type="button" class="mes-update-btn" id="mes-update-now">Update</button>
      <button type="button" class="mes-update-btn secondary" id="mes-update-later">Later</button>
    </div>
  `;

  document.body.appendChild(popup);

  const updateBtn = popup.querySelector("#mes-update-now");
  const laterBtn = popup.querySelector("#mes-update-later");

  updateBtn?.addEventListener("click", () => {
    localStorage.setItem(UPDATE_DISMISSED_KEY, status.latestVersion);
    if (status.downloadUrl) {
      openUrlInNewTab(status.downloadUrl);
    }
    closeUpdatePopup();
  });

  laterBtn?.addEventListener("click", () => {
    localStorage.setItem(UPDATE_LAST_PROMPTED_KEY, getDayKey());
    closeUpdatePopup();
  });
}

function shouldPromptForUpdate(status) {
  if (!status?.updateAvailable || !status?.latestVersion) {
    return false;
  }

  const dismissedVersion = localStorage.getItem(UPDATE_DISMISSED_KEY);
  if (dismissedVersion === status.latestVersion) {
    return false;
  }

  const lastPromptedDay = localStorage.getItem(UPDATE_LAST_PROMPTED_KEY);
  return lastPromptedDay !== getDayKey();
}

async function checkForExtensionUpdates(trigger = "page-load") {
  try {
    let status = await safeSendMessage({ type: "GET_UPDATE_STATUS" });
    if (!status || !status.checkedAt) {
      status = await safeSendMessage({ type: "CHECK_FOR_UPDATES", trigger });
    }

    if (shouldPromptForUpdate(status)) {
      showUpdatePopup(status);
      localStorage.setItem(UPDATE_LAST_PROMPTED_KEY, getDayKey());
    }

    return status;
  } catch (error) {
    console.warn("Update check failed", error);
    return null;
  }
}

const isAlreadyEnhanced = (cell) =>
  cell.querySelector(`.${COPY_BUTTON_CLASS}`) !== null;

const createCopyButton = (pnText) => {
  const button = document.createElement("button");
  button.type = "button";
  button.className = COPY_BUTTON_CLASS;
  button.textContent = "+";
  button.setAttribute("aria-label", `Copy PN ${pnText}`);
  button.dataset.copy = pnText;
  return button;
};

const createCopyAllButton = () => {
  const button = document.createElement("button");
  button.type = "button";
  button.className = COPY_ALL_BUTTON_CLASS;
  button.textContent = "Copy All PNs";
  button.setAttribute("aria-label", "Copy all part numbers");
  return button;
};

const createCopyAllPnQtyButton = () => {
  const button = document.createElement("button");
  button.type = "button";
  button.className = COPY_ALL_PN_QTY_BUTTON_CLASS;
  button.textContent = "Copy All PNs + QTY";
  button.setAttribute("aria-label", "Copy all part numbers and quantities");
  return button;
};

const createCopyLpnButton = (lpnText) => {
  const button = document.createElement("button");
  button.type = "button";
  button.className = LPN_COPY_BUTTON_CLASS;
  button.textContent = "LPN";
  button.setAttribute("aria-label", `Copy LPN ${lpnText}`);
  button.dataset.copy = lpnText;
  return button;
};

const createCopyMoButton = (moText) => {
  const button = document.createElement("button");
  button.type = "button";
  button.className = LPN_COPY_BUTTON_CLASS;
  button.textContent = "MO";
  button.setAttribute("aria-label", `Copy MO ${moText}`);
  button.dataset.copy = moText;
  return button;
};

const createCopyLpnLocationButton = (locText) => {
  const button = document.createElement("button");
  button.type = "button";
  button.className = LPN_LOC_COPY_BUTTON_CLASS;
  button.textContent = "+";
  button.setAttribute("aria-label", `Copy location ${locText}`);
  button.dataset.copy = locText;
  return button;
};

const createRefreshLpnButton = () => {
  const button = document.createElement("button");
  button.type = "button";
  button.className = "mes-refresh-lpn-btn";
  button.textContent = "↻ Refresh LPNs";
  button.setAttribute("aria-label", "Refresh LPN information");
  button.addEventListener("click", refreshLpnData);
  return button;
};

const createCopyAllLpnButton = () => {
  const button = document.createElement("button");
  button.type = "button";
  button.className = COPY_ALL_LPN_BUTTON_CLASS;
  button.textContent = "Copy All LPNs";
  button.setAttribute("aria-label", "Copy all LPNs");
  return button;
};

const createCopyAllLpnLocButton = () => {
  const button = document.createElement("button");
  button.type = "button";
  button.className = COPY_ALL_LPN_LOC_BUTTON_CLASS;
  button.textContent = "Copy All LPNs + Location";
  button.setAttribute("aria-label", "Copy all LPNs and locations");
  return button;
};

const createCopyWoButton = (woText) => {
  const button = document.createElement("button");
  button.type = "button";
  button.className = COPY_WO_BUTTON_CLASS;
  button.textContent = "WO";
  button.setAttribute("aria-label", `Copy WO ${woText}`);
  button.dataset.copy = woText;
  return button;
};

const createCopyAssetButton = () => {
  const button = document.createElement("button");
  button.type = "button";
  button.className = COPY_ASSET_BUTTON_CLASS;
  button.textContent = "ASSET";
  button.setAttribute("aria-label", "Download asset attachment");
  return button;
};

const createCopySRButton = () => {
  const button = document.createElement("button");
  button.type = "button";
  button.className = COPY_SR_BUTTON_CLASS;
  button.textContent = "ST";
  button.setAttribute("aria-label", "Copy service template");
  return button;
};

const wrapPnCell = (cell) => {
  if (isAlreadyEnhanced(cell)) {
    return;
  }

  const pnText = cell.textContent?.trim();
  if (!pnText) {
    return;
  }

  const wrapper = document.createElement("span");
  wrapper.className = WRAPPER_CLASS;

  const textSpan = document.createElement("span");
  textSpan.className = "mes-pn-text";
  textSpan.textContent = pnText;

  const copyButton = createCopyButton(pnText);

  wrapper.appendChild(textSpan);
  wrapper.appendChild(copyButton);

  cell.textContent = "";
  cell.appendChild(wrapper);
};

const enhancePnCells = () => {
  const pnCells = document.querySelectorAll("td.icva-link");
  pnCells.forEach((cell) => wrapPnCell(cell));
};

const wrapLpnCell = (cell) => {
  if (cell.querySelector(`.${LPN_COPY_BUTTON_CLASS}`)) {
    return;
  }

  const lpnText = cell.textContent?.trim();
  if (!lpnText) {
    return;
  }

  const wrapper = document.createElement("span");
  wrapper.className = WRAPPER_CLASS;

  const textSpan = document.createElement("span");
  textSpan.className = "mes-pn-text";
  textSpan.textContent = lpnText;

  const moMatch = lpnText.split("-").find((part) => part.startsWith("MO"));
  const moText = moMatch || lpnText;
  const copyMoButton = createCopyMoButton(moText);
  const copyButton = createCopyLpnButton(lpnText);

  wrapper.appendChild(textSpan);
  wrapper.appendChild(copyMoButton);
  wrapper.appendChild(copyButton);

  cell.textContent = "";
  cell.appendChild(wrapper);
};

const normalizeLocationText = (text) => text.replace(/\.\.+$/g, "").trim();

const wrapLpnLocationCell = (cell) => {
  if (cell.querySelector(`.${LPN_LOC_COPY_BUTTON_CLASS}`)) {
    return;
  }

  const locTextRaw = cell.textContent?.trim();
  if (!locTextRaw) {
    return;
  }

  const locText = normalizeLocationText(locTextRaw);

  const wrapper = document.createElement("span");
  wrapper.className = WRAPPER_CLASS;

  const textSpan = document.createElement("span");
  textSpan.className = "mes-pn-text";
  textSpan.textContent = locTextRaw;

  const copyButton = createCopyLpnLocationButton(locText);

  wrapper.appendChild(textSpan);
  wrapper.appendChild(copyButton);

  cell.textContent = "";
  cell.appendChild(wrapper);
};

const enhanceLpnTable = () => {
  const lpnRows = document.querySelectorAll("#wo-lpn-body tr");
  lpnRows.forEach((row) => {
    const lpnCell = row.querySelector("td:nth-child(1)");
    const locCell = row.querySelector("td:nth-child(2)");
    if (lpnCell) {
      wrapLpnCell(lpnCell);
    }
    if (locCell) {
      wrapLpnLocationCell(locCell);
    }
  });
};

const collectAllPnValues = () => {
  const pnCells = document.querySelectorAll("td.icva-link");
  return Array.from(pnCells)
    .map((cell) => {
      const text = cell.querySelector(".mes-pn-text")?.textContent;
      return text?.trim() || cell.textContent?.trim();
    })
    .filter(Boolean);
};

const collectAllPnQtyValues = () => {
  const pnCells = document.querySelectorAll("td.icva-link");
  return Array.from(pnCells)
    .map((cell) => {
      const pnText = cell.querySelector(".mes-pn-text")?.textContent?.trim() ||
        cell.textContent?.trim();
      const row = cell.closest("tr");
      const qtyCell = row ? row.querySelector("td:nth-child(3)") : null;
      const qtyRaw = qtyCell?.textContent?.trim();
      const qtyText = qtyRaw ? qtyRaw.split("/")[0].trim() : null;
      if (!pnText || !qtyText) {
        return null;
      }
      return `PN: ${pnText} QTY: ${qtyText}`;
    })
    .filter(Boolean);
};

const collectAllLpnValues = () => {
  const lpnCells = document.querySelectorAll("#wo-lpn-body td:nth-child(1)");
  return Array.from(lpnCells)
    .map((cell) => cell.querySelector(".mes-pn-text")?.textContent?.trim() || cell.textContent?.trim())
    .filter(Boolean);
};

const collectAllLpnLocationValues = () => {
  const lpnRows = document.querySelectorAll("#wo-lpn-body tr");
  return Array.from(lpnRows)
    .map((row) => {
      const lpnCell = row.querySelector("td:nth-child(1)");
      const locCell = row.querySelector("td:nth-child(2)");
      const lpnText = lpnCell?.querySelector(".mes-pn-text")?.textContent?.trim() || lpnCell?.textContent?.trim();
      const locTextRaw = locCell?.querySelector(".mes-pn-text")?.textContent?.trim() || locCell?.textContent?.trim();
      const locText = locTextRaw ? normalizeLocationText(locTextRaw) : "";
      if (!lpnText || !locText) {
        return null;
      }
      return `LPN: ${lpnText} LOCATION: ${locText}`;
    })
    .filter(Boolean);
};

const ensureCopyAllButton = () => {
  if (document.querySelector(`.${COPY_ALL_BUTTON_CLASS}`)) {
    return;
  }

  const firstPnCell = document.querySelector("td.icva-link");
  if (!firstPnCell) {
    return;
  }

  const table = firstPnCell.closest("table");
  if (!table) {
    return;
  }

  const thead = table.querySelector("thead");
  if (!thead) {
    return;
  }

  const headerCells = Array.from(thead.querySelectorAll("th"));
  const targetHeader = headerCells.find((cell) =>
    /\bitem\b/i.test(cell.textContent || "")
  );

  if (!targetHeader) {
    return;
  }

  const button = createCopyAllButton();
  const buttonPnQty = createCopyAllPnQtyButton();
  targetHeader.appendChild(button);
  targetHeader.appendChild(buttonPnQty);
};

const ensureLpnHeaderButtons = () => {
  if (document.querySelector(`.${COPY_ALL_LPN_BUTTON_CLASS}`)) {
    return;
  }

  const lpnHeader = document.querySelector("#wo-lpn-p");
  if (!lpnHeader) {
    return;
  }

  const headerContainer = lpnHeader.closest("div");
  const rightContainer = headerContainer?.querySelector(".flex.items-end") || headerContainer?.lastElementChild;
  if (!rightContainer || !(rightContainer instanceof HTMLElement)) {
    return;
  }

  const wrapper = document.createElement("span");
  wrapper.className = "mes-lpn-header-actions";

  const refreshBtn = createRefreshLpnButton();
  const copyAll = createCopyAllLpnButton();
  const copyAllLoc = createCopyAllLpnLocButton();

  wrapper.appendChild(refreshBtn);
  wrapper.appendChild(copyAll);
  wrapper.appendChild(copyAllLoc);
  rightContainer.appendChild(wrapper);
};

const refreshLpnData = async () => {
  const workOrderId = getWorkOrderId();
  if (!workOrderId) {
    console.error('Could not find work order ID');
    return;
  }

  const refreshBtn = document.querySelector('.mes-refresh-lpn-btn');
  if (refreshBtn) {
    refreshBtn.disabled = true;
    refreshBtn.textContent = '↻ Refreshing...';
  }

  try {
    const response = await fetch(
      `https://apirouter.apps.wwt.com/api/forward/mes-api/workOrderInfo?jobHeaderId=${workOrderId}`,
      {
        method: 'GET',
        credentials: 'include',
        headers: {
          'Accept': 'application/json',
          'Content-Type': 'application/json'
        }
      }
    );

    if (!response.ok) {
      throw new Error(`Failed to fetch work order data: ${response.status}`);
    }

    const workOrderData = await response.json();
    const lpnData = workOrderData.onHandLPNs || [];

    // Update the LPN table
    const lpnTableBody = document.querySelector('#wo-lpn-body');
    if (lpnTableBody) {
      // Clear existing rows
      lpnTableBody.innerHTML = '';

      // Add new rows
      lpnData.forEach(lpn => {
        const row = document.createElement('tr');
        row.setAttribute('data-v-512c6e58', '');

        // LPN cell
        const lpnCell = document.createElement('td');
        lpnCell.setAttribute('data-v-512c6e58', '');
        lpnCell.textContent = lpn.licensePlateNumber;
        row.appendChild(lpnCell);

        // Location cell
        const locCell = document.createElement('td');
        locCell.setAttribute('data-v-512c6e58', '');
        locCell.textContent = lpn.location || '';
        row.appendChild(locCell);

        lpnTableBody.appendChild(row);
      });

      // Re-enhance the table with copy buttons
      enhanceLpnTable();
    }

    if (refreshBtn) {
      refreshBtn.textContent = '✓ Refreshed!';
      setTimeout(() => {
        refreshBtn.textContent = '↻ Refresh LPNs';
        refreshBtn.disabled = false;
      }, 2000);
    }
  } catch (error) {
    console.error('Error refreshing LPN data:', error);
    if (refreshBtn) {
      refreshBtn.textContent = '✗ Error';
      refreshBtn.classList.add('is-error');
      setTimeout(() => {
        refreshBtn.textContent = '↻ Refresh LPNs';
        refreshBtn.classList.remove('is-error');
        refreshBtn.disabled = false;
      }, 2000);
    }
  }
};

const getWorkOrderId = () => {
  const match = window.location.pathname.match(/\/orders\/(\d+)/);
  return match ? match[1] : null;
};

const ensureWoCopyButton = () => {
  const header = document.querySelector("#order-details-header-h3");
  if (!header) {
    return;
  }

  const headerText = header.textContent || "";
  const match = headerText.match(/WO#\s*(\d+)/i);
  if (!match) {
    return;
  }

  const woText = match[1];

  if (!header.querySelector(`.${COPY_WO_BUTTON_CLASS}`)) {
    const button = createCopyWoButton(woText);
    button.classList.add("mes-wo-inline");
    header.appendChild(button);
  }

  if (!header.querySelector(`.${COPY_ASSET_BUTTON_CLASS}`)) {
    const assetButton = createCopyAssetButton();
    assetButton.classList.add("mes-wo-inline");
    header.appendChild(assetButton);
  }

  if (!header.querySelector(`.${COPY_SR_BUTTON_CLASS}`)) {
    const srButton = createCopySRButton();
    srButton.classList.add("mes-wo-inline");
    header.appendChild(srButton);
  }
};

const fetchAttachmentIds = async (orderId) => {
  const url = `https://apirouter.apps.wwt.com/api/forward/associations/mes2-work-order/${orderId}?type=attachment&embedded=false&pageSize=100&includeSecure=true`;
  const response = await fetch(url, { credentials: "include" });
  if (!response.ok) {
    return [];
  }

  const payload = await response.json();
  const data = payload?.attachment?.data?.data || [];
  return data
    .map((item) => item?.resources?.find((res) => res.type === "attachment")?.id)
    .filter(Boolean);
};

const fetchAttachments = async (ids) => {
  if (ids.length === 0) {
    return [];
  }

  const query = ids.join(",");
  const url = `https://apirouter.apps.wwt.com/api/forward/attachments?ids=${encodeURIComponent(query)}&includeSecure=true&pagination=false`;
  const response = await fetch(url, { credentials: "include" });
  if (!response.ok) {
    return [];
  }

  const payload = await response.json();
  return payload?.data || [];
};

const findMatchingAttachment = (attachments, woText) => {
  const normalized = attachments.map((item) => ({
    item,
    name: (item?.fileName || "").toLowerCase(),
  }));

  const qcKeywords = ["asset qc", "qc"].filter(Boolean);
  const assetKeywords = ["asset sheet", "asset"].filter(Boolean);

  const qcMatch = normalized.find(({ name }) =>
    qcKeywords.some((keyword) => name.includes(keyword))
  );
  if (qcMatch) {
    return qcMatch.item;
  }

  const assetMatch = normalized.find(({ name }) =>
    assetKeywords.some((keyword) => name.includes(keyword))
  );
  return assetMatch ? assetMatch.item : null;
};

const downloadAttachment = async (attachment) => {
  if (!attachment?.url) {
    return;
  }

  const filename = attachment.fileName || undefined;
  try {
    await safeSendMessage({
      type: "download",
      url: attachment.url,
      filename,
    });
  } catch (e) {
    console.error('Download message error:', e);
    if (e.message && e.message.includes('refresh this page')) {
      alert(e.message);
    }
  }
};

const copyText = async (text) => {
  if (navigator.clipboard?.writeText) {
    await navigator.clipboard.writeText(text);
    return true;
  }

  const tempInput = document.createElement("textarea");
  tempInput.value = text;
  tempInput.setAttribute("readonly", "");
  tempInput.style.position = "absolute";
  tempInput.style.left = "-9999px";
  document.body.appendChild(tempInput);
  tempInput.select();
  const success = document.execCommand("copy");
  document.body.removeChild(tempInput);
  return success;
};

const handleCopyClick = async (event) => {
  const target = event.target;
  if (!(target instanceof HTMLElement)) {
    return;
  }

  if (
    !target.classList.contains(COPY_BUTTON_CLASS) &&
    !target.classList.contains(COPY_ALL_BUTTON_CLASS) &&
    !target.classList.contains(COPY_ALL_PN_QTY_BUTTON_CLASS) &&
    !target.classList.contains(LPN_COPY_BUTTON_CLASS) &&
    !target.classList.contains(LPN_LOC_COPY_BUTTON_CLASS) &&
    !target.classList.contains(COPY_ALL_LPN_BUTTON_CLASS) &&
    !target.classList.contains(COPY_ALL_LPN_LOC_BUTTON_CLASS) &&
    !target.classList.contains(COPY_WO_BUTTON_CLASS) &&
    !target.classList.contains(COPY_ASSET_BUTTON_CLASS) &&
    !target.classList.contains(COPY_SR_BUTTON_CLASS)
  ) {
    return;
  }

  event.preventDefault();
  event.stopPropagation();

  const isCopyAll = target.classList.contains(COPY_ALL_BUTTON_CLASS);
  const isCopyAllPnQty = target.classList.contains(
    COPY_ALL_PN_QTY_BUTTON_CLASS
  );
  const isCopyAllLpn = target.classList.contains(COPY_ALL_LPN_BUTTON_CLASS);
  const isCopyAllLpnLoc = target.classList.contains(
    COPY_ALL_LPN_LOC_BUTTON_CLASS
  );
  const isCopyWo = target.classList.contains(COPY_WO_BUTTON_CLASS);
  const isCopyAsset = target.classList.contains(COPY_ASSET_BUTTON_CLASS);
  const isCopySR = target.classList.contains(COPY_SR_BUTTON_CLASS);
  const pnText = target.dataset.copy;
  const values = isCopyAll
    ? collectAllPnValues()
    : isCopyAllPnQty
      ? collectAllPnQtyValues()
      : isCopyAllLpn
        ? collectAllLpnValues()
        : isCopyAllLpnLoc
          ? collectAllLpnLocationValues()
          : isCopyWo
            ? pnText
              ? [pnText]
              : []
            : isCopyAsset
              ? []
              : isCopySR
                ? []
                : pnText
                  ? [pnText]
                  : [];
  if (values.length === 0) {
    if (!isCopyAsset && !isCopySR) {
      return;
    }
  }

  try {
    if (isCopyAsset) {
      const orderId = getWorkOrderId();
      const header = document.querySelector("#order-details-header-h3");
      const headerText = header?.textContent || "";
      const woMatch = headerText.match(/WO#\s*(\d+)/i);
      const woText = woMatch ? woMatch[1] : null;
      if (!orderId) {
        alert('Could not find work order ID');
        return;
      }

      // Show loading state
      const originalText = target.textContent;
      target.textContent = "⏳";
      target.disabled = true;

      try {
        const ids = await fetchAttachmentIds(orderId);
        const attachments = await fetchAttachments(ids);
        const match = findMatchingAttachment(attachments, woText);
        
        target.disabled = false;
        
        if (match) {
          try {
            await downloadAttachment(match);
            target.textContent = "✓";
            target.classList.add("is-copied");
            window.setTimeout(() => {
              target.textContent = originalText || "ASSET";
              target.classList.remove("is-copied");
            }, 1400);
          } catch (downloadError) {
            console.error('Download error:', downloadError);
            target.textContent = "✗";
            target.classList.add("is-error");
            alert('Failed to download asset sheet.');
            window.setTimeout(() => {
              target.textContent = originalText || "ASSET";
              target.classList.remove("is-error");
            }, 1400);
          }
        } else {
          target.textContent = "✗";
          target.classList.add("is-error");
          alert('Asset sheet not found. Please check attachments manually.');
          window.setTimeout(() => {
            target.textContent = originalText || "ASSET";
            target.classList.remove("is-error");
          }, 1400);
        }
      } catch (fetchError) {
        console.error('Fetch attachments error:', fetchError);
        target.disabled = false;
        target.textContent = "✗";
        target.classList.add("is-error");
        alert('Failed to fetch attachments. Please try again.');
        window.setTimeout(() => {
          target.textContent = originalText || "ASSET";
          target.classList.remove("is-error");
        }, 1400);
      }
      return;
    } else if (isCopySR) {
      const orderId = getWorkOrderId();
      if (!orderId) {
        alert('Could not find work order ID');
        return;
      }
      
      // Show loading state
      const originalText = target.textContent;
      target.textContent = "⏳";
      target.disabled = true;
      
      try {
        const response = await fetch(
          `https://apirouter.apps.wwt.com/api/forward/mes-api/workOrderInfo?jobHeaderId=${orderId}`,
          {
            method: 'GET',
            credentials: 'include',
            headers: {
              'Accept': 'application/json',
              'Content-Type': 'application/json'
            }
          }
        );
        
        target.disabled = false;
        
        if (!response.ok) {
          target.textContent = "✗";
          target.classList.add("is-error");
          alert('Failed to fetch work order information');
          setTimeout(() => {
            target.textContent = originalText || "ST";
            target.classList.remove("is-error");
          }, 1400);
          return;
        }
        
        const workOrderData = await response.json();
        const serviceTemplate = workOrderData.serviceTemplateName || '';
        
        if (!serviceTemplate) {
          target.textContent = "✗";
          target.classList.add("is-error");
          alert('No service template found for this work order');
          setTimeout(() => {
            target.textContent = originalText || "ST";
            target.classList.remove("is-error");
          }, 1400);
          return;
        }
        
        const success = await copyText(serviceTemplate);
        if (!success) {
          target.textContent = "✗";
          target.classList.add("is-error");
          alert('Failed to copy service template to clipboard');
          setTimeout(() => {
            target.textContent = originalText || "ST";
            target.classList.remove("is-error");
          }, 1400);
          return;
        }
        
        // Success - show checkmark
        target.textContent = "✓";
        target.classList.add("is-copied");
        setTimeout(() => {
          target.textContent = originalText || "ST";
          target.classList.remove("is-copied");
        }, 1400);
      } catch (error) {
        console.error('SR fetch error:', error);
        target.disabled = false;
        target.textContent = "✗";
        target.classList.add("is-error");
        alert('Network error. Please try again.');
        setTimeout(() => {
          target.textContent = originalText || "ST";
          target.classList.remove("is-error");
        }, 1400);
      }
      return;
    } else {
      const success = await copyText(values.join("\n"));
      if (!success) {
        return;
      }
    }

    const originalText = target.textContent;
    target.textContent = "✓";
    target.classList.add("is-copied");
    window.setTimeout(() => {
      target.textContent =
        originalText ??
        (isCopyAll
          ? "Copy All PNs"
          : isCopyAllPnQty
            ? "Copy All PNs + QTY"
            : isCopyAllLpn
              ? "Copy All LPNs"
              : isCopyAllLpnLoc
                ? "Copy All LPNs + Location"
                : isCopyWo
                  ? "WO"
                  : isCopyAsset
                    ? "ASSET"
                    : isCopySR
                      ? "ST"
                      : "+");
      target.classList.remove("is-copied");
    }, 1400);
  } catch (error) {
    console.error("MES PN copy failed", error);
  }
};

// ===== TOOLS MENU =====
const createToolsMenu = () => {
  const existing = document.getElementById('mes-tools-menu');
  if (existing) return;

  const headerRight = document.querySelector('.header-right');
  if (!headerRight) return;

  const menuContainer = document.createElement('details');
  menuContainer.id = 'mes-tools-menu';
  menuContainer.className = 'mes-tools-menu';

  const summary = document.createElement('summary');
  summary.className = 'mes-tools-summary';
  summary.innerHTML = `
    <span class="mes-tools-icon">
      <svg width="16" height="16" viewBox="0 0 16 16" fill="currentColor" xmlns="http://www.w3.org/2000/svg">
        <path fill-rule="evenodd" d="M8 4.754a3.246 3.246 0 1 0 0 6.492 3.246 3.246 0 0 0 0-6.492zM5.754 8a2.246 2.246 0 1 1 4.492 0 2.246 2.246 0 0 1-4.492 0z"/>
        <path fill-rule="evenodd" d="M9.796 1.343c-.527-1.79-3.065-1.79-3.592 0l-.094.319a.873.873 0 0 1-1.255.52l-.292-.16c-1.64-.892-3.433.902-2.54 2.541l.159.292a.873.873 0 0 1-.52 1.255l-.319.094c-1.79.527-1.79 3.065 0 3.592l.319.094a.873.873 0 0 1 .52 1.255l-.16.292c-.892 1.64.901 3.434 2.541 2.54l.292-.159a.873.873 0 0 1 1.255.52l.094.319c.527 1.79 3.065 1.79 3.592 0l.094-.319a.873.873 0 0 1 1.255-.52l.292.16c1.64.893 3.434-.902 2.54-2.541l-.159-.292a.873.873 0 0 1 .52-1.255l.319-.094c1.79-.527 1.79-3.065 0-3.592l-.319-.094a.873.873 0 0 1-.52-1.255l.16-.292c.893-1.64-.902-3.433-2.541-2.54l-.292.159a.873.873 0 0 1-1.255-.52l-.094-.319z"/>
      </svg>
    </span>
    <span class="mes-tools-text">Tools</span>
  `;

  const dropdown = document.createElement('div');
  dropdown.className = 'mes-tools-dropdown';
  dropdown.innerHTML = `
    <div class="mes-tools-list">
      <div class="mes-tools-label">— MO Tools —</div>
      <button class="mes-tool-item" data-tool="mo-auto-prm">MO PRM/INFO</button>
      <button class="mes-tool-item" data-tool="mo-tracker">MO STATUS (Not Working)</button>
      <div class="mes-tools-label">— Sheet Updater —</div>
      <button class="mes-tool-item" data-tool="ib-update-tracker">IB Update Tracker</button>
      <div class="mes-tools-label">— Reportal —</div>
      <button class="mes-tool-item" data-tool="inventory-reportal">Inventory Reportal</button>
      <button class="mes-tool-item" data-tool="outbound-reportal">Outbound Reportal</button>
      <div class="mes-tools-label">— Serial Checkers —</div>
      <button class="mes-tool-item" data-tool="serial-scan-lsc">LSC Serial Checker</button>
      <button class="mes-tool-item" data-tool="serial-scan-ibc">IBC Serial Checker</button>
      <button class="mes-tool-item" data-tool="check-updates">Check for Updates</button>
      <div class="mes-tools-label">— Button Style —</div>
      <div class="mes-button-customization">
        <label class="mes-customization-row" for="mes-btn-primary-color">
          <span>Primary</span>
          <input id="mes-btn-primary-color" type="color" />
        </label>
        <label class="mes-customization-row" for="mes-btn-hover-color">
          <span>Hover</span>
          <input id="mes-btn-hover-color" type="color" />
        </label>
        <label class="mes-customization-row" for="mes-btn-text-color">
          <span>Text</span>
          <input id="mes-btn-text-color" type="color" />
        </label>
        <div class="mes-customization-actions">
          <button type="button" class="mes-customization-btn" id="mes-btn-style-apply">Apply</button>
          <button type="button" class="mes-customization-btn" id="mes-btn-style-reset">Reset</button>
        </div>
      </div>
    </div>
  `;

  menuContainer.appendChild(summary);
  menuContainer.appendChild(dropdown);

  const aiIcon = headerRight.querySelector('.bs-icon.cursor-pointer');
  // Firefox: Verify element is actually a child before using insertBefore
  if (aiIcon && aiIcon.parentNode === headerRight) {
    headerRight.insertBefore(menuContainer, aiIcon);
  } else if (headerRight.firstChild && headerRight.firstChild.parentNode === headerRight) {
    headerRight.insertBefore(menuContainer, headerRight.firstChild);
  } else {
    // Fallback to appendChild if insertBefore conditions not met
    headerRight.appendChild(menuContainer);
  }

  syncButtonControls(currentButtonSettings);
};

// ===== MO AUTO PRM/COPY SCRIPT =====
const moAutoPrmCopy = async () => {
  const BASE_MES = 'https://apirouter.apps.wwt.com/api/forward/mes-api';
  const BASE_PRM_BUILD = 'https://apirouter.apps.wwt.com/prm-api';
  const BASE_PRM_SEARCH = 'https://apirouter.apps.wwt.com/api/forward/prm-api';
  const TZ = 'America/Chicago';

  const listPayload = {
    limit: 1000,
    offset: 0,
    releasedOnly: true,
    includeCompPercent: true,
    includeLocations: true,
    includeOrderManager: true,
    facilityCode: 'WPC',
    showUnassignedOnly: false,
    orderBy: [
      { orderField: 'priority', orderDirection: 'ASC' },
      { orderField: 'criticalRatio', orderDirection: 'ASC' },
      { orderField: 'documentNumber', orderDirection: 'ASC' }
    ],
    labArea: 'L4',
    dynamicFilters: [
      {
        name: 'labDestinationName',
        displayAs: 'Lab Destination',
        type: 'dynamicLov',
        dataType: 'string',
        operator: '%',
        availableOperators: ['=', '?', '%'],
        lovName: 'labDestinations',
        values: ['%naic1%']
      }
    ]
  };

  function showToast(message, success = true) {
    const existing = document.getElementById('mo-toast-wwt');
    if (existing) existing.remove();
    const t = document.createElement('div');
    t.id = 'mo-toast-wwt';
    t.textContent = message;
    t.style.cssText = `
      position: fixed;
      bottom: 16px;
      right: 16px;
      padding: 6px 10px;
      background: ${success ? '#2e7d32' : '#c62828'};
      color: #fff;
      font-size: 12px;
      font-family: system-ui,-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;
      border-radius: 4px;
      z-index: 99999;
      box-shadow: 0 2px 6px rgba(0,0,0,0.3);
      opacity: 0;
      transition: opacity .2s ease;
    `;
    document.body.appendChild(t);
    requestAnimationFrame(() => { t.style.opacity = '1'; });
    setTimeout(() => {
      t.style.opacity = '0';
      setTimeout(() => t.remove(), 200);
    }, 2500);
  }

  function formatBuildDueDate(info) {
    const ops = Array.isArray(info.operations) ? info.operations : [];
    const allServices = ops.flatMap(op => Array.isArray(op.jobServices) ? op.jobServices : []);
    const buildSvc = allServices.find(svc => svc.serviceName === 'Build');
    if (!buildSvc) return 'N/A';
    const end = buildSvc.serviceEndDate || buildSvc.serviceStartDate;
    if (!end) return 'N/A';
    const d = new Date(end);
    const day = d.toLocaleString('en-US', { day: 'numeric', timeZone: TZ });
    const month = d.toLocaleString('en-US', { month: 'short', timeZone: TZ });
    return `${day}-${month}`;
  }

  async function fetchJson(url, opts = {}) {
    const res = await fetch(url, {
      credentials: 'include',
      headers: { Accept: 'application/json', ...(opts.headers || {}) },
      ...opts
    });
    if (!res.ok) throw new Error(`HTTP ${res.status} for ${url}`);
    const text = await res.text();
    if (text.trim().startsWith('<!DOCTYPE') || text.trim().startsWith('<html'))
      throw new Error(`Got HTML instead of JSON from ${url}`);
    return JSON.parse(text);
  }

  async function updateWorkOrderHeader(payload) {
    const res = await fetch(`${BASE_MES}/updateWorkOrderHeader`, {
      method: 'PATCH',
      credentials: 'include',
      headers: { 'Content-Type': 'application/json', Accept: 'application/json' },
      body: JSON.stringify(payload)
    });
    if (!res.ok) {
      const body = await res.text().catch(() => '');
      throw new Error(`updateWorkOrderHeader HTTP ${res.status}: ${body}`);
    }
  }

  function extractMsfCode(wfg) {
    if (!wfg) return null;
    const firstPart = String(wfg).split('.')[0];
    if (/^MSF-\d+$/i.test(firstPart)) return firstPart;
    const m = String(wfg).match(/MSF-\d+/i);
    return m ? m[0] : null;
  }

  async function getMostRecentPrmIdFromMsf(msfCode) {
    const url = `${BASE_PRM_SEARCH}/buildConfigVersion/search?prmname=${encodeURIComponent(msfCode)}&status=Published`;
    const arr = await fetchJson(url);
    if (!Array.isArray(arr) || arr.length === 0) return null;
    const published = arr.filter(x => x.status?.status === 'Published');
    if (published.length === 0) return null;
    let best = published[0];
    for (const item of published.slice(1)) {
      if ((item.version || 0) > (best.version || 0)) {
        best = item;
      }
    }
    return best.id || null;
  }

  function parseMoInput(raw) {
    const s = String(raw || '').trim();
    if (!s) return [];
    const matches = s.match(/MO\d+/gi) || [];
    const uniq = [];
    const seen = new Set();
    for (const m of matches) {
      const mo = m.toUpperCase();
      if (!seen.has(mo)) {
        seen.add(mo);
        uniq.push(mo);
      }
    }
    return uniq;
  }

  function woHasMo(wo, mo) {
    const s = String(wo?.moveOrderNumbers || '');
    if (!s) return false;
    return s.split(',').map(x => x.trim()).filter(Boolean).includes(mo);
  }

  async function buildLineForMo(mo, wo) {
    if (!wo) return { mo, line: `${mo}\t\t(No work order found)\t\tN/A`, ok: false };
    
    let { documentNumber, jobHeaderId, prmId } = wo;
    
    if (!prmId) {
      showToast(`No PRM on ${mo}, trying to add…`, false);
      const msfCode = extractMsfCode(wo.wipFinishedGood);
      if (!msfCode) {
        return { mo, line: `${mo}\t${documentNumber || ''}\t(Could not find MSF code)\t\tN/A`, ok: false };
      }
      
      const newPrmId = await getMostRecentPrmIdFromMsf(msfCode);
      if (!newPrmId) {
        return { mo, line: `${mo}\t${documentNumber || ''}\t(No Published PRM for ${msfCode})\t\tN/A`, ok: false };
      }
      
      const updatePayload = {
        orderTypeCode: wo.orderTypeCode,
        siteProject: wo.siteProject ?? null,
        prodStatus: wo.prodStatus,
        prodAssignmentGrp: wo.prodAssignmentGrp ?? null,
        priority: wo.priority,
        delayReasonCode: wo.delayReasonCode ?? null,
        delayResponsibilityCode: wo.delayResponsibilityCode ?? null,
        rootCauseCode: wo.rootCauseCode ?? null,
        expediteFlag: wo.expediteFlag ?? 'N',
        btsFlag: wo.btsFlag ?? null,
        crateOrderedFlag: wo.crateOrderedFlag ?? 'N',
        workstations: wo.workstations ?? null,
        rootCauseDescription: wo.rootCauseDescription ?? null,
        jobHeaderId: wo.jobHeaderId,
        prmId: newPrmId
      };
      
      await updateWorkOrderHeader(updatePayload);
      wo.prmId = newPrmId;
      prmId = newPrmId;
      showToast(`PRM set for ${mo}`, true);
    }
    
    const prmJson = await fetchJson(`${BASE_PRM_BUILD}/buildConfigVersion/${encodeURIComponent(prmId)}`);
    const buildName = prmJson.buildConfig?.buildName || prmJson.buildConfig?.prmName || '(Build name not found)';
    
    const woInfo = await fetchJson(`${BASE_MES}/workOrderInfo?jobHeaderId=${encodeURIComponent(jobHeaderId)}`);
    const buildDue = formatBuildDueDate(woInfo);
    
    const line = `${mo}\t${documentNumber || ''}\t${buildName}\t\t${buildDue}`;
    return { mo, line, ok: true };
  }

  const raw = prompt('Enter MO numbers (comma / space / newline). Example:\nMO156913878, MO156913879');
  const mos = parseMoInput(raw);
  
  if (mos.length === 0) {
    showToast('No valid MOs found', false);
    return;
  }
  
  showToast(`Loading ${mos.length} MO(s)…`, true);
  
  const list = await fetchJson(`${BASE_MES}/workOrders`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(listPayload)
  });
  
  if (!Array.isArray(list)) {
    showToast('Unexpected workOrders response', false);
    return;
  }
  
  const woByMo = new Map();
  for (const mo of mos) {
    const wo = list.find(o => woHasMo(o, mo));
    if (wo) woByMo.set(mo, wo);
  }
  
  const results = [];
  for (let i = 0; i < mos.length; i++) {
    const mo = mos[i];
    try {
      const wo = woByMo.get(mo) || null;
      const r = await buildLineForMo(mo, wo);
      results.push(r);
    } catch (err) {
      console.error('Error on', mo, err);
      results.push({ mo, line: `${mo}\t\t(Error – see console)\t\tN/A`, ok: false });
    }
  }
  
  const out = results.map(r => r.line).join('\n');
  console.log(out);
  
  try {
    await navigator.clipboard.writeText(out);
    showToast(`Copied ${results.length} line(s)`, true);
  } catch (e) {
    console.warn('Could not copy to clipboard:', e);
    showToast('Copy failed – see console', false);
  }
  
  console.table(results.map(r => ({ MO: r.mo, OK: r.ok, Line: r.line })));
};

// ===== MO TRACKER STATUS UPDATER SCRIPT =====
const moTrackerStatusUpdater = async () => {
  const BASE_MES = 'https://apirouter.apps.wwt.com/api/forward/mes-api';
  const BASE_PRM = 'https://apirouter.apps.wwt.com/prm-api';

  const listPayload = {
    limit: 1000,
    offset: 0,
    releasedOnly: true,
    includeCompPercent: true,
    includeLocations: true,
    includeOrderManager: true,
    facilityCode: 'WPC',
    showUnassignedOnly: false,
    orderBy: [
      { orderField: 'priority', orderDirection: 'ASC' },
      { orderField: 'criticalRatio', orderDirection: 'ASC' },
      { orderField: 'documentNumber', orderDirection: 'ASC' }
    ],
    labArea: 'L4',
    dynamicFilters: [
      {
        name: 'labDestinationName',
        displayAs: 'Lab Destination',
        type: 'dynamicLov',
        dataType: 'string',
        operator: '%',
        availableOperators: ['=', '?', '%'],
        lovName: 'labDestinations',
        values: ['%naic1%']
      }
    ]
  };

  function showToast(message, success = true) {
    const existing = document.getElementById('mo-toast-wwt');
    if (existing) existing.remove();
    const t = document.createElement('div');
    t.id = 'mo-toast-wwt';
    t.textContent = message;
    t.style.cssText = `
      position: fixed;
      bottom: 16px;
      right: 16px;
      padding: 6px 10px;
      background: ${success ? '#2e7d32' : '#c62828'};
      color: #fff;
      font-size: 12px;
      font-family: system-ui,-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;
      border-radius: 4px;
      z-index: 99999;
      box-shadow: 0 2px 6px rgba(0,0,0,0.3);
      opacity: 0;
      transition: opacity .2s ease;
    `;
    document.body.appendChild(t);
    requestAnimationFrame(() => { t.style.opacity = '1'; });
    setTimeout(() => {
      t.style.opacity = '0';
      setTimeout(() => t.remove(), 200);
    }, 2500);
  }

  async function fetchJson(url, opts = {}) {
    const res = await fetch(url, {
      credentials: 'include',
      headers: { Accept: 'application/json', ...(opts.headers || {}) },
      ...opts
    });
    if (!res.ok) throw new Error(`HTTP ${res.status} for ${url}`);
    const text = await res.text();
    if (text.trim().startsWith('<!DOCTYPE') || text.trim().startsWith('<html')) {
      throw new Error(`Got HTML instead of JSON from ${url}`);
    }
    return JSON.parse(text);
  }

  function isInDrpLocation(info) {
    const lpns = Array.isArray(info.onHandLPNs) ? info.onHandLPNs : [];
    for (const lpn of lpns) {
      const loc = String(lpn.location || '');
      if (loc.startsWith('ZL4MMTDRP') || loc.startsWith('ZL4XDR')) {
        return true;
      }
    }
    return false;
  }

  function hasAnyLpn(info) {
    const lpns = Array.isArray(info.onHandLPNs) ? info.onHandLPNs : [];
    return lpns.length > 0;
  }

  function getFirstLocation(info) {
    const lpns = Array.isArray(info.onHandLPNs) ? info.onHandLPNs : [];
    if (!lpns.length) return '';
    return String(lpns[0].location || '');
  }

  function getNestedStatus(info, opName, svcName) {
    const ops = Array.isArray(info.operations) ? info.operations : [];
    for (const op of ops) {
      if (op.description !== opName) continue;
      const svcs = Array.isArray(op.jobServices) ? op.jobServices : [];
      for (const svc of svcs) {
        if (svc.serviceName === svcName) {
          return svc.serviceStatus || 'N/A';
        }
      }
    }
    return 'N/A';
  }

  function analyzeBuild(info) {
    const ops = Array.isArray(info.operations) ? info.operations : [];
    let buildSvc = null;
    let buildOpSeq = null;

    for (const op of ops) {
      if (!Array.isArray(op.jobServices)) continue;
      for (const svc of op.jobServices) {
        if (svc.serviceName === 'Build') {
          buildSvc = svc;
          buildOpSeq = op.operationSeqNum;
          break;
        }
      }
      if (buildSvc) break;
    }

    const buildStatus = buildSvc?.serviceStatus || 'N/A';
    const buildStatusU = (buildStatus || '').toUpperCase();

    const buildStarted =
      buildStatusU === 'WORKING' ||
      buildStatusU === 'COMPLETE' ||
      !!buildSvc?.actualStartDate ||
      !!buildSvc?.serviceStartDate;

    let pastBuild = false;
    if (buildOpSeq != null) {
      for (const op of ops) {
        if (typeof op.operationSeqNum !== 'number') continue;
        if (op.operationSeqNum <= buildOpSeq) continue;
        if (!Array.isArray(op.jobServices)) continue;
        for (const svc of op.jobServices) {
          const st = (svc.serviceStatus || '').toUpperCase();
          if (st && st !== 'NOT_STARTED') {
            pastBuild = true;
            break;
          }
        }
        if (pastBuild) break;
      }
    }

    return { buildStatus, buildStatusU, buildStarted, pastBuild };
  }

  async function getBuildName(prmId) {
    if (!prmId) return '';
    try {
      const prmJson = await fetchJson(
        `${BASE_PRM}/buildConfigVersion/${encodeURIComponent(prmId)}`
      );
      return (
        prmJson.buildConfig?.buildName ||
        prmJson.buildConfig?.prmName ||
        ''
      );
    } catch {
      return '';
    }
  }

  const raw = prompt('Paste MO list (one per line, e.g. MO157028664):');
  if (!raw) {
    showToast('No MO list provided', false);
    return;
  }

  const moList = raw
    .split(/[\r\n,]+/)
    .map(s => s.trim())
    .filter(s => /^MO\d+$/.test(s));

  if (moList.length === 0) {
    showToast('No valid MO numbers found', false);
    return;
  }

  const moSet = new Set(moList);

  let workOrders;
  try {
    workOrders = await fetchJson(`${BASE_MES}/workOrders`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(listPayload)
    });
  } catch (e) {
    console.error('Error fetching workOrders:', e);
    showToast('Error fetching workOrders – see console', false);
    return;
  }

  if (!Array.isArray(workOrders)) {
    console.error('Unexpected workOrders format:', workOrders);
    showToast('Unexpected workOrders response', false);
    return;
  }

  const moToWO = {};
  for (const wo of workOrders) {
    const mos = String(wo.moveOrderNumbers || '')
      .split(',')
      .map(s => s.trim())
      .filter(Boolean);
    for (const mo of mos) {
      if (moSet.has(mo) && !moToWO[mo]) {
        moToWO[mo] = wo;
      }
    }
  }

  const drpIssues = [];
  const lines = [];

  for (const mo of moList) {
    const wo = moToWO[mo];

    let statusOut = '';
    let moOut = '';
    let docOut = '';
    let nameOut = '';

    if (!wo) {
      lines.push(`Open\t\t\t`);
      continue;
    }

    const jobHeaderId = wo.jobHeaderId;
    const prmId = wo.prmId;

    let info;
    try {
      info = await fetchJson(
        `${BASE_MES}/workOrderInfo?jobHeaderId=${encodeURIComponent(jobHeaderId)}`
      );
    } catch (e) {
      console.error('Error fetching workOrderInfo for', mo, e);
      lines.push(`Open\t\t\t`);
      continue;
    }

    const inDrp = isInDrpLocation(info);
    const hasLpn = hasAnyLpn(info);

    const infoCurrentService = info.currentService || wo.currentService || '';

    const invStatus = getNestedStatus(info, 'Inbound', 'Inventory');
    const pickStatus = getNestedStatus(info, 'Stage', 'Picking');
    const invU = (invStatus || '').toUpperCase();
    const pickU = (pickStatus || '').toUpperCase();

    const { buildStatus, buildStatusU, buildStarted, pastBuild } = analyzeBuild(info);

    if (inDrp) {
      moOut = mo;
      docOut = wo.documentNumber || '';
      nameOut = await getBuildName(prmId);

      if (invU === 'COMPLETE') {
        statusOut = 'Inventoried';
      } else if (['NOT_STARTED', 'WORKING', 'COMPLETE'].includes(pickU)) {
        statusOut = 'Picking';
      } else if (buildStatusU === 'NOT_STARTED') {
        statusOut = 'Inventoried';
      } else {
        statusOut = '';
      }

      const safeServices = ['Inventory', 'Picking', 'Rack Prep'];
      const isSafeCurrent = safeServices.includes(infoCurrentService);

      if (!isSafeCurrent && buildStatusU !== 'NOT_STARTED' && (buildStarted || pastBuild)) {
        drpIssues.push({
          MO: mo,
          Document: wo.documentNumber || '',
          CurrentService: infoCurrentService || '',
          BuildStatus: buildStatus,
          Location: getFirstLocation(info),
          InventoryStatus: invStatus,
          PickingStatus: pickStatus
        });
      }

      lines.push(`${statusOut}\t${moOut}\t${docOut}\t${nameOut}`);
    } else {
      const pickingActive = ['NOT_STARTED', 'WORKING', 'COMPLETE'].includes(pickU);

      // Check if inventory is complete first, then picking, then open
      if (invU === 'COMPLETE') {
        moOut = mo;
        docOut = wo.documentNumber || '';
        nameOut = await getBuildName(prmId);
        lines.push(`Inventoried\t${moOut}\t${docOut}\t${nameOut}`);
      } else if (pickingActive) {
        moOut = mo;
        docOut = wo.documentNumber || '';
        nameOut = await getBuildName(prmId);
        lines.push(`Picking\t${moOut}\t${docOut}\t${nameOut}`);
      } else {
        lines.push(`Open\t\t\t`);
      }
    }
  }

  const output = lines.join('\n');
  console.log('=== DRP/Open summary ===');
  console.log(output);

  try {
    await navigator.clipboard.writeText(output);
    showToast('Summary copied to clipboard', true);
  } catch (e) {
    console.warn('Clipboard copy failed:', e);
    showToast('Copy failed – see console', false);
  }

  if (drpIssues.length > 0) {
    console.log('=== Orders in or past Build still in DRP being worked on ===');
    console.table(drpIssues);
  } else {
    console.log('No orders in/past Build in DRP with unsafe current service.');
  }
};

// ===== SERIAL SCAN FUNCTIONALITY =====
const REQUIRED_FIELDS_IBC = {
  base: new Set(["Asset Number", "Rack Position"]),
  msf: new Set(["MSF Part"]),
  msfAsset: new Set(["MSF Asset"])
};

const REQUIRED_FIELDS_LSC = new Set(["Asset Number", "Rack Position", "Warranty Start", "Warranty End"]);
const isMacField = (p) => /mac\s*address/i.test(p || "");

const classifyEntry = (attrs = [], mode = 'IBC') => {
  if (mode === 'LSC') {
    const prompts = attrs.map(a => a.attributePrompt);
    const unique = new Set(prompts);
    const isExactFour = unique.size === 4;
    const isExactMatch = isExactFour && [...unique].every(p => REQUIRED_FIELDS_LSC.has(p));
    return isExactMatch ? "MAIN" : "NODE";
  }
  return "MAIN";
};

const missingPrompts = (attrs = [], mode = 'IBC') => {
  const missing = [];
  
  if (mode === 'LSC') {
    // LSC mode (old): Check MAC address fields separately
    for (const a of attrs) {
      if (isMacField(a.attributePrompt)) continue;
      const v = a?.value;
      if (v === null || v === "" || (Array.isArray(v) && v.length === 0)) {
        missing.push(a.attributePrompt || "(Unnamed Attribute)");
      }
    }
  } else {
    // IBC mode (new): Check base required fields
    const prompts = new Set(attrs.map(a => a.attributePrompt));
    
    for (const field of REQUIRED_FIELDS_IBC.base) {
      const attr = attrs.find(a => a.attributePrompt === field);
      if (!attr || attr.value === null || attr.value === "" || (Array.isArray(attr.value) && attr.value.length === 0)) {
        missing.push(field);
      }
    }

    // If item has MSF Part field, it's required
    if (prompts.has("MSF Part")) {
      const attr = attrs.find(a => a.attributePrompt === "MSF Part");
      if (!attr || attr.value === null || attr.value === "" || (Array.isArray(attr.value) && attr.value.length === 0)) {
        missing.push("MSF Part");
      }
    }

    // If item has MSF Asset field, it's required
    if (prompts.has("MSF Asset")) {
      const attr = attrs.find(a => a.attributePrompt === "MSF Asset");
      if (!attr || attr.value === null || attr.value === "" || (Array.isArray(attr.value) && attr.value.length === 0)) {
        missing.push("MSF Asset");
      }
    }
  }
  
  return [...new Set(missing)];
};

const makeReport = (rows, mode = 'IBC') => {
  const affected = [];
  const lines = [];
  for (const row of rows) {
    const serial = row.serialNumber || "(unknown serial)";
    const attrs = Array.isArray(row.attributes) ? row.attributes : [];
    const role = classifyEntry(attrs, mode);
    const miss = missingPrompts(attrs, mode);
    if (miss.length > 0) {
      affected.push({ serial, role, miss });
      lines.push(`• ${serial}  [${role === "MAIN" ? "Main Device" : "Node"}] missing: ${miss.join(", ")}`);
    }
  }
  return { affected, lines };
};

const openSerialScanModal = ({ total, affected, lines }) => {
  const overlay = document.createElement("div");
  overlay.className = "ssc-overlay";
  const modal = document.createElement("div");
  modal.className = "ssc-modal";
  const countBadge = (txt, cls = "") => `<span class="ssc-badge ${cls}">${txt}</span>`;
  const rolePill = (role) => `<span class="ssc-pill ${role === "NODE" ? "node" : ""}">${role === "MAIN" ? "Main Device" : "Node"}</span>`;
  const listHTML = affected.map(a => `
    <div class="ssc-item">
      <div class="ssc-row">
        <strong class="ssc-serial">${a.serial}</strong>
        ${rolePill(a.role)}
        <span class="ssc-badge warn">${a.miss.length} missing</span>
      </div>
      <div class="ssc-missing"><strong>Missing:</strong> ${a.miss.join(", ")}</div>
    </div>
  `).join("");

  modal.innerHTML = `
    <div class="ssc-header">
      <h3 class="ssc-title">Missing Attributes</h3>
      <div class="ssc-actions">
        <button class="ssc-btn" id="sscCopy">Copy</button>
        <button class="ssc-btn primary" id="sscClose">Close</button>
      </div>
    </div>
    <div class="ssc-badges">
      ${countBadge(`${total} total`)}
      ${countBadge(`${affected.length} with missing`, "warn")}
    </div>
    <div class="ssc-list">${affected.length ? listHTML : `<div class="ssc-item">All serials complete (MAC fields ignored).</div>`}</div>
    <div class="ssc-pre" id="sscPre">${lines.join("\n")}</div>
  `;
  document.body.appendChild(overlay);
  document.body.appendChild(modal);
  const closeAll = () => { overlay.remove(); modal.remove(); };
  modal.querySelector("#sscClose").addEventListener("click", closeAll);
  overlay.addEventListener("click", closeAll);
  modal.querySelector("#sscCopy").addEventListener("click", async () => {
    const txt = modal.querySelector("#sscPre").textContent || "";
    try {
      await navigator.clipboard.writeText(txt);
      const b = modal.querySelector("#sscCopy");
      b.textContent = "Copied";
      setTimeout(() => b.textContent = "Copy", 1100);
    } catch {
      alert("Copy failed.");
    }
  });
};

// Helper function to show toast notification for serial scan
function showSerialScanToast(message, success = true) {
  const existing = document.getElementById('serial-scan-toast');
  if (existing) existing.remove();
  const t = document.createElement('div');
  t.id = 'serial-scan-toast';
  t.textContent = message;
  t.style.cssText = `
    position: fixed;
    bottom: 16px;
    right: 16px;
    padding: 12px 18px;
    background: ${success ? '#2e7d32' : '#f57c00'};
    color: #fff;
    font-size: 14px;
    font-family: system-ui,-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;
    border-radius: 6px;
    z-index: 99999;
    box-shadow: 0 3px 10px rgba(0,0,0,0.3);
    opacity: 0;
    transition: opacity .2s ease;
    max-width: 500px;
    line-height: 1.4;
  `;
  document.body.appendChild(t);
  requestAnimationFrame(() => { t.style.opacity = '1'; });
  setTimeout(() => {
    t.style.opacity = '0';
    setTimeout(() => t.remove(), 200);
  }, 4000);
}

async function runSerialScan(mode = 'IBC') {
  try {
    let jobHeaderId = getWorkOrderId();
    console.log('Running serial scan, jobHeaderId:', jobHeaderId);
    if (!jobHeaderId) {
      jobHeaderId = prompt("Enter Work Order (jobHeaderId):");
      if (!jobHeaderId) {
        showSerialScanToast("No jobHeaderId provided", false);
        return;
      }
    }
    
    const docType = mode === 'LSC' ? 'CFG_LTO' : 'WIP_JOB';
    console.log('Serial scan mode:', mode, 'docType:', docType);
    
    console.log('Sending SCAN_ORDER message...');
    const resp = await safeSendMessage({
      type: "SCAN_ORDER",
      jobHeaderId,
      orgId: "3346",
      docType: docType
    });
    
    console.log('Received response:', resp);
    console.log('Response type:', typeof resp);
    console.log('Response ok:', resp?.ok);
    
    if (!resp) {
      throw new Error("No response received from background script");
    }
    
    if (!resp?.ok) {
      throw new Error(resp?.error || "Background error");
    }

    // Process the response data
    const rows = resp.serials;
    if (!Array.isArray(rows)) {
      throw new Error("Invalid response format - expected array of serials");
    }
    
    const expectedDocType = mode === 'LSC' ? 'CFG_LTO' : 'WIP_JOB';
    
    // Check if we have any serials
    if (rows.length === 0) {
      // Check if it might be the wrong document type
      const actualDocType = resp.info?.documentSource || resp.info?.documentType;
      if (actualDocType && actualDocType !== expectedDocType) {
        const suggestion = actualDocType === 'WIP_JOB' ? 'IB Scan Checker' : 'Lab Scan Checker';
        showSerialScanToast(
          `No serials found. This is a ${actualDocType} order. Use ${suggestion} instead.`,
          false
        );
      } else {
        showSerialScanToast(`No serials found for this work order in ${mode} mode.`, false);
      }
      return;
    }
    
    // Check if serials match the expected document type
    const firstSerial = rows[0];
    const actualDocType = firstSerial?.documentType;
    
    if (actualDocType && actualDocType !== expectedDocType) {
      const suggestion = actualDocType === 'WIP_JOB' ? 'IB Scan Checker' : 'Lab Scan Checker';
      showSerialScanToast(
        `Wrong document type! This is a ${actualDocType} order. Use ${suggestion} instead.`,
        false
      );
      return;
    }
    
    const total = rows.length;
    const { affected, lines } = makeReport(rows, mode);
    
    console.log('Debug - Total serials:', total);
    console.log('Debug - Affected count:', affected.length);
    
    openSerialScanModal({ total, affected, lines });
  } catch (e) {
    console.error("[Serial Scan] Error:", e);
    if (e.message && e.message.includes('refresh this page')) {
      showSerialScanToast(e.message, false);
    } else {
      showSerialScanToast(`Scan failed: ${e.message}`, false);
    }
  }
}

// ===== IB UPDATE TRACKER =====
async function runIBUpdateTracker() {
  const WEBAPP = 'https://script.google.com/macros/s/AKfycbxV2MzkL-d7m-Wq_sskGCIUKhTaWmAk8kuVDq71uVYACbZ2OD7LfgKx2LDCaITNAvs/exec';
  const TOKEN_STORAGE_KEY = 'mes-ib-sheet-token';
  const TYPE = 'IBMOR';

  const promptForToken = (existingToken = '') => {
    const input = prompt('Enter IB Update Tracker token:', existingToken) || '';
    return input.trim();
  };

  const isAuthFailure = (status, errorMessage = '') => {
    if (status === 401 || status === 403) return true;
    const msg = String(errorMessage || '').toLowerCase();
    return (
      msg.includes('token') ||
      msg.includes('auth') ||
      msg.includes('unauthorized') ||
      msg.includes('forbidden') ||
      msg.includes('invalid password') ||
      msg.includes('wrong password')
    );
  };

  let token = localStorage.getItem(TOKEN_STORAGE_KEY) || '';
  if (!token) {
    token = promptForToken();
    if (!token) {
      showSerialScanToast('IB Update Tracker token is required', false);
      return;
    }
  }

  const URL_LIST = 'https://apirouter.apps.wwt.com/api/forward/mes-api/workOrders';
  const URL_INFO = 'https://apirouter.apps.wwt.com/api/forward/mes-api/workOrderInfo?jobHeaderId=';

  const postBody = {
    limit: 1000,
    offset: 0,
    releasedOnly: true,
    includeCompPercent: true,
    includeLocations: true,
    includeOrderManager: true,
    facilityCode: 'WPC',
    showUnassignedOnly: false,
    labArea: 'L4',
    orderBy: [
      { orderField: 'priority', orderDirection: 'ASC' },
      { orderField: 'criticalRatio', orderDirection: 'ASC' },
      { orderField: 'documentNumber', orderDirection: 'ASC' }
    ],
    dynamicFilters: [
      {
        name: 'serviceTemplateName',
        displayAs: 'Service Template',
        type: 'dynamicLov',
        dataType: 'string',
        operator: '%',
        lovName: 'templates',
        values: ['%IB%']
      },
      {
        name: 'labDestinationName',
        displayAs: 'Lab Destination',
        type: 'dynamicLov',
        dataType: 'string',
        operator: '%',
        lovName: 'labDestinations',
        availableOperators: ['=', '?', '%'],
        values: ['%naic1%']
      }
    ]
  };

  try {
    showSerialScanToast(`Grabbing ${TYPE} orders…`, true);

    const lr = await fetch(URL_LIST, {
      method: 'POST',
      credentials: 'include',
      headers: {
        'Content-Type': 'application/json',
        'Accept': 'application/json'
      },
      body: JSON.stringify(postBody)
    });

    if (!lr.ok) throw new Error('List HTTP ' + lr.status);
    const list = await lr.json();
    
    if (!Array.isArray(list) || !list.length) {
      showSerialScanToast('No orders found', false);
      return;
    }

    showSerialScanToast(`Fetching details for ${list.length} orders…`, true);

    const infos = (await Promise.all(list.map(async o => {
      try {
        const r = await fetch(URL_INFO + encodeURIComponent(o.jobHeaderId), {
          credentials: 'include',
          headers: { 'Accept': 'application/json' }
        });
        if (!r.ok) return null;
        const i = await r.json();
        return {
          documentNumber: i.documentNumber,
          currentService: i.currentService,
          currentServiceStatus: i.currentServiceStatus,
          serviceTemplateName: i.serviceTemplateName,
          operations: i.operations,
          onHandLPNs: i.onHandLPNs
        };
      } catch {
        return null;
      }
    }))).filter(Boolean);

    if (!infos.length) {
      showSerialScanToast('No order details retrieved', false);
      return;
    }

    const uploadWithToken = async (uploadToken) => {
      const up = await fetch(WEBAPP + '?token=' + encodeURIComponent(uploadToken), {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify({ type: TYPE, infos, token: uploadToken })
      });
      const j = await up.json().catch(() => ({}));
      return { up, j };
    };

    showSerialScanToast(`Uploading ${infos.length} orders…`, true);
    let { up, j } = await uploadWithToken(token);

    if (!up.ok || !j.ok) {
      const authFailed = isAuthFailure(up.status, j.error || up.statusText || up.status);
      if (authFailed) {
        localStorage.removeItem(TOKEN_STORAGE_KEY);
        const newToken = promptForToken('');
        if (!newToken) {
          showSerialScanToast('IB Update Tracker token is required', false);
          return;
        }
        token = newToken;
        showSerialScanToast('Retrying upload with new token…', true);
        ({ up, j } = await uploadWithToken(token));
      }
    }

    if (up.ok && j.ok) {
      localStorage.setItem(TOKEN_STORAGE_KEY, token);
      showSerialScanToast(
        `✅ Sent ${j.wrote || infos.length} ${TYPE} rows (LPN docs: ${j.lpnDocs || 0}, LPNs: ${j.lpnCount || 0})`,
        true
      );
    } else {
      showSerialScanToast(
        `⚠️ Upload failed: ${j.error || up.statusText || up.status}`,
        false
      );
    }
  } catch (e) {
    console.error('IB Update Tracker error:', e);
    showSerialScanToast(`❌ Error: ${e?.message || String(e)}`, false);
  }
}

// Helper function to check and open reportal intro tab if needed
async function checkAndOpenReportalIntro() {
  const REPORTAL_LAST_OPENED_KEY = 'mes-reportal-last-opened';
  const EIGHT_HOURS_MS = 8 * 60 * 60 * 1000; // 8 hours in milliseconds
  
  try {
    const stored = localStorage.getItem(REPORTAL_LAST_OPENED_KEY);
    const lastOpened = stored ? parseInt(stored, 10) : 0;
    const now = Date.now();
    
    // Check if 8 hours have passed since last opening
    if (now - lastOpened >= EIGHT_HOURS_MS) {
      // Update the timestamp before opening to prevent duplicate opens
      localStorage.setItem(REPORTAL_LAST_OPENED_KEY, String(now));
      
      // Send message to background script to open and close the tab
      await safeSendMessage({
        type: 'OPEN_REPORTAL_INTRO'
      });
    }
  } catch (error) {
    console.error('Error checking reportal intro:', error);
    // Don't throw - continue with normal reportal execution
  }
}

async function runIbiReportal({ lookupPayload, targetTemplateName, downloadBaseName }) {
  const IBI_TEST_URL = 'https://reports.wwt.com/ibi_apps/WFServlet';
  const FINAL_FEX = 'app/wwt_gah_direct.fex';
  const GAH_CODE = '1083900';
  const payload = lookupPayload;

  const extractTemplateInfo = (bodyText) => {
    if (!bodyText) return null;
    const start = bodyText.indexOf('<fxf');
    const end = bodyText.indexOf('</fxf>');
    if (start === -1 || end === -1) return null;

    const xmlOnly = bodyText.slice(start, end + 6);
    const doc = new DOMParser().parseFromString(xmlOnly, 'text/xml');
    const parserError = doc.querySelector('parsererror');
    if (parserError) return null;

    const rows = Array.from(doc.querySelectorAll('table > tr'));
    if (!rows.length) return null;

    const templates = rows
      .map((row) => {
        const nameCell = row.querySelector('td[colnum="c0"]');
        const templateCell = row.querySelector('td[colnum="c1"]');
        const templateName = (nameCell?.getAttribute('rawvalue') || nameCell?.textContent || '').trim();
        const templateString = (templateCell?.getAttribute('rawvalue') || templateCell?.textContent || '').trim();
        const templateMatch = templateString.match(/\(\}(\d+)\s*$/);
        const templateId = templateMatch ? templateMatch[1] : '';
        if (!templateName || !templateId || !templateString) return null;
        return { templateId, templateString, templateName };
      })
      .filter(Boolean);

    if (!templates.length) return null;

    const exactMatch = templates.find(
      (t) => t.templateName.toLowerCase() === String(targetTemplateName || '').toLowerCase()
    );
    return exactMatch || templates[0];
  };

  try {
    const response = await safeSendMessage({
      type: 'IBI_TEST_POST',
      url: IBI_TEST_URL,
      payload
    });

    if (!response?.ok) {
      throw new Error(response?.error || `HTTP ${response?.status || 'unknown'} ${response?.statusText || ''}`.trim());
    }

    const outputType = response.outputType || response.contentType || 'unknown output';
    if (!response.bodyText) {
      throw new Error('No XML body returned');
    }

    const templateInfo = extractTemplateInfo(response.bodyText);
    if (templateInfo) {
      const dateSuffix = getDayKey();
      const safeBaseName = (downloadBaseName || 'IBI Report').replace(/[\\/:*?"<>|]/g, '_').trim();
      const datedFileName = `${safeBaseName} ${dateSuffix}.xlsx`;

      const downloadResponse = await safeSendMessage({
        type: 'IBI_TEST_DOWNLOAD_EXCEL',
        url: IBI_TEST_URL,
        filename: datedFileName,
        fields: {
          IBIMR_action: 'MR_RUN_FEX',
          IBIMR_sub_action: 'MR_STD_REPORT',
          IBIMR_drill: 'RUNNID',
          IBIMR_folder: '#guidedadhocy',
          IBIMR_domain: 'commonto/commonto.htm',
          IBIMR_fex: FINAL_FEX,
          P_INSTANCE: 'ERP',
          WFFMT: 'EXL07',
          wffmt_hidden: 'EXL07',
          P_GAH_CODE: GAH_CODE,
          P_TEMPLATE_ID: templateInfo.templateId,
          template_id_hidden: templateInfo.templateId,
          P_TEMPLATE_STRING: templateInfo.templateString,
          tmplt_string_hidden: templateInfo.templateString,
          P_REPORT_TILE: templateInfo.templateName || 'MES IBI Report',
          P_REPORT_TITLE: templateInfo.templateName || 'MES IBI Report'
        }
      });

      if (!downloadResponse?.ok) {
        throw new Error(downloadResponse?.error || `Download failed (${downloadResponse?.status || 'unknown'})`);
      }

      showSerialScanToast(`Excel download started: ${downloadResponse.filename || 'IBI_Report.xlsx'}`, true);
    } else {
      throw new Error('Template data not found in XML response');
    }

    showSerialScanToast(`IBI test POST succeeded (${response.status}) ${outputType}`, true);
    if (response.bodyPreview) {
      console.log('IBI test response preview:', response.bodyPreview);
    }
  } catch (error) {
    console.error('IBI test POST failed:', error);
    showSerialScanToast(`IBI test POST failed: ${error?.message || String(error)}`, false);
  }
}

async function runInventoryReportal() {
  // Check and open reportal intro tab if needed
  await checkAndOpenReportalIntro();
  
  const payload = `IBIMR_action=MR_RUN_FEX&IBIMR_sub_action=MR_STD_REPORT&IBIMR_drill=RUNNID&IBIMR_folder=%23guidedadhocy&IBIMR_domain=commonto/commonto.htm&IBIMR_fex=app/wwt_guided_ad_hoc_otf_ajax_fex.fex&P_INSTANCE=ERP&P_SELECT_STATEMENTS=TEMPLATE_NAME,TEMPLATE_HEADER||'(}'||TEMPLATE_SEGMENTS||'(}'||nvl(TEMPLATE_ORDERBY,'NONE')||'(}'||nvl(REPORT_TITLE,'NONE')||'(}'||TEMPLATE_ID&P_LOOKUP_V=TEMPLATE_LOOKUP_V&P_USER_WHERE=FOC_NONE&P_WHERE_STATEMENTS=and business_unit = 'NAIC1%20Inventory' and gah_code = '1083900'&P_GAH_CODE=1083900`;
  return runIbiReportal({
    lookupPayload: payload,
    targetTemplateName: 'Org 40 at NAIC1',
    downloadBaseName: 'Inventory GAH'
  });
}

async function runOutboundReportal() {
  // Check and open reportal intro tab if needed
  await checkAndOpenReportalIntro();
  
  const payload = `IBIMR_action=MR_RUN_FEX&IBIMR_sub_action=MR_STD_REPORT&IBIMR_drill=RUNNID&IBIMR_folder=%23guidedadhocy&IBIMR_domain=commonto/commonto.htm&IBIMR_fex=app/wwt_guided_ad_hoc_otf_ajax_fex.fex&P_INSTANCE=ERP&P_SELECT_STATEMENTS=TEMPLATE_NAME,TEMPLATE_HEADER||'(}'||TEMPLATE_SEGMENTS||'(}'||nvl(TEMPLATE_ORDERBY,'NONE')||'(}'||nvl(REPORT_TITLE,'NONE')||'(}'||TEMPLATE_ID&P_LOOKUP_V=TEMPLATE_LOOKUP_V&P_USER_WHERE=FOC_NONE&P_WHERE_STATEMENTS=and business_unit = 'Outbound' and gah_code = '1083900'&P_GAH_CODE=1083900`;
  return runIbiReportal({
    lookupPayload: payload,
    targetTemplateName: 'Yee NAIC1',
    downloadBaseName: 'Outbound GAH'
  });
}

// ===== TOOLS MENU EVENT HANDLER =====
document.addEventListener('click', (e) => {
  const applyStyleButton = e.target.closest('#mes-btn-style-apply');
  if (applyStyleButton) {
    const settings = getButtonSettingsFromControls();
    applyButtonSettings(settings);
    saveButtonSettings(settings);
    return;
  }

  const resetStyleButton = e.target.closest('#mes-btn-style-reset');
  if (resetStyleButton) {
    const defaults = { ...DEFAULT_BUTTON_SETTINGS };
    applyButtonSettings(defaults);
    saveButtonSettings(defaults);
    syncButtonControls(defaults);
    return;
  }

  const toolButton = e.target.closest('.mes-tool-item');
  if (!toolButton) return;

  const tool = toolButton.dataset.tool;
  const menu = document.getElementById('mes-tools-menu');

  if (menu) menu.removeAttribute('open');

  if (tool === 'mo-auto-prm') {
    moAutoPrmCopy();
  } else if (tool === 'mo-tracker') {
    moTrackerStatusUpdater();
  } else if (tool === 'ib-update-tracker') {
    runIBUpdateTracker();
  } else if (tool === 'inventory-reportal') {
    runInventoryReportal();
  } else if (tool === 'outbound-reportal') {
    runOutboundReportal();
  } else if (tool === 'serial-scan-lsc') {
    runSerialScan('LSC');
  } else if (tool === 'serial-scan-ibc') {
    runSerialScan('IBC');
  } else if (tool === 'check-updates') {
    checkForExtensionUpdates('manual-tool-check');
  }
});

document.addEventListener("click", handleCopyClick, true);

// Function to download and modify placard for a specific WO ID
async function downloadModifiedPlacardByWoId(workOrderId, buttonElement = null) {
  try {
    // Show loading indicator
    if (buttonElement) {
      buttonElement.disabled = true;
      const originalText = buttonElement.textContent;
      buttonElement.textContent = 'Downloading...';
      buttonElement.dataset.originalText = originalText;
    }
    // Fetch work order data with credentials using correct endpoint
    const response = await fetch(
      `https://apirouter.apps.wwt.com/api/forward/mes-api/workOrderInfo?jobHeaderId=${workOrderId}`,
      {
        method: 'GET',
        credentials: 'include',
        headers: {
          'Accept': 'application/json',
          'Content-Type': 'application/json'
        }
      }
    );
    
    if (!response.ok) {
      throw new Error(`Failed to fetch work order data: ${response.status} ${response.statusText}`);
    }
    
    const workOrderData = await response.json();
    
    console.log('Work order data fetched:', workOrderData);
    
    // Send message to background script to modify and download placard
    const message = {
      type: 'modifyPlacardExcel',
      url: `https://apirouter.apps.wwt.com/api/forward/mes-api/workOrders/${workOrderId}/placard`,
      workOrderData: workOrderData,
      filename: `Placard_${workOrderData.documentNumber}.xlsx`
    };
    
    console.log('Sending message to background script:', message);
    
    let bgResponse;
    try {
      bgResponse = await safeSendMessage(message);
      console.log('Response from background script:', bgResponse);
    } catch (error) {
      console.error('Error sending message to background:', error);
      throw error;
    }
    
    if (buttonElement) {
      buttonElement.disabled = false;
      buttonElement.textContent = buttonElement.dataset.originalText || 'Download Placard (Auto-filled)';
    }
    
    if (bgResponse && bgResponse.ok) {
      console.log('Placard downloaded successfully for WO', workOrderId);
      if (buttonElement) {
        buttonElement.textContent = '✓ Downloaded!';
        setTimeout(() => {
          buttonElement.textContent = buttonElement.dataset.originalText || 'Download Placard (Auto-filled)';
        }, 2000);
      }
      return { success: true, woId: workOrderId, woNumber: workOrderData.documentNumber };
    } else {
      console.error('Failed to download placard:', bgResponse?.error);
      return { success: false, woId: workOrderId, error: bgResponse?.error || 'Unknown error' };
    }
  } catch (error) {
    console.error('Error downloading placard for WO', workOrderId, ':', error);
    
    // Check if it's an extension context error
    if (error.message && error.message.includes('refresh this page')) {
      alert(error.message);
    }
    
    if (buttonElement) {
      buttonElement.disabled = false;
      buttonElement.textContent = buttonElement.dataset.originalText || 'Download Placard (Auto-filled)';
    }
    return { success: false, woId: workOrderId, error: error.message };
  }
}

// Function to download and modify placard
async function downloadModifiedPlacard() {
  try {
    // Get work order ID from the current page
    const workOrderMatch = window.location.pathname.match(/\/orders\/(\d+)/);
    if (!workOrderMatch) {
      console.error('Could not find work order ID in URL');
      alert('Could not find work order ID. Please make sure you are on a work order page.');
      return;
    }
    
    const workOrderId = workOrderMatch[1];
    
    // Use the helper function with button reference
    const btn = document.getElementById('placard-download-btn');
    await downloadModifiedPlacardByWoId(workOrderId, btn);
    
  } catch (error) {
    console.error('Error downloading placard:', error);
    alert('Error: ' + error.message);
    
    const btn = document.getElementById('placard-download-btn');
    if (btn) {
      btn.disabled = false;
      btn.textContent = 'Download Placard (Auto-filled)';
    }
  }
}

// Intercept "Generate Placard" menu click
function interceptGeneratePlacardClick() {
  // Use event delegation to catch clicks on dynamically created menu items
  document.addEventListener('click', (e) => {
    // Check if clicked element or its parent is the Generate Placard menu item
    const target = e.target.closest('.p-tieredmenu-item');
    if (!target) return;
    
    // Check if this is the "Generate Placard" item
    const label = target.querySelector('.p-tieredmenu-item-label');
    if (label && label.textContent.trim() === 'Generate Placard') {
      // Prevent the default placard generation
      e.preventDefault();
      e.stopPropagation();
      e.stopImmediatePropagation();
      
      console.log('Intercepted Generate Placard click - using auto-filled version');
      
      // Call our modified placard download function
      downloadModifiedPlacard();
      
      // Close the menu
      const menu = target.closest('.p-tieredmenu');
      if (menu) {
        menu.style.display = 'none';
      }
      
      return false;
    }
  }, true); // Use capture phase to intercept before other handlers
}

// Add placard download button to the page (optional - now we override the menu)
function addPlacardDownloadButton() {
  // Only add if on a work order details page
  if (!window.location.pathname.match(/\/orders\/\d+/)) {
    return;
  }
  
  // Don't add if already exists
  if (document.getElementById('placard-download-btn')) {
    return;
  }
  
  const observer = new MutationObserver(() => {
    // Look for the actions area or header
    const actionsArea = document.querySelector('.order-actions, .work-order-actions, [class*="action"], header');
    
    if (actionsArea && !document.getElementById('placard-download-btn')) {
      const downloadBtn = document.createElement('button');
      downloadBtn.id = 'placard-download-btn';
      downloadBtn.textContent = 'Download Placard (Auto-filled)';
      downloadBtn.className = 'btn btn-success';
      downloadBtn.style.cssText = 'margin: 5px; padding: 8px 12px; background-color: #28a745; color: white; border: none; border-radius: 4px; cursor: pointer; font-weight: 500;';
      downloadBtn.addEventListener('click', downloadModifiedPlacard);
      
      actionsArea.appendChild(downloadBtn);
      observer.disconnect();
    }
  });
  
  observer.observe(document.body, {
    childList: true,
    subtree: true
  });
}

currentButtonSettings = loadButtonSettings();
applyButtonSettings(currentButtonSettings);

enhancePnCells();
ensureCopyAllButton();
enhanceLpnTable();
ensureLpnHeaderButtons();
ensureWoCopyButton();
createToolsMenu();
// addPlacardDownloadButton();
interceptGeneratePlacardClick(); // Intercept the Generate Placard menu click
checkForExtensionUpdates('content-startup');

const observer = new MutationObserver(() => {
  enhancePnCells();
  ensureCopyAllButton();
  enhanceLpnTable();
  ensureLpnHeaderButtons();
  ensureWoCopyButton();
  createToolsMenu();
  // addPlacardDownloadButton();
});

observer.observe(document.body, {
  childList: true,
  subtree: true,
});
