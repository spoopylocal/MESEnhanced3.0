// Load ExcelJS library for Excel modification (CSP-safe version)
try {
  importScripts('exceljs.bare.min.js');
  console.log('ExcelJS import successful');
} catch (e) {
  console.error('Failed to import ExcelJS:', e);
}

// Browser API compatibility (Chrome/Firefox)
if (typeof browser === 'undefined') {
  var browser = chrome;
}

console.log('Background script loading...');
console.log('ExcelJS loaded:', typeof ExcelJS !== 'undefined');
if (typeof ExcelJS !== 'undefined') {
  console.log('ExcelJS version:', ExcelJS.version || 'unknown');
}

const UPDATE_FEED_URLS = [
  'https://raw.githubusercontent.com/spoopylocal/MESEnhanced3.0/main/update-feed.json',
  'https://raw.githubusercontent.com/spoopylocal/MESEnhanced3.0/master/update-feed.json',
  browser.runtime.getURL('update-feed.json')
];
const UPDATE_ALARM_NAME = 'mes-enhanced-daily-update-check';
const UPDATE_STATUS_KEY = 'mesUpdateStatus';

// IBMOR -> Sheet sync config/state
const IBMOR_SYNC_WEBAPP_URL = 'https://long-fog-579c.thomascarnellpc.workers.dev';
const IBMOR_SYNC_TYPE = 'IBMOR';
const IBMOR_SYNC_FIXED_TOKEN = 'ihateben';
const IBMOR_SYNC_TOKEN_KEY = 'mes-ib-sheet-token';
const IBMOR_SYNC_CONFIG_KEY = 'mes-ibmor-sheet-sync-config';
const IBMOR_SYNC_STATUS_KEY = 'mes-ibmor-sheet-sync-status';
const IBMOR_SYNC_CLIENT_ID_KEY = 'mes-ibmor-sheet-sync-client-id';
const IBMOR_SYNC_ALARM_NAME = 'mes-ibmor-sheet-sync-5m';
const IBMOR_SYNC_WINDOW_MS = 5 * 60 * 1000;
const IBMOR_SYNC_URL_LIST = 'https://apirouter.apps.wwt.com/api/forward/mes-api/workOrders';
const IBMOR_SYNC_URL_INFO = 'https://apirouter.apps.wwt.com/api/forward/mes-api/workOrderInfo?jobHeaderId=';
const IBMOR_SYNC_URL_FLAGS_BASE = 'https://apirouter.apps.wwt.com/api/forward/mes-api/workOrder/';
const IBMOR_SYNC_POST_BODY = {
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
    { orderField: 'documentNumber' }
  ],
  labArea: 'L4',
  dynamicFilters: [
    {
      name: 'labDestinationName',
      displayAs: 'Lab Destination',
      type: 'dynamicLov',
      dataType: 'string',
      operator: '=',
      availableOperators: ['=', '\u2260', '%'],
      lovName: 'labDestinations',
      filterValues: [
        'ABM Lab Production',
        'Build',
        'Certification',
        'FLX',
        'INT',
        'L2 Expansion',
        'L2 Production',
        'L3 Networking',
        'L3 Production',
        'Microwave',
        'NAIC1 L3',
        'NAIC1 Rack & Stack',
        'RF',
        'Rack & Stack',
        'Repair',
        'Test',
        'VIC Lab'
      ],
      values: ['NAIC1 Rack & Stack']
    }
  ]
};

let ibmorSyncInFlight = null;

function compareVersions(versionA, versionB) {
  const a = String(versionA || '0').split('.').map((n) => parseInt(n, 10) || 0);
  const b = String(versionB || '0').split('.').map((n) => parseInt(n, 10) || 0);
  const maxLength = Math.max(a.length, b.length);
  for (let i = 0; i < maxLength; i += 1) {
    const av = a[i] || 0;
    const bv = b[i] || 0;
    if (av > bv) return 1;
    if (av < bv) return -1;
  }
  return 0;
}

async function saveUpdateStatus(status) {
  await browser.storage.local.set({ [UPDATE_STATUS_KEY]: status });
  return status;
}

async function getSavedUpdateStatus() {
  const saved = await browser.storage.local.get(UPDATE_STATUS_KEY);
  return saved?.[UPDATE_STATUS_KEY] || null;
}

async function checkForUpdates(trigger = 'manual') {
  const currentVersion = browser.runtime.getManifest().version;
  try {
    let payload = null;
    let resolvedFeedUrl = null;
    const failures = [];

    for (const feedUrl of UPDATE_FEED_URLS) {
      try {
        const response = await fetch(feedUrl, {
          method: 'GET',
          cache: 'no-cache',
          headers: { 'Accept': 'application/json' }
        });

        if (!response.ok) {
          failures.push(`${feedUrl} -> HTTP ${response.status}`);
          continue;
        }

        payload = await response.json();
        resolvedFeedUrl = feedUrl;
        break;
      } catch (feedError) {
        failures.push(`${feedUrl} -> ${feedError?.message || String(feedError)}`);
      }
    }

    if (!payload) {
      throw new Error(`Update feed unavailable. Tried: ${failures.join(' | ')}`);
    }

    const latestVersion = String(payload?.latestVersion || currentVersion);
    const notes = Array.isArray(payload?.notes)
      ? payload.notes.filter((item) => typeof item === 'string').slice(0, 6)
      : [];
    const downloadUrl = String(payload?.downloadUrl || '');
    const updateAvailable = compareVersions(latestVersion, currentVersion) > 0;

    const status = {
      checkedAt: Date.now(),
      trigger,
      currentVersion,
      latestVersion,
      updateAvailable,
      notes,
      downloadUrl,
      feedUrl: resolvedFeedUrl,
      ok: true
    };

    return saveUpdateStatus(status);
  } catch (error) {
    const status = {
      checkedAt: Date.now(),
      trigger,
      currentVersion,
      latestVersion: currentVersion,
      updateAvailable: false,
      notes: [],
      downloadUrl: '',
      ok: false,
      error: error?.message || String(error)
    };
    return saveUpdateStatus(status);
  }
}

function setupUpdateAlarm() {
  browser.alarms.create(UPDATE_ALARM_NAME, {
    delayInMinutes: 1,
    periodInMinutes: 24 * 60
  });
}

function setupIbmorSyncAlarm() {
  browser.alarms.create(IBMOR_SYNC_ALARM_NAME, {
    delayInMinutes: 1,
    periodInMinutes: 5
  });
}

function nowIso() {
  return new Date().toISOString();
}

function isEligibleMesOrdersUrl(url) {
  if (typeof url !== 'string' || !url) return false;
  return /^https:\/\/mes\.apps\.wwt\.com\/orders(?:\/[^\/?#]+)?(?:[?#].*)?$/.test(url);
}

async function getActiveEligibleMesOrdersTab() {
  const tabs = await browser.tabs.query({ active: true });
  for (const tab of tabs) {
    const url = String(tab?.url || '');
    if (isEligibleMesOrdersUrl(url)) {
      return {
        tabId: tab.id,
        windowId: tab.windowId,
        url
      };
    }
  }
  return null;
}

async function getOrCreateIbmorClientId() {
  const saved = await browser.storage.local.get(IBMOR_SYNC_CLIENT_ID_KEY);
  let id = saved?.[IBMOR_SYNC_CLIENT_ID_KEY];
  if (id && typeof id === 'string') {
    return id;
  }

  id = `${Date.now()}-${Math.random().toString(36).slice(2, 10)}`;
  await browser.storage.local.set({ [IBMOR_SYNC_CLIENT_ID_KEY]: id });
  return id;
}

async function getIbmorSyncConfig() {
  return {
    enabled: true,
    token: IBMOR_SYNC_FIXED_TOKEN,
    windowMs: IBMOR_SYNC_WINDOW_MS
  };
}

async function setIbmorSyncConfig(nextConfig = {}) {
  // Intentionally no-op: sync now uses a fixed token and is always enabled.
  return getIbmorSyncConfig();
}

async function saveIbmorSyncStatus(status = {}) {
  const payload = {
    ...status,
    checkedAt: Date.now(),
    checkedAtIso: nowIso()
  };
  await browser.storage.local.set({ [IBMOR_SYNC_STATUS_KEY]: payload });
  return payload;
}

async function getIbmorSyncStatus() {
  const saved = await browser.storage.local.get(IBMOR_SYNC_STATUS_KEY);
  return saved?.[IBMOR_SYNC_STATUS_KEY] || null;
}

async function ibmorFetchJson(url, opts) {
  const response = await fetch(url, opts);
  if (!response.ok) {
    throw new Error(`HTTP ${response.status} ${url}`);
  }
  return response.json();
}

async function ibmorPostToSheet(token, payload) {
  const response = await fetch(`${IBMOR_SYNC_WEBAPP_URL}?token=${encodeURIComponent(token)}`, {
    method: 'POST',
    headers: {
      'Content-Type': 'text/plain;charset=utf-8'
    },
    body: JSON.stringify(payload)
  });

  const text = await response.text().catch(() => '');
  let json = {};
  try {
    json = JSON.parse(text);
  } catch (e) {
    json = {};
  }

  return {
    response,
    text,
    json
  };
}

async function runIbmorSheetSync(trigger = 'alarm') {
  if (ibmorSyncInFlight) {
    return ibmorSyncInFlight;
  }

  ibmorSyncInFlight = (async () => {
    const config = await getIbmorSyncConfig();

    const lastStatus = await getIbmorSyncStatus();
    const lastCheckedAt = Number(lastStatus?.checkedAt || 0);
    const withinLocalWindow = lastCheckedAt > 0 && (Date.now() - lastCheckedAt) < config.windowMs;
    if (withinLocalWindow) {
      return saveIbmorSyncStatus({
        ok: true,
        ran: false,
        skipped: true,
        reason: 'local-recent-check-window',
        trigger,
        previousCheckedAt: lastCheckedAt,
        previousCheckedAtIso: lastStatus?.checkedAtIso || ''
      });
    }

    const activeMesTab = await getActiveEligibleMesOrdersTab();
    if (!activeMesTab) {
      return saveIbmorSyncStatus({
        ok: true,
        ran: false,
        skipped: true,
        reason: 'no-eligible-mes-orders-tab',
        trigger
      });
    }

    const clientId = await getOrCreateIbmorClientId();

    try {
      // Server-side claim check: if another user already uploaded in the 5-minute window, skip heavy API calls.
      const claim = await ibmorPostToSheet(config.token, {
        type: IBMOR_SYNC_TYPE,
        action: 'claimWindow',
        windowMs: config.windowMs,
        clientId,
        trigger
      });

      if (!claim.response.ok) {
        return saveIbmorSyncStatus({
          ok: false,
          ran: false,
          skipped: true,
          reason: 'claim-http-failed',
          trigger,
          statusCode: claim.response.status,
          details: claim.text.slice(0, 300)
        });
      }

      const shouldUpload = claim.json?.shouldUpload !== false;
      if (!shouldUpload) {
        return saveIbmorSyncStatus({
          ok: true,
          ran: false,
          skipped: true,
          reason: claim.json?.reason || 'recent-update-window',
          trigger,
          activeTabUrl: activeMesTab.url,
          lastUploadedAt: claim.json?.lastUploadedAt || null,
          sourceClientId: claim.json?.clientId || null
        });
      }

      const list = await ibmorFetchJson(IBMOR_SYNC_URL_LIST, {
        method: 'POST',
        credentials: 'include',
        headers: {
          'Content-Type': 'application/json',
          'Accept': 'application/json'
        },
        body: JSON.stringify(IBMOR_SYNC_POST_BODY)
      });

      if (!Array.isArray(list) || list.length === 0) {
        return saveIbmorSyncStatus({
          ok: true,
          ran: true,
          skipped: false,
          reason: 'no-orders',
          trigger,
          orderCount: 0,
          flagCount: 0
        });
      }

      const results = await Promise.all(list.map(async (order) => {
        const id = order?.jobHeaderId;
        if (!id) return null;

        try {
          const info = await ibmorFetchJson(`${IBMOR_SYNC_URL_INFO}${encodeURIComponent(id)}`, {
            credentials: 'include',
            headers: {
              'Accept': 'application/json'
            }
          });

          let flags = [];
          try {
            flags = await ibmorFetchJson(`${IBMOR_SYNC_URL_FLAGS_BASE}${encodeURIComponent(id)}/flags`, {
              credentials: 'include',
              headers: {
                'Accept': 'application/json'
              }
            });
          } catch (e) {
            flags = [];
          }

          return { info, flags };
        } catch (e) {
          return null;
        }
      }));

      const infos = [];
      const flagDetails = [];

      for (const result of results.filter(Boolean)) {
        infos.push(result.info);
        const flags = Array.isArray(result.flags) ? result.flags : [];
        for (const flag of flags) {
          flagDetails.push({
            ...flag,
            documentNumber: result.info?.documentNumber || '',
            jobHeaderId: result.info?.jobHeaderId || ''
          });
        }
      }

      const upload = await ibmorPostToSheet(config.token, {
        type: IBMOR_SYNC_TYPE,
        action: 'upload',
        windowMs: config.windowMs,
        clientId,
        trigger,
        infos,
        flagDetails
      });

      if (!upload.response.ok || upload.json?.ok !== true) {
        return saveIbmorSyncStatus({
          ok: false,
          ran: true,
          skipped: false,
          reason: 'upload-failed',
          trigger,
          statusCode: upload.response.status,
          details: upload.text.slice(0, 300),
          orderCount: infos.length,
          flagCount: flagDetails.length
        });
      }

      return saveIbmorSyncStatus({
        ok: true,
        ran: true,
        skipped: false,
        reason: 'uploaded',
        trigger,
        activeTabUrl: activeMesTab.url,
        orderCount: upload.json?.wrote?.IBMOR_Orders ?? infos.length,
        flagCount: upload.json?.wrote?.FlagDetails ?? flagDetails.length,
        wrote: upload.json?.wrote || null
      });
    } catch (error) {
      return saveIbmorSyncStatus({
        ok: false,
        ran: false,
        skipped: false,
        reason: 'exception',
        trigger,
        error: error?.message || String(error)
      });
    }
  })();

  try {
    return await ibmorSyncInFlight;
  } finally {
    ibmorSyncInFlight = null;
  }
}

browser.runtime.onStartup.addListener(() => {
  setupUpdateAlarm();
  setupIbmorSyncAlarm();
  checkForUpdates('startup');
  runIbmorSheetSync('startup');
});

browser.runtime.onInstalled.addListener(() => {
  setupUpdateAlarm();
  setupIbmorSyncAlarm();
  checkForUpdates('installed');
  runIbmorSheetSync('installed');
});

browser.alarms.onAlarm.addListener((alarm) => {
  if (alarm?.name === UPDATE_ALARM_NAME) {
    checkForUpdates('daily');
  }

  if (alarm?.name === IBMOR_SYNC_ALARM_NAME) {
    runIbmorSheetSync('alarm');
  }
});

setupUpdateAlarm();
setupIbmorSyncAlarm();

// Helper function to extract GICLAB ticket from labNotes
function extractGICLABTicket(labNotes) {
  if (!labNotes) return null;
  const match = labNotes.match(/GICLAB(\d+)/);
  return match ? `GICLAB${match[1]}` : null;
}

// Helper function to fetch power type from PRM API
async function fetchPowerType(prmId) {
  try {
    const response = await fetch(
      `https://apirouter.apps.wwt.com/prm-api/buildConfigVersion/${prmId}/run-list/ac-power`,
      {
        method: 'GET',
        credentials: 'include',
        headers: {
          'Accept': 'application/json',
          'Content-Type': 'application/json'
        }
      }
    );
    if (!response.ok) return 'N/A';
    
    const data = await response.json();
    // Get whipInstallationOrientation from first PDU entry
    if (data && data.length > 0 && data[0].pduMaterial) {
      return data[0].pduMaterial.whipInstallationOrientation || 'N/A';
    }
    return 'N/A';
  } catch (error) {
    console.error('Error fetching power type:', error);
    return 'N/A';
  }
}

// Helper function to prepare Excel data from work order
async function prepareExcelData(workOrderData) {
  const gicLabTicket = extractGICLABTicket(workOrderData.labNotes);
  const powerType = workOrderData.prmId 
    ? await fetchPowerType(workOrderData.prmId)
    : 'N/A';
  
  return {
    updateByLabel: [
      { label: 'MES Work Order:', valueCol: 2, value: workOrderData.documentNumber || '' },
      { label: 'Sales Order:', valueCol: 2, value: 'N/A' },
      { label: 'Move Order(s):', valueCol: 2, value: workOrderData.moveOrderNumbers || '' },
      { label: 'Order Type:', valueCol: 1, value: workOrderData.orderTypeCode || '' },
      { label: 'BTS:', value: 'N/A' },
      { label: 'GICLAB Ticket #:', valueCol: 2, value: gicLabTicket || 'N/A' },
      { label: 'GICLABTASK Ticket #:', valueCol: 2, value: 'N/A' },
      { label: 'Site/Project:', valueCol: 2, value: 'N/A' },
      { label: 'Power Type:', valueCol: 2, value: powerType },
      { label: 'Customer PO / Build Request #:', valueCol: 2, value: workOrderData.custPoNumber || 'N/A' },
      { label: 'Rack ID (If Applicable):', valueCol: 2, value: 'N/A' },
      { label: 'Customer Specific:', valueCol: 2, value: 'N/A' },
      { label: 'Project #', value: '10004703' },
      { label: 'Task #', value: '1002' },
      { label: '# of CarePack Boxes:', value: 'N/A' },
      { label: 'Internal(Int) or External(Ext):', value: 'N/A' }
    ]
  };
}

// Handle SCAN_ORDER message
async function handleScanOrder(message) {
  // 1) Get documentHeaderId
  console.log('Fetching work order info for jobHeaderId:', message.jobHeaderId);
  const workOrderInfoUrl = `https://apirouter.apps.wwt.com/api/forward/mes-api/workOrderInfo?jobHeaderId=${encodeURIComponent(message.jobHeaderId)}`;
  const infoResponse = await fetch(workOrderInfoUrl, {
    method: 'GET',
    credentials: 'include',
    cache: 'no-cache',
    mode: 'cors'
  });
  
  console.log('Work order info response status:', infoResponse.status);
  if (!infoResponse.ok) {
    throw new Error(`Failed to fetch work order info: ${infoResponse.status} ${infoResponse.statusText}`);
  }
  
  const info = await infoResponse.json();
  console.log('Work order info received:', info);
  
  const docHeaderId = info?.documentHeaderId ?? 
    // Deep fallback search for documentHeaderId
    (function deepFind(obj) {
      const key = /document.*header.*id/i;
      const stack = [obj];
      const seen = new Set();
      while (stack.length) {
        const cur = stack.pop();
        if (!cur || typeof cur !== 'object' || seen.has(cur)) continue;
        seen.add(cur);
        for (const [k, v] of Object.entries(cur)) {
          if (key.test(k) && (typeof v === 'number' || /^\d+$/.test(String(v)))) {
            return String(v);
          }
          if (v && typeof v === 'object') stack.push(v);
        }
      }
      return null;
    })(info);
  
  console.log('Found documentHeaderId:', docHeaderId);
  if (!docHeaderId) {
    throw new Error('documentHeaderId not found in workOrderInfo response.');
  }
  
  // 2) Get serial attributes
  const serialAttrsUrl = `https://apirouter.apps.wwt.com/api/forward/serial-attributes/${encodeURIComponent(docHeaderId)}/lab-order?organizationId=${encodeURIComponent(message.orgId || '3346')}&documentType=${encodeURIComponent(message.docType || 'CFG_LTO')}&includeValidations=true&includeAttributes=true`;
  console.log('Fetching serial attributes from:', serialAttrsUrl);
  const serialsResponse = await fetch(serialAttrsUrl, {
    method: 'GET',
    credentials: 'include',
    cache: 'no-cache',
    mode: 'cors'
  });
  
  console.log('Serial attributes response status:', serialsResponse.status);
  if (!serialsResponse.ok) {
    throw new Error(`Failed to fetch serial attributes: ${serialsResponse.status} ${serialsResponse.statusText}`);
  }
  
  const serials = await serialsResponse.json();
  console.log('Serial attributes received, count:', Array.isArray(serials) ? serials.length : 'not array');
  console.log('Returning success response with', Array.isArray(serials) ? serials.length : 0, 'serials');
  
  return { 
    ok: true, 
    docHeaderId, 
    info, 
    serials: Array.isArray(serials) ? serials : [] 
  };
}

console.log('Setting up message listener...');
browser.runtime.onMessage.addListener((message, sender, sendResponse) => {
  console.log('Message received in background:', message);
  
  // Wrapper to handle async messages properly in Manifest V2
  const handleAsync = async () => {
    if (message?.type === 'GET_UPDATE_STATUS') {
      return getSavedUpdateStatus();
    }

    if (message?.type === 'CHECK_FOR_UPDATES') {
      return checkForUpdates(message?.trigger || 'manual');
    }

    if (message?.type === 'GET_IBMOR_SYNC_CONFIG') {
      const config = await getIbmorSyncConfig();
      return {
        ok: true,
        enabled: config.enabled,
        hasToken: Boolean(config.token),
        windowMs: config.windowMs
      };
    }

    if (message?.type === 'SET_IBMOR_SYNC_CONFIG') {
      const config = await setIbmorSyncConfig({
        enabled: message?.enabled,
        token: message?.token
      });

      return {
        ok: true,
        enabled: config.enabled,
        hasToken: Boolean(config.token),
        windowMs: config.windowMs
      };
    }

    if (message?.type === 'GET_IBMOR_SYNC_STATUS') {
      return getIbmorSyncStatus();
    }

    if (message?.type === 'RUN_IBMOR_SYNC_NOW') {
      return runIbmorSheetSync(message?.trigger || 'manual');
    }

    if (message?.type === "download") {
      return new Promise((resolve) => {
        chrome.downloads.download(
          {
            url: message.url,
            filename: message.filename || undefined,
            saveAs: false,
            conflictAction: "uniquify",
          },
          () => {
            resolve({ ok: true });
          }
        );
      });
    }

    if (message?.type === "IBI_TEST_POST") {
      try {
        const response = await fetch(message.url, {
          method: 'POST',
          credentials: 'include',
          headers: {
            'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8'
          },
          body: message.payload || ''
        });

        const text = await response.text().catch(() => '');
        return {
          ok: response.ok,
          status: response.status,
          statusText: response.statusText,
          outputType: response.headers.get('X-IBI-OutputType') || '',
          contentType: response.headers.get('Content-Type') || '',
          bodyText: text,
          bodyPreview: text.slice(0, 500)
        };
      } catch (error) {
        return { ok: false, error: error?.message || String(error) };
      }
    }

    if (message?.type === "IBI_TEST_DOWNLOAD_EXCEL") {
      try {
        const fields = message?.fields && typeof message.fields === 'object' ? message.fields : {};
        const payload = new URLSearchParams(
          Object.entries(fields).reduce((acc, [key, value]) => {
            acc[key] = value == null ? '' : String(value);
            return acc;
          }, {})
        ).toString();

        const response = await fetch(message.url, {
          method: 'POST',
          credentials: 'include',
          headers: {
            'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8'
          },
          body: payload
        });

        const contentType = response.headers.get('Content-Type') || '';
        const outputType = response.headers.get('X-IBI-OutputType') || '';

        if (!response.ok) {
          const errorText = await response.text().catch(() => '');
          return {
            ok: false,
            status: response.status,
            statusText: response.statusText,
            outputType,
            contentType,
            error: errorText.slice(0, 500) || `HTTP ${response.status} ${response.statusText}`
          };
        }

        let downloadResponse = response;
        let downloadContentType = contentType;
        let wrapperPreview = '';

        if (/text\/html/i.test(contentType)) {
          const html = await response.text().catch(() => '');
          wrapperPreview = html.slice(0, 500);

          const redirectMatch = html.match(/location\.replace\((['"])([^'"]+)\1\)/i);
          const binaryPath = redirectMatch?.[2] || '';

          if (binaryPath) {
            const binaryUrl = new URL(binaryPath, message.url).toString();
            downloadResponse = await fetch(binaryUrl, {
              method: 'GET',
              credentials: 'include'
            });
            downloadContentType = downloadResponse.headers.get('Content-Type') || '';

            if (!downloadResponse.ok) {
              const redirectError = await downloadResponse.text().catch(() => '');
              return {
                ok: false,
                status: downloadResponse.status,
                statusText: downloadResponse.statusText,
                outputType,
                contentType: downloadContentType,
                error: redirectError.slice(0, 500) || `GETBINARY failed: HTTP ${downloadResponse.status}`
              };
            }
          } else if (!/EXL07|excel/i.test(outputType)) {
            return {
              ok: false,
              status: response.status,
              statusText: response.statusText,
              outputType,
              contentType,
              error: wrapperPreview || 'Server returned HTML without Excel redirect'
            };
          }
        }

        if (/text\/xml/i.test(downloadContentType) && !/EXL07|excel/i.test(outputType)) {
          const maybeError = await downloadResponse.text().catch(() => '');
          return {
            ok: false,
            status: downloadResponse.status,
            statusText: downloadResponse.statusText,
            outputType,
            contentType: downloadContentType,
            error: maybeError.slice(0, 500) || 'Server returned XML instead of Excel file'
          };
        }

        const blob = await downloadResponse.blob();
        const buffer = await blob.arrayBuffer();
        const bytes = new Uint8Array(buffer);
        let binary = '';
        const chunkSize = 0x8000;
        for (let i = 0; i < bytes.length; i += chunkSize) {
          const chunk = bytes.subarray(i, i + chunkSize);
          binary += String.fromCharCode(...chunk);
        }
        const base64 = btoa(binary);
        const mime = downloadContentType || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
        const dataUrl = `data:${mime};base64,${base64}`;
        const rawName = String(message.filename || 'IBI_Report.xlsx');
        const filename = rawName.toLowerCase().endsWith('.xlsx') ? rawName : `${rawName}.xlsx`;

        const downloadId = await chrome.downloads.download({
          url: dataUrl,
          filename,
          saveAs: false,
          conflictAction: 'uniquify'
        });

        return {
          ok: true,
          status: downloadResponse.status,
          outputType,
          contentType: downloadContentType,
          downloadId,
          filename
        };
      } catch (error) {
        return { ok: false, error: error?.message || String(error) };
      }
    }

    if (message?.type === "SCAN_ORDER") {
      console.log('SCAN_ORDER request received:', message.jobHeaderId);
      try {
        const result = await handleScanOrder(message);
        console.log('handleScanOrder resolved with result:', result);
        return result;
      } catch (e) {
        console.error('handleScanOrder rejected with error:', e);
        return { ok: false, error: e?.message || String(e) };
      }
    }

    if (message?.type === "modifyExcel") {
      try {
        await modifyAndDownloadExcel(message.url, message.extraData, message.filename);
        return { ok: true };
      } catch (error) {
        return { ok: false, error: error.message };
      }
    }

    if (message?.type === "modifyPlacardExcel") {
      console.log('=== RECEIVED modifyPlacardExcel MESSAGE ===');
      console.log('Message data:', JSON.stringify(message, null, 2));
      try {
        console.log('Preparing Excel data...');
        const extraData = await prepareExcelData(message.workOrderData);
        console.log('Excel data prepared:', JSON.stringify(extraData, null, 2));
        console.log('Downloading and modifying Excel from:', message.url);
        await modifyAndDownloadExcel(message.url, extraData, message.filename);
        console.log('=== Excel modified and downloaded successfully ===');
        return { ok: true };
      } catch (error) {
        console.error('=== ERROR modifying placard ===', error);
        console.error('Stack:', error.stack);
        return { ok: false, error: error.message };
      }
    }

    if (message?.type === "OPEN_REPORTAL_INTRO") {
      try {
        const tab = await chrome.tabs.create({
          url: 'https://reports.wwt.com/',
          active: true
        });
        
        // Close the tab after 5 seconds
        setTimeout(() => {
          chrome.tabs.remove(tab.id).catch(err => {
            console.log('Tab already closed or error closing:', err);
          });
        }, 5000);
        
        return { ok: true, tabId: tab.id };
      } catch (error) {
        console.error('Error opening reportal intro tab:', error);
        return { ok: false, error: error?.message || String(error) };
      }
    }
  };

  // Execute async handler and send response
  handleAsync()
    .then(response => {
      if (response !== undefined) {
        console.log('Sending response:', response);
        // Ensure response is serializable by converting to JSON and back
        try {
          const serialized = JSON.parse(JSON.stringify(response));
          sendResponse(serialized);
        } catch (e) {
          console.error('Error serializing response:', e);
          sendResponse({ ok: false, error: 'Failed to serialize response' });
        }
      }
    })
    .catch(error => {
      console.error('Handler error:', error);
      sendResponse({ ok: false, error: error.message || String(error) });
    });

  return true; // Keep channel open for async response
});

async function modifyAndDownloadExcel(url, extraData, filename) {
  try {
    // Check if ExcelJS is available
    if (typeof ExcelJS === 'undefined') {
      console.warn('ExcelJS library not available. Downloading unmodified placard.');
      
      // Fetch the file with credentials first, then download
      const response = await fetch(url, {
        method: 'GET',
        credentials: 'include',
        headers: {
          'Accept': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        }
      });
      
      if (!response.ok) {
        throw new Error(`Failed to fetch Excel file: ${response.status} ${response.statusText}`);
      }
      
      const arrayBuffer = await response.arrayBuffer();
      
      // Convert to base64 data URL (Manifest V3 compatible)
      const base64 = btoa(
        new Uint8Array(arrayBuffer).reduce((data, byte) => data + String.fromCharCode(byte), '')
      );
      const downloadUrl = `data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,${base64}`;
      
      await browser.downloads.download({
        url: downloadUrl,
        filename: filename || 'placard.xlsx',
        saveAs: false,
        conflictAction: 'uniquify'
      });
      
      return;
    }
    
    // Fetch the Excel file with credentials
    const response = await fetch(url, {
      method: 'GET',
      credentials: 'include',
      headers: {
        'Accept': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      }
    });
    
    if (!response.ok) {
      throw new Error(`Failed to fetch Excel file: ${response.status} ${response.statusText}`);
    }
    
    const arrayBuffer = await response.arrayBuffer();

    // Parse the Excel file with ExcelJS (preserves all formatting)
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(arrayBuffer);

    // Get the first sheet
    const worksheet = workbook.worksheets[0];
    console.log('Worksheet loaded:', worksheet.name, 'Rows:', worksheet.rowCount);

    // Update cells in specific rows that contain a label (fully preserves formatting)
    if (extraData && extraData.updateByLabel) {
      console.log('Applying', extraData.updateByLabel.length, 'updates...');
      extraData.updateByLabel.forEach(({ label, valueCol, value }) => {
        console.log('Looking for label:', label, 'to update with:', value);
        let updateCount = 0;
        // Iterate through all rows (including empty ones)
        worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
          // Check each cell in the row for the label
          let labelFound = false;
          row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
            if (!labelFound && cell.value && cell.value.toString().includes(label)) {
              // Found the label
              console.log('Found label "' + label + '" at row', rowNumber, 'col', colNumber);
              labelFound = true;
              let targetCell;
              if (valueCol !== undefined) {
                // Use absolute column position (0-indexed, so add 1 for ExcelJS)
                targetCell = row.getCell(valueCol + 1);
                console.log('Updating absolute column', valueCol + 1, 'with:', value);
              } else {
                // Update the cell to the RIGHT of the label (next column)
                targetCell = row.getCell(colNumber + 1);
                console.log('Updating cell to the right (col', colNumber + 1, ') with:', value);
              }
              targetCell.value = value;
              updateCount++;
              // Formatting is automatically preserved by ExcelJS
            }
          });
        });
        if (updateCount === 0) {
          console.warn('Label "' + label + '" was not found in the worksheet!');
        }
      });
      console.log('All updates applied');
    }

    // Generate modified Excel file
    const buffer = await workbook.xlsx.writeBuffer();
    console.log('Excel buffer generated, size:', buffer.byteLength);

    // Convert buffer to base64 data URL (Manifest V3 compatible)
    const base64 = btoa(
      new Uint8Array(buffer).reduce((data, byte) => data + String.fromCharCode(byte), '')
    );
    const downloadUrl = `data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,${base64}`;

    // Download the modified file
    await browser.downloads.download({
      url: downloadUrl,
      filename: filename || 'modified_placard.xlsx',
      saveAs: false,
      conflictAction: 'uniquify'
    });
    
    console.log('Download initiated successfully');
  } catch (error) {
    console.error('Error modifying Excel:', error);
    throw error;
  }
}
