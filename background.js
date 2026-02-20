// Load ExcelJS library for Excel modification
try {
  importScripts('exceljs.min.js');
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
      console.log('Received modifyPlacardExcel message:', message);
      try {
        console.log('Preparing Excel data...');
        const extraData = await prepareExcelData(message.workOrderData);
        console.log('Excel data prepared:', extraData);
        console.log('Downloading and modifying Excel from:', message.url);
        await modifyAndDownloadExcel(message.url, extraData, message.filename);
        console.log('Excel modified and downloaded successfully');
        return { ok: true };
      } catch (error) {
        console.error('Error modifying placard:', error);
        return { ok: false, error: error.message };
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
      
      // Manifest V2 supports URL.createObjectURL
      const blob = new Blob([arrayBuffer], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      });
      const downloadUrl = URL.createObjectURL(blob);
      
      await browser.downloads.download({
        url: downloadUrl,
        filename: filename || 'placard.xlsx',
        saveAs: false,
        conflictAction: 'uniquify'
      });
      
      // Clean up blob URL
      setTimeout(() => URL.revokeObjectURL(downloadUrl), 1000);
      
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

    // Update cells in specific rows that contain a label (fully preserves formatting)
    if (extraData && extraData.updateByLabel) {
      extraData.updateByLabel.forEach(({ label, valueCol, value }) => {
        // Iterate through all rows
        worksheet.eachRow((row, rowNumber) => {
          // Check each cell in the row for the label
          let labelFound = false;
          row.eachCell((cell, colNumber) => {
            if (!labelFound && cell.value && cell.value.toString().includes(label)) {
              // Found the label
              labelFound = true;
              let targetCell;
              if (valueCol !== undefined) {
                // Use absolute column position (0-indexed, so add 1 for ExcelJS)
                targetCell = row.getCell(valueCol + 1);
              } else {
                // Update the cell to the RIGHT of the label (next column)
                targetCell = row.getCell(colNumber + 1);
              }
              targetCell.value = value;
              // Formatting is automatically preserved by ExcelJS
            }
          });
        });
      });
    }

    // Generate modified Excel file
    const buffer = await workbook.xlsx.writeBuffer();

    // Manifest V2 supports URL.createObjectURL
    const blob = new Blob([buffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    });
    const downloadUrl = URL.createObjectURL(blob);

    // Download the modified file
    await browser.downloads.download({
      url: downloadUrl,
      filename: filename || 'modified_placard.xlsx',
      saveAs: false,
      conflictAction: 'uniquify'
    });
    
    // Clean up blob URL
    setTimeout(() => URL.revokeObjectURL(downloadUrl), 1000);
  } catch (error) {
    console.error('Error modifying Excel:', error);
    throw error;
  }
}
