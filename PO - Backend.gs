// ============================================================
//  PO - Backend.gs
//  Google Apps Script backend for Purchase Order Dashboard
//  Sheet columns (A–P):
//    A = PO Number            B = PO Date
//    C = Expected Delivery    D = Reference#
//    E = Status               F = Vendor Name
//    G = Attention            H = Warehouse Name
//    I = Item Name            J = SKU
//    K = Item Desc            L = Quantity
//    M = Usage Unit           N = Item Price
//    O = Item Total           P = Delivery Instructions
//
//  Multi-item model: multiple rows share the same PO Number.
//  PO-level fields (A–H, P) are duplicated across rows;
//  item-level fields (I–O) are unique per row.
// ============================================================

const SHEET_NAME         = 'PO - Database';
const AUDIT_SHEET_NAME   = 'Audit Log';
const CATALOG_SHEET_NAME = 'Item Catalog';
const VENDOR_CATALOG_SHEET_NAME = 'Vendor Catalog';
const RECEIVING_SHEET_NAME      = 'Receiving Log';
const COMMENTS_SHEET_NAME       = 'PO Comments';
const WAREHOUSE_CATALOG_SHEET_NAME = 'Warehouse Catalog';
const INVENTORY_SHEET_NAME      = 'Inventory';
const TEMPLATES_SHEET_NAME      = 'PO Templates';
const NOTIFICATION_EMAIL = ''; // Set to recipient email, e.g. 'purchasing@soleavenue.com'
                               // If blank, falls back to Session.getActiveUser().getEmail()

// ------------------------------------------------------------
// Router – serves the correct HTML page based on ?page= param
// ------------------------------------------------------------
function doGet(e) {
  try {
    const page = (e && e.parameter && e.parameter.page) ? e.parameter.page : 'form';

    let templateName;
    if (page === 'form') {
      templateName = 'PO - Purchase Order Form';
    } else if (page === 'dashboard') {
      templateName = 'PO - Dashboard';
    } else if (page === 'edit') {
      templateName = 'PO - Edit Purchase Order Form';
    } else if (page === 'orders') {
      templateName = 'PO - Purchase Orders';
    } else if (page === 'receives') {
      templateName = 'PO - Purchase Receives';
    } else if (page === 'inventory') {
      templateName = 'PO - Inventory';
    } else {
      templateName = 'PO - Purchase Order Form';
    }

    const template = HtmlService.createTemplateFromFile(templateName);
    // Embed the real exec URL so HTML pages can navigate without guessing.
    template.scriptUrl = ScriptApp.getService().getUrl();
    // Embed the ?po= param so the Edit page can pre-load a PO on arrival.
    template.poParam   = (e && e.parameter && e.parameter.po) ? e.parameter.po : '';

    return template.evaluate()
      .setTitle('Purchase Order Dashboard')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  } catch (err) {
    const html = `<!DOCTYPE html><html><head><title>Error</title>
      <style>body{font-family:sans-serif;padding:3rem;color:#333;}
      h2{color:#c00;} code{background:#f4f4f4;padding:4px 8px;border-radius:4px;}</style>
      </head><body>
      <h2>Page Load Error</h2>
      <p><strong>Message:</strong> ${err.message}</p>
      <p>Make sure the following HTML files exist in your Apps Script project with <em>exact</em> names:</p>
      <ul>
        <li><code>PO - Dashboard</code></li>
        <li><code>PO - Purchase Orders</code></li>
        <li><code>PO - Purchase Order Form</code></li>
        <li><code>PO - Edit Purchase Order Form</code></li>
        <li><code>PO - Purchase Receives</code></li>
        <li><code>PO - Inventory</code></li>
      </ul>
      <p>After adding or renaming files, create a <strong>new deployment</strong> for the changes to take effect.</p>
      </body></html>`;
    return HtmlService.createHtmlOutput(html).setTitle('Error – PO Dashboard');
  }
}

// ------------------------------------------------------------
// Helper: get the PO - Database sheet
// ------------------------------------------------------------
function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('Sheet "' + SHEET_NAME + '" not found.');
  return sheet;
}

// ------------------------------------------------------------
// Helper: append a row to the Audit Log sheet
// ------------------------------------------------------------
function logAudit_(action, poNumber, details) {
  try {
    const ss  = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(AUDIT_SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(AUDIT_SHEET_NAME);
      sheet.appendRow(['Timestamp', 'Action', 'PO Number', 'User Email', 'Details']);
      sheet.setFrozenRows(1);
    }
    const email = (function() {
      try { return Session.getActiveUser().getEmail() || 'unknown'; }
      catch (e) { return 'unknown'; }
    })();
    sheet.appendRow([new Date(), String(action), String(poNumber || ''), email, String(details || '')]);
  } catch (err) {
    Logger.log('Audit log error: ' + err.message);
  }
}

// ------------------------------------------------------------
// Helper: send an HTML email notification when a PO is Issued
// ------------------------------------------------------------
function sendPOEmail_(poData) {
  try {
    const recipient = NOTIFICATION_EMAIL ||
      (function() { try { return Session.getActiveUser().getEmail(); } catch (e) { return ''; } })();
    if (!recipient) return;

    const itemRows = (poData.items || []).map(function(item) {
      var qty  = parseInt(item.qty) || 0;
      var rate = parseFloat(item.rate) || 0;
      var tot  = parseFloat(item.total) || (rate * qty);
      return '<tr>' +
        '<td style="padding:6px 10px;border-bottom:1px solid #eee;">' + (item.itemName || '') + '</td>' +
        '<td style="padding:6px 10px;border-bottom:1px solid #eee;">' + (item.sku || '') + '</td>' +
        '<td style="padding:6px 10px;border-bottom:1px solid #eee;text-align:center;">' + qty + '</td>' +
        '<td style="padding:6px 10px;border-bottom:1px solid #eee;text-align:center;">' + (item.usageUnit || '') + '</td>' +
        '<td style="padding:6px 10px;border-bottom:1px solid #eee;text-align:right;">\u20B1' + rate.toFixed(2) + '</td>' +
        '<td style="padding:6px 10px;border-bottom:1px solid #eee;text-align:right;">\u20B1' + tot.toFixed(2) + '</td>' +
        '</tr>';
    }).join('');

    var grandTotal = (poData.items || []).reduce(function(s, it) {
      return s + (parseFloat(it.total) || (parseFloat(it.rate) || 0) * (parseInt(it.qty) || 0));
    }, 0);

    var htmlBody =
      '<div style="font-family:sans-serif;max-width:700px;margin:0 auto;">' +
        '<div style="background:#000;color:#fff;padding:20px 24px;border-radius:8px 8px 0 0;">' +
          '<h1 style="margin:0;font-size:22px;letter-spacing:1px;">SOLE AVENUE</h1>' +
          '<p style="margin:4px 0 0;font-size:13px;opacity:0.7;">Purchase Order Notification</p>' +
        '</div>' +
        '<div style="background:#f8f8f8;padding:20px 24px;border:1px solid #eee;">' +
          '<h2 style="margin:0 0 16px;font-size:16px;color:#333;">PO <strong>' + poData.po + '</strong> has been <span style="color:#1a6b3c;">Issued</span></h2>' +
          '<table style="width:100%;border-collapse:collapse;margin-bottom:12px;font-size:13px;">' +
            '<tr><td style="padding:4px 0;color:#666;width:150px;">Vendor</td><td style="font-weight:600;">' + (poData.vendor || '') + '</td></tr>' +
            '<tr><td style="padding:4px 0;color:#666;">PO Date</td><td>' + (poData.poDate || '') + '</td></tr>' +
            '<tr><td style="padding:4px 0;color:#666;">Expected Delivery</td><td>' + (poData.deliveryDate || '\u2014') + '</td></tr>' +
            '<tr><td style="padding:4px 0;color:#666;">Warehouse</td><td>' + (poData.warehouse || '\u2014') + '</td></tr>' +
          '</table>' +
          '<table style="width:100%;border-collapse:collapse;font-size:13px;">' +
            '<thead><tr style="background:#f0f0f0;">' +
              '<th style="padding:8px 10px;text-align:left;">Item</th>' +
              '<th style="padding:8px 10px;text-align:left;">SKU</th>' +
              '<th style="padding:8px 10px;text-align:center;">Qty</th>' +
              '<th style="padding:8px 10px;text-align:center;">Unit</th>' +
              '<th style="padding:8px 10px;text-align:right;">Price</th>' +
              '<th style="padding:8px 10px;text-align:right;">Total</th>' +
            '</tr></thead>' +
            '<tbody>' + itemRows + '</tbody>' +
          '</table>' +
          '<div style="text-align:right;margin-top:12px;font-size:15px;font-weight:700;">Grand Total: \u20B1' + grandTotal.toFixed(2) + '</div>' +
        '</div>' +
        '<div style="background:#eee;padding:10px 24px;font-size:11px;color:#888;border-radius:0 0 8px 8px;">Automated notification from Sole Avenue PO System.</div>' +
      '</div>';

    MailApp.sendEmail({
      to:       recipient,
      subject:  'PO Issued: ' + poData.po + ' \u2014 ' + (poData.vendor || '') + ' (Sole Avenue)',
      htmlBody: htmlBody
    });
  } catch (err) {
    Logger.log('Email notification error: ' + err.message);
  }
}

// ------------------------------------------------------------
// Generate next unique PO number in format PO-XXXXX
// ------------------------------------------------------------
function generatePONumber() {
  try {
    const sheet   = getSheet_();
    const lastRow = sheet.getLastRow();
    let maxNum    = 0;

    if (lastRow > 1) {
      const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      data.forEach(function(row) {
        const match = String(row[0]).match(/(\d+)$/);
        if (match) {
          const num = parseInt(match[1], 10);
          if (num > maxNum) maxNum = num;
        }
      });
    }

    return 'PO-' + String(maxNum + 1).padStart(5, '0');
  } catch (err) {
    return 'PO-00001';
  }
}

// ------------------------------------------------------------
// Helper: format a YYYY-MM-DD string to MM/dd/yyyy
// ------------------------------------------------------------
function formatDeliveryDate_(dateStr) {
  if (!dateStr) return '';
  try {
    return Utilities.formatDate(
      new Date(dateStr + 'T00:00:00'),
      Session.getScriptTimeZone(),
      'MM/dd/yyyy'
    );
  } catch (e) {
    return String(dateStr);
  }
}

// ------------------------------------------------------------
// Helper: map a raw 16-element row to a PO object (used internally)
// ------------------------------------------------------------
function rowToObject_(row) {
  return {
    po:                   String(row[0]  || ''),
    poDate:               row[1] instanceof Date
                            ? Utilities.formatDate(row[1], Session.getScriptTimeZone(), 'MM/dd/yyyy')
                            : String(row[1]  || ''),
    deliveryDate:         row[2] instanceof Date
                            ? Utilities.formatDate(row[2], Session.getScriptTimeZone(), 'MM/dd/yyyy')
                            : String(row[2]  || ''),
    reference:            String(row[3]  || ''),
    status:               String(row[4]  || ''),
    vendor:               String(row[5]  || ''),
    attention:            String(row[6]  || ''),
    warehouse:            String(row[7]  || ''),
    itemName:             String(row[8]  || ''),
    sku:                  String(row[9]  || ''),
    itemDesc:             String(row[10] || ''),
    qty:                  row[11],
    usageUnit:            String(row[12] || ''),
    rate:                 row[13],
    total:                row[14],
    deliveryInstructions: String(row[15] || '')
  };
}

// ------------------------------------------------------------
// Helper: extract item-level fields from a row (cols I–O)
// ------------------------------------------------------------
function rowToItemObject_(row) {
  return {
    itemName:  String(row[8]  || ''),
    sku:       String(row[9]  || ''),
    itemDesc:  String(row[10] || ''),
    qty:       parseInt(row[11])   || 0,
    usageUnit: String(row[12] || ''),
    rate:      parseFloat(row[13]) || 0,
    total:     parseFloat(row[14]) || 0
  };
}

// ------------------------------------------------------------
// Helper: group raw sheet rows into one object per PO number.
// Preserves insertion order; accumulates items[].
// ------------------------------------------------------------
function groupRowsByPO_(rows) {
  const orderMap = {};
  const result   = [];

  rows.forEach(function(row) {
    const poNum = String(row[0] || '').trim();
    if (!poNum) return;

    if (orderMap[poNum] === undefined) {
      orderMap[poNum] = result.length;
      result.push({
        po:      poNum,
        poDate:  row[1] instanceof Date
                   ? Utilities.formatDate(row[1], Session.getScriptTimeZone(), 'MM/dd/yyyy')
                   : String(row[1] || ''),
        deliveryDate: row[2] instanceof Date
                   ? Utilities.formatDate(row[2], Session.getScriptTimeZone(), 'MM/dd/yyyy')
                   : String(row[2] || ''),
        reference:            String(row[3]  || ''),
        status:               String(row[4]  || ''),
        vendor:               String(row[5]  || ''),
        attention:            String(row[6]  || ''),
        warehouse:            String(row[7]  || ''),
        deliveryInstructions: String(row[15] || ''),
        items:             [],
        poTotal:           0,
        totalQtyOrdered:   0,
        qtyReceived:       0
      });
    }

    const idx  = orderMap[poNum];
    const item = rowToItemObject_(row);
    result[idx].items.push(item);
    result[idx].poTotal         += item.total;
    result[idx].totalQtyOrdered += item.qty;
  });

  result.forEach(function(po) {
    po.itemCount = po.items.length;
  });

  return result;
}

// ------------------------------------------------------------
// CREATE – Append new PO rows (one per item)
// formData shape:
//   { po, deliveryDate, reference, status, vendor, attention,
//     warehouse, deliveryInstructions,
//     items: [{itemName, sku, itemDesc, qty, usageUnit, rate}] }
// ------------------------------------------------------------
function processForm(formData) {
  try {
    const sheet = getSheet_();

    // Duplicate check
    const existing = searchPO(formData.po);
    if (existing) return 'Duplicate: PO "' + formData.po + '" already exists.';

    // PO-level required field validation
    if (!String(formData.po     || '').trim()) return 'Error: PO number is missing.';
    if (!String(formData.vendor || '').trim()) return 'Error: Vendor is required.';
    if (!String(formData.status || '').trim()) return 'Error: Order status is required.';

    // Items array validation
    const items = formData.items;
    if (!Array.isArray(items) || items.length === 0) {
      return 'Error: At least one item is required.';
    }
    for (let i = 0; i < items.length; i++) {
      const item = items[i];
      const n    = i + 1;
      if (!String(item.itemName || '').trim()) return 'Error: Item ' + n + ' name is required.';
      const qty  = parseInt(item.qty)    || 0;
      const rate = parseFloat(item.rate) || 0;
      if (qty  < 1)  return 'Error: Item ' + n + ' quantity must be at least 1.';
      if (rate <= 0) return 'Error: Item ' + n + ' price must be greater than 0.';
    }

    const poDate    = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MM/dd/yyyy');
    const delivDate = formatDeliveryDate_(formData.deliveryDate);

    // Append one sheet row per item
    items.forEach(function(item) {
      const qty  = parseInt(item.qty)    || 0;
      const rate = parseFloat(item.rate) || 0;
      sheet.appendRow([
        String(formData.po).trim(),                       // A: PO Number
        poDate,                                            // B: PO Date (auto)
        delivDate,                                         // C: Expected Delivery
        String(formData.reference    || ''),               // D: Reference#
        String(formData.status),                           // E: Status
        String(formData.vendor).trim(),                    // F: Vendor Name
        String(formData.attention    || ''),               // G: Attention
        String(formData.warehouse    || ''),               // H: Warehouse Name
        String(item.itemName).trim(),                      // I: Item Name
        String(item.sku              || ''),               // J: SKU
        String(item.itemDesc         || ''),               // K: Item Desc
        qty,                                               // L: Quantity
        String(item.usageUnit        || ''),               // M: Usage Unit
        rate,                                              // N: Item Price
        rate * qty,                                        // O: Item Total (server-calculated)
        String(formData.deliveryInstructions || '')        // P: Delivery Instructions
      ]);
    });

    // Audit log
    logAudit_('Created', String(formData.po).trim(),
      'Vendor: ' + String(formData.vendor).trim() + ', Items: ' + items.length + ', Status: ' + String(formData.status));

    // Email notification if status is Issued
    if (String(formData.status) === 'Issued') {
      sendPOEmail_({
        po: String(formData.po).trim(), poDate: poDate, deliveryDate: delivDate,
        vendor: String(formData.vendor).trim(), warehouse: String(formData.warehouse || ''),
        items: items.map(function(item) {
          var q = parseInt(item.qty) || 0, r = parseFloat(item.rate) || 0;
          return { itemName: String(item.itemName||'').trim(), sku: String(item.sku||''), qty: q, usageUnit: String(item.usageUnit||''), rate: r, total: r * q };
        })
      });
    }

    // Auto-save items to catalog (fire-and-forget)
    try {
      items.forEach(function(item) {
        if (String(item.itemName || '').trim()) {
          addCatalogItem({ itemName: String(item.itemName).trim(), sku: String(item.sku||''), itemDesc: String(item.itemDesc||''), usageUnit: String(item.usageUnit||''), defaultRate: parseFloat(item.rate) || 0 });
        }
      });
    } catch (e) {}

    // Auto-save vendor to vendor catalog (fire-and-forget)
    try {
      var vName = String(formData.vendor || '').trim();
      if (vName) addVendorCatalogItem({ vendorName: vName, contactPerson: String(formData.attention || '') });
    } catch (e) {}

    // Auto-save warehouse to warehouse catalog (fire-and-forget)
    try {
      var wName = String(formData.warehouse || '').trim();
      if (wName) addWarehouseCatalogItem({ warehouseName: wName });
    } catch (e) {}

    return 'Success';
  } catch (err) {
    return 'Error: ' + err.message;
  }
}

// ------------------------------------------------------------
// READ – Return all POs grouped by PO number
// ------------------------------------------------------------
function getAllPOs() {
  try {
    const sheet   = getSheet_();
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return [];

    const data = sheet.getRange(2, 1, lastRow - 1, 16).getValues();
    const rows = data.filter(row => String(row[0]).trim() !== '');
    var pos = groupRowsByPO_(rows);

    // Join per-PO received quantities from the Receiving Log
    try {
      var recSheet = getReceivingSheet_();
      var recLast  = recSheet.getLastRow();
      if (recLast > 1) {
        var recData = recSheet.getRange(2, 1, recLast - 1, 5).getValues();
        // Build lookup: PO → {itemKey: orderedQty}
        var poItemQty = {};
        pos.forEach(function(po) {
          var qty = {};
          po.items.forEach(function(it) {
            var k = String(it.itemName||'').trim().toLowerCase()+'||'+String(it.sku||'').trim().toLowerCase();
            qty[k] = (qty[k] || 0) + (parseInt(it.qty) || 0);
          });
          poItemQty[po.po.toUpperCase()] = qty;
        });
        var recMap = {};
        var recItemMap = {};
        recData.forEach(function(row) {
          var po = String(row[1]).trim().toUpperCase();
          if (!po) return;
          var qty = parseInt(row[4]) || 0;
          recMap[po] = (recMap[po] || 0) + qty;
          var k = String(row[2]||'').trim().toLowerCase()+'||'+String(row[3]||'').trim().toLowerCase();
          if (!recItemMap[po]) recItemMap[po] = {};
          recItemMap[po][k] = (recItemMap[po][k] || 0) + qty;
        });
        pos.forEach(function(po) {
          var poKey = po.po.toUpperCase();
          po.qtyReceived = recMap[poKey] || 0;
          // Count SKUs received that are not on the PO at all
          var orderedQty = poItemQty[poKey] || {};
          var receivedItems = recItemMap[poKey] || {};
          var unexpectedCount = 0;
          Object.keys(receivedItems).forEach(function(k) {
            if (!orderedQty.hasOwnProperty(k)) {
              unexpectedCount++;
            }
          });
          po.unexpectedCount = unexpectedCount;
        });
      }
    } catch (e) {}

    return pos;
  } catch (err) {
    return [];
  }
}

// ------------------------------------------------------------
// READ – Find a single PO by PO number (case-insensitive)
//        Returns a grouped object with items[] array, or null.
// ------------------------------------------------------------
function searchPO(poNumber) {
  try {
    const sheet   = getSheet_();
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return null;

    const data   = sheet.getRange(2, 1, lastRow - 1, 16).getValues();
    const target = String(poNumber).toLowerCase().trim();

    const matchingRows = data.filter(function(row) {
      return String(row[0]).toLowerCase().trim() === target;
    });

    if (matchingRows.length === 0) return null;

    const grouped = groupRowsByPO_(matchingRows);
    return grouped[0] || null;
  } catch (err) {
    return null;
  }
}

// ------------------------------------------------------------
// UPDATE – Delete all rows for the PO then re-append all items.
//          PO Date (col B) is preserved from the original rows.
//          formData shape same as processForm but with originalPo.
// ------------------------------------------------------------
function updatePO(formData) {
  try {
    const sheet   = getSheet_();
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return 'PO not found';

    // Read cols A–B to locate rows and capture original PO Date
    const data      = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    const searchKey = String(formData.originalPo || formData.po).toLowerCase().trim();

    const rowIndices    = [];
    let existingPoDate  = null;

    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]).toLowerCase().trim() === searchKey) {
        rowIndices.push(i + 2); // 1-based sheet row
        if (existingPoDate === null) existingPoDate = data[i][1];
      }
    }

    if (rowIndices.length === 0) return 'PO not found';

    // Capture old status before deleting rows (for email trigger)
    var oldStatus = '';
    try { oldStatus = String(sheet.getRange(rowIndices[0], 5).getValue() || ''); } catch (e) {}

    // Validate items before touching the sheet
    const items = formData.items;
    if (!Array.isArray(items) || items.length === 0) {
      return 'Error: At least one item is required.';
    }

    const delivDate = formatDeliveryDate_(formData.deliveryDate);

    // Delete existing rows in reverse order to prevent row-index shifting
    for (let r = rowIndices.length - 1; r >= 0; r--) {
      sheet.deleteRow(rowIndices[r]);
    }

    // Re-append all items as new rows at the bottom
    items.forEach(function(item) {
      const qty  = parseInt(item.qty)    || 0;
      const rate = parseFloat(item.rate) || 0;
      sheet.appendRow([
        String(formData.po).trim(),
        existingPoDate,                                    // preserve original PO Date
        delivDate,
        String(formData.reference    || ''),
        String(formData.status),
        String(formData.vendor),
        String(formData.attention    || ''),
        String(formData.warehouse    || ''),
        String(item.itemName).trim(),
        String(item.sku              || ''),
        String(item.itemDesc         || ''),
        qty,
        String(item.usageUnit        || ''),
        rate,
        rate * qty,                                        // server-recalculated
        String(formData.deliveryInstructions || '')
      ]);
    });

    // Audit log
    logAudit_('Updated', String(formData.po).trim(),
      'Vendor: ' + String(formData.vendor).trim() + ', Items: ' + items.length + ', Status: ' + String(formData.status));

    // Auto-save vendor to vendor catalog (fire-and-forget)
    try {
      var vName = String(formData.vendor || '').trim();
      if (vName) addVendorCatalogItem({ vendorName: vName, contactPerson: String(formData.attention || '') });
    } catch (e) {}

    // Auto-save warehouse to warehouse catalog (fire-and-forget)
    try {
      var wName = String(formData.warehouse || '').trim();
      if (wName) addWarehouseCatalogItem({ warehouseName: wName });
    } catch (e) {}

    // Email notification if status changed to Issued
    if (String(formData.status) === 'Issued' && oldStatus !== 'Issued') {
      var fmtPoDate = existingPoDate instanceof Date
        ? Utilities.formatDate(existingPoDate, Session.getScriptTimeZone(), 'MM/dd/yyyy')
        : String(existingPoDate || '');
      sendPOEmail_({
        po: String(formData.po).trim(), poDate: fmtPoDate, deliveryDate: delivDate,
        vendor: String(formData.vendor).trim(), warehouse: String(formData.warehouse || ''),
        items: items.map(function(item) {
          var q = parseInt(item.qty) || 0, r = parseFloat(item.rate) || 0;
          return { itemName: String(item.itemName||'').trim(), sku: String(item.sku||''), qty: q, usageUnit: String(item.usageUnit||''), rate: r, total: r * q };
        })
      });
    }

    return 'Success';
  } catch (err) {
    return 'Error: ' + err.message;
  }
}

// ------------------------------------------------------------
// DELETE – Remove ALL rows for a PO number
// ------------------------------------------------------------
function deletePO(poNumber) {
  try {
    const sheet   = getSheet_();
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return 'PO not found';

    const data   = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    const target = String(poNumber).toLowerCase().trim();

    const toDelete = [];
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]).toLowerCase().trim() === target) {
        toDelete.push(i + 2);
      }
    }

    if (toDelete.length === 0) return 'PO not found';

    // Delete in reverse order to avoid row-index shifting
    for (let r = toDelete.length - 1; r >= 0; r--) {
      sheet.deleteRow(toDelete[r]);
    }

    logAudit_('Deleted', String(poNumber).trim(), 'Rows removed: ' + toDelete.length);
    return 'Success';
  } catch (err) {
    return 'Error: ' + err.message;
  }
}

// ------------------------------------------------------------
// STATUS UPDATE – Change only the Status column (col E) for a PO.
//                Much faster than full updatePO (reads 5 cols, writes 1).
// ------------------------------------------------------------
function updatePOStatus(poNumber, newStatus) {
  try {
    var VALID = ['Draft', 'Issued', 'Partial', 'Closed'];
    if (VALID.indexOf(newStatus) === -1) return 'Error: Invalid status "' + newStatus + '".';

    var sheet   = getSheet_();
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return 'PO not found';

    var data   = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
    var target = String(poNumber).toLowerCase().trim();
    var oldStatus   = '';
    var rowsUpdated = 0;

    for (var i = 0; i < data.length; i++) {
      if (String(data[i][0]).toLowerCase().trim() === target) {
        if (!oldStatus) oldStatus = String(data[i][4] || '');
        sheet.getRange(i + 2, 5).setValue(newStatus);
        rowsUpdated++;
      }
    }

    if (rowsUpdated === 0) return 'PO not found';

    logAudit_('StatusChanged', String(poNumber).trim(), oldStatus + ' \u2192 ' + newStatus);

    if (newStatus === 'Issued' && oldStatus !== 'Issued') {
      var poData = searchPO(poNumber);
      if (poData) sendPOEmail_(poData);
    }

    return 'Success';
  } catch (err) {
    return 'Error: ' + err.message;
  }
}

// ------------------------------------------------------------
// Helper: get or create the Item Catalog sheet
// Columns: A=Item Name, B=SKU, C=Description, D=Default Unit, E=Default Price
// ------------------------------------------------------------
function getCatalogSheet_() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CATALOG_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CATALOG_SHEET_NAME);
    sheet.appendRow(['Item Name', 'SKU', 'Description', 'Default Unit', 'Default Price']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// ------------------------------------------------------------
// READ – Return all items from the Item Catalog
// ------------------------------------------------------------
function getItemCatalog() {
  try {
    var sheet   = getCatalogSheet_();
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return [];
    var data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
    return data
      .filter(function(row) { return String(row[0]).trim() !== ''; })
      .map(function(row) {
        return {
          itemName:    String(row[0] || ''),
          sku:         String(row[1] || ''),
          itemDesc:    String(row[2] || ''),
          usageUnit:   String(row[3] || ''),
          defaultRate: parseFloat(row[4]) || 0
        };
      });
  } catch (err) {
    return [];
  }
}

// ------------------------------------------------------------
// CREATE – Add item to catalog (no duplicate names)
// ------------------------------------------------------------
function addCatalogItem(data) {
  try {
    var sheet   = getCatalogSheet_();
    var lastRow = sheet.getLastRow();
    var name    = String(data.itemName || '').trim();
    if (!name) return 'Error: Item name is required.';
    if (lastRow > 1) {
      var existing = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      var isDupe   = existing.some(function(row) {
        return String(row[0]).toLowerCase().trim() === name.toLowerCase();
      });
      if (isDupe) return 'Duplicate';
    }
    sheet.appendRow([name, String(data.sku || ''), String(data.itemDesc || ''), String(data.usageUnit || ''), parseFloat(data.defaultRate) || 0]);
    return 'Success';
  } catch (err) {
    return 'Error: ' + err.message;
  }
}

// ------------------------------------------------------------
// Helper: get or create the Vendor Catalog sheet
// Columns: A=Vendor Name, B=Contact Person, C=Email, D=Phone, E=Address
// ------------------------------------------------------------
function getVendorCatalogSheet_() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(VENDOR_CATALOG_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(VENDOR_CATALOG_SHEET_NAME);
    sheet.appendRow(['Vendor Name', 'Contact Person', 'Email', 'Phone', 'Address']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// ------------------------------------------------------------
// READ – Return all vendors from the Vendor Catalog
// ------------------------------------------------------------
function getVendorCatalog() {
  try {
    var sheet   = getVendorCatalogSheet_();
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return [];
    var data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
    return data
      .filter(function(row) { return String(row[0]).trim() !== ''; })
      .map(function(row) {
        return {
          vendorName:    String(row[0] || ''),
          contactPerson: String(row[1] || ''),
          email:         String(row[2] || ''),
          phone:         String(row[3] || ''),
          address:       String(row[4] || '')
        };
      });
  } catch (err) {
    return [];
  }
}

// ------------------------------------------------------------
// CREATE – Add vendor to catalog (no duplicate names)
// ------------------------------------------------------------
function addVendorCatalogItem(data) {
  try {
    var sheet   = getVendorCatalogSheet_();
    var lastRow = sheet.getLastRow();
    var name    = String(data.vendorName || '').trim();
    if (!name) return 'Error: Vendor name is required.';
    if (lastRow > 1) {
      var existing = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      var isDupe   = existing.some(function(row) {
        return String(row[0]).toLowerCase().trim() === name.toLowerCase();
      });
      if (isDupe) return 'Duplicate';
    }
    sheet.appendRow([
      name,
      String(data.contactPerson || ''),
      String(data.email || ''),
      String(data.phone || ''),
      String(data.address || '')
    ]);
    return 'Success';
  } catch (err) {
    return 'Error: ' + err.message;
  }
}

// ============================================================
//  RECEIVING / PARTIAL DELIVERY TRACKING
// ============================================================

function getReceivingSheet_() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(RECEIVING_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(RECEIVING_SHEET_NAME);
    sheet.appendRow(['Timestamp', 'PO Number', 'Item Name', 'SKU', 'Qty Received', 'Received By', 'Notes']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// Record receiving for items — auto-closes PO when fully received
function recordReceiving(data) {
  try {
    var poNumber = String(data.poNumber || '').trim();
    if (!poNumber) return 'Error: PO number is required.';
    var items = data.items || [];
    if (!items.length) return 'Error: No items to receive.';

    var sheet = getReceivingSheet_();
    var now   = new Date();
    var user  = '';
    try { user = Session.getActiveUser().getEmail(); } catch (e) {}
    var notes = String(data.notes || '');

    items.forEach(function(item) {
      var qty = parseInt(item.qtyReceived) || 0;
      if (qty <= 0) return;
      sheet.appendRow([now, poNumber, String(item.itemName || ''), String(item.sku || ''), qty, user, notes]);
      // Update inventory
      try { updateInventoryFromReceiving_(String(item.itemName || ''), String(item.sku || ''), qty); } catch (e) {}
    });

    logAudit_('Receiving', poNumber, 'Received ' + items.length + ' item(s)');

    // Check receiving progress → auto-set Partial or Closed
    try {
      var summary = getReceivingSummary(poNumber);
      var allReceived = summary.every(function(s) { return s.remaining <= 0; });
      if (allReceived && summary.length > 0) {
        updatePOStatus(poNumber, 'Closed');
      } else if (summary.some(function(s) { return s.qtyReceived > 0; })) {
        // At least one item received but not all → Partial (unless already Closed)
        var currentPO = searchPO(poNumber);
        if (currentPO && currentPO.status !== 'Closed') {
          updatePOStatus(poNumber, 'Partial');
        }
      }
    } catch (e) {}

    return 'Success';
  } catch (err) {
    return 'Error: ' + err.message;
  }
}

// Get per-item receiving summary for a PO
function getReceivingSummary(poNumber) {
  try {
    var po = String(poNumber || '').trim().toUpperCase();

    // Get ordered items from PO Database
    var dbSheet = getSheet_();
    var dbLast  = dbSheet.getLastRow();
    var ordered = {};
    if (dbLast > 1) {
      var dbData = dbSheet.getRange(2, 1, dbLast - 1, 16).getValues();
      dbData.forEach(function(row) {
        if (String(row[0]).trim().toUpperCase() === po) {
          var key = String(row[8] || '') + '||' + String(row[9] || '');
          if (!ordered[key]) {
            ordered[key] = { itemName: String(row[8] || ''), sku: String(row[9] || ''), qtyOrdered: 0, qtyReceived: 0 };
          }
          ordered[key].qtyOrdered += parseInt(row[11]) || 0;
        }
      });
    }

    // Sum received from Receiving Log
    var recSheet = getReceivingSheet_();
    var recLast  = recSheet.getLastRow();
    if (recLast > 1) {
      var recData = recSheet.getRange(2, 1, recLast - 1, 5).getValues();
      recData.forEach(function(row) {
        if (String(row[1]).trim().toUpperCase() === po) {
          var key = String(row[2] || '') + '||' + String(row[3] || '');
          if (ordered[key]) {
            ordered[key].qtyReceived += parseInt(row[4]) || 0;
          }
        }
      });
    }

    return Object.keys(ordered).map(function(key) {
      var item = ordered[key];
      return {
        itemName:    item.itemName,
        sku:         item.sku,
        qtyOrdered:  item.qtyOrdered,
        qtyReceived: item.qtyReceived,
        remaining:   item.qtyOrdered - item.qtyReceived
      };
    });
  } catch (err) {
    return [];
  }
}

// Get receiving history for a PO
function getReceivingHistory(poNumber) {
  try {
    var po    = String(poNumber || '').trim().toUpperCase();
    var sheet = getReceivingSheet_();
    var last  = sheet.getLastRow();
    if (last <= 1) return [];

    var data    = sheet.getRange(2, 1, last - 1, 7).getValues();
    var results = [];
    data.forEach(function(row) {
      if (String(row[1]).trim().toUpperCase() === po) {
        results.push({
          timestamp:   row[0] instanceof Date ? Utilities.formatDate(row[0], Session.getScriptTimeZone(), 'MM/dd/yyyy HH:mm') : String(row[0]),
          itemName:    String(row[2] || ''),
          sku:         String(row[3] || ''),
          qtyReceived: parseInt(row[4]) || 0,
          receivedBy:  String(row[5] || ''),
          notes:       String(row[6] || '')
        });
      }
    });
    return results.reverse(); // newest first
  } catch (err) {
    return [];
  }
}

// ============================================================
//  PO COMMENTS / NOTES
// ============================================================

function getCommentsSheet_() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(COMMENTS_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(COMMENTS_SHEET_NAME);
    sheet.appendRow(['Timestamp', 'PO Number', 'User Email', 'Comment']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function addComment(poNumber, comment) {
  try {
    var po  = String(poNumber || '').trim();
    var txt = String(comment || '').trim();
    if (!po)  return 'Error: PO number is required.';
    if (!txt) return 'Error: Comment cannot be empty.';

    var sheet = getCommentsSheet_();
    var user  = '';
    try { user = Session.getActiveUser().getEmail(); } catch (e) {}

    sheet.appendRow([new Date(), po, user, txt]);
    return 'Success';
  } catch (err) {
    return 'Error: ' + err.message;
  }
}

function getComments(poNumber) {
  try {
    var po    = String(poNumber || '').trim().toUpperCase();
    var sheet = getCommentsSheet_();
    var last  = sheet.getLastRow();
    if (last <= 1) return [];

    var data    = sheet.getRange(2, 1, last - 1, 4).getValues();
    var results = [];
    data.forEach(function(row) {
      if (String(row[1]).trim().toUpperCase() === po) {
        results.push({
          timestamp: row[0] instanceof Date ? Utilities.formatDate(row[0], Session.getScriptTimeZone(), 'MM/dd/yyyy HH:mm') : String(row[0]),
          email:     String(row[2] || ''),
          comment:   String(row[3] || '')
        });
      }
    });
    return results.reverse(); // newest first
  } catch (err) {
    return [];
  }
}

// ============================================================
//  WAREHOUSE CATALOG
// ============================================================

function getWarehouseCatalogSheet_() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(WAREHOUSE_CATALOG_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(WAREHOUSE_CATALOG_SHEET_NAME);
    sheet.appendRow(['Warehouse Name', 'Address', 'Contact Person', 'Phone']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getWarehouseCatalog() {
  try {
    var sheet   = getWarehouseCatalogSheet_();
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return [];
    var data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
    return data
      .filter(function(row) { return String(row[0]).trim() !== ''; })
      .map(function(row) {
        return {
          warehouseName: String(row[0] || ''),
          address:       String(row[1] || ''),
          contactPerson: String(row[2] || ''),
          phone:         String(row[3] || '')
        };
      });
  } catch (err) {
    return [];
  }
}

function addWarehouseCatalogItem(data) {
  try {
    var sheet   = getWarehouseCatalogSheet_();
    var lastRow = sheet.getLastRow();
    var name    = String(data.warehouseName || '').trim();
    if (!name) return 'Error: Warehouse name is required.';
    if (lastRow > 1) {
      var existing = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      var isDupe   = existing.some(function(row) {
        return String(row[0]).toLowerCase().trim() === name.toLowerCase();
      });
      if (isDupe) return 'Duplicate';
    }
    sheet.appendRow([name, String(data.address || ''), String(data.contactPerson || ''), String(data.phone || '')]);
    return 'Success';
  } catch (err) {
    return 'Error: ' + err.message;
  }
}

// ============================================================
//  INVENTORY / LOW STOCK
// ============================================================

function getInventorySheet_() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(INVENTORY_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(INVENTORY_SHEET_NAME);
    sheet.appendRow(['Item Name', 'SKU', 'Current Stock', 'Reorder Level', 'Last Updated']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getInventory() {
  try {
    var sheet   = getInventorySheet_();
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return [];
    var data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
    return data
      .filter(function(row) { return String(row[0]).trim() !== ''; })
      .map(function(row) {
        return {
          itemName:     String(row[0] || ''),
          sku:          String(row[1] || ''),
          currentStock: parseInt(row[2]) || 0,
          reorderLevel: parseInt(row[3]) || 0,
          lastUpdated:  row[4] instanceof Date ? Utilities.formatDate(row[4], Session.getScriptTimeZone(), 'MM/dd/yyyy HH:mm') : String(row[4] || '')
        };
      });
  } catch (err) {
    return [];
  }
}

function updateInventoryItem(data) {
  try {
    var sheet   = getInventorySheet_();
    var lastRow = sheet.getLastRow();
    var name    = String(data.itemName || '').trim();
    var sku     = String(data.sku || '').trim();
    if (!name) return 'Error: Item name is required.';

    // Find existing row
    if (lastRow > 1) {
      var existing = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
      for (var i = 0; i < existing.length; i++) {
        if (String(existing[i][0]).toLowerCase().trim() === name.toLowerCase() &&
            String(existing[i][1]).toLowerCase().trim() === sku.toLowerCase()) {
          var rowIdx = i + 2;
          sheet.getRange(rowIdx, 3).setValue(parseInt(data.currentStock) || 0);
          sheet.getRange(rowIdx, 4).setValue(parseInt(data.reorderLevel) || 0);
          sheet.getRange(rowIdx, 5).setValue(new Date());
          return 'Success';
        }
      }
    }

    // New row
    sheet.appendRow([name, sku, parseInt(data.currentStock) || 0, parseInt(data.reorderLevel) || 0, new Date()]);
    return 'Success';
  } catch (err) {
    return 'Error: ' + err.message;
  }
}

function getLowStockItems() {
  try {
    var items = getInventory();
    return items.filter(function(item) {
      return item.reorderLevel > 0 && item.currentStock <= item.reorderLevel;
    });
  } catch (err) {
    return [];
  }
}

// Private: update inventory when receiving items
function updateInventoryFromReceiving_(itemName, sku, qtyReceived) {
  var sheet   = getInventorySheet_();
  var lastRow = sheet.getLastRow();
  var name    = String(itemName || '').trim();
  var skuStr  = String(sku || '').trim();

  if (lastRow > 1) {
    var existing = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
    for (var i = 0; i < existing.length; i++) {
      if (String(existing[i][0]).toLowerCase().trim() === name.toLowerCase() &&
          String(existing[i][1]).toLowerCase().trim() === skuStr.toLowerCase()) {
        var rowIdx = i + 2;
        var current = parseInt(existing[i][2]) || 0;
        sheet.getRange(rowIdx, 3).setValue(current + qtyReceived);
        sheet.getRange(rowIdx, 5).setValue(new Date());
        return;
      }
    }
  }

  // New inventory entry
  sheet.appendRow([name, skuStr, qtyReceived, 0, new Date()]);
}

// ============================================================
//  PO TEMPLATES
// ============================================================

function getTemplatesSheet_() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(TEMPLATES_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(TEMPLATES_SHEET_NAME);
    sheet.appendRow(['Template Name', 'Vendor', 'Attention', 'Warehouse', 'Delivery Instructions', 'Items JSON']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function saveTemplate(data) {
  try {
    var sheet   = getTemplatesSheet_();
    var lastRow = sheet.getLastRow();
    var name    = String(data.templateName || '').trim();
    if (!name) return 'Error: Template name is required.';

    // Dupe check
    if (lastRow > 1) {
      var existing = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      var isDupe   = existing.some(function(row) {
        return String(row[0]).toLowerCase().trim() === name.toLowerCase();
      });
      if (isDupe) return 'Error: A template with that name already exists.';
    }

    var itemsJson = JSON.stringify(data.items || []);
    sheet.appendRow([
      name,
      String(data.vendor || ''),
      String(data.attention || ''),
      String(data.warehouse || ''),
      String(data.deliveryInstructions || ''),
      itemsJson
    ]);
    return 'Success';
  } catch (err) {
    return 'Error: ' + err.message;
  }
}

function getTemplates() {
  try {
    var sheet   = getTemplatesSheet_();
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return [];
    var data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
    return data
      .filter(function(row) { return String(row[0]).trim() !== ''; })
      .map(function(row) {
        var items = [];
        try { items = JSON.parse(row[5]); } catch (e) {}
        return {
          templateName:         String(row[0] || ''),
          vendor:               String(row[1] || ''),
          attention:            String(row[2] || ''),
          warehouse:            String(row[3] || ''),
          deliveryInstructions: String(row[4] || ''),
          items:                items
        };
      });
  } catch (err) {
    return [];
  }
}

function deleteTemplate(templateName) {
  try {
    var sheet   = getTemplatesSheet_();
    var lastRow = sheet.getLastRow();
    var name    = String(templateName || '').trim().toLowerCase();
    if (!name || lastRow <= 1) return 'Error: Template not found.';

    var data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = data.length - 1; i >= 0; i--) {
      if (String(data[i][0]).toLowerCase().trim() === name) {
        sheet.deleteRow(i + 2);
        return 'Success';
      }
    }
    return 'Error: Template not found.';
  } catch (err) {
    return 'Error: ' + err.message;
  }
}

// ------------------------------------------------------------
// SUMMARY – Aggregate stats for the dashboard cards
//           getAllPOs() returns one entry per unique PO.
// ------------------------------------------------------------
function getSummary() {
  try {
    const pos = getAllPOs();
    var lowStockCount = 0;
    try { lowStockCount = getLowStockItems().length; } catch (e) {}
    return {
      total:         pos.length,
      draft:         pos.filter(p => p.status === 'Draft').length,
      issued:        pos.filter(p => p.status === 'Issued').length,
      partial:       pos.filter(p => p.status === 'Partial').length,
      closed:        pos.filter(p => p.status === 'Closed').length,
      totalValue:    pos.reduce((sum, p) => sum + (parseFloat(p.poTotal) || 0), 0),
      lowStockCount: lowStockCount
    };
  } catch (err) {
    return { total: 0, draft: 0, issued: 0, closed: 0, totalValue: 0, lowStockCount: 0 };
  }
}
