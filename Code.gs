const SHEET_NAMES = {
  CATALOG: 'Catalog',
  BATCHES: 'Batches',
  SALES: 'Sales',
  LEDGER: 'Ledger',
  CUSTOMERS: 'Customers',
  SETTINGS: 'Settings'
};

function doGet() {
  ensureSheetsInitialized();
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('MediCare POS & IAS 2 Ledger')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getCacheVersion() {
  const props = PropertiesService.getScriptProperties();
  let ver = props.getProperty('CACHE_VERSION');
  if (!ver) {
    ver = String(Date.now());
    props.setProperty('CACHE_VERSION', ver);
  }
  return ver;
}

function bumpCacheVersion() {
  const ver = String(Date.now());
  PropertiesService.getScriptProperties().setProperty('CACHE_VERSION', ver);
  return ver;
}

function ensureSheetsInitialized() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const schemas = {
    [SHEET_NAMES.CATALOG]: ['ID', 'Company', 'Name', 'Generic', 'SellPrice', 'Img', 'PcsPerStrip', 'StripsPerBox', 'MinStock'],
    [SHEET_NAMES.BATCHES]: ['BatchNo', 'Date', 'Company', 'MedID', 'MedName', 'Generic', 'InitialQty', 'RemainingQty', 'GrossUnitCost', 'CommRate', 'NetUnitCost', 'Expiry', 'AlertDays'],
    [SHEET_NAMES.SALES]: ['InvoiceNo', 'Date', 'Time', 'CustomerPhone', 'CustomerName', 'Subtotal', 'Discount', 'GrandTotal', 'GrossCost', 'NetCost', 'RetailMarkup', 'RealizedComm', 'GrossProfit', 'Status', 'ItemsJSON', 'CashReceived', 'ChangeAmount'],
    [SHEET_NAMES.LEDGER]: ['EntryID', 'Date', 'Type', 'Ref', 'Party', 'Details', 'Debit', 'Credit', 'ProfitImpact'],
    [SHEET_NAMES.CUSTOMERS]: ['Phone', 'Name', 'LastVisit', 'TotalVisits'],
    [SHEET_NAMES.SETTINGS]: ['Key', 'Value']
  };

  Object.keys(schemas).forEach(function(name) {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      sheet.appendRow(schemas[name]);
      sheet.getRange(1, 1, 1, schemas[name].length).setFontWeight('bold');
    } else {
      const lastCol = sheet.getLastColumn();
      const existingHeaders = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
      if (existingHeaders.length < schemas[name].length) {
        for (let c = existingHeaders.length; c < schemas[name].length; c++) {
          sheet.getRange(1, c + 1).setValue(schemas[name][c]).setFontWeight('bold');
        }
      }
    }
  });

  const setSheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
  if (setSheet.getLastRow() <= 1) {
    setSheet.appendRow(['shopName', 'MediCare Model Pharmacy']);
    setSheet.appendRow(['shopPhone', '+880 1711-000000']);
    setSheet.appendRow(['shopAddress', 'Dhaka, Bangladesh']);
    setSheet.appendRow(['shopFooter', 'Thank you for your visit.']);
    setSheet.appendRow(['alertEmail', Session.getActiveUser().getEmail() || '']);
  }

  const custSheet = ss.getSheetByName(SHEET_NAMES.CUSTOMERS);
  custSheet.getRange('A:A').setNumberFormat('@');
}

function normalizePhone(rawPhone) {
  let p = String(rawPhone || '').trim().replace(/^'+/, '');
  if (!p) return '';
  if (/^1[3-9]\d{8}$/.test(p)) {
    p = '0' + p;
  }
  return p;
}

function formatDateYMD(val) {
  if (!val) return '';
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(val).trim();
}

function getAllAppData() {
  ensureSheetsInitialized();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const catRows = ss.getSheetByName(SHEET_NAMES.CATALOG).getDataRange().getValues();
  const catalog = [];
  for (let i = 1; i < catRows.length; i++) {
    const r = catRows[i];
    if (!r[0]) continue;
    catalog.push({
      id: Number(r[0]),
      company: String(r[1] || ''),
      name: String(r[2] || ''),
      generic: String(r[3] || ''),
      sellPrice: Number(r[4]) || 0,
      img: String(r[5] || ''),
      pcsPerStrip: Math.max(1, Number(r[6]) || 1),
      stripsPerBox: Math.max(1, Number(r[7]) || 1),
      minStock: Math.max(0, Number(r[8]) || 0)
    });
  }

  const batchRows = ss.getSheetByName(SHEET_NAMES.BATCHES).getDataRange().getValues();
  const batches = [];
  for (let i = 1; i < batchRows.length; i++) {
    const r = batchRows[i];
    if (!r[0]) continue;
    batches.push({
      batchNo: String(r[0]),
      date: formatDateYMD(r[1]),
      company: String(r[2] || ''),
      medId: Number(r[3]),
      medName: String(r[4] || ''),
      generic: String(r[5] || ''),
      initialQty: Number(r[6]) || 0,
      remainingQty: Number(r[7]) || 0,
      grossUnitCost: Number(r[8]) || 0,
      commRate: Number(r[9]) || 0,
      netUnitCost: Number(r[10]) || 0,
      expiry: formatDateYMD(r[11]),
      alertDays: Number(r[12]) || 15
    });
  }

  const salesRows = ss.getSheetByName(SHEET_NAMES.SALES).getDataRange().getValues();
  const sales = [];
  for (let i = 1; i < salesRows.length; i++) {
    const r = salesRows[i];
    if (!r[0]) continue;
    let parsedItems = [];
    try { parsedItems = JSON.parse(r[14] || '[]'); } catch (e) {}
    const grandTotalVal = Number(r[7]) || 0;
    const cashRecVal = r[15] !== '' && r[15] !== undefined ? Number(r[15]) : grandTotalVal;
    const changeVal = r[16] !== '' && r[16] !== undefined ? Number(r[16]) : Math.max(0, cashRecVal - grandTotalVal);

    sales.push({
      invoiceNo: String(r[0]),
      date: formatDateYMD(r[1]),
      time: String(r[2] || ''),
      customerPhone: normalizePhone(r[3]),
      customerName: String(r[4] || ''),
      subtotal: Number(r[5]) || 0,
      discount: Number(r[6]) || 0,
      grandTotal: grandTotalVal,
      grossCost: Number(r[8]) || 0,
      netCost: Number(r[9]) || 0,
      retailMarkup: Number(r[10]) || 0,
      realizedComm: Number(r[11]) || 0,
      grossProfit: Number(r[12]) || 0,
      status: String(r[13] || 'Completed'),
      items: parsedItems,
      cashReceived: cashRecVal,
      changeAmount: changeVal
    });
  }
  sales.reverse();

  const ledRows = ss.getSheetByName(SHEET_NAMES.LEDGER).getDataRange().getValues();
  const ledger = [];
  for (let i = 1; i < ledRows.length; i++) {
    const r = ledRows[i];
    if (!r[0]) continue;
    ledger.push({
      entryId: String(r[0]),
      date: formatDateYMD(r[1]),
      type: String(r[2] || ''),
      ref: String(r[3] || ''),
      party: String(r[4] || ''),
      details: String(r[5] || ''),
      debit: Number(r[6]) || 0,
      credit: Number(r[7]) || 0,
      profitImpact: Number(r[8]) || 0
    });
  }
  ledger.reverse();

  const custRows = ss.getSheetByName(SHEET_NAMES.CUSTOMERS).getDataRange().getValues();
  const customers = [];
  for (let i = 1; i < custRows.length; i++) {
    const r = custRows[i];
    const phone = normalizePhone(r[0]);
    if (!phone) continue;
    customers.push({
      phone: phone,
      name: String(r[1] || ''),
      lastVisit: formatDateYMD(r[2]),
      totalVisits: Number(r[3]) || 1
    });
  }

  const setRows = ss.getSheetByName(SHEET_NAMES.SETTINGS).getDataRange().getValues();
  const settings = {};
  for (let i = 1; i < setRows.length; i++) {
    if (setRows[i][0]) settings[String(setRows[i][0])] = String(setRows[i][1] || '');
  }

  return {
    cacheVersion: getCacheVersion(),
    catalog: catalog,
    batches: batches,
    sales: sales,
    ledger: ledger,
    customers: customers,
    settings: settings
  };
}

function checkAndSendLowStockAlerts(medIdsToCheck) {
  if (!medIdsToCheck || medIdsToCheck.length === 0) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const setRows = ss.getSheetByName(SHEET_NAMES.SETTINGS).getDataRange().getValues();
  const settings = {};
  for (let i = 1; i < setRows.length; i++) {
    if (setRows[i][0]) settings[String(setRows[i][0])] = String(setRows[i][1] || '');
  }

  const recipientEmail = (settings.alertEmail || '').trim() || Session.getActiveUser().getEmail();
  if (!recipientEmail) return;

  const catRows = ss.getSheetByName(SHEET_NAMES.CATALOG).getDataRange().getValues();
  const batchRows = ss.getSheetByName(SHEET_NAMES.BATCHES).getDataRange().getValues();

  const stockMap = {};
  for (let b = 1; b < batchRows.length; b++) {
    const mId = Number(batchRows[b][3]);
    const rem = Number(batchRows[b][7]) || 0;
    if (rem > 0) {
      stockMap[mId] = (stockMap[mId] || 0) + rem;
    }
  }

  const props = PropertiesService.getScriptProperties();
  const alertedMeds = [];

  medIdsToCheck.forEach(function(targetId) {
    const idNum = Number(targetId);
    for (let c = 1; c < catRows.length; c++) {
      if (Number(catRows[c][0]) === idNum) {
        const company = String(catRows[c][1] || '');
        const medName = String(catRows[c][2] || '');
        const generic = String(catRows[c][3] || '');
        const minStock = Math.max(0, Number(catRows[c][8]) || 0);
        const currentStock = stockMap[idNum] || 0;
        const propKey = 'LOW_STOCK_NOTIFIED_' + idNum;

        if (minStock > 0 && currentStock <= minStock) {
          if (props.getProperty(propKey) !== '1') {
            alertedMeds.push({
              name: medName,
              generic: generic,
              company: company,
              currentStock: currentStock,
              minStock: minStock
            });
            props.setProperty(propKey, '1');
          }
        } else if (currentStock > minStock) {
          props.deleteProperty(propKey);
        }
        break;
      }
    }
  });

  if (alertedMeds.length > 0) {
    const shopName = settings.shopName || 'MediCare POS';
    const subject = '[' + shopName + '] Low Inventory Alert: ' + alertedMeds.map(function(m) { return m.name; }).join(', ');

    const lines = [
      'Low Inventory Alert from ' + shopName,
      'Date: ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd hh:mm a'),
      '',
      'The following medicine(s) have reached or dropped below the minimum stock threshold:',
      ''
    ];

    alertedMeds.forEach(function(m, idx) {
      lines.push((idx + 1) + '. Product Name: ' + m.name + ' (' + m.generic + ')');
      lines.push('   Company: ' + m.company);
      lines.push('   Current Stock: ' + m.currentStock + ' pcs/bottles');
      lines.push('   Minimum Stock Limit: ' + m.minStock + ' pcs/bottles');
      lines.push('');
    });

    lines.push('Please arrange a wholesale purchase order soon.');

    try {
      MailApp.sendEmail(recipientEmail, subject, lines.join('\n'));
    } catch (e) {
      Logger.log('Low stock email error: ' + e.message);
    }
  }
}

function saveCatalogMedicine(med) {
  ensureSheetsInitialized();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.CATALOG);
  const rows = sheet.getDataRange().getValues();

  const pcsPerStrip = Math.max(1, Number(med.pcsPerStrip) || 1);
  const stripsPerBox = Math.max(1, Number(med.stripsPerBox) || 1);
  const minStock = Math.max(0, Number(med.minStock) || 0);

  if (med.id) {
    for (let i = 1; i < rows.length; i++) {
      if (Number(rows[i][0]) === Number(med.id)) {
        const finalImg = med.img !== undefined ? med.img : rows[i][5];
        sheet.getRange(i + 1, 2, 1, 8).setValues([[
          med.company,
          med.name,
          med.generic,
          Number(med.sellPrice) || 0,
          finalImg,
          pcsPerStrip,
          stripsPerBox,
          minStock
        ]]);
        PropertiesService.getScriptProperties().deleteProperty('LOW_STOCK_NOTIFIED_' + Number(med.id));
        bumpCacheVersion();
        return getAllAppData();
      }
    }
  }

  let maxId = 0;
  for (let i = 1; i < rows.length; i++) {
    const idVal = Number(rows[i][0]) || 0;
    if (idVal > maxId) maxId = idVal;
  }

  sheet.appendRow([
    maxId + 1,
    med.company,
    med.name,
    med.generic,
    Number(med.sellPrice) || 0,
    med.img || '',
    pcsPerStrip,
    stripsPerBox,
    minStock
  ]);
  bumpCacheVersion();
  return getAllAppData();
}

function updateSellingPrice(medId, newPrice) {
  ensureSheetsInitialized();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.CATALOG);
  const rows = sheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (Number(rows[i][0]) === Number(medId)) {
      sheet.getRange(i + 1, 5).setValue(Number(newPrice) || 0);
      break;
    }
  }
  bumpCacheVersion();
  return getAllAppData();
}

function savePurchaseBatch(payload) {
  ensureSheetsInitialized();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const batchSheet = ss.getSheetByName(SHEET_NAMES.BATCHES);
  const ledgerSheet = ss.getSheetByName(SHEET_NAMES.LEDGER);
  const rows = batchSheet.getDataRange().getValues();

  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const baseQty = Number(payload.baseQty) || Number(payload.qty) || 0;
  const grossUnit = Number(payload.grossUnitCost) || 0;
  const commRate = Number(payload.commRate) || 0;
  const netUnit = Number((grossUnit * (1 - commRate / 100)).toFixed(4));
  const netPayable = Number((baseQty * netUnit).toFixed(2));
  const unitNote = payload.purchaseUnitLabel ? ' (' + payload.purchaseUnitLabel + ')' : '';

  if (payload.isEdit && payload.batchNo) {
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(payload.batchNo)) {
        const oldInit = Number(rows[i][6]) || 0;
        const oldRem = Number(rows[i][7]) || 0;
        const soldSoFar = Math.max(0, oldInit - oldRem);
        const newRem = Math.max(0, baseQty - soldSoFar);

        batchSheet.getRange(i + 1, 3, 1, 11).setValues([[
          payload.company,
          Number(payload.medId),
          payload.medName,
          payload.generic,
          baseQty,
          newRem,
          grossUnit,
          commRate,
          netUnit,
          payload.expiry,
          Number(payload.alertDays) || 15
        ]]);

        const ledRows = ledgerSheet.getDataRange().getValues();
        for (let j = 1; j < ledRows.length; j++) {
          if (String(ledRows[j][3]) === String(payload.batchNo) && String(ledRows[j][2]) === 'Wholesale Purchase') {
            ledgerSheet.getRange(j + 1, 5, 1, 5).setValues([[
              payload.company,
              payload.medName + ' (' + baseQty + ' pcs' + unitNote + ' @ ' + netUnit.toFixed(2) + ' TK net)',
              0,
              netPayable,
              0
            ]]);
            break;
          }
        }
        checkAndSendLowStockAlerts([Number(payload.medId)]);
        bumpCacheVersion();
        return getAllAppData();
      }
    }
  }

  batchSheet.appendRow([
    payload.batchNo,
    today,
    payload.company,
    Number(payload.medId),
    payload.medName,
    payload.generic,
    baseQty,
    baseQty,
    grossUnit,
    commRate,
    netUnit,
    payload.expiry,
    Number(payload.alertDays) || 15
  ]);

  ledgerSheet.appendRow([
    'LED-' + Date.now(),
    today,
    'Wholesale Purchase',
    payload.batchNo,
    payload.company,
    payload.medName + ' (' + baseQty + ' pcs' + unitNote + ' @ ' + netUnit.toFixed(2) + ' TK net)',
    0,
    netPayable,
    0
  ]);

  checkAndSendLowStockAlerts([Number(payload.medId)]);
  bumpCacheVersion();
  return getAllAppData();
}

function deletePurchaseBatch(batchNo) {
  ensureSheetsInitialized();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const batchSheet = ss.getSheetByName(SHEET_NAMES.BATCHES);
  const ledgerSheet = ss.getSheetByName(SHEET_NAMES.LEDGER);

  let affectedMedId = null;
  const rows = batchSheet.getDataRange().getValues();
  for (let i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][0]) === String(batchNo)) {
      affectedMedId = Number(rows[i][3]);
      batchSheet.deleteRow(i + 1);
      break;
    }
  }

  const ledRows = ledgerSheet.getDataRange().getValues();
  for (let j = ledRows.length - 1; j >= 1; j--) {
    if (String(ledRows[j][3]) === String(batchNo)) {
      ledgerSheet.deleteRow(j + 1);
    }
  }

  if (affectedMedId) checkAndSendLowStockAlerts([affectedMedId]);
  bumpCacheVersion();
  return getAllAppData();
}

function processBatchDispositionQty(batchNo, actionType, qtyToProcess) {
  ensureSheetsInitialized();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const batchSheet = ss.getSheetByName(SHEET_NAMES.BATCHES);
  const ledgerSheet = ss.getSheetByName(SHEET_NAMES.LEDGER);
  const rows = batchSheet.getDataRange().getValues();
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

  let affectedMedId = null;
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(batchNo)) {
      const company = String(rows[i][2]);
      affectedMedId = Number(rows[i][3]);
      const medName = String(rows[i][4]);
      const remQty = Number(rows[i][7]) || 0;
      const netUnit = Number(rows[i][10]) || 0;
      const qty = Math.min(remQty, Math.max(1, Number(qtyToProcess) || remQty));
      if (qty <= 0) break;

      const carryingVal = Number((qty * netUnit).toFixed(2));
      batchSheet.getRange(i + 1, 8).setValue(remQty - qty);

      if (actionType === 'return') {
        ledgerSheet.appendRow([
          'LED-' + Date.now(),
          today,
          'Supplier Return',
          batchNo,
          company,
          'Returned ' + qty + ' pcs of ' + medName + ' (Full Credit)',
          carryingVal,
          0,
          0
        ]);
      } else {
        ledgerSheet.appendRow([
          'LED-' + Date.now(),
          today,
          'Stock Destroyed',
          batchNo,
          company,
          'Destroyed ' + qty + ' pcs of ' + medName + ' (IAS 2 Write-off)',
          0,
          0,
          -carryingVal
        ]);
      }
      break;
    }
  }

  if (affectedMedId) checkAndSendLowStockAlerts([affectedMedId]);
  bumpCacheVersion();
  return getAllAppData();
}

function restoreBatchesFromItems(batchSheet, items) {
  const rows = batchSheet.getDataRange().getValues();
  items.forEach(function(item) {
    const allocs = item.batchAllocations || [];
    allocs.forEach(function(al) {
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][0]) === String(al.batchNo)) {
          const curRem = Number(rows[i][7]) || 0;
          rows[i][7] = curRem + Number(al.qty);
          batchSheet.getRange(i + 1, 8).setValue(rows[i][7]);
          break;
        }
      }
    });
  });
}

function completeOrUpdateSale(payload) {
  ensureSheetsInitialized();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const batchSheet = ss.getSheetByName(SHEET_NAMES.BATCHES);
  const salesSheet = ss.getSheetByName(SHEET_NAMES.SALES);
  const ledgerSheet = ss.getSheetByName(SHEET_NAMES.LEDGER);
  const custSheet = ss.getSheetByName(SHEET_NAMES.CUSTOMERS);

  const now = new Date();
  const today = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const timeStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'hh:mm a');

  const salesRows = salesSheet.getDataRange().getValues();
  let targetSalesRow = -1;
  let invoiceNo = payload.editInvoiceNo || '';

  if (invoiceNo) {
    for (let i = 1; i < salesRows.length; i++) {
      if (String(salesRows[i][0]) === String(invoiceNo)) {
        targetSalesRow = i + 1;
        let oldItems = [];
        try { oldItems = JSON.parse(salesRows[i][14] || '[]'); } catch (e) {}
        restoreBatchesFromItems(batchSheet, oldItems);
        break;
      }
    }
  } else {
    invoiceNo = 'INV-' + Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyMMdd') + '-' + String(salesRows.length).padStart(3, '0');
  }

  const batchRows = batchSheet.getDataRange().getValues();
  const activeBatches = [];
  for (let i = 1; i < batchRows.length; i++) {
    activeBatches.push({
      rowIndex: i + 1,
      batchNo: String(batchRows[i][0]),
      date: formatDateYMD(batchRows[i][1]),
      medId: Number(batchRows[i][3]),
      remainingQty: Number(batchRows[i][7]) || 0,
      grossUnitCost: Number(batchRows[i][8]) || 0,
      netUnitCost: Number(batchRows[i][10]) || 0
    });
  }
  activeBatches.sort(function(a, b) {
    return a.date.localeCompare(b.date) || a.batchNo.localeCompare(b.batchNo);
  });

  let subtotal = 0;
  let totalGrossCost = 0;
  let totalNetCost = 0;
  const processedItems = [];
  const soldMedIds = [];

  payload.items.forEach(function(item) {
    const medIdNum = Number(item.medId);
    if (soldMedIds.indexOf(medIdNum) === -1) soldMedIds.push(medIdNum);

    const baseQty = Number(item.baseQty) || Number(item.qty) || 0;
    const displayQty = Number(item.displayQty) || baseQty;
    const unitLabel = String(item.unitLabel || 'Pcs');
    const unitSellPrice = Number(item.unitSellPrice) || Number(item.sellPrice) || 0;
    const baseSellPrice = baseQty > 0 ? (unitSellPrice * displayQty) / baseQty : Number(item.sellPrice) || 0;

    let needed = baseQty;
    let itemGrossCost = 0;
    let itemNetCost = 0;
    const allocations = [];

    for (let k = 0; k < activeBatches.length; k++) {
      const b = activeBatches[k];
      if (needed <= 0) break;
      if (b.medId !== medIdNum || b.remainingQty <= 0) continue;

      const take = Math.min(b.remainingQty, needed);
      b.remainingQty -= take;
      needed -= take;

      itemGrossCost += take * b.grossUnitCost;
      itemNetCost += take * b.netUnitCost;
      allocations.push({
        batchNo: b.batchNo,
        qty: take,
        grossUnitCost: b.grossUnitCost,
        netUnitCost: b.netUnitCost
      });
    }

    if (needed > 0) {
      throw new Error('Insufficient stock for ' + item.name);
    }

    const lineTotal = Number((displayQty * unitSellPrice).toFixed(2));
    subtotal += lineTotal;
    totalGrossCost += itemGrossCost;
    totalNetCost += itemNetCost;

    processedItems.push({
      medId: medIdNum,
      name: item.name,
      generic: item.generic,
      unitType: item.unitType || 'pcs',
      unitLabel: unitLabel,
      multiplier: Number(item.multiplier) || 1,
      displayQty: displayQty,
      unitSellPrice: unitSellPrice,
      sellPrice: baseSellPrice,
      qty: baseQty,
      baseQty: baseQty,
      lineTotal: lineTotal,
      batchAllocations: allocations
    });
  });

  activeBatches.forEach(function(b) {
    batchSheet.getRange(b.rowIndex, 8).setValue(b.remainingQty);
  });

  const discount = Math.max(0, Number(payload.discount) || 0);
  const grandTotal = Number(Math.max(0, subtotal - discount).toFixed(2));
  const rawCashReceived = Number(payload.cashReceived);
  const cashReceived = !isNaN(rawCashReceived) && rawCashReceived > 0 ? Number(rawCashReceived.toFixed(2)) : grandTotal;
  const changeAmount = Number(Math.max(0, cashReceived - grandTotal).toFixed(2));

  const retailMarkup = Number((grandTotal - totalGrossCost).toFixed(2));
  const realizedComm = Number((totalGrossCost - totalNetCost).toFixed(2));
  const grossProfit = Number((grandTotal - totalNetCost).toFixed(2));

  const custPhone = normalizePhone(payload.customerPhone);
  const custName = String(payload.customerName || '').trim();

  const rowValues = [
    invoiceNo,
    today,
    timeStr,
    custPhone ? "'" + custPhone : '',
    custName,
    Number(subtotal.toFixed(2)),
    discount,
    grandTotal,
    Number(totalGrossCost.toFixed(2)),
    Number(totalNetCost.toFixed(2)),
    retailMarkup,
    realizedComm,
    grossProfit,
    'Completed',
    JSON.stringify(processedItems),
    cashReceived,
    changeAmount
  ];

  if (targetSalesRow > 0) {
    salesSheet.getRange(targetSalesRow, 1, 1, rowValues.length).setValues([rowValues]);
  } else {
    salesSheet.appendRow(rowValues);
  }

  const ledRows = ledgerSheet.getDataRange().getValues();
  let updatedLedger = false;
  const partyLabel = [custName, custPhone].filter(Boolean).join(' - ') || 'Walk-in Customer';
  const itemSummary = processedItems.map(function(i) { return i.name + ' (' + i.displayQty + ' ' + i.unitLabel + ')'; }).join(', ');

  for (let j = 1; j < ledRows.length; j++) {
    if (String(ledRows[j][3]) === String(invoiceNo) && String(ledRows[j][2]) === 'Retail Sale') {
      ledgerSheet.getRange(j + 1, 5, 1, 5).setValues([[partyLabel, itemSummary, grandTotal, 0, grossProfit]]);
      updatedLedger = true;
      break;
    }
  }
  if (!updatedLedger) {
    ledgerSheet.appendRow(['LED-' + Date.now(), today, 'Retail Sale', invoiceNo, partyLabel, itemSummary, grandTotal, 0, grossProfit]);
  }

  if (custPhone) {
    const custRows = custSheet.getDataRange().getValues();
    let foundCust = false;
    for (let c = 1; c < custRows.length; c++) {
      if (normalizePhone(custRows[c][0]) === custPhone) {
        const visits = (Number(custRows[c][3]) || 0) + 1;
        custSheet.getRange(c + 1, 1).setNumberFormat('@').setValue(custPhone);
        custSheet.getRange(c + 1, 2, 1, 3).setValues([[custName || custRows[c][1], today, visits]]);
        foundCust = true;
        break;
      }
    }
    if (!foundCust) {
      const nextRow = custSheet.getLastRow() + 1;
      custSheet.getRange(nextRow, 1).setNumberFormat('@').setValue(custPhone);
      custSheet.getRange(nextRow, 2, 1, 3).setValues([[custName, today, 1]]);
    }
  }

  checkAndSendLowStockAlerts(soldMedIds);
  bumpCacheVersion();
  const appData = getAllAppData();
  const invoiceObj = appData.sales.find(function(s) { return s.invoiceNo === invoiceNo; });
  return { appData: appData, invoice: invoiceObj };
}

function processFullSaleReturn(invoiceNo) {
  ensureSheetsInitialized();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const batchSheet = ss.getSheetByName(SHEET_NAMES.BATCHES);
  const salesSheet = ss.getSheetByName(SHEET_NAMES.SALES);
  const ledgerSheet = ss.getSheetByName(SHEET_NAMES.LEDGER);

  const salesRows = salesSheet.getDataRange().getValues();
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

  for (let i = 1; i < salesRows.length; i++) {
    if (String(salesRows[i][0]) === String(invoiceNo)) {
      if (String(salesRows[i][13]) === 'Returned') break;

      let items = [];
      try { items = JSON.parse(salesRows[i][14] || '[]'); } catch (e) {}
      restoreBatchesFromItems(batchSheet, items);

      const medIds = items.map(function(it) { return Number(it.medId); });
      checkAndSendLowStockAlerts(medIds);

      const grandTotal = Number(salesRows[i][7]) || 0;
      const grossProfit = Number(salesRows[i][12]) || 0;
      const party = [salesRows[i][4], normalizePhone(salesRows[i][3])].filter(Boolean).join(' - ') || 'Walk-in Customer';

      salesSheet.getRange(i + 1, 14).setValue('Returned');
      ledgerSheet.appendRow([
        'LED-' + Date.now(),
        today,
        'Sale Return',
        invoiceNo,
        party,
        'Full customer return & stock restored to FIFO batches',
        0,
        grandTotal,
        -grossProfit
      ]);
      break;
    }
  }

  bumpCacheVersion();
  return getAllAppData();
}

function saveShopSettings(settings) {
  ensureSheetsInitialized();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
  const rows = sheet.getDataRange().getValues();

  Object.keys(settings).forEach(function(key) {
    let found = false;
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === key) {
        sheet.getRange(i + 1, 2).setValue(settings[key]);
        found = true;
        break;
      }
    }
    if (!found) {
      sheet.appendRow([key, settings[key]]);
    }
  });

  bumpCacheVersion();
  return getAllAppData();
}
