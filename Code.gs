/**
 * ==========================================================
 * CODE.GS - VERSI FINAL (AUTO MERGE + IMPORT CSV FIXED)
 * ==========================================================
 */
const CONFIG = {
  APP_NAME: "WMS Inventory System",
  SHEET_NAMES: {
    MASTER_VIEW: "MasterProduk", 
    INBOUND: "StokMasuk",
    BATCH_PRODUCT: "BatchProduct", 
    SUPPORT: "DataPendukung"
  },
  HEADERS: {
    BATCH: 13, 
    INBOUND: 17 
  },
  // --- KONFIGURASI MERGE & CSV TEMPLATE ---
  INBOUND_MATCH_KEYS: [
    "kodeBarang", "typeLokasi", "lokasi", "noBatch",
    "expiredDate", "typeSticker", "sizeCover", "size", "keterangan"
  ],
  CSV_TEMPLATE: {
    HEADERS: [
      "KODE", "QTY", "TYPE LOKASI", "LOKASI", "BATCH", "EXPIRED",
      "TYPE STICKER", "SIZE COVER", "SIZE PRODUK", "KETERANGAN"
    ],
    EXAMPLE: [
      "SNB-0001", "10", "Excel", "EX01", "BATCH-123", "2026-12-31",
      "STICKER VIAL", "Kecil", "100ML", "Contoh Data"
    ]
  },
  SUPPORT_MAP: { LOKASI_KODE: 1, LOKASI_TIPE: 2, SIZE: 4, STICKER: 6, COVER: 8 }
};

// --- 1. ROUTING HALAMAN ---
function doGet(e) {
  var page = e.parameter.page || 'dashboard';
  var template;
  
  if (page == 'dashboard') template = HtmlService.createTemplateFromFile('Dashboard');
  else if (page == 'barang-masuk') template = HtmlService.createTemplateFromFile('BarangMasuk');
  else if (page == 'riwayat-masuk') template = HtmlService.createTemplateFromFile('RiwayatMasuk');
  else if (page == 'master-produk') template = HtmlService.createTemplateFromFile('MasterProduk');
  else if (page == 'laporan-barang') template = HtmlService.createTemplateFromFile('BatchProduct');
  else if (page == 'data-pendukung') template = HtmlService.createTemplateFromFile('DataPendukung');
  else template = HtmlService.createTemplateFromFile('Dashboard');
  
  template.page = page; 
  template.scriptUrl = ScriptApp.getService().getUrl();
  template.appName = CONFIG.APP_NAME;
  return template.evaluate().setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename, params) {
  try {
    var template = HtmlService.createTemplateFromFile(filename);
    if (params) { for (var key in params) { template[key] = params[key]; } }
    return template.evaluate().getContent();
  } catch (e) { return "Error: File " + filename + " not found."; }
}

function getAppConfig() { return CONFIG; }

// --- 2. FUNGSI TRANSAKSI INBOUND ---
function processInboundTransaction(header, items) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sHist = _ensureSheet(CONFIG.SHEET_NAMES.INBOUND);
  const sBatch = _ensureSheet(CONFIG.SHEET_NAMES.BATCH_PRODUCT);
  
  let trxId = _genId(sHist);

  try {
    items.forEach(item => {
      sHist.appendRow([
        trxId, header.tanggalMasuk, header.poProduk, header.poSticker, header.shipmentDate,
        item.kodeBarang, item.namaProduk, item.brand, 
        (item.noBatch||"-"), (item.expiredDate||"-"), (item.typeSticker||"-"), (item.sizeCover||"-"), (item.size||"-"),
        item.qtyMasuk, item.typeLokasi, item.lokasi, (item.keterangan||"-")
      ]);
      _updateStock(sBatch, item);
    });
    SpreadsheetApp.flush();
    return { success: true, message: "Validasi Sukses!", trxId: trxId };
  } catch(e) {
    return { success: false, message: "Error: " + e.message };
  }
}

function _updateStock(sheet, item) {
  let kode = String(item.kodeBarang).toUpperCase().trim();
  let loc = String(item.lokasi).toUpperCase().trim();
  let batch = String(item.noBatch || "-").toUpperCase().trim();
  let stick = String(item.typeSticker || "-").toUpperCase().trim();
  let cover = String(item.sizeCover || "-").toUpperCase().trim();
  let size = String(item.size || "-").toUpperCase().trim();
  
  let data = sheet.getDataRange().getValues();
  let found = false;

  for(let i=1; i<data.length; i++) {
    if (String(data[i][1]).toUpperCase() == kode &&
        String(data[i][6]).toUpperCase() == loc &&
        String(data[i][7]).toUpperCase() == batch &&
        String(data[i][9]).toUpperCase() == stick &&
        String(data[i][10]).toUpperCase() == cover &&
        String(data[i][11]).toUpperCase() == size) {
      
      let newQty = Number(data[i][4]) + Number(item.qtyMasuk);
      sheet.getRange(i+1, 5).setValue(newQty);
      sheet.getRange(i+1, 13).setValue(new Date());
      found = true;
      break;
    }
  }

  if(!found) {
    let id = "BATCH-" + Math.floor(Math.random()*1000000);
    sheet.appendRow([
      id, kode, item.namaProduk, item.brand, item.qtyMasuk, 
      item.typeLokasi, loc, batch, (item.expiredDate||"-"), 
      stick, cover, size, new Date()
    ]);
  }
}

// --- 3. BACA HISTORY (MONITORING) - GROUPING ---
function getInboundHistory(start, end) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAMES.INBOUND);
  if(!sheet || sheet.getLastRow() < 2) return [];

  const data = sheet.getRange(2,1,sheet.getLastRow()-1, 17).getDisplayValues(); 
  let groups = {};
  let isFilterOn = (start && start !== "" && end && end !== "");
  
  for (let i = 0; i < data.length; i++) {
    let r = data[i];
    let id = r[0]; 
    if (!id) continue;

    let rawDate = String(r[1]).trim();
    let dateForFilter = rawDate;
    if (rawDate.match(/^\d{1,2}\/\d{1,2}\/\d{4}$/)) {
       let parts = rawDate.split('/');
       dateForFilter = `${parts[2]}-${parts[1].padStart(2,'0')}-${parts[0].padStart(2,'0')}`;
    }

    let pass = true;
    if (isFilterOn) {
        if (dateForFilter.match(/^\d{4}-\d{2}-\d{2}$/)) {
            if (dateForFilter < start || dateForFilter > end) pass = false;
        }
    }

    if (pass) {
      if (!groups[id]) {
        groups[id] = {
          id: id, tgl: rawDate, poProduk: r[2], poSticker: r[3], shipment: r[4],
          user: "Admin", totalQty: 0, items: []
        };
      }
      groups[id].items.push({
        kode: r[5], nama: r[6], brand: r[7], batch: r[8], exp: r[9],
        typeSticker: r[10], sizeCover: r[11], size: r[12],
        qty: Number(r[13]), typeLoc: r[14], lokasi: r[15], notes: r[16] || "-" 
      });
      groups[id].totalQty += Number(r[13]);
    }
  }
  return Object.values(groups).reverse();
}

// --- 4. DATA PENDUKUNG & HELPER ---
function getBatchStockList(filterType) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAMES.BATCH_PRODUCT);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, CONFIG.HEADERS.BATCH).getDisplayValues();
  let stockList = [];
  for (let i = 0; i < data.length; i++) {
    let r = data[i]; if (!r[1]) continue; 
    let item = {
      idBatch: r[0], kode: r[1], nama: r[2], brand: r[3], qty: Number(r[4]) || 0,
      typeLoc: r[5], lokasi: r[6], batch: r[7], exp: r[8], sticker: r[9], cover: r[10], size: r[11], lastUpdate: r[12]
    };
    if (filterType === 'positive' && item.qty <= 0) continue;
    if (filterType === 'zero' && item.qty > 0) continue;
    stockList.push(item);
  }
  return stockList;
}

function _ensureSheet(name) {
  let s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if(!s) s = SpreadsheetApp.getActiveSpreadsheet().insertSheet(name); return s;
}
function _genId(sheet) {
  let pre = "IN-" + Utilities.formatDate(new Date(), "Asia/Jakarta", "yyMMdd") + "-";
  return pre + (sheet.getLastRow() < 2 ? "001" : String(sheet.getLastRow()).padStart(3,'0'));
}
function getMasterData() {
  const s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("MasterProduk");
  return (s && s.getLastRow() > 1) ? s.getRange(2, 1, s.getLastRow()-1, 4).getValues() : [];
}
function getSupportingData() {
  const s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DataPendukung");
  if(!s || s.getLastRow() < 2) return {lokasi:[], size:[], typeSticker:[], sizeCover:[]};
  const d = s.getRange(2, 1, s.getLastRow()-1, 9).getValues();
  let res = {lokasi:[], size:[], typeSticker:[], sizeCover:[]};
  d.forEach(r => {
    if(r[0]) res.lokasi.push(r[1] ? r[0]+" - "+r[1] : r[0]);
    if(r[3]) res.size.push(r[3]);
    if(r[5]) res.typeSticker.push(r[5]);
    if(r[7]) res.sizeCover.push(r[7]);
  });
  return res;
}
function addSupportItem(cat, v1, v2) {
  const s = _ensureSheet("DataPendukung");
  let col = (cat=='lokasi')?1 : (cat=='size')?4 : (cat=='typeSticker')?6 : 8;
  let r = s.getLastRow()+1; s.getRange(r, col).setValue(v1);
  if(v2) s.getRange(r, col+1).setValue(v2); return {success:true};
}
function updateSupportItem(cat, oldV, newV, newT) { return {success:true}; }
function deleteSupportItem(cat, val) {
  const s = _ensureSheet("DataPendukung");
  let col = (cat=='lokasi')?1 : (cat=='size')?4 : (cat=='typeSticker')?6 : 8;
  let d = s.getDataRange().getValues(); let search = (cat=='lokasi') ? val.split(" - ")[0] : val;
  for(let i=0; i<d.length; i++) { if(String(d[i][col-1]) == String(search)) { s.getRange(i+1, col).clearContent(); if(cat=='lokasi') s.getRange(i+1, col+1).clearContent(); break; } }
}
function addMasterData(form) {
  const s = _ensureSheet("MasterProduk"); s.appendRow([getNextId(), form.nama, form.brand, form.notes]); return {success:true};
}
function updateMasterData(form) {
  const s = _ensureSheet("MasterProduk"); const d = s.getDataRange().getValues();
  for(let i=1; i<d.length; i++) { if(String(d[i][0]) == form.kode) { s.getRange(i+1, 2).setValue(form.nama); s.getRange(i+1, 3).setValue(form.brand); s.getRange(i+1, 4).setValue(form.notes); return {success:true}; } }
}
function deleteMasterData(kode) {
  const s = _ensureSheet("MasterProduk"); const d = s.getDataRange().getValues();
  for(let i=1; i<d.length; i++) { if(String(d[i][0]) == kode) { s.deleteRow(i+1); return {success:true}; } }
}
function getNextId() {
  const s = _ensureSheet("MasterProduk"); if(s.getLastRow()<2) return "SNB-0001";
  const last = s.getRange(s.getLastRow(), 1).getValue(); const num = parseInt(last.split("-")[1]) + 1;
  return "SNB-" + String(num).padStart(4,'0');
}
function deleteInboundTransaction(id) {
  const s = _ensureSheet("StokMasuk"); const d = s.getDataRange().getValues(); let deleted = false;
  for(let i=d.length-1; i>=1; i--) { if(String(d[i][0]) == id) { s.deleteRow(i+1); deleted = true; } }
  if(deleted) return {success:true, message:"Transaksi Dihapus"}; return {success:false, message:"ID Tidak Ditemukan"};
}
