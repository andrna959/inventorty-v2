// ==========================================================
// CONFIGURATION - FINAL VERSION (INTEGRATED WITH BATCH_MAP)
// ==========================================================
const CONFIG = {
  APP_NAME: "WMS Inventory System",
  
  // 1. INFO PERUSAHAAN (Untuk Header Surat Jalan & Dokumen)
  COMPANY_INFO: {
    NAME: "PT SOCIAL BELLA INDONESIA",
    ADDRESS: "St. Moritz Office Tower Lt. 15, Jl. Puri Indah Raya, Jakarta Barat 11610",
    WAREHOUSE: "Gudang Cikupa, Tangerang"
  },

  // 2. KONFIGURASI TANDA TANGAN DINAMIS
  SIGNATURES: [
    { 
      title: "PACKER", 
      defaultName: "Warehouse Staff", 
      showDate: true,
      isCustomer: false
    },
    { 
      title: "SECURITY", 
      defaultName: "LP", 
      showDate: true,
      isCustomer: false
    },
    { 
      title: "RECEIVED BY", 
      defaultName: "", 
      isCustomer: true, // Ambil nama customer dari data transaksi secara otomatis
      showDate: true 
    }
  ],

  // 3. NAMA SHEET (Database)
  SHEET_NAMES: {
    MASTER_VIEW: "MasterProduk", 
    INBOUND: "StokMasuk",
    OUTBOUND: "StokKeluar",
    BATCH_PRODUCT: "BatchProduct", 
    SUPPORT: "DataPendukung",
    PUTAWAY_LOG: "RiwayatPutaway" // Log khusus untuk mencatat histori pemindahan rak
  },

  // 4. PEMETAAN KOLOM BATCH PRODUCT (Koordinat Database Utama)
  // Memetakan indeks array (0-12) ke nama kolom di sheet BatchProduct
  BATCH_MAP: {
    ID: 0,           // Kolom A
    KODE: 1,         // Kolom B
    NAMA: 2,         // Kolom C
    BRAND: 3,        // Kolom D
    QTY: 4,          // Kolom E
    TYPE_LOC: 5,     // Kolom F
    LOKASI: 6,       // Kolom G
    BATCH: 7,        // Kolom H
    EXP: 8,          // Kolom I
    STICKER: 9,      // Kolom J
    COVER: 10,       // Kolom K
    SIZE: 11,        // Kolom L
    LAST_UPDATE: 12  // Kolom M
  },

  // 5. SETTING MODUL PUTAWAY
  PUTAWAY_SETTINGS: {
    DEFAULT_TRANSIT_LOC: "TRANSIT", // Lokasi default barang yang baru masuk (Inbound)
    MODES: {
      PARTIAL: "PARTIAL MOVE", // Pemindahan sebagian Qty
      ALL: "MOVE ALL"          // Pemindahan satu batch utuh
    }
  },

  // 6. KONFIGURASI TEKNIS & HEADERS
  HEADERS: {
    BATCH: 13, // Total 13 kolom di sheet BatchProduct
    INBOUND: 17 
  },

  // Kunci pencocokan untuk logika penggabungan stok (Merge Logic)
  INBOUND_MATCH_KEYS: [
    "kodeBarang", "typeLokasi", "lokasi", "noBatch",
    "expiredDate", "typeSticker", "sizeCover", "size", "keterangan"
  ],

  // Template untuk fitur Import CSV
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

  // Pemetaan kolom pada sheet DataPendukung
  SUPPORT_MAP: { 
    LOKASI_KODE: 1, 
    LOKASI_TIPE: 2, 
    SIZE: 4, 
    STICKER: 6, 
    COVER: 8 
  }
};

// ==========================================================
// ROUTING (DO GET) - UPDATED FOR PUTAWAY MODULE
// ==========================================================
function doGet(e) {
  var page = e.parameter.page || 'dashboard';
  var template;
  
  // Routing Halaman
  if (page == 'dashboard') {
    template = HtmlService.createTemplateFromFile('Dashboard');
  } 
  else if (page == 'BarangMasuk') {
    template = HtmlService.createTemplateFromFile('BarangMasuk');
  } 
  else if (page == 'RiwayatMasuk') {
    template = HtmlService.createTemplateFromFile('RiwayatMasuk');
  } 
  else if (page == 'BarangKeluar') {
    template = HtmlService.createTemplateFromFile('BarangKeluar');
  } 
  else if (page == 'RiwayatKeluar') {
    template = HtmlService.createTemplateFromFile('RiwayatKeluar');
  }
  // --- START: MODUL PUTAWAY ROUTING ---
  else if (page == 'PutawayPartial') {
    template = HtmlService.createTemplateFromFile('PutawayPartial');
  }
  else if (page == 'PutawayAll') {
    template = HtmlService.createTemplateFromFile('PutawayAll');
  }
  else if (page == 'RiwayatPutaway') {
    template = HtmlService.createTemplateFromFile('RiwayatPutaway');
  }
  // --- END: MODUL PUTAWAY ROUTING ---
  else if (page == 'CetakSuratJalan') {
    template = HtmlService.createTemplateFromFile('CetakSuratJalan');
    template.trxId = e.parameter.id; 
  }
  else {
    template = HtmlService.createTemplateFromFile('Dashboard');
  }
  
  // Parameter Global untuk Template
  template.page = page; 
  template.scriptUrl = ScriptApp.getService().getUrl();
  template.appName = CONFIG.APP_NAME;
  
  return template.evaluate()
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setTitle(CONFIG.APP_NAME + " - " + page);
}

// ==========================================================
// FUNGSI KHUSUS PRINTING & OUTBOUND
// ==========================================================

// 1. AMBIL DATA RIWAYAT KELUAR (Untuk Tabel Monitoring)
function getOutboundHistory(start, end) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAMES.OUTBOUND);
  if(!sheet || sheet.getLastRow() < 2) return [];

  const data = sheet.getRange(2, 1, sheet.getLastRow()-1, 12).getDisplayValues(); 
  let groups = {};
  
  let isFilterOn = (start && start !== "" && end && end !== "");

  for (let i = 0; i < data.length; i++) {
    let r = data[i];
    let id = r[0]; 
    if (!id) continue;

    let rawDate = String(r[1]); 
    let dateObj = new Date(rawDate);
    let dateStr = Utilities.formatDate(dateObj, "Asia/Jakarta", "yyyy-MM-dd");

    if (isFilterOn) {
       if (dateStr < start || dateStr > end) continue;
    }

    if (!groups[id]) {
      groups[id] = {
        id: id,
        tgl: Utilities.formatDate(dateObj, "Asia/Jakarta", "dd/MM/yyyy"), 
        docNo: r[2],    
        customer: r[3], 
        notes: r[11],
        totalQty: 0,
        itemCount: 0
      };
    }
    
    groups[id].totalQty += Number(r[7]); 
    groups[id].itemCount += 1; 
  }

  return Object.values(groups).reverse();
}

// 2. AMBIL DATA CETAK (Untuk Surat Jalan)
function getOutboundPrintData(trxId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAMES.OUTBOUND);
  if (!sheet || sheet.getLastRow() < 2) return null;

  const data = sheet.getDataRange().getValues();
  // Filter baris berdasarkan ID Transaksi
  const rows = data.filter(r => String(r[0]) === String(trxId));

  if (rows.length === 0) return null;

  const firstRow = rows[0];
  
  // Siapkan Data Header
  const header = {
    company: CONFIG.COMPANY_INFO, // Data Perusahaan dari Config
    signatures: CONFIG.SIGNATURES, // Data Tanda Tangan dari Config
    
    trxId: firstRow[0],
    tanggal: Utilities.formatDate(new Date(firstRow[1]), "Asia/Jakarta", "dd MMMM yyyy"),
    docNo: firstRow[2],
    customer: firstRow[3],
    notes: firstRow[11] || "-"
  };

  // Siapkan Data Barang (Items)
  const items = rows.map(r => {
    return {
      sku: r[4],
      nama: r[5],
      brand: r[6],
      qty: r[7],
      batch: r[8],
      exp: r[9] ? Utilities.formatDate(new Date(r[9]), "Asia/Jakarta", "yyyy-MM-dd") : "-",
      lokasi: r[10]
    };
  });

  return { header: header, items: items };
}

function include(filename, params) {
  try {
    var template = HtmlService.createTemplateFromFile(filename);
    if (params) { for (var key in params) { template[key] = params[key]; } }
    return template.evaluate().getContent();
  } catch (e) { return "Error: File " + filename + " not found."; }
}

function getAppConfig() { return CONFIG; }

// --- 2. FUNGSI TRANSAKSI INBOUND (TIDAK DISENTUH) ---
function processInboundTransaction(header, items) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sHist = _ensureSheet(CONFIG.SHEET_NAMES.INBOUND);
  const sBatch = _ensureSheet(CONFIG.SHEET_NAMES.BATCH_PRODUCT);
  
  let trxId = _genId(sHist, "IN-"); // [REVISI KECIL] Parameterisasi Prefix agar fungsi _genId bisa dipakai Outbound

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

// --- 3. BACA HISTORY INBOUND (TIDAK DISENTUH) ---
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

// --- 4. DATA PENDUKUNG & HELPER (DIREVISI BAGIAN UPDATE/DELETE) ---
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

// [REVISI] Menambahkan parameter prefix agar bisa dipakai untuk IN- dan OUT-
function _genId(sheet, prefix) {
  let pre = (prefix || "IN-") + Utilities.formatDate(new Date(), "Asia/Jakarta", "yyMMdd") + "-";
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

// [REVISI] Mengisi logika update yang sebelumnya kosong
function updateSupportItem(cat, oldV, newV, newT) {
  const s = _ensureSheet("DataPendukung");
  let col = (cat=='lokasi')?1 : (cat=='size')?4 : (cat=='typeSticker')?6 : 8;
  const data = s.getDataRange().getValues();
  // Parsing oldV jika kategori lokasi (format "Kode - Tipe")
  let searchVal = (cat == 'lokasi' && oldV.includes(" - ")) ? oldV.split(" - ")[0] : oldV;
  
  for(let i=0; i<data.length; i++) {
    if(String(data[i][col-1]) == String(searchVal)) {
      s.getRange(i+1, col).setValue(newV);
      if(cat == 'lokasi' && newT) s.getRange(i+1, col+1).setValue(newT);
      return {success:true};
    }
  }
  return {success:false, message:"Data tidak ditemukan."};
}

// [REVISI] Menggunakan deleteRow agar tidak ada baris kosong
function deleteSupportItem(cat, val) {
  const s = _ensureSheet("DataPendukung");
  let col = (cat=='lokasi')?1 : (cat=='size')?4 : (cat=='typeSticker')?6 : 8;
  let d = s.getDataRange().getValues(); 
  let search = (cat=='lokasi') ? val.split(" - ")[0] : val;
  for(let i=0; i<d.length; i++) { 
    if(String(d[i][col-1]) == String(search)) { 
      s.deleteRow(i+1); // Pakai deleteRow, bukan clearContent
      return {success:true};
    } 
  }
  return {success:false};
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

// ==========================================================
// --- 5. FUNGSI BARU: OUTBOUND (BARANG KELUAR) ---
// ==========================================================

// [BARU] Mengambil stok tersedia untuk picking, diurutkan Expired (FEFO)
function getAvailableStockForPicking(sku) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAMES.BATCH_PRODUCT);
  if (!sheet || sheet.getLastRow() < 2) return [];

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, CONFIG.HEADERS.BATCH).getDisplayValues();
  let availableList = [];
  sku = String(sku).toUpperCase().trim();

  for (let i = 0; i < data.length; i++) {
    let r = data[i];
    // Index: 1=Kode, 4=Qty, 5=TypeLoc, 6=Loc, 7=Batch, 8=Exp, 9=Sticker, 10=Cover, 11=Size
    if (String(r[1]).toUpperCase().trim() === sku) {
      let qty = Number(r[4]);
      if (qty > 0) { 
        availableList.push({
          batch: r[7], lokasi: r[6], typeLoc: r[5], exp: r[8],
          qty: qty, sticker: r[9], cover: r[10], size: r[11]
        });
      }
    }
  }
  // Sort FEFO (Expired duluan di atas)
  return availableList.sort((a, b) => {
     if (!a.exp || a.exp === '-') return 1;
     if (!b.exp || b.exp === '-') return -1;
     return new Date(a.exp) - new Date(b.exp);
  });
}

// [BARU] Memproses transaksi keluar: Simpan Log & Kurangi Stok
function processOutboundTransaction(header, cart) {
  const sOut = _ensureSheet(CONFIG.SHEET_NAMES.OUTBOUND);
  const sBatch = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAMES.BATCH_PRODUCT);
  
  let trxId = _genId(sOut, "OUT-"); // Generate ID Outbound

  try {
    let batchData = sBatch.getDataRange().getValues();
    
    cart.forEach(item => {
      item.pickDetails.forEach(pick => {
         // 1. Update Stok (Deduksi) di BatchProduct
         for(let i=1; i<batchData.length; i++) {
            let row = batchData[i];
            if(String(row[1]) == item.sku && String(row[6]) == pick.lokasi && String(row[7]) == pick.batch) {
               let newQty = Number(row[4]) - Number(pick.qty);
               if(newQty < 0) newQty = 0;
               sBatch.getRange(i+1, 5).setValue(newQty);
               sBatch.getRange(i+1, 13).setValue(new Date()); // Last Update
               break; 
            }
         }
         // 2. Simpan Log ke StokKeluar
         sOut.appendRow([
            trxId, header.tanggal, header.docNo, header.customer,
            item.sku, item.nama, item.brand,
            pick.qty, pick.batch, pick.exp, pick.lokasi,
            header.notes
         ]);
      });
    });
    SpreadsheetApp.flush();
    return { success: true, message: "Outbound Berhasil Disimpan!", trxId: trxId };
  } catch (e) {
    return { success: false, message: "Error: " + e.message };
  }
}


// ==========================================================
// MODUL PUTAWAY - FINAL (BATCH & LOCATION MOVE)
// ==========================================================

// 1. CARI SATU KOTAK (Untuk Partial Move)
function getBatchDetailsForPutaway(searchId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAMES.BATCH_PRODUCT);
  if (!sheet || sheet.getLastRow() < 2) return null;
  
  const search = String(searchId).toUpperCase().trim();
  const BM = CONFIG.BATCH_MAP;
  
  // Cari ID Batch di Kolom A
  const finder = sheet.getRange("A:A").createTextFinder(search).matchCase(false).matchEntireCell(true).findNext();
  
  if (finder) {
    const rowIndex = finder.getRow();
    const rowData = sheet.getRange(rowIndex, 1, 1, CONFIG.HEADERS.BATCH).getValues()[0];
    
    return {
      idBatch: rowData[BM.ID], kode: rowData[BM.KODE], nama: rowData[BM.NAMA],
      qty: Number(rowData[BM.QTY]), lokasi: rowData[BM.LOKASI], batch: rowData[BM.BATCH]
    };
  }
  return null;
}

// 2. CARI SEMUA ISI RAK (Untuk Move All Lokasi) - BARU!
function getItemsAtLocation(locationName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAMES.BATCH_PRODUCT);
  if (!sheet || sheet.getLastRow() < 2) return null;
  
  const search = String(locationName).toUpperCase().trim();
  const BM = CONFIG.BATCH_MAP;
  const data = sheet.getDataRange().getValues();
  let itemsFound = [];
  
  // Mencari semua barang yang punya alamat rak yang sama
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][BM.LOKASI]).toUpperCase().trim() === search) {
      itemsFound.push({
        rowIndex: i + 1, idBatch: data[i][BM.ID],
        nama: data[i][BM.NAMA], qty: Number(data[i][BM.QTY])
      });
    }
  }
  return itemsFound.length > 0 ? itemsFound : null;
}

// 3. PINDAH SEBAGIAN (Partial Move)
function processPutawayPartial(moveData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sBatch = ss.getSheetByName(CONFIG.SHEET_NAMES.BATCH_PRODUCT);
  const sLog = ss.getSheetByName(CONFIG.SHEET_NAMES.PUTAWAY_LOG);
  const BM = CONFIG.BATCH_MAP;

  try {
    const finder = sBatch.getRange("A:A").createTextFinder(moveData.idBatch).findNext();
    if (!finder) throw new Error("Barang tidak ditemukan!");
    
    const rowIndex = finder.getRow();
    const rowData = sBatch.getRange(rowIndex, 1, 1, CONFIG.HEADERS.BATCH).getValues()[0];
    const currentQty = Number(rowData[BM.QTY]);
    const moveQty = Number(moveData.qtyMove);
    
    if (moveQty > currentQty) throw new Error("Stok tidak cukup!");
    
    sBatch.getRange(rowIndex, BM.QTY + 1).setValue(currentQty - moveQty);
    sBatch.getRange(rowIndex, BM.LAST_UPDATE + 1).setValue(new Date());

    _updateStock(sBatch, { 
      kodeBarang: rowData[BM.KODE], namaProduk: rowData[BM.NAMA], brand: rowData[BM.BRAND],
      qtyMasuk: moveQty, typeLokasi: moveData.targetTypeLoc, lokasi: moveData.targetLoc,
      noBatch: rowData[BM.BATCH], expiredDate: rowData[BM.EXP], typeSticker: rowData[BM.STICKER],
      sizeCover: rowData[BM.COVER], size: rowData[BM.SIZE]
    });

    sLog.appendRow([_genId(sLog, "PW-"), new Date(), "Admin", rowData[BM.KODE], rowData[BM.NAMA], rowData[BM.BATCH], rowData[BM.EXP], rowData[BM.STICKER], rowData[BM.COVER], rowData[BM.SIZE], rowData[BM.LOKASI], moveData.targetLoc, moveQty, "PARTIAL MOVE"]);
    return { success: true, message: "Berhasil pindah eceran!" };
  } catch (e) { return { success: false, message: e.message }; }
}

// 4. PINDAH SEMUA ISI RAK (Location Move All) - BARU!
function processLocationMoveAll(moveData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sBatch = ss.getSheetByName(CONFIG.SHEET_NAMES.BATCH_PRODUCT);
  const sLog = ss.getSheetByName(CONFIG.SHEET_NAMES.PUTAWAY_LOG);
  const BM = CONFIG.BATCH_MAP;

  try {
    moveData.items.forEach(item => {
      const rowData = sBatch.getRange(item.rowIndex, 1, 1, CONFIG.HEADERS.BATCH).getValues()[0];
      const oldLoc = rowData[BM.LOKASI];

      sBatch.getRange(item.rowIndex, BM.TYPE_LOC + 1).setValue(moveData.targetTypeLoc);
      sBatch.getRange(item.rowIndex, BM.LOKASI + 1).setValue(moveData.targetLoc);
      sBatch.getRange(item.rowIndex, BM.LAST_UPDATE + 1).setValue(new Date());

      sLog.appendRow([_genId(sLog, "PW-LOC-"), new Date(), "Admin", rowData[BM.KODE], rowData[BM.NAMA], rowData[BM.BATCH], rowData[BM.EXP], rowData[BM.STICKER], rowData[BM.COVER], rowData[BM.SIZE], oldLoc, moveData.targetLoc, rowData[BM.QTY], "LOCATION MOVE ALL"]);
    });
    return { success: true, message: "Satu rak berhasil dikosongkan!" };
  } catch (e) { return { success: false, message: e.message }; }
}

// 5. LIHAT RIWAYAT
function getPutawayHistory() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAMES.PUTAWAY_LOG);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 14).getDisplayValues();
  return data.map(r => ({ id: r[0], tgl: r[1], kode: r[3], nama: r[4], batch: r[5], asal: r[10], tujuan: r[11], qty: r[12], mode: r[13] })).reverse();
}
