/**
 * ================================================================================
 * SETUP DATABASE & SAFE MIGRATION - RUMAH AIR BOGOR RESTO
 * ================================================================================
 * File      : setup.gs
 * Deskripsi : Inisialisasi otomatis dan migrasi non-destruktif untuk Google Sheets.
 * Standar   : ZettBOT v4.1 (Nama tab bebas underscore, idempoten, safe migration).
 * ================================================================================
 */

/**
 * Fungsi inisialisasi utama database Google Sheets.
 * Aman dijalankan berulang kali (Safe Migration): tidak akan menghapus atau menimpa
 * data yang sudah ada di spreadsheet.
 */
function initDatabase() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error('Spreadsheet aktif tidak ditemukan. Pastikan script ini terhubung ke Google Sheets.');
  }

  // 1. Inisialisasi Tab ProfilResto
  setupProfilResto(ss);

  // 2. Inisialisasi Tab DaftarMenu
  setupDaftarMenu(ss);

  // 3. Inisialisasi Tab PromoAcara
  setupPromoAcara(ss);

  // 4. Inisialisasi Tab DataReservasi
  setupDataReservasi(ss);

  // 5. Inisialisasi Tab BannerMenu (Database Gambar Banner Seksi Menu)
  setupBannerMenu(ss);

  // 6. Inisialisasi Tab GaleriResto (Database Foto, Video & Keterangan Galeri)
  setupGaleriResto(ss);

  SpreadsheetApp.flush();
  Logger.log('Inisialisasi & Migrasi Database Rumah Air Bogor Resto Sukses!');
  return 'Database Rumah Air Bogor Resto berhasil disiapkan dan siap digunakan!';
}

/**
 * 1. Setup Tab ProfilResto (Pengaturan Profil & Informasi Restoran)
 */
function setupProfilResto(ss) {
  var sheetName = 'ProfilResto';
  var sheet = ss.getSheetByName(sheetName);
  var headers = ['Kunci', 'Nilai'];

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    formatHeaderRow(sheet, headers.length);
  } else {
    syncHeaders(sheet, headers);
  }

  // Data profil default jika sheet baru kosong
  if (sheet.getLastRow() <= 1) {
    var defaultProfile = [
      ['NamaResto', 'Rumah Air Bogor Resto'],
      ['Tagline', 'Indonesian Culinary Heritage'],
      ['NoWhatsApp', '081111861853'],
      ['NoWhatsAppIntl', '6281111861853'],
      ['Alamat', 'Jl. Raya Boulevard CBD Bogor Nirwana Residence, Kota Bogor, Jawa Barat 16132'],
      ['JamOperasional', 'Senin-Jumat 11:00 - 19:00 WIB, Sabtu-Minggu/Tgl Merah 09:30 - 21:00 WIB'],
      ['JamBuka', '11:00'],
      ['JamTutup', '19:00'],
      ['LinkMaps', 'https://maps.google.com/?q=Rumah+Air+Bogor+Nirwana+Residence'],
      ['LinkInstagram', 'https://instagram.com/rumahairbogor'],
      ['LinkTikTok', 'https://tiktok.com/@rumahairbogor'],
      ['LogoURL', 'https://drive.google.com/file/d/165WesmJN1ixY3SPzVeWvjmKVGAkCMTA-/view?usp=drive_link'],
      ['HeroImageURL', 'https://drive.google.com/file/d/1WTMWbiQ9B2S7F27ThIlJSdjwKVi6kEq5/view?usp=drive_link'],
      ['LinkMenuPDF', 'https://drive.google.com/file/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/view?usp=sharing'],
      ['LinkPdfEventCelebration', 'https://drive.google.com/file/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/view?usp=sharing'],
      ['LinkPdfWedding', 'https://drive.google.com/file/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/view?usp=sharing'],
      ['LinkPdfMeeting', 'https://drive.google.com/file/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/view?usp=sharing'],
      ['FotoGaleriResto', 'https://drive.google.com/file/d/1WTMWbiQ9B2S7F27ThIlJSdjwKVi6kEq5/view?usp=drive_link'],
      ['VideoGaleriResto', '']
    ];
    sheet.getRange(2, 1, defaultProfile.length, 2).setValues(defaultProfile);
    sheet.autoResizeColumns(1, 2);
  } else {
    // Safe Migration: Cek apakah kunci-kunci baru sudah ada di spreadsheet
    var existingKeys = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getDisplayValues();
    var keyMap = {};
    for (var k = 0; k < existingKeys.length; k++) {
      if (existingKeys[k][0]) {
        keyMap[existingKeys[k][0].trim()] = true;
      }
    }

    var newRequiredKeys = [
      ['LinkMenuPDF', 'https://drive.google.com/file/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/view?usp=sharing'],
      ['LinkPdfEventCelebration', 'https://drive.google.com/file/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/view?usp=sharing'],
      ['LinkPdfWedding', 'https://drive.google.com/file/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/view?usp=sharing'],
      ['LinkPdfMeeting', 'https://drive.google.com/file/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/view?usp=sharing'],
      ['FotoGaleriResto', 'https://drive.google.com/file/d/1WTMWbiQ9B2S7F27ThIlJSdjwKVi6kEq5/view?usp=drive_link'],
      ['VideoGaleriResto', '']
    ];

    for (var n = 0; n < newRequiredKeys.length; n++) {
      var itemKey = newRequiredKeys[n][0];
      var itemVal = newRequiredKeys[n][1];
      if (!keyMap[itemKey]) {
        sheet.appendRow([itemKey, itemVal]);
      }
    }
  }
}

/**
 * 2. Setup Tab DaftarMenu (Katalog Menu Makanan & Minuman)
 */
function setupDaftarMenu(ss) {
  var sheetName = 'DaftarMenu';
  var sheet = ss.getSheetByName(sheetName);
  var headers = ['IDMenu', 'Kategori', 'NamaMenu', 'Deskripsi', 'Harga', 'FotoURL', 'Status', 'Badge'];

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    formatHeaderRow(sheet, headers.length);
  } else {
    syncHeaders(sheet, headers);
  }

  // Isi data menu autentik jika sheet masih kosong
  if (sheet.getLastRow() <= 1) {
    var starterMenu = [
      ['MNU-001', 'Ikan & Seafood', 'Gurame Bakar Cobek Khas RAB', 'Gurame segar bakar bumbu cobek khas Sunda dengan sambal jahe pedas gurih meresap sempurna', 85000, 'https://images.unsplash.com/photo-1580476262798-bddd9f4b7369?w=600&q=80', 'Tersedia', 'Best Seller'],
      ['MNU-002', 'Ikan & Seafood', 'Gurame Terbang Renyah', 'Gurame goreng renyah mekar dengan cocolan sambal terasi dadak dan lalapan segar kebun', 82000, 'https://images.unsplash.com/photo-1534422298391-e4f8c172dddb?w=600&q=80', 'Tersedia', 'Favorit'],
      ['MNU-003', 'Makanan Khas', 'Nasi Timbel Komplit Ayam Bakar', 'Nasi bungkus daun pisang, ayam bakar rempah, tahu, tempe bacem, sayur asem, sambal & lalapan', 48000, 'https://images.unsplash.com/photo-1626777552726-4a6b54c97e46?w=600&q=80', 'Tersedia', 'Best Seller'],
      ['MNU-004', 'Makanan Khas', 'Sop Buntut Spesial Kuah Bening', 'Buntut sapi empuk pilihan dengan kuah kaldu rempah gurih hangat, wortel & taburan emping', 78000, 'https://images.unsplash.com/photo-1547592166-23ac45744acd?w=600&q=80', 'Tersedia', 'Rekomendasi'],
      ['MNU-005', 'Makanan Khas', 'Ayam Goreng Lengkuas Saung Danau', 'Ayam kampung goreng empuk bertabur serundeng lengkuas harum gurih khas Parahyangan', 38000, 'https://images.unsplash.com/photo-1626082927389-6cd097cdc6ec?w=600&q=80', 'Tersedia', ''],
      ['MNU-006', 'Sayur & Sambal', 'Karedok Sayur Segar Saus Kacang', 'Sayuran mentah segar renyah dipadu siraman bumbu kacang kencur khas Sunda yang gurih wangi', 22000, 'https://images.unsplash.com/photo-1540420773420-3366772f4999?w=600&q=80', 'Tersedia', 'Favorit'],
      ['MNU-007', 'Sayur & Sambal', 'Sayur Asem Gurih Segar', 'Sayur asem kuah bening asam manis segar dengan melinjo, jagung manis, labu siam & kacang panjang', 18000, 'https://images.unsplash.com/photo-1546069901-ba9599a7e63c?w=600&q=80', 'Tersedia', ''],
      ['MNU-008', 'Sayur & Sambal', 'Sambal Dadak Cobek & Lalapan', 'Ulekan cabai rawit merah, terasi bakar dan tomat segar disajikan langsung di cobek tanah liat', 12000, 'https://images.unsplash.com/photo-1596797038530-2c107229654b?w=600&q=80', 'Tersedia', ''],
      ['MNU-009', 'Minuman', 'Es Kelapa Muda Jeruk Batok', 'Air kelapa muda murni dalam batok utuh dengan perasan jeruk nipis segar & gula aren alami', 28000, 'https://images.unsplash.com/photo-1513558161293-cdaf765ed2fd?w=600&q=80', 'Tersedia', 'Best Seller'],
      ['MNU-010', 'Minuman', 'Es Cincau Hijau Santan Pandan', 'Cincau hijau alami daun camcao dengan santan gurih segar dan sirup gula kelapa wangi pandan', 20000, 'https://images.unsplash.com/photo-1551024709-8f23befc6f87?w=600&q=80', 'Tersedia', 'Favorit'],
      ['MNU-011', 'Minuman', 'Wedang Jahe Rempah Hangat', 'Minuman seduhan jahe merah bakar, serai, cengkeh dan gula batu alami penghangat suasana', 18000, 'https://images.unsplash.com/photo-1544787219-7f47ccb76574?w=600&q=80', 'Tersedia', ''],
      ['MNU-012', 'Paket Rombongan', 'Paket Botram Saung Danau (4-5 Pax)', '1 Gurame Bakar Cobek, 4 Nasi Timbel, 1 Porsi Ayam Lengkuas, 1 Sayur Asem, Karedok, Sambal Lalap & 4 Es Kelapa', 295000, 'https://images.unsplash.com/photo-1555939594-58d7cb561ad1?w=600&q=80', 'Tersedia', 'Paling Hemat'],
      ['MNU-013', 'Paket Rombongan', 'Paket Wisata Keluarga Besar (8-10 Pax)', '2 Gurame Bakar/Goreng, 8 Nasi Liwet Kastrol, 2 Porsi Ayam Bakar, Cumi Saus Padang, Tumis Kangkung, Sayur Asem & Minuman', 580000, 'https://images.unsplash.com/photo-1504674900247-0877df9cc836?w=600&q=80', 'Tersedia', 'Rekomendasi']
    ];
    sheet.getRange(2, 1, starterMenu.length, headers.length).setValues(starterMenu);
    sheet.autoResizeColumns(1, headers.length);
  }
}

/**
 * 3. Setup Tab PromoAcara (Banner Promosi & Pengumuman)
 */
function setupPromoAcara(ss) {
  var sheetName = 'PromoAcara';
  var sheet = ss.getSheetByName(sheetName);
  var headers = ['IDPromo', 'JudulPromo', 'Keterangan', 'BannerURL', 'Status'];

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    formatHeaderRow(sheet, headers.length);
  } else {
    syncHeaders(sheet, headers);
  }

  // Isi data promo jika sheet masih kosong
  if (sheet.getLastRow() <= 1) {
    var starterPromo = [
      ['PRM-001', 'Paket Botram Keluarga di Saung Danau', 'Nikmati makan lesehan di saung apung atas danau dengan semilir angin sejuk dan panorama asri. Diskon khusus 10% untuk reservasi hari kerja (Senin - Kamis).', 'https://drive.google.com/file/d/1WTMWbiQ9B2S7F27ThIlJSdjwKVi6kEq5/view?usp=drive_link', 'Aktif'],
      ['PRM-002', 'Paket Gathering Perusahaan & Arisan VIP', 'Fasilitas saung luas kapasitas hingga 100 pax lengkap dengan sound system, proyektor, dan pemandangan danau yang asri.', 'https://drive.google.com/file/d/1WTMWbiQ9B2S7F27ThIlJSdjwKVi6kEq5/view?usp=drive_link', 'Aktif']
    ];
    sheet.getRange(2, 1, starterPromo.length, headers.length).setValues(starterPromo);
    sheet.autoResizeColumns(1, headers.length);
  }
}

/**
 * 4. Setup Tab DataReservasi (Pencatatan Masuk Booking Tamu)
 */
function setupDataReservasi(ss) {
  var sheetName = 'DataReservasi';
  var sheet = ss.getSheetByName(sheetName);
  var headers = ['NoReservasi', 'WaktuInput', 'NamaTamu', 'NomorWA', 'TanggalAcara', 'JamAcara', 'JumlahPax', 'PilihanArea', 'CatatanKhusus', 'StatusReservasi'];

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    formatHeaderRow(sheet, headers.length);
  } else {
    syncHeaders(sheet, headers);
  }

  // Sheet reservasi tidak diisi dummy data berlebihan agar bersih untuk operasional nyata
  sheet.autoResizeColumns(1, headers.length);
}

/**
 * Helper: Format Header Row agar tampil profesional & tegas
 */
function formatHeaderRow(sheet, colCount) {
  var headerRange = sheet.getRange(1, 1, 1, colCount);
  headerRange.setBackground('#134E4A'); // Deep Forest Teal
  headerRange.setFontColor('#FFFFFF');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
  sheet.setFrozenRows(1);
}

/**
 * Helper: Sinkronisasi Header Baru (Safe Migration)
 * Menambahkan kolom baru ke paling kanan tanpa menghapus data yang sudah ada.
 */
function syncHeaders(sheet, expectedHeaders) {
  var lastCol = sheet.getLastColumn();
  if (lastCol === 0) {
    sheet.getRange(1, 1, 1, expectedHeaders.length).setValues([expectedHeaders]);
    formatHeaderRow(sheet, expectedHeaders.length);
    return;
  }

  var existingHeaders = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  var missingHeaders = [];

  for (var i = 0; i < expectedHeaders.length; i++) {
    if (existingHeaders.indexOf(expectedHeaders[i]) === -1) {
      missingHeaders.push(expectedHeaders[i]);
    }
  }

  if (missingHeaders.length > 0) {
    var startCol = lastCol + 1;
    sheet.getRange(1, startCol, 1, missingHeaders.length).setValues([missingHeaders]);
    formatHeaderRow(sheet, startCol + missingHeaders.length - 1);
  }
}

/**
 * 5. Setup Tab BannerMenu (Database Gambar Banner Seksi Menu Regullar)
 */
function setupBannerMenu(ss) {
  var sheetName = 'BannerMenu';
  var sheet = ss.getSheetByName(sheetName);
  var headers = ['IDBanner', 'Judul', 'FotoURL', 'Status'];

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    formatHeaderRow(sheet, headers.length);
  } else {
    syncHeaders(sheet, headers);
  }

  // Isi data default banner menu jika sheet masih kosong
  if (sheet.getLastRow() <= 1) {
    var defaultBanners = [
      ['BM-001', 'Buku Menu & Kuliner Khas Nusantara', 'https://drive.google.com/file/d/1WTMWbiQ9B2S7F27ThIlJSdjwKVi6kEq5/view?usp=drive_link', 'Aktif']
    ];
    sheet.getRange(2, 1, defaultBanners.length, headers.length).setValues(defaultBanners);
    sheet.autoResizeColumns(1, headers.length);
  }
}

/**
 * 6. Setup Tab GaleriResto (Database Foto, Video & Keterangan Galeri Interaktif)
 */
function setupGaleriResto(ss) {
  var sheetName = 'GaleriResto';
  var sheet = ss.getSheetByName(sheetName);
  var headers = ['IDGaleri', 'Judul', 'MediaURL', 'Tipe', 'Status'];

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    formatHeaderRow(sheet, headers.length);
  } else {
    syncHeaders(sheet, headers);
  }

  // Isi data default galeri jika sheet masih kosong (Hanya 1 Foto Resmi Rumah Air, NOL Unsplash)
  if (sheet.getLastRow() <= 1) {
    var defaultGallery = [
      ['GL-001', 'Saung Utama & Danau Nirwana', 'https://drive.google.com/file/d/1WTMWbiQ9B2S7F27ThIlJSdjwKVi6kEq5/view?usp=drive_link', 'Foto', 'Aktif']
    ];
    sheet.getRange(2, 1, defaultGallery.length, headers.length).setValues(defaultGallery);
    sheet.autoResizeColumns(1, headers.length);
  }
}
