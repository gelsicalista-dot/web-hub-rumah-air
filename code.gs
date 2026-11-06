/**
 * ================================================================================
 * BACKEND CONTROLLER & API - RUMAH AIR BOGOR RESTO
 * ================================================================================
 * File      : code.gs
 * Deskripsi : Handler doGet Web App, API pengambilan data profil, menu, promo,
 *             serta transaksi reservasi dengan penomoran unik otomatis & LockService.
 * Standar   : ZettBOT v4.1 (Index-Based Mapping, getDisplayValues, SpreadsheetApp.flush)
 * ================================================================================
 */

// KONSTANTA PEMETAAN INDEKS KOLOM (INDEX-BASED MAPPING)
// Kebal terhadap modifikasi teks header oleh pengguna di Google Sheets.

const COL_PROFIL = {
  KUNCI: 0, // Kolom A
  NILAI: 1  // Kolom B
};

const COL_MENU = {
  ID: 0,        // Kolom A: IDMenu
  KATEGORI: 1,  // Kolom B: Kategori
  NAMA: 2,      // Kolom C: NamaMenu
  DESKRIPSI: 3, // Kolom D: Deskripsi
  HARGA: 4,     // Kolom E: Harga
  FOTO: 5,      // Kolom F: FotoURL
  STATUS: 6,    // Kolom G: Status
  BADGE: 7      // Kolom H: Badge
};

const COL_PROMO = {
  ID: 0,        // Kolom A: IDPromo
  JUDUL: 1,     // Kolom B: JudulPromo
  KETERANGAN: 2,// Kolom C: Keterangan
  BANNER: 3,    // Kolom D: BannerURL
  STATUS: 4     // Kolom E: Status
};

const COL_BANNER_MENU = {
  ID: 0,        // Kolom A: IDBanner
  JUDUL: 1,     // Kolom B: Judul
  FOTO: 2,      // Kolom C: FotoURL
  STATUS: 3     // Kolom D: Status
};

const COL_GALERI = {
  ID: 0,        // Kolom A: IDGaleri
  JUDUL: 1,     // Kolom B: Judul
  MEDIA: 2,     // Kolom C: MediaURL
  TIPE: 3,      // Kolom D: Tipe
  STATUS: 4     // Kolom E: Status
};

const COL_RESERVASI = {
  NO_RESERVASI: 0, // Kolom A: NoReservasi
  WAKTU_INPUT: 1,  // Kolom B: WaktuInput
  NAMA: 2,         // Kolom C: NamaTamu
  WA: 3,           // Kolom D: NomorWA
  TANGGAL: 4,      // Kolom E: TanggalAcara
  JAM: 5,          // Kolom F: JamAcara
  PAX: 6,          // Kolom G: JumlahPax
  AREA: 7,         // Kolom H: PilihanArea
  CATATAN: 8,      // Kolom I: CatatanKhusus
  STATUS: 9        // Kolom J: StatusReservasi
};

/**
 * Handler HTTP GET untuk menyajikan antarmuka Web App
 */
function doGet(e) {
  var htmlOutput = HtmlService.createHtmlOutputFromFile('index')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
    .setTitle('Rumah Air Bogor Resto | Official Web Hub & Pamflet Digital')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  return htmlOutput;
}

/**
 * Mengambil seluruh data awal aplikasi secara terpadu (Profile, Menu, Promo, Status Jam Buka)
 * Meminimalisir network round-trip antara klien dan server.
 */
function getAppInitialData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      return { success: false, message: 'Spreadsheet aktif tidak ditemukan.' };
    }

    var profile = fetchProfileData(ss);
    var menuList = fetchMenuList(ss);
    var promoList = fetchPromoList(ss);
    var bannerMenuList = fetchBannerMenuList(ss);
    var galeriList = fetchGaleriList(ss);
    var openStatus = calculateOpenStatus(profile);

    return {
      success: true,
      profile: profile,
      menu: menuList,
      promo: promoList,
      bannerMenu: bannerMenuList,
      galeri: galeriList,
      openStatus: openStatus,
      timestamp: Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy HH:mm:ss')
    };
  } catch (error) {
    Logger.log('Error in getAppInitialData: ' + error.toString());
    return {
      success: false,
      message: 'Gagal memuat data: ' + error.message
    };
  }
}

/**
 * Helper: Ambil data Profil Resto
 */
function fetchProfileData(ss) {
  var sheet = ss.getSheetByName('ProfilResto');
  var profile = {
    NamaResto: 'Rumah Air Bogor Resto',
    Tagline: 'Sensasi Kuliner Saung Danau & Wisata Keluarga Khas Sunda Nusantara',
    NoWhatsApp: '081111861853',
    NoWhatsAppIntl: '6281111861853',
    Alamat: 'Jl. Raya Boulevard CBD Bogor Nirwana Residence, Kota Bogor, Jawa Barat 16132',
    JamOperasional: '10:00 - 21:00 WIB (Buka Setiap Hari)',
    JamBuka: '10:00',
    JamTutup: '21:00',
    LinkMaps: 'https://maps.google.com/?q=Rumah+Air+Bogor+Nirwana+Residence',
    LinkInstagram: 'https://instagram.com/rumahairbogor',
    LinkTikTok: 'https://tiktok.com/@rumahairbogor',
    LogoURL: 'https://drive.google.com/file/d/165WesmJN1ixY3SPzVeWvjmKVGAkCMTA-/view?usp=drive_link',
    HeroImageURL: 'https://drive.google.com/file/d/1WTMWbiQ9B2S7F27ThIlJSdjwKVi6kEq5/view?usp=drive_link',
    LinkMenuPDF: 'https://drive.google.com/file/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/view?usp=sharing',
    LinkPdfEventCelebration: 'https://drive.google.com/file/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/view?usp=sharing',
    LinkPdfWedding: 'https://drive.google.com/file/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/view?usp=sharing',
    LinkPdfMeeting: 'https://drive.google.com/file/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/view?usp=sharing',
    FotoGaleriResto: 'https://drive.google.com/file/d/1WTMWbiQ9B2S7F27ThIlJSdjwKVi6kEq5/view?usp=drive_link',
    VideoGaleriResto: ''
  };

  if (!sheet || sheet.getLastRow() <= 1) return profile;

  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getDisplayValues();
  for (var i = 0; i < data.length; i++) {
    var key = data[i][COL_PROFIL.KUNCI].toString().trim();
    var val = data[i][COL_PROFIL.NILAI].toString().trim();
    if (key) {
      profile[key] = val;
    }
  }

  // Sanitasi nomor WhatsApp internasional (+62)
  if (profile.NoWhatsApp) {
    var cleanWA = profile.NoWhatsApp.replace(/[^0-9]/g, '');
    if (cleanWA.indexOf('0') === 0) {
      cleanWA = '62' + cleanWA.substring(1);
    }
    profile.NoWhatsAppIntl = cleanWA;
  }

  // Auto-Convert Google Drive Image links to Direct CDN URLs
  profile.LogoURL = formatDirectImageUrl(profile.LogoURL);

  // Parse Multi-Image Hero Slider (Mendukung beberapa foto dipisah koma atau baris baru)
  var heroRaw = profile.HeroImageURL || '';
  var heroArr = [];
  var splitItems = heroRaw.split(/[\n,]+/);
  for (var h = 0; h < splitItems.length; h++) {
    var itemUrl = splitItems[h].trim();
    if (itemUrl) {
      heroArr.push(formatDirectImageUrl(itemUrl));
    }
  }

  // Filter Cerdas Anti-Demo: Jika ada foto asli pengguna (Google Drive / custom link),
  // buang secara otomatis seluruh link demo Unsplash agar tidak tercampur di slider!
  var customImages = [];
  for (var c = 0; c < heroArr.length; c++) {
    if (heroArr[c].indexOf('images.unsplash.com') === -1) {
      customImages.push(heroArr[c]);
    }
  }

  if (customImages.length > 0) {
    heroArr = customImages; // 100% foto murni milik pengguna
  } else if (heroArr.length === 0) {
    heroArr.push('https://lh3.googleusercontent.com/d/1WTMWbiQ9B2S7F27ThIlJSdjwKVi6kEq5');
  }

  profile.HeroImages = heroArr;
  profile.HeroImageURL = heroArr[0]; // Backward compatibility

  // Parse Galeri Foto Rumah Air Bogor
  var galeriRaw = profile.FotoGaleriResto || '';
  var galeriArr = [];
  var splitGaleri = galeriRaw.split(/[\n,]+/);
  for (var g = 0; g < splitGaleri.length; g++) {
    var galeriItem = splitGaleri[g].trim();
    if (galeriItem) {
      galeriArr.push(formatDirectImageUrl(galeriItem));
    }
  }

  // Jika galeri dari spreadsheet masih kosong, sertakan foto resmi default
  if (galeriArr.length === 0) {
    galeriArr.push('https://lh3.googleusercontent.com/d/1WTMWbiQ9B2S7F27ThIlJSdjwKVi6kEq5');
  }
  profile.FotoGaleri = galeriArr;

  // Parse Video Galeri Rumah Air Bogor (Google Drive Video / YouTube / MP4)
  var videoRaw = profile.VideoGaleriResto || '';
  var videoArr = [];
  var splitVideo = videoRaw.split(/[\n,]+/);
  for (var v = 0; v < splitVideo.length; v++) {
    var videoItem = splitVideo[v].trim();
    if (videoItem) {
      videoArr.push(videoItem);
    }
  }
  profile.VideoGaleri = videoArr;

  return profile;
}

/**
 * Helper: Ambil daftar Menu Makanan & Minuman
 */
function fetchMenuList(ss) {
  var sheet = ss.getSheetByName('DaftarMenu');
  var menuList = [];

  if (!sheet || sheet.getLastRow() <= 1) return menuList;

  var lastRow = sheet.getLastRow();
  var numCols = 8;
  var rows = sheet.getRange(2, 1, lastRow - 1, numCols).getDisplayValues();

  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var status = row[COL_MENU.STATUS] ? row[COL_MENU.STATUS].toString().trim() : 'Tersedia';

    // Lewati jika menu nonaktif/habis jika diinginkan, namun tetap sertakan status
    menuList.push({
      id: row[COL_MENU.ID] || ('MNU-' + (i + 1)),
      kategori: row[COL_MENU.KATEGORI] || 'Lainnya',
      nama: row[COL_MENU.NAMA] || '',
      deskripsi: row[COL_MENU.DESKRIPSI] || '',
      hargaFormatted: formatRupiahString(row[COL_MENU.HARGA]),
      hargaRaw: parseNumber(row[COL_MENU.HARGA]),
      foto: formatDirectImageUrl(row[COL_MENU.FOTO] || 'https://images.unsplash.com/photo-1546069901-ba9599a7e63c?w=600&q=80'),
      status: status,
      badge: row[COL_MENU.BADGE] || ''
    });
  }

  return menuList;
}

/**
 * Helper: Ambil daftar Promo & Banner Berjalan
 */
function fetchPromoList(ss) {
  var sheet = ss.getSheetByName('PromoAcara');
  var promoList = [];

  if (!sheet || sheet.getLastRow() <= 1) return promoList;

  var lastRow = sheet.getLastRow();
  var numCols = Math.max(sheet.getLastColumn(), 5);
  var headerValues = sheet.getRange(1, 1, 1, numCols).getDisplayValues()[0];

  // Deteksi indeks kolom dinamis
  var idCol = 0;
  var judulCol = 1;
  var ketCol = 2;
  var bannerCol = 3;
  var statusCol = 4;

  for (var c = 0; c < headerValues.length; c++) {
    var h = headerValues[c].toString().trim().toLowerCase();
    if (h.indexOf('status') !== -1) statusCol = c;
    else if (h.indexOf('id') !== -1) idCol = c;
    else if (h.indexOf('judul') !== -1) judulCol = c;
    else if (h.indexOf('keterangan') !== -1 || h.indexOf('deskripsi') !== -1) ketCol = c;
    else if (h.indexOf('banner') !== -1 || h.indexOf('foto') !== -1 || h.indexOf('gambar') !== -1) bannerCol = c;
  }

  var rows = sheet.getRange(2, 1, lastRow - 1, numCols).getDisplayValues();

  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var status = row[statusCol] ? row[statusCol].toString().trim().toLowerCase() : '';

    // Hanya lolos jika status benar-benar aktif (mengabaikan tidak aktif, nonaktif, off, dsb)
    if (status === 'aktif' || status === 'active') {
      var rawBanner = row[bannerCol] ? row[bannerCol].toString().trim() : '';
      promoList.push({
        id: row[idCol] || ('PRM-' + (i + 1)),
        judul: row[judulCol] || 'Promo Spesial Rumah Air Bogor',
        keterangan: row[ketCol] || '',
        banner: rawBanner ? formatDirectImageUrl(rawBanner) : 'https://lh3.googleusercontent.com/d/1WTMWbiQ9B2S7F27ThIlJSdjwKVi6kEq5',
        status: 'Aktif'
      });
    }
  }

  return promoList;
}

/**
 * Helper: Ambil daftar foto Banner Menu Regullar dari tab BannerMenu
 */
function fetchBannerMenuList(ss) {
  var sheet = ss.getSheetByName('BannerMenu');
  var bannerList = [];
  var fallbackBanner = 'https://lh3.googleusercontent.com/d/1WTMWbiQ9B2S7F27ThIlJSdjwKVi6kEq5';

  if (!sheet || sheet.getLastRow() <= 1) {
    return [fallbackBanner];
  }

  var lastRow = sheet.getLastRow();
  var rows = sheet.getRange(2, 1, lastRow - 1, 4).getDisplayValues();

  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var status = row[COL_BANNER_MENU.STATUS] ? row[COL_BANNER_MENU.STATUS].toString().trim() : 'Aktif';
    var rawUrl = row[COL_BANNER_MENU.FOTO] ? row[COL_BANNER_MENU.FOTO].toString().trim() : '';

    if (status.toLowerCase() === 'aktif' && rawUrl) {
      bannerList.push(formatDirectImageUrl(rawUrl));
    }
  }

  if (bannerList.length === 0) {
    bannerList.push(fallbackBanner);
  }

  return bannerList;
}

/**
 * Helper: Ambil daftar foto, video & keterangan dari tab GaleriResto
 */
function fetchGaleriList(ss) {
  var sheet = ss.getSheetByName('GaleriResto');
  var galeriList = [];
  var fallbackItem = {
    id: 'GL-001',
    judul: 'Saung Utama & Danau Nirwana',
    url: 'https://lh3.googleusercontent.com/d/1WTMWbiQ9B2S7F27ThIlJSdjwKVi6kEq5',
    rawUrl: 'https://drive.google.com/file/d/1WTMWbiQ9B2S7F27ThIlJSdjwKVi6kEq5/view?usp=drive_link',
    tipe: 'Foto',
    status: 'Aktif'
  };

  if (!sheet || sheet.getLastRow() <= 1) {
    return [fallbackItem];
  }

  var lastRow = sheet.getLastRow();
  var rows = sheet.getRange(2, 1, lastRow - 1, 5).getDisplayValues();

  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var status = row[COL_GALERI.STATUS] ? row[COL_GALERI.STATUS].toString().trim() : 'Aktif';
    var rawUrl = row[COL_GALERI.MEDIA] ? row[COL_GALERI.MEDIA].toString().trim() : '';

    if (status.toLowerCase() === 'aktif' && rawUrl) {
      galeriList.push({
        id: row[COL_GALERI.ID] || ('GL-' + (i + 1)),
        judul: row[COL_GALERI.JUDUL] || 'Suasana Rumah Air Bogor',
        url: formatDirectImageUrl(rawUrl),
        rawUrl: rawUrl,
        tipe: row[COL_GALERI.TIPE] || 'Foto',
        status: status
      });
    }
  }

  if (galeriList.length === 0) {
    galeriList.push(fallbackItem);
  }

  return galeriList;
}

/**
 * Helper: Hitung status buka/tutup resto realtime (WIB)
 */
function calculateOpenStatus(profile) {
  var now = new Date();
  var hours = now.getHours();
  var minutes = now.getMinutes();
  var currentMinutes = hours * 60 + minutes;

  // Standar: 10:00 (600 menit) sampai 21:00 (1260 menit)
  var openTime = 10 * 60;
  var closeTime = 21 * 60;

  var isOpen = currentMinutes >= openTime && currentMinutes <= closeTime;

  return {
    isOpen: isOpen,
    label: isOpen ? 'Buka Sekarang' : 'Tutup Sementara',
    jamInfo: profile.JamOperasional || '10:00 - 21:00 WIB'
  };
}

/**
 * TRANSAKSI: Simpan Reservasi Masuk & Buat Tautan WhatsApp
 * Dilindungi LockService untuk mencegah tabrakan penomoran ID otomatis saat multi-user.
 */
function submitReservasi(payload) {
  var lock = LockService.getScriptLock();
  try {
    // Tunggu antrean lock hingga 15 detik
    lock.waitLock(15000);

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      return { success: false, message: 'Koneksi database tidak tersedia.' };
    }

    var sheet = ss.getSheetByName('DataReservasi');
    if (!sheet) {
      // Auto create jika belum ada
      setupDataReservasi(ss);
      sheet = ss.getSheetByName('DataReservasi');
    }

    // Ekstraksi data primitif murni & sanitasi input
    var namaTamu = sanitizeString(payload.nama);
    var nomorWA = sanitizeString(payload.nomorWA);
    var tanggalAcara = sanitizeString(payload.tanggal);
    var jamAcara = sanitizeString(payload.jam);
    var jumlahPax = parseInt(payload.pax, 10) || 1;
    var pilihanArea = sanitizeString(payload.area) || 'Saung Atas Air';
    var catatanKhusus = sanitizeString(payload.catatan) || '-';

    if (!namaTamu || !nomorWA || !tanggalAcara || !jamAcara) {
      return { success: false, message: 'Mohon lengkapi seluruh formulir reservasi wajib.' };
    }

    // Generator Penomoran Unik Sekuensial Otomatis: RAB-YYYYMM-XXX
    var now = new Date();
    var periodPrefix = 'RAB-' + Utilities.formatDate(now, 'Asia/Jakarta', 'yyyyMM') + '-';
    var nextSeq = getNextSequenceNumber(sheet, periodPrefix);
    var noReservasi = periodPrefix + padZero(nextSeq, 3);
    var waktuInput = Utilities.formatDate(now, 'Asia/Jakarta', 'dd/MM/yyyy HH:mm:ss') + ' WIB';

    // Tambahkan baris baru ke sheet DataReservasi
    var newRow = [
      noReservasi,
      waktuInput,
      namaTamu,
      nomorWA,
      tanggalAcara,
      jamAcara,
      jumlahPax,
      pilihanArea,
      catatanKhusus,
      'Menunggu Konfirmasi'
    ];

    sheet.appendRow(newRow);

    // Paksa sinkronisasi langsung ke Google Sheets fisik
    SpreadsheetApp.flush();

    // Ambil nomor WA resmi Resto dari Profil
    var profile = fetchProfileData(ss);
    var targetPhone = profile.NoWhatsAppIntl || '6281111861853';

    // Susun pesan WhatsApp yang rapi, elegan & profesional
    var waMessage =
      'Halo Admin *Rumah Air Bogor Resto*,\n' +
      'Saya ingin mengonfirmasi *Reservasi Meja / Saung* dengan detail berikut:\n\n' +
      '📋 *No. Reservasi:* ' + noReservasi + '\n' +
      '👤 *Nama Pemesan:* ' + namaTamu + '\n' +
      '📱 *Nomor Kontak:* ' + nomorWA + '\n' +
      '📅 *Tanggal:* ' + tanggalAcara + '\n' +
      '⏰ *Waktu/Jam:* ' + jamAcara + ' WIB\n' +
      '👥 *Jumlah Tamu:* ' + jumlahPax + ' Orang (Pax)\n' +
      '🏞️ *Pilihan Area:* ' + pilihanArea + '\n' +
      '📝 *Catatan Khusus:* ' + catatanKhusus + '\n\n' +
      'Mohon informasikan ketersediaan saung dan kelengkapan reservasi kami. Terima kasih!';

    var waUrl = 'https://api.whatsapp.com/send?phone=' + targetPhone + '&text=' + encodeURIComponent(waMessage);

    return {
      success: true,
      noReservasi: noReservasi,
      waUrl: waUrl,
      message: 'Reservasi Anda berhasil dicatat dengan nomor ' + noReservasi + '!'
    };

  } catch (error) {
    Logger.log('Error submitReservasi: ' + error.toString());
    return {
      success: false,
      message: 'Terjadi kendala saat mencatat reservasi: ' + error.message
    };
  } finally {
    // Lepas penguncian database
    lock.releaseLock();
  }
}

/**
 * Logika Backend Anti-Tabrakan: Memindai nomor urut tertinggi (MAX + 1)
 */
function getNextSequenceNumber(sheet, prefix) {
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 1;

  var idValues = sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues();
  var maxSeq = 0;

  for (var i = 0; i < idValues.length; i++) {
    var idStr = idValues[i][0];
    if (idStr && idStr.indexOf(prefix) === 0) {
      var seqPart = idStr.substring(prefix.length);
      var seqNum = parseInt(seqPart, 10);
      if (!isNaN(seqNum) && seqNum > maxSeq) {
        maxSeq = seqNum;
      }
    }
  }

  return maxSeq + 1;
}

/**
 * Helper Utilities
 */
function padZero(num, size) {
  var s = num + '';
  while (s.length < size) s = '0' + s;
  return s;
}

function sanitizeString(str) {
  if (!str) return '';
  return str.toString().replace(/[<>]/g, '').trim();
}

function parseNumber(val) {
  if (typeof val === 'number') return val;
  if (!val) return 0;
  var cleaned = val.toString().replace(/[^0-9]/g, '');
  return parseInt(cleaned, 10) || 0;
}

function formatRupiahString(val) {
  var num = parseNumber(val);
  return 'Rp ' + num.toString().replace(/\B(?=(\d{3})+(?!\d))/g, '.');
}

/**
 * Helper: Mengonversi tautan Google Drive menjadi link gambar langsung CDN (lh3.googleusercontent.com)
 */
function formatDirectImageUrl(url) {
  if (!url) return '';
  var str = url.toString().trim();

  // Pola 1: drive.google.com/file/d/FILE_ID/view...
  var match1 = str.match(/\/file\/d\/([a-zA-Z0-9_-]+)/);
  if (match1 && match1[1]) {
    return 'https://lh3.googleusercontent.com/d/' + match1[1];
  }

  // Pola 2: drive.google.com/open?id=FILE_ID atau uc?id=FILE_ID
  var match2 = str.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (match2 && match2[1]) {
    return 'https://lh3.googleusercontent.com/d/' + match2[1];
  }

  return str;
}

