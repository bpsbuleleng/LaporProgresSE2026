/**
 * Aplikasi Pelaporan Progres Pendataan SE2026 — BPS Kabupaten Buleleng
 * Backend Google Apps Script (Web App).
 *
 * Database : Google Spreadsheet (sheet: master, laporan, riwayat)
 * Foto      : Google Drive (1 foto terbaru per subSLS)
 *
 * Alur: Petugas (PPL/PML) melapor progres per subSLS + unggah screenshot Fasih Mobile.
 */

// ====== Konfigurasi sumber data ======
const SPREADSHEET_ID = '1C_JzJhRYb7n_qxzHdNjKlnIbPZ28k3JbyRqXKN5TjEI'; // database
const FOLDER_ID      = '1oFg7jRb1Pu4B7HHW3jcH0Zwhi-mIiuj_';            // folder foto
const TZ             = 'Asia/Makassar';                                 // WITA (Bali/Buleleng)
// Sumber hasil sistem (dashboard monitoring publik, tanpa login)
const MONITORING_URL = 'https://bpsbuleleng.github.io/MonitoringSE2026/';

// Nama sheet + header (urutan kolom wajib sama)
const SHEETS = {
  master:  ['id_subsls', 'kecamatan', 'desa_kelurahan', 'sls', 'subsls', 'ppl', 'pml'],
  laporan: ['id_subsls', 'open', 'submit', 'reject', 'approved', 'url_foto', 'updated_at'],
  riwayat: ['tanggal', 'id_subsls', 'open', 'submit', 'reject', 'approved'],
  // Muatan/target wilayah per subSLS (referensi prefilled, diisi via impor)
  muatan:  ['id_subsls', 'jml_keluarga', 'keluarga_lainnya', 'jml_usaha', 'usaha_lainnya']
};

// ====== Entry point Web App ======
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Lapor Progres SE2026 — BPS Buleleng')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ====== Setup: buat sheet + header bila belum ada ======
function setup() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  Object.keys(SHEETS).forEach(function (name) {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    const header = SHEETS[name];
    // Tulis header hanya bila baris pertama masih kosong
    if (sh.getRange(1, 1).getValue() === '') {
      sh.getRange(1, 1, 1, header.length).setValues([header])
        .setFontWeight('bold').setBackground('#134E8E').setFontColor('#ffffff');
      sh.setFrozenRows(1);
    }
    // Simpan tanggal & updated_at sebagai teks agar pencocokan konsisten
    if (name === 'riwayat') sh.getRange('A2:A').setNumberFormat('@');
    if (name === 'laporan') sh.getRange('G2:G').setNumberFormat('@');
  });
  return 'Setup selesai. Sheet master, laporan, riwayat siap digunakan.';
}

// ====== Helper: ambil sheet / folder dengan pesan error jelas ======
function getSheet_(name) {
  const sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(name);
  if (!sh) throw new Error('Sheet "' + name + '" tidak ditemukan. Jalankan fungsi setup() dahulu.');
  return sh;
}

function getFolder_() {
  try {
    return DriveApp.getFolderById(FOLDER_ID);
  } catch (e) {
    throw new Error('Folder foto tidak ditemukan atau tidak dapat diakses.');
  }
}

// Ambil sheet, buat otomatis + header bila belum ada
function getOrCreateSheet_(name) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, SHEETS[name].length).setValues([SHEETS[name]])
      .setFontWeight('bold').setBackground('#134E8E').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }
  return sh;
}

// Baca sheet yang boleh belum ada (tidak melempar error) — aman untuk data opsional
function readSheetSafe_(name) {
  try { return readSheet_(name); } catch (e) { return []; }
}

// Peta muatan per id_subsls (kosong bila sheet belum ada / belum diisi)
function readMuatanMap_() {
  const map = {};
  readSheetSafe_('muatan').forEach(function (r) {
    map[String(r.id_subsls).trim()] = {
      jml_keluarga: num_(r.jml_keluarga),
      keluarga_lainnya: num_(r.keluarga_lainnya),
      jml_usaha: num_(r.jml_usaha),
      usaha_lainnya: num_(r.usaha_lainnya)
    };
  });
  return map;
}

/**
 * Ambil hasil sistem per subSLS dari dashboard Monitoring SE2026 dan isi sheet `muatan`.
 * Halaman menyimpan datanya sebagai window.DASH_DATA = { fields:[...], rows:[[...]] }.
 * id_subsls dibentuk dari idsls (14 digit) + kdsub (2 digit).
 * Jalankan manual, atau pasang trigger time-based agar otomatis.
 */
function updateMuatan() {
  const res = UrlFetchApp.fetch(MONITORING_URL, { muteHttpExceptions: true, followRedirects: true });
  if (res.getResponseCode() !== 200) {
    throw new Error('Gagal mengambil halaman monitoring (HTTP ' + res.getResponseCode() + ').');
  }
  const html = res.getContentText();

  // Potong blok JSON: dari '{' setelah penanda sampai penutup </script>
  const i = html.indexOf('window.DASH_DATA');
  if (i < 0) throw new Error('Struktur halaman monitoring berubah: window.DASH_DATA tidak ditemukan.');
  const st = html.indexOf('{', i);
  const en = html.indexOf('</script>', st);
  if (st < 0 || en < 0) throw new Error('Struktur halaman monitoring berubah: blok data tidak lengkap.');
  let raw = html.substring(st, en).trim();
  if (raw.slice(-1) === ';') raw = raw.slice(0, -1);

  let D;
  try { D = JSON.parse(raw); }
  catch (e) { throw new Error('Data monitoring gagal diurai: ' + e.message); }

  const f = D.fields || [];
  const ix = {};
  ['idsls', 'kdsub', 'rkkel', 'rkkellain', 'rkush', 'rkushlain'].forEach(function (k) {
    ix[k] = f.indexOf(k);
    if (ix[k] < 0) throw new Error('Kolom "' + k + '" tidak ada di data monitoring.');
  });

  const out = (D.rows || []).map(function (r) {
    const sid = String(r[ix.idsls]) + ('0' + String(r[ix.kdsub])).slice(-2); // 14 + 2 digit
    return [sid, num_(r[ix.rkkel]), num_(r[ix.rkkellain]), num_(r[ix.rkush]), num_(r[ix.rkushlain])];
  });
  if (!out.length) throw new Error('Data monitoring kosong.');

  const sh = getOrCreateSheet_('muatan');
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 5).clearContent();
  sh.getRange(2, 1, out.length, 1).setNumberFormat('@');  // id sebagai teks (hindari notasi ilmiah)
  sh.getRange(2, 1, out.length, 5).setValues(out);

  const info = String(D.updated || '-');
  PropertiesService.getScriptProperties().setProperty('muatan_updated', info);
  return 'Hasil sistem diperbarui: ' + out.length + ' subSLS (data monitoring per ' + info + ').';
}

// Impor massal data muatan (menimpa seluruh isi sheet muatan)
function simpanMuatan(rows) {
  if (!rows || !rows.length) throw new Error('Data muatan kosong.');
  const sh = getOrCreateSheet_('muatan');
  const out = rows.map(function (r) {
    return [String(r.id_subsls).trim(), num_(r.jml_keluarga), num_(r.keluarga_lainnya),
            num_(r.jml_usaha), num_(r.usaha_lainnya)];
  });
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 5).clearContent();
  sh.getRange(2, 1, out.length, 5).setValues(out);
  return 'Muatan tersimpan untuk ' + out.length + ' subSLS.';
}

// Baca seluruh isi sheet menjadi array objek {header: nilai}
function readSheet_(name) {
  const sh = getSheet_(name);
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  const header = values[0];
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (row.join('') === '') continue; // lewati baris kosong
    const obj = {};
    header.forEach(function (h, c) { obj[String(h)] = row[c]; });
    rows.push(obj);
  }
  return rows;
}

// ====== Data untuk Halaman 1 (Pelaporan) ======
function getReportingData() {
  const master = readSheet_('master').map(normalizeMaster_);
  const laporan = {};
  readSheet_('laporan').forEach(function (r) {
    laporan[String(r.id_subsls).trim()] = {
      open: num_(r.open), submit: num_(r.submit), reject: num_(r.reject),
      approved: num_(r.approved), url_foto: String(r.url_foto || ''),
      updated_at: String(r.updated_at || '')
    };
  });
  return {
    master: master, laporan: laporan, muatan: readMuatanMap_(),
    muatan_updated: PropertiesService.getScriptProperties().getProperty('muatan_updated') || ''
  };
}

// ====== Data untuk Halaman 2 (Dashboard) ======
function getDashboardData() {
  const master = readSheet_('master').map(normalizeMaster_);
  const laporan = {};
  readSheet_('laporan').forEach(function (r) {
    laporan[String(r.id_subsls).trim()] = {
      open: num_(r.open), submit: num_(r.submit), reject: num_(r.reject),
      approved: num_(r.approved), url_foto: String(r.url_foto || ''),
      updated_at: String(r.updated_at || '')
    };
  });
  const riwayat = readSheet_('riwayat').map(function (r) {
    return {
      tanggal: fmtTanggal_(r.tanggal),
      id_subsls: String(r.id_subsls).trim(),
      open: num_(r.open), submit: num_(r.submit),
      reject: num_(r.reject), approved: num_(r.approved)
    };
  });
  return { master: master, laporan: laporan, riwayat: riwayat, muatan: readMuatanMap_() };
}

// ====== Simpan laporan (upsert laporan + riwayat + foto) ======
/**
 * payload = {
 *   id_subsls: string,
 *   open|submit|reject|approved: number,
 *   foto: { data: base64Tanpa prefix, mime: 'image/jpeg' } | null
 * }
 */
function saveLaporan(payload) {
  if (!payload || !payload.id_subsls) throw new Error('Data tidak lengkap: id subSLS kosong.');
  const id = String(payload.id_subsls).trim();

  // Validasi id ada di master
  const adaDiMaster = readSheet_('master').some(function (m) {
    return String(m.id_subsls).trim() === id;
  });
  if (!adaDiMaster) throw new Error('subSLS ' + id + ' tidak terdaftar di master.');

  // Validasi angka >= 0
  const open     = validNum_(payload.open, 'open');
  const submit   = validNum_(payload.submit, 'submit');
  const reject   = validNum_(payload.reject, 'reject');
  const approved = validNum_(payload.approved, 'approved');

  // ---- Foto: ganti foto lama bila ada unggahan baru ----
  const shLaporan = getSheet_('laporan');
  const existing  = findRow_(shLaporan, id);          // baris laporan lama (jika ada)
  let urlFoto = existing.found ? String(existing.data[5] || '') : '';

  if (payload.foto && payload.foto.data) {
    const folder = getFolder_();
    let file;
    try {
      hapusFotoLama_(folder, id);                      // selalu sisakan 1 foto terbaru
      const stamp = Utilities.formatDate(new Date(), TZ, 'yyyyMMdd-HHmmss');
      const blob  = Utilities.newBlob(
        Utilities.base64Decode(payload.foto.data),
        payload.foto.mime || 'image/jpeg',
        id + '_' + stamp + '.jpg'
      );
      file = folder.createFile(blob);
    } catch (e) {
      // Umumnya karena akun pemilik Web App belum punya izin EDITOR pada folder
      throw new Error('Foto gagal diunggah: ' + e.message +
        '. Pastikan akun pemilik Web App memiliki izin Editor pada folder Drive tujuan.');
    }
    // Bagikan agar thumbnail tampil untuk pengguna Web App (non-fatal bila dibatasi domain)
    try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
    // URL thumbnail Drive dapat ditampilkan langsung di <img>
    urlFoto = 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w1280';
  }

  const now = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss');

  // ---- Upsert laporan ----
  const rowLaporan = [id, open, submit, reject, approved, urlFoto, now];
  if (existing.found) {
    shLaporan.getRange(existing.row, 1, 1, rowLaporan.length).setValues([rowLaporan]);
  } else {
    shLaporan.appendRow(rowLaporan);
  }

  // ---- Upsert riwayat (kunci: id_subsls + tanggal hari ini) ----
  upsertRiwayat_(id, open, submit, reject, approved);

  return {
    ok: true,
    id_subsls: id,
    laporan: {
      open: open, submit: submit, reject: reject, approved: approved,
      url_foto: urlFoto, updated_at: now
    }
  };
}

// ---- Timpa/isi baris riwayat untuk id + tanggal hari ini ----
function upsertRiwayat_(id, open, submit, reject, approved) {
  const sh = getSheet_('riwayat');
  const hari = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  const values = sh.getDataRange().getValues();
  const baris = [hari, id, open, submit, reject, approved];
  for (let i = 1; i < values.length; i++) {
    const tgl = fmtTanggal_(values[i][0]);
    const rid = String(values[i][1]).trim();
    if (tgl === hari && rid === id) {                  // sudah ada → timpa
      sh.getRange(i + 1, 1, 1, baris.length).setValues([baris]);
      return;
    }
  }
  sh.appendRow(baris);                                 // belum ada → tambah
}

// ---- Hapus semua foto lama ber-prefix "{id}_" di folder ----
function hapusFotoLama_(folder, id) {
  const prefix = id + '_';
  const it = folder.searchFiles('title contains "' + prefix.replace(/"/g, '') + '"');
  while (it.hasNext()) {
    const f = it.next();
    if (f.getName().indexOf(prefix) === 0) f.setTrashed(true);
  }
}

// ---- Cari baris berdasarkan id_subsls (kolom A); kembalikan {found,row,data} ----
function findRow_(sh, id) {
  const values = sh.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]).trim() === id) {
      return { found: true, row: i + 1, data: values[i] };
    }
  }
  return { found: false, row: -1, data: null };
}

// ====== Util kecil ======
function normalizeMaster_(m) {
  return {
    id_subsls: String(m.id_subsls).trim(),
    kecamatan: String(m.kecamatan || '').trim(),
    desa_kelurahan: String(m.desa_kelurahan || '').trim(),
    sls: String(m.sls || '').trim(),
    subsls: String(m.subsls || '').trim(),
    ppl: String(m.ppl || '').trim(),
    pml: String(m.pml || '').trim()
  };
}

function num_(v) { const n = Number(v); return isFinite(n) ? n : 0; }

function validNum_(v, nama) {
  const n = Number(v);
  if (!isFinite(n) || n < 0) throw new Error('Nilai "' + nama + '" harus angka ≥ 0.');
  return Math.floor(n);
}

// Format nilai tanggal (Date atau string) menjadi 'yyyy-MM-dd'
function fmtTanggal_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  return String(v || '').trim();
}

// ====== Diagnostik izin (jalankan manual di editor bila foto gagal) ======
function cekAkses() {
  const out = [];
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    out.push('✔ Spreadsheet terbaca: ' + ss.getName());
  } catch (e) { out.push('✗ Spreadsheet GAGAL: ' + e.message); }
  try {
    const f = DriveApp.getFolderById(FOLDER_ID);
    out.push('✔ Folder terbaca: ' + f.getName());
    const t = f.createFile(Utilities.newBlob('cek', 'text/plain', '__cek_akses.txt'));
    t.setTrashed(true);
    out.push('✔ Tulis ke folder OK — izin Editor tersedia.');
  } catch (e) {
    out.push('✗ Tulis folder GAGAL: ' + e.message +
      ' → bagikan folder Drive sebagai EDITOR ke akun ' + Session.getEffectiveUser().getEmail() + '.');
  }
  const msg = out.join('\n');
  Logger.log(msg);
  return msg;
}

// ====== Ambil foto sebagai data URL (fallback preview bila thumbnail gagal) ======
function getFoto(fileId) {
  try {
    const blob = DriveApp.getFileById(fileId).getBlob();
    return 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes());
  } catch (e) {
    throw new Error('Foto tidak dapat dimuat: ' + e.message);
  }
}
