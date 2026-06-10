// ============================================================
// Code.gs - Google Apps Script Backend
// Sistem Absensi SMKN Rakit Kulim - Face + GPS
// ============================================================

// ==================== KONFIGURASI ==========================
const CONFIG = {
  SHEET_SISWA:   "DataSiswa",
  SHEET_ABSEN:   "Absensi",
  SHEET_SETTING: "Pengaturan",
  SHEET_ADMIN:   "Admin",
  SHEET_LIBUR:   "HariLibur",
  SHEET_JADWAL:  "JadwalHari",
  RADIUS_METER:  100,
};

// ==================== MAIN HANDLER ==========================
function doGet(e) {
  const action = e.parameter.action;
  let result;
  try {
    switch (action) {
      case "getSetting":    result = getSetting(); break;
      case "getSiswa":      result = getSiswaList(); break;
      case "getAbsensi":    result = getAbsensi(e.parameter); break;
      case "getRekapBulan": result = getRekapBulan(e.parameter); break;
      case "getLibur":      result = getLibur(); break;
      case "getJadwal":     result = getJadwal(); break;
      default: result = { status: "error", message: "Action tidak dikenal" };
    }
  } catch (err) {
    result = { status: "error", message: err.message };
  }
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  let body, result;
  try {
    body = JSON.parse(e.postData.contents);
    switch (body.action) {
      case "absen":         result = absenSiswa(body); break;
      case "loginAdmin":    result = loginAdmin(body); break;
      case "daftarSiswa":   result = daftarSiswa(body); break;
      case "updateSiswa":   result = updateSiswa(body); break;
      case "hapusSiswa":    result = hapusSiswa(body); break;
      case "saveSetting":   result = saveSetting(body); break;
      case "hapusAbsensi":  result = hapusAbsensi(body); break;
      case "tambahLibur":   result = tambahLibur(body); break;
      case "hapusLibur":    result = hapusLibur(body); break;
      case "saveJadwal":    result = saveJadwal(body); break;
      default: result = { status: "error", message: "Action tidak dikenal" };
    }
  } catch (err) {
    result = { status: "error", message: err.message };
  }
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ==================== INISIALISASI SHEET ====================
function initSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Sheet DataSiswa
  let sh = ss.getSheetByName(CONFIG.SHEET_SISWA);
  if (!sh) {
    sh = ss.insertSheet(CONFIG.SHEET_SISWA);
    sh.appendRow(["NIS","Nama","Kelas","Foto_Descriptor","Tanggal_Daftar"]);
    sh.getRange(1,1,1,5).setFontWeight("bold").setBackground("#1a73e8").setFontColor("#ffffff");
  }

  // Sheet Absensi
  sh = ss.getSheetByName(CONFIG.SHEET_ABSEN);
  if (!sh) {
    sh = ss.insertSheet(CONFIG.SHEET_ABSEN);
    sh.appendRow(["ID","NIS","Nama","Kelas","Tanggal","Waktu","Status","Latitude","Longitude","Jarak_Meter","Catatan"]);
    sh.getRange(1,1,1,11).setFontWeight("bold").setBackground("#1a73e8").setFontColor("#ffffff");
  }

  // Sheet Pengaturan
  sh = ss.getSheetByName(CONFIG.SHEET_SETTING);
  if (!sh) {
    sh = ss.insertSheet(CONFIG.SHEET_SETTING);
    sh.appendRow(["Key","Value"]);
    sh.appendRow(["NamaSekolah","SMKN Rakit Kulim"]);
    sh.appendRow(["Alamat","Jl. Pendidikan No. 1"]);
    sh.appendRow(["LatSekolah","-0.5897"]);
    sh.appendRow(["LngSekolah","101.4478"]);
    sh.appendRow(["RadiusMeter","100"]);
    sh.appendRow(["JamMasuk","07:00"]);
    sh.appendRow(["JamPulang","15:00"]);
    sh.appendRow(["JamTelat","07:30"]);
    sh.getRange(1,1,1,2).setFontWeight("bold").setBackground("#1a73e8").setFontColor("#ffffff");
  }

  // Sheet Admin
  sh = ss.getSheetByName(CONFIG.SHEET_ADMIN);
  if (!sh) {
    sh = ss.insertSheet(CONFIG.SHEET_ADMIN);
    sh.appendRow(["Username","Password","Nama","Role"]);
    sh.appendRow(["admin","admin123","Administrator","superadmin"]);
    sh.getRange(1,1,1,4).setFontWeight("bold").setBackground("#1a73e8").setFontColor("#ffffff");
  }

  // Sheet HariLibur
  sh = ss.getSheetByName(CONFIG.SHEET_LIBUR);
  if (!sh) {
    sh = ss.insertSheet(CONFIG.SHEET_LIBUR);
    sh.appendRow(["ID","TglMulai","TglAkhir","Keterangan","DibuatOleh"]);
    sh.getRange(1,1,1,5).setFontWeight("bold").setBackground("#1a73e8").setFontColor("#ffffff");
  }

  // Sheet JadwalHari — jadwal khusus per hari
  sh = ss.getSheetByName(CONFIG.SHEET_JADWAL);
  if (!sh) {
    sh = ss.insertSheet(CONFIG.SHEET_JADWAL);
    sh.appendRow(["Hari","NamaHari","JamMasuk","JamTelat","JamPulang","Aktif"]);
    // Default: Senin s/d Sabtu, isi dari setting umum kecuali Senin & Jumat
    const defaults = [
      ["0","Minggu","07:00","07:30","15:00","false"],
      ["1","Senin","07:15","07:45","15:00","true"],
      ["2","Selasa","07:00","07:30","15:00","true"],
      ["3","Rabu","07:00","07:30","15:00","true"],
      ["4","Kamis","07:00","07:30","15:00","true"],
      ["5","Jumat","07:00","07:30","11:30","true"],
      ["6","Sabtu","07:00","07:30","13:00","true"],
    ];
    defaults.forEach(row => sh.appendRow(row));
    sh.getRange(1,1,1,6).setFontWeight("bold").setBackground("#1a73e8").setFontColor("#ffffff");
  }

  return { status: "ok", message: "Semua sheet berhasil diinisialisasi" };
}

// ==================== JADWAL HARI ===========================
function getJadwal() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_JADWAL);
  if (!sh) {
    // Kembalikan default hardcode jika sheet belum ada
    return { status: "ok", data: [
      {hari:"0",namaHari:"Minggu",jamMasuk:"07:00",jamTelat:"07:30",jamPulang:"15:00",aktif:false},
      {hari:"1",namaHari:"Senin",jamMasuk:"07:15",jamTelat:"07:45",jamPulang:"15:00",aktif:true},
      {hari:"2",namaHari:"Selasa",jamMasuk:"07:00",jamTelat:"07:30",jamPulang:"15:00",aktif:true},
      {hari:"3",namaHari:"Rabu",jamMasuk:"07:00",jamTelat:"07:30",jamPulang:"15:00",aktif:true},
      {hari:"4",namaHari:"Kamis",jamMasuk:"07:00",jamTelat:"07:30",jamPulang:"15:00",aktif:true},
      {hari:"5",namaHari:"Jumat",jamMasuk:"07:00",jamTelat:"07:30",jamPulang:"11:30",aktif:true},
      {hari:"6",namaHari:"Sabtu",jamMasuk:"07:00",jamTelat:"07:30",jamPulang:"13:00",aktif:true},
    ]};
  }
  const rows = sh.getDataRange().getValues();
  const list = [];
  for (let i = 1; i < rows.length; i++) {
    if (!rows[i][0].toString()) continue;
    list.push({
      hari:      rows[i][0].toString(),
      namaHari:  rows[i][1].toString(),
      jamMasuk:  rows[i][2].toString(),
      jamTelat:  rows[i][3].toString(),
      jamPulang: rows[i][4].toString(),
      aktif:     rows[i][5].toString() === "true"
    });
  }
  return { status: "ok", data: list };
}

function saveJadwal(body) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_JADWAL);
  if (!sh) return { status: "error", message: "Sheet JadwalHari tidak ditemukan. Jalankan initSheets() dulu." };
  const rows = sh.getDataRange().getValues();
  // body.data = array of {hari, jamMasuk, jamTelat, jamPulang, aktif}
  for (const item of body.data) {
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0].toString() === item.hari.toString()) {
        sh.getRange(i+1, 3).setValue(item.jamMasuk);
        sh.getRange(i+1, 4).setValue(item.jamTelat);
        sh.getRange(i+1, 5).setValue(item.jamPulang);
        sh.getRange(i+1, 6).setValue(item.aktif ? "true" : "false");
        break;
      }
    }
  }
  return { status: "ok", message: "Jadwal berhasil disimpan" };
}

// ==================== HARI LIBUR ============================
function getLibur() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_LIBUR);
  if (!sh) return { status: "ok", data: [] };
  const data = sh.getDataRange().getValues();
  const list = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    list.push({
      id:          data[i][0].toString(),
      tglMulai:    data[i][1].toString(),
      tglAkhir:    data[i][2].toString(),
      keterangan:  data[i][3].toString()
    });
  }
  return { status: "ok", data: list };
}

function tambahLibur(body) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_LIBUR);
  if (!sh) return { status: "error", message: "Sheet HariLibur tidak ditemukan. Jalankan initSheets() dulu." };
  const id = "LBR" + new Date().getTime();
  sh.appendRow([id, body.tglMulai, body.tglAkhir, body.keterangan, "admin"]);
  return { status: "ok", message: "Hari libur ditambahkan" };
}

function hapusLibur(body) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_LIBUR);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() === body.id.toString()) {
      sh.deleteRow(i + 1);
      return { status: "ok", message: "Hari libur dihapus" };
    }
  }
  return { status: "error", message: "Data tidak ditemukan" };
}

// ==================== PENGATURAN ============================
function getSetting() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_SETTING);
  if (!sh) return { status: "error", message: "Sheet Pengaturan tidak ditemukan" };
  const data = sh.getDataRange().getValues();
  const setting = {};
  for (let i = 1; i < data.length; i++) {
    setting[data[i][0]] = data[i][1];
  }
  return { status: "ok", data: setting };
}

function saveSetting(body) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_SETTING);
  const data = sh.getDataRange().getValues();
  const updates = body.data;
  for (const key in updates) {
    let found = false;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === key) {
        sh.getRange(i+1, 2).setValue(updates[key]);
        found = true; break;
      }
    }
    if (!found) sh.appendRow([key, updates[key]]);
  }
  return { status: "ok", message: "Pengaturan disimpan" };
}

// ==================== DATA SISWA ============================
function getSiswaList() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_SISWA);
  if (!sh) return { status: "error", message: "Sheet DataSiswa tidak ditemukan" };
  const data = sh.getDataRange().getValues();
  const list = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) {
      list.push({
        nis:          data[i][0].toString(),
        nama:         data[i][1],
        kelas:        data[i][2],
        descriptor:   data[i][3] ? JSON.parse(data[i][3]) : null,
        tanggalDaftar: data[i][4]
      });
    }
  }
  return { status: "ok", data: list };
}

function daftarSiswa(body) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_SISWA);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() === body.nis.toString()) {
      return { status: "error", message: "NIS sudah terdaftar" };
    }
  }
  sh.appendRow([body.nis, body.nama, body.kelas, JSON.stringify(body.descriptor), getWIBDate()]);
  return { status: "ok", message: "Siswa berhasil didaftarkan" };
}

function updateSiswa(body) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_SISWA);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() === body.nis.toString()) {
      sh.getRange(i+1, 2).setValue(body.nama);
      sh.getRange(i+1, 3).setValue(body.kelas);
      if (body.descriptor) sh.getRange(i+1, 4).setValue(JSON.stringify(body.descriptor));
      return { status: "ok", message: "Data siswa diperbarui" };
    }
  }
  return { status: "error", message: "Siswa tidak ditemukan" };
}

function hapusSiswa(body) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_SISWA);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() === body.nis.toString()) {
      sh.deleteRow(i+1);
      return { status: "ok", message: "Siswa dihapus" };
    }
  }
  return { status: "error", message: "Siswa tidak ditemukan" };
}

// ==================== ABSENSI ===============================
function absenSiswa(body) {
  const now  = new Date();
  const wib  = new Date(now.getTime() + 7 * 60 * 60 * 1000);
  const pad  = n => String(n).padStart(2, "0");
  const tanggal = wib.getUTCFullYear() + "-" + pad(wib.getUTCMonth()+1) + "-" + pad(wib.getUTCDate());
  const hariIdx = wib.getUTCDay(); // 0=Minggu, 1=Senin, …

  // ── Cek hari libur ──────────────────────────────────────
  const shLibur = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_LIBUR);
  if (shLibur) {
    const rows = shLibur.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (!rows[i][0]) continue;
      if (tanggal >= rows[i][1].toString() && tanggal <= rows[i][2].toString()) {
        return { status: "error", message: "Hari ini adalah hari libur: " + rows[i][3] };
      }
    }
  }

  // ── Ambil jadwal hari ini ────────────────────────────────
  const jadwalRes = getJadwal();
  const jadwalList = jadwalRes.data || [];
  const jadwalHariIni = jadwalList.find(j => j.hari === String(hariIdx));

  // Jika hari tidak aktif (mis. Minggu) tolak absen
  if (jadwalHariIni && jadwalHariIni.aktif === false) {
    return { status: "error", message: "Hari " + (jadwalHariIni.namaHari || "ini") + " bukan hari sekolah." };
  }

  // Gunakan jam dari jadwal hari, fallback ke setting umum
  const setting = getSetting().data;
  const jamTelat  = jadwalHariIni ? jadwalHariIni.jamTelat  : (setting.JamTelat  || "07:30");
  const jamPulang = jadwalHariIni ? jadwalHariIni.jamPulang : (setting.JamPulang || "15:00");

  // ── Cek GPS ─────────────────────────────────────────────
  const lat1   = parseFloat(setting.LatSekolah);
  const lng1   = parseFloat(setting.LngSekolah);
  const radius = parseFloat(setting.RadiusMeter) || CONFIG.RADIUS_METER;
  const lat2   = parseFloat(body.lat);
  const lng2   = parseFloat(body.lng);
  const jarak  = hitungJarak(lat1, lng1, lat2, lng2);
  if (jarak > radius) {
    return { status: "error", message: `Anda berada ${Math.round(jarak)}m dari sekolah. Batas: ${radius}m` };
  }

  const waktu = pad(wib.getUTCHours()) + ":" + pad(wib.getUTCMinutes()) + ":" + pad(wib.getUTCSeconds());

  // ── Cek sudah absen ──────────────────────────────────────
  const shAbsen = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_ABSEN);
  const dataAbsen = shAbsen.getDataRange().getValues();
  for (let i = 1; i < dataAbsen.length; i++) {
    if (dataAbsen[i][1].toString() === body.nis.toString() &&
        dataAbsen[i][4].toString() === tanggal) {
      return { status: "error", message: "Anda sudah absen hari ini pukul " + dataAbsen[i][5] };
    }
  }

  // ── Tentukan status ──────────────────────────────────────
  const [jB, mB] = jamTelat.split(":").map(Number);
  const menitAbsen = wib.getUTCHours() * 60 + wib.getUTCMinutes();
  const menitBatas = jB * 60 + mB;
  const status = menitAbsen <= menitBatas ? "Hadir" : "Terlambat";

  const id = "ABS" + now.getTime();
  shAbsen.appendRow([id, body.nis, body.nama, body.kelas, tanggal, waktu, status, lat2, lng2, Math.round(jarak), ""]);

  return {
    status: "ok",
    message: `Absensi berhasil! Status: ${status}`,
    data: {
      tanggal, waktu, status,
      jarak: Math.round(jarak),
      jamMasuk:  jadwalHariIni ? jadwalHariIni.jamMasuk  : (setting.JamMasuk  || "07:00"),
      jamPulang: jamPulang,
      hariNama:  jadwalHariIni ? jadwalHariIni.namaHari : ""
    }
  };
}

function getAbsensi(params) {
  const sh   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_ABSEN);
  const data = sh.getDataRange().getValues();
  const list = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    const row = {
      id: data[i][0], nis: data[i][1].toString(), nama: data[i][2],
      kelas: data[i][3], tanggal: data[i][4], waktu: data[i][5],
      status: data[i][6], lat: data[i][7], lng: data[i][8],
      jarak: data[i][9], catatan: data[i][10]
    };
    if (params.tanggal && row.tanggal.toString() !== params.tanggal) continue;
    if (params.kelas   && row.kelas   !== params.kelas)   continue;
    if (params.nis     && row.nis     !== params.nis)      continue;
    list.push(row);
  }
  return { status: "ok", data: list };
}

function getRekapBulan(params) {
  const sh   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_ABSEN);
  const data = sh.getDataRange().getValues();
  const rekap = {};
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    const tgl = data[i][4] ? data[i][4].toString() : "";
    if (!tgl.startsWith(params.bulan)) continue;
    const nis = data[i][1].toString();
    if (!rekap[nis]) rekap[nis] = { nis, nama: data[i][2], kelas: data[i][3], hadir: 0, terlambat: 0 };
    if (data[i][6] === "Hadir")      rekap[nis].hadir++;
    else if (data[i][6] === "Terlambat") rekap[nis].terlambat++;
  }
  return { status: "ok", data: Object.values(rekap) };
}

function hapusAbsensi(body) {
  const sh   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_ABSEN);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() === body.id.toString()) {
      sh.deleteRow(i+1);
      return { status: "ok", message: "Data absensi dihapus" };
    }
  }
  return { status: "error", message: "Data tidak ditemukan" };
}

// ==================== LOGIN ADMIN ===========================
function loginAdmin(body) {
  const sh   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_ADMIN);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === body.username && data[i][1] === body.password) {
      return { status: "ok", data: { nama: data[i][2], role: data[i][3] } };
    }
  }
  return { status: "error", message: "Username atau password salah" };
}

// ==================== UTILITIES =============================
function hitungJarak(lat1, lng1, lat2, lng2) {
  const R    = 6371000;
  const dLat = (lat2 - lat1) * Math.PI / 180;
  const dLng = (lng2 - lng1) * Math.PI / 180;
  const a = Math.sin(dLat/2) * Math.sin(dLat/2) +
            Math.cos(lat1 * Math.PI/180) * Math.cos(lat2 * Math.PI/180) *
            Math.sin(dLng/2) * Math.sin(dLng/2);
  return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1-a));
}

function getWIBDate() {
  const wib = new Date(new Date().getTime() + 7*60*60*1000);
  const p   = n => String(n).padStart(2,"0");
  return wib.getUTCFullYear() + "-" + p(wib.getUTCMonth()+1) + "-" + p(wib.getUTCDate());
}
