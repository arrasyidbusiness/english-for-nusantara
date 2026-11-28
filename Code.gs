// =====================================================================
// Code.gs — English For Nusantara (VERSI TERBARU, PENGGANTI PENUH)
// Perubahan: (1) kolom nilai lengkap & otomatis di Google Sheets,
// (2) penggabungan data di server per kolom (nilai terbaru menang, 0 tidak hilang),
// (3) kunci penulisan agar data tidak saling menimpa.
// Cara pakai: hapus semua isi Code.gs lama, tempel seluruh isi file ini, Simpan,
// lalu Deploy > Manage deployments > Edit > Version: New version > Deploy.
// =====================================================================

const ROOT_FOLDER_NAME = 'English_For_Nusantara_Database_Permanen';

function getOrCreateRootFolder() {
  const folders = DriveApp.getFoldersByName(ROOT_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(ROOT_FOLDER_NAME);
}

function getTargetSpreadsheet(rootFolder) {
  // Jika Apps Script dibuka dari Google Sheets (bound script), langsung pakai sheet tersebut
  try {
    const active = SpreadsheetApp.getActiveSpreadsheet();
    if (active) return active;
  } catch (e) {}

  // Jika standalone script, otomatis buat/buka file Google Sheets di dalam folder Drive
  const sheetName = 'Database_English_For_Nusantara';
  const files = rootFolder.getFilesByName(sheetName);
  if (files.hasNext()) {
    return SpreadsheetApp.openById(files.next().getId());
  }
  const created = SpreadsheetApp.create(sheetName);
  DriveApp.getFileById(created.getId()).moveTo(rootFolder);
  return created;
}

function getOrCreateTab(ss, tabName, headers) {
  let sheet = ss.getSheetByName(tabName);
  if (!sheet) {
    sheet = ss.insertSheet(tabName);
  }
  if (sheet.getLastRow() === 0 && headers && headers.length) {
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#102542').setFontColor('#FFFFFF');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// Urutan alami kode TP: "TP 7.2" < "TP 7.10" < "TP 8.1"
function compareTpKeys_(a, b) {
  const na = String(a).match(/\d+/g) || [];
  const nb = String(b).match(/\d+/g) || [];
  for (let i = 0; i < Math.max(na.length, nb.length); i++) {
    const x = Number(na[i] || 0), y = Number(nb[i] || 0);
    if (x !== y) return x - y;
  }
  return String(a).localeCompare(String(b));
}

function writeTable_(sheet, headers, rows) {
  sheet.clearContents();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold').setBackground('#102542').setFontColor('#FFFFFF');
  sheet.setFrozenRows(1);
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows); // angka 0 tetap ditulis sebagai 0
  }
}

function syncToGoogleSheets(rootFolder, registry, allTeacherSpaces) {
  const ss = getTargetSpreadsheet(rootFolder);

  // 1. Tab Data_Guru
  const guruHeaders = ['Email Guru', 'Nama Lengkap', 'NIP', 'Sekolah', 'Jenjang Diampu', 'Status', 'Paket', 'Berlaku Sampai', 'Terakhir Update'];
  const guruSheet = getOrCreateTab(ss, 'Data_Guru', guruHeaders);
  const teachers = registry.teachers || [];
  const guruRows = teachers.map(t => [
    t.email || '',
    t.nama || '',
    t.nip || '-',
    t.sekolah || '',
    (t.allowedGrades || ['IX']).join(', '),
    t.status || 'ACTIVE',
    t.plan || 'TRIAL',
    t.expiresAt || '',
    new Date().toLocaleString('id-ID')
  ]);
  writeTable_(guruSheet, guruHeaders, guruRows);

  // 2. Tab Data_Siswa_Nilai: kolom UH per TP dibuat OTOMATIS dari data nilai yang ada
  const siswaSheet = getOrCreateTab(ss, 'Data_Siswa_Nilai', ['Sekolah']);
  const studentRaw = [];
  const tpKeySet = {};

  // 3. Tab Data_Materi_Soal (Seluruh Materi Ajar, Buku Panduan & Bank Soal UH/STS/UAS/TKA)
  const materiSoalHeaders = [
    'Sekolah', 'Guru Pengampu', 'Jenis Konten', 'Jenjang/Kode', 'Judul / Paket Soal',
    'Jumlah Butir / Lampiran', 'Status Tampil Siswa', 'Ringkasan / Tautan', 'Update Terakhir'
  ];
  const materiSoalSheet = getOrCreateTab(ss, 'Data_Materi_Soal', materiSoalHeaders);
  const materiSoalRows = [];

  Object.keys(allTeacherSpaces).forEach(emailKey => {
    const sp = allTeacherSpaces[emailKey];
    if (!sp) return;
    const sekolah = sp.profile && sp.profile.sekolah ? sp.profile.sekolah : '';
    const guruNama = sp.profile && sp.profile.nama ? sp.profile.nama : emailKey;
    const timestampStr = new Date().toLocaleString('id-ID');

    // Kumpulkan data siswa & nilai
    if (sp.rombels) {
      Object.keys(sp.rombels).forEach(rId => {
        const r = sp.rombels[rId];
        (r.students || []).forEach(st => {
          Object.keys(st.scores || {}).forEach(k => { tpKeySet[k] = true; });
          studentRaw.push({ sekolah, guruNama, r, rId, st, timestampStr });
        });
      });
    }

    // Kumpulkan baris Materi Ajar & Buku Panduan
    (sp.uploadedResources || []).forEach(res => {
      materiSoalRows.push([
        sekolah,
        guruNama,
        res.category || 'Materi Ajar',
        'Kelas ' + (res.grade || '') + ' (' + (res.chapterCode || '-') + ')',
        res.title || '',
        res.fileName ? 'File: ' + res.fileName : 'Teks / Tautan',
        res.visibleToStudent ? 'Aktif di Siswa' : 'Tersembunyi',
        (res.externalUrl || res.content || '').slice(0, 180),
        res.uploadedAt || timestampStr
      ]);
    });

    // Kumpulkan baris Bank Soal (UH, STS, UAS, TKA, Remedial TKA)
    if (sp.exams) {
      Object.keys(sp.exams).forEach(examKey => {
        const ex = sp.exams[examKey];
        if (!ex) return;
        const isRemedial = /^TKA-.+-R$/.test(examKey);
        const jenisUjian = isRemedial ? 'Bank Soal Remedial TKA'
          : examKey.indexOf('TKA-') === 0 ? 'Bank Soal TKA'
          : examKey.indexOf('STS-') === 0 ? 'Bank Soal STS'
          : examKey.indexOf('UAS-') === 0 ? 'Bank Soal UAS/SAS'
          : 'Bank Soal UH';
        const qCount = Array.isArray(ex.questions) ? ex.questions.length : 0;
        const firstQuestion = qCount > 0 ? ex.questions[0].q : '-';

        materiSoalRows.push([
          sekolah,
          guruNama,
          jenisUjian,
          examKey,
          ex.title || examKey,
          qCount + ' Butir Soal (' + (ex.duration || 30) + ' Menit)',
          ex.published ? 'Aktif di Siswa' : 'Nonaktif',
          'Contoh Soal 1: ' + String(firstQuestion).slice(0, 140),
          timestampStr
        ]);
      });
    }
  });

  // Susun kolom Data_Siswa_Nilai (UH per TP otomatis, urut alami)
  const tpKeys = Object.keys(tpKeySet).sort(compareTpKeys_);
  const siswaHeaders = [
    'Sekolah', 'Guru Pengampu', 'Kelas/Rombel', 'Kode Kelas', 'NIS', 'Nama Siswa', 'L/P', 'PIN', 'Hadir (dari 30)'
  ].concat(tpKeys.map(k => 'UH ' + k)).concat([
    'STS 1', 'SAS 1', 'Remedial 1', 'Katrol 1',
    'STS 2', 'SAS 2', 'Remedial 2', 'Katrol 2',
    'Skor TKA 1', 'Skor TKA 2', 'Skor Remedial TKA', 'Pelanggaran Ujian', 'Update Terakhir'
  ]);

  const studentRows = studentRaw.map(o => {
    const st = o.st;
    const r = o.r;
    const grade = r.grade || '';
    const hadirCount = Array.isArray(st.att) ? st.att.filter(v => v === 'H').length : 30;
    const examRes = st.examResults || {};
    const scores = st.scores || {};
    let pelanggaran = 0;
    Object.keys(st.examViolations || {}).forEach(k => {
      pelanggaran += Number((st.examViolations[k] || {}).count || 0);
    });
    const tka = suffix => {
      const v = examRes['TKA-' + grade + '-' + suffix];
      return v === undefined || v === null ? '' : v;
    };
    return [
      o.sekolah,
      o.guruNama,
      r.label || o.rId,
      r.classCode || '',
      st.nis || '',
      st.name || '',
      st.gender || 'L',
      st.pin || '1234',
      hadirCount + '/30'
    ].concat(tpKeys.map(k => (Object.prototype.hasOwnProperty.call(scores, k) ? Number(scores[k]) : '')))
     .concat([
      Number(st.sts1 || 0), Number(st.sas1 || 0), Number(st.remedial1 || 0), Number(st.katrol1 || 0),
      Number(st.sts2 || 0), Number(st.sas2 || 0), Number(st.remedial2 || 0), Number(st.katrol2 || 0),
      tka('1'), tka('2'), tka('R'), pelanggaran, o.timestampStr
    ]);
  });

  writeTable_(siswaSheet, siswaHeaders, studentRows);
  writeTable_(materiSoalSheet, materiSoalHeaders, materiSoalRows);

  // 4. Tab Token_Lisensi
  const tokenHeaders = ['Kode Token', 'Paket', 'Masa Aktif (Hari)', 'Digunakan Oleh'];
  const tokenSheet = getOrCreateTab(ss, 'Token_Lisensi', tokenHeaders);
  const tokens = registry.tokens || [];
  const tokRows = tokens.map(tk => [tk.code || '', tk.plan || '', tk.days || 30, tk.usedBy || 'Tersedia']);
  writeTable_(tokenSheet, tokenHeaders, tokRows);
}

// Tab "Nilai_<nama guru>": rekap lengkap termasuk Nilai Akhir & Status (dihitung oleh aplikasi guru)
function tulisRekapNilai_(rootFolder, tbl) {
  if (!tbl || !tbl.headers || !tbl.rows) return;
  const ss = getTargetSpreadsheet(rootFolder);
  const nama = String(tbl.sheetName || 'Nilai_Guru').replace(/[\[\]\*\?:\/\\]/g, ' ').substring(0, 90);
  let sheet = ss.getSheetByName(nama);
  if (!sheet) sheet = ss.insertSheet(nama);
  writeTable_(sheet, tbl.headers, tbl.rows);
}

// =====================================================================
// PENGGABUNGAN DATA DI SERVER (nilai terbaru per kolom menang)
// Setiap perubahan nilai di aplikasi diberi cap waktu per kolom (_ts).
// =====================================================================
const TRACKED_SCALARS_ = ['sts1', 'sas1', 'sts2', 'sas2', 'remedial1', 'remedial2', 'katrol1', 'katrol2', 'wa'];
const TRACKED_MAPS_ = ['scores', 'examResults', 'examViolations', 'lkpdSubmissions', 'interaktif'];

function normName_(s) {
  return String(s || '').trim().toLowerCase().replace(/\s+/g, ' ');
}

// studentSender=true: pengirim adalah siswa -> hanya kolom milik siswa yang boleh diambil;
// data lain (kehadiran, PIN, dll.) tetap dari data tersimpan (milik guru).
function mergeStudent_(stored, incoming, studentSender) {
  const sts = stored._ts || {};
  const its = incoming._ts || {};
  const out = studentSender ? Object.assign({}, stored) : Object.assign({}, stored, incoming);

  TRACKED_SCALARS_.forEach(k => {
    const s = sts[k] || 0, i = its[k] || 0;
    if (i > s) out[k] = incoming[k];
    else if (s > i) out[k] = stored[k];
    else if (studentSender) out[k] = stored[k] !== undefined ? stored[k] : incoming[k];
    else out[k] = incoming[k] !== undefined ? incoming[k] : stored[k];
  });

  TRACKED_MAPS_.forEach(mk => {
    const sm = stored[mk] || {};
    const im = incoming[mk] || {};
    const res = {};
    const keys = {};
    Object.keys(sm).forEach(k => { keys[k] = true; });
    Object.keys(im).forEach(k => { keys[k] = true; });
    Object.keys(keys).forEach(k => {
      const path = mk + '.' + k;
      const inS = Object.prototype.hasOwnProperty.call(sm, k);
      const inI = Object.prototype.hasOwnProperty.call(im, k);
      if (!inI) { res[k] = sm[k]; return; }
      if (!inS) { res[k] = im[k]; return; }
      const s = sts[path] || 0, i = its[path] || 0;
      if (i > s) res[k] = im[k];
      else if (s > i) res[k] = sm[k];
      else res[k] = studentSender ? sm[k] : im[k];
    });
    out[mk] = res;
  });

  if (Array.isArray(incoming.completedChapters) || Array.isArray(stored.completedChapters)) {
    const set = {};
    (stored.completedChapters || []).concat(incoming.completedChapters || []).forEach(c => { set[c] = true; });
    out.completedChapters = Object.keys(set);
  }

  const ts = {};
  Object.keys(sts).forEach(k => { ts[k] = sts[k]; });
  Object.keys(its).forEach(k => { ts[k] = Math.max(ts[k] || 0, its[k] || 0); });
  out._ts = ts;
  return out;
}

function mergeTeacherSpace_(stored, incoming, senderRole, studentName) {
  if (!stored) return incoming;
  const studentSender = senderRole === 'student';
  const result = JSON.parse(JSON.stringify(studentSender ? stored : incoming));
  if (!result.rombels) result.rombels = {};
  const sRombels = stored.rombels || {};
  const iRombels = incoming.rombels || {};
  const ids = {};
  Object.keys(sRombels).forEach(k => { ids[k] = true; });
  Object.keys(iRombels).forEach(k => { ids[k] = true; });
  const me = normName_(studentName);

  Object.keys(ids).forEach(rId => {
    const sr = sRombels[rId];
    const ir = iRombels[rId];
    if (!sr && ir) {
      if (!studentSender) result.rombels[rId] = ir;
      return;
    }
    if (sr && !ir) {
      result.rombels[rId] = sr;
      return;
    }
    const storedList = sr.students || [];
    const incomingList = ir.students || [];
    const merged = [];
    const used = {};

    if (studentSender) {
      storedList.forEach(p => {
        const key = normName_(p.name);
        used[key] = true;
        const inc = incomingList.filter(x => normName_(x.name) === key)[0];
        if (inc && (!me || me === key)) merged.push(mergeStudent_(p, inc, true));
        else merged.push(p);
      });
      // Siswa baru yang mendaftar mandiri (belum ada di data tersimpan)
      incomingList.forEach(x => {
        const key = normName_(x.name);
        if (!used[key] && (!me || me === key)) merged.push(x);
      });
    } else {
      incomingList.forEach(x => {
        const key = normName_(x.name);
        used[key] = true;
        const old = storedList.filter(p => normName_(p.name) === key)[0];
        merged.push(old ? mergeStudent_(old, x, false) : x);
      });
      // Siswa yang hanya ada di data tersimpan (mis. mendaftar mandiri dari perangkat lain)
      storedList.forEach(p => {
        if (!used[normName_(p.name)]) merged.push(p);
      });
    }
    result.rombels[rId] = Object.assign({}, studentSender ? sr : ir, { students: merged });
  });
  return result;
}


function readJsonFile(folder, fileName, defaultData) {
  const files = folder.getFilesByName(fileName);
  if (!files.hasNext()) return defaultData;
  try {
    const content = files.next().getBlob().getDataAsString('UTF-8');
    return JSON.parse(content);
  } catch (e) {
    return defaultData;
  }
}

function writeJsonFile(folder, fileName, dataObj) {
  const jsonStr = JSON.stringify(dataObj, null, 2);
  const files = folder.getFilesByName(fileName);
  if (files.hasNext()) {
    const file = files.next();
    file.setContent(jsonStr);
    return file.getUrl();
  } else {
    const created = folder.createFile(fileName, jsonStr, MimeType.PLAIN_TEXT);
    return created.getUrl();
  }
}

function loadAllTeacherSpacesMap(rootFolder) {
  const map = {};
  const files = rootFolder.getFiles();
  while (files.hasNext()) {
    const f = files.next();
    const name = f.getName();
    if (name.indexOf('teacher_') === 0 && name.slice(-5) === '.json') {
      try {
        const parsed = JSON.parse(f.getBlob().getDataAsString('UTF-8'));
        if (parsed && parsed.teacherKey && parsed.teacherSpace) {
          map[parsed.teacherKey] = parsed.teacherSpace;
        }
      } catch (e) {}
    }
  }
  return map;
}

// =====================================================================
// NOTIFIKASI WHATSAPP (opsional) — dikirim dari server lewat gateway WhatsApp
// Token disimpan di Script Properties (BUKAN di kode). Cara menyimpan token:
//   1) Isi token di fungsi simpanTokenWA() di bawah,
//   2) pilih fungsi simpanTokenWA di menu fungsi lalu tekan Jalankan (sekali saja),
//   3) setelah berhasil, hapus token dari kode.
// Gateway bawaan: Fonnte (api.fonnte.com). Tanpa token, fitur ini otomatis tidak aktif.
// =====================================================================
function simpanTokenWA() {
  PropertiesService.getScriptProperties().setProperty('WA_TOKEN', 'ISI_TOKEN_FONNTE_DI_SINI');
}

function waNorm_(n) {
  let d = String(n || '').replace(/[^0-9]/g, '');
  if (!d) return '';
  if (d.indexOf('62') === 0) return d;
  if (d.indexOf('0') === 0) return '62' + d.slice(1);
  if (d.indexOf('8') === 0) return '62' + d;
  return d;
}

// Daftar nomor valid dari teks "0812..., 0813...; 0814..." (koma/titik koma/baris baru), tanpa duplikat.
function waList_(v) {
  const list = String(v || '').split(/[,;\n]+/).map(waNorm_).filter(function (x) { return x.length >= 9 && x.length <= 15; });
  return list.filter(function (x, i) { return list.indexOf(x) === i; });
}

// Mengirim satu pesan ke satu atau beberapa nomor (dipisah koma). Mengembalikan {ok, detail}.
function waKirim_(targets, message) {
  const token = PropertiesService.getScriptProperties().getProperty('WA_TOKEN');
  if (!token || token === 'ISI_TOKEN_FONNTE_DI_SINI') return { ok: false, detail: 'Token WA belum diatur' };
  const uniq = waList_(Array.isArray(targets) ? targets.join(',') : targets);
  if (!uniq.length) return { ok: false, detail: 'Nomor tujuan tidak valid' };
  const res = UrlFetchApp.fetch('https://api.fonnte.com/send', {
    method: 'post',
    headers: { Authorization: token },
    payload: { target: uniq.join(','), message: message, countryCode: '62' },
    muteHttpExceptions: true
  });
  return waHasil_(res);
}

// Hasil panggilan Fonnte: sukses hanya bila HTTP 2xx dan status tidak false
function waHasil_(res) {
  const code = res.getResponseCode();
  const text = String(res.getContentText());
  let j = null;
  try { j = JSON.parse(text); } catch (e) {}
  const ok = code >= 200 && code < 300 && !(j && j.status === false);
  const detail = (j && (j.reason || j.detail)) ? String(j.reason || j.detail) : text.slice(0, 200);
  return { ok: ok, detail: detail, json: j };
}

// Kunci keamanan opsional: bila WA_API_KEY diisi di Script properties, aksi WA wajib menyertakan kunci yang sama
function waKunciValid_(body) {
  const k = PropertiesService.getScriptProperties().getProperty('WA_API_KEY');
  return !k || String(body.apiKey || '') === k;
}

// Batas jenjang untuk edit materi interaktif: guru hanya boleh mengubah jenjang yang ia ampu.
// Perubahan jenjang lain dari klien diabaikan; isi jenjang lain yang sudah tersimpan tetap dipertahankan.
const ROMAWI_KE_ANGKA_ = { VII: '7', VIII: '8', IX: '9' };

function batasiEditMateri_(stored, incoming, allowedRoman) {
  const boleh = (allowedRoman || []).map(function (x) { return ROMAWI_KE_ANGKA_[x]; }).filter(Boolean);
  const cek = function (k) { return boleh.indexOf(String(k).split('|')[0]) >= 0; };
  const out = { chapters: {}, meetings: {} };
  const sCh = (stored && stored.chapters) || {}, sMe = (stored && stored.meetings) || {};
  const iCh = (incoming && incoming.chapters) || {}, iMe = (incoming && incoming.meetings) || {};
  Object.keys(sCh).forEach(function (k) { if (!cek(k)) out.chapters[k] = sCh[k]; });
  Object.keys(sMe).forEach(function (k) { if (!cek(k)) out.meetings[k] = sMe[k]; });
  Object.keys(iCh).forEach(function (k) { if (cek(k)) out.chapters[k] = iCh[k]; });
  Object.keys(iMe).forEach(function (k) { if (cek(k)) out.meetings[k] = iMe[k]; });
  return out;
}

// Kirim ke grup WhatsApp (target = ID grup berakhiran @g.us)
function waKirimGrup_(groupId, message) {
  const token = PropertiesService.getScriptProperties().getProperty('WA_TOKEN');
  if (!token || token === 'ISI_TOKEN_FONNTE_DI_SINI') return { ok: false, detail: 'Token WA belum diatur' };
  const id = String(groupId || '').trim();
  if (!/^[0-9A-Za-z._-]+@g\.us$/.test(id)) return { ok: false, detail: 'ID grup tidak valid (harus berakhiran @g.us)' };
  const res = UrlFetchApp.fetch('https://api.fonnte.com/send', {
    method: 'post',
    headers: { Authorization: token },
    payload: { target: id, message: String(message || '').slice(0, 4000) },
    muteHttpExceptions: true
  });
  return waHasil_(res);
}

// Daftar grup yang diikuti nomor Fonnte. refresh=true meminta Fonnte mengambil ulang dari WhatsApp (cukup sekali).
function waDaftarGrup_(refresh) {
  const token = PropertiesService.getScriptProperties().getProperty('WA_TOKEN');
  if (!token || token === 'ISI_TOKEN_FONNTE_DI_SINI') return { ok: false, detail: 'Token WA belum diatur', groups: [] };
  const opt = { method: 'post', headers: { Authorization: token }, muteHttpExceptions: true };
  if (refresh) UrlFetchApp.fetch('https://api.fonnte.com/fetch-group', opt);
  const r = waHasil_(UrlFetchApp.fetch('https://api.fonnte.com/get-whatsapp-group', opt));
  const data = (r.json && Array.isArray(r.json.data)) ? r.json.data : [];
  const groups = data.map(function (g) { return { id: g.id, name: g.name }; })
    .filter(function (g) { return g.id && /@g\.us$/.test(String(g.id)); });
  return { ok: true, detail: r.detail, groups: groups };
}

const MAKS_PELANGGARAN_ = 3;

function labelUjian_(k) {
  return /^TKA-.+-R$/.test(k) ? 'Remedial TKA' : k;
}

function gradeDariKunciUjian_(k) {
  let m = String(k).match(/-(VII|VIII|IX)-/);
  if (m) return m[1];
  m = String(k).match(/TP (\d)/);
  if (m) return m[1] === '7' ? 'VII' : m[1] === '8' ? 'VIII' : 'IX';
  return '';
}

function petaSiswa_(sp) {
  const map = {};
  const rombels = (sp && sp.rombels) || {};
  Object.keys(rombels).forEach(function (rId) {
    const r = rombels[rId];
    (r.students || []).forEach(function (st) {
      map[rId + '|' + normName_(st.name)] = { rId: rId, label: r.label || rId, grade: r.grade || '', st: st };
    });
  });
  return map;
}

// Membandingkan data tersimpan (lama) dengan hasil gabungan (baru) untuk mencari kejadian baru.
function deteksiEventWa_(oldSp, newSp, senderRole) {
  const res = { guru: [], siswa: {}, siaran: [], siswaSendiri: {} };
  if (!oldSp || !newSp) return res;
  const oldM = petaSiswa_(oldSp);
  const newM = petaSiswa_(newSp);

  if (senderRole === 'student') {
    Object.keys(newM).forEach(function (key) {
      const n = newM[key], o = oldM[key], st = n.st;
      const nama = st.name + ' (' + n.label + ')';
      if (!o) { res.guru.push('👤 Siswa baru bergabung: ' + nama); return; }
      const sendiri = [];
      // Nomor WA baru didaftarkan siswa -> pesan penyambutan sekaligus uji nomor
      if (st.wa && !o.st.wa) sendiri.push('📱 Nomor WhatsApp Anda sudah terdaftar. Kabar ujian, LKPD, dan nilai akan dikirim ke nomor ini.');
      const oe = o.st.examResults || {};
      Object.keys(st.examResults || {}).forEach(function (k) {
        if (oe[k] !== st.examResults[k]) {
          res.guru.push('✅ ' + nama + ' selesai mengerjakan ' + labelUjian_(k) + ' — skor ' + st.examResults[k]);
          sendiri.push('✅ Anda telah menyelesaikan ' + labelUjian_(k) + '. Skor Anda: ' + st.examResults[k]);
        }
      });
      const ol = o.st.lkpdSubmissions || {};
      Object.keys(st.lkpdSubmissions || {}).forEach(function (tp) {
        if (!ol[tp]) {
          res.guru.push('📝 ' + nama + ' mengirim jawaban LKPD ' + tp);
          sendiri.push('📝 Jawaban LKPD ' + tp + ' Anda sudah terkirim ke guru.');
        }
      });
      const ov = o.st.examViolations || {};
      Object.keys(st.examViolations || {}).forEach(function (k) {
        const c = (st.examViolations[k] || {}).count || 0;
        const oc = (ov[k] || {}).count || 0;
        if (c > oc) {
          res.guru.push('⚠️ ' + nama + ' melanggar aturan ujian ' + labelUjian_(k) + ' (total ' + c + 'x)');
          sendiri.push('⚠️ Pelanggaran aturan ujian ' + labelUjian_(k) + ' tercatat (total ' + c + 'x): berpindah tab/aplikasi atau memakai penerjemah. Setelah ' + MAKS_PELANGGARAN_ + 'x ujian dikumpulkan otomatis.');
        }
      });
      if (sendiri.length) res.siswaSendiri[key] = { nama: st.name, wa: st.wa, lines: sendiri };
    });
    return res;
  }

  // Pengirim guru -> kabar untuk siswa
  Object.keys(newM).forEach(function (key) {
    const o = oldM[key];
    if (!o) return;
    const st = newM[key].st, os = o.st;
    const lines = [];
    const osub = os.lkpdSubmissions || {};
    Object.keys(st.lkpdSubmissions || {}).forEach(function (tp) {
      const sc = (st.lkpdSubmissions[tp] || {}).score;
      const osc = (osub[tp] || {}).score;
      if (sc !== null && sc !== undefined && sc !== osc) lines.push('🎯 LKPD ' + tp + ' sudah dinilai: ' + sc);
    });
    [1, 2].forEach(function (i) {
      if (Number(st['remedial' + i] || 0) > 0 && st['remedial' + i] !== os['remedial' + i]) lines.push('🔁 Nilai remedial Semester ' + i + ': ' + st['remedial' + i]);
      if (Number(st['katrol' + i] || 0) > 0 && st['katrol' + i] !== os['katrol' + i]) lines.push('⭐ Katrol +' + st['katrol' + i] + ' (Semester ' + i + ')');
    });
    if (lines.length) res.siswa[key] = { nama: st.name, wa: st.wa, lines: lines };
  });

  const oldEx = oldSp.exams || {};
  const newEx = newSp.exams || {};
  Object.keys(newEx).forEach(function (k) {
    if (newEx[k] && newEx[k].published && !(oldEx[k] && oldEx[k].published)) {
      res.siaran.push({ grade: gradeDariKunciUjian_(k), text: '📝 Ujian baru tersedia: ' + (newEx[k].title || k), guruText: '📢 Anda mengaktifkan ujian: ' + (newEx[k].title || k) });
    }
  });
  const oldRes = oldSp.uploadedResources || [];
  (newSp.uploadedResources || []).forEach(function (x) {
    if (!x.visibleToStudent) return;
    const idn = x.id || x.title;
    const had = oldRes.some(function (y) { return (y.id || y.title) === idn && y.visibleToStudent; });
    if (!had) res.siaran.push({ grade: x.grade || '', text: '📚 Materi baru dari guru: ' + (x.title || 'Materi'), guruText: '📚 Anda menambahkan materi: ' + (x.title || 'Materi') });
  });
  return res;
}

function kirimNotifikasiWa_(oldSp, newSp, senderRole) {
  if (!newSp || !newSp.profile) return;
  const prof = newSp.profile;
  const ev = deteksiEventWa_(oldSp, newSp, senderRole);
  const footer = '\n\n— English For Nusantara';
  const guruAktif = prof.waNotifGuru !== false && waList_(prof.waGuru).length > 0;
  const siswaAktif = prof.waNotifSiswa !== false;
  let kirim = 0;

  // 1) Ke guru: semua aktivitas siswa
  if (ev.guru.length && guruAktif) {
    const baris = ev.guru.slice(0, 15).join('\n');
    const lebih = ev.guru.length > 15 ? '\n…dan ' + (ev.guru.length - 15) + ' kejadian lainnya' : '';
    waKirim_(prof.waGuru, '📣 Pembaruan dari siswa\n' + baris + lebih + footer);
    kirim++;
  }

  let dikabari = 0;
  if (siswaAktif) {
    // 2) Ke siswa itu sendiri (selesai ujian + skor, LKPD terkirim, pelanggaran, nomor terdaftar)
    Object.keys(ev.siswaSendiri).forEach(function (key) {
      if (kirim >= 40) return;
      const o = ev.siswaSendiri[key];
      if (!waList_(o.wa).length) return;
      waKirim_(o.wa, 'Halo ' + o.nama + ',\n' + o.lines.join('\n') + footer);
      kirim++;
    });

    // 3) Ke siswa tertentu (nilai/LKPD/remedial/katrol dari guru)
    Object.keys(ev.siswa).forEach(function (key) {
      if (kirim >= 40) return;
      const o = ev.siswa[key];
      if (!waList_(o.wa).length) return;
      waKirim_(o.wa, 'Halo ' + o.nama + ',\n' + o.lines.join('\n') + footer);
      kirim++;
      dikabari++;
    });

    // 4) Siaran ke seluruh siswa satu jenjang (ujian/materi baru dari guru)
    const grades = {};
    ev.siaran.forEach(function (x) {
      const g = x.grade || '*';
      (grades[g] = grades[g] || []).push(x.text);
    });
    Object.keys(grades).forEach(function (g) {
      const nomor = [];
      Object.keys(newSp.rombels || {}).forEach(function (rId) {
        const r = newSp.rombels[rId];
        if (g !== '*' && r.grade !== g) return;
        (r.students || []).forEach(function (st) { waList_(st.wa).forEach(function (n) { nomor.push(n); }); });
      });
      if (nomor.length) {
        waKirim_(nomor, 'Halo siswa' + (g !== '*' ? ' kelas ' + g : '') + ',\n' + grades[g].join('\n') + '\nSilakan buka portal English For Nusantara.' + footer);
        dikabari += nomor.length;
      }
    });
  }

  // 5) Ke guru: ringkasan tindakan guru sendiri (ujian/materi diaktifkan, nilai diperbarui) + jumlah siswa yang dikabari
  if (senderRole !== 'student' && guruAktif) {
    const baris = [];
    ev.siaran.forEach(function (x) { if (x.guruText) baris.push(x.guruText); });
    const nSiswaNilai = Object.keys(ev.siswa).length;
    if (nSiswaNilai) baris.push('🎯 Anda memperbarui nilai/LKPD/remedial/katrol untuk ' + nSiswaNilai + ' siswa');
    if (baris.length) {
      if (siswaAktif) baris.push('✉️ Kabar diteruskan ke ' + dikabari + ' nomor siswa');
      waKirim_(prof.waGuru, '🗂️ Tindakan Anda tercatat\n' + baris.slice(0, 15).join('\n') + footer);
    }
  }
}

// Data untuk siswa tidak boleh memuat nomor WA siapa pun
function bersihkanUntukSiswa_(space) {
  const copy = JSON.parse(JSON.stringify(space));
  if (copy.profile) delete copy.profile.waGuru;
  Object.keys(copy.rombels || {}).forEach(function (rId) {
    delete copy.rombels[rId].waGroup;
    (copy.rombels[rId].students || []).forEach(function (st) { delete st.wa; });
  });
  return copy;
}

// =====================================================================
// LISENSI, PAKET, PESANAN & IZIN EDIT MATERI
// Harga & rekening bisa diubah admin (menu Pesanan). Aksi admin dijaga ADMIN_KEY di Script properties.
// Pembayaran dikonfirmasi manual oleh admin (cocokkan transfer), lalu kode token diberikan ke guru.
// =====================================================================
const PAKET_DEFAULT_ = {
  'PRO-1-BULAN': { label: 'Paket 1 Bulan', days: 30, harga: 25000 },
  'PRO-1-TAHUN': { label: 'Paket 1 Tahun', days: 365, harga: 250000 },
  'PRO-2-TAHUN': { label: 'Paket 2 Tahun', days: 730, harga: 450000 }
};

function adminOk_(body) {
  const k = PropertiesService.getScriptProperties().getProperty('ADMIN_KEY');
  return !!k && String(body.adminKey || '') === k;
}

function getRegistry_(root) {
  const r = readJsonFile(root, 'master_registry.json', { teachers: [], tokens: [], classIndex: {} });
  if (!r.teachers) r.teachers = [];
  if (!r.tokens) r.tokens = [];
  if (!r.classIndex) r.classIndex = {};
  if (!r.orders) r.orders = [];
  if (!r.prices) r.prices = {};
  return r;
}

function hargaPaket_(reg) {
  const out = {};
  Object.keys(PAKET_DEFAULT_).forEach(function (k) {
    const d = PAKET_DEFAULT_[k];
    out[k] = { label: d.label, days: d.days, harga: (reg.prices[k] && reg.prices[k].harga) || d.harga };
  });
  return out;
}

function kodeAcak_(prefix) {
  return prefix + '-' + Utilities.getUuid().replace(/-/g, '').substring(0, 8).toUpperCase();
}

function tanggalISO_(d) {
  return Utilities.formatDate(d, 'Asia/Jakarta', 'yyyy-MM-dd');
}

// Guru tidak boleh mengubah lisensi dari perangkatnya: nilai lisensi dari server dipertahankan
function jagaLisensi_(incoming, saved) {
  if (!saved || !saved.teacherRecord || !incoming) return incoming;
  const out = Object.assign({}, incoming);
  ['plan', 'expiresAt', 'status'].forEach(function (k) {
    if (saved.teacherRecord[k] !== undefined) out[k] = saved.teacherRecord[k];
  });
  return out;
}

function lisensiAksi_(action, body, root) {
  const reg = getRegistry_(root);

  if (action === 'getPrices') {
    return { status: 'ok', paket: hargaPaket_(reg), paymentInfo: reg.paymentInfo || '' };
  }

  if (action === 'createOrder') {
    const email = String(body.email || '').trim().toLowerCase();
    const paket = hargaPaket_(reg)[String(body.plan || '')];
    if (!email || !paket) return { status: 'error', message: 'Paket atau email tidak valid' };
    const order = { code: kodeAcak_('ORD'), email: email, plan: body.plan, days: paket.days, amount: paket.harga, status: 'MENUNGGU', tokenCode: '', createdAt: new Date().toISOString() };
    reg.orders.push(order);
    writeJsonFile(root, 'master_registry.json', reg);
    return { status: 'ok', order: order, paymentInfo: reg.paymentInfo || '' };
  }

  if (action === 'getOrder') {
    const email = String(body.email || '').trim().toLowerCase();
    const code = String(body.code || '').trim().toUpperCase();
    const order = reg.orders.filter(function (o) { return o.code === code && o.email === email; })[0];
    if (!order) return { status: 'error', message: 'Pesanan tidak ditemukan' };
    return { status: 'ok', order: { code: order.code, plan: order.plan, amount: order.amount, status: order.status, tokenCode: order.status === 'LUNAS' ? order.tokenCode : '' } };
  }

  if (action === 'confirmOrder') {
    if (!adminOk_(body)) return { status: 'error', message: 'Kunci admin salah' };
    const code = String(body.code || '').trim().toUpperCase();
    const order = reg.orders.filter(function (o) { return o.code === code; })[0];
    if (!order) return { status: 'error', message: 'Pesanan tidak ditemukan' };
    if (order.status !== 'MENUNGGU') return { status: 'error', message: 'Pesanan sudah diproses' };
    const token = { code: kodeAcak_('EFN'), plan: order.plan, days: order.days, usedBy: '', orderCode: order.code };
    reg.tokens.push(token);
    order.status = 'LUNAS';
    order.tokenCode = token.code;
    writeJsonFile(root, 'master_registry.json', reg);
    return { status: 'ok', order: order };
  }

  if (action === 'redeemToken') {
    const email = String(body.email || '').trim().toLowerCase();
    const code = String(body.code || '').trim().toUpperCase();
    const token = reg.tokens.filter(function (t) { return String(t.code).toUpperCase() === code && !t.usedBy; })[0];
    if (!token) return { status: 'error', message: 'Kode tidak valid atau sudah dipakai' };
    const file = 'teacher_' + email.replace(/[^a-z0-9]/g, '_') + '.json';
    const saved = readJsonFile(root, file, null);
    let teacher = reg.teachers.filter(function (t) { return String(t.email).toLowerCase() === email; })[0];
    if (!teacher && saved && saved.teacherRecord) teacher = Object.assign({}, saved.teacherRecord);
    if (!teacher) return { status: 'error', message: 'Akun guru tidak ditemukan. Sinkronkan data dari perangkat guru dahulu.' };
    const today = new Date();
    const dasar = teacher.expiresAt && teacher.expiresAt > tanggalISO_(today) ? new Date(teacher.expiresAt + 'T00:00:00+07:00') : today;
    dasar.setDate(dasar.getDate() + Number(token.days || 30));
    teacher.expiresAt = tanggalISO_(dasar);
    teacher.plan = token.plan;
    teacher.status = 'ACTIVE';
    token.usedBy = email;
    const order = token.orderCode ? reg.orders.filter(function (o) { return o.code === token.orderCode; })[0] : null;
    if (order) order.status = 'SELESAI';
    const idx = reg.teachers.findIndex(function (t) { return String(t.email).toLowerCase() === email; });
    if (idx >= 0) reg.teachers[idx] = teacher; else reg.teachers.push(teacher);
    writeJsonFile(root, 'master_registry.json', reg);
    if (saved) { saved.teacherRecord = teacher; writeJsonFile(root, file, saved); }
    return { status: 'ok', teacherRecord: teacher };
  }

  if (action === 'setPrices') {
    if (!adminOk_(body)) return { status: 'error', message: 'Kunci admin salah' };
    const prices = body.prices || {};
    Object.keys(PAKET_DEFAULT_).forEach(function (k) {
      if (prices[k] && Number(prices[k].harga) >= 0) reg.prices[k] = { harga: Number(prices[k].harga) };
    });
    if (body.paymentInfo !== undefined) reg.paymentInfo = String(body.paymentInfo).slice(0, 500);
    writeJsonFile(root, 'master_registry.json', reg);
    return { status: 'ok', paket: hargaPaket_(reg), paymentInfo: reg.paymentInfo || '' };
  }

  if (action === 'listOrders') {
    if (!adminOk_(body)) return { status: 'error', message: 'Kunci admin salah' };
    return { status: 'ok', orders: reg.orders, paket: hargaPaket_(reg), paymentInfo: reg.paymentInfo || '' };
  }

  if (action === 'setPermissions') {
    if (!adminOk_(body)) return { status: 'error', message: 'Kunci admin salah' };
    const email = String(body.email || '').trim().toLowerCase();
    const file = 'teacher_' + email.replace(/[^a-z0-9]/g, '_') + '.json';
    const saved = readJsonFile(root, file, null);
    if (!saved) return { status: 'error', message: 'Guru belum pernah sinkron ke cloud' };
    const grades = (body.editGrades || []).filter(function (g) { return ['VII', 'VIII', 'IX'].indexOf(g) >= 0; });
    saved.permissions = { editGrades: grades };
    writeJsonFile(root, file, saved);
    return { status: 'ok', editGrades: grades };
  }

  if (action === 'adminGetSpace') {
    if (!adminOk_(body)) return { status: 'error', message: 'Kunci admin salah' };
    const email = String(body.email || '').trim().toLowerCase();
    const saved = readJsonFile(root, 'teacher_' + email.replace(/[^a-z0-9]/g, '_') + '.json', null);
    if (!saved) return { status: 'error', message: 'Guru belum pernah sinkron ke cloud' };
    return { status: 'ok', interaktifEdits: (saved.teacherSpace && saved.teacherSpace.interaktifEdits) || {}, editGrades: (saved.permissions && saved.permissions.editGrades) || [] };
  }

  if (action === 'adminSaveEdits') {
    if (!adminOk_(body)) return { status: 'error', message: 'Kunci admin salah' };
    const email = String(body.email || '').trim().toLowerCase();
    const file = 'teacher_' + email.replace(/[^a-z0-9]/g, '_') + '.json';
    const saved = readJsonFile(root, file, null);
    if (!saved || !saved.teacherSpace) return { status: 'error', message: 'Guru belum pernah sinkron ke cloud' };
    saved.teacherSpace.interaktifEdits = body.edits || {};
    writeJsonFile(root, file, saved);
    return { status: 'ok' };
  }

  return null;
}

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', service: 'English For Nusantara Sheets & Drive Active' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  let lock = null;
  try {
    const body = JSON.parse(e.postData.contents || '{}');
    const action = body.action || '';
    const root = getOrCreateRootFolder();

    // Aksi yang menulis data dikunci agar beberapa pengirim tidak saling menimpa
    if (['syncTeacherDriveBackup', 'adminSync', 'createOrder', 'confirmOrder', 'redeemToken', 'setPrices', 'setPermissions', 'adminSaveEdits'].indexOf(action) >= 0) {
      lock = LockService.getScriptLock();
      lock.waitLock(30000);
    }

    const lisensi = lisensiAksi_(action, body, root);
    if (lisensi) return jsonResponse(lisensi);

    // 1. Simpan Backup Ruang Guru (digabung dengan data tersimpan) + Tulis Otomatis ke Google Sheets
    if (action === 'syncTeacherDriveBackup') {
      const email = String(body.email || '').trim().toLowerCase();
      if (!email || !body.teacherSpace) {
        return jsonResponse({ status: 'error', message: 'Email atau data kosong' });
      }
      const safeFileName = 'teacher_' + email.replace(/[^a-z0-9]/g, '_') + '.json';
      const senderRole = body.senderRole || 'teacher';
      const saved = readJsonFile(root, safeFileName, null);

      const mergedSpace = mergeTeacherSpace_(
        saved && saved.teacherSpace ? saved.teacherSpace : null,
        body.teacherSpace,
        senderRole,
        body.studentName
      );
      if (senderRole !== 'student') {
        const jenjangGuru = (saved && saved.permissions && saved.permissions.editGrades) || [];
        const sEdits = saved && saved.teacherSpace ? saved.teacherSpace.interaktifEdits : null;
        mergedSpace.interaktifEdits = batasiEditMateri_(sEdits, body.teacherSpace.interaktifEdits, jenjangGuru);
      }
      // Data lisensi guru hanya boleh diubah oleh guru, bukan oleh siswa
      const teacherRecord = senderRole === 'student'
        ? ((saved && saved.teacherRecord) || body.teacherRecord || null)
        : (body.teacherRecord || (saved && saved.teacherRecord) || null);

      const lisensiSimpan = senderRole === 'student' ? teacherRecord : jagaLisensi_(teacherRecord, saved);
      const fileUrl = writeJsonFile(root, safeFileName, {
        teacherKey: email,
        teacherRecord: lisensiSimpan,
        teacherSpace: mergedSpace,
        permissions: saved && saved.permissions ? saved.permissions : undefined,
        updatedAt: new Date().toISOString()
      });

      const registry = readJsonFile(root, 'master_registry.json', { teachers: [], tokens: [], classIndex: {} });
      if (!registry.teachers) registry.teachers = [];
      if (!registry.classIndex) registry.classIndex = {};
      if (teacherRecord && senderRole !== 'student') {
        // Lisensi (paket, masa berlaku, status) dijaga server: salinan lokal tidak boleh memundurkannya
        const rekamRegistry = jagaLisensi_(teacherRecord, { teacherRecord: registry.teachers.filter(t => String(t.email).toLowerCase() === email)[0] || null });
        const idx = registry.teachers.findIndex(t => String(t.email).toLowerCase() === email);
        if (idx >= 0) registry.teachers[idx] = rekamRegistry;
        else registry.teachers.push(rekamRegistry);
      }

      const rombels = mergedSpace.rombels || {};
      Object.keys(rombels).forEach(rId => {
        const code = String(rombels[rId].classCode || '').trim().toUpperCase();
        if (code) registry.classIndex[code] = email;
      });
      writeJsonFile(root, 'master_registry.json', registry);

      // Tulis ke Google Sheets (semua tab) dari data yang sudah digabung
      const allSpaces = loadAllTeacherSpacesMap(root);
      allSpaces[email] = mergedSpace;
      syncToGoogleSheets(root, registry, allSpaces);

      // Rekap lengkap per guru (termasuk Nilai Akhir) dari aplikasi guru
      if (senderRole !== 'student' && body.gradeTable) {
        tulisRekapNilai_(root, body.gradeTable);
      }

      // Notifikasi WhatsApp (jika diaktifkan guru & token tersedia); kegagalan tidak mengganggu sinkron
      try {
        kirimNotifikasiWa_(saved && saved.teacherSpace ? saved.teacherSpace : null, mergedSpace, senderRole);
      } catch (waErr) {
        console.error('Notifikasi WA gagal: ' + waErr);
      }

      return jsonResponse({ status: 'ok', fileUrl: fileUrl });
    }

    // 2. Muat Data Guru saat Login dari Perangkat Baru / Setelah Update GitHub
    if (action === 'loadTeacherDriveBackup') {
      const email = String(body.email || '').trim().toLowerCase();
      const safeFileName = 'teacher_' + email.replace(/[^a-z0-9]/g, '_') + '.json';
      const saved = readJsonFile(root, safeFileName, null);
      if (!saved) return jsonResponse({ status: 'not_found' });
      return jsonResponse({
        status: 'ok',
        teacherRecord: saved.teacherRecord,
        teacherSpace: saved.teacherSpace,
        permissions: saved.permissions || { editGrades: [] }
      });
    }

    // 3. Muat Kelas Berdasarkan Kode Kelas untuk Login Siswa Lintas Perangkat
    if (action === 'loadClassByCode') {
      const code = String(body.classCode || '').trim().toUpperCase();
      const registry = readJsonFile(root, 'master_registry.json', { teachers: [], tokens: [], classIndex: {} });
      const ownerEmail = registry.classIndex[code];
      if (!ownerEmail) return jsonResponse({ status: 'not_found' });

      const safeFileName = 'teacher_' + ownerEmail.replace(/[^a-z0-9]/g, '_') + '.json';
      const saved = readJsonFile(root, safeFileName, null);
      if (!saved) return jsonResponse({ status: 'not_found' });

      return jsonResponse({
        status: 'ok',
        teacherKey: ownerEmail,
        teacherSpace: bersihkanUntukSiswa_(saved.teacherSpace)
      });
    }

    // 4. Sinkronisasi Panel Admin Pusat + Perbarui Google Sheets
    if (action === 'adminSync') {
      const registry = readJsonFile(root, 'master_registry.json', { teachers: [], tokens: [], classIndex: {} });
      const mergedTeachers = {};
      (registry.teachers || []).forEach(t => { mergedTeachers[String(t.email).toLowerCase()] = t; });
      (body.teachers || []).forEach(t => {
        const k = String(t.email).toLowerCase();
        if (!mergedTeachers[k]) mergedTeachers[k] = t;
      });

      const mergedTokens = {};
      (registry.tokens || []).forEach(tk => { mergedTokens[tk.code] = tk; });
      (body.tokens || []).forEach(tk => {
        if (!mergedTokens[tk.code]) mergedTokens[tk.code] = tk;
      });

      registry.teachers = Object.values(mergedTeachers);
      registry.tokens = Object.values(mergedTokens);
      writeJsonFile(root, 'master_registry.json', registry);

      const allSpaces = loadAllTeacherSpacesMap(root);
      syncToGoogleSheets(root, registry, allSpaces);

      const reg2 = getRegistry_(root);
      return jsonResponse({
        status: 'ok',
        teachers: registry.teachers,
        tokens: registry.tokens,
        orders: reg2.orders,
        prices: hargaPaket_(reg2),
        paymentInfo: reg2.paymentInfo || ''
      });
    }

    // Kirim rekap ke grup WA kelas
    if (action === 'sendWaGroup') {
      if (!waKunciValid_(body)) return jsonResponse({ status: 'error', message: 'Kunci keamanan WA salah' });
      const g = waKirimGrup_(body.target, body.message);
      return jsonResponse(g.ok ? { status: 'ok' } : { status: 'error', message: g.detail });
    }

    // Daftar grup WA untuk dipilih guru
    if (action === 'listWaGroups') {
      if (!waKunciValid_(body)) return jsonResponse({ status: 'error', message: 'Kunci keamanan WA salah' });
      const d = waDaftarGrup_(!!body.refresh);
      return jsonResponse(d.ok ? { status: 'ok', groups: d.groups, detail: d.detail } : { status: 'error', message: d.detail });
    }

    // Pesan tes WhatsApp dari menu Data Kelas
    if (action === 'testWa') {
      if (!waKunciValid_(body)) return jsonResponse({ status: 'error', message: 'Kunci keamanan WA salah' });
      const r = waKirim_(body.target, '✅ Tes notifikasi English For Nusantara berhasil.\nPesan ini dikirim dari server aplikasi.');
      return jsonResponse(r.ok ? { status: 'ok' } : { status: 'error', message: r.detail });
    }

    // Uji koneksi dari aplikasi (Cek Status Sistem)
    if (action === 'ping') {
      return jsonResponse({ status: 'ok', pong: true });
    }

    return jsonResponse({ status: 'error', message: 'Unknown action' });
  } catch (err) {
    return jsonResponse({ status: 'error', message: String(err) });
  } finally {
    if (lock) {
      try { lock.releaseLock(); } catch (x) {}
    }
  }
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
