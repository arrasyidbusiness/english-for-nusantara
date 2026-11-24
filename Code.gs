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
    'Sekolah', 'Guru Pengampu', 'Kelas/Rombel', 'Kode Kelas', 'Nama Siswa', 'L/P', 'PIN', 'Hadir (dari 30)'
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
const TRACKED_SCALARS_ = ['sts1', 'sas1', 'sts2', 'sas2', 'remedial1', 'remedial2', 'katrol1', 'katrol2'];
const TRACKED_MAPS_ = ['scores', 'examResults', 'examViolations', 'lkpdSubmissions'];

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
    if (action === 'syncTeacherDriveBackup' || action === 'adminSync') {
      lock = LockService.getScriptLock();
      lock.waitLock(30000);
    }

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
      // Data lisensi guru hanya boleh diubah oleh guru, bukan oleh siswa
      const teacherRecord = senderRole === 'student'
        ? ((saved && saved.teacherRecord) || body.teacherRecord || null)
        : (body.teacherRecord || (saved && saved.teacherRecord) || null);

      const fileUrl = writeJsonFile(root, safeFileName, {
        teacherKey: email,
        teacherRecord: teacherRecord,
        teacherSpace: mergedSpace,
        updatedAt: new Date().toISOString()
      });

      const registry = readJsonFile(root, 'master_registry.json', { teachers: [], tokens: [], classIndex: {} });
      if (!registry.teachers) registry.teachers = [];
      if (!registry.classIndex) registry.classIndex = {};
      if (teacherRecord && senderRole !== 'student') {
        const idx = registry.teachers.findIndex(t => String(t.email).toLowerCase() === email);
        if (idx >= 0) registry.teachers[idx] = teacherRecord;
        else registry.teachers.push(teacherRecord);
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
        teacherSpace: saved.teacherSpace
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
        teacherSpace: saved.teacherSpace
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

      return jsonResponse({
        status: 'ok',
        teachers: registry.teachers,
        tokens: registry.tokens
      });
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
