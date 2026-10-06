const GC_SPREADSHEET_ID = '18oh2WCDf5p6xSyE1_sDOxCSY6HtV87fpMoEfhDm9cRs';
const GC_ROSTER_SHEET_NAME = 'SCHEDULE UPDATE';
const GC_REPORTS_SHEET_NAME = 'LAPORAN UPDATE SAMPEL';
const GC_HS_ALBUM_SHEET_NAME = 'ALBUM HS';
const GC_HS_DRIVE_FOLDER_NAME = 'GC Update Sampel - Album HS';
const GC_HS_MAX_IMAGE_BYTES = 100 * 1024;
const GC_HS_ALBUM_HEADERS = [
  'Kode Sampel',
  'Ni',
  'Fe',
  'SiO2',
  'MgO',
  'SM',
  'Link Foto',
  'Timestamp Upload'
];
const GC_REPORT_HEADERS = [
  'Tanggal',
  'Area PIT',
  'Shift',
  'Metode',
  'ID Metode',
  'Jumlah Sampel',
  'Timestamp Pengumpulan',
  'Nama Pelapor'
];

function doGet(e) {
  try {
    const parameters = e && e.parameter ? e.parameter : {};
    if (parameters.resource === 'roster') {
      const year = Number(parameters.year);
      if (!Number.isInteger(year) || year < 1900 || year > 9999) {
        throw new Error('Parameter year wajib berupa tahun kalender yang valid.');
      }

      return jsonOutput_({
        ok: true,
        data: readRoster_(year),
        error: null
      });
    }

    if (parameters.resource === 'laporan') {
      return jsonOutput_({
        ok: true,
        data: readReports_(),
        error: null
      });
    }
    if (parameters.resource === 'album-hs') {
      return jsonOutput_({
        ok: true,
        data: readHsAlbum_(),
        error: null
      });
    }

    throw new Error('Parameter resource harus bernilai roster atau laporan.');
  } catch (error) {
    return errorOutput_(error);
  }
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error('Body request JSON wajib diisi.');
    }

    let request;
    try {
      request = JSON.parse(e.postData.contents);
    } catch (parseError) {
      throw new Error('Body request bukan JSON yang valid.');
    }
    if (!request || typeof request !== 'object' || Array.isArray(request)) {
      throw new Error('Body request harus berupa objek JSON.');
    }

    const parameters = e.parameter || {};
    if (parameters.resource === 'album-hs' || request.resource === 'album-hs') {
      return jsonOutput_({
        ok: true,
        data: saveHsAlbumEntry_(request),
        error: null
      });
    }

    const submissionId = requiredText_(request.submissionId, 'Submission ID');
    if (submissionId.length > 128) {
      throw new Error('Submission ID tidak boleh lebih dari 128 karakter.');
    }
    const report = validateReport_(request);
    const saveResult = saveReportIdempotently_(report, submissionId);

    return jsonOutput_({
      ok: true,
      data: Object.assign(reportResponse_(report), {
        submissionId: submissionId,
        status: saveResult.alreadyExists ? 'already_exists' : 'created'
      }),
      error: null
    });
  } catch (error) {
    return errorOutput_(error);
  }
}

function saveReportIdempotently_(report, submissionId) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const spreadsheet = SpreadsheetApp.openById(GC_SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName(GC_REPORTS_SHEET_NAME);
    if (!sheet) {
      throw new Error('Sheet "' + GC_REPORTS_SHEET_NAME + '" tidak ditemukan.');
    }

    let headerMap = getReportHeaderMap_(sheet);
    if (headerMap.submissionId === undefined) {
      const submissionIdColumn = sheet.getLastColumn() + 1;
      sheet.getRange(1, submissionIdColumn).setValue('Submission ID');
      headerMap = getReportHeaderMap_(sheet);
    }

    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      const storedIds = sheet.getRange(
        2,
        headerMap.submissionId + 1,
        lastRow - 1,
        1
      ).getDisplayValues();
      if (storedIds.some(row => String(row[0]).trim() === submissionId)) {
        return { alreadyExists: true };
      }
    }

    const row = new Array(sheet.getLastColumn()).fill('');
    row[headerMap['Tanggal']] = report.tanggal;
    row[headerMap['Area PIT']] = report.pit;
    row[headerMap['Shift']] = report.shift;
    row[headerMap['Metode']] = report.metode;
    row[headerMap['ID Metode']] = report.idMetode;
    row[headerMap['Jumlah Sampel']] = report.jumlahSampel;
    row[headerMap['Timestamp Pengumpulan']] = report.timestampPengumpulan;
    row[headerMap['Nama Pelapor']] = report.namaPelapor;
    row[headerMap.submissionId] = submissionId;
    sheet.appendRow(row);
    return { alreadyExists: false };
  } finally {
    lock.releaseLock();
  }
}

function saveHsAlbumEntry_(request) {
  const entry = validateHsAlbumEntry_(request);
  const spreadsheet = SpreadsheetApp.openById(GC_SPREADSHEET_ID);
  const sheet = getOrCreateHsAlbumSheet_(spreadsheet);
  const folder = getOrCreateHsAlbumFolder_();
  const timestamp = new Date();
  const safeCode = entry.kodeSampel.replace(/[^\w-]+/g, '_').slice(0, 80);
  const fileName = safeCode + '_' + Utilities.formatDate(
    timestamp,
    Session.getScriptTimeZone(),
    'yyyyMMdd_HHmmss'
  ) + '.jpg';

  const file = folder.createFile(
    Utilities.newBlob(entry.imageBytes, 'image/jpeg', fileName)
  );

  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const link = file.getUrl();
    sheet.appendRow([
      entry.kodeSampel,
      entry.ni,
      entry.fe,
      entry.siO2,
      entry.mgo,
      entry.sm,
      link,
      timestamp
    ]);

    return {
      kodeSampel: entry.kodeSampel,
      ni: entry.ni,
      fe: entry.fe,
      siO2: entry.siO2,
      mgo: entry.mgo,
      sm: entry.sm,
      linkFoto: link,
      timestampUpload: timestamp.toISOString()
    };
  } catch (error) {
    file.setTrashed(true);
    throw error;
  }
}

function validateHsAlbumEntry_(request) {
  const kodeSampel = requiredText_(request.kodeSampel, 'Kode Sampel');
  const ni = requiredAssayValue_(request.ni, 'Ni');
  const fe = requiredAssayValue_(request.fe, 'Fe');
  const siO2 = requiredAssayValue_(request.siO2, 'SiO2');
  const mgo = requiredAssayValue_(request.mgo, 'MgO');
  const sm = requiredAssayValue_(request.sm, 'SM');
  const image = request.image;

  if (!image || image.mimeType !== 'image/jpeg' || typeof image.base64 !== 'string') {
    throw new Error('Foto wajib dikirim dalam format JPEG hasil kompresi.');
  }

  const imageBytes = Utilities.base64Decode(image.base64);
  if (imageBytes.length === 0 || imageBytes.length > GC_HS_MAX_IMAGE_BYTES) {
    throw new Error('Ukuran foto harus lebih dari 0 dan maksimal 100 KB.');
  }

  return {
    kodeSampel: kodeSampel,
    ni: ni,
    fe: fe,
    siO2: siO2,
    mgo: mgo,
    sm: sm,
    imageBytes: imageBytes
  };
}

function requiredAssayValue_(value, fieldName) {
  if (value === undefined || value === null || String(value).trim() === '') {
    throw new Error('Field wajib diisi: ' + fieldName + '.');
  }

  const numericValue = Number(value);
  if (!Number.isFinite(numericValue) || numericValue < 0) {
    throw new Error(fieldName + ' harus berupa angka valid yang tidak negatif.');
  }
  return numericValue;
}

function getOrCreateHsAlbumSheet_(spreadsheet) {
  let sheet = spreadsheet.getSheetByName(GC_HS_ALBUM_SHEET_NAME);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(GC_HS_ALBUM_SHEET_NAME);
    sheet.getRange(1, 1, 1, GC_HS_ALBUM_HEADERS.length).setValues([GC_HS_ALBUM_HEADERS]);
    return sheet;
  }

  const lastColumn = sheet.getLastColumn();
  if (lastColumn === 0) {
    sheet.getRange(1, 1, 1, GC_HS_ALBUM_HEADERS.length).setValues([GC_HS_ALBUM_HEADERS]);
    return sheet;
  }

  const headers = sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0]
    .map(header => String(header).trim());
  if (headers.length !== GC_HS_ALBUM_HEADERS.length ||
      GC_HS_ALBUM_HEADERS.some((header, index) => headers[index] !== header)) {
    throw new Error('Header sheet "' + GC_HS_ALBUM_SHEET_NAME + '" tidak sesuai konfigurasi Album HS.');
  }
  return sheet;
}

function getOrCreateHsAlbumFolder_() {
  const folders = DriveApp.getFoldersByName(GC_HS_DRIVE_FOLDER_NAME);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(GC_HS_DRIVE_FOLDER_NAME);
}

function authorizeDrive_() {
  DriveApp.getFoldersByName(GC_HS_DRIVE_FOLDER_NAME);
}

function readHsAlbum_() {
  const spreadsheet = SpreadsheetApp.openById(GC_SPREADSHEET_ID);
  const sheet = getOrCreateHsAlbumSheet_(spreadsheet);
  if (sheet.getLastRow() < 2) {
    return [];
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0]
    .map(header => String(header).trim());
  if (headers.length !== GC_HS_ALBUM_HEADERS.length ||
      GC_HS_ALBUM_HEADERS.some((header, index) => headers[index] !== header)) {
    throw new Error('Header sheet "' + GC_HS_ALBUM_SHEET_NAME + '" tidak sesuai konfigurasi Album HS.');
  }

  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();
  return rows
    .filter(row => row.some(value => value !== '' && value !== null))
    .map(row => ({
      kodeSampel: String(row[0] || ''),
      ni: row[1],
      fe: row[2],
      siO2: row[3],
      mgo: row[4],
      sm: row[5],
      linkFoto: String(row[6] || ''),
      timestampUpload: reportTimestampValue_(row[7])
    }))
    .reverse();
}

function validateReport_(request) {
  const tanggal = requiredText_(request.tanggal, 'Tanggal');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(tanggal) || !isValidDate_(tanggal)) {
    throw new Error('Tanggal wajib menggunakan format YYYY-MM-DD yang valid.');
  }

  const pit = requiredText_(request.pit, 'Area PIT');
  const shift = requiredText_(request.shift, 'Shift');
  const metode = requiredText_(request.methode || request.metode, 'Metode');
  const idMetode = requiredText_(request.idMetode || request.id, 'ID Metode');
  const namaPelapor = requiredText_(request.namaPelapor || request.petugas, 'Nama Pelapor');
  const jumlahSampel = Number(request.jumlahSampel !== undefined
    ? request.jumlahSampel
    : request.jumlah);
  if (!Number.isInteger(jumlahSampel) || jumlahSampel <= 0) {
    throw new Error('Jumlah Sampel wajib berupa bilangan bulat lebih dari 0.');
  }

  const timestampInput = request.timestampPengumpulan || request.createdAt;
  const timestampPengumpulan = timestampInput
    ? new Date(timestampInput)
    : new Date();
  if (isNaN(timestampPengumpulan.getTime())) {
    throw new Error('Timestamp Pengumpulan tidak valid.');
  }

  return {
    tanggal: tanggal,
    pit: pit,
    shift: shift,
    metode: metode,
    idMetode: idMetode,
    jumlahSampel: jumlahSampel,
    timestampPengumpulan: timestampPengumpulan,
    namaPelapor: namaPelapor
  };
}

function readReports_() {
  const spreadsheet = SpreadsheetApp.openById(GC_SPREADSHEET_ID);
  const sheet = spreadsheet.getSheetByName(GC_REPORTS_SHEET_NAME);
  if (!sheet) {
    throw new Error('Sheet "' + GC_REPORTS_SHEET_NAME + '" tidak ditemukan.');
  }

  const headerMap = getReportHeaderMap_(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return [];
  }

  const rows = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  return rows
    .filter(row => GC_REPORT_HEADERS.some(header => {
      const value = row[headerMap[header]];
      return value !== '' && value !== null;
    }))
    .map(row => ({
      tanggal: reportDateValue_(row[headerMap['Tanggal']]),
      pit: String(row[headerMap['Area PIT']] || ''),
      shift: String(row[headerMap['Shift']] || ''),
      methode: String(row[headerMap['Metode']] || ''),
      idMetode: String(row[headerMap['ID Metode']] || ''),
      jumlah: row[headerMap['Jumlah Sampel']],
      timestampPengumpulan: reportTimestampValue_(
        row[headerMap['Timestamp Pengumpulan']]
      ),
      namaPelapor: String(row[headerMap['Nama Pelapor']] || '')
    }));
}

function getReportHeaderMap_(sheet) {
  const lastColumn = sheet.getLastColumn();
  if (lastColumn === 0) {
    throw new Error('Header sheet "' + GC_REPORTS_SHEET_NAME + '" belum tersedia.');
  }

  const headers = sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0]
    .map(header => String(header).trim().toLowerCase());
  const headerMap = {};

  GC_REPORT_HEADERS.forEach(header => {
    const matches = [];
    headers.forEach((value, index) => {
      if (value === header.toLowerCase()) {
        matches.push(index);
      }
    });
    if (matches.length !== 1) {
      throw new Error(
        'Header "' + header + '" harus tersedia tepat satu kali pada sheet "' +
        GC_REPORTS_SHEET_NAME + '".'
      );
    }
    headerMap[header] = matches[0];
  });

  const submissionIdColumns = [];
  headers.forEach((value, index) => {
    if (value === 'submission id') {
      submissionIdColumns.push(index);
    }
  });
  if (submissionIdColumns.length > 1) {
    throw new Error('Header "Submission ID" hanya boleh tersedia satu kali.');
  }
  if (submissionIdColumns.length === 1) {
    headerMap.submissionId = submissionIdColumns[0];
  }

  return headerMap;
}

function requiredText_(value, fieldName) {
  if (value === undefined || value === null || String(value).trim() === '') {
    throw new Error('Field wajib diisi: ' + fieldName + '.');
  }
  return String(value).trim();
}

function isValidDate_(dateText) {
  const parts = dateText.split('-').map(Number);
  const date = new Date(Date.UTC(parts[0], parts[1] - 1, parts[2]));
  return date.getUTCFullYear() === parts[0] &&
    date.getUTCMonth() === parts[1] - 1 &&
    date.getUTCDate() === parts[2];
}

function reportDateValue_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return value === null || value === undefined ? '' : String(value);
}

function reportTimestampValue_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return value.toISOString();
  }
  return value === null || value === undefined ? '' : String(value);
}

function reportResponse_(report) {
  return {
    tanggal: report.tanggal,
    pit: report.pit,
    shift: report.shift,
    methode: report.metode,
    idMetode: report.idMetode,
    jumlah: report.jumlahSampel,
    timestampPengumpulan: report.timestampPengumpulan.toISOString(),
    namaPelapor: report.namaPelapor
  };
}

function errorOutput_(error) {
  return jsonOutput_({
    ok: false,
    data: null,
    error: {
      message: error && error.message ? error.message : String(error)
    }
  });
}

function readRoster_(year) {
  const spreadsheet = SpreadsheetApp.openById(GC_SPREADSHEET_ID);
  const sheet = spreadsheet.getSheetByName(GC_ROSTER_SHEET_NAME);
  if (!sheet) {
    throw new Error('Sheet "' + GC_ROSTER_SHEET_NAME + '" tidak ditemukan.');
  }

  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (lastRow === 0 || lastColumn < 3) {
    return [];
  }

  const range = sheet.getRange(1, 1, lastRow, lastColumn);
  const values = range.getValues();
  const displayValues = range.getDisplayValues();
  const calendar = findCalendar_(values, displayValues, year);
  const employees = [];
  let currentPosition = '';

  for (let rowIndex = calendar.rowIndex + 1; rowIndex < lastRow; rowIndex += 1) {
    const nama = String(displayValues[rowIndex][0] || '').trim();
    const jabatan = String(displayValues[rowIndex][1] || '').trim();

    const groupName = rosterGroupName_(nama);
    const positionGroup = rosterPositionGroup_(groupName);
    if (positionGroup) {
      currentPosition = positionGroup;
      continue;
    }
    if (groupName) {
      continue;
    }
    if (!nama || !jabatan) {
      continue;
    }

    const jadwal = calendar.columns.map(column => ({
      tanggal: isoDate_(year, calendar.month, column.day),
      hari: weekdayName_(year, calendar.month, column.day),
      kode: String(displayValues[rowIndex][column.columnIndex] || '')
    }));

    employees.push({
      nama: nama,
      jabatan: jabatan,
      posisiKerja: currentPosition,
      pit: currentPosition,
      jadwal: jadwal
    });
  }

  return employees;
}

function rosterGroupName_(nama) {
  const normalizedName = String(nama || '').toUpperCase().replace(/\s+/g, ' ').trim();
  return isGroupRow_(normalizedName) ? normalizedName : '';
}

function rosterPositionGroup_(groupName) {
  const positionGroups = {
    'LEADER TEAM': 'LEADER TEAM',
    'TEAM SAMPLER': 'TEAM SAMPLER',
    'BLOK A': 'BLOK A',
    'BLOK B': 'BLOK B',
    'INFRAS & BETA': 'INFRAS & BETA',
    BETA: 'BETA',
    DRIVER: 'DRIVER'
  };
  const normalized = String(groupName || '').toUpperCase().replace(/\s+/g, ' ').trim();
  return positionGroups[normalized] || '';
}

function findCalendar_(values, displayValues, year) {
  let bestRowIndex = -1;
  let bestColumns = [];

  for (let rowIndex = 0; rowIndex < values.length; rowIndex += 1) {
    let expectedDay = 1;
    const columns = [];

    for (let columnIndex = 2; columnIndex < values[rowIndex].length; columnIndex += 1) {
      const day = calendarDay_(values[rowIndex][columnIndex], displayValues[rowIndex][columnIndex]);
      if (day === expectedDay) {
        columns.push({ columnIndex: columnIndex, day: day });
        expectedDay += 1;
      } else if (columns.length > 0) {
        break;
      }
    }

    if (columns.length > bestColumns.length) {
      bestRowIndex = rowIndex;
      bestColumns = columns;
    }
  }

  if (bestRowIndex < 0 || bestColumns.length < 28) {
    throw new Error('Header kalender tanggal 1–28 atau lebih tidak ditemukan pada sheet.');
  }

  const month = findMonth_(displayValues, bestRowIndex, bestColumns);
  if (!month) {
    throw new Error('Nama bulan pada header kalender tidak ditemukan pada sheet.');
  }

  const daysInMonth = new Date(Date.UTC(year, month, 0)).getUTCDate();
  if (bestColumns.length > daysInMonth) {
    bestColumns = bestColumns.slice(0, daysInMonth);
  }

  return {
    rowIndex: bestRowIndex,
    month: month,
    columns: bestColumns
  };
}

function calendarDay_(value, displayValue) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return value.getDate();
  }

  if (typeof value === 'number' && Number.isInteger(value) && value >= 1 && value <= 31) {
    return value;
  }

  const text = String(displayValue || '').trim();
  const match = text.match(/(?:^|\D)(\d{1,2})\s*$/);
  if (!match) {
    return null;
  }

  const day = Number(match[1]);
  return day >= 1 && day <= 31 ? day : null;
}

function findMonth_(displayValues, calendarRowIndex, calendarColumns) {
  const monthMap = {
    january: 1, jan: 1, januari: 1,
    february: 2, feb: 2, februari: 2,
    march: 3, mar: 3, maret: 3,
    april: 4, apr: 4,
    may: 5, mei: 5,
    june: 6, jun: 6, juni: 6,
    july: 7, jul: 7, juli: 7,
    august: 8, aug: 8, agustus: 8, agu: 8,
    september: 9, sep: 9,
    october: 10, oct: 10, oktober: 10, okt: 10,
    november: 11, nov: 11,
    december: 12, dec: 12, desember: 12, des: 12
  };
  const firstCalendarColumn = calendarColumns[0].columnIndex;
  const lastCalendarColumn = calendarColumns[calendarColumns.length - 1].columnIndex;
  const firstHeaderRow = Math.max(0, calendarRowIndex - 3);

  for (let rowIndex = firstHeaderRow; rowIndex <= calendarRowIndex; rowIndex += 1) {
    for (let columnIndex = firstCalendarColumn; columnIndex <= lastCalendarColumn; columnIndex += 1) {
      const text = String(displayValues[rowIndex][columnIndex] || '').toLowerCase();
      const words = text.match(/[a-z]+/g) || [];
      for (let wordIndex = 0; wordIndex < words.length; wordIndex += 1) {
        if (monthMap[words[wordIndex]]) {
          return monthMap[words[wordIndex]];
        }
      }
    }
  }

  return null;
}

function isGroupRow_(name) {
  const normalized = String(name || '').toUpperCase().replace(/\s+/g, ' ').trim();
  const groupNames = [
    'LEADER TEAM',
    'TEAM SAMPLER',
    'BLOK A',
    'BLOK B',
    'INFRAS & BETA',
    'BETA',
    'DRIVER'
  ];

  return groupNames.indexOf(normalized) !== -1;
}

function isoDate_(year, month, day) {
  return String(year).padStart(4, '0') + '-' +
    String(month).padStart(2, '0') + '-' +
    String(day).padStart(2, '0');
}

function weekdayName_(year, month, day) {
  const weekdays = ['Minggu', 'Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'];
  return weekdays[new Date(Date.UTC(year, month - 1, day)).getUTCDay()];
}

function jsonOutput_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function authorizeDrive() {
  DriveApp.getFoldersByName(GC_HS_DRIVE_FOLDER_NAME);
}

function authorizeDriveWrite() {
  const folder = DriveApp.createFolder('GC_Update_Sampel_AUTH_TEST');
  folder.setTrashed(true);
}
