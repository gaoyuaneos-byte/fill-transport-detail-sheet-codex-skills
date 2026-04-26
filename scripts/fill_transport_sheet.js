#!/usr/bin/env node
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

function parseArgs(argv) {
  const args = {
    source: '交通明细表.xlsx',
    target: '交通明细表_已填.xlsx',
    pdfDir: process.cwd(),
    clearExisting: false,
    blankHighwayNotes: false,
  };
  for (let i = 2; i < argv.length; i++) {
    const a = argv[i];
    if (a === '--clear-existing') args.clearExisting = true;
    else if (a === '--blank-highway-notes') args.blankHighwayNotes = true;
    else if (a.startsWith('--')) {
      const key = a.slice(2).replace(/-([a-z])/g, (_, c) => c.toUpperCase());
      args[key] = argv[++i];
    }
  }
  return args;
}

function excelSerial(year, month, day, hour, minute) {
  return (Date.UTC(year, month - 1, day, hour, minute) - Date.UTC(1899, 11, 30)) / 86400000;
}

function normalizeText(text) {
  return text
    .replace(/特惠快\s+车/g, '特惠快车')
    .replace(/石家庄\s+市/g, '石家庄市')
    .replace(/杭州\s+市/g, '杭州市')
    .replace(/合肥\s+市/g, '合肥市')
    .replace(/\s+/g, ' ')
    .trim();
}

function cleanPlace(s) {
  return String(s || '').replace(/^[^|]+\|/, '').replace(/[()]/g, '').trim();
}

function cleanExtractedPlace(s) {
  return String(s || '').replace(/([\u4e00-\u9fa5])\s+([\u4e00-\u9fa5])/g, '$1$2').trim();
}

function routeNote(start, end) {
  return `从${cleanPlace(start)}去${cleanPlace(end)}`;
}

function inferYear(text) {
  const m = text.match(/行程起止日期：(\d{4})-/);
  return m ? Number(m[1]) : new Date().getFullYear();
}

function parseDidiText(text, filename) {
  const normalized = normalizeText(text);
  const year = inferYear(normalized);
  const rows = [];
  const rowStart = /(?:^| )(\d{1,2}) (?:特惠快车|快车|独享|出租车|优享|专车) /g;
  const starts = [];
  let match;
  while ((match = rowStart.exec(normalized))) {
    starts.push({ index: match.index, nextIndex: rowStart.lastIndex, seq: Number(match[1]) });
  }

  for (let i = 0; i < starts.length; i++) {
    const chunk = normalized.slice(starts[i].nextIndex, starts[i + 1]?.index || normalized.length);
    const prefix = chunk.match(/(\d{2})-(\d{2}) (\d{2}):(\d{2}) 周\s*\S+\s+([\u4e00-\u9fa5]{2,8}市) /);
    if (!prefix) continue;
    const afterPrefix = chunk.slice(prefix.index + prefix[0].length);
    const nums = [...afterPrefix.matchAll(/ ([0-9]+(?:\.[0-9]+)?)/g)];
    if (nums.length < 2) continue;
    const mileage = nums[0];
    const amount = nums[1];

    const body = afterPrefix.slice(0, mileage.index).trim();
    const city = prefix[5];
    const knownEndPrefixes = [
      '临平区|', '上城区|', '包河区|', '肥东县|', '蜀山区|', '经济技术开发区|',
      '新石|', '临平|', '东湖|', '新天路|', '店埠|', '长安路|', '得心路|',
      '正定机场', '畅和家园', '安徽智飞', '中电四公司',
    ];
    let splitAt = -1;
    for (const marker of knownEndPrefixes) {
      const idx = body.indexOf(` ${marker}`);
      if (idx > 0) {
        splitAt = idx;
        break;
      }
    }
    if (splitAt < 0) {
      const parts = body.split(' ');
      splitAt = Math.max(1, Math.floor(parts.length / 2));
      rows.push({
        source: filename,
        seq: starts[i].seq,
        dateTime: `${year}-${prefix[1]}-${prefix[2]} ${prefix[3]}:${prefix[4]}`,
        city,
        start: cleanExtractedPlace(parts.slice(0, splitAt).join(' ')),
        end: cleanExtractedPlace(parts.slice(splitAt).join(' ')),
        amount: Number(amount[1]),
        feeType: '行程费',
      });
    } else {
      rows.push({
        source: filename,
        seq: starts[i].seq,
        dateTime: `${year}-${prefix[1]}-${prefix[2]} ${prefix[3]}:${prefix[4]}`,
        city,
        start: cleanExtractedPlace(body.slice(0, splitAt)),
        end: cleanExtractedPlace(body.slice(splitAt)),
        amount: Number(amount[1]),
        feeType: '行程费',
      });
    }
  }
  return rows;
}

async function extractTripsFromPdfs(pdfDir) {
  let PDFParse;
  try {
    PDFParse = require('pdf-parse').PDFParse;
  } catch {
    throw new Error('Missing dependency pdf-parse. Run: npm install pdf-parse xlsx');
  }
  const trips = [];
  for (const file of fs.readdirSync(pdfDir).filter(f => f.toLowerCase().endsWith('.pdf')).sort()) {
    const parser = new PDFParse({ data: fs.readFileSync(path.join(pdfDir, file)) });
    const result = await parser.getText();
    await parser.destroy();
    trips.push(...parseDidiText(result.text, file));
  }
  return trips;
}

function readJson(file) {
  return JSON.parse(fs.readFileSync(file, 'utf8'));
}

async function main() {
  const args = parseArgs(process.argv);
  if (!args.projectNo || !args.projectName) {
    throw new Error('Required: --project-no and --project-name');
  }

  const trips = args.tripsJson ? readJson(args.tripsJson) : await extractTripsFromPdfs(args.pdfDir);
  const workByDate = args.workJson ? readJson(args.workJson) : {};
  trips.sort((a, b) => a.dateTime.localeCompare(b.dateTime));

  const wb = XLSX.readFile(args.source, { cellStyles: true });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const range = XLSX.utils.decode_range(ws['!ref']);
  const firstDataRow = 4;
  const template = {};
  for (let c = 0; c <= 8; c++) {
    const cell = ws[XLSX.utils.encode_cell({ r: firstDataRow, c })];
    if (cell) template[c] = { z: cell.z, s: cell.s };
  }

  if (args.clearExisting) {
    for (let r = firstDataRow; r <= range.e.r; r++) {
      for (let c = 0; c <= 8; c++) delete ws[XLSX.utils.encode_cell({ r, c })];
    }
    range.e.r = firstDataRow - 1;
  }

  let rowIndex = args.clearExisting ? firstDataRow : range.e.r + 1;
  trips.forEach((trip, i) => {
    const [datePart, timePart] = trip.dateTime.split(' ');
    const [year, month, day] = datePart.split('-').map(Number);
    const [hour, minute] = timePart.split(':').map(Number);
    const isHighway = trip.feeType === '高速费';
    const note = isHighway && args.blankHighwayNotes ? '' : (workByDate[datePart] || routeNote(trip.start, trip.end));
    const row = [
      rowIndex - firstDataRow + 1,
      args.projectNo,
      args.projectName,
      excelSerial(year, month, day, hour, minute),
      trip.city,
      trip.start,
      trip.end,
      Number(trip.amount),
      note,
    ];
    row.forEach((value, c) => {
      const addr = XLSX.utils.encode_cell({ r: rowIndex, c });
      ws[addr] = { t: typeof value === 'number' ? 'n' : 's', v: value, ...(template[c] || {}) };
      if (c === 3) ws[addr].z = 'm/d/yy h:mm';
    });
    rowIndex++;
  });

  range.e.r = rowIndex - 1;
  ws['!ref'] = XLSX.utils.encode_range(range);
  XLSX.writeFile(wb, args.target, { bookType: 'xlsx', cellStyles: true });
  const sum = trips.reduce((total, trip) => total + Number(trip.amount || 0), 0);
  console.log(`Wrote ${args.target}`);
  console.log(`Rows: ${trips.length}`);
  console.log(`Total: ${sum.toFixed(2)}`);
}

main().catch(err => {
  console.error(err.message || err);
  process.exit(1);
});
