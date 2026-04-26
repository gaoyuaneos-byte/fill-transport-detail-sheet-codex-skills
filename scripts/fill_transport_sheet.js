#!/usr/bin/env node
'use strict';

const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

// ============================================================================
// Argument parsing
// ============================================================================

function parseArgs(argv) {
  const args = {
    source: '交通明细表.xlsx',
    target: '交通明细表_已填.xlsx',
    pdfDir: process.cwd(),
    sourceFormat: 'auto',
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
  if (!args.projectNo || !args.projectName) {
    throw new Error(
      '缺少必填参数。用法:\n' +
      '  node fill_transport_sheet.js --project-no "PAEEXXXX" --project-name "XX项目" [--source "交通明细表.xlsx"] [--target "输出.xlsx"] [--clear-existing] [--blank-highway-notes] [--source-format didi|hellobike|auto] [--pdf-dir <dir>] [--trips-json <file>] [--work-json <file>]'
    );
  }
  return args;
}

// ============================================================================
// Text normalization
// ============================================================================

/**
 * Normalize PDF-extracted text: fix known service-type splits and generic
 * "XX 市" city-name splits, then collapse remaining whitespace.
 */
function normalizeText(text) {
  return text
    // Known ride-service types that get split across lines in Didi PDFs
    .replace(/特惠快\s+车/g, '特惠快车')
    .replace(/特惠\s+快车/g, '特惠快车')
    .replace(/顺风\s+车/g, '顺风车')
    .replace(/滴滴快\s+车/g, '滴滴快车')
    .replace(/滴滴专\s+车/g, '滴滴专车')
    .replace(/惊喜\s+特价/g, '惊喜特价')
    // Fix time splits: "15: 55" → "15:55"
    .replace(/(\d{2}):\s+(\d{2})/g, '$1:$2')
    // Strip page markers that pollute number extraction
    .replace(/页码[：:]\s*\d+\s*\/\s*\d+/g, '')
    .replace(/--\s*\d+\s+of\s+\d+\s*--/g, '')
    // Generic: "XX 市" → "XX市"
    .replace(/([\p{Script=Han}]{2,8})\s+市/gu, '$1市')
    // Collapse remaining whitespace to single spaces
    .replace(/\s+/g, ' ')
    .trim();
}

// ============================================================================
// Address helpers
// ============================================================================

/**
 * Strip the "district|" prefix from a place name, keeping only the detail part.
 * "XX县|XX地名" → "XX地名"
 */
function cleanPlace(s) {
  return String(s || '').replace(/^[^|]+\|/, '').replace(/[()]/g, '').trim();
}

/**
 * Generate a route-note string from start/end places.
 */
function routeNote(start, end) {
  return `从${cleanPlace(start)}去${cleanPlace(end)}`;
}

/**
 * Split a combined address body into [startAddr, endAddr].
 *
 * Didi PDF bodies look like:
 *   "XX县|XX地名 XX区|XX站-进站口"
 *
 * Strategy: find all "XXX|" patterns (district/road markers).
 * The *second* "XXX|" typically begins the end address.
 * Falls back to a midpoint split when no pipe markers are present.
 */
function splitAddressPair(body) {
  const pipePat = /([\p{Script=Han}]{1,8})\|/gu;
  const pipes = [...body.matchAll(pipePat)];

  if (pipes.length >= 2) {
    const splitIdx = pipes[1].index;
    return [body.slice(0, splitIdx).trim(), body.slice(splitIdx).trim()];
  }

  if (pipes.length === 1) {
    // One pipe: start address has the district prefix; split the remainder.
    const after = pipes[0].index + pipes[0][0].length;
    const start = body.slice(0, after).trim();
    const rest = body.slice(after).trim();
    const parts = rest.split(/\s+/);
    const mid = Math.max(1, Math.floor(parts.length / 2));
    return [
      (start + ' ' + parts.slice(0, mid).join(' ')).trim(),
      parts.slice(mid).join(' '),
    ];
  }

  // No pipes — split by spaces roughly in half.
  const parts = body.split(/\s+/);
  const mid = Math.max(1, Math.floor(parts.length / 2));
  return [parts.slice(0, mid).join(' '), parts.slice(mid).join(' ')];
}

// ============================================================================
// Date helpers
// ============================================================================

function inferYear(text) {
  const m = text.match(/行程起止日期[：:]\s*(\d{4})-/);
  return m ? Number(m[1]) : new Date().getFullYear();
}

function excelSerial(year, month, day, hour, minute) {
  return (Date.UTC(year, month - 1, day, hour, minute) - Date.UTC(1899, 11, 30)) / 86400000;
}

// ============================================================================
// Didi PDF parser
// ============================================================================

/**
 * Parse Didi itinerary PDF text into trip records.
 * Handles: 快车, 特惠快车, 专车, 优享, 出租车, 顺风车, 独享
 */
function parseDidiText(text, filename) {
  const normalized = normalizeText(text);
  const year = inferYear(normalized);
  const rows = [];

  // Match ride type lines: "序号 车型 ..."
  const rideTypes = '(?:特惠快车|滴滴快车|滴滴专车|快车|专车|优享|出租车|顺风车|独享|惊喜特价)';
  const rowStart = new RegExp(`(?:^| )(\\d{1,2}) ${rideTypes} `, 'g');
  const starts = [];
  let m;
  while ((m = rowStart.exec(normalized))) {
    starts.push({ index: m.index, nextIdx: rowStart.lastIndex, seq: Number(m[1]) });
  }

  for (let i = 0; i < starts.length; i++) {
    const chunk = normalized.slice(starts[i].nextIdx, starts[i + 1]?.index || normalized.length);

    // Prefix: MM-DD HH:mm 周X 城市市
    const prefix = chunk.match(/(\d{2})-(\d{2}) (\d{2}):(\d{2}) 周\s*\S+\s+([\p{Script=Han}]{2,8}市) /u);
    if (!prefix) continue;

    const afterPrefix = chunk.slice(prefix.index + prefix[0].length);

    // Find numeric fields (mileage, amount) at the end
    const nums = [...afterPrefix.matchAll(/ ([0-9]+(?:\.[0-9]+)?)/g)];
    if (nums.length < 2) continue;

    const body = afterPrefix.slice(0, nums[nums.length - 2].index).trim();
    const amount = Number(nums[nums.length - 1][1]);

    const [month, day, hour, minute] = prefix.slice(1, 5);
    const city = prefix[5];

    const [start, end] = splitAddressPair(body);

    rows.push({
      source: filename,
      seq: starts[i].seq,
      dateTime: `${year}-${month}-${day} ${hour}:${minute}`,
      city,
      start: start.replace(/\s+/g, ''),
      end: end.replace(/\s+/g, ''),
      amount,
      feeType: '行程费',
    });
  }

  return rows;
}

// ============================================================================
// Hellobike / tabular parser
// ============================================================================

/**
 * Parse Hellobike (哈啰) or generic pasted table text into trip records.
 *
 * Expected formats (auto-detected):
 *   Tab-separated:
 *     2026-01-21\t10:27\tXX市\tXX县XX地名\tXX区XX站\t39.10
 *   Multi-space:
 *     2026-01-21  10:27  XX市  XX县XX地名  XX区XX站  39.10
 *
 * Lines that look like headers (containing 日期/起点/终点 etc.) are skipped.
 */
function parseHellobikeText(text, filename) {
  const lines = text.split(/\r?\n/).filter(l => l.trim());
  const rows = [];

  for (const line of lines) {
    // Skip header / summary lines
    if (/^(日期|时间|序号|行程|合计|总计|小计)/.test(line.trim())) continue;
    if (/起点.*终点|终点.*起点/.test(line)) continue;

    // Try to find a date-time pattern
    const dtMatch = line.match(/(\d{4})[-/](\d{1,2})[-/](\d{1,2})\s+(\d{1,2}):(\d{2})/);
    if (!dtMatch) continue;

    const before = line.slice(0, dtMatch.index).trim();
    const after = line.slice(dtMatch.index + dtMatch[0].length).trim();

    // Skip if there's non-numeric junk before the date (e.g. row numbers are fine)
    if (before && !/^\d+$/.test(before)) continue;

    // Split remaining fields: tab or 2+ spaces
    const fields = after.split(/\t|\s{2,}/).filter(Boolean);
    if (fields.length < 3) continue;

    // Fields: city, start, end, [extra...], amount
    const city = fields[0];
    // Amount is the last numeric-looking field
    let amountIdx = -1;
    for (let i = fields.length - 1; i >= 0; i--) {
      if (/^\d+\.?\d*$/.test(fields[i])) { amountIdx = i; break; }
    }
    if (amountIdx < 2) continue;

    const amount = Number(fields[amountIdx]);
    const addrFields = fields.slice(1, amountIdx);

    // Split address fields roughly in half
    const mid = Math.max(1, Math.floor(addrFields.length / 2));
    const start = addrFields.slice(0, mid).join('');
    const end = addrFields.slice(mid).join('');

    rows.push({
      source: filename,
      seq: null,
      dateTime: `${dtMatch[1]}-${dtMatch[2].padStart(2, '0')}-${dtMatch[3].padStart(2, '0')} ${dtMatch[4]}:${dtMatch[5]}`,
      city,
      start,
      end,
      amount,
      feeType: '行程费',
    });
  }

  // Assign sequence numbers
  rows.forEach((r, i) => { r.seq = i + 1; });
  return rows;
}

// ============================================================================
// Format detection & extraction dispatch
// ============================================================================

function detectFormat(filename, text) {
  const lower = filename.toLowerCase();
  if (/哈[啰罗喽]|hellobike|顺风车/.test(lower)) return 'hellobike';
  if (/滴滴|didi/.test(lower)) return 'didi';
  // Content-based detection
  if (text && /特惠快车|滴滴快车|滴滴专车/.test(text)) return 'didi';
  if (text && /\d{4}[-/]\d{1,2}[-/]\d{1,2}\s+\d{1,2}:\d{2}/.test(text) && !/周\s*\S+\s+市/.test(text)) return 'hellobike';
  return 'didi'; // default
}

async function extractTripsFromPdfs(pdfDir, sourceFormat) {
  let PDFParse;
  try {
    PDFParse = require('pdf-parse').PDFParse;
  } catch (e) {
    throw new Error(
      '缺少 pdf-parse 依赖。请运行:\n  npm install pdf-parse xlsx\n' +
      '如果已安装但仍然报错，检查是否在正确的目录下运行。'
    );
  }

  const pdfFiles = fs.readdirSync(pdfDir)
    .filter(f => f.toLowerCase().endsWith('.pdf'))
    .sort();

  if (pdfFiles.length === 0) {
    throw new Error(`在 ${pdfDir} 中没有找到 PDF 文件。`);
  }

  const trips = [];
  for (const file of pdfFiles) {
    const filePath = path.join(pdfDir, file);
    const parser = new PDFParse({ data: fs.readFileSync(filePath) });
    const result = await parser.getText();
    await parser.destroy();

    const format = sourceFormat === 'auto' ? detectFormat(file, result.text) : sourceFormat;

    if (format === 'hellobike') {
      const parsed = parseHellobikeText(result.text, file);
      if (parsed.length === 0) {
        console.warn(`[警告] ${file}: 未能从哈啰格式解析出任何行程，尝试滴滴格式...`);
        trips.push(...parseDidiText(result.text, file));
      } else {
        trips.push(...parsed);
      }
    } else {
      const parsed = parseDidiText(result.text, file);
      if (parsed.length === 0) {
        console.warn(
          `[警告] ${file}: PDF文本提取可能不完整（文本长度=${result.text.length}）。\n` +
          `  如果PDF只有空白表格，请直接粘贴行程数据并使用 --trips-json 参数。`
        );
      }
      trips.push(...parsed);
    }
  }

  return trips;
}

// ============================================================================
// Excel helpers
// ============================================================================

const HEADER_SIGNALS = ['序号', '项目号', '项目名称', '日期', '起点', '终点', '金额', '说明'];

/**
 * Auto-detect the first data row (right after the header row).
 * Returns a 0-based row index. Falls back to 4 if detection fails.
 */
function detectDataStartRow(ws, range) {
  // Scan first 15 rows for the header
  const maxScan = Math.min(range.e.r, 15);
  for (let r = 0; r <= maxScan; r++) {
    const cells = [];
    for (let c = 0; c <= Math.min(range.e.c, 8); c++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      cells.push(ws[addr] ? String(ws[addr].v || '') : '');
    }
    const matchCount = HEADER_SIGNALS.filter(h => cells.some(c => c.includes(h))).length;
    if (matchCount >= 5) return r + 1; // Data starts after header
  }
  return 4; // Conservative default
}

/**
 * Validate that the source workbook exists and has the expected structure.
 */
function validateSource(filePath) {
  if (!fs.existsSync(filePath)) {
    throw new Error(`源文件不存在: ${filePath}`);
  }
  const wb = XLSX.readFile(filePath);
  if (!wb.SheetNames.length) {
    throw new Error(`${filePath} 不包含任何工作表。`);
  }
  return wb;
}

function readJson(file) {
  if (!fs.existsSync(file)) {
    throw new Error(`JSON 文件不存在: ${file}`);
  }
  return JSON.parse(fs.readFileSync(file, 'utf8'));
}

// ============================================================================
// Main fill logic
// ============================================================================

async function main() {
  const args = parseArgs(process.argv);

  // --- 1. Get trips ---
  let trips;
  if (args.tripsJson) {
    trips = readJson(args.tripsJson);
  } else {
    trips = await extractTripsFromPdfs(args.pdfDir, args.sourceFormat);
  }

  if (trips.length === 0) {
    throw new Error('没有解析到任何行程记录。请检查 PDF 文件或使用 --trips-json 手动提供数据。');
  }

  // --- 2. Load work-content map ---
  const workByDate = args.workJson ? readJson(args.workJson) : {};

  // --- 3. Sort by departure time ---
  trips.sort((a, b) => a.dateTime.localeCompare(b.dateTime));

  // --- 4. Load & validate workbook ---
  const wb = validateSource(args.source);
  const sheetName = wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  const range = XLSX.utils.decode_range(ws['!ref'] || 'A1:I1');

  // --- 5. Detect data start row ---
  const firstDataRow = detectDataStartRow(ws, range);

  // --- 6. Capture template style from the first data row ---
  const template = {};
  for (let c = 0; c <= 8; c++) {
    const cell = ws[XLSX.utils.encode_cell({ r: firstDataRow, c })];
    if (cell) template[c] = { z: cell.z, s: cell.s };
  }

  // --- 7. Clear old data rows ---
  if (args.clearExisting) {
    for (let r = firstDataRow; r <= range.e.r; r++) {
      for (let c = 0; c <= 8; c++) delete ws[XLSX.utils.encode_cell({ r, c })];
    }
    range.e.r = firstDataRow - 1;
  }

  // --- 8. Write trip rows ---
  let rowIndex = args.clearExisting ? firstDataRow : range.e.r + 1;

  for (const trip of trips) {
    const [datePart, timePart] = trip.dateTime.split(' ');
    const [year, month, day] = datePart.split('-').map(Number);
    const [hour, minute] = timePart.split(':').map(Number);
    const isHighway = trip.feeType === '高速费';

    let note;
    if (isHighway && args.blankHighwayNotes) {
      note = '';
    } else if (workByDate[datePart]) {
      note = workByDate[datePart];
    } else {
      note = routeNote(trip.start, trip.end);
    }

    const values = [
      rowIndex - firstDataRow + 1,   // 序号
      args.projectNo,                // 项目号
      args.projectName,              // 项目名称
      excelSerial(year, month, day, hour, minute), // 日期
      trip.city,                     // 城市
      trip.start,                    // 起点
      trip.end,                      // 终点
      Number(trip.amount.toFixed(2)), // 金额
      note,                          // 说明
    ];

    values.forEach((value, c) => {
      const addr = XLSX.utils.encode_cell({ r: rowIndex, c });
      ws[addr] = {
        t: typeof value === 'number' ? 'n' : 's',
        v: value,
        ...(template[c] || {}),
      };
      if (c === 3) ws[addr].z = 'm/d/yy h:mm';
    });

    rowIndex++;
  }

  // --- 9. Update range & write ---
  range.e.r = rowIndex - 1;
  ws['!ref'] = XLSX.utils.encode_range(range);

  XLSX.writeFile(wb, args.target, { bookType: 'xlsx', cellStyles: true });

  // --- 10. Summary ---
  const sum = trips.reduce((t, trip) => t + Number(trip.amount || 0), 0);
  const highwayCount = trips.filter(t => t.feeType === '高速费').length;
  console.log(`输出文件: ${args.target}`);
  console.log(`行程行数: ${trips.length}${highwayCount ? ` (含 ${highwayCount} 条高速费)` : ''}`);
  console.log(`金额合计: ¥${sum.toFixed(2)}`);
  console.log(`数据起始行: ${firstDataRow + 1}`);
}

main().catch(err => {
  console.error('[错误] ' + (err.message || err));
  process.exit(1);
});
