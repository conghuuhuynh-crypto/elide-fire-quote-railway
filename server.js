/**
 * server.js — Elide Fire Quote Server (Railway deployment)
 * adm-zip render XML + LibreOffice → PDF (không Carbone, không watermark)
 */

require('dotenv').config({ path: require('path').join(__dirname, '..', '.env') });
const express  = require('express');
const AdmZip   = require('adm-zip');
const path     = require('path');
const fs       = require('fs');
const https    = require('https');
const FormData = require('form-data');
const { exec, execSync } = require('child_process');
const os       = require('os');
const crypto   = require('crypto');

const app  = express();
const PORT = process.env.PORT || 3333;

// Python script để patch docx zip (viết 1 lần khi server start)
const PY_PATCH_SCRIPT = path.join(os.tmpdir(), 'patch_docx_elide.py');
fs.writeFileSync(PY_PATCH_SCRIPT, `import sys,zipfile
src,dst,xp=sys.argv[1],sys.argv[2],sys.argv[3]
xml=open(xp,'rb').read()
entries={}
with zipfile.ZipFile(src,'r') as z:
    for info in z.infolist():
        entries[info.filename]=(info,z.read(info.filename))
with zipfile.ZipFile(dst,'w',zipfile.ZIP_DEFLATED) as zout:
    for name,(info,data) in entries.items():
        if name=='word/document.xml': zout.writestr(name,xml)
        else: zout.writestr(info,data)
`, 'utf8');

// Prevent crash on unhandled errors
process.on('uncaughtException',  e => console.error('[uncaughtException]',  e.message));
process.on('unhandledRejection', e => console.error('[unhandledRejection]', e));

const TEMPLATE   = path.join(__dirname, 'templates', 'quote-template.docx');
const QUOTES_DIR = path.join(__dirname, 'outputs', 'quotes');
const SOFFICE    = process.platform === 'win32'
  ? '"C:\\Program Files\\LibreOffice\\program\\soffice.exe"'
  : 'soffice';
const PYTHON3    = process.platform === 'win32'
  ? '"C:\\Users\\Admin\\AppData\\Local\\Programs\\Python\\Python312\\python.exe"'
  : 'python3';

// NocoDB config — đọc từ .env, fallback cho Railway (xóa fallback sau khi set Railway Variables)
const NOCODB_HOST  = process.env.NOCODB_HOST  || 'nocodb-production-4d61.up.railway.app';
const NOCODB_TOKEN = process.env.NOCODB_TOKEN || '';
const NOCODB_BASE  = process.env.NOCODB_BASE  || 'p49wwa1uzmjtv1e';
const TABLE_NV     = process.env.NOCODB_TABLE_NV  || 'mbxi5rjran05biu'; // Nhan_vien
const TABLE_BG     = process.env.NOCODB_TABLE_BG  || 'mnfhtr9jysetk07'; // Bao_gia
const TABLE_SP     = process.env.NOCODB_TABLE_SP  || 'm1isvr6ljrp2klj'; // San_pham
const TABLE_HD     = process.env.NOCODB_TABLE_HD  || 'mudqz3rj45htmui'; // Hop_dong
const TABLE_CHAT   = process.env.NOCODB_TABLE_CHAT || 'muy359ghdcu7vo2'; // Chat_history

// AI Chat config
const { OpenAI } = require('openai');
const OPENROUTER_API_KEY = process.env.OPENROUTER_API_KEY || '';
const CHAT_MODEL = 'anthropic/claude-haiku-4.5'; // OpenRouter model ID
const openaiClient = new OpenAI({
  baseURL: 'https://openrouter.ai/api/v1',
  apiKey: OPENROUTER_API_KEY || 'sk-placeholder',
  defaultHeaders: {
    'HTTP-Referer': process.env.APP_URL || 'https://elide-fire-quote-railway-production.up.railway.app',
    'X-Title': 'Elide Fire Quote App'
  }
});
if (!OPENROUTER_API_KEY) console.warn('⚠️  OPENROUTER_API_KEY chưa được set — Chat AI sẽ không hoạt động');

// Cảnh báo sớm nếu thiếu biến bắt buộc
if (!NOCODB_TOKEN) console.warn('⚠️  NOCODB_TOKEN chưa được set — NocoDB calls sẽ thất bại');

const CONTRACT_TEMPLATE = path.join(__dirname, 'templates', 'contract-template.docx');

// Job queue
const jobs = {};

// ---- Helpers ----

/**
 * De-splice v2: gộp {{field}} bị Word cắt ra nhiều <w:r> runs, GIỮ NGUYÊN formatting (bold, size…)
 *
 * Nguyên nhân lỗi: Word thường tách {{field}} thành nhiều runs — run đầu chứa "{{", các run giữa
 * chứa tên field (có thể có <w:b/>), run cuối chứa "}}". De-splice đơn giản chỉ giữ rPr của
 * run đầu → mất bold. Hàm này tìm TOÀN BỘ chuỗi run chứa {{...}}, merge rPr từ TẤT CẢ run,
 * sau đó thay bằng 1 run duy nhất với rPr đã merge.
 */
function deSpliceFields(xml) {
  const parts = [];
  let pos = 0;

  while (pos < xml.length) {
    const braceStart = xml.indexOf('{{', pos);
    if (braceStart === -1) { parts.push(xml.slice(pos)); break; }

    // Walk forward to find matching }}
    let i = braceStart + 2, foundEnd = -1, foundEnd2 = -1;
    while (i < xml.length) {
      if (xml[i] === '<') {
        const gt = xml.indexOf('>', i);
        if (gt === -1) break;
        i = gt + 1;
        continue;
      }
      if (xml[i] === '}') {
        // Peek forward skipping XML tags to find next non-tag char
        let j = i + 1;
        while (j < xml.length && xml[j] === '<') {
          const gt = xml.indexOf('>', j);
          if (gt === -1) break;
          j = gt + 1;
        }
        if (j < xml.length && xml[j] === '}') { foundEnd = i; foundEnd2 = j; break; }
      }
      i++;
    }
    if (foundEnd === -1) { parts.push(xml.slice(pos)); break; }

    const inner = xml.slice(braceStart + 2, foundEnd);
    const clean = inner.replace(/<[^>]+>/g, '').replace(/\s+/g, ' ').trim();
    const endPos = foundEnd2 + 1; // sau ký tự } thứ hai (có thể cách xa về mặt XML)

    if (!inner.includes('<')) {
      // Single run — không cần xử lý thêm
      parts.push(xml.slice(pos, endPos));
      pos = endPos;
      continue;
    }

    // Split across runs — tìm <w:r> đầu chứa {{
    let runStart = braceStart, sp = braceStart;
    while (sp >= 0) {
      const idx = xml.lastIndexOf('<w:r', sp);
      if (idx === -1) break;
      const ch = xml[idx + 4];
      if (ch === '>' || ch === ' ') { runStart = idx; break; }
      sp = idx - 1;
    }

    // Tìm cuối </w:r> sau }}
    const lastRunClose = xml.indexOf('</w:r>', endPos);
    const runSeqEnd = lastRunClose !== -1 ? lastRunClose + 6 : endPos;

    // Thu thập và merge tất cả rPr trong chuỗi run
    const runSeq = xml.slice(runStart, runSeqEnd);
    const mergedTags = {};
    for (const m of runSeq.matchAll(/<w:rPr>([\s\S]*?)<\/w:rPr>/g)) {
      // Self-closing tags: <w:b/>, <w:bCs/>, <w:sz w:val="24"/>, <w:lang w:val="..."/>…
      for (const tm of m[1].matchAll(/<(w:\w+)(?:\s[^>]*)?\s*\/>/g)) {
        mergedTags[tm[1]] = tm[0]; // key = tag name, last value wins (size etc.)
      }
    }

    const mergedRPr = Object.keys(mergedTags).length > 0
      ? `<w:rPr>${Object.values(mergedTags).join('')}</w:rPr>`
      : '';

    parts.push(xml.slice(pos, runStart));
    parts.push(`<w:r>${mergedRPr}<w:t xml:space="preserve">{{${clean}}}</w:t></w:r>`);
    pos = runSeqEnd;
  }

  return parts.join('');
}

function escXml(s) {
  return String(s || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function fixCellLineSpacing(rowXml, tagStr, lineValue, before = 0, after = 0) {
  const ti = rowXml.indexOf(tagStr);
  if (ti === -1) return rowXml;
  const tcS = rowXml.lastIndexOf('<w:tc>', ti);
  const tcE = rowXml.indexOf('</w:tc>', ti) + 7;
  if (tcS === -1 || tcE === 6) return rowXml;
  let cell = rowXml.slice(tcS, tcE);
  const spacing = `<w:spacing w:before="${before}" w:after="${after}" w:line="${lineValue}" w:lineRule="auto"/>`;
  if (cell.includes('<w:spacing ')) {
    cell = cell.replace(/<w:spacing[^/]*\/>/, spacing);
  } else if (cell.includes('</w:pPr>')) {
    cell = cell.replace('</w:pPr>', `${spacing}</w:pPr>`);
  }
  return rowXml.slice(0, tcS) + cell + rowXml.slice(tcE);
}

function fixCellAlign(rowXml, tagStr, align) {
  const ti = rowXml.indexOf(tagStr);
  if (ti === -1) return rowXml;
  const tcS = rowXml.lastIndexOf('<w:tc>', ti);
  const tcE = rowXml.indexOf('</w:tc>', ti) + 7;
  if (tcS === -1 || tcE === 6) return rowXml;
  let cell = rowXml.slice(tcS, tcE);
  if (cell.includes('<w:jc ')) {
    cell = cell.replace(/<w:jc w:val="[^"]*"\/>/, `<w:jc w:val="${align}"/>`);
  } else if (cell.includes('</w:pPr>')) {
    cell = cell.replace('</w:pPr>', `<w:jc w:val="${align}"/></w:pPr>`);
  }
  return rowXml.slice(0, tcS) + cell + rowXml.slice(tcE);
}

function moTaToRuns(text, templateRun) {
  const rPrMatch = templateRun.match(/<w:rPr>[\s\S]*?<\/w:rPr>/);
  const rPr = rPrMatch ? rPrMatch[0] : '';
  // Bold rPr: thêm <w:b/><w:bCs/> nếu chưa có
  const boldRPr = rPr
    ? (rPr.includes('<w:b/>') ? rPr : rPr.replace('</w:rPr>', '<w:b/><w:bCs/></w:rPr>'))
    : '<w:rPr><w:b/><w:bCs/></w:rPr>';
  // Normal rPr: bỏ bold nếu có
  const normalRPr = rPr
    .replace(/<w:b\/>/g, '').replace(/<w:bCs\/>/g, '').replace(/<w:b \/>/g, '');

  const BOLD_LINES = 3; // 3 dòng đầu in đậm
  const lines = String(text || '').split('\n');
  return lines.map((line, i) => {
    const rp = i < BOLD_LINES ? boldRPr : normalRPr;
    return `<w:r>${rp}<w:t xml:space="preserve">${escXml(line)}</w:t></w:r>` +
      (i < lines.length - 1 ? `<w:r>${rp}<w:br/></w:r>` : '');
  }).join('');
}

// ---- Định dạng ngày tiếng Việt ----
function formatDateVietnamese(dateStr) {
  if (!dateStr) return '';
  const m = String(dateStr).match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
  if (m) return `Ngày ${parseInt(m[1])} tháng ${parseInt(m[2])} năm ${m[3]}`;
  return dateStr;
}

// ---- Số tiền bằng chữ (Tiếng Việt) ----
function soTienBangChu(amount) {
  const donvi = ['', 'một', 'hai', 'ba', 'bốn', 'năm', 'sáu', 'bảy', 'tám', 'chín'];
  if (!amount || amount === 0) return 'Không đồng';
  function docNhom(n) {
    const tram = Math.floor(n / 100), chuc = Math.floor((n % 100) / 10), dv = n % 10;
    let s = '';
    if (tram > 0) s += donvi[tram] + ' trăm ';
    if (chuc === 0 && dv > 0 && tram > 0) s += 'lẻ ' + donvi[dv];
    else if (chuc === 1) s += 'mười ' + (dv === 5 ? 'lăm' : dv > 0 ? donvi[dv] : '');
    else if (chuc > 1) s += donvi[chuc] + ' mươi ' + (dv === 1 ? 'mốt' : dv === 5 ? 'lăm' : dv > 0 ? donvi[dv] : '');
    else if (dv > 0) s += donvi[dv];
    return s.trim();
  }
  const hang = ['', 'nghìn', 'triệu', 'tỷ'];
  let groups = [], n = Math.round(amount);
  while (n > 0) { groups.push(n % 1000); n = Math.floor(n / 1000); }
  let result = '';
  for (let i = groups.length - 1; i >= 0; i--) {
    if (groups[i] > 0) result += docNhom(groups[i]) + (hang[i] ? ' ' + hang[i] : '') + ' ';
  }
  return result.trim().charAt(0).toUpperCase() + result.trim().slice(1) + ' đồng';
}

/**
 * Render contract template — hỗ trợ cú pháp {{field-name}} từ file Word
 * items: array [{mo_ta, don_vi_tinh, so_luong, don_gia}] để duplicate product row
 */
function renderContractTemplate(templatePath, fieldMap, items) {
  const tmpDocx = path.join(os.tmpdir(), `contract_${crypto.randomBytes(4).toString('hex')}.docx`);
  fs.copyFileSync(templatePath, tmpDocx);

  const zip = new AdmZip(tmpDocx);
  const xmlEntry = zip.getEntry('word/document.xml');
  if (!xmlEntry) return tmpDocx;

  let xml = xmlEntry.getData().toString('utf8');

  // De-splice: gộp {{field}} bị Word cắt ra nhiều runs, GIỮ NGUYÊN bold/size
  xml = deSpliceFields(xml);

  // Duplicate product row nếu có nhiều items
  if (items && items.length > 1) {
    // Tìm row chứa {{Mo-ta-san-pham}}
    const rowRegex = /<w:tr[ >][\s\S]*?<\/w:tr>/g;
    let productRowStr = null, productRowIndex = -1;
    let m;
    while ((m = rowRegex.exec(xml)) !== null) {
      if (m[0].includes('{{Mo-ta-san-pham}}')) {
        productRowStr = m[0];
        productRowIndex = m.index;
        break;
      }
    }
    if (productRowStr) {
      const fmtN = n => Math.round(n).toLocaleString('vi-VN');
      const originalRowLength = productRowStr.length; // lưu length gốc TRƯỚC khi modify

      // Fix alignment & line-spacing cho product row (giống quote)
      productRowStr = fixCellAlign(productRowStr, '{{Mo-ta-san-pham}}', 'left');
      productRowStr = fixCellAlign(productRowStr, '{{Don-gia}}',        'right');
      productRowStr = fixCellAlign(productRowStr, '{{Thanh-tien}}',     'right');
      productRowStr = fixCellLineSpacing(productRowStr, '{{Mo-ta-san-pham}}', 300, 80, 80);

      // Tìm templateRun chứa {{Mo-ta-san-pham}} để lấy formatting (bold/font)
      const moTaPos = productRowStr.indexOf('{{Mo-ta-san-pham}}');
      let moTaRunStart = -1, moTaRunEnd = -1, moTaRun = '';
      if (moTaPos !== -1) {
        let sp = moTaPos;
        while (sp >= 0) {
          const idx = productRowStr.lastIndexOf('<w:r', sp);
          if (idx === -1) break;
          const ch = productRowStr[idx + 4];
          if (ch === '>' || ch === ' ') { moTaRunStart = idx; break; }
          sp = idx - 1;
        }
        moTaRunEnd = productRowStr.indexOf('</w:r>', moTaPos) + 6;
        moTaRun    = moTaRunStart !== -1 ? productRowStr.slice(moTaRunStart, moTaRunEnd) : '';
      }

      const expandedRows = items.map((it, idx) => {
        const sl  = parseFloat(it.so_luong) || 0;
        const gia = parseFloat((it.don_gia||'').toString().replace(/\./g,'').replace(/,/g,'.')) || 0;
        const tt  = sl * gia;
        let row = productRowStr;
        // Mo-ta-san-pham: dùng moTaToRuns (bold + line-break)
        if (moTaRunStart !== -1) {
          row = row.slice(0, moTaRunStart) + moTaToRuns(it.mo_ta || '', moTaRun) + row.slice(moTaRunEnd);
        } else {
          row = row.split('{{Mo-ta-san-pham}}').join(escXml(it.mo_ta || ''));
        }
        row = row
          .split('{{stt}}').join(escXml(String(idx + 1)))
          .split('{{Đon-vi-tinh}}').join(escXml(it.don_vi_tinh || 'Cái'))
          .split('{{So-luong}}').join(escXml(String(sl)))
          .split('{{Don-gia}}').join(escXml(fmtN(gia)))
          .split('{{Thanh-tien}}').join(escXml(fmtN(tt)));
        return row;
      }).join('');
      xml = xml.slice(0, productRowIndex) + expandedRows + xml.slice(productRowIndex + originalRowLength);
    }
  }

  // Thay thế các field còn lại (bỏ qua product-row fields đã xử lý)
  const skipFields = items && items.length > 1
    ? ['stt', 'Mo-ta-san-pham', 'Đon-vi-tinh', 'So-luong', 'Don-gia', 'Thanh-tien']
    : [];
  for (const [key, val] of Object.entries(fieldMap)) {
    if (skipFields.includes(key)) continue;
    // Mo-ta-san-pham: dùng moTaToRuns để giữ định dạng nhiều dòng (bold 3 dòng đầu)
    if (key === 'Mo-ta-san-pham') {
      const moTaPos = xml.indexOf('{{Mo-ta-san-pham}}');
      if (moTaPos !== -1) {
        let runStart = -1, sp = moTaPos;
        while (sp >= 0) {
          const idx = xml.lastIndexOf('<w:r', sp);
          if (idx === -1) break;
          const ch = xml[idx + 4];
          if (ch === '>' || ch === ' ') { runStart = idx; break; }
          sp = idx - 1;
        }
        const runEnd = xml.indexOf('</w:r>', moTaPos) + 6;
        const moTaRun = runStart !== -1 ? xml.slice(runStart, runEnd) : '';
        if (runStart !== -1) {
          xml = xml.slice(0, runStart) + moTaToRuns(String(val || ''), moTaRun) + xml.slice(runEnd);
        } else {
          xml = xml.split('{{Mo-ta-san-pham}}').join(escXml(String(val || '')));
        }
      }
      continue;
    }
    xml = xml.split(`{{${key}}}`).join(escXml(String(val || '')));
  }

  const xmlTmp = tmpDocx + '.xml';
  fs.writeFileSync(xmlTmp, xml, 'utf8');
  try {
    execSync(`${PYTHON3} "${PY_PATCH_SCRIPT}" "${templatePath}" "${tmpDocx}" "${xmlTmp}"`, { timeout: 15000 });
  } finally {
    try { fs.unlinkSync(xmlTmp); } catch(_) {}
  }
  return tmpDocx;
}

/**
 * Render toàn bộ template bằng adm-zip (không Carbone):
 * - Expand items[] rows
 * - Thay thế tất cả {d.xxx} fields
 * Returns path to rendered docx (tmp file)
 */
function renderDocxTemplate(templatePath, data, items) {
  const tmpDocx = path.join(os.tmpdir(), `render_${crypto.randomBytes(4).toString('hex')}.docx`);
  fs.copyFileSync(templatePath, tmpDocx);

  const zip = new AdmZip(tmpDocx);
  const xmlEntry = zip.getEntry('word/document.xml');
  if (!xmlEntry) return tmpDocx; // fallback: trả về bản copy gốc

  let xml = xmlEntry.getData().toString('utf8');

  // 0. De-splice: gộp {d.xxx} bị Word cắt ra nhiều <w:r> runs
  xml = xml.replace(/\{d\.((?:[^<}]|<[^>]+>)*?)\}/g, (match, inner) => {
    const varName = inner.replace(/<[^>]+>/g, '');
    if (inner !== varName) return `{d.${varName}}`;
    return match;
  });

  // 0b. De-splice: gộp {{xxx}} bị Word cắt ra (vd: {{Ten_cong_ty}} trong template v2)
  xml = deSpliceFields(xml);

  // 1. Expand items rows
  if (items && items.length > 0) {
    let searchPos = 0, rowStart = -1, rowEnd = -1;
    while (true) {
      const s = xml.indexOf('<w:tr ', searchPos);
      if (s === -1) break;
      const e = xml.indexOf('</w:tr>', s) + 7;
      if (xml.slice(s, e).includes('d.items[i]')) { rowStart = s; rowEnd = e; break; }
      searchPos = e;
    }

    if (rowStart !== -1) {
      let templateRow = xml.slice(rowStart, rowEnd);
      // Fix alignment trước khi detect positions
      templateRow = fixCellAlign(templateRow, '{d.items[i].mo_ta}',    'left');
      templateRow = fixCellAlign(templateRow, '{d.items[i].don_gia}',  'right');
      templateRow = fixCellAlign(templateRow, '{d.items[i].thanh_tien}', 'right');
      templateRow = fixCellLineSpacing(templateRow, '{d.items[i].mo_ta}', 300, 80, 80); // 1.25x + 4pt top/bottom
      const moTaPos = templateRow.indexOf('{d.items[i].mo_ta}');
      let moTaRunStart = -1, moTaRunEnd = -1, moTaRun = '';
      if (moTaPos !== -1) {
        // Tìm <w:r> hoặc <w:r (có attribute) — không tìm <w:rPr> hay <w:rStyle>
        let sp = moTaPos;
        while (sp >= 0) {
          const idx = templateRow.lastIndexOf('<w:r', sp);
          if (idx === -1) break;
          const ch = templateRow[idx + 4]; // char sau <w:r
          if (ch === '>' || ch === ' ') { moTaRunStart = idx; break; }
          sp = idx - 1;
        }
        moTaRunEnd = templateRow.indexOf('</w:r>', moTaPos) + 6;
        moTaRun    = moTaRunStart !== -1 ? templateRow.slice(moTaRunStart, moTaRunEnd) : '';
      }

      const expandedRows = items.map(item => {
        let row = templateRow;
        if (moTaRunStart !== -1) {
          row = row.slice(0, moTaRunStart) + moTaToRuns(item.mo_ta, moTaRun) + row.slice(moTaRunEnd);
        } else {
          row = row.replace('{d.items[i].mo_ta}', escXml(item.mo_ta));
        }
        row = row.replace('{d.items[i].stt}',        escXml(item.stt));
        row = row.replace('{d.items[i].so_luong}',   escXml(item.so_luong));
        row = row.replace('{d.items[i].don_gia}',    escXml(item.don_gia));
        row = row.replace('{d.items[i].thanh_tien}', escXml(item.thanh_tien));
        return row;
      }).join('');

      xml = xml.slice(0, rowStart) + expandedRows + xml.slice(rowEnd);
    }
  }

  // 2. Thay thế tất cả {d.xxx} fields
  for (const [key, val] of Object.entries(data)) {
    xml = xml.split(`{d.${key}}`).join(escXml(String(val)));
  }

  // 2b. Thay thế {{CamelCase_field}} (template v2 dùng cho ten_cong_ty etc.) — case-insensitive
  for (const [key, val] of Object.entries(data)) {
    xml = xml.replace(new RegExp(`\\{\\{${key}\\}\\}`, 'gi'), escXml(String(val)));
  }

  // Dùng Python3 script để update zip — đảm bảo ZIP format chuẩn
  const xmlTmp = tmpDocx + '.xml';
  fs.writeFileSync(xmlTmp, xml, 'utf8');
  try {
    execSync(`${PYTHON3} "${PY_PATCH_SCRIPT}" "${TEMPLATE}" "${tmpDocx}" "${xmlTmp}"`, { timeout: 15000 });
  } finally {
    try { fs.unlinkSync(xmlTmp); } catch(_) {}
  }

  return tmpDocx;
}

app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(express.static(path.join(__dirname, 'public')));
app.use('/assets', express.static(path.join(__dirname, 'assets')));
app.use('/download', express.static(QUOTES_DIR));

app.get('/', (req, res) => res.sendFile(path.join(__dirname, 'public', 'index.html')));
app.get('/health', (req, res) => res.json({ status: 'ok', version: 'v25-bold-fix' }));

// Helper: NocoDB GET với timeout
function nocoGet(path, res) {
  const options = { hostname: NOCODB_HOST, path, headers: { 'xc-token': NOCODB_TOKEN } };
  const req = https.get(options, r => {
    let d = '';
    r.on('data', c => d += c);
    r.on('end', () => { try { res.json(JSON.parse(d).list || []); } catch(e) { res.json([]); } });
  });
  req.on('error', () => { if (!res.headersSent) res.json([]); });
  req.setTimeout(25000, () => { req.destroy(); if (!res.headersSent) res.json([]); });
}

// API nhân viên
app.get('/api/employees', (req, res) => {
  nocoGet(`/api/v1/db/data/noco/${NOCODB_BASE}/${TABLE_NV}?limit=100`, res);
});

// API sản phẩm
app.get('/api/products', (req, res) => {
  nocoGet(`/api/v1/db/data/noco/${NOCODB_BASE}/${TABLE_SP}?limit=200&sort=Id`, res);
});

// API danh sách báo giá cũ (để nạp lại form)
app.get('/api/quotes', (req, res) => {
  const search = (req.query.search || '').trim();
  const limit  = Math.min(parseInt(req.query.limit) || 30, 50);
  let qs = `limit=${limit}&sort=-Id`;
  if (search) {
    const s = encodeURIComponent(search);
    qs += `&where=(Ten_cong_ty,like,%25${s}%25)~or(So_bao_gia,like,%25${s}%25)`;
  }
  nocoGet(`/api/v1/db/data/noco/${NOCODB_BASE}/${TABLE_BG}?${qs}`, res);
});

// API job status
app.get('/api/job/:id', (req, res) => {
  const job = jobs[req.params.id];
  if (!job) return res.status(404).json({ status: 'not_found' });
  res.json(job);
});

// Helper: upload PDF
function uploadPdfToNocoDB(pdfPath, filename) {
  return new Promise((resolve) => {
    try {
      const form = new FormData();
      form.append('file', fs.createReadStream(pdfPath), { filename, contentType: 'application/pdf' });
      const options = {
        hostname: NOCODB_HOST,
        path: `/api/v1/db/storage/upload?path=noco/${NOCODB_BASE}/Bao_gia/File_PDF`,
        method: 'POST',
        headers: { ...form.getHeaders(), 'xc-token': NOCODB_TOKEN }
      };
      const req = https.request(options, r => {
        let d = '';
        r.on('data', c => d += c);
        r.on('end', () => {
          try {
            const parsed = JSON.parse(d);
            const att = Array.isArray(parsed) ? parsed[0] : parsed;
            resolve(att && att.path ? att : null);
          } catch (e) { resolve(null); }
        });
      });
      req.on('error', () => resolve(null));
      form.pipe(req);
    } catch (e) { resolve(null); }
  });
}

// Helper: lưu NocoDB
function saveQuoteToNocoDB(record) {
  return new Promise((resolve, reject) => {
    const body = Buffer.from(JSON.stringify(record));
    const options = {
      hostname: NOCODB_HOST,
      path: `/api/v1/db/data/noco/${NOCODB_BASE}/${TABLE_BG}`,
      method: 'POST',
      headers: { 'xc-token': NOCODB_TOKEN, 'Content-Type': 'application/json', 'Content-Length': body.length }
    };
    const req = https.request(options, r => {
      let d = '';
      r.on('data', c => d += c);
      r.on('end', () => {
        try { const res = JSON.parse(d); if (res.Id) resolve(res.Id); else reject(new Error('No Id: ' + d)); }
        catch (e) { reject(e); }
      });
    });
    req.on('error', reject);
    req.write(body);
    req.end();
  });
}

// Worker
async function runJob(jobId, b) {
  const job = jobs[jobId];
  const fmt = (n) => n.toLocaleString('vi-VN');

  let rawItems = [];
  if (Array.isArray(b.items)) rawItems = b.items;
  else if (typeof b.items === 'string') { try { rawItems = JSON.parse(b.items); } catch (_) {} }

  const validItems = rawItems.filter(it => parseFloat(it.so_luong) > 0);
  const ckTong = parseFloat(b.chiet_khau_tong) || 0;
  let tongTruoCK = 0;

  const items = validItems.map((it, idx) => {
    const qty   = parseFloat(it.so_luong) || 0;
    const price = parseFloat(it.don_gia)  || 0;
    const ck    = parseFloat(it.chiet_khau) || 0;
    const tt    = qty * price * (1 - ck / 100);
    tongTruoCK += tt;
    return {
      stt:        String(idx + 1).padStart(2, '0'),
      mo_ta:      it.mo_ta || it.model || '',
      so_luong:   String(qty),
      don_gia:    fmt(price * (1 - ck / 100)),
      thanh_tien: fmt(tt)
    };
  });

  const total = tongTruoCK * (1 - ckTong / 100);

  const data = {
    ten_cong_ty:       b.ten_cong_ty       || '',
    ten_phong_ban:     b.ten_phong_ban     || '',
    ten_nguoi_lien_he: b.ten_nguoi_lien_he || '',
    email_khach_hang:  b.email_khach_hang  || '',
    sdt_khach_hang:    b.sdt_khach_hang    || '',
    phien_ban:         b.phien_ban         || 'Phát hành lần đầu',
    ngay_bao_gia:      b.ngay_bao_gia      || new Date().toLocaleDateString('vi-VN'),
    so_bao_gia:        b.so_bao_gia        || '',
    ten_du_an:         b.ten_du_an         || '',
    bo_phan:           b.nv_bo_phan        || '',
    ten_nhan_vien:     b.nv_ten            || '',
    email_nhan_vien:   b.nv_email          || '',
    sdt_nhan_vien:     b.nv_sdt            || '',
    truoc_chiet_khau:  fmt(tongTruoCK),
    chiet_khau:        ckTong > 0 ? `${ckTong}%` : '0',
    tong_thanh_tien:   fmt(total),
  };

  const soSlug = (b.so_bao_gia || 'bao-gia').replace(/[\/\\:*?"<>|]/g, '-').trim();
  if (!fs.existsSync(QUOTES_DIR)) fs.mkdirSync(QUOTES_DIR, { recursive: true });
  const outPdf = path.join(QUOTES_DIR, `${soSlug}.pdf`);

  // Render docx với adm-zip (không Carbone)
  let renderedDocx;
  try {
    renderedDocx = renderDocxTemplate(TEMPLATE, data, items);
  } catch (e) {
    job.status = 'error'; job.error = 'Render error: ' + e.message; return;
  }

  // Lưu bản Word (trước khi LibreOffice convert)
  const outDocx = path.join(QUOTES_DIR, `${soSlug}.docx`);
  try { fs.copyFileSync(renderedDocx, outDocx); } catch (_) {}

  // Copy rendered docx vào QUOTES_DIR để LibreOffice output đúng chỗ
  const tmpDocx = path.join(QUOTES_DIR, `_tmp_${jobId}.docx`);
  try { fs.copyFileSync(renderedDocx, tmpDocx); } catch (e) {
    job.status = 'error'; job.error = 'Copy docx error: ' + e.message; return;
  }
  try { fs.unlinkSync(renderedDocx); } catch (_) {}

  const cmd = `${SOFFICE} --headless --convert-to pdf --outdir "${QUOTES_DIR}" "${tmpDocx}"`;
  console.log('[LO] cmd:', cmd);
  exec(cmd, { timeout: 120000 }, (err2, stdout, stderr) => {
    try { fs.unlinkSync(tmpDocx); } catch (_) {}
    if (err2) { job.status = 'error'; job.error = 'LibreOffice: ' + err2.message + ' stderr:' + stderr; return; }

    const libreOut = path.join(QUOTES_DIR, `_tmp_${jobId}.pdf`);
    if (!fs.existsSync(libreOut)) {
      const files = (() => { try { return fs.readdirSync(QUOTES_DIR); } catch(_) { return []; } })();
      job.status = 'error';
      job.error = 'PDF not found. Files:[' + files.join(',') + '] stderr:' + stderr;
      return;
    }
    try { fs.renameSync(libreOut, outPdf); } catch (_) {}

    const finalPath = fs.existsSync(outPdf) ? outPdf : libreOut;
    const finalName = path.basename(finalPath);
    const appUrl    = process.env.APP_URL || 'https://elide-fire-quote-railway-production.up.railway.app';

    job.status      = 'done';
    job.url         = `${appUrl}/download/${finalName}`;  // absolute (cho Railway)
    job.relativeUrl = `/download/${finalName}`;            // relative (cho local download)
    job.filename    = finalName;
    job.docxUrl     = `/download/${soSlug}.docx`;
    console.log('✅ PDF ready:', finalName);

    // Map per-product items (Techideas / Lovingcare)
    const findItem = (kw) => validItems.find(it =>
      String(it.mo_ta || it.model || '').toLowerCase().includes(kw.toLowerCase())
    );
    const itTech = findItem('techideas');
    const itLove = findItem('lovingcare');
    const calcTT = (it) =>
      (parseFloat(it.so_luong)||0) * (parseFloat(it.don_gia)||0) * (1 - (parseFloat(it.chiet_khau)||0)/100);

    const record = {
      So_bao_gia: b.so_bao_gia || '', Ngay_bao_gia: b.ngay_bao_gia || '',
      Phien_ban: b.phien_ban || '', Ten_du_an: b.ten_du_an || '',
      Ten_cong_ty: b.ten_cong_ty || '', Phong_ban_KH: b.ten_phong_ban || '',
      Nguoi_lien_he: b.ten_nguoi_lien_he || '', SDT_khach_hang: b.sdt_khach_hang || '',
      Email_khach_hang: b.email_khach_hang || '', NV_bo_phan: b.nv_bo_phan || '',
      NV_ten: b.nv_ten || '', NV_email: b.nv_email || '', NV_sdt: b.nv_sdt || '',
      CK_Tong_don: ckTong, Tong_thanh_toan: total,
      SL_Techideas:        itTech ? parseFloat(itTech.so_luong)    || 0 : 0,
      DonGia_Techideas:    itTech ? parseFloat(itTech.don_gia)     || 0 : 0,
      CK_Techideas:        itTech ? parseFloat(itTech.chiet_khau)  || 0 : 0,
      ThanhTien_Techideas: itTech ? calcTT(itTech) : 0,
      SL_Lovingcare:        itLove ? parseFloat(itLove.so_luong)   || 0 : 0,
      DonGia_Lovingcare:    itLove ? parseFloat(itLove.don_gia)    || 0 : 0,
      CK_Lovingcare:        itLove ? parseFloat(itLove.chiet_khau) || 0 : 0,
      ThanhTien_Lovingcare: itLove ? calcTT(itLove) : 0,
    };
    Promise.resolve()
      .then(() => uploadPdfToNocoDB(finalPath, finalName))
      .then(att => { if (att) record.File_PDF = [att]; return saveQuoteToNocoDB(record); })
      .then(() => console.log('✅ NocoDB saved'))
      .catch(e => console.error('NocoDB error:', e.message));
  });
}

// API generate
app.post('/api/generate', (req, res) => {
  const jobId = crypto.randomBytes(6).toString('hex');
  jobs[jobId] = { status: 'processing' };
  setTimeout(() => { delete jobs[jobId]; }, 3600000);
  setImmediate(() => runJob(jobId, req.body).catch(e => {
    if (jobs[jobId]) { jobs[jobId].status = 'error'; jobs[jobId].error = 'Unhandled: ' + e.message; }
    console.error('[runJob crash]', e.message);
  }));
  res.json({ jobId });
});

// ---- API Hợp đồng ----

// Lấy danh sách hợp đồng cũ
app.get('/api/contracts', (req, res) => {
  if (!TABLE_HD) return res.json([]);
  const search = (req.query.search || '').trim();
  const limit  = Math.min(parseInt(req.query.limit) || 30, 50);
  let qs = `limit=${limit}&sort=-Id`;
  if (search) {
    const s = encodeURIComponent(search);
    qs += `&where=(Ten_cong_ty,like,%25${s}%25)~or(So_hop_dong,like,%25${s}%25)`;
  }
  nocoGet(`/api/v1/db/data/noco/${NOCODB_BASE}/${TABLE_HD}?${qs}`, res);
});

// Tạo hợp đồng mới
app.post('/api/generate-contract', (req, res) => {
  const jobId = crypto.randomBytes(6).toString('hex');
  jobs[jobId] = { status: 'processing' };
  setTimeout(() => { delete jobs[jobId]; }, 3600000);
  setImmediate(() => runContractJob(jobId, req.body).catch(e => {
    if (jobs[jobId]) { jobs[jobId].status = 'error'; jobs[jobId].error = 'Unhandled: ' + e.message; }
    console.error('[runContractJob crash]', e.message);
  }));
  res.json({ jobId });
});

async function runContractJob(jobId, b) {
  const job = jobs[jobId];
  const fmt = n => Math.round(n).toLocaleString('vi-VN');

  // Hỗ trợ cả items array (multi-row) lẫn single-field (legacy)
  const items = Array.isArray(b.items) && b.items.length ? b.items : [{
    mo_ta:       b.mo_ta_san_pham || '',
    don_vi_tinh: b.don_vi_tinh || 'Cái',
    so_luong:    parseFloat(b.so_luong) || 0,
    don_gia:     parseFloat((b.don_gia || '').toString().replace(/\./g,'').replace(/,/g,'.')) || 0,
  }];

  // Tính tổng từ tất cả items
  let thanhTienTong = 0;
  items.forEach(it => {
    const sl  = parseFloat(it.so_luong) || 0;
    const gia = parseFloat((it.don_gia || '').toString().replace(/\./g,'').replace(/,/g,'.')) || 0;
    thanhTienTong += sl * gia;
  });
  const thueSuat = parseFloat(b.thue_suat) || 8;
  const thue   = Math.round(thanhTienTong * thueSuat / 100);
  const tongHD = thanhTienTong + thue;

  // Tổng hợp cho template (single-row template hiện tại)
  const firstItem  = items[0];
  const moTa       = items.length === 1 ? firstItem.mo_ta : items.map((it,i) => `${i+1}. ${it.mo_ta}`).join('\n');
  const donViTinh  = firstItem.don_vi_tinh || 'Cái';
  const soLuongSum = items.reduce((s, it) => s + (parseFloat(it.so_luong) || 0), 0);

  // Tính đơn giá cho trường hợp 1 sản phẩm
  const firstDonGia = parseFloat((firstItem.don_gia||'').toString().replace(/\./g,'').replace(/,/g,'.')) || 0;

  const fieldMap = {
    'SO-HOP-DONG':                                   b.so_hop_dong || '',
    'ngay-ky-hop-dong':                              formatDateVietnamese(b.ngay_ky_hop_dong || ''),
    'TEN-CONG-TY':                                   b.ten_cong_ty || '',
    'Dia-chi':                                       b.dia_chi || '',
    'Ma-so-thue':                                    b.ma_so_thue || '',
    'So-tai-khoan-ngan-hang':                        b.so_tai_khoan_ngan_hang || '',
    'Ten-nguoi-dai-dien':                            b.ten_nguoi_dai_dien || '',
    'Chuc-vu':                                       b.chuc_vu || '',
    'stt':                                           '1',
    'Mo-ta-san-pham':                                moTa,
    'Đon-vi-tinh':                                   donViTinh,
    'So-luong':                                      String(soLuongSum),
    'Don-gia':                                       fmt(firstDonGia),
    'Thanh-tien':                                    fmt(thanhTienTong),
    'Tong-thanh-tien':                               fmt(thanhTienTong),
    'Thue-gia-tri-gia-tang':                         fmt(thue),
    'Tong-gia-tri-hop-dong':                          fmt(tongHD),
    'TONG-GIA-TRI-HOP-DONG':                         fmt(tongHD),
    'So-tien-tong-gia-tring-hop-dong-bang-chu':      soTienBangChu(tongHD),
    'Thoi-gian-giao-hang':                           b.thoi_gian_giao_hang || '',
    'Dia-diem-giao-hang':                            b.dia_diem_giao_hang || '',
  };

  const CONTRACTS_DIR = path.join(__dirname, 'outputs', 'contracts');
  if (!fs.existsSync(CONTRACTS_DIR)) fs.mkdirSync(CONTRACTS_DIR, { recursive: true });

  const soSlug = (b.so_hop_dong || 'hop-dong').replace(/[\/\\:*?"<>|]/g, '-').trim();

  let renderedDocx;
  try {
    renderedDocx = renderContractTemplate(CONTRACT_TEMPLATE, fieldMap, items);
  } catch(e) {
    job.status = 'error'; job.error = 'Render error: ' + e.message; return;
  }

  // Export Word
  const outDocx = path.join(CONTRACTS_DIR, `${soSlug}.docx`);
  try { fs.copyFileSync(renderedDocx, outDocx); } catch(_) {}

  // Export PDF
  const tmpDocx = path.join(CONTRACTS_DIR, `_tmp_${jobId}.docx`);
  try { fs.copyFileSync(renderedDocx, tmpDocx); } catch(e) {
    job.status = 'error'; job.error = 'Copy docx error: ' + e.message; return;
  }
  try { fs.unlinkSync(renderedDocx); } catch(_) {}

  const outPdf = path.join(CONTRACTS_DIR, `${soSlug}.pdf`);
  const cmd = `${SOFFICE} --headless --convert-to pdf --outdir "${CONTRACTS_DIR}" "${tmpDocx}"`;
  exec(cmd, { timeout: 120000 }, (err2, stdout, stderr) => {
    try { fs.unlinkSync(tmpDocx); } catch(_) {}
    if (err2) { job.status = 'error'; job.error = 'LibreOffice: ' + err2.message; return; }

    const libreOut = path.join(CONTRACTS_DIR, `_tmp_${jobId}.pdf`);
    if (!fs.existsSync(libreOut)) { job.status = 'error'; job.error = 'PDF not found'; return; }
    try { fs.renameSync(libreOut, outPdf); } catch(_) {}

    job.status    = 'done';
    job.pdfUrl    = `/download-contract/${soSlug}.pdf`;
    job.docxUrl   = `/download-contract/${soSlug}.docx`;
    job.filename  = soSlug;
    console.log('✅ Contract ready:', soSlug);

    // Lưu NocoDB nếu có table ID
    if (TABLE_HD) {
      const record = {
        So_hop_dong: b.so_hop_dong || '', Ngay_ky: b.ngay_ky_hop_dong || '',
        Ten_cong_ty: b.ten_cong_ty || '', Dia_chi: b.dia_chi || '',
        Ma_so_thue: b.ma_so_thue || '', So_tai_khoan: b.so_tai_khoan_ngan_hang || '',
        Nguoi_dai_dien: b.ten_nguoi_dai_dien || '', Chuc_vu: b.chuc_vu || '',
        NV_ten: b.nv_ten || '', NV_bo_phan: b.nv_bo_phan || '', NV_email: b.nv_email || '', NV_sdt: b.nv_sdt || '',
        Mo_ta_san_pham: moTa, So_luong: soLuongSum,
        Don_gia: items.length === 1 ? (parseFloat((firstItem.don_gia||'').toString().replace(/\./g,'').replace(/,/g,'.')) || 0) : 0,
        Tong_gia_tri: tongHD,
        Thoi_gian_giao_hang: b.thoi_gian_giao_hang || '',
        Dia_diem_giao_hang: b.dia_diem_giao_hang || '',
      };
      Promise.resolve()
        .then(() => uploadPdfToNocoDB(outPdf, `${soSlug}.pdf`))
        .then(att => { if (att) record.File_PDF = [att]; return saveToNocoDB(TABLE_HD, record); })
        .catch(e => console.error('NocoDB contract error:', e.message));
    }
  });
}

function saveToNocoDB(tableId, record) {
  return new Promise((resolve, reject) => {
    const body = Buffer.from(JSON.stringify(record));
    const options = {
      hostname: NOCODB_HOST,
      path: `/api/v1/db/data/noco/${NOCODB_BASE}/${tableId}`,
      method: 'POST',
      headers: { 'xc-token': NOCODB_TOKEN, 'Content-Type': 'application/json', 'Content-Length': body.length }
    };
    const req = https.request(options, r => {
      let d = ''; r.on('data', c => d += c);
      r.on('end', () => { try { resolve(JSON.parse(d)); } catch(e) { reject(e); } });
    });
    req.on('error', reject); req.write(body); req.end();
  });
}

app.use('/download-contract', express.static(path.join(__dirname, 'outputs', 'contracts')));

// ============================================================
// AI CHAT
// ============================================================

// Rate limiting: 10 req/phút/IP
const rateLimitMap = new Map();
function checkRateLimit(ip) {
  const now = Date.now();
  const entry = rateLimitMap.get(ip) || { count: 0, reset: now + 60000 };
  if (now > entry.reset) { entry.count = 0; entry.reset = now + 60000; }
  entry.count++;
  rateLimitMap.set(ip, entry);
  return entry.count <= 10;
}
setInterval(() => {
  const now = Date.now();
  for (const [k, v] of rateLimitMap.entries()) { if (now > v.reset) rateLimitMap.delete(k); }
}, 120000);

// Load 20 tin nhắn gần nhất của session từ NocoDB
function loadChatHistory(sessionId) {
  return new Promise(resolve => {
    const qs = `limit=20&sort=-Id&where=(Session_id,eq,${encodeURIComponent(sessionId)})`;
    const req = https.get({
      hostname: NOCODB_HOST,
      path: `/api/v1/db/data/noco/${NOCODB_BASE}/${TABLE_CHAT}?${qs}`,
      headers: { 'xc-token': NOCODB_TOKEN }
    }, r => {
      let d = '';
      r.on('data', c => d += c);
      r.on('end', () => {
        try { resolve((JSON.parse(d).list || []).reverse()); } catch { resolve([]); }
      });
    });
    req.on('error', () => resolve([]));
    req.setTimeout(10000, () => { req.destroy(); resolve([]); });
  });
}

// Lưu tin nhắn vào NocoDB (async, không block)
function saveChatMessage(sessionId, role, content, activeTab) {
  saveToNocoDB(TABLE_CHAT, {
    Session_id: sessionId,
    Role: role,
    Content: content,
    Active_tab: activeTab || ''
  }).catch(e => console.error('[chat save error]', e.message));
}

// Tool definitions cho Claude
const chatTools = [
  {
    type: 'function',
    function: {
      name: 'query_quotes',
      description: 'Truy vấn danh sách báo giá từ hệ thống. Dùng để thống kê hoặc tìm báo giá cụ thể.',
      parameters: {
        type: 'object',
        properties: {
          search: { type: 'string', description: 'Tìm theo tên công ty hoặc số báo giá (bỏ trống để lấy tất cả)' },
          limit:  { type: 'number', description: 'Số kết quả tối đa (mặc định 20)' }
        }
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'query_contracts',
      description: 'Truy vấn danh sách hợp đồng từ hệ thống.',
      parameters: {
        type: 'object',
        properties: {
          search: { type: 'string', description: 'Tìm theo tên công ty hoặc số hợp đồng' },
          limit:  { type: 'number', description: 'Số kết quả tối đa (mặc định 20)' }
        }
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'query_employees',
      description: 'Lấy danh sách nhân viên trong hệ thống.',
      parameters: { type: 'object', properties: {} }
    }
  },
  {
    type: 'function',
    function: {
      name: 'prefill_quote_form',
      description: 'Điền thông tin vào form báo giá. Gọi khi đã thu thập đủ thông tin từ user.',
      parameters: {
        type: 'object',
        properties: {
          ten_cong_ty:       { type: 'string', description: 'Tên công ty khách hàng' },
          ten_phong_ban:     { type: 'string', description: 'Tên phòng ban' },
          ten_nguoi_lien_he: { type: 'string', description: 'Tên người liên hệ' },
          email_khach_hang:  { type: 'string', description: 'Email khách hàng' },
          sdt_khach_hang:    { type: 'string', description: 'Số điện thoại' },
          ten_du_an:         { type: 'string', description: 'Tên dự án' },
          items: {
            type: 'array',
            description: 'Danh sách sản phẩm',
            items: {
              type: 'object',
              properties: {
                mo_ta:      { type: 'string', description: 'Mô tả sản phẩm' },
                so_luong:   { type: 'number', description: 'Số lượng' },
                don_gia:    { type: 'number', description: 'Đơn giá (VNĐ)' },
                chiet_khau: { type: 'number', description: 'Chiết khấu (%)' }
              }
            }
          }
        }
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'prefill_contract_form',
      description: 'Điền thông tin vào form hợp đồng. Gọi khi đã thu thập đủ thông tin từ user.',
      parameters: {
        type: 'object',
        properties: {
          so_hop_dong:             { type: 'string' },
          ngay_ky_hop_dong:        { type: 'string', description: 'Ngày ký (YYYY-MM-DD)' },
          ten_cong_ty:             { type: 'string' },
          dia_chi:                 { type: 'string' },
          ma_so_thue:              { type: 'string' },
          so_tai_khoan_ngan_hang:  { type: 'string' },
          ten_nguoi_dai_dien:      { type: 'string' },
          chuc_vu:                 { type: 'string' },
          thoi_gian_giao_hang:     { type: 'string' },
          dia_diem_giao_hang:      { type: 'string' },
          items: {
            type: 'array',
            items: {
              type: 'object',
              properties: {
                mo_ta:       { type: 'string' },
                don_vi_tinh: { type: 'string' },
                so_luong:    { type: 'number' },
                don_gia:     { type: 'number' }
              }
            }
          }
        }
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'switch_tab',
      description: 'Chuyển sang tab phù hợp với yêu cầu của user.',
      parameters: {
        type: 'object',
        required: ['tab'],
        properties: {
          tab: { type: 'string', enum: ['quote', 'contract'], description: '"quote" = Báo giá, "contract" = Hợp đồng' }
        }
      }
    }
  }
];

// Thực thi tool call
async function executeTool(name, args) {
  const queryNoco = (tableId, search, searchFields, limit) => new Promise(resolve => {
    const lim = Math.min(limit || 20, 50);
    let qs = `limit=${lim}&sort=-Id`;
    if (search) {
      const s = encodeURIComponent(search);
      const cond = searchFields.map(f => `(${f},like,%25${s}%25)`).join('~or');
      qs += `&where=${cond}`;
    }
    const req = https.get({
      hostname: NOCODB_HOST,
      path: `/api/v1/db/data/noco/${NOCODB_BASE}/${tableId}?${qs}`,
      headers: { 'xc-token': NOCODB_TOKEN }
    }, r => {
      let d = ''; r.on('data', c => d += c);
      r.on('end', () => { try { resolve(JSON.parse(d).list || []); } catch { resolve([]); } });
    });
    req.on('error', () => resolve([]));
    req.setTimeout(15000, () => { req.destroy(); resolve([]); });
  });

  if (name === 'query_quotes')    return queryNoco(TABLE_BG, args.search, ['Ten_cong_ty','So_bao_gia'], args.limit);
  if (name === 'query_contracts') return queryNoco(TABLE_HD, args.search, ['Ten_cong_ty','So_hop_dong'], args.limit);
  if (name === 'query_employees') return queryNoco(TABLE_NV, '', [], 50);
  // Client-side actions: chỉ cần acknowledge
  if (['prefill_quote_form','prefill_contract_form','switch_tab'].includes(name)) return { success: true };
  return { error: 'Unknown tool' };
}

// Build system prompt với context hiện tại
function buildSystemPrompt(activeTab, formContext) {
  const today = new Date().toLocaleDateString('vi-VN');
  const tabName = activeTab === 'contract' ? 'Hợp đồng' : 'Báo giá';

  let formStr = '';
  if (formContext && typeof formContext === 'object') {
    const filled = Object.entries(formContext)
      .filter(([, v]) => v && String(v).trim())
      .map(([k, v]) => `  - ${k}: ${v}`)
      .join('\n');
    if (filled) formStr = `\nForm đang điền dở:\n${filled}\n`;
  }

  return `Bạn là trợ lý AI của Elide Fire Vietnam — phân phối bóng chữa cháy tự động Elide Fire (Đan Mạch, bảo hành 5 năm).

Ngày: ${today} | Tab: ${tabName}
${formStr}
Sản phẩm: Techideas 1.4kg — 2.500.000đ | Lovingcare 0.4kg — 1.950.000đ | VAT 8%

QUY TẮC TRẢ LỜI (bắt buộc):
- Ngắn gọn. Chỉ trả lời đúng điều được hỏi, không diễn giải thêm nếu user không yêu cầu
- Dùng bullet point khi liệt kê. KHÔNG dùng bảng markdown
- Hạn chế emoji — chỉ dùng khi thực sự cần nhấn mạnh, không dùng ở đầu mỗi dòng
- Khi tra cứu dữ liệu: tóm tắt kết quả bằng bullet point, không paste nguyên bảng số liệu
- Tiếng Việt, thân thiện nhưng súc tích

CHỨC NĂNG:
- Tư vấn sản phẩm theo nhu cầu/diện tích
- Thu thập thông tin → điền form báo giá hoặc hợp đồng
- Tra cứu báo giá, hợp đồng, nhân viên

- Không hỏi lại trường đã có trong form
- Khi prefill: điền đủ mọi thông tin đã thu thập
- Nếu cần tab khác: gọi switch_tab trước khi prefill`;
}

// POST /api/chat — streaming SSE
app.post('/api/chat', async (req, res) => {
  const ip = (req.headers['x-forwarded-for'] || req.socket.remoteAddress || 'unknown').split(',')[0].trim();

  if (!checkRateLimit(ip)) {
    return res.status(429).json({ error: 'Quá nhiều yêu cầu, vui lòng chờ 1 phút rồi thử lại.' });
  }
  if (!OPENROUTER_API_KEY) {
    return res.status(503).json({ error: 'AI chưa được cấu hình. Vui lòng liên hệ admin.' });
  }

  const { message, sessionId, activeTab, formContext } = req.body;
  if (!message || !sessionId) return res.status(400).json({ error: 'Thiếu message hoặc sessionId' });

  // SSE setup
  res.setHeader('Content-Type', 'text/event-stream');
  res.setHeader('Cache-Control', 'no-cache');
  res.setHeader('Connection', 'keep-alive');
  res.flushHeaders();

  const send = (data) => { try { res.write(`data: ${JSON.stringify(data)}\n\n`); } catch (_) {} };

  // Lưu tin nhắn user ngay lập tức
  saveChatMessage(sessionId, 'user', message, activeTab);

  try {
    const history = await loadChatHistory(sessionId);
    const messages = [
      { role: 'system', content: buildSystemPrompt(activeTab, formContext || {}) },
      ...history.slice(-19).map(h => ({ role: h.Role, content: h.Content })),
      { role: 'user', content: message }
    ];

    let fullResponse = '';

    // Agentic loop: tối đa 3 vòng tool calls
    for (let iter = 0; iter < 3; iter++) {
      const stream = await openaiClient.chat.completions.create({
        model: CHAT_MODEL,
        messages,
        tools: chatTools,
        tool_choice: 'auto',
        max_tokens: 1500,
        stream: true
      });

      // Accumulate streaming response
      const toolCallMap = {};
      let textContent = '';
      let finishReason = null;

      for await (const chunk of stream) {
        const choice = chunk.choices[0];
        if (!choice) continue;
        finishReason = choice.finish_reason || finishReason;
        const delta = choice.delta;

        // Text streaming → gửi thẳng về client
        if (delta?.content) {
          textContent += delta.content;
          send({ type: 'text', content: delta.content });
        }

        // Accumulate tool calls (có thể nhiều chunks)
        if (delta?.tool_calls) {
          for (const tc of delta.tool_calls) {
            if (!toolCallMap[tc.index]) toolCallMap[tc.index] = { id: '', name: '', args: '' };
            if (tc.id)                   toolCallMap[tc.index].id   += tc.id;
            if (tc.function?.name)       toolCallMap[tc.index].name += tc.function.name;
            if (tc.function?.arguments)  toolCallMap[tc.index].args += tc.function.arguments;
          }
        }
      }

      const toolCalls = Object.values(toolCallMap);

      // Không có tool call → xong
      if (!toolCalls.length || finishReason === 'stop') {
        fullResponse = textContent;
        break;
      }

      // Có tool calls → execute
      messages.push({
        role: 'assistant',
        content: textContent || null,
        tool_calls: toolCalls.map(tc => ({
          id: tc.id, type: 'function',
          function: { name: tc.name, arguments: tc.args }
        }))
      });

      for (const tc of toolCalls) {
        let args = {};
        try { args = JSON.parse(tc.args || '{}'); } catch (_) {}

        send({ type: 'tool_start', name: tc.name });

        // Gửi action về client để xử lý phía frontend
        if (['prefill_quote_form', 'prefill_contract_form', 'switch_tab'].includes(tc.name)) {
          send({ type: 'action', name: tc.name, data: args });
        }

        const result = await executeTool(tc.name, args);
        messages.push({ role: 'tool', tool_call_id: tc.id, content: JSON.stringify(result) });
      }
      // Tiếp tục vòng lặp để lấy response text sau tool calls
    }

    send({ type: 'done' });
    res.end();

    if (fullResponse) saveChatMessage(sessionId, 'assistant', fullResponse, activeTab);

  } catch (e) {
    console.error('[chat error]', e.message);
    send({ type: 'error', message: 'Xin lỗi, tôi đang gặp sự cố. Vui lòng thử lại sau.' });
    send({ type: 'done' });
    res.end();
  }
});

// GET /api/chat/history/:sessionId — load lịch sử chat
app.get('/api/chat/history/:sessionId', async (req, res) => {
  try {
    const history = await loadChatHistory(req.params.sessionId);
    res.json(history);
  } catch (e) {
    res.json([]);
  }
});

app.listen(PORT, () => {
  console.log(`✅ Elide Fire Quote Server running on port ${PORT}`);
});
