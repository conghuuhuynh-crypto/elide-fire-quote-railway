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
const http     = require('http');
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

// Graceful shutdown on unhandled errors
process.on('uncaughtException', e => {
  console.error('[uncaughtException]', e.message);
  setTimeout(() => process.exit(1), 1000); // cho PM2 tự restart
});
process.on('unhandledRejection', e => {
  console.error('[unhandledRejection]', e);
});

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

// Auto-detect HTTP vs HTTPS for NocoDB (local = http, remote = https)
const nocoLib = NOCODB_HOST.startsWith('localhost') || NOCODB_HOST.startsWith('127.') ? http : https;
const nocoPort = NOCODB_HOST.startsWith('localhost') || NOCODB_HOST.startsWith('127.') ? parseInt(NOCODB_HOST.split(':')[1]) || 8080 : 443;
const nocoHostname = NOCODB_HOST.split(':')[0];

// AI Chat config
const { OpenAI } = require('openai');
const OPENROUTER_API_KEY = process.env.OPENROUTER_API_KEY || '';
const CHAT_MODEL = process.env.CHAT_MODEL || 'anthropic/claude-haiku-4.5'; // Chat bot báo giá — giữ Haiku để tiết kiệm
const openaiClient = new OpenAI({
  baseURL: 'https://openrouter.ai/api/v1',
  apiKey: OPENROUTER_API_KEY || 'sk-placeholder',
  defaultHeaders: {
    'HTTP-Referer': process.env.APP_URL || 'https://app.elidefire.com.vn',
    'X-Title': 'Elide Fire App'
  }
});
if (!OPENROUTER_API_KEY) console.warn('⚠️  OPENROUTER_API_KEY chưa được set — Chat AI sẽ không hoạt động');

// ── Claude Code CLI — Content & SEO Agent (chất lượng tương đương chạy /seo /content thật) ──
const ANTHROPIC_API_KEY  = process.env.ANTHROPIC_API_KEY || '';
const CLAUDE_PROJECT_DIR = process.env.CLAUDE_PROJECT_DIR || path.join(__dirname, '..'); // thư mục chứa CLAUDE.md
if (!ANTHROPIC_API_KEY)  console.warn('⚠️  ANTHROPIC_API_KEY chưa được set — Content/SEO Agent sẽ không hoạt động');
if (!fs.existsSync(path.join(CLAUDE_PROJECT_DIR, 'CLAUDE.md'))) console.warn('⚠️  CLAUDE.md không tìm thấy tại CLAUDE_PROJECT_DIR — Claude CLI thiếu context');

// ── Knowledge Base — đọc từ skill files thật lúc startup ────────────────────
const KNOWLEDGE_DIR = path.join(__dirname, 'knowledge');

function loadSkill(filename) {
  try {
    return fs.readFileSync(path.join(KNOWLEDGE_DIR, filename), 'utf8').trim();
  } catch(e) {
    console.warn(`⚠️  Không đọc được ${filename}:`, e.message);
    return '';
  }
}

const SKILL_SEO     = loadSkill('skill-seo.md');
const SKILL_CONTENT = loadSkill('skill-content.md');
const KB_BRAND      = loadSkill('products.md');

// Log để xác nhận skill files đã được đọc
console.log(`[CMS] skill-seo.md: ${SKILL_SEO.length} ký tự`);
console.log(`[CMS] skill-content.md: ${SKILL_CONTENT.length} ký tự`);
console.log(`[CMS] products.md: ${KB_BRAND.length} ký tự`);

const KB_KEYWORDS = 'B2C: bong chua chay gia dinh/xe oto/gia/hieu qua | B2B: nha xuong/tu dien/tu server/pccc tu dong | Brand: bong chua chay elide fire';
if (!SKILL_SEO)     console.warn('⚠️  skill-seo.md trống — SEO Agent sẽ thiếu context');
if (!SKILL_CONTENT) console.warn('⚠️  skill-content.md trống — Content Agent sẽ thiếu context');
if (!KB_BRAND)      console.warn('⚠️  products.md trống — AI sẽ thiếu thông tin sản phẩm');

// Cảnh báo sớm nếu thiếu biến bắt buộc
if (!NOCODB_TOKEN) console.warn('⚠️  NOCODB_TOKEN chưa được set — NocoDB calls sẽ thất bại');

const CONTRACT_TEMPLATE = path.join(__dirname, 'templates', 'contract-template-v2.docx');

// Job queue
const jobs = {};

// In-memory chat histories: sessionId → [{role, content}]
// Chỉ lưu user + assistant text (không lưu tool calls)
// Persist sang NocoDB sau mỗi turn — load lại khi server restart
const chatHistories = new Map();
const CHAT_HISTORY_MAX = 20;  // 10 lượt hội thoại
const CHAT_SESSIONS_MAX = 500; // giới hạn số session in-memory tránh leak

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

app.use(express.json({ limit: '25mb' }));
app.use(express.urlencoded({ extended: true, limit: '25mb' }));
app.use(express.static(path.join(__dirname, 'public')));
app.use('/assets', express.static(path.join(__dirname, 'assets')));
app.use('/download', express.static(QUOTES_DIR));

app.get('/', (req, res) => res.sendFile(path.join(__dirname, 'public', 'index.html')));
app.get('/health', (req, res) => res.json({ status: 'ok', version: 'v46-all-fields-verified' }));

// GET /admin/schema — scan field names thực tế từ NocoDB, so sánh với bot config
app.get('/admin/schema', async (req, res) => {
  const fetchFields = (tableId, label) => new Promise(resolve => {
    const req2 = nocoLib.get({
      hostname: nocoHostname, port: nocoPort,
      path: `/api/v1/db/data/noco/${NOCODB_BASE}/${tableId}?limit=1`,
      headers: { 'xc-token': NOCODB_TOKEN }
    }, r => {
      let d = ''; r.on('data', c => d += c);
      r.on('end', () => {
        try {
          const row = (JSON.parse(d).list || [])[0] || {};
          const fields = Object.keys(row).filter(k => !['Id','CreatedAt','UpdatedAt'].includes(k));
          resolve({ label, tableId, fields, error: null });
        } catch (e) { resolve({ label, tableId, fields: [], error: e.message }); }
      });
    });
    req2.on('error', e => resolve({ label, tableId, fields: [], error: e.message }));
    req2.setTimeout(10000, () => { req2.destroy(); resolve({ label, tableId, fields: [], error: 'timeout' }); });
  });

  const [bg, hd, nv] = await Promise.all([
    fetchFields(TABLE_BG, 'Bao_gia'),
    fetchFields(TABLE_HD, 'Hop_dong'),
    fetchFields(TABLE_NV, 'Nhan_vien'),
  ]);

  // Field names bot đang dùng (hardcoded trong formatToolResult + system prompt)
  const botConfig = {
    Bao_gia:   ['So_bao_gia','Ngay_bao_gia','Phien_ban','Ten_cong_ty','Ten_du_an','Nguoi_lien_he','SDT_khach_hang','Email_khach_hang','Phong_ban_KH','SL_Techideas','DonGia_Techideas','ThanhTien_Techideas','SL_Lovingcare','DonGia_Lovingcare','ThanhTien_Lovingcare','CK_Tong_don','Tong_thanh_toan','NV_ten','NV_bo_phan','NV_sdt','NV_email'],
    Hop_dong:  ['So_hop_dong','Ten_cong_ty','Ngay_ky','Nguoi_dai_dien','Chuc_vu','Dia_chi','Ma_so_thue','So_tai_khoan','Mo_ta_san_pham','So_luong','Don_gia','Tong_gia_tri','Thoi_gian_giao_hang','Dia_diem_giao_hang','NV_ten','NV_bo_phan','NV_sdt','NV_email'],
    Nhan_vien: ['Ten_nhan_vien','Bo_phan','So_dien_thoai','Email'],
  };

  const diff = (actual, expected) => ({
    ok:      expected.filter(f => actual.includes(f)),
    missing: expected.filter(f => !actual.includes(f)),  // bot dùng nhưng NocoDB không có
    extra:   actual.filter(f => !expected.includes(f)),  // NocoDB có nhưng bot chưa dùng
  });

  res.json({
    Bao_gia:   { ...bg,   diff: diff(bg.fields,   botConfig.Bao_gia)   },
    Hop_dong:  { ...hd,   diff: diff(hd.fields,   botConfig.Hop_dong)  },
    Nhan_vien: { ...nv,   diff: diff(nv.fields,   botConfig.Nhan_vien) },
  });
});

// Helper: NocoDB GET với timeout
function nocoGet(path, res) {
  const options = { hostname: nocoHostname, port: nocoPort, path, headers: { 'xc-token': NOCODB_TOKEN } };
  const req = nocoLib.get(options, r => {
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
  nocoGet(`/api/v1/db/data/noco/${NOCODB_BASE}/${TABLE_SP}?limit=200`, res);
});

// API danh sách báo giá cũ (để nạp lại form)
app.get('/api/quotes', (req, res) => {
  const search = (req.query.search || '').trim();
  const limit  = Math.min(parseInt(req.query.limit) || 30, 50);
  let qs = `limit=${limit}&sort=-so_bao_gia`;
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
        hostname: nocoHostname, port: nocoPort,
        path: `/api/v1/db/storage/upload?path=noco/${NOCODB_BASE}/Bao_gia/File_PDF`,
        method: 'POST',
        headers: { ...form.getHeaders(), 'xc-token': NOCODB_TOKEN }
      };
      const req = nocoLib.request(options, r => {
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
      hostname: nocoHostname, port: nocoPort,
      path: `/api/v1/db/data/noco/${NOCODB_BASE}/${TABLE_BG}`,
      method: 'POST',
      headers: { 'xc-token': NOCODB_TOKEN, 'Content-Type': 'application/json', 'Content-Length': body.length }
    };
    const req = nocoLib.request(options, r => {
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
      Items: JSON.stringify(validItems),
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
  let qs = `limit=${limit}&sort=-so_hop_dong`;
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
        Items: JSON.stringify(items),
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
      hostname: nocoHostname, port: nocoPort,
      path: `/api/v1/db/data/noco/${NOCODB_BASE}/${tableId}`,
      method: 'POST',
      headers: { 'xc-token': NOCODB_TOKEN, 'Content-Type': 'application/json', 'Content-Length': body.length }
    };
    const req = nocoLib.request(options, r => {
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

// Track NocoDB row IDs per session để tránh re-search
const sessionRowIds = new Map();

// PATCH một row NocoDB
function patchNocoDB(tableId, rowId, data) {
  return new Promise((resolve, reject) => {
    const body = Buffer.from(JSON.stringify(data));
    const opts = {
      hostname: nocoHostname, port: nocoPort,
      path: `/api/v1/db/data/noco/${NOCODB_BASE}/${tableId}/${rowId}`,
      method: 'PATCH',
      headers: { 'xc-token': NOCODB_TOKEN, 'Content-Type': 'application/json', 'Content-Length': body.length }
    };
    const req = nocoLib.request(opts, r => {
      let d = ''; r.on('data', c => d += c);
      r.on('end', () => { try { resolve(JSON.parse(d)); } catch { resolve({}); } });
    });
    req.on('error', reject);
    req.setTimeout(10000, () => { req.destroy(); reject(new Error('timeout')); });
    req.write(body); req.end();
  });
}

// Load session từ NocoDB — 1 row per session, Content = JSON array [{role, content}]
function loadChatSession(sessionId) {
  return new Promise(resolve => {
    const qs = `limit=1&where=(Session_id,eq,${encodeURIComponent(sessionId)})`;
    const req = nocoLib.get({
      hostname: nocoHostname, port: nocoPort,
      path: `/api/v1/db/data/noco/${NOCODB_BASE}/${TABLE_CHAT}?${qs}`,
      headers: { 'xc-token': NOCODB_TOKEN }
    }, r => {
      let d = ''; r.on('data', c => d += c);
      r.on('end', () => {
        try {
          const row = (JSON.parse(d).list || [])[0];
          if (!row) return resolve([]);
          if (row.Id) sessionRowIds.set(sessionId, row.Id);
          try { resolve(JSON.parse(row.Content || '[]')); } catch { resolve([]); }
        } catch { resolve([]); }
      });
    });
    req.on('error', () => resolve([]));
    req.setTimeout(10000, () => { req.destroy(); resolve([]); });
  });
}

// Upsert session: 1 row per session trong NocoDB (async, không block)
function upsertChatSession(sessionId, messages, activeTab) {
  const payload = { Content: JSON.stringify(messages), Active_tab: activeTab || '' };
  const rowId = sessionRowIds.get(sessionId);
  if (rowId) {
    patchNocoDB(TABLE_CHAT, rowId, payload).catch(e => console.error('[chat patch]', e.message));
    return;
  }
  // Lần đầu: search xem row đã tồn tại chưa
  const qs = `limit=1&where=(Session_id,eq,${encodeURIComponent(sessionId)})`;
  const req = nocoLib.get({
    hostname: nocoHostname, port: nocoPort,
    path: `/api/v1/db/data/noco/${NOCODB_BASE}/${TABLE_CHAT}?${qs}`,
    headers: { 'xc-token': NOCODB_TOKEN }
  }, r => {
    let d = ''; r.on('data', c => d += c);
    r.on('end', () => {
      try {
        const row = (JSON.parse(d).list || [])[0];
        if (row?.Id) {
          sessionRowIds.set(sessionId, row.Id);
          patchNocoDB(TABLE_CHAT, row.Id, payload).catch(e => console.error('[chat patch]', e.message));
        } else {
          saveToNocoDB(TABLE_CHAT, { Session_id: sessionId, Role: 'session', ...payload })
            .then(r => { if (r?.Id) sessionRowIds.set(sessionId, r.Id); })
            .catch(e => console.error('[chat create]', e.message));
        }
      } catch (e) { console.error('[chat upsert]', e.message); }
    });
  });
  req.on('error', e => console.error('[chat upsert get]', e.message));
  req.setTimeout(10000, () => req.destroy());
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
          nv_ten:            { type: 'string', description: 'Tên nhân viên phụ trách — phải khớp chính xác với Ten_nhan_vien trong danh sách nhân viên' },
          ten_cong_ty:       { type: 'string', description: 'Tên công ty khách hàng' },
          ten_phong_ban:     { type: 'string', description: 'Tên phòng ban' },
          ten_nguoi_lien_he: { type: 'string', description: 'Tên người liên hệ' },
          email_khach_hang:  { type: 'string', description: 'Email khách hàng' },
          sdt_khach_hang:    { type: 'string', description: 'Số điện thoại khách hàng' },
          ten_du_an:         { type: 'string', description: 'Tên dự án' },
          items: {
            type: 'array',
            description: 'Danh sách sản phẩm',
            items: {
              type: 'object',
              properties: {
                model:      { type: 'string', description: 'Mã model sản phẩm: "TECHIDEAS" (1.4kg, 2.5tr) hoặc "LOVINGCARE" (0.4kg, 1.95tr)' },
                mo_ta:      { type: 'string', description: 'Mô tả thêm nếu có' },
                so_luong:   { type: 'number', description: 'Số lượng' },
                don_gia:    { type: 'number', description: 'Đơn giá (VNĐ) — bỏ trống để dùng giá mặc định' },
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
          nv_ten:                  { type: 'string', description: 'Tên nhân viên phụ trách — phải khớp chính xác với Ten_nhan_vien trong danh sách nhân viên' },
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
                model:       { type: 'string', description: '"TECHIDEAS" hoặc "LOVINGCARE"' },
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

// Validate & sanitize tool args trước khi execute
function validateArgs(toolName, raw) {
  switch (toolName) {
    case 'query_employees':
      return {}; // tool này không nhận args — force empty
    case 'query_quotes':
    case 'query_contracts':
      return {
        search: typeof raw.search === 'string' ? raw.search.trim().slice(0, 200) : '',
        limit:  Number.isInteger(raw.limit) ? Math.min(Math.max(raw.limit, 1), 50) : 20
      };
    case 'switch_tab':
      return { tab: ['quote', 'contract'].includes(raw.tab) ? raw.tab : 'quote' };
    case 'prefill_quote_form':
    case 'prefill_contract_form':
      // Đảm bảo items là array
      if (raw.items && !Array.isArray(raw.items)) raw.items = [];
      return raw;
    default:
      return raw;
  }
}

// Thực thi tool call
async function executeTool(name, args) {
  const queryNoco = (tableId, search, searchFields, exactFields, limit) => new Promise(resolve => {
    try {
      const lim = Math.min(limit || 20, 50);
      let qs = `limit=${lim}&sort=-so_bao_gia`;
      if (search) {
        const s = encodeURIComponent(search.trim()); // encode mọi ký tự kể cả tiếng Việt
        const conds = searchFields.map(f => {
          if (exactFields && exactFields.includes(f)) return `(${f},eq,${s})`;
          return `(${f},like,%25${s}%25)`;
        });
        qs += `&where=${conds.join('~or')}`;
      }
      const req = nocoLib.get({
        hostname: nocoHostname, port: nocoPort,
        path: `/api/v1/db/data/noco/${NOCODB_BASE}/${tableId}?${qs}`,
        headers: { 'xc-token': NOCODB_TOKEN }
      }, r => {
        let d = ''; r.on('data', c => d += c);
        r.on('end', () => { try { resolve(JSON.parse(d).list || []); } catch { resolve([]); } });
      });
      req.on('error', () => resolve([]));
      req.setTimeout(15000, () => { req.destroy(); resolve([]); });
    } catch (e) {
      console.error('[queryNoco]', e.message);
      resolve([]);
    }
  });

  if (name === 'query_quotes')    return queryNoco(TABLE_BG, args.search, ['Ten_cong_ty','So_bao_gia'], ['So_bao_gia'], args.limit);
  if (name === 'query_contracts') return queryNoco(TABLE_HD, args.search, ['Ten_cong_ty','So_hop_dong'], ['So_hop_dong'], args.limit);
  if (name === 'query_employees') return queryNoco(TABLE_NV, '', [], [], 50);
  // Client-side actions: chỉ cần acknowledge
  if (['prefill_quote_form','prefill_contract_form','switch_tab'].includes(name)) return { success: true };
  return { error: 'Unknown tool' };
}

// Format tool result — chống hallucination
function formatToolResult(toolName, result) {
  const STRICT = '\n[QUY TẮC: Chỉ dùng đúng dữ liệu trên. Không suy luận, không bổ sung, không bịa thêm.]';

  // Client-side actions: không cần format
  if (['prefill_quote_form','prefill_contract_form','switch_tab'].includes(toolName)) {
    return JSON.stringify({ success: true });
  }

  // Mảng rỗng
  if (Array.isArray(result) && result.length === 0) {
    return `KHÔNG TÌM THẤY DỮ LIỆU. Trả lời: "Không có dữ liệu phù hợp." Không được bịa thêm.`;
  }

  // query_employees — chỉ giữ fields cần thiết
  if (toolName === 'query_employees') {
    const clean = (Array.isArray(result) ? result : [result]).map(r => ({
      Id:             r.Id,
      Ten_nhan_vien:  r.Ten_nhan_vien || null,
      Bo_phan:        r.Bo_phan       || null,
      Email:          r.Email         || null,
      SDT:            r.So_dien_thoai || r.SDT || null  // So_dien_thoai là field thực tế trong NocoDB
    }));
    return `DANH SÁCH NHÂN VIÊN (${clean.length} người):\n` +
      clean.map(r => `- Id=${r.Id} | Tên: ${r.Ten_nhan_vien || 'N/A'} | Bộ phận: ${r.Bo_phan || 'N/A'} | SDT: ${r.SDT || 'không có'} | Email: ${r.Email || 'không có'}`).join('\n') +
      STRICT;
  }

  // query_quotes — toàn bộ fields từ NocoDB (verified by /admin/schema)
  if (toolName === 'query_quotes') {
    const list = Array.isArray(result) ? result : [result];
    const clean = list.map(r => ({
      So_bao_gia:          r.So_bao_gia          || null,
      Ngay_bao_gia:        r.Ngay_bao_gia        || null,
      Phien_ban:           r.Phien_ban           || null,
      Ten_cong_ty:         r.Ten_cong_ty         || null,
      Ten_du_an:           r.Ten_du_an           || null,
      Nguoi_lien_he:       r.Nguoi_lien_he       || null,
      SDT_khach_hang:      r.SDT_khach_hang      || null,
      Email_khach_hang:    r.Email_khach_hang    || null,
      Phong_ban_KH:        r.Phong_ban_KH        || null,
      SL_Techideas:        r.SL_Techideas        || null,
      DonGia_Techideas:    r.DonGia_Techideas    || null,
      ThanhTien_Techideas: r.ThanhTien_Techideas || null,
      SL_Lovingcare:       r.SL_Lovingcare       || null,
      DonGia_Lovingcare:   r.DonGia_Lovingcare   || null,
      ThanhTien_Lovingcare:r.ThanhTien_Lovingcare|| null,
      CK_Tong_don:         r.CK_Tong_don         || null,
      Tong_thanh_toan:     r.Tong_thanh_toan     || null,
      NV_ten:              r.NV_ten              || null,
      NV_bo_phan:          r.NV_bo_phan          || null,
      NV_sdt:              r.NV_sdt              || null,
      NV_email:            r.NV_email            || null,
    }));
    return `KẾT QUẢ BÁO GIÁ (${clean.length} bản ghi):\n` +
      JSON.stringify(clean, null, 2) + STRICT;
  }

  // query_contracts — toàn bộ fields từ NocoDB (verified by /admin/schema)
  if (toolName === 'query_contracts') {
    const list = Array.isArray(result) ? result : [result];
    const clean = list.map(r => ({
      So_hop_dong:         r.So_hop_dong         || null,
      Ten_cong_ty:         r.Ten_cong_ty         || null,
      Ngay_ky:             r.Ngay_ky             || null,
      Nguoi_dai_dien:      r.Nguoi_dai_dien      || null,
      Chuc_vu:             r.Chuc_vu             || null,
      Dia_chi:             r.Dia_chi             || null,
      Ma_so_thue:          r.Ma_so_thue          || null,
      So_tai_khoan:        r.So_tai_khoan        || null,
      Mo_ta_san_pham:      r.Mo_ta_san_pham      || null,
      So_luong:            r.So_luong            || null,
      Don_gia:             r.Don_gia             || null,
      Tong_gia_tri:        r.Tong_gia_tri        || null,
      Thoi_gian_giao_hang: r.Thoi_gian_giao_hang || null,
      Dia_diem_giao_hang:  r.Dia_diem_giao_hang  || null,
      NV_ten:              r.NV_ten              || null,
      NV_bo_phan:          r.NV_bo_phan          || null,
      NV_sdt:              r.NV_sdt              || null,
      NV_email:            r.NV_email            || null,
    }));
    return `KẾT QUẢ HỢP ĐỒNG (${clean.length} bản ghi):\n` +
      JSON.stringify(clean, null, 2) + STRICT;
  }

  // Default
  return JSON.stringify(result) + STRICT;
}

// Build system prompt với context hiện tại
function buildSystemPrompt(activeTab, formContext) {
  const today = new Date().toLocaleDateString('vi-VN');
  const tabName = activeTab === 'contract' ? 'Hợp đồng' : 'Báo giá';

  let formStr = '';
  if (formContext && typeof formContext === 'object') {
    const filled = Object.entries(formContext)
      .filter(([, v]) => v && String(v).trim())
      .map(([k, v]) => `  ${k}: ${v}`)
      .join('\n');
    if (filled) formStr = `\n[Form hiện tại — Tab ${tabName}]\n${filled}\n`;
  }

  return `Bạn là trợ lý AI nội bộ của Công ty Cổ phần Kỹ thuật Môi trường Tinh Tuệ — phân phối độc quyền bóng chữa cháy Elide Fire tại Việt Nam.
Ngày: ${today} | Tab: ${tabName}
${formStr}
== TOOLS & KHI NÀO DÙNG ==

| Tool | Dùng khi |
|---|---|
| query_quotes | Tra cứu báo giá — theo tên công ty hoặc số báo giá |
| query_contracts | Tra cứu hợp đồng — theo tên công ty hoặc số hợp đồng |
| query_employees | Lấy DS nhân viên — KHÔNG truyền args |
| prefill_quote_form | Điền form báo giá |
| prefill_contract_form | Điền form hợp đồng |
| switch_tab | Chuyển tab trước khi prefill tab kia |

QUY TẮC TOOL — BẮT BUỘC:
- Mọi thông tin (SĐT, email, số tiền, tên...) CHỈ lấy từ kết quả tool. KHÔNG TỰ NGHĨ RA.
- Hỏi SĐT/email/bộ phận nhân viên → gọi query_employees trước, so tên → trả đúng giá trị
- Hỏi về báo giá/hợp đồng cụ thể → gọi query_quotes/query_contracts trước
- Form có So_bao_gia hoặc So_hop_dong → CÓ THỂ đã lưu → gọi query để xác nhận, không tự kết luận
- Tool trả null → "không có dữ liệu". Tool trả rỗng → "Không tìm thấy trong hệ thống". KHÔNG suy diễn.

== VÍ DỤ ĐÚNG / SAI ==

[1] User hỏi SĐT nhân viên Huỳnh Công Hữu:
  ❌ SAI: Trả ngay "Số điện thoại là 0987654321"
  ✅ ĐÚNG: Gọi query_employees → tìm Ten_nhan_vien = "Huỳnh Công Hữu" → đọc SDT → trả đúng

[2] Form có So_bao_gia = "EF-2026-03-001", user hỏi "đã lưu chưa?":
  ❌ SAI: Kết luận "Báo giá chưa được lưu" từ form context
  ✅ ĐÚNG: Gọi query_quotes search="EF-2026-03-001" → có kết quả → "Đã lưu", rỗng → "Không tìm thấy"

[3] Form đã có data, user nói "đổi tên công ty thành ABC":
  ❌ SAI: Hỏi lại "Bạn có muốn cập nhật không?"
  ✅ ĐÚNG: Gọi prefill ngay, giữ tất cả field cũ + ten_cong_ty = "ABC" → "Đã cập nhật tên công ty"

[4] User nói "chọn nhân viên Nguyễn Văn A phụ trách":
  ❌ SAI: Tự đặt nv_ten = "Nguyễn Văn A" rồi prefill
  ✅ ĐÚNG: Gọi query_employees → xác nhận tên chính xác → prefill với tên lấy từ tool

[5] User hỏi "báo giá của Công ty ABC giá trị bao nhiêu?":
  ❌ SAI: Bịa số tiền
  ✅ ĐÚNG: Gọi query_quotes search="Công ty ABC" → đọc Tong_thanh_toan từ kết quả → trả lời

== PREFILL FORM ==

Trường hợp 1 — Form trống, user muốn tạo mới:
  Thu thập thông tin → khi đủ nói "Bạn xác nhận để tôi điền form?" → sau xác nhận mới gọi prefill

Trường hợp 2 — Form có data, user sửa field (từ khóa: "đổi/sửa/thay/xóa/cập nhật"):
  → Gọi prefill NGAY (không hỏi): giữ nguyên tất cả field cũ, chỉ đổi field được yêu cầu
  → Xóa field: truyền "" (chuỗi rỗng). Sau đó nói "Đã cập nhật [tên field]"

Mapping params:
  Báo giá: ten_cong_ty | ten_phong_ban | ten_nguoi_lien_he | sdt_khach_hang | email_khach_hang | ten_du_an | nv_ten | items[]
  Hợp đồng: ten_cong_ty | dia_chi | ma_so_thue | ten_nguoi_dai_dien | chuc_vu | so_hop_dong | nv_ten | items[]
  items[]: { model: "TECHIDEAS"|"LOVINGCARE", so_luong, don_gia }

Nhân viên: BẮT BUỘC gọi query_employees trước → lấy Ten_nhan_vien chính xác → truyền vào nv_ten

== SẢN PHẨM ==
- TECHIDEAS 1.4kg — 2.500.000đ — nhà xưởng, kho, nhà máy, tủ điện công nghiệp
- LOVINGCARE 0.4kg — 1.950.000đ — gia đình, xe ô tô, văn phòng, căn hộ
- Lắp cách nguồn lửa 20-30cm, tự kích hoạt 3-30 giây, dập 360°. Không cần vận hành. Tuổi thọ 5 năm.
- Chứng nhận ISO 9001, CE, EN615. Eureka Gold, WIPO Gold.

== SCHEMA ==
BÁO GIÁ: So_bao_gia | Ngay_bao_gia | Phien_ban | Ten_cong_ty | Ten_du_an | Nguoi_lien_he | SDT_khach_hang | Email_khach_hang | Phong_ban_KH | SL_Techideas | DonGia_Techideas | ThanhTien_Techideas | SL_Lovingcare | DonGia_Lovingcare | ThanhTien_Lovingcare | CK_Tong_don | Tong_thanh_toan | NV_ten | NV_bo_phan | NV_sdt | NV_email
HỢP ĐỒNG: So_hop_dong | Ten_cong_ty | Ngay_ky | Nguoi_dai_dien | Chuc_vu | Dia_chi | Ma_so_thue | So_tai_khoan | Mo_ta_san_pham | So_luong | Don_gia | Tong_gia_tri | Thoi_gian_giao_hang | Dia_diem_giao_hang | NV_ten | NV_bo_phan | NV_sdt | NV_email
NHÂN VIÊN: Id | Ten_nhan_vien | Bo_phan | So_dien_thoai | Email
  → NV_ten (báo giá/hợp đồng) = Ten_nhan_vien (nhân viên) — là cùng 1 người

== QUY TẮC TRẢ LỜI ==
- Ngắn gọn — chỉ trả lời đúng điều được hỏi
- Bullet point khi liệt kê, không dùng bảng markdown, không emoji đầu dòng
- Tóm tắt kết quả query, không paste nguyên JSON
- Không hỏi lại field đã có trong form context`;
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

  try {
    // In-memory history cho session hiện tại (reset khi redeploy — không dùng NocoDB history)
    const sessionHistory = chatHistories.get(sessionId) || [];
    const messages = [
      { role: 'system', content: buildSystemPrompt(activeTab, formContext || {}) },
      ...sessionHistory,
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
        temperature: 0,
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
        let rawArgs = {};
        try { rawArgs = JSON.parse(tc.args || '{}'); } catch (_) {}
        const args = validateArgs(tc.name, rawArgs);

        send({ type: 'tool_start', name: tc.name });

        // Gửi action về client để xử lý phía frontend
        if (['prefill_quote_form', 'prefill_contract_form', 'switch_tab'].includes(tc.name)) {
          send({ type: 'action', name: tc.name, data: args });
        }

        const t0 = Date.now();
        const result = await executeTool(tc.name, args);
        const count = Array.isArray(result) ? result.length + ' records' : (result?.error ? 'error' : 'ok');
        console.log(`[tool] ${tc.name} → ${count} (${Date.now() - t0}ms)`);

        messages.push({ role: 'tool', tool_call_id: tc.id, content: formatToolResult(tc.name, result) });
      }
      // Tiếp tục vòng lặp để lấy response text sau tool calls
    }

    send({ type: 'done' });
    res.end();

    if (fullResponse) {
      // Cập nhật in-memory history
      sessionHistory.push({ role: 'user', content: message });
      sessionHistory.push({ role: 'assistant', content: fullResponse });
      const updated = sessionHistory.slice(-CHAT_HISTORY_MAX);
      chatHistories.set(sessionId, updated);
      // Giới hạn số session in-memory (xóa entry cũ nhất khi vượt giới hạn)
      if (chatHistories.size > CHAT_SESSIONS_MAX) {
        const oldest = chatHistories.keys().next().value;
        chatHistories.delete(oldest);
      }
      // Lưu toàn bộ session vào 1 row NocoDB
      upsertChatSession(sessionId, updated, activeTab);
    }

  } catch (e) {
    console.error('[chat error]', e.message, e.stack?.split('\n')[1]);
    send({ type: 'error', message: 'Xin lỗi, tôi đang gặp sự cố. Vui lòng thử lại sau. (' + e.message.slice(0, 80) + ')' });
    send({ type: 'done' });
    res.end();
  }
});

// GET /api/chat/history/:sessionId — load lịch sử chat (1 session = 1 row)
app.get('/api/chat/history/:sessionId', async (req, res) => {
  try {
    // Nếu session đang có trong memory → dùng luôn (tránh gọi NocoDB)
    const inMem = chatHistories.get(req.params.sessionId);
    if (inMem && inMem.length) return res.json(inMem);
    const history = await loadChatSession(req.params.sessionId);
    // Restore vào memory nếu load được
    if (history.length) chatHistories.set(req.params.sessionId, history);
    res.json(history);
  } catch (e) {
    res.json([]);
  }
});

// ═══════════════════════════════════════════════════════════════════════════
// CMS — Content Publisher
// ═══════════════════════════════════════════════════════════════════════════

const CMS_ROOT    = path.join(__dirname, '..');
const CMS_CONTENT = path.join(CMS_ROOT, 'outputs', 'content');
const CMS_IMAGES  = path.join(CMS_ROOT, 'outputs', 'images');
const WP_DOMAIN   = process.env.WP_DOMAIN || 'elidefire.com.vn';
const WP_APP_PASS = process.env.WP_APP_PASS || '';
const WP_AUTH     = Buffer.from('admin.tech@tinhtue.vn:' + WP_APP_PASS).toString('base64');
if (!WP_APP_PASS) console.warn('⚠️  WP_APP_PASS chua duoc set -- CMS publish se khong hoat dong');

// ── Anthropic SDK ── callClaude voi Prompt Caching ──
const Anthropic = require('@anthropic-ai/sdk');

// callClaude: system prompt (cached) + user message
// Cache TTL = 5 phut — cac call tiep theo chi ton 10% chi phi cho system prompt
async function callClaude(systemText, userText) {
  if (!ANTHROPIC_API_KEY) throw new Error('ANTHROPIC_API_KEY chua duoc set');
  const client = new Anthropic({ apiKey: ANTHROPIC_API_KEY });
  const msg = await client.messages.create({
    model: 'claude-sonnet-4-5',
    max_tokens: 8192,
    system: [{
      type: 'text',
      text: systemText,
      cache_control: { type: 'ephemeral' }  // Cache system prompt 5 phut
    }],
    messages: [{ role: 'user', content: userText }]
  });
  const text = msg.content?.[0]?.text?.trim();
  if (!text) throw new Error('Anthropic API khong tra ve noi dung');
  return text;
}

// Backward compat — dung cho cac route cu
async function runClaudeAgent(prompt) {
  return callClaude('Ban la AI assistant chuyen nghiep cua Elide Fire Vietnam.', prompt);
}


function cmsParseFrontmatter(raw) {
  const match = raw.match(/^---\r?\n([\s\S]*?)\r?\n---/);
  if (!match) return { meta: {}, body: raw };
  const meta = {};
  match[1].split(/\r?\n/).forEach(line => {
    const c = line.indexOf(':');
    if (c < 0) return;
    meta[line.slice(0, c).trim()] = line.slice(c + 1).trim();
  });
  return { meta, body: raw.slice(match[0].length).trim() };
}

function cmsMdToHtml(md) {
  const scripts = [];
  md = md.replace(/<script[\s\S]*?<\/script>/gi, m => { scripts.push(m); return `%%S${scripts.length-1}%%`; });
  let html = md
    .replace(/^### (.+)$/gm, '<h3>$1</h3>').replace(/^## (.+)$/gm, '<h2>$1</h2>').replace(/^# (.+)$/gm, '<h1>$1</h1>')
    .replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>').replace(/\*(.+?)\*/g, '<em>$1</em>')
    .replace(/^> (.+)$/gm, '<blockquote><p>$1</p></blockquote>').replace(/^---$/gm, '<hr>')
    .replace(/!\[([^\]]*)\]\(([^)]+)\)/g, '<img src="$2" alt="$1" style="max-width:100%">')
    .replace(/\[([^\]]+)\]\(([^)]+)\)/g, '<a href="$2">$1</a>');
  html = html.replace(/(\|.+\|\r?\n)+/g, table => {
    const rows = table.trim().split(/\r?\n/).filter(r => !r.match(/^\|[-| :]+\|$/));
    return '<table class="wp-table">\n' + rows.map((row, i) => {
      const cells = row.split('|').slice(1, -1), tag = i === 0 ? 'th' : 'td';
      return '<tr>' + cells.map(c => `<${tag}>${c.trim()}</${tag}>`).join('') + '</tr>';
    }).join('\n') + '\n</table>';
  });
  html = html.replace(/(^- .+$\r?\n?)+/gm, block =>
    '<ul>' + block.trim().split(/\r?\n/).map(l => '<li>' + l.replace(/^- /, '') + '</li>').join('') + '</ul>\n'
  );
  html = html.split(/\r?\n\r?\n/).map(p => {
    p = p.trim(); if (!p) return '';
    if (p.match(/^<(h[1-6]|ul|ol|blockquote|hr|table|script|img|%%S)/)) return p;
    return '<p>' + p.replace(/\r?\n/g, '<br>') + '</p>';
  }).join('\n');
  scripts.forEach((s, i) => { html = html.replace(`%%S${i}%%`, s); });
  return html;
}

function cmsWpRequest(method, endpoint, body, extraHeaders = {}) {
  return new Promise((resolve, reject) => {
    const isBuffer = Buffer.isBuffer(body);
    const data = isBuffer ? body : (body ? JSON.stringify(body) : undefined);
    const headers = { 'Authorization': 'Basic ' + WP_AUTH, ...extraHeaders };
    if (!isBuffer && data) { headers['Content-Type'] = headers['Content-Type'] || 'application/json'; headers['Content-Length'] = Buffer.byteLength(data); }
    const req = https.request({ hostname: WP_DOMAIN, path: endpoint, method, headers }, res => {
      let d = ''; res.on('data', c => d += c);
      res.on('end', () => { try { resolve(JSON.parse(d)); } catch(e) { reject(new Error(d.slice(0, 300))); } });
    });
    req.setTimeout(25000, () => { req.destroy(new Error('WP API timeout (25s)')); });
    req.on('error', reject);
    if (data) req.write(data);
    req.end();
  });
}

// Serve content-ui.html
app.get('/content', (req, res) => {
  res.sendFile(path.join(__dirname, '..', 'tools', 'content-ui.html'));
});

// Serve images from outputs/images/
app.get('/api/cms/img/:filename', (req, res) => {
  const imgPath = path.join(CMS_IMAGES, req.params.filename);
  if (!fs.existsSync(imgPath)) { res.status(404).end(); return; }
  const ext  = imgPath.split('.').pop().toLowerCase();
  const mime = { jpg:'image/jpeg', jpeg:'image/jpeg', png:'image/png', gif:'image/gif', webp:'image/webp' }[ext] || 'image/jpeg';
  res.setHeader('Content-Type', mime);
  res.end(fs.readFileSync(imgPath));
});

// List posts
app.get('/api/cms/posts', (req, res) => {
  if (!fs.existsSync(CMS_CONTENT)) { res.json([]); return; }
  const posts = fs.readdirSync(CMS_CONTENT).filter(f => f.endsWith('.md')).sort().reverse().map(f => {
    try {
      const { meta } = cmsParseFrontmatter(fs.readFileSync(path.join(CMS_CONTENT, f), 'utf8'));
      return { filename: f, title: meta['Meta Title'] || f, keyword: meta['Focus Keyword'] || '', slug: meta['URL Slug'] || '' };
    } catch { return { filename: f, title: f, keyword: '', slug: '' }; }
  });
  res.json(posts);
});

// List images
app.get('/api/cms/images', (req, res) => {
  if (!fs.existsSync(CMS_IMAGES)) { res.json([]); return; }
  res.json(fs.readdirSync(CMS_IMAGES).filter(f => /\.(jpg|jpeg|png|gif|webp)$/i.test(f)).sort().reverse());
});

// Get post content
app.get('/api/cms/post/:filename', (req, res) => {
  const filePath = path.join(CMS_CONTENT, req.params.filename);
  if (!fs.existsSync(filePath)) { res.status(404).json({ error: 'Not found' }); return; }
  const raw = fs.readFileSync(filePath, 'utf8');
  const { meta, body } = cmsParseFrontmatter(raw);
  const h1Match = body.match(/^# (.+)$/m);
  res.json({ meta, body, html: cmsMdToHtml(body.replace(/^# .+\n?/, '').trim()), raw, title: h1Match ? h1Match[1] : (meta['Meta Title'] || req.params.filename) });
});

// Save post
app.put('/api/cms/post/:filename', express.json({ limit: '5mb' }), (req, res) => {
  const filePath = path.join(CMS_CONTENT, req.params.filename);
  fs.writeFileSync(filePath, req.body.raw, 'utf8');
  res.json({ ok: true });
});

// AI Generate content
app.post('/api/cms/generate', express.json({ limit: '1mb' }), async (req, res) => {
  try {
    const { title, brief, keyword } = req.body;
    if (!OPENROUTER_API_KEY) throw new Error('OPENROUTER_API_KEY chưa được cấu hình');

    const prompt = `Bạn là chuyên gia content marketing cho bóng chữa cháy tự động Elide Fire (nhập khẩu từ Đan Mạch, phân phối độc quyền tại Việt Nam bởi Công ty Kỹ thuật Môi trường Tinh Tuệ).

Viết bài blog SEO đầy đủ bằng tiếng Việt theo thông tin sau:
${title ? `Tiêu đề: ${title}` : ''}
${keyword ? `Từ khóa SEO chính: ${keyword}` : ''}
${brief ? `Brief/Outline:\n${brief}` : ''}

Yêu cầu:
- Viết bằng Markdown, bắt đầu bằng # (H1) là tiêu đề bài
- Dùng ## và ### cho các phần và tiêu đề phụ
- Độ dài 800–1200 từ, đủ thông tin, đọc dễ hiểu
- Tự nhiên, chuyên nghiệp, thuyết phục — không sáo rỗng
- Lồng ghép tự nhiên từ khóa SEO (không nhồi nhét)
- Kết thúc bằng CTA kêu gọi liên hệ hoặc mua hàng
- Proof points có thể dùng: 145 quốc gia, 40 triệu người dùng, 9 giải thưởng quốc tế, tự kích hoạt trong 3–30 giây, 5 năm không bảo dưỡng, chứng nhận CE & ISO 9001:2015
- Giá tham khảo: Techideas 1.4kg: 2.500.000 VNĐ | Lovingcare 0.4kg: 1.950.000 VNĐ
- CHỈ xuất nội dung Markdown — không giải thích thêm`;

    const completion = await openaiClient.chat.completions.create({
      model: CHAT_MODEL,
      messages: [{ role: 'user', content: prompt }],
      max_tokens: 3000,
      temperature: 0.7
    });

    const content = completion.choices?.[0]?.message?.content || '';
    if (!content) throw new Error('AI không trả về nội dung');
    res.json({ content });
  } catch(e) {
    res.status(500).json({ error: e.message });
  }
});

// Publish to WordPress — hỗ trợ cả base64 (browser upload) và cms:filename (server folder)

// ── SEO Post-Processor — đảm bảo tiêu chí bằng code, không phụ thuộc AI ─────
// ── SEO Structural Fix (sync) — H2, first para, paragraph length ─────────────
function seoStructuralFix(html, keyword) {
  if (!keyword || !html) return html;
  const kw    = keyword.toLowerCase();
  const kwCap = keyword.charAt(0).toUpperCase() + keyword.slice(1);
  const stripT = h => h.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim();

  // 1. Keyword trong đoạn đầu
  const firstParaMatch = html.match(/<p>(.*?)<\/p>/s);
  if (firstParaMatch && !stripT(firstParaMatch[1]).toLowerCase().includes(kw)) {
    html = html.replace(firstParaMatch[0], `<p><strong>${kwCap}</strong> — ${firstParaMatch[1]}</p>`);
  }

  // 2. Keyword trong H2
  const h2list  = [...html.matchAll(/<h2[^>]*>(.*?)<\/h2>/gi)];
  const kwWords = kw.split(/\s+/).filter(w => w.length > 2);
  const hasKwH2 = h2list.some(m => kwWords.every(w => stripT(m[1]).toLowerCase().includes(w)));
  if (!hasKwH2 && h2list.length > 0) {
    const firstH2 = h2list[0];
    html = html.replace(firstH2[0], firstH2[0].replace(firstH2[1], `${kwCap}: ${firstH2[1]}`));
  }

  // 3. Break đoạn văn dài (> 500 ký tự)
  let safety = 0;
  while (safety++ < 20) {
    const m = /<p>([^<]{500,})<\/p>/.exec(html);
    if (!m) break;
    const mid = Math.floor(m[1].length / 2);
    const cut = m[1].indexOf('. ', mid);
    if (cut > 0 && cut < m[1].length - 30) {
      html = html.replace(m[0], `<p>${m[1].slice(0, cut + 1)}</p>\n<p>${m[1].slice(cut + 2)}</p>`);
    } else break;
  }

  return html;
}

// ── SEO Refinement Pass (async) — Claude tích hợp keyword tự nhiên ───────────
async function seoRefinementPass(html, keyword) {
  if (!keyword || !html || !ANTHROPIC_API_KEY) return html;

  const stripT  = h => h.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim();
  const kwEsc   = keyword.toLowerCase().replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const plain   = stripT(html).toLowerCase();
  const wc      = plain.split(/\s+/).filter(w => w.length > 0).length;
  const kwCount = (plain.match(new RegExp(kwEsc, 'g')) || []).length;
  const density = kwCount / wc * 100;

  // Gọi API nếu density < 0.8% — đảm bảo Rank Math pass 0.5%
  if (density >= 0.8) return html;

  const needed = Math.max(Math.ceil(wc * 0.008) - kwCount, 3); // target 0.8%
  console.log(`[SEO Refine] density=${density.toFixed(2)}% kwCount=${kwCount} wc=${wc} needed=${needed}`);

  try {
    const system = `Bạn là SEO editor chuyên tiếng Việt. Nhiệm vụ: chỉnh sửa HTML content để tăng tần suất từ khóa.
LUẬT BẮT BUỘC:
- Chỉ chỉnh SỬA câu văn có sẵn, KHÔNG thêm đoạn mới
- Giữ nguyên toàn bộ HTML tags, links, headings
- Keyword phải xuất hiện TỰ NHIÊN trong câu, không lặp máy móc
- Trả về TOÀN BỘ HTML đã chỉnh, không thêm gì khác ngoài HTML`;

    const userMsg = `Từ khóa: "${keyword}"
Hiện tại: ${kwCount} lần (${density.toFixed(1)}%) — cần thêm ~${needed} lần để đạt 0.6%

Tích hợp từ khóa tự nhiên vào nội dung bên dưới:
${html}`;

    const refined = await callClaude(system, userMsg);

    // Validate: refined phải dài hơn 50% so với input và chứa HTML tags
    if (refined && refined.length > html.length * 0.5 && refined.includes('<p>')) {
      // Strip markdown code blocks nếu Claude wrap trong ```html
      const cleaned = refined.replace(/^```html?\n?/i, '').replace(/\n?```$/,'').trim();
      const plainAfter = cleaned.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').toLowerCase();
      const kwAfter = (plainAfter.match(new RegExp(kwEsc, 'g')) || []).length;
      console.log('[SEO Refine] done: ' + kwCount + ' -> ' + kwAfter + ' occurrences');
      return cleaned;
    }
    console.warn('[SEO Refine] validation failed len=' + (refined && refined.length) + ' vs ' + html.length);
  } catch(e) {
    console.error('[SEO Refine] error:', e.message);
  }

  return html; // fallback: trả về HTML gốc nếu API lỗi
}

app.post('/api/cms/publish', express.json({ limit: '20mb' }), async (req, res) => {
  try {
    const {
      title, content, meta, status, date, slug, keyword, keywords, type,
      // Base64 upload từ browser
      featuredBase64, featuredName,
      contentImagesB64,           // [{index, name, base64}]
      // Legacy: đọc từ server folder
      featuredFilename
    } = req.body;
    if (!title || !content) throw new Error('Thiếu title hoặc content');

    // ── Plain text → HTML converter ──────────────────────────────────────────
    function plainToHtml(text) {
      // Replace image placeholders: [ảnh: upload:ID] → placeholder kept for later replacement
      // First handle upload:INDEX replacements after image uploads
      const blocks = text.split(/\n\n+/);
      return blocks.map(block => {
        block = block.trim();
        if (!block) return '';
        // Image placeholder line
        if (block.match(/^\[ảnh:\s*upload:\d+\]$/)) return block; // handled below
        // Short standalone line (heading candidate): all caps or < 80 chars, no sentence-ending punctuation
        if (!block.includes('\n') && block.length <= 80 && !/[.,;]$/.test(block) && block === block.toUpperCase() && block.length > 3) {
          return `<h2>${block}</h2>`;
        }
        // Multi-line list (lines starting with - or số.)
        const lines = block.split('\n');
        if (lines.length > 1 && lines.every(l => /^[-•\d]/.test(l.trim()))) {
          return '<ul>' + lines.map(l => `<li>${l.replace(/^[-•]\s*/, '').replace(/^\d+\.\s*/, '')}</li>`).join('') + '</ul>';
        }
        return '<p>' + block.replace(/\n/g, '<br>') + '</p>';
      }).filter(Boolean).join('\n\n');
    }

    // Categories: 94 = Tin tức, điều chỉnh theo loại bài
    const categories = type === 'product' ? [94, 95] : [94];
    const tags       = [105, 103, 116, 108, 104, 114];

    // ── Helper: parse data URI → {mime, buf} ─────────────────────────────────
    function parseDataUri(dataUri) {
      const m = dataUri.match(/^data:([^;]+);base64,(.+)$/);
      if (!m) throw new Error('Invalid data URI');
      return { mime: m[1], buf: Buffer.from(m[2], 'base64') };
    }

    // ── Upload content images (base64, từ browser) ───────────────────────────
    const uploadedByIndex = {}; // index → URL
    for (const img of (contentImagesB64 || [])) {
      try {
        const { mime, buf } = parseDataUri(img.base64);
        const media = await cmsWpRequest('POST', '/wp-json/wp/v2/media', buf, {
          'Content-Type': mime,
          'Content-Disposition': `attachment; filename="${img.name}"`,
          'Content-Length': buf.length
        });
        if (media.source_url) uploadedByIndex[img.index] = media.source_url;
      } catch(e2) { console.error('[CMS] Content img upload failed:', e2.message); }
    }

    // ── Upload content images (legacy: cms:filename từ server folder) ─────────
    const uploadedByName = {};
    const cmsImgRegex = /!\[([^\]]*)\]\(cms:([^)]+)\)/g;
    for (const m of [...content.matchAll(cmsImgRegex)]) {
      const filename = m[2].trim();
      if (uploadedByName[filename]) continue;
      try {
        const imgPath = path.join(CMS_IMAGES, filename);
        if (!fs.existsSync(imgPath)) { uploadedByName[filename] = null; continue; }
        const buf  = fs.readFileSync(imgPath);
        const ext  = filename.split('.').pop().toLowerCase();
        const mime = ext === 'png' ? 'image/png' : ext === 'gif' ? 'image/gif' : 'image/jpeg';
        const media = await cmsWpRequest('POST', '/wp-json/wp/v2/media', buf, {
          'Content-Type': mime,
          'Content-Disposition': `attachment; filename="${filename}"`,
          'Content-Length': buf.length
        });
        uploadedByName[filename] = media.source_url || null;
      } catch(e2) { uploadedByName[filename] = null; }
    }

    // ── Thay placeholders trong content ──────────────────────────────────────
    let processedContent = content
      // [ảnh: upload:INDEX] or [ảnh: upload:INDEX — filename.jpg] — plain text format
      .replace(/\[ảnh:\s*upload:(\d+)[^\]]*\]/g, (match, idx) => {
        const url = uploadedByIndex[parseInt(idx)];
        return url ? `<img src="${url}" alt="" style="max-width:100%;height:auto;border-radius:6px;margin:16px 0">` : '';
      })
      // upload:INDEX in markdown format (legacy)
      .replace(/!\[([^\]]*)\]\(upload:(\d+)\)/g, (match, alt, idx) => {
        const url = uploadedByIndex[parseInt(idx)];
        return url ? `<img src="${url}" alt="${alt}" style="max-width:100%">` : '';
      })
      // cms:filename (server folder legacy)
      .replace(cmsImgRegex, (match, alt, filename) => {
        const url = uploadedByName[filename.trim()];
        return url ? `![${alt}](${url})` : match;
      });

    // Convert plain text → HTML (nếu content không phải markdown)
    // Detect nếu content có markdown syntax
    const isMarkdown = /^#{1,3} |^\*\*|^- |\*\*[^*]+\*\*/m.test(processedContent);
    let htmlContent = isMarkdown
      ? cmsMdToHtml(processedContent.replace(/^# .+\n?/, '').trim())
      : plainToHtml(processedContent);

    // ── SEO Post-Process — deterministic boost trước auto-inject ──────────────
    // ── SEO: structural fix (sync) — áp dụng ngay trước publish ─────────────
    htmlContent = seoStructuralFix(htmlContent, keyword || '');
    // refinement pass (async) chạy background sau khi publish xong

    // ── Auto-inject SEO links nếu chưa có ────────────────────────────────────
    // Internal links — bắt buộc theo quy chuẩn SEO
    if (!/href=["'][^"']*elidefire\.com\.vn\/san-pham\//.test(htmlContent)) {
      htmlContent += '\n<p><strong>Xem sản phẩm phù hợp:</strong> <a href="https://elidefire.com.vn/san-pham/bong-chua-chay-elide-fire-lovingcare">Bóng chữa cháy Elide Fire LOVINGCARE 0.4kg</a> (cho gia đình, xe ô tô, văn phòng) | <a href="https://elidefire.com.vn/san-pham/bong-chua-chay-elide-fire-techideas">Bóng chữa cháy Elide Fire TECHIDEAS 1.4kg</a> (cho nhà xưởng, kho, công nghiệp).</p>';
    }
    // External dofollow link — nguồn uy tín PCCC
    if (!/href=["'][^"']*pccc\.gov\.vn/.test(htmlContent)) {
      htmlContent += '\n<p><em>Nguồn tham khảo: <a href="https://www.pccc.gov.vn">Cục Cảnh sát Phòng cháy, chữa cháy và Cứu nạn cứu hộ</a> — cơ quan quản lý nhà nước về PCCC tại Việt Nam.</em></p>';
    }

    // ── Upload ảnh đại diện ───────────────────────────────────────────────────
    let mediaId = 0;

    // Ưu tiên base64 từ browser
    if (featuredBase64) {
      try {
        const { mime, buf } = parseDataUri(featuredBase64);
        const filename = featuredName || 'featured.jpg';
        const media = await cmsWpRequest('POST', '/wp-json/wp/v2/media', buf, {
          'Content-Type': mime,
          'Content-Disposition': `attachment; filename="${filename}"`,
          'Content-Length': buf.length
        });
        if (media.id) {
          mediaId = media.id;
          await cmsWpRequest('POST', `/wp-json/wp/v2/media/${mediaId}`, { alt_text: keyword || title });
        }
      } catch(e2) { console.error('[CMS] Featured upload failed:', e2.message); }
    } else if (featuredFilename) {
      // Legacy: đọc từ server folder
      try {
        const imgPath = path.join(CMS_IMAGES, featuredFilename);
        if (fs.existsSync(imgPath)) {
          const buf  = fs.readFileSync(imgPath);
          const ext  = featuredFilename.split('.').pop().toLowerCase();
          const mime = ext === 'png' ? 'image/png' : ext === 'gif' ? 'image/gif' : 'image/jpeg';
          const media = await cmsWpRequest('POST', '/wp-json/wp/v2/media', buf, {
            'Content-Type': mime,
            'Content-Disposition': `attachment; filename="${featuredFilename}"`,
            'Content-Length': buf.length
          });
          if (media.id) {
            mediaId = media.id;
            await cmsWpRequest('POST', `/wp-json/wp/v2/media/${mediaId}`, { alt_text: keyword || title });
          }
        }
      } catch(e2) { console.error('[CMS] Featured upload failed:', e2.message); }
    }

    // ── Tạo post ─────────────────────────────────────────────────────────────
    const postPayload = {
      title, content: htmlContent,
      status: status || 'draft',
      slug:   toSlug(slug || title || ''),
      categories, tags,
      meta: {
        rank_math_focus_keyword: keyword || '',
        rank_math_title:         title,
        rank_math_description:   meta || ''
      }
    };
    if (date)    postPayload.date            = new Date(date).toISOString().replace(/\.\d{3}Z$/, '');
    if (mediaId) postPayload.featured_media  = mediaId;

    const post = await cmsWpRequest('POST', '/wp-json/wp/v2/posts', postPayload);
    if (!post.id) throw new Error(JSON.stringify(post).slice(0, 300));

    // Trả về ngay cho user — không chờ SEO refinement
    res.json({ postId: post.id, url: `https://${WP_DOMAIN}/?p=${post.id}`, slug: post.slug, status: post.status });

    // Background: SEO refinement pass (Anthropic API ~10-20s)
    seoRefinementPass(htmlContent, keyword || '').then(async refined => {
      if (refined && refined !== htmlContent) {
        try {
          await cmsWpRequest('POST', `/wp-json/wp/v2/posts/${post.id}`, { content: refined });
          console.log(`[CMS] SEO refinement applied to post ${post.id}`);
        } catch(e) { console.error('[CMS] SEO refinement update failed:', e.message); }
      }
    }).catch(e => console.error('[CMS] SEO refinement bg error:', e.message));
  } catch(e) {
    res.status(500).json({ error: e.message });
  }
});

// Strip Markdown khỏi text — dùng cho outline + content trước khi trả về client
// ── Vietnamese slug helper — proper diacritics removal ──────────────────────
function toSlug(str) {
  if (!str) return '';
  return str
    .normalize('NFD')
    .replace(/[̀-ͯ]/g, '')
    .replace(/[đĐ]/g, d => d === 'đ' ? 'd' : 'D')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '');
}

function stripMarkdown(text) {
  if (!text) return text;
  return text
    .replace(/^#{1,6}\s+(.+)$/gm, (_, t) => t.toUpperCase()) // ## Heading → HEADING
    .replace(/\*\*(.+?)\*\*/gs, '$1')     // **bold** → bold
    .replace(/\*(.+?)\*/gs, '$1')         // *italic* → italic
    .replace(/`{3}[\s\S]*?`{3}/g, '')     // ```code block``` → xóa
    .replace(/`(.+?)`/g, '$1')            // `inline code` → plain
    .replace(/\[(.+?)\]\(.+?\)/g, '$1')   // [link](url) → link text
    .replace(/^---+$/gm, '')              // --- horizontal rule → xóa
    .replace(/^>\s*/gm, '')              // > blockquote → xóa dấu
    .replace(/\n{3,}/g, '\n\n')          // 3+ dòng trống → 2 dòng
    .trim();
}

// ── Generate outline (SEO Agent) ─────────────────────────────────────────────
app.post('/api/cms/generate-outline', express.json({ limit: '1mb' }), async (req, res) => {
  try {
    const { topic, keyword, sourceUrls } = req.body;
    if (!topic) throw new Error('Thiếu chủ đề (topic)');

    const now = new Date();
    const currentDate = `tháng ${now.getMonth() + 1}/${now.getFullYear()}`;

    // Prompt gửi thẳng cho Claude CLI — CLI tự load CLAUDE.md + memory + context đầy đủ
    // System prompt: full SEO skill + products.md (cached)
    const outlineSystem = SKILL_SEO + '\n\n---\n\n' + KB_BRAND;

    const prompt = `NHIỆM VỤ: Tạo OUTLINE (khung bài) SEO cho bài blog Elide Fire Vietnam.

Chủ đề: ${topic}
${keyword ? `Từ khóa mục tiêu: ${keyword}` : ''}
${sourceUrls && sourceUrls.length ? `Nguồn tham khảo:\n${sourceUrls.map(u => '  ' + u).join('\n')}` : ''}
Ngày hiện tại: ${currentDate}

OUTLINE = KHUNG BÀI THÔI. Mỗi bullet tối đa 8 từ. Không viết câu đầy đủ.

Cấu trúc bắt buộc:
MỞ BÀI (100-150 từ): hook + từ khóa + freshness
H2 1: [tiêu đề câu hỏi] (200-250 từ) → 3 bullet <= 8 từ/bullet
H2 2: [tiêu đề câu hỏi] (200-250 từ) → 3 bullet <= 8 từ/bullet
H2 3: [tiêu đề câu hỏi] (200-250 từ) → 3 bullet <= 8 từ/bullet
H2 4: [tiêu đề câu hỏi] (200-250 từ) → 3 bullet <= 8 từ/bullet
H2 FAQ (150 từ) → 3 cặp Q/A
KẾT BÀI + CTA (75-100 từ)

TITLE không chứa năm. Tổng outline không quá 150 từ.

Trả về CHÍNH XÁC format sau — không thêm text nào khác:
TITLE: [tiêu đề H1 ≤65 ký tự, có từ khóa, ưu tiên chứa số (vd: 5 lý do, 3 bước, 7 cách...)]
META: [130-155 ký tự, từ khóa nguyên văn, có CTA]
SLUG: [slug sinh từ keyword, không phải title — chỉ keyword viết thường không dấu, nối bằng dấu gạch]
KEYWORD: [từ khóa chính 2-5 từ tiếng Việt có dấu]
OUTLINE:
[outline text thuần]`;

    const raw = await callClaude(outlineSystem, prompt);
    if (!raw) throw new Error('Claude không trả về kết quả');

    // Parse delimiter-based (không dùng JSON — tránh lỗi multiline string)
    const titleM   = raw.match(/^TITLE:\s*(.+)$/m);
    const metaM    = raw.match(/^META:\s*(.+)$/m);
    const slugM    = raw.match(/^SLUG:\s*(.+)$/m);
    const keywordM = raw.match(/^KEYWORD:\s*(.+)$/m);
    const outlineIdx = raw.indexOf('\nOUTLINE:');
    const outlineRaw = outlineIdx >= 0
      ? raw.slice(outlineIdx + '\nOUTLINE:'.length).trim()
      : raw;

    if (!titleM) throw new Error('AI không trả về đúng format — thử lại');

    res.json({
      title:   titleM[1].trim(),
      meta:    metaM?.[1].trim()    || '',
      slug:    toSlug(slugM?.[1].trim() || ''),
      keyword: keywordM?.[1].trim() || '',
      outline: stripMarkdown(outlineRaw)
    });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

// ── Generate full content from approved outline (Content Agent) ───────────────
app.post('/api/cms/generate-from-outline', express.json({ limit: '2mb' }), async (req, res) => {
  try {
    const { topic, keyword, outline } = req.body;
    if (!outline) throw new Error('Thiếu outline');

    const { title: outlineTitle, meta: outlineMeta, slug: outlineSlug } = req.body;

    // System prompt: full Content skill + products.md (cached)
    const contentSystem = SKILL_CONTENT + '\n\n---\n\n' + KB_BRAND;

    const prompt = `NHIỆM VỤ: Viết bài blog hoàn chỉnh từ outline đã được duyệt.

Từ khóa: ${keyword || '(xem outline)'}
${topic ? `Chủ đề: ${topic}` : ''}
${outlineTitle ? `Tiêu đề H1: ${outlineTitle}` : ''}

OUTLINE ĐÃ DUYỆT:
${outline}

QUY TẮC FORMAT OUTPUT BẮT BUỘC:
- Tiêu đề phần: dùng ## (H2), không dùng ###, không all-caps
- Nội dung: plain text, đoạn cách nhau 1 dòng trắng
- KHÔNG tự thêm phần ngoài outline đã duyệt

SEO BẮT BUỘC:
- Từ khóa "${keyword || 'từ khóa chính'}" NGUYÊN VĂN trong 100 từ đầu
- Từ khóa NGUYÊN VĂN trong ≥2 heading ## (không tách, không thay thế)
- Mật độ ~1%: bài 1.000 từ → từ khóa NGUYÊN VĂN ≥10 lần (KHÔNG chèn thêm từ vào giữa keyword)
- Dùng ≥3 proof points từ brand
- External link đến pccc.gov.vn
- Internal links đến trang sản phẩm

Độ dài: ~200-250 từ/phần — không cắt bớt

Trả về CHÍNH XÁC format sau — không thêm text nào khác:
TITLE: [tiêu đề H1, từ khóa ở đầu, ưu tiên chứa số (vd: 5 lý do, 3 bước, 7 cách...)]
META: [130-155 ký tự, từ khóa nguyên văn, có CTA]
SLUG: [slug-khong-dau]
CONTENT:
[toàn bộ nội dung text thuần]`;

    const raw = await callClaude(contentSystem, prompt);
    if (!raw) throw new Error('Claude không trả về nội dung');

    // Parse delimiter-based (tránh lỗi JSON với nội dung dài)
    const titleM   = raw.match(/^TITLE:\s*(.+)$/m);
    const metaM    = raw.match(/^META:\s*(.+)$/m);
    const slugM    = raw.match(/^SLUG:\s*(.+)$/m);
    const contentIdx = raw.indexOf('\nCONTENT:');
    const contentRaw = contentIdx >= 0 ? raw.slice(contentIdx + '\nCONTENT:'.length).trim() : raw;

    if (!titleM) throw new Error('AI không trả về đúng format — thử lại');

    res.json({
      title:   titleM[1].trim(),
      meta:    metaM?.[1].trim()    || outlineMeta  || '',
      slug:    toSlug(slugM?.[1].trim() || outlineSlug || ''),
      content: contentRaw
    });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

// ── Optimize SEO after publish ────────────────────────────────────────────────
app.post('/api/cms/optimize-seo', express.json({ limit: '2mb' }), async (req, res) => {
  try {
    const { postId, keyword, title, content, meta } = req.body;
    if (!postId) throw new Error('Thiếu postId');

    let metaTitle = title || '';
    let metaDesc  = meta  || '';

    // Luôn dùng callClaude để viết SEO title/desc với keyword NGUYÊN VĂN
    if (ANTHROPIC_API_KEY && keyword) {
      try {
        const seoPrompt = `Viết SEO title và meta description cho bài blog.

Keyword (NGUYÊN VĂN, KHÔNG thay đổi): "${keyword}"
Title bài: ${title || ''}

YÊU CẦU BẮT BUỘC:
- SEO title: ≤60 ký tự, chứa "${keyword}" NGUYÊN VĂN liên tục
- Meta description: 130-155 ký tự, chứa "${keyword}" NGUYÊN VĂN, có CTA

Chỉ trả về JSON (không thêm gì khác):
{"metaTitle": "...", "metaDesc": "..."}`;

        const aiRaw = await callClaude('Bạn là SEO copywriter. Chỉ trả về JSON theo yêu cầu, không thêm text nào khác.', seoPrompt);
        const jsonMatch = aiRaw.match(/\{[\s\S]*\}/);
        if (jsonMatch) {
          const parsed = JSON.parse(jsonMatch[0]);
          const kw = keyword.toLowerCase();
          if (parsed.metaTitle && parsed.metaTitle.toLowerCase().includes(kw)) metaTitle = parsed.metaTitle;
          if (parsed.metaDesc  && parsed.metaDesc.toLowerCase().includes(kw))  metaDesc  = parsed.metaDesc;
        }
      } catch(aiErr) { console.error('[SEO Opt] AI error:', aiErr.message); }
    }

    // Post-process: thêm số vào SEO title nếu chưa có (Rank Math cộng điểm)
    if (metaTitle && !/\d/.test(metaTitle) && keyword) {
      const numbers = ['5', '3', '7', '4', '6'];
      const prefixes = ['5 lý do ', '3 cách ', '7 lợi ích ', '4 bước ', '6 điều '];
      const pick = prefixes[Math.floor(Math.random() * prefixes.length)];
      // Chỉ thêm nếu title chưa đủ dài
      if (metaTitle.length + pick.length <= 62) {
        metaTitle = pick + metaTitle.charAt(0).toLowerCase() + metaTitle.slice(1);
      }
    }

    // Cập nhật Rank Math fields qua WP REST API
    const payload = {
      meta: {
        rank_math_focus_keyword: keyword || '',
        rank_math_title:         metaTitle,
        rank_math_description:   metaDesc
      }
    };
    const result = await cmsWpRequest('POST', `/wp-json/wp/v2/posts/${postId}`, payload);
    if (!result.id) throw new Error(JSON.stringify(result).slice(0, 200));
    res.json({ ok: true, postId: result.id, metaTitle, metaDesc });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

// ── Verify SEO — 14 tiêu chí Rank Math (đồng bộ với tools/verify-seo.js) ─────
app.get('/api/cms/verify-seo/:postId', async (req, res) => {
  try {
    const postId = req.params.postId;
    const post = await cmsWpRequest('GET', `/wp-json/wp/v2/posts/${postId}?context=edit`);
    if (!post.id) throw new Error('Post không tồn tại');

    const keyword  = (post.meta?.rank_math_focus_keyword || '').toLowerCase().trim();
    const seoTitle = (post.meta?.rank_math_title || '').toLowerCase();
    const seoDesc  = (post.meta?.rank_math_description || '').toLowerCase();
    const slug     = post.slug || '';
    const rawHtml  = post.content?.raw || '';
    const stripH   = h => h.replace(/<script[\s\S]*?<\/script>/gi,' ').replace(/<style[\s\S]*?<\/style>/gi,' ').replace(/<[^>]+>/g,' ').replace(/&nbsp;/g,' ').replace(/&[a-z]+;/g,' ').replace(/\s+/g,' ').trim();
    const plain    = stripH(rawHtml).toLowerCase();
    const wc       = plain.split(/\s+/).filter(w => w.length > 0).length;
    const mediaId  = post.featured_media;

    if (!keyword) return res.json({ passed: 0, total: 0, pct: 0, noKeyword: true, results: [] });

    // Fetch media alt text (nếu có featured image)
    let mediaAlt = '';
    if (mediaId) {
      try {
        const media = await cmsWpRequest('GET', `/wp-json/wp/v2/media/${mediaId}`);
        mediaAlt = (media.alt_text || '').toLowerCase();
      } catch(e2) { /* ignore — sẽ fail criterion 8 */ }
    }

    // Fetch tất cả post khác để check keyword uniqueness (tối đa 50)
    let otherPosts = [];
    try {
      const others = await cmsWpRequest('GET', `/wp-json/wp/v2/posts?per_page=50&status=publish,draft&exclude=${postId}`);
      if (Array.isArray(others)) otherPosts = others;
    } catch(e2) { /* ignore — sẽ skip criterion 14 */ }

    const results = [];
    const check = (label, ok, detail) => results.push({ label, ok, detail: detail || '' });

    // 1. Keyword trong SEO title
    check('Keyword trong SEO title', seoTitle.includes(keyword));
    // 2. Keyword trong meta description
    check('Keyword trong meta description', seoDesc.includes(keyword));
    // 3. Keyword trong URL slug (exact contiguous match — same as Rank Math)
    const kwSlug = keyword.normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/đ/gi,'d').replace(/[^a-z0-9\s]/gi,'').trim().replace(/\s+/g,'-').toLowerCase();
    check('Keyword trong URL slug', slug.includes(kwSlug));
    // 4. URL slug ≤ 75 ký tự
    check(`URL slug ≤75 ký tự (${slug.length} ký tự)`, slug.length <= 75);
    // 5. Keyword trong 10% đầu bài
    const kwPos = plain.indexOf(keyword);
    check('Keyword trong 10% đầu bài', kwPos >= 0 && kwPos <= Math.floor(plain.length * 0.1));
    // 6. Keyword tìm thấy trong nội dung
    const kwCount = (plain.match(new RegExp(keyword.replace(/[.*+?^${}()|[\]\\]/g,'\\$&'),'g')) || []).length;
    check(`Keyword trong nội dung (${kwCount} lần)`, kwCount >= 1);
    // 7. Độ dài nội dung ≥ 600 từ
    check(`Độ dài nội dung (${wc} từ)`, wc >= 600);
    // 8. Keyword trong H2/H3 (partial match — all keyword words must appear in heading)
    const headings = [...rawHtml.matchAll(/<h[23][^>]*>([\s\S]*?)<\/h[23]>/gi)].map(m => stripH(m[1]).toLowerCase());
    const kwWords  = keyword.split(/\s+/).filter(w => w.length > 2);
    check('Keyword trong H2/H3', headings.some(h => kwWords.every(w => h.includes(w))));
    // 9. Alt text ảnh đại diện chứa keyword
    if (mediaId) {
      check('Alt text ảnh đại diện chứa keyword', mediaAlt.includes(keyword), mediaAlt ? `"${mediaAlt.slice(0,60)}"` : 'Alt text đang trống');
    } else {
      check('Ảnh đại diện (featured image)', false, 'Chưa set featured image');
    }
    // 10. Mật độ từ khóa 0.5–2.5%
    const density    = wc > 0 ? (kwCount / wc * 100) : 0;
    const kwWordCount = keyword.split(/\s+/).length;
    const minDensity  = kwWordCount <= 2 ? 0.5 : kwWordCount <= 4 ? 0.3 : 0.2; // scale threshold by keyword length
    check(`Mật độ từ khóa ${density.toFixed(1)}% (${minDensity}–2.5%)`, density >= minDensity && density <= 2.5);
    // 11. External links ≥ 1
    const extLinks = [...rawHtml.matchAll(/href="(https?:\/\/(?!(?:www\.)?elidefire)[^"]+)"/g)].map(m => m[1]);
    check(`External links (${extLinks.length} link)`, extLinks.length >= 1);
    // 12. External links là dofollow (không có nofollow)
    const nofollowLinks = [...rawHtml.matchAll(/href="(https?:\/\/(?!(?:www\.)?elidefire)[^"]+)"[^>]*rel="[^"]*nofollow[^"]*"/g)];
    const dofollowOk = extLinks.length >= 1 && nofollowLinks.length === 0;
    check('External links là dofollow', dofollowOk, nofollowLinks.length ? `${nofollowLinks.length} nofollow link` : extLinks.length ? 'Tất cả dofollow ✓' : 'Không có external link');
    // 13. Internal links ≥ 1
    const intLinks = [...rawHtml.matchAll(/href="((?:https?:\/\/(?:www\.)?elidefire[^"]*|\/[^"#]+))"/g)].map(m => m[1]);
    check(`Internal links (${intLinks.length} link)`, intLinks.length >= 1);
    // 14. Keyword chưa dùng ở bài khác
    const duplicate = otherPosts.find(p => (p.meta?.rank_math_focus_keyword || '').toLowerCase().trim() === keyword);
    check('Keyword chưa dùng ở bài khác', !duplicate, duplicate ? `Trùng với post ID ${duplicate.id}` : 'Unique ✓');

    const passed = results.filter(r => r.ok).length;
    const total  = results.length;
    res.json({ passed, total, pct: Math.round(passed / total * 100), keyword, results });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.listen(PORT, () => {
  console.log(`✅ Elide Fire Quote Server running on port ${PORT}`);
});
