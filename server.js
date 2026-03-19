/**
 * server.js — Elide Fire Quote Server (Railway deployment)
 * Carbone open-source + LibreOffice → PDF không watermark
 */

const express  = require('express');
const carbone  = require('carbone');
const JSZip    = require('jszip');
const path     = require('path');
const fs       = require('fs');
const https    = require('https');
const FormData = require('form-data');
const { exec } = require('child_process');
const os       = require('os');
const crypto   = require('crypto');

const app  = express();
const PORT = process.env.PORT || 3333;

const TEMPLATE   = path.join(__dirname, 'templates', 'quote-template.docx');
const QUOTES_DIR = path.join(__dirname, 'outputs', 'quotes');
const SOFFICE    = process.platform === 'win32'
  ? '"C:\\Program Files\\LibreOffice\\program\\soffice.exe"'
  : 'soffice';

// NocoDB config
const NOCODB_HOST  = 'nocodb-production-4d61.up.railway.app';
const NOCODB_TOKEN = 'cDiEKkF4wmvUroENBM_LrZb6VXQ6K5MlKgzXS7bA';
const NOCODB_BASE  = 'p49wwa1uzmjtv1e';
const TABLE_NV     = 'mbxi5rjran05biu';   // Nhan_vien
const TABLE_BG     = 'mnfhtr9jysetk07';   // Bao_gia
const TABLE_SP     = 'm1isvr6ljrp2klj';   // San_pham

// Job queue — lưu trạng thái từng job trong bộ nhớ
const jobs = {};

// Helper: escape XML đặc biệt
function escXml(s) {
  return String(s || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

// Helper: convert text có \n thành Word runs với <w:br/>
function moTaToRuns(text, templateRun) {
  const rPrMatch = templateRun.match(/<w:rPr>[\s\S]*?<\/w:rPr>/);
  const rPr = rPrMatch ? rPrMatch[0] : '';
  const lines = String(text || '').split('\n');
  return lines.map((line, i) =>
    `<w:r>${rPr}<w:t xml:space="preserve">${escXml(line)}</w:t></w:r>` +
    (i < lines.length - 1 ? `<w:r>${rPr}<w:br/></w:r>` : '')
  ).join('');
}

// Pre-process template: thay thế items[i] rows bằng actual rows
async function expandTemplateItems(templatePath, items) {
  const templateBuf = fs.readFileSync(templatePath);
  if (!items || items.length === 0) return templateBuf;

  const zip = await JSZip.loadAsync(templateBuf);
  let xml = await zip.file('word/document.xml').async('string');

  // Tìm dòng template chứa {d.items[i]...}
  let searchPos = 0, rowStart = -1, rowEnd = -1;
  while (true) {
    const s = xml.indexOf('<w:tr ', searchPos);
    if (s === -1) break;
    const e = xml.indexOf('</w:tr>', s) + 7;
    if (xml.slice(s, e).includes('d.items[i]')) { rowStart = s; rowEnd = e; break; }
    searchPos = e;
  }
  if (rowStart === -1) return templateBuf; // không tìm thấy → trả template gốc

  const templateRow = xml.slice(rowStart, rowEnd);

  // Tìm run chứa mo_ta để xử lý newlines riêng
  const moTaRunRegex = /<w:r\b[^>]*>(?:(?!<\/w:r>)[\s\S])*?\{d\.items\[i\]\.mo_ta\}(?:(?!<\/w:r>)[\s\S])*?<\/w:r>/;

  const expandedRows = items.map(item => {
    let row = templateRow;
    // Xử lý mo_ta (có thể có \n)
    const moTaMatch = row.match(moTaRunRegex);
    if (moTaMatch) {
      row = row.replace(moTaRunRegex, moTaToRuns(item.mo_ta, moTaMatch[0]));
    } else {
      row = row.replace(/\{d\.items\[i\]\.mo_ta\}/g, escXml(item.mo_ta));
    }
    // Thay thế các field còn lại
    row = row.replace(/\{d\.items\[i\]\.stt\}/g,        escXml(item.stt));
    row = row.replace(/\{d\.items\[i\]\.so_luong\}/g,   escXml(item.so_luong));
    row = row.replace(/\{d\.items\[i\]\.don_gia\}/g,    escXml(item.don_gia));
    row = row.replace(/\{d\.items\[i\]\.thanh_tien\}/g, escXml(item.thanh_tien));
    return row;
  }).join('');

  xml = xml.slice(0, rowStart) + expandedRows + xml.slice(rowEnd);
  zip.file('word/document.xml', xml);
  return await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
}

app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use('/assets', express.static(path.join(__dirname, 'assets')));
app.use('/assets', express.static(path.join(__dirname, 'public', 'assets')));
app.use('/download', express.static(QUOTES_DIR));

// Serve form
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// Health check
app.get('/health', (req, res) => res.json({ status: 'ok' }));

// API lấy danh sách nhân viên
app.get('/api/employees', (req, res) => {
  const options = {
    hostname: NOCODB_HOST,
    path: `/api/v1/db/data/noco/${NOCODB_BASE}/${TABLE_NV}?limit=100`,
    headers: { 'xc-token': NOCODB_TOKEN }
  };
  https.get(options, r => {
    let d = '';
    r.on('data', c => d += c);
    r.on('end', () => {
      try { res.json(JSON.parse(d).list || []); }
      catch (e) { res.json([]); }
    });
  }).on('error', () => res.json([]));
});

// API lấy danh sách sản phẩm
app.get('/api/products', (req, res) => {
  const options = {
    hostname: NOCODB_HOST,
    path: `/api/v1/db/data/noco/${NOCODB_BASE}/${TABLE_SP}?limit=200&sort=Id`,
    headers: { 'xc-token': NOCODB_TOKEN }
  };
  https.get(options, r => {
    let d = '';
    r.on('data', c => d += c);
    r.on('end', () => {
      try { res.json(JSON.parse(d).list || []); }
      catch (e) { res.json([]); }
    });
  }).on('error', () => res.json([]));
});

// API kiểm tra trạng thái job
app.get('/api/job/:id', (req, res) => {
  const job = jobs[req.params.id];
  if (!job) return res.status(404).json({ status: 'not_found' });
  res.json(job);
});

// Helper: upload PDF lên NocoDB storage
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

// Helper: lưu record vào NocoDB Bao_gia
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
        try {
          const res = JSON.parse(d);
          if (res.Id) resolve(res.Id);
          else reject(new Error('No Id: ' + d));
        } catch (e) { reject(e); }
      });
    });
    req.on('error', reject);
    req.write(body);
    req.end();
  });
}

// Worker: chạy nền — tạo PDF + lưu NocoDB
function runJob(jobId, b) {
  const job = jobs[jobId];
  const fmt = (n) => n.toLocaleString('vi-VN');

  // Parse items array (từ JSON body)
  let rawItems = [];
  if (Array.isArray(b.items)) {
    rawItems = b.items;
  } else if (typeof b.items === 'string') {
    try { rawItems = JSON.parse(b.items); } catch (_) {}
  }

  // Chỉ lấy items có số lượng > 0
  const validItems = rawItems.filter(it => parseFloat(it.so_luong) > 0);

  // Tính thành tiền từng dòng + tổng
  const ckTong = parseFloat(b.chiet_khau_tong) || 0;
  let tongTruoCK = 0;

  const carboneItems = validItems.map((it, idx) => {
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
    items:             carboneItems,
    truoc_chiet_khau:  fmt(tongTruoCK),
    chiet_khau:        ckTong > 0 ? `${ckTong}%` : '0',
    tong_thanh_tien:   fmt(total),
  };

  const soSlug  = (b.so_bao_gia || 'bao-gia').replace(/[\/\\:*?"<>|]/g, '-').trim();
  const tmpDocx = path.join(os.tmpdir(), `${soSlug}-${jobId}.docx`);
  const outPdf  = path.join(QUOTES_DIR, `${soSlug}.pdf`);

  if (!fs.existsSync(QUOTES_DIR)) fs.mkdirSync(QUOTES_DIR, { recursive: true });

  // Pre-process: expand items rows trực tiếp trong XML (Carbone array không hỗ trợ Word table reliably)
  expandTemplateItems(TEMPLATE, carboneItems).then(templateBuf => {
    // Xóa items khỏi data để Carbone không cần xử lý
    const carboneData = { ...data };
    delete carboneData.items;

    carbone.render(templateBuf, carboneData, {}, (err, result) => {
    if (err) { job.status = 'error'; job.error = err.message; return; }

    fs.writeFileSync(tmpDocx, result);

    const cmd = `${SOFFICE} --headless --convert-to pdf --outdir "${QUOTES_DIR}" "${tmpDocx}"`;
    exec(cmd, { timeout: 120000 }, async (err2) => {
      try { fs.unlinkSync(tmpDocx); } catch (_) {}

      if (err2) { job.status = 'error'; job.error = err2.message; return; }

      const tmpBasename = path.basename(tmpDocx, '.docx') + '.pdf';
      const libreOut    = path.join(QUOTES_DIR, tmpBasename);
      try { fs.renameSync(libreOut, outPdf); } catch (_) {}

      const finalPath = fs.existsSync(outPdf) ? outPdf : libreOut;
      const finalName = path.basename(finalPath);
      const appUrl    = 'https://elide-fire-quote-railway-production.up.railway.app';

      // Job done
      job.status   = 'done';
      job.url      = `${appUrl}/download/${finalName}`;
      job.filename = finalName;

      // Lưu NocoDB nền
      const record = {
        So_bao_gia:       b.so_bao_gia        || '',
        Ngay_bao_gia:     b.ngay_bao_gia      || '',
        Phien_ban:        b.phien_ban          || '',
        Ten_du_an:        b.ten_du_an          || '',
        Ten_cong_ty:      b.ten_cong_ty        || '',
        Phong_ban_KH:     b.ten_phong_ban      || '',
        Nguoi_lien_he:    b.ten_nguoi_lien_he  || '',
        SDT_khach_hang:   b.sdt_khach_hang     || '',
        Email_khach_hang: b.email_khach_hang   || '',
        NV_bo_phan:       b.nv_bo_phan         || '',
        NV_ten:           b.nv_ten             || '',
        NV_email:         b.nv_email           || '',
        NV_sdt:           b.nv_sdt             || '',
        Items_JSON:       JSON.stringify(validItems),
        CK_Tong_don:      ckTong,
        Tong_thanh_toan:  total,
      };
      try {
        const att = await uploadPdfToNocoDB(finalPath, finalName);
        if (att) record.File_PDF = [att];
        await saveQuoteToNocoDB(record);
        console.log('✅ NocoDB saved');
      } catch (e) { console.error('NocoDB error:', e.message); }
    });
  });
  }).catch(e => { job.status = 'error'; job.error = 'Template error: ' + e.message; });
}

// API xuất PDF — trả jobId ngay lập tức
app.post('/api/generate', (req, res) => {
  const jobId = crypto.randomBytes(6).toString('hex');
  jobs[jobId] = { status: 'processing' };

  // Dọn job cũ sau 1 giờ
  setTimeout(() => { delete jobs[jobId]; }, 3600000);

  // Chạy nền
  setImmediate(() => runJob(jobId, req.body));

  res.json({ jobId });
});

app.listen(PORT, () => {
  console.log(`✅ Elide Fire Quote Server running on port ${PORT}`);
});
