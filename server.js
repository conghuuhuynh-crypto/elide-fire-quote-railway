/**
 * server.js — Elide Fire Quote Server (Railway deployment)
 * Carbone open-source + LibreOffice → PDF không watermark
 */

const express  = require('express');
const carbone  = require('carbone');
const AdmZip   = require('adm-zip');
const path     = require('path');
const fs       = require('fs');
const https    = require('https');
const FormData = require('form-data');
const { exec } = require('child_process');
const os       = require('os');
const crypto   = require('crypto');

const app  = express();
const PORT = process.env.PORT || 3333;

// Prevent crash on unhandled errors
process.on('uncaughtException',   e => console.error('[uncaughtException]',   e.message));
process.on('unhandledRejection',  e => console.error('[unhandledRejection]',  e));

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

// Job queue
const jobs = {};

// ---- Helpers ----

function escXml(s) {
  return String(s || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function moTaToRuns(text, templateRun) {
  const rPrMatch = templateRun.match(/<w:rPr>[\s\S]*?<\/w:rPr>/);
  const rPr = rPrMatch ? rPrMatch[0] : '';
  const lines = String(text || '').split('\n');
  return lines.map((line, i) =>
    `<w:r>${rPr}<w:t xml:space="preserve">${escXml(line)}</w:t></w:r>` +
    (i < lines.length - 1 ? `<w:r>${rPr}<w:br/></w:r>` : '')
  ).join('');
}

// Pre-process template: expand items[] rows dùng adm-zip (pure JS)
function expandTemplateItems(templatePath, items) {
  if (!items || items.length === 0) return templatePath;

  const tmpDocx = path.join(os.tmpdir(), `tmpl_${crypto.randomBytes(4).toString('hex')}.docx`);

  try {
    fs.copyFileSync(templatePath, tmpDocx);

    const zip = new AdmZip(tmpDocx);
    const xmlEntry = zip.getEntry('word/document.xml');
    if (!xmlEntry) { fs.unlinkSync(tmpDocx); return templatePath; }

    let xml = xmlEntry.getData().toString('utf8');

    // Tìm dòng template chứa {d.items[i]...}
    let searchPos = 0, rowStart = -1, rowEnd = -1;
    while (true) {
      const s = xml.indexOf('<w:tr ', searchPos);
      if (s === -1) break;
      const e = xml.indexOf('</w:tr>', s) + 7;
      if (xml.slice(s, e).includes('d.items[i]')) { rowStart = s; rowEnd = e; break; }
      searchPos = e;
    }
    if (rowStart === -1) { fs.unlinkSync(tmpDocx); return templatePath; }

    const templateRow = xml.slice(rowStart, rowEnd);

    // Tìm run chứa mo_ta để xử lý xuống dòng
    const moTaPos = templateRow.indexOf('{d.items[i].mo_ta}');
    let moTaRunStart = -1, moTaRunEnd = -1, moTaRun = '';
    if (moTaPos !== -1) {
      moTaRunStart = templateRow.lastIndexOf('<w:r', moTaPos);
      moTaRunEnd   = templateRow.indexOf('</w:r>', moTaPos) + 6;
      moTaRun      = templateRow.slice(moTaRunStart, moTaRunEnd);
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

    zip.updateFile('word/document.xml', Buffer.from(xml, 'utf8'));
    zip.writeZip(tmpDocx);

    return tmpDocx;
  } catch (e) {
    try { fs.unlinkSync(tmpDocx); } catch (_) {}
    throw e;
  }
}

app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use('/assets', express.static(path.join(__dirname, 'assets')));
app.use('/assets', express.static(path.join(__dirname, 'public', 'assets')));
app.use('/download', express.static(QUOTES_DIR));

app.get('/', (req, res) => res.sendFile(path.join(__dirname, 'public', 'index.html')));
app.get('/health', (req, res) => res.json({ status: 'ok', version: 'adm-zip-v3' }));

// API nhân viên
app.get('/api/employees', (req, res) => {
  const options = {
    hostname: NOCODB_HOST,
    path: `/api/v1/db/data/noco/${NOCODB_BASE}/${TABLE_NV}?limit=100`,
    headers: { 'xc-token': NOCODB_TOKEN }
  };
  https.get(options, r => {
    let d = '';
    r.on('data', c => d += c);
    r.on('end', () => { try { res.json(JSON.parse(d).list || []); } catch (e) { res.json([]); } });
  }).on('error', () => res.json([]));
});

// API sản phẩm
app.get('/api/products', (req, res) => {
  const options = {
    hostname: NOCODB_HOST,
    path: `/api/v1/db/data/noco/${NOCODB_BASE}/${TABLE_SP}?limit=200&sort=Id`,
    headers: { 'xc-token': NOCODB_TOKEN }
  };
  https.get(options, r => {
    let d = '';
    r.on('data', c => d += c);
    r.on('end', () => { try { res.json(JSON.parse(d).list || []); } catch (e) { res.json([]); } });
  }).on('error', () => res.json([]));
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
    truoc_chiet_khau:  fmt(tongTruoCK),
    chiet_khau:        ckTong > 0 ? `${ckTong}%` : '0',
    tong_thanh_tien:   fmt(total),
  };

  const soSlug  = (b.so_bao_gia || 'bao-gia').replace(/[\/\\:*?"<>|]/g, '-').trim();
  const tmpDocx = path.join(os.tmpdir(), `${soSlug}-${jobId}.docx`);
  const outPdf  = path.join(QUOTES_DIR, `${soSlug}.pdf`);

  if (!fs.existsSync(QUOTES_DIR)) fs.mkdirSync(QUOTES_DIR, { recursive: true });

  let patchedTemplatePath = TEMPLATE;

  try {
    patchedTemplatePath = expandTemplateItems(TEMPLATE, carboneItems);
  } catch (e) {
    job.status = 'error';
    job.error  = 'Template expand error: ' + e.message;
    return;
  }

  carbone.render(patchedTemplatePath, data, {}, (err, result) => {
    // Dọn tmp file sau khi Carbone đọc xong
    if (patchedTemplatePath !== TEMPLATE) {
      try { fs.unlinkSync(patchedTemplatePath); } catch (_) {}
    }

    if (err) { job.status = 'error'; job.error = 'Carbone: ' + err.message; return; }

    try { fs.writeFileSync(tmpDocx, result); } catch (e) {
      job.status = 'error'; job.error = 'Write docx: ' + e.message; return;
    }

    const cmd = `${SOFFICE} --headless --convert-to pdf --outdir "${QUOTES_DIR}" "${tmpDocx}"`;
    exec(cmd, { timeout: 120000 }, (err2) => {
      try { fs.unlinkSync(tmpDocx); } catch (_) {}
      if (err2) { job.status = 'error'; job.error = 'LibreOffice: ' + err2.message; return; }

      const tmpBasename = path.basename(tmpDocx, '.docx') + '.pdf';
      const libreOut    = path.join(QUOTES_DIR, tmpBasename);
      try { fs.renameSync(libreOut, outPdf); } catch (_) {}

      const finalPath = fs.existsSync(outPdf) ? outPdf : libreOut;
      const finalName = path.basename(finalPath);
      const appUrl    = 'https://elide-fire-quote-railway-production.up.railway.app';

      job.status   = 'done';
      job.url      = `${appUrl}/download/${finalName}`;
      job.filename = finalName;

      const record = {
        So_bao_gia: b.so_bao_gia || '', Ngay_bao_gia: b.ngay_bao_gia || '',
        Phien_ban: b.phien_ban || '', Ten_du_an: b.ten_du_an || '',
        Ten_cong_ty: b.ten_cong_ty || '', Phong_ban_KH: b.ten_phong_ban || '',
        Nguoi_lien_he: b.ten_nguoi_lien_he || '', SDT_khach_hang: b.sdt_khach_hang || '',
        Email_khach_hang: b.email_khach_hang || '', NV_bo_phan: b.nv_bo_phan || '',
        NV_ten: b.nv_ten || '', NV_email: b.nv_email || '', NV_sdt: b.nv_sdt || '',
        Items_JSON: JSON.stringify(validItems), CK_Tong_don: ckTong, Tong_thanh_toan: total,
      };
      Promise.resolve()
        .then(() => uploadPdfToNocoDB(finalPath, finalName))
        .then(att => { if (att) record.File_PDF = [att]; return saveQuoteToNocoDB(record); })
        .then(() => console.log('✅ NocoDB saved'))
        .catch(e => console.error('NocoDB error:', e.message));
    });
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

app.listen(PORT, () => {
  console.log(`✅ Elide Fire Quote Server running on port ${PORT}`);
});
