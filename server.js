/**
 * server.js — Elide Fire Quote Server (Railway deployment)
 * Carbone open-source + LibreOffice → PDF không watermark
 */

const express  = require('express');
const carbone  = require('carbone');
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

// Job queue — lưu trạng thái từng job trong bộ nhớ
const jobs = {};

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

  const qty1   = parseFloat(b.so_luong_01) || 0;
  const price1 = parseFloat((b.gia_sp_01 || '').replace(/[^0-9]/g, '')) || 0;
  const ck1    = parseFloat(b.chiet_khau_01) || 0;
  const qty2   = parseFloat(b.so_luong_02) || 0;
  const price2 = parseFloat((b.gia_sp_02 || '').replace(/[^0-9]/g, '')) || 0;
  const ck2    = parseFloat(b.chiet_khau_02) || 0;
  const ckTong = parseFloat(b.chiet_khau_tong) || 0;

  const tt1        = qty1 * price1 * (1 - ck1 / 100);
  const tt2        = qty2 * price2 * (1 - ck2 / 100);
  const tongTruoCK = tt1 + tt2;
  const total      = tongTruoCK * (1 - ckTong / 100);
  const fmt        = (n) => n.toLocaleString('vi-VN');

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
    so_luong_01:       String(qty1 || '0'),
    gia_sp_01:         fmt(price1 * (1 - ck1 / 100)),
    thanh_tien_01:     fmt(tt1),
    so_luong_02:       String(qty2 || '0'),
    gia_sp_02:         fmt(price2 * (1 - ck2 / 100)),
    thanh_tien_02:     fmt(tt2),
    truoc_chiet_khau:  fmt(tongTruoCK),
    chiet_khau:        ckTong > 0 ? `${ckTong}%` : '0',
    tong_thanh_tien:   fmt(total),
    bo_phan:           b.nv_bo_phan || '',
    ten_nhan_vien:     b.nv_ten     || '',
    email_nhan_vien:   b.nv_email   || '',
    sdt_nhan_vien:     b.nv_sdt     || '',
  };

  const soSlug  = (b.so_bao_gia || 'bao-gia').replace(/[\/\\:*?"<>|]/g, '-').trim();
  const tmpDocx = path.join(os.tmpdir(), `${soSlug}-${jobId}.docx`);
  const outPdf  = path.join(QUOTES_DIR, `${soSlug}.pdf`);

  if (!fs.existsSync(QUOTES_DIR)) fs.mkdirSync(QUOTES_DIR, { recursive: true });

  carbone.render(TEMPLATE, data, {}, (err, result) => {
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

      // Job done — cập nhật trạng thái
      job.status   = 'done';
      job.url      = `${appUrl}/download/${finalName}`;
      job.filename = finalName;

      // Lưu NocoDB nền
      const record = {
        So_bao_gia: b.so_bao_gia || '', Ngay_bao_gia: b.ngay_bao_gia || '',
        Phien_ban: b.phien_ban || '', Ten_du_an: b.ten_du_an || '',
        Ten_cong_ty: b.ten_cong_ty || '', Phong_ban_KH: b.ten_phong_ban || '',
        Nguoi_lien_he: b.ten_nguoi_lien_he || '', SDT_khach_hang: b.sdt_khach_hang || '',
        Email_khach_hang: b.email_khach_hang || '', NV_bo_phan: b.nv_bo_phan || '',
        NV_ten: b.nv_ten || '', NV_email: b.nv_email || '', NV_sdt: b.nv_sdt || '',
        SL_Techideas: qty1, DonGia_Techideas: price1, CK_Techideas: ck1, ThanhTien_Techideas: tt1,
        SL_Lovingcare: qty2, DonGia_Lovingcare: price2, CK_Lovingcare: ck2, ThanhTien_Lovingcare: tt2,
        CK_Tong_don: ckTong, Tong_thanh_toan: total,
      };
      try {
        const att = await uploadPdfToNocoDB(finalPath, finalName);
        if (att) record.File_PDF = [att];
        await saveQuoteToNocoDB(record);
        console.log('✅ NocoDB saved');
      } catch (e) { console.error('NocoDB error:', e.message); }
    });
  });
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
