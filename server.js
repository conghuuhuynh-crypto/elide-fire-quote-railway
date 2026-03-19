/**
 * server.js — Elide Fire Quote Server (Railway deployment)
 * Carbone open-source + LibreOffice → PDF không watermark
 */

const express = require('express');
const carbone  = require('carbone');
const path     = require('path');
const fs       = require('fs');
const https    = require('https');
const { exec } = require('child_process');
const os       = require('os');

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

app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use('/assets', express.static(path.join(__dirname, 'assets')));
app.use('/assets', express.static(path.join(__dirname, 'public', 'assets')));

// Serve form
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// Health check
app.get('/health', (req, res) => res.json({ status: 'ok' }));

// API lấy danh sách nhân viên từ NocoDB
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
      try {
        const data = JSON.parse(d);
        res.json(data.list || []);
      } catch (e) {
        res.json([]);
      }
    });
  }).on('error', () => res.json([]));
});

// Helper: upload PDF lên NocoDB storage
function uploadPdfToNocoDB(pdfPath, filename) {
  return new Promise((resolve) => {
    try {
      const fileData = fs.readFileSync(pdfPath);
      const boundary = 'FormBoundary' + Date.now();
      const body = Buffer.concat([
        Buffer.from(`--${boundary}\r\nContent-Disposition: form-data; name="file"; filename="${filename}"\r\nContent-Type: application/pdf\r\n\r\n`),
        fileData,
        Buffer.from(`\r\n--${boundary}--\r\n`)
      ]);
      const options = {
        hostname: NOCODB_HOST,
        path: '/api/v1/storage/upload',
        method: 'POST',
        headers: {
          'xc-token': NOCODB_TOKEN,
          'Content-Type': `multipart/form-data; boundary=${boundary}`,
          'Content-Length': body.length
        }
      };
      const req = https.request(options, r => {
        let d = '';
        r.on('data', c => d += c);
        r.on('end', () => {
          try { resolve(JSON.parse(d)); } catch (e) { resolve(null); }
        });
      });
      req.on('error', () => resolve(null));
      req.write(body);
      req.end();
    } catch (e) {
      resolve(null);
    }
  });
}

// Helper: lưu record vào NocoDB Bao_gia
function saveQuoteToNocoDB(record) {
  return new Promise((resolve) => {
    const body = Buffer.from(JSON.stringify(record));
    const options = {
      hostname: NOCODB_HOST,
      path: `/api/v1/db/data/noco/${NOCODB_BASE}/${TABLE_BG}`,
      method: 'POST',
      headers: {
        'xc-token': NOCODB_TOKEN,
        'Content-Type': 'application/json',
        'Content-Length': body.length
      }
    };
    const req = https.request(options, r => {
      let d = '';
      r.on('data', c => d += c);
      r.on('end', () => resolve());
    });
    req.on('error', () => resolve());
    req.write(body);
    req.end();
  });
}

// API xuất PDF
app.post('/api/generate', (req, res) => {
  const b = req.body;

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

  const discountAmt = tongTruoCK - total;
  const chietKhauDisplay = ckTong > 0
    ? `-${ckTong}% (-${fmt(discountAmt)} VNĐ)`
    : '0';

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
    chiet_khau:        chietKhauDisplay,
    tong_thanh_tien:   fmt(total),
    bo_phan:           b.nv_bo_phan        || '',
    ten_nhan_vien:     b.nv_ten            || '',
    email_nhan_vien:   b.nv_email          || '',
    sdt_nhan_vien:     b.nv_sdt            || '',
  };

  const soSlug  = (b.so_bao_gia || 'bao-gia').replace(/[\/\\:*?"<>|]/g, '-').trim();
  const tmpDocx = path.join(os.tmpdir(), `${soSlug}.docx`);
  const outPdf  = path.join(QUOTES_DIR, `${soSlug}.pdf`);

  if (!fs.existsSync(QUOTES_DIR)) fs.mkdirSync(QUOTES_DIR, { recursive: true });

  carbone.render(TEMPLATE, data, {}, (err, result) => {
    if (err) return res.status(500).send('Lỗi render: ' + err.message);

    fs.writeFileSync(tmpDocx, result);

    const cmd = `${SOFFICE} --headless --convert-to pdf --outdir "${QUOTES_DIR}" "${tmpDocx}"`;
    exec(cmd, { timeout: 60000 }, async (err2) => {
      try { fs.unlinkSync(tmpDocx); } catch (_) {}

      if (err2) return res.status(500).send('Lỗi convert PDF: ' + err2.message);

      const tmpBasename = path.basename(tmpDocx, '.docx') + '.pdf';
      const libreOut    = path.join(QUOTES_DIR, tmpBasename);
      try { fs.renameSync(libreOut, outPdf); } catch (_) {}

      const finalPath = fs.existsSync(outPdf) ? outPdf : libreOut;
      const finalName = path.basename(finalPath);

      // Lưu vào NocoDB trước khi gửi PDF
      try {
        const attachment = await uploadPdfToNocoDB(finalPath, finalName);
        await saveQuoteToNocoDB({
          So_bao_gia:       b.so_bao_gia        || '',
          Ngay_bao_gia:     b.ngay_bao_gia       || '',
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
          SL_Techideas:     qty1,
          DonGia_Techideas: price1,
          CK_Techideas:     ck1,
          ThanhTien_Techideas: tt1,
          SL_Lovingcare:    qty2,
          DonGia_Lovingcare: price2,
          CK_Lovingcare:    ck2,
          ThanhTien_Lovingcare: tt2,
          CK_Tong_don:      ckTong,
          Tong_thanh_toan:  total,
          File_PDF:         attachment ? [attachment] : null,
        });
      } catch (e) {
        console.error('NocoDB save error:', e.message);
      }

      // Gửi PDF về trình duyệt
      res.setHeader('Content-Disposition', `attachment; filename="${finalName}"`);
      res.setHeader('Content-Type', 'application/pdf');
      res.sendFile(finalPath);
    });
  });
});

app.listen(PORT, () => {
  console.log(`✅ Elide Fire Quote Server running on port ${PORT}`);
});
