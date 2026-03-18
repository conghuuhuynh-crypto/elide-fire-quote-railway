/**
 * server.js — Elide Fire Quote Server (Railway deployment)
 * Carbone open-source + LibreOffice → PDF không watermark
 */

const express = require('express');
const carbone  = require('carbone');
const path     = require('path');
const fs       = require('fs');
const { exec } = require('child_process');
const os       = require('os');

const app  = express();
const PORT = process.env.PORT || 3333;

const TEMPLATE   = path.join(__dirname, 'templates', 'quote-template.docx');
const QUOTES_DIR = path.join(__dirname, 'outputs', 'quotes');
const SOFFICE    = process.platform === 'win32'
  ? '"C:\\Program Files\\LibreOffice\\program\\soffice.exe"'
  : 'soffice';

app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use('/assets', express.static(path.join(__dirname, 'assets')));

// Serve form
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// Health check cho Railway
app.get('/health', (req, res) => res.json({ status: 'ok' }));

// API xuất PDF
app.post('/api/generate', (req, res) => {
  const b = req.body;

  const qty1   = parseFloat(b.so_luong_01) || 0;
  const price1 = parseFloat((b.gia_sp_01 || '').replace(/[^0-9]/g, '')) || 0;
  const ck1    = parseFloat(b.chiet_khau_01) || 0;
  const qty2   = parseFloat(b.so_luong_02) || 0;
  const price2 = parseFloat((b.gia_sp_02 || '').replace(/[^0-9]/g, '')) || 0;
  const ck2    = parseFloat(b.chiet_khau_02) || 0;

  const tt1   = qty1 * price1 * (1 - ck1 / 100);
  const tt2   = qty2 * price2 * (1 - ck2 / 100);
  const total = tt1 + tt2;
  const fmt   = (n) => n.toLocaleString('vi-VN');

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
    tong_thanh_tien:   fmt(total),
  };

  const soSlug  = (b.so_bao_gia || 'bao-gia').replace(/[\/\\:*?"<>|]/g, '-').trim();
  const tmpDocx = path.join(os.tmpdir(), `${soSlug}.docx`);
  const outPdf  = path.join(QUOTES_DIR, `${soSlug}.pdf`);

  if (!fs.existsSync(QUOTES_DIR)) fs.mkdirSync(QUOTES_DIR, { recursive: true });

  carbone.render(TEMPLATE, data, {}, (err, result) => {
    if (err) return res.status(500).send('Lỗi render: ' + err.message);

    fs.writeFileSync(tmpDocx, result);

    const cmd = `${SOFFICE} --headless --convert-to pdf --outdir "${QUOTES_DIR}" "${tmpDocx}"`;
    exec(cmd, { timeout: 60000 }, (err2) => {
      try { fs.unlinkSync(tmpDocx); } catch (_) {}

      if (err2) return res.status(500).send('Lỗi convert PDF: ' + err2.message);

      const tmpBasename = path.basename(tmpDocx, '.docx') + '.pdf';
      const libreOut    = path.join(QUOTES_DIR, tmpBasename);
      try { fs.renameSync(libreOut, outPdf); } catch (_) {}

      const finalPath = fs.existsSync(outPdf) ? outPdf : libreOut;
      const finalName = path.basename(finalPath);

      res.setHeader('Content-Disposition', `attachment; filename="${finalName}"`);
      res.setHeader('Content-Type', 'application/pdf');
      res.sendFile(finalPath);
    });
  });
});

app.listen(PORT, () => {
  console.log(`✅ Elide Fire Quote Server running on port ${PORT}`);
});
