function salinFile(tanggalTerbit) {
  const namaFileSalinan = `Sertifikat Seminar ${tanggalTerbit}`;

  // buka file blanko
  const file = DriveApp.getFileById(env('idSlideBlanko'));

  // salin file blanko
  const idSlideSalinan = file.makeCopy(namaFileSalinan).getId();

  return idSlideSalinan;
}


function editFile(idSlideSalinan, namaPeserta, kategori) {
  const slides = SlidesApp.openById(idSlideSalinan);
  const slide = slides.getSlides()[0];
  const shapes = slide.getShapes();

  shapes.forEach(shape => {
    if (shape.getText) {
      const textRange = shape.getText();
      textRange.replaceAllText('<<Nama Lengkap>>', namaPeserta);
      textRange.replaceAllText('<<Kategori>>', kategori);
    }
  });

  slides.saveAndClose();
}


function eksporFile(idSlideSalinan, tanggalTerbit) {
  // nama file
  const namaFile = `Sertifikat Seminar ${tanggalTerbit}.pdf`;

  // jadikan pdf dan ubah nama file
  const pdf = DriveApp.getFileById(idSlideSalinan).getBlob().getAs('application/pdf');
  pdf.setName(namaFile);

  // pindah file pdf ke folder ekspor
  const sertifikatPdf = DriveApp.getFolderById(env('idFolder')).createFile(pdf);
  const idSertifikatPdf = sertifikatPdf.getId();

  return idSertifikatPdf;
}


function ambilUrlPdf(idSertifikatPdf) {
  // buka file pdf
  const sertifikatPdf = DriveApp.getFileById(idSertifikatPdf);

  // ubah permission file pdf
  sertifikatPdf.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  // ambil url
  const urlSertifikatPdf = sertifikatPdf.getUrl();

  return urlSertifikatPdf;
}


function hapusFileSalinan(idSlideSalinan) {
  // hapus slides salinan
  DriveApp.getFileById(idSlideSalinan).setTrashed(true);
}

function tulisPesan(namaPeserta, urlSertifikatPdf, namaAcara = 'Dharma Shanti', tanggalAcara = '17 April 2025') {
  const pesan = `*E-SERTIFIKAT - ${namaPeserta} - ${namaAcara}*

Assalamualaikum warahmatullahi wabarakatuh,  
Salam sejahtera bagi kita semua,  
Shalom,  
Om Swastiastu,  
Namo Buddhaya,  
Salam Kebajikan.

Kami mengucapkan terima kasih atas partisipasi aktif Anda dalam *Seminar ${namaAcara}*.

Semoga kegiatan ini menjadi sumber inspirasi dan memperkuat nilai-nilai spiritual dalam kehidupan kita bersama.

Sertifikat ini menjadi bukti kontribusi dan semangat belajar Anda.

📄 *Nama Peserta:* ${namaPeserta}
📅 *Tanggal Acara:* ${tanggalAcara}

🔗 *Link Download Sertifikat:*
${urlSertifikatPdf}

📌 *Catatan Penting:*
Jika link di atas tidak bisa diklik, silakan:
1️⃣ Simpan nomor ini sebagai kontak  
2️⃣ Atau balas pesan ini dengan teks sembarang, misal: *"OK"*

Terima kasih atas keikutsertaannya.

*Salam sejahtera bagi kita semua.*;`

  return pesan;
}


function kirimWa(nomor, pesan) {
  const headers = {
    'Authorization': env('wablasToken'),
    'content-type' : 'application/json'
  };

  const payload = JSON.stringify(
    {
      'data':
      [{
        'phone' : nomor,
        'message': pesan,
        'secret': false,
        'retry': false,
        'isGroup': false
      }]
    }
  );
  
  const option = {
    'method': 'POST',
    'headers': headers,
    'payload': payload
  };

  const response = UrlFetchApp.fetch(env('wablasUrl'), option);
  const responseJson = JSON.parse(response.getContentText());

  Logger.log(responseJson);

  return responseJson['message'];
}


function bacaForm(e) {
  // baca inputan form
  const namaPeserta = e.namedValues['Nama Lengkap'][0].trim();
  const noWaPeserta = e.namedValues['No. WhatsApps'][0].trim();
  const kategori = e.namedValues['Kategori'][0].trim(); // Tambahkan ini!

  // tanggal dan waktu
  const now = new Date;
  const tanggalTerbit = `${now.getDate()}-${now.getMonth()+1}-${now.getFullYear()}_${now.getHours()}.${now.getMinutes()}.${now.getSeconds()}`;


  // proses
  const idSlideSalinan = salinFile(tanggalTerbit);
  editFile(idSlideSalinan, namaPeserta, kategori);
  const idSertifikatPdf = eksporFile(idSlideSalinan, tanggalTerbit);
  const urlSertifikatPdf = ambilUrlPdf(idSertifikatPdf);
  hapusFileSalinan(idSlideSalinan);

  // kirim wa
  const pesan = tulisPesan(namaPeserta, urlSertifikatPdf);
  const report = kirimWa(noWaPeserta, pesan);

  // buka sheet
  const spreadsheet = SpreadsheetApp.openById(env('idSheet'));
  const sheet = spreadsheet.getSheetByName('Form Responses 1');
  const lastRow = sheet.getLastRow();

  // tulis link pdf
  sheet.getRange(`L${lastRow}`).setValue(urlSertifikatPdf);

  // tulis hasil pengiriman
  sheet.getRange(`M${lastRow}`).setValue(report);
}

// function testBacaForm() {
//   const e = {
//     namedValues: {
//       "Nama Lengkap": ["Tes Nama"],
//       "No. WhatsApps": ["628"],
//       "Kategori": ["Umum"]
//     }
//   };
//   bacaForm(e);
// }
