// =========================
// KONFIGURASI
// =========================

const SPREADSHEET_ID = "1Gm0CZApL--kZ2xuLOP8ybBtA-6sa-79pbYvo95RpnIc";

const TEMPLATE_1 = "1WOrBHxJQfzPdsb10swXXAL-Y5J13zL6wIQZK01bufCk";
const TEMPLATE_2 = "1-PUC_trme474l2BwOFUseX5nN-YxLcoqwf1Wb_T-fdY";
const TEMPLATE_3 = "1mEkPgpImtv1Rg-TFZmJOqBLW2HR-zNNXKgWP7D4RW1c";

// =========================
// WEB APP
// =========================

function doGet() {
  return HtmlService
    .createHtmlOutputFromFile("Index")
    .setTitle("Input Verifikasi Ijazah");
}

// =========================
// CARI DATA
// =========================

function getDataByNim(nim) {
  try {
    if (!nim) return { ditemukan: false };

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName("Data_Surat");

    if (!sheet) {
      return { ditemukan: false, error: "Sheet tidak ditemukan" };
    }

    const data = sheet.getDataRange().getDisplayValues();

    const nimCari = String(nim).trim().replace(/\s+/g, "");

    for (let i = 1; i < data.length; i++) {
      const nimSheet = String(data[i][5]).trim().replace(/\s+/g, "");

      if (nimSheet === nimCari) {
        return {
          ditemukan: true,
          nama: data[i][2],
          tempat_lahir: data[i][3],
          tanggal_lahir: data[i][4],
          nim: data[i][5],
          fakultas: data[i][6],
          prodi: data[i][7],
          nomor_seri: data[i][8]
        };
      }
    }

    return { ditemukan: false };

  } catch (e) {
    return { ditemukan: false, error: e.message };
  }
}


// =========================
// SIMPAN DATA
// =========================

function simpanData(data) {

  try {

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    const sheet = ss.getSheetByName("Data_Surat");

    if (!sheet) {
      throw new Error("Sheet Data_Surat tidak ditemukan");
    }
    if (!data || !data.nama || !data.nim) {
  throw new Error("Data tidak lengkap");
}

    const nomorSurat =
      generateNomorSurat(sheet);

    const fileInfo =
      generateSurat(data, nomorSurat);

    sheet.appendRow([
      sheet.getLastRow(),
      nomorSurat,
      data.nama,
      data.tempat_lahir,
      data.tanggal_lahir,
      data.nim,
      data.fakultas,
      data.prodi,
      data.nomor_seri,
      data.tanggal_lulus,
      data.tanggal_ijazah,
      data.template,
      fileInfo.docx,
      new Date()
    ]);


    return {
      success: true,
      nomor: nomorSurat,
      docx: fileInfo.docx,
      pdf: fileInfo.pdf,
      edit: fileInfo.edit
    };

  } catch (err) {

    Logger.log(err);

    return {
      success: false,
      message: err.message
    };

  }
}

// =========================
// NOMOR SURAT
// =========================

function generateNomorSurat(sheet) {

const no = sheet.getLastRow() > 1 ? sheet.getLastRow() - 1 : 1;

  const bulanRomawi = [
    "I","II","III","IV","V","VI",
    "VII","VIII","IX","X","XI","XII"
  ];

  return Utilities.formatString(
    "%03d/D.3-II/%s/%s",
    no,
    bulanRomawi[new Date().getMonth()],
    new Date().getFullYear()
  );
}

// =========================
// GENERATE SURAT
// =========================

function generateSurat(data, nomorSurat) {

  let templateId = TEMPLATE_1;

  if (data.template === "2") {
    templateId = TEMPLATE_2;
  }

  if (data.template === "3") {
    templateId = TEMPLATE_3;
  }

 // 1. Ambil folder tujuan berdasarkan ID dari link yang Anda berikan
  const folderTujuan = DriveApp.getFolderById("1_tgvi3seSSxv8H50Ac-AIYtJE7YJLQtY");

  // 2. Gandakan template langsung ke dalam folder tersebut
  const fileCopy = DriveApp
    .getFileById(templateId)
    .makeCopy(
      "SURAT_" + nomorSurat.replace(/\//g, "-"),
      folderTujuan // <--- Menambahkan folder tujuan di sini
    );

  const fileId =
    fileCopy.getId();

  const doc =
    DocumentApp.openById(fileId);

  const body =
    doc.getBody();

  const tanggalLahir =
    formatTanggalIndonesia(
      data.tanggal_lahir
    );

  const tanggalIjazah =
    formatTanggalIndonesia(
      data.tanggal_ijazah
    );

  let tanggalLulus = "";

  if (data.tanggal_lulus) {

    tanggalLulus =
      formatTanggalIndonesia(
        data.tanggal_lulus
      );

  }

  const tanggalCetak =
    (data.template === "3")
      ? formatTanggalInggris()
      : formatTanggalCetakIndonesia();

  body.replaceText("\\[NOMOR_SURAT\\]", nomorSurat);
  body.replaceText("\\[NAMA\\]", data.nama || "");
  body.replaceText("\\[TEMPAT_LAHIR\\]", data.tempat_lahir || "");
  body.replaceText("\\[TANGGAL_LAHIR\\]", tanggalLahir);
  body.replaceText("\\[NIM\\]", data.nim || "");
  body.replaceText("\\[FAKULTAS\\]", data.fakultas || "");
  body.replaceText("\\[PRODI\\]", data.prodi || "");
  body.replaceText("\\[NO_IJAZAH\\]", data.nomor_seri || "");
  body.replaceText("\\[TGL_LULUS\\]", tanggalLulus);
  body.replaceText("\\[TGL_IJAZAH\\]", tanggalIjazah);
  body.replaceText("\\[TANGGAL_CETAK\\]", tanggalCetak);

  doc.saveAndClose();

  return {

    docx:
      "https://docs.google.com/document/d/" +
      fileId +
      "/export?format=docx",

    pdf:
      "https://docs.google.com/document/d/" +
      fileId +
      "/export?format=pdf",

    edit:
      "https://docs.google.com/document/d/" +
      fileId +
      "/edit"
  };
}

// =========================
// FORMAT TANGGAL
// =========================

function formatTanggalIndonesia(tgl) {

  if (!tgl) return "";

  const d = new Date(tgl);
  if (isNaN(d.getTime())) return "";

  const bulan = [
    "Januari","Februari","Maret","April","Mei","Juni",
    "Juli","Agustus","September","Oktober","November","Desember"
  ];

  return `${d.getDate()} ${bulan[d.getMonth()]} ${d.getFullYear()}`;
}

function formatTanggalCetakIndonesia() {

  return (
    "Yogyakarta, " +
    formatTanggalIndonesia(
      new Date()
    )
  );
}

function formatTanggalInggris() {

  return (
    "Yogyakarta, " +
    Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      "MMMM dd, yyyy"
    )
  );
}