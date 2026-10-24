function testGenerateSemuaTaruna() {
  var SPREADSHEET_ID = "1R67WyOUpf8IqZ1rNRDlMig_G88uA9QOORKRXHQKTy64";
  var templateId = "1l46acEx60-FrU3s0ZB7tNQsteD9Azej_PEHIABx70nE";
  var folderId = "1fRSB2Xgpoi7yk5YmX1junaVJpq3d6zyb";
  
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Form Responses 1");
  var data = sheet.getDataRange().getValues();

  var grouped = {};

  // ========================
  // GROUP SEMUA DATA
  // ========================
  var periode = getPeriodeMingguan();

  for (var i = 1; i < data.length; i++) {
  
    var nama = data[i][3];       
    var nit = data[i][4];

    var tanggaldata = new Date(data[i][2]);
    if (tanggaldata < periode.start || tanggaldata > periode.end) continue;

    if (!nama) continue;


    var key = nama + "_" + nit;

    if (!grouped[key]) {
      grouped[key] = {
        nama: nama,
        nit: nit,
        email: data[i][1],
        prodi: data[i][5],
        dosen: data[i][7],
        satuan: data[i][6],
        lokasi: data[i][8],
        instruktur: data[i][9],
        rows: []
      };
    }

    grouped[key].rows.push({
      tanggal: tanggaldata,
      jenis: data[i][10],
      uraian: data[i][11],
      dokumentasi: data[i][12],
      instruktur: data[i][9]
    });
  }

  // ========================
  // LOOP PROSES DOKUMEN
  // ========================
  for (var key in grouped) {

    var obj = grouped[key];
    var nama = obj.nama;

    try {
      // 1. Copy Template
      var tempFile = DriveApp.getFileById(templateId).makeCopy("Temp_" + nama);
      var doc = DocumentApp.openById(tempFile.getId());
      var body = doc.getBody();

      // 2. Isi Header
      body.replaceText("<<Nama Taruna>>", nama);
      body.replaceText("<<NIT>>", obj.nit);
      body.replaceText("<<Prodi>>", obj.prodi);
      body.replaceText("<<Dosen Pembimbing>>", obj.dosen);
      body.replaceText("<<Satuan Pendidikan>>", obj.satuan.toUpperCase());
      body.replaceText("<<Lokasi Kegiatan>>", obj.lokasi.toUpperCase());

      var listInstruktur = getAllInstruktur(obj.rows);
      var formatted = formatInstrukturRapi(listInstruktur, 38);

      body.replaceText("<<Nama Instruktur/Teknisi>>", formatted);

      // 3. Isi Tabel
      var found = body.findText("{{Tabel}}");

      if (found) {
        var paragraph = found.getElement().getParent();

        while (paragraph.getType() !== DocumentApp.ElementType.PARAGRAPH) {
          paragraph = paragraph.getParent();
        }

        var parent = paragraph.getParent();
        var index = parent.getChildIndex(paragraph);

        paragraph.removeFromParent();

        // MENAMBAHKAN "No" PADA HEADER
        var table = parent.insertTable(index, [
          ["No", "Tanggal", "Topik Pembelajaran", "Uraian Pembelajaran", "Dokumentasi"]
        ]);

        table.setBorderWidth(1);
        
        // MENYESUAIKAN LEBAR KOLOM (Total 5 kolom sekarang)
        table.setColumnWidth(0, 25);  // Lebar untuk kolom nomor
        table.setColumnWidth(1, 70);  // Tanggal
        table.setColumnWidth(2, 100); // Jenis Kegiatan
        table.setColumnWidth(3, 200); // Uraian
        table.setColumnWidth(4, 120); // Dokumentasi

        // MENAMBAHKAN VARIABEL INDEX (i) UNTUK PENOMORAN
        obj.rows.forEach(function(r, i) {
          var row = table.appendTableRow();

          row.appendTableCell((i + 1).toString()); // Mengisi nomor otomatis (mulai dari 1)
          row.appendTableCell(formatTanggal(r.tanggal));
          row.appendTableCell(r.jenis);
          row.appendTableCell(r.uraian);

          var cellDok = row.appendTableCell("");
          
          if (r.dokumentasi && r.dokumentasi.toString().startsWith("http")) {
            cellDok.setText("Lihat Dokumentasi");
            var text = cellDok.editAsText();
            text.setLinkUrl(0, text.getText().length - 1, r.dokumentasi);
          } else {
            cellDok.setText("-");
          }
        });
        
        body.appendParagraph("");
        buatTabelTTD(body, body.getNumChildren(), listInstruktur);
      }

      doc.saveAndClose();

      // ========================
      // KONVERSI KE DOCX
      // ========================
      var docId = tempFile.getId();

      var url = "https://docs.google.com/feeds/download/documents/export/Export?id=" + docId + "&exportFormat=docx";

      var response = UrlFetchApp.fetch(url, {
        method: "get",
        headers: {
          "Authorization": "Bearer " + ScriptApp.getOAuthToken()
        }
      });

      var today = new Date();

      var namaFileSafe = nama.replace(/[^\w\s]/gi, "").replace(/\s+/g, "_");

      var startFile = formatTanggalFile(periode.start);
      var endFile = formatTanggalFile(periode.end);

      var fileName = "Jurnal_Harian_" + namaFileSafe + "_" + startFile + "_sd_" + endFile + ".docx";

      var docxBlob = response.getBlob().setName(fileName);

      // Kirim email
      var startDisplay = formatTanggalIndo(periode.start);
      var endDisplay = formatTanggalIndo(periode.end)
      var subject = "Jurnal Harian Program Pembelajaran Taruna Berprestasi (PTB) Bidang Budi Daya Ikan Periode " + startDisplay + " s/d " + endDisplay;
      var emailbodyHTML = 
          "Halo <b>" + nama + "</b>,<br><br>" +
          "Berikut adalah jurnal harian Anda untuk tanggal " + startDisplay + " s/d " + endDisplay + ".<br>" + "Silakan berkoordinasi dengan Instruktur/Teknisi Anda untuk ditandatangani.<br>" + "Sebelum meminta tandatangan, Anda dapat mengedit jurnal ini jika terdapat kesalahan atau ada hal yang ingin diperbaiki. Kemudian, masukkan foto dokumentasi pada kolom yang tersedia.<br><br>" + 
          "Terima kasih dan tetap semangat.<br><br>" + 
          "<i>(Email ini terkirim secara otomatis setiap hari Senin oleh sistem, mohon untuk tidak dibalas.)</i><br><br>";
      GmailApp.sendEmail(
        obj.email,
        subject,
        "",
        {
          name: "Jurnal PTB Budi Daya 2026",
          htmlBody: emailbodyHTML,
          attachments: [docxBlob]
        }
      );

      // Simpan ke folder
      DriveApp.getFolderById(folderId).createFile(docxBlob);

      tempFile.setTrashed(true);

      Logger.log("SUKSES: " + nama);

    } catch (err) {
      Logger.log("ERROR di " + nama + ": " + err.toString());
    }
  }
}

function formatTanggal(date) {
  return Utilities.formatDate(date, "GMT+7", "dd/MM/yyyy");
}

function splitInstruktur(text) {
  if (!text) return [];

  // Pisah awal berdasarkan koma
  var parts = text.split(",");

  var result = [];
  var current = "";

  parts.forEach(function(p) {
    p = p.trim();

    
    if (/^[A-Z][a-z]/.test(p)) {
      if (current) result.push(current.trim());
      current = p;
    } else {
      // dianggap lanjutan gelar
      current += ", " + p;
    }
  });

  if (current) result.push(current.trim());

  return result;
}

function formatList(text) {
  if (!text) return "";

  var items = text.split(",");
  var result = "";

  for (var i = 0; i < items.length; i++) {
    if (i === 0) {
      result += (i + 1) + ". " + items[i].trim();
    } else {
      result += "\n    " + (i + 1) + ". " + items[i].trim(); // indent
    }
  }

  return result;
}

function formatInstrukturRapi(list, labelLength) {
  var indent = " ".repeat(labelLength); // sejajar dengan setelah ": "
  var result = "";

  list.forEach(function(item, i) {
    if (i === 0) {
      result += (i + 1) + ". " + item;
    } else {
      result += "\n" + indent + (i + 1) + ". " + item;
    }
  });

  return result;
}

function formatTanggalIndo(date) {
  var bulan = [
    "januari","februari","maret","april","mei","juni",
    "juli","agustus","september","oktober","november","desember"
  ];

  var tgl = date.getDate().toString().padStart(2, '0');
  var bln = bulan[date.getMonth()];
  bln = bln.charAt(0).toUpperCase() + bln.slice(1);
  var thn = date.getFullYear();

  return tgl + " " + bln + " " + thn;
}

function formatTanggalFile(date) {
  var yyyy = date.getFullYear();
  var mm = (date.getMonth() + 1).toString().padStart(2, '0');
  var dd = date.getDate().toString().padStart(2, '0');

  return yyyy + "-" + mm + "-" + dd;
}

function getPeriodeMingguan() {
  var today = new Date();

  var end = new Date(today);
  end.setDate(end.getDate() - end.getDay());
  end.setHours(23, 59, 59, 999);

  var start = new Date(end);
  start.setDate(start.getDate() - 6);
  start.setHours(0, 0, 0, 0);

  return { start: start, end: end };
}

function buatTabelTTD(parent, index, listInstruktur) {

  var maxKolom = 3;
  var jumlah = listInstruktur.length;
  var jumlahBaris = Math.ceil(jumlah / maxKolom);

  for (var b = 0; b < jumlahBaris; b++) {

    var isLastRow = (b === jumlahBaris - 1);

    // hitung sisa kolom di baris terakhir
    var sisa = jumlah % maxKolom;
    if (sisa === 0) sisa = maxKolom;

    // =========================
    // BARIS HEADER
    // =========================
    var headerRow = [];

    for (var k = 0; k < maxKolom; k++) {

      var idx = b * maxKolom + k;

      if (idx < jumlah) {
        headerRow.push("Instruktur/Teknisi, " + (idx + 1));

      } else if (isLastRow && k === maxKolom - 1) {
        
        headerRow.push("Taruna/i,");

      } else {
        headerRow.push("");
      }
    }

    var table = parent.insertTable(index++, [headerRow]);
    table.setBorderWidth(0);

    // =========================
    // BARIS AREA TTD
    // =========================
    var rowTTD = table.appendTableRow();
    rowTTD.setMinimumHeight(80);

    for (var k = 0; k < maxKolom; k++) {
      rowTTD.appendTableCell("");
    }

    var rowText = table.appendTableRow();

    for (var k = 0; k < maxKolom; k++) {

      var idx = b * maxKolom + k;
      var cell = rowText.appendTableCell("");

      if (idx < jumlah) {
        cell.setText("(TTD dan Nama Lengkap)");

      } else if (isLastRow && k === maxKolom - 1) {
        cell.setText("(TTD dan Nama Lengkap)");

      } else {
        cell.setText("");
      }
    }

    parent.insertParagraph(index++, "");
  }
  if (jumlah % maxKolom === 0) {

    var table = parent.insertTable(index++, [
      ["", "", "Taruna/i"]
    ]);
    table.setBorderWidth(0);

    var rowTTD = table.appendTableRow();
    rowTTD.setMinimumHeight(80);

    rowTTD.appendTableCell("");
    rowTTD.appendTableCell("");
    rowTTD.appendTableCell("");

    var rowText = table.appendTableRow();
    rowText.appendTableCell("");
    rowText.appendTableCell("");
    rowText.appendTableCell("(TTD dan Nama Lengkap)");
  }
}

function getAllInstruktur(rows) {

  var all = [];

  rows.forEach(function(r) {
    var list = splitInstruktur(r.instruktur);
    all = all.concat(list);
  });

  // hapus duplikasi
  var unique = [...new Set(all)];

  return unique;
}
