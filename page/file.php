<!DOCTYPE html>
<html lang="en">
  <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>Document LHP</title>
    <link rel="stylesheet" href="../style/app.css" />
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.17.1/xlsx.full.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/exceljs/dist/exceljs.min.js"></script>
    <!-- <style>
      @page {
        margin: 20mm;
        size: "[216, 356]" landscape;
      }
    </style> -->
  </head>
  <body>
    <button id="convertPDF">Export ke PDF</button>
    <button id="exportBtn">Export ke Excel</button>

    <div id="template">
      <div class="lampiran">
        <p>LAMPIRAN LAPORAN HASIL PEMERIKSAAN</p>
        <p id="nomor">NOMOR : 900.1.13.1/S-24./PBB-P2/424.102/2024</p>
        <p>TANGGAL : 0</p>
      </div>

      <h1>PERUBAHAN DATA SUBJEK/OBJEK PAJAK</h1>
      <h3 id="jenis_mutasi">JENIS MUTASI:</h3>
      <div class="table-content">
        <div class="kecamatan">
          <p id="kecamatan">KECAMATAN :</p>
          <p id="desa">DESA/KELURAHAN :</p>
        </div>
        <table>
          <thead>
            <tr>
              <th colspan="7" id="data_lama">DATA LAMA</th>
              <th colspan="7" id="data_baru">DATA BARU</th>
              <th rowspan="3" class="keterangan">KETERANGAN</th>
            </tr>
            <tr>
              <td rowspan="2" class="no">NO</td>
              <th rowspan="2" class="nop">NOP</th>
              <th rowspan="2">NAMA</th>
              <th colspan="2" id="kolom_bumi1">BUMI</th>
              <th colspan="2" id="kolom_bang1">BANGUNAN</th>
              <th rowspan="2">NO</th>
              <th rowspan="2" class="nop">NOP</th>
              <th rowspan="2">NAMA</th>
              <th colspan="2" id="kolom_bumi2">BUMI</th>
              <th colspan="2" id="kolom_bang2">BANGUNAN</th>
            </tr>
            <tr>
              <th id="kolom_luas">LUAS</th>
              <th>ZNT</th>
              <th>LUAS</th>
              <th id="klas1">KLAS</th>
              <th>LUAS</th>
              <th>ZNT</th>
              <th>LUAS</th>
              <th>KLAS</th>
            </tr>
          </thead>
          <tbody></tbody>
        </table>
      </div>
      <div class="ttd">
        <p>Petugas Peneliti</p>
        <br /><br /><br />
        <u id="petugas_peneliti"
          >Mochamad asasasGrisvian Gema Elvitra, A. Md. A. Pjk.</u
        >
        <p id="nip_peneliti">NIP. 20000130 202201 1 003</p>
        <div class="ttd-1">
          <div>
            <p>Diperiksa Oleh,</p>
            <p>Kasubbid Pendaftaran dan Pendataan</p>
            <br /><br /><br />
            <u id="kasubbid_pendataan">Muhammad Khoriri, SE</u>
            <p id="eselon_kasubbid_pendataan">Penata Tingkat I (III/d)</p>
            <p id="nip_kasubbid_pendataan">NIP. 19681012 199003 1 010</p>
          </div>
          <div>
            <p>Diperiksa Oleh,</p>
            <p>Kasubbid Perhitungan dan Penetapan</p>
            <br /><br /><br />
            <u id="kasubbid_penetapan">Sanca Dwi Anggoro, S.Kom, MM</u>
            <p id="eselon_kasubbid_penetapan">Penata (III/C)</p>
            <p id="nip_kasubbid_penetapan">NIP. 19851214 201001 1 009</p>
          </div>
        </div>
      </div>
    </div>

    <script>
      document.addEventListener("DOMContentLoaded", function () {
        const urlParams = new URLSearchParams(window.location.search);
        // const idMutasi = urlParams.get("idmutasi_pbb"); // Ambil ID dari parameter URL

        const idMutasi = 1;

        if (idMutasi) {
          fetch(
            `https://api.pendapatan.pasuruankab.go.id/api/get/mutasi/pbb?idmutasi_pbb=${idMutasi}`,
            {
              method: "GET",
              headers: {
                Authorization: "Bearer 8853147364",
                "Content-Type": "application/json",
              },
            }
          )
            .then((response) => response.json())
            .then((data) => {
              if (
                data.status === "success" &&
                Array.isArray(data.data) &&
                data.data.length > 0
              ) {
                const mutasiData = data.data; // Data berbentuk array
                isiTabel(mutasiData);
                isiPejabat(mutasiData[0]);
                isiInfoLain(mutasiData[0]);
                isiNomor(mutasiData[0]);
                formatTanggal(mutasiData[0]);
              } else {
                console.error("Gagal mengambil data:", data.message);
              }
            })
            .catch((error) => console.error("Error:", error));
        } else {
          console.error("ID Mutasi tidak ditemukan dalam parameter URL");
        }

        function isiTabel(mutasiData) {
          const tableBody = document.querySelector("#template table tbody");
          tableBody.innerHTML = "";

          let totalLuasBumi = 0;
          let totalLuasBangunan = 0;
          let totalLuasBumiBaru = 0;
          let totalLuasBangunanBaru = 0;
          let totalLuasBumi1 = 0;
          let totalLuasBangunan1 = 0;
          let nomorUrut = 1;
          let nomorUrut1 = 1;

          function formatNOP(nop) {
            if (nop && nop.length === 18) {
              return `${nop.substring(0, 2)}.${nop.substring(
                2,
                5
              )}.${nop.substring(5, 8)}.${nop.substring(8, 11)}-${nop.substring(
                11,
                15
              )}-${nop.substring(15, 18)}`;
            }
            return nop;
          }

          mutasiData.forEach((mutasi, index) => {
            const daftarNop = JSON.parse(mutasi.nop); // Ambil daftar NOP
            // const detail = JSON.parse(daftarNop[0].detail); // Ambil detail dari daftar NOP pertama
            const nopList = JSON.parse(mutasi.nop);
            const nopBaruList = JSON.parse(mutasi.nop_baru);

            const nopDisplay = nopList.join("<br>");
            const nopBaruDisplay = nopBaruList.join("<br>");

            let totalLuas = 0;
              let totalBangunan = 0;
              let totalLuasBaru = 0;
              let totalBangunanBaru = 0;

            if (mutasi.pemecahan_penyatuan === "penyatuan") {
              // Mutasi Penyatuan
              // const detailBaru = JSON.parse(daftarNop[0].detail); // Ambil detail dari daftar NOP pertama
              
              nopBaruList.forEach((nopBaru) => {
                totalLuasBaru += parseInt(nopBaru.lbumi);
                totalBangunanBaru += parseInt(nopBaru.lbng) || 0;
              });

              totalLuasBumiBaru += totalLuasBaru;
              totalLuasBangunanBaru += totalBangunanBaru;

              nopList.forEach((nop, nopIndex) => {
                totalLuas += parseInt(nop.lbumi);
                totalBangunan += parseInt(nop.lbng) || 0;

                const row = tableBody.insertRow();
                row.innerHTML = `
                <td>${nomorUrut++}</td>
                <td>${formatNOP(nop.nop.trim())}</td>
                <td>${nop.nama.trim()}</td>
                <td>${nop.lbumi}</td>
                <td>${nop.znt}</td>
                <td>${nop.lbng}</td>
                <td>${nop.sjpt}</td>
                ${
                  nopIndex === 0
                    ? `
                <td rowspan="${nopList.length}">${nomorUrut1}</td>
                <td rowspan="${nopList.length}">${nopBaruList
                        .map((item) => formatNOP(item.nop_baru.trim()))
                        .join("<br>")}</td>
                <td rowspan="${nopList.length}">${nopBaruList
                        .map((item) => item.nama_baru.trim())
                        .join("<br>")}</td>
                <td rowspan="${nopList.length}">${nopBaruList[0].lbumi}</td>
                <td rowspan="${nopList.length}">${nopBaruList[0].znt}</td>
                <td rowspan="${nopList.length}">${nopBaruList[0].lbng}</td>
                <td rowspan="${nopList.length}">${nopBaruList[0].sjpt}</td>
                <td rowspan="${nopList.length}">${mutasi.keterangan.trim()}</td>
            `
                    : ``
                }
            `;
              });

              totalLuasBumi += totalLuas;
              totalLuasBangunan += totalBangunan;
            } else {
              // Mutasi Pemecahan
              nopBaruList.forEach((nop, nopIndex) => {
                const row = tableBody.insertRow();
              row.innerHTML = `
        ${
            nopIndex === 0
                ? `
                    <td rowspan="${nopBaruList.length}">${nomorUrut}</td>
                    <td rowspan="${nopBaruList.length}">${formatNOP(nopList[0].nop.trim())}</td>
                    <td rowspan="${nopBaruList.length}">${nopList[0].nama.trim()}</td>
                    <td rowspan="${nopBaruList.length}">${nopList[0].lbumi}</td>
                    <td rowspan="${nopBaruList.length}">${nopList[0].znt}</td>
                    <td rowspan="${nopBaruList.length}">${nopList[0].lbng}</td>
                    <td rowspan="${nopBaruList.length}">${nopList[0].sjpt}</td>
                `
                : `
                `
        }
        <td>${nomorUrut1++}</td>
        <td>${formatNOP(nop.nop_baru.trim())}</td>
        <td>${nop.nama_baru.trim()}</td>
        <td>${nop.lbumi}</td>
        <td>${nop.znt}</td>
        <td>${nop.lbng}</td>
        <td>${nop.sjpt}</td>
        ${
            nopIndex === 0
                ? `<td rowspan="${nopBaruList.length}">${mutasi.keterangan.trim()}</td>`
                : ``
        }
    `;
              });
              
              console.log(nopBaruList[0].nop_baru);

              totalLuasBumi += parseInt(nopList[0].lbumi) || 0;
              totalLuasBangunan += parseInt(nopList[0].lbng) || 0;

              nopBaruList.forEach((nopBaru) => {
                totalLuasBaru += parseInt(nopBaru.lbumi);
                totalBangunanBaru += parseInt(nopBaru.lbng) || 0;
              });

              totalLuasBumiBaru += totalLuasBaru;
              totalLuasBangunanBaru += totalBangunanBaru;
            }
          });
          const totalRow = tableBody.insertRow();
          totalRow.innerHTML = `
        <td colspan="3" id="totalLB">TOTAL</td>
        <td>${totalLuasBumi}</td>
        <td></td>
        <td class="totalLB1">${totalLuasBangunan}</td>
        <td></td>
        <td colspan="3" id="totalLB">TOTAL</td>
        <td>${totalLuasBumiBaru}</td>
        <td></td>
        <td class="totalLB1">${totalLuasBangunanBaru}</td>
        <td></td>
        <td></td>
    `;
        }
        function isiInfoLain(data) {
          document.getElementById("kecamatan").textContent =
            "KECAMATAN : " + data.kecamatan;
          document.getElementById("desa").textContent =
            "DESA/KELURAHAN : " + data.desa;
          document.getElementById("jenis_mutasi").textContent =
            "JENIS MUTASI : " + data.pemecahan_penyatuan.toUpperCase();
          console.log(data.desa);
          console.log(data.kecamatan);
        }

        function isiPejabat(data) {
          const pejabat = JSON.parse(data.pejabat_terkait);

          document.getElementById("petugas_peneliti").textContent =
            pejabat[0].nama;
          document.getElementById("nip_peneliti").textContent =
            "NIP. " + pejabat[0].nip;
          document.getElementById("kasubbid_pendataan").textContent =
            pejabat[1].nama;
          document.getElementById("eselon_kasubbid_pendataan").textContent =
            pejabat[1].golongan;
          document.getElementById("nip_kasubbid_pendataan").textContent =
            "NIP. " + pejabat[1].nip;
          document.getElementById("kasubbid_penetapan").textContent =
            pejabat[2].nama;
          document.getElementById("eselon_kasubbid_penetapan").textContent =
            pejabat[2].golongan;
          document.getElementById("nip_kasubbid_penetapan").textContent =
            "NIP. " + pejabat[2].nip;
        }

        function formatTanggal(data) {
          const tanggalApi = data.createddate;
          const date = new Date(tanggalApi);
          const day = String(date.getDate()).padStart(2, "0");
          const month = String(date.getMonth() + 1).padStart(2, "0");
          const year = date.getFullYear();

          const tanggalIndonesia = `${day}-${month}-${year}`;
          document.querySelector(".lampiran p:nth-child(3)").textContent =
            "TANGGAL : " + tanggalIndonesia;
        }
        function isiNomor(data) {
          document.getElementById("nomor").textContent =
            "NOMOR : " + data.nomer;
        }

        document
          .getElementById("exportBtn")
          .addEventListener("click", function () {
            const wb = XLSX.utils.book_new();
            const htmlContent = document.getElementById("template");
            let data = [];
            let rowIndex = 0; // Indeks baris saat ini

            let merges = [];

            function addEmptyRows(count) {
              for (let i = 0; i < count; i++) {
                data.push([]);
                rowIndex++;
              }
            }

            const lampiranParagraphs =
              htmlContent.querySelectorAll(".lampiran p");
            lampiranParagraphs.forEach((p) => {
              let rowData = [];
              for (let i = 0; i < 12; i++) {
                rowData.push("");
              }
              rowData.push(p.textContent);
              data.push(rowData);
              merges.push({
                s: { r: rowIndex, c: 17 },
                e: { r: rowIndex, c: 19 },
              }); // Tambahkan merge cell
              rowIndex++;
            });

            // Ambil judul (<h1>)
            const headings = htmlContent.querySelectorAll("h1");
            headings.forEach((h) => {
              data.push([h.textContent]);
              rowIndex++;
            });

            addEmptyRows(1);

            addEmptyRows(1); // Tambahkan baris kosong

            // Ambil div kecamatan
            const kecamatans = htmlContent.querySelectorAll(".kecamatan");
            kecamatans.forEach((kec) => {
              let kecamatanText = "";
              let desaText = "";

              kec.querySelectorAll("p").forEach((p) => {
                if (p.id === "kecamatan") {
                  kecamatanText = p.textContent;
                } else if (p.id === "desa") {
                  desaText = p.textContent;
                }
              });

              let rowData = [kecamatanText];
              for (let i = 1; i < 7; i++) {
                rowData.push("");
              }
              rowData.push(desaText);
              data.push(rowData);

              rowIndex++;
            });

            addEmptyRows(1);

            function formatNOP(nop) {
              if (nop && nop.length === 18) {
                return `'${nop.substring(0, 2)}.${nop.substring(
                  2,
                  5
                )}.${nop.substring(5, 8)}.${nop.substring(
                  8,
                  11
                )}-${nop.substring(11, 15)}-${nop.substring(15, 18)}`;
              }
              return `'${nop}`; // Tambahkan apostrof ke nilai NOP yang tidak diformat
            }

            // Format NOP di data tabel
            for (let i = 0; i < data.length; i++) {
              if (data[i].length > 1) {
                // Pastikan baris memiliki data
                data[i][1] = formatNOP(data[i][1]); // Format NOP
                data[i][8] = formatNOP(data[i][8]); // Format NOP baru
              }
            }

            // Ambil tabel (<table>)
            const tables = htmlContent.querySelectorAll("table");
            tables.forEach((table) => {
              table.querySelectorAll("tr").forEach((row) => {
                const rowData = [];
                let colIndex = 0;
                row.querySelectorAll("th, td").forEach((cell) => {
                  rowData.push(cell.textContent);
                  colIndex++;
                  if (cell.id == "klas1") {
                    for (let i = 1; i < 4; i++) {
                      rowData.push("");
                    }
                  }
                  if (cell.id == "data_lama" || cell.id == "data_baru") {
                    for (let i = 1; i < 7; i++) {
                      rowData.push("");
                    }
                  }
                  if (cell.id == "totalLB") {
                    for (let i = 1; i < 3; i++) {
                      rowData.push("");
                    }
                  }
                  if (cell.id == "kolom_luas") {
                    for (let i = 1; i < 2; i++) {
                      rowData.push("LUAS");
                      rowData.push("ZNT");
                      rowData.push("LUAS");
                    }
                  }

                  if (
                    cell.id == "kolom_bumi1" ||
                    cell.id == "kolom_bang1" ||
                    cell.id == "kolom_bumi2" ||
                    cell.id == "kolom_bang2"
                  ) {
                    rowData.push("");
                    colIndex++;
                  }
                });
                data.push(rowData);
                rowIndex++;
              });
            });

            const ws = XLSX.utils.aoa_to_sheet(data);
            ws["!merges"] = merges;

            const excelData = XLSX.utils.sheet_to_json(ws, { header: 1 });

            // Buat workbook exceljs
            const excelWorkbook = new ExcelJS.Workbook();
            const excelWorksheet = excelWorkbook.addWorksheet("Content");

            // Tambahkan data ke worksheet exceljs
            excelWorksheet.addRows(excelData);

            const petugasPeneliti =
              document.getElementById("petugas_peneliti").textContent;
            const nipPeneliti =
              document.getElementById("nip_peneliti").textContent;
            const kasubbidPendataan =
              document.getElementById("kasubbid_pendataan").textContent;
            const eselonKasubbidPendataan = document.getElementById(
              "eselon_kasubbid_pendataan"
            ).textContent;
            const nipKasubbidPendataan = document.getElementById(
              "nip_kasubbid_pendataan"
            ).textContent;
            const kasubbidPenetapan =
              document.getElementById("kasubbid_penetapan").textContent;
            const eselonKasubbidPenetapan = document.getElementById(
              "eselon_kasubbid_penetapan"
            ).textContent;
            const nipKasubbidPenetapan = document.getElementById(
              "nip_kasubbid_penetapan"
            ).textContent;

            // Tambahkan data ke worksheet Excel
            excelWorksheet.getCell("F" + (data.length + 3)).value =
              "Petugas Peneliti";
            excelWorksheet.getCell("F" + (data.length + 7)).value =
              petugasPeneliti;
            excelWorksheet.getCell("F" + (data.length + 8)).value = nipPeneliti;

            excelWorksheet.getCell("C" + (data.length + 12)).value =
              "Diperiksa Oleh,";
            excelWorksheet.getCell("C" + (data.length + 13)).value =
              "Kasubbid Pendaftaran dan Pendataan";
            excelWorksheet.getCell("C" + (data.length + 17)).value =
              kasubbidPendataan;
            excelWorksheet.getCell("C" + (data.length + 18)).value =
              eselonKasubbidPendataan;
            excelWorksheet.getCell("C" + (data.length + 19)).value =
              nipKasubbidPendataan;

            excelWorksheet.getCell("I" + (data.length + 12)).value =
              "Diperiksa Oleh,";
            excelWorksheet.getCell("I" + (data.length + 13)).value =
              "Kasubbid Perhitungan dan Penetapan";
            excelWorksheet.getCell("I" + (data.length + 17)).value =
              kasubbidPenetapan;
            excelWorksheet.getCell("I" + (data.length + 18)).value =
              eselonKasubbidPenetapan;
            excelWorksheet.getCell("I" + (data.length + 19)).value =
              nipKasubbidPenetapan;

            // Merge & Center semua merge cell manual
            excelWorksheet.mergeCells("M1:O1");
            excelWorksheet.mergeCells("M2:O2");
            excelWorksheet.mergeCells("M3:O3");
            excelWorksheet.mergeCells("A4:O4");
            excelWorksheet.mergeCells("A7:C7");
            excelWorksheet.mergeCells("Q12:T12");
            excelWorksheet.mergeCells("A10:A11");
            excelWorksheet.mergeCells("B10:B11");
            excelWorksheet.mergeCells("C10:C11");
            excelWorksheet.mergeCells("D10:E10");
            excelWorksheet.mergeCells("F10:G10");
            excelWorksheet.mergeCells("K10:L10");
            excelWorksheet.mergeCells("M10:N10");
            excelWorksheet.mergeCells("H10:H11");
            excelWorksheet.mergeCells("I10:I11");
            excelWorksheet.mergeCells("J10:J11");
            excelWorksheet.mergeCells("A9:G9");
            excelWorksheet.mergeCells("H9:N9");
            excelWorksheet.mergeCells("O9:O11");
            excelWorksheet.mergeCells("A13:C13");
            excelWorksheet.mergeCells("H13:J13");
            excelWorksheet.mergeCells("H7:J7");
            excelWorksheet.mergeCells("F16:H16");
            excelWorksheet.mergeCells("F20:H20");
            excelWorksheet.mergeCells("F21:H21");
            excelWorksheet.mergeCells("C25:E25");
            excelWorksheet.mergeCells("C26:E26");
            excelWorksheet.mergeCells("C30:E30");
            excelWorksheet.mergeCells("C31:E31");
            excelWorksheet.mergeCells("C32:E32");
            excelWorksheet.mergeCells("I25:J25");
            excelWorksheet.mergeCells("I26:J26");
            excelWorksheet.mergeCells("I30:J30");
            excelWorksheet.mergeCells("I31:J31");
            excelWorksheet.mergeCells("I32:J32");

            excelWorksheet.getColumn("A").width = 5;
            excelWorksheet.getColumn("B").width = 30;
            excelWorksheet.getColumn("C").width = 20;
            excelWorksheet.getColumn("I").width = 30;
            excelWorksheet.getColumn("J").width = 20;
            excelWorksheet.getColumn("O").width = 30;

            let totalLBRowIndex = 0;
            for (let i = 0; i < data.length; i++) {
              for (let j = 0; j < data[i].length; j++) {
                if (data[i][j] === "TOTAL" && data[i][0] == "TOTAL") {
                  totalLBRowIndex = i + 1;
                  break;
                }
              }
              if (totalLBRowIndex > 0) {
                break;
              }
            }

            // Atur border untuk rentang A9:O(totalLBRowIndex)
            if (totalLBRowIndex > 0) {
              for (let i = 9; i <= totalLBRowIndex; i++) {
                for (let j = 1; j <= 15; j++) {
                  excelWorksheet.getCell(i, j).border = {
                    top: { style: "thin" },
                    left: { style: "thin" },
                    bottom: { style: "thin" },
                    right: { style: "thin" },
                  };
                }
              }
            }

            for (let i = 65; i <= 80; i++) {
              // 65 = A, 80 = P
              const col = String.fromCharCode(i);
              excelWorksheet.getColumn(col).eachCell((cell) => {
                cell.alignment = {
                  wrapText: true,
                  horizontal: "center",
                  vertical: "middle",
                };
              });
            }

            // Simpan file Excel
            excelWorkbook.xlsx.writeBuffer().then(function (buffer) {
              const blob = new Blob([buffer], {
                type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
              });
              const url = window.URL.createObjectURL(blob);
              const anchor = document.createElement("a");
              anchor.href = url;
              anchor.download = "Lampiran LHP.xlsx";
              anchor.click();
            });
          });
      });
    </script>
    <script>
      document
        .getElementById("convertPDF")
        .addEventListener("click", function () {
          const element = document.getElementById("template");
          const style = document.createElement("style");
          style.innerHTML = `
            #nop, #nop_baru {
                width: 250px !important;
                word-wrap: break-word !important;
            }
        `;
          element.appendChild(style);

          html2canvas(element, { scale: 2 }).then(function (canvas) {
            const imgData = canvas.toDataURL("image/png");
            const pdf = new jspdf.jsPDF({
              orientation: "landscape",
              unit: "mm",
              format: "a4", // Menggunakan format A4
            });
            const imgProps = pdf.getImageProperties(imgData);
            const pdfWidth = pdf.internal.pageSize.getWidth();
            const pdfHeight = pdf.internal.pageSize.getHeight();
            let heightLeft = imgProps.height;
            let position = 0;

            pdf.addImage(
              imgData,
              "PNG",
              0,
              position,
              pdfWidth,
              (imgProps.height * pdfWidth) / imgProps.width
            );
            heightLeft -= pdfHeight;

            const margin = 10;

            // Hitung lebar dan tinggi gambar yang disesuaikan
            const imgWidthAdjusted = pdfWidth - 2 * margin;
            const imgHeightAdjusted =
              (imgProps.height * imgWidthAdjusted) / imgProps.width;

            pdf.save("Lampiran LHP.pdf");
            element.removeChild(style);
          });
        });
    </script>
  </body>
</html>
