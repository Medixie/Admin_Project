document.getElementById("laporanForm").addEventListener("submit", function(e) {
  e.preventDefault();

  const nama = document.getElementById("namaAdmin").value;
  const tanggal = document.getElementById("tanggalLaporan").value;
  const mpp = document.getElementById("reportMpp").value;
  const issue = document.getElementById("reportIssue").value;
  const material = document.getElementById("reportMaterial").value;
  const progres = document.getElementById("reportProgres").value;
  const ceklist = document.getElementById("reportCeklist").value;

  const table = document.getElementById("rekapTable").getElementsByTagName("tbody")[0];
  const newRow = table.insertRow();

  newRow.innerHTML = `
    <td>${nama}</td>
    <td>${tanggal}</td>
    <td>${mpp}</td>
    <td>${issue}</td>
    <td>${material}</td>
    <td>${progres}</td>
    <td>${ceklist}</td>
  `;

  document.getElementById("laporanForm").reset();
});

// Export ke Excel
document.getElementById("exportExcel").addEventListener("click", function () {
  let table = document.getElementById("rekapTable");
  let wb = XLSX.utils.table_to_book(table, { sheet: "Laporan Harian" });
  XLSX.writeFile(wb, "LaporanHarian.xlsx");
});
