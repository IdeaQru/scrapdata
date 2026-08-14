async function exportToExcel(totalPages = 9) {
  console.log("%c[POSSI Exporter] Memuat library XLSX...", "color: #007bff; font-weight: bold;");
  
  // 1. Load library SheetJS otomatis di browser
  if (typeof XLSX === 'undefined') {
    await new Promise((resolve, reject) => {
      const script = document.createElement('script');
      script.src = "https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js";
      script.onload = resolve;
      script.onerror = () => reject(new Error("Gagal memuat pustaka XLSX"));
      document.head.appendChild(script);
    });
  }

  console.log("%c[POSSI Exporter] Memulai pengambilan data 9 halaman...", "color: #007bff; font-weight: bold;");

  let allRows = [];
  let headersExtracted = false;

  for (let page = 1; page <= totalPages; page++) {
    console.log(`Mengambil data halaman ${page} dari ${totalPages}...`);
    const url = `https://possijatim.my.id/nias/existing?page=${page}`;
    
    try {
      const res = await fetch(url);
      const html = await res.text();
      const doc = new DOMParser().parseFromString(html, 'text/html');
      const table = doc.querySelector('table');

      if (!table) {
        console.warn(`Halaman ${page}: Tabel tidak ditemukan.`);
        continue;
      }

      // Ambil Header Tabel (hanya dari halaman 1)
      if (!headersExtracted) {
        const ths = table.querySelectorAll('thead th, tr:first-child th');
        if (ths.length > 0) {
          const headers = Array.from(ths).map(th => th.innerText.trim().replace(/\s+/g, ' '));
          allRows.push(headers);
          headersExtracted = true;
        }
      }

      // Ambil Baris Data
      const trs = table.querySelectorAll('tbody tr');
      const rows = trs.length > 0 ? trs : table.querySelectorAll('tr:not(:first-child)');

      rows.forEach(tr => {
        const tds = tr.querySelectorAll('td');
        if (tds.length > 0) {
          const rowData = Array.from(tds).map(td => td.innerText.trim().replace(/\s+/g, ' '));
          allRows.push(rowData);
        }
      });
    } catch (err) {
      console.error(`Gagal mengambil data halaman ${page}:`, err);
    }
  }

  if (allRows.length === 0) {
    alert("Gagal mengambil data atau tabel tidak ditemukan.");
    return;
  }

  // 2. Buat File Excel (.xlsx) Murni
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(allRows);
  XLSX.utils.book_append_sheet(wb, ws, "Data Existing");
  
  // 3. Auto Download File .xlsx
  XLSX.writeFile(wb, "data_possi_existing_page1_9.xlsx");
  console.log("%c[POSSI Exporter] BERHASIL! File data_possi_existing_page1_9.xlsx telah diunduh.", "color: #28a745; font-weight: bold;");
}

// Jalankan otomatis
exportToExcel(9);
