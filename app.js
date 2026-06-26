// Pindahkan toggleDetail ke scope global agar bisa dibaca oleh atribut onclick HTML
function toggleDetail(id) {
    const el = document.getElementById(`detail-${id}`);
    if (el) {
        el.style.display = (el.style.display === "none") ? "block" : "none";
    }
}

document.addEventListener("DOMContentLoaded", function () {

    let fuse;
    let dataKBLI = [];

    const kategoriMap = {
        "esensial": "kategori-esensial",
        "umum ditemukan": "kategori-umum",
        "potensial": "kategori-potensial"
    };
        
    async function init() {
        try {
            const response = await fetch("data/kbli_populer_riau2.xlsx");
            
            if (!response.ok) {
                throw new Error("File Excel tidak ditemukan atau gagal dimuat.");
            }

            const arrayBuffer = await response.arrayBuffer();
            const workbook = XLSX.read(arrayBuffer, { type: "array" });
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            const rows = XLSX.utils.sheet_to_json(worksheet);
            

            dataKBLI = rows.map(row => ({
                kbli1: row.kbli1 || "",
                kbli5: String(row.kbli5 || ""),
                judul: row.judul || "",
                kategori: row.kategori || "",
                contoh: row.contoh_usaha
                    ? String(row.contoh_usaha).split(";").map(x => x.trim()).filter(Boolean)
                    : [],
                keyword: row.kata_kunci
                    ? String(row.kata_kunci).split(";").map(x => x.trim()).filter(Boolean)
                    : []
            }));

            buatFuse();
            
            // PERBAIKAN: Jangan panggil tampilkan(dataKBLI) di sini agar dashboard kosong saat awal dibuka.
            // Sembunyikan text jumlah hasil di awal
            document.getElementById("result-count").innerHTML = ""; 

        } catch (error) {
            console.error(error);
            document.getElementById("results").innerHTML =
                `<p style="color: red;">Gagal membaca file Excel. Error: ${error.message}</p>`;
        }
    }

    function buatFuse() {
        fuse = new Fuse(dataKBLI, {
            includeScore: true,
            threshold: 0.35,
            ignoreLocation: true,
            keys: [
                { name: "kbli5", weight: 1 },
                { name: "judul", weight: 5 },
                { name: "contoh", weight: 5 },
                { name: "keyword", weight: 5 }
            ]
        });
    }

    function tampilkan(data) {
        const hasil = document.getElementById("results");
        const countEl = document.getElementById("result-count");

        // Jika data kosong (tidak ada kata kunci / tidak ada hasil match)
        if (data.length === 0) {
            countEl.innerHTML = "Tidak ada hasil yang ditemukan";
            hasil.innerHTML = "";
            return;
        }

        countEl.innerHTML = `Ditemukan ${data.length} hasil`;
        hasil.innerHTML = "";

        data.forEach((item, index) => {
            const kategoriClass =
                kategoriMap[String(item.kategori).trim().toLowerCase()] || "kategori-umum";

            hasil.innerHTML += `
                <div class="card">
                    <div class="kbli">${item.kbli5}</div>
                    <div><b>KBLI 1 Digit:</b> ${item.kbli1}</div>
                    <div class="judul">${item.judul}</div>
                    <div class="kategori ${kategoriClass}">
                        ${item.kategori}</div>

                    <button class="detail-button" onclick="toggleDetail(${index})">
                        Lihat Detail
                    </button>

                    <div id="detail-${index}" class="detail" style="display:none;">
                        <hr>
                        <b>Contoh Usaha</b>
                        <ul>${
                            item.contoh.length
                                ? item.contoh.map(x => `<li>${x}</li>`).join("")
                                : "<li>Tidak tersedia</li>"
                        }</ul>

                        <b>Kata Kunci</b>
                        <ul>${
                            item.keyword.length
                                ? item.keyword.map(x => `<li>${x}</li>`).join("")
                                : "<li>Tidak tersedia</li>"
                        }</ul>
                    </div>
                </div>
            `;
        });
    }

    function cariData() {
        const keyword = document.getElementById("search").value.trim();

        // PERBAIKAN: Jika input kosong, kosongkan layar kembali (jangan tampilkan semua data)
        if (!keyword) {
            document.getElementById("result-count").innerHTML = "";
            document.getElementById("results").innerHTML = "";
            return;
        }

        const result = fuse.search(keyword);
        const hasilPencarian = result.slice(0, 20).map(x => x.item);

        tampilkan(hasilPencarian);
    }

    // Event Listener untuk Tombol dan Input
    document.getElementById("btnCari")?.addEventListener("click", cariData);
    
    // Jika Anda ingin pencarian otomatis saat mengetik tetap aktif, biarkan baris ini.
    // Jika ingin hasil HANYA keluar saat tombol Cari/Enter ditekan, hapus baris di bawah ini:
    document.getElementById("search")?.addEventListener("input", cariData);

    document.getElementById("search")?.addEventListener("keypress", function (e) {
        if (e.key === "Enter") cariData();
    });

    init();
});