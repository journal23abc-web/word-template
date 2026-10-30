# Template Paste — Add-in Word

Tempel isi artikel per bagian (Introduction, Method, Results and Discussion, Conclusion, References, …), lalu add-in menyesuaikan formatnya dengan template jurnal secara otomatis:

- **Font, ukuran, perataan, spasi sebelum/sesudah, spasi baris** untuk tiap jenis paragraf
- **Deteksi sub-heading otomatis** (kalimat pendek tanpa titik, teks bold, heading dari Google Docs/ChatGPT, heading bernomor seperti `2.1`) → dibuat bold sesuai template. Di Results and Discussion, *Results* dan *Discussion* jadi tingkat 2, sub-judul di bawahnya tingkat 3
- **Referensi**: hanging indent 1,27 cm, italic otomatis gaya APA (nama jurnal + volume, atau judul buku) bila tempelan tidak membawa italic, DOI jadi hyperlink, baris terputus dari PDF disambung
- **Caption** (`Table 1.` / `Figure 1.` → label bold, judul normal; caption gaya APA dua baris "Figure 1" + judul italic otomatis digabung, dan italic pada judul caption dihilangkan), **Note.**, dan **tabel** yang ditempel (garis gaya APA, font 10 pt)
- **Menambah, bukan mengganti**: isi masuk di bawah konten yang sudah ada pada bagian itu. Bila judul bagian belum ada, judulnya dibuat di urutan yang benar
- **Abstrak**: kartu ABSTRACT mengganti teks di kotak abstrak (bukan menambah). AIEDS memakai format terstruktur: *Purpose – / Methodology – / Findings – / Research limitations – / Originality –*. Label dikenali otomatis dari tempelan (juga *Design/methodology/approach:*, *Tujuan:*, dll.), dibuat **bold**, pemisahnya diseragamkan menjadi " – ", dan komponennya disatukan dalam satu paragraf yang dipisah *line break*, 10 pt rata kiri. JEPSI memakai satu paragraf 9 pt justify. Baris *Keywords* di tempelan dilewati. Abstrak lama bisa dikembalikan, dan jumlah kata dibandingkan dengan batas abstrak jurnal
- **Cek kesesuaian** (tab **Cek**): mencocokkan *To cite this article*, baik di badan dokumen, tabel, maupun footer, dengan artikelnya. Yang dicek: penulis (nama, urutan, inisial), judul, nama jurnal, volume/nomor, tahun, halaman (dibandingkan dengan nomor halaman awal dokumen), DOI (format, kelengkapan, kecocokan vX iY dan DOI di header), serta tanda baca (titik ganda, spasi ganda, placeholder). Juga memeriksa **article history**: tanggal terbaca dan valid, urutan Received ≤ Revised ≤ Accepted, dan jarak Received → Accepted **maks. 3 bulan** (bisa diubah per jurnal), tidak melewati bulan terbit, serta tanggal di masa depan. Selain itu memeriksa header/footer (placeholder "Author", "Title…", "XXXX"; running head penulis dan judul; nama jurnal) dan batas kata
- **Hapus isi bagian**: kosongkan seluruh isi satu bagian di dokumen (atau semua bagian sekaligus) dengan judul bagian tetap ada, misalnya untuk mengganti Introduction dengan versi baru. Ada konfirmasi dan tombol **Kembalikan isi yang dihapus** (selama panel masih terbuka), dan Ctrl+Z tetap bisa dipakai
- **Rapikan isi yang ada**: terapkan format template ke paragraf yang sudah terlanjur ditempel manual (italic tidak diubah), **termasuk tabel** di bagian itu: garis vertikal dihapus, garis gaya APA (atas, bawah judul kolom, bawah tabel), font & ukuran tabel, judul kolom bold, perataan sel
- **Kutipan langsung** (mis. kutipan wawancara “…” Guru, S1) dikenali dan dibuat menjorok + italic
- **Hitung kata** per bagian dan total terhadap batas jurnal (aturan bisa diatur per jurnal), plus abstrak
- **Profil template**: semua angka bisa diubah, dan profil bisa diduplikat untuk jurnal lain (ekspor/impor JSON)

**Template bawaan** (pilih di menu **Jurnal** di bagian atas tab Tempel):

| | AIEDS | JEPSI | VOICE |
|---|---|---|---|
| Font | Cambria | Cambria | Cambria |
| Judul bagian | 12 pt bold KAPITAL, tengah, sesudah 6 pt | 11 pt bold KAPITAL, tengah, sesudah 6 pt | 12 pt bold KAPITAL, tengah, 6/6 pt |
| Sub-heading | 11 pt bold, 3/3 pt | 11 pt bold, 0/6 pt | 11 pt bold, 6/6 pt |
| Isi | 11 pt justify, 3/3 pt, spasi tunggal | 11 pt justify, 0/6 pt, spasi tunggal | 11 pt justify, 6/6 pt |
| Caption / Note. | label bold / Note. tegak | label bold / *Note.* italic | label bold; tabel 10 pt, gambar 11 pt |
| Abstrak | terstruktur, label bold " – ", 10 pt rata kiri, 1 paragraf dengan line break | 1 paragraf, 9 pt justify, sesudah 8 pt, spasi 1,08 | terstruktur, paragraf terpisah, hanya label bold, 10 pt justify |
| Kutipan langsung | menjorok 1,27 cm, italic | menjorok 1,27 cm, italic | menjorok 1,27 cm, italic |
| Tabel | 10 pt, garis APA, isi rata tengah | 10 pt, garis APA, isi rata kiri | 10 pt, garis APA, isi rata kiri |
| Referensi | 11 pt, hanging 1,27 cm, sesudah 6 pt | 11 pt, hanging 1,27 cm, sesudah 6 pt | 11 pt, hanging 0,83 cm, sesudah 4 pt, tanpa italic |

**Batas kata diatur sendiri per jurnal** di tab **Hitung kata → Aturan hitung kata**: maksimal/minimal kata isi, maksimal kata abstrak, bagian mana yang ikut dihitung (mis. referensi, Author Contribution, AI Disclosure), dan teks tabel. 0 = tidak dibatasi.

---

## 1. Hosting (sekali saja, gratis lewat GitHub Pages)

Add-in Word harus dibuka dari alamat **https**. Cara termudah:

1. Buat repositori baru di GitHub, misalnya `template-paste` (boleh *Public*).
2. Unggah **seluruh isi folder ini** (`manifest.xml`, `README.html`, folder `src/`, `assets/`) ke repositori itu.
3. Buka **Settings → Pages**, lalu pada *Source* pilih **Deploy from a branch**, branch `main`, folder `/ (root)`, dan klik **Save**.
4. Tunggu 1–2 menit. Alamat situs Anda akan seperti `https://NAMAANDA.github.io/template-paste`.
5. Buka `manifest.xml` di Notepad, lalu gunakan **Find & Replace**: ganti semua `https://YOUR-SITE` dengan alamat tadi (tanpa `/` di akhir). Simpan file itu di komputer Anda. (File di GitHub tidak perlu diubah.)

Cek: buka `https://NAMAANDA.github.io/template-paste/src/taskpane.html` di browser. Panel akan tampil dengan tulisan "Mode pratinjau".

## 2. Memasang di Word

### Word di Windows (Microsoft 365 / 2019 / 2021)
1. Buat folder, misalnya `C:\AddinWord`, lalu salin `manifest.xml` ke folder itu.
2. Klik kanan folder → **Properties → Sharing → Share…** → tambahkan diri Anda → **Share**. Catat alamat jaringannya, misalnya `\\NAMA-PC\AddinWord`.
3. Di Word: **File → Options → Trust Center → Trust Center Settings → Trusted Add-in Catalogs**. Isi *Catalog Url* dengan `\\NAMA-PC\AddinWord`, klik **Add catalog**, centang **Show in Menu**, klik OK, lalu tutup dan buka Word lagi.
4. **Home → Add-ins** (atau **Insert → My Add-ins**) → tab **SHARED FOLDER** → **Template Paste** → **Add**.

### Word di Mac
1. Buka Finder → **Go → Go to Folder…** → `~/Library/Containers/com.microsoft.Word/Data/Documents/wef` (buat folder `wef` bila belum ada).
2. Salin `manifest.xml` ke folder itu, lalu buka ulang Word.
3. **Insert → Add-ins → My Add-ins** → pilih **Template Paste**.

### Word di web (office.com)
**Insert → Add-ins → My Add-ins → Upload My Add-in** → pilih `manifest.xml`.

Setelah terpasang, tombol **Template Paste** muncul di tab **Home**.

---

## 3. Cara pakai

1. Buka naskah template jurnal (yang sudah ada judul INTRODUCTION, METHOD, dst., atau masih kosong pun bisa).
2. Klik **Template Paste** di tab Home, lalu pilih jurnalnya di menu **Jurnal**.
3. Di tab **Tempel**, buka kartu bagian (misalnya *INTRODUCTION*), klik kotak bergaris putus-putus, lalu tekan **Ctrl+V**. Anda bisa menempel berkali-kali untuk menambah isi.
4. Periksa pratinjau. Tiap baris punya menu jenis (*Isi / Heading 2 / Heading 3 / Kutipan / Caption / Catatan / Referensi / Lewati*). Ubah bila deteksi keliru.
5. Klik **Sisipkan ke dokumen**, atau **Sisipkan semua bagian** di bawah untuk memasukkan semua sekaligus.
6. Bila hasilnya tidak sesuai, klik **Hapus sisipan terakhir** (atau Ctrl+Z).
7. Untuk **mengganti** isi suatu bagian: klik **Hapus isi bagian ini di dokumen** di kartu bagian itu, lalu tempel isi barunya. Judul bagian tidak ikut terhapus. Perhatikan bahwa *semua* isi di antara judul itu dan judul bagian berikutnya ikut terhapus, termasuk tabel, gambar, dan section break.

**Mode "Posisi kursor"**: isi disisipkan di posisi kursor, bukan di akhir bagian.

**Gambar** di tempelan diabaikan. Sisipkan gambar langsung di Word, lalu tempel caption-nya lewat add-in.

**Tips akurasi deteksi**: menyalin dari Word, Google Docs, atau ChatGPT paling akurat karena format bold/italic ikut terbawa. Untuk teks polos (Notepad/PDF), sub-heading dikenali dari ciri "pendek, diawali huruf kapital, tanpa titik di akhir".

## 4. Template untuk jurnal lain

Tab **Template** → **Duplikat**, beri nama (misalnya "Jurnal X"), lalu ubah ukuran, spasi, hanging, kutipan, dan **Daftar bagian artikel** (judul bagian beserta nama lainnya, misalnya *Pendahuluan*, *Metode Penelitian*). Profil tersimpan di komputer itu. Gunakan **Ekspor** untuk menyalin profil ke komputer lain.

## 5. Struktur file

```
manifest.xml          definisi add-in (ganti YOUR-SITE)
src/taskpane.html     panel
src/taskpane.js       antarmuka
src/engine.js         deteksi heading/referensi/abstrak & aturan format (tanpa Office.js)
src/check.js          pemeriksa kesesuaian (cite, article history, header/footer)
src/word-ops.js       operasi ke dokumen Word (Office.js, WordApi 1.3+)
assets/               ikon
```

Kebutuhan: Word 2016 ke atas / Microsoft 365 / Word Mac / Word web. Fitur "Hapus sisipan terakhir" butuh WordApi 1.4 (Microsoft 365 / Word 2021).

## 6. Catatan versi

**v6.1:**
- Perbaikan error *GeneralException (TableCell.cellIndex)* saat mengganti/mengosongkan abstrak di tabel yang punya sel gabungan (template JEPSI). Add-in tidak lagi memakai nomor baris/kolom sel; sel abstrak dikenali dari isinya.
- Rapikan tabel juga tidak lagi memakai nomor sel, jadi aman untuk tabel bersel gabungan.

**v6 (perbaikan bug):**
- Abstrak yang disalin dari dalam tabel (seperti di template JEPSI/AIEDS/VOICE) sekarang terbaca. Kalau seluruh tabel halaman depan ikut tersalin, *Article info*, *Keywords* beserta butirnya, *Article history*, *cite*, dan lisensi otomatis dilewati.
- Mengganti abstrak yang terdiri dari beberapa paragraf (VOICE) memakai ulang paragraf yang ada, sehingga tidak error di sel tabel dan latar warnanya terjaga.
- Hitung kata abstrak tidak lagi ikut menghitung teks *To cite this article* (JEPSI).
- Panel tetap aktif walaupun Word lambat siap saat pertama dibuka.
- Kompatibel dengan Word for Mac versi lama (tanpa regex *lookbehind*).
- Rapikan: label caption ("Table 1.") dibold, baris persamaan "(1)" tidak diubah.
- Menyisipkan setelah daftar bernomor/berpoin tidak membuat isi baru ikut bernomor.
- Catatan kaki/komentar dari clipboard Word tidak ikut tertempel.
- "Figure 1" yang berdiri sendiri dikenali sebagai caption, bukan sub-heading.
