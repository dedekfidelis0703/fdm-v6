# FDM Extractor

Aplikasi web untuk ekstraksi data dari file Excel FDM (Formulir Data Masuk) ke format Excel dengan perhitungan NJOP dan PBB otomatis.

## Fitur

- ✅ **Upload Multiple Files** - Upload hingga 50 file Excel sekaligus
- ✅ **Ekstraksi Otomatis** - Ekstrak data dengan logika yang sama persis dengan script Python
- ✅ **Download Hasil** - File hasil ekstraksi didownload secara otomatis
- ✅ **Download Ulang** - Tombol untuk mendownload ulang hasil ekstraksi
- ✅ **Ekstraksi Baru** - Mulai ekstraksi baru dengan mudah
- ✅ **Privasi Terjaga** - File tidak disimpan di server

## Cara Penggunaan

1. **Upload FDM** - Klik tombol "Upload FDM" dan pilih file Excel (.xlsx, .xls, .xlsm)
2. **Ekstrak Sekarang** - Klik tombol "Ekstrak Sekarang" untuk memulai ekstraksi
3. **Download Ulang** - Jika file belum terdownload, klik "Download Ulang Hasil Ekstraksi"
4. **Ekstraksi Lain** - Klik "Ekstraksi FDM Lain" untuk memulai proses baru

## Teknologi

- React + TypeScript
- Vite
- Tailwind CSS
- shadcn/ui
- SheetJS (xlsx) untuk pemrosesan Excel
- file-saver untuk download file

## Deploy ke Vercel

### Opsi 1: Deploy dari GitHub

1. Buat repository baru di GitHub
2. Upload semua file project ini ke repository
3. Login ke [Vercel](https://vercel.com) dengan akun GitHub
4. Klik "Add New Project"
5. Pilih repository FDM Extractor
6. Klik "Deploy"

### Opsi 2: Deploy dengan Vercel CLI

```bash
# Install Vercel CLI
npm i -g vercel

# Login ke Vercel
vercel login

# Deploy
vercel --prod
```

## Struktur Project

```
fdm-extractor/
├── src/
│   ├── components/ui/     # Komponen UI (shadcn)
│   ├── sections/          # Komponen halaman
│   ├── types/             # TypeScript types
│   ├── utils/             # Utility functions
│   ├── App.tsx            # Root component
│   └── main.tsx           # Entry point
├── dist/                  # Build output
├── index.html
├── package.json
├── tailwind.config.js
├── tsconfig.json
└── vite.config.ts
```

## Konversi dari Python

Logika ekstraksi di file `src/utils/fdmExtractor.ts` adalah konversi langsung dari script Python `Ekstrak_FDM_V6.py` dengan:

- Definisi item yang sama persis
- Fungsi pencarian sheet pintar (`getSheetSmart`)
- Ekstraksi data statis dan dinamis
- Generate Excel dengan rumus-rumus yang identik
- Sheet "1. Hasil" dan "2. Kesimpulan" dengan formula yang sama

## Lisensi

Private - Untuk penggunaan internal
