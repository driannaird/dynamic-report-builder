# Excel Export Node.js dengan Template

Project Express.js untuk export Excel menggunakan template `.xlsx` dengan library ExcelJS.

## 🚀 Fitur

- ✅ Export Excel berdasarkan template `.xlsx`
- ✅ Data dinamis dari API
- ✅ Custom data melalui POST request
- ✅ Auto calculate total
- ✅ Professional formatting dari template
- ✅ TypeScript support

## 📁 Struktur Project

```
excel-export-node/
├── templates/
│   └── template.xlsx          # Template Excel
├── src/
│   └── index.ts              # Main server file
├── dist/                     # Compiled JavaScript
├── package.json
├── tsconfig.json
└── README.md
```

## 🔧 Installation

1. **Clone atau download project**
2. **Install dependencies:**
   ```bash
   npm install
   ```

## 🏃‍♂️ Menjalankan Server

### Development (with auto-reload)

```bash
npm run dev
```

### Production

```bash
npm run build
npm start
```

Server akan berjalan di: `http://localhost:3000`

## 📊 API Endpoints

### 1. GET `/export`

Download Excel dengan data default

**Response:** File Excel dengan nama `laporan-penjualan.xlsx`

### 2. POST `/export-custom`

Download Excel dengan data custom

**Body:**

```json
{
  "data": [
    ["Produk X", 5, 15000],
    ["Produk Y", 3, 25000]
  ],
  "filename": "laporan-saya.xlsx"
}
```

**Response:** File Excel dengan nama custom

### 3. GET `/template-info`

Melihat informasi template Excel

**Response:**

```json
{
  "sheetName": "Sheet1",
  "rowCount": 10,
  "columnCount": 16,
  "title": "Laporan Penjualan",
  "templatePath": "/path/to/template.xlsx"
}
```

### 4. GET `/`

Info API dan dokumentasi endpoints

## 🧪 Testing API

### Menggunakan curl

**Download Excel default:**

```bash
curl -o laporan.xlsx http://localhost:3000/export
```

**Download Excel custom:**

```bash
curl -X POST http://localhost:3000/export-custom \
  -H "Content-Type: application/json" \
  -d '{
    "data": [
      ["iPhone 14", 2, 15000000],
      ["Samsung S23", 1, 12000000]
    ],
    "filename": "penjualan-handphone.xlsx"
  }' \
  -o penjualan-handphone.xlsx
```

**Cek info template:**

```bash
curl http://localhost:3000/template-info
```

### Menggunakan Postman

1. **GET Request ke** `http://localhost:3000/export`

   - Set response type ke "Send and Download"

2. **POST Request ke** `http://localhost:3000/export-custom`
   - Body type: raw JSON
   - Content:
     ```json
     {
       "data": [
         ["Laptop ASUS", 3, 8500000],
         ["Mouse Logitech", 5, 250000]
       ],
       "filename": "electronics.xlsx"
     }
     ```

## 📋 Format Data

Data yang dikirim harus dalam format array 2 dimensi:

```javascript
[
  [nama_produk, quantity, harga],
  [nama_produk, quantity, harga],
  // ...
];
```

**Contoh:**

```javascript
[
  ["Laptop", 2, 8000000], // String, Number, Number
  ["Mouse", 5, 150000],
  ["Keyboard", 3, 350000],
];
```

## 🎨 Kustomisasi Template

1. **Buka file** `templates/template.xlsx`
2. **Edit sesuai kebutuhan:**
   - Cell A1: Judul laporan
   - Row 2: Header kolom (opsional)
   - Row 3 ke bawah: Area data (akan diisi otomatis)
3. **Simpan template**
4. **Restart server**

## 🔧 Konfigurasi

### Mengubah Port Server

Edit file `src/index.ts`:

```typescript
const port = 3001; // Ganti port di sini
```

### Mengubah Lokasi Template

Edit path template di `src/index.ts`:

```typescript
const templatePath = path.join(process.cwd(), "templates/my-template.xlsx");
```

## 🛠️ Development

### Scripts tersedia:

```bash
npm run dev      # Development dengan auto-reload
npm run build    # Compile TypeScript
npm start        # Jalankan production
npm run watch    # Development dengan nodemon
```

### Dependencies:

- **express**: Web framework
- **exceljs**: Excel manipulation library
- **typescript**: TypeScript compiler
- **@types/express**: TypeScript types untuk Express

## 📝 Troubleshooting

### Template tidak ditemukan

- Pastikan file `template.xlsx` ada di folder `templates/`
- Cek nama file dan ekstensi yang benar

### Error saat export

- Cek format data yang dikirim
- Pastikan array data berisi numbers untuk quantity dan harga

### Port sudah digunakan

- Ganti port di `src/index.ts`
- Atau kill process yang menggunakan port 3000

## 🎯 Use Cases

- **Laporan penjualan bulanan**
- **Invoice otomatis**
- **Export data inventory**
- **Laporan keuangan**
- **Data transaksi**

## 📄 License

MIT License - bebas digunakan untuk project pribadi maupun komersial.
