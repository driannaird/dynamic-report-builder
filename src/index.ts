import express, { Request, Response } from "express";
import ExcelJS from "exceljs";
import path from "path";
import fs from "fs";

const app = express();
const port = 3000;

// Middleware
app.use(express.json());

app.get("/export", async (req: Request, res: Response) => {
  try {
    // Load workbook dari file template
    const workbook = new ExcelJS.Workbook();
    const templatePath = path.join(process.cwd(), "templates/template.xlsx");

    // Cek apakah file template ada
    if (!fs.existsSync(templatePath)) {
      return res.status(404).send("Template file tidak ditemukan");
    }

    await workbook.xlsx.readFile(templatePath);

    const worksheet = workbook.getWorksheet(1); // ambil sheet pertama

    if (!worksheet) {
      return res.status(500).send("Worksheet tidak ditemukan");
    }

    // Contoh data yang akan diisi
    const data: [string, number, number][] = [
      ["Produk A", 10, 20000],
      ["Produk B", 5, 50000],
      ["Produk C", 8, 30000],
      ["Produk D", 15, 25000],
      ["Produk E", 7, 40000],
    ];

    let rowIndex = 4; // Mulai dari baris ke-3
    data.forEach((item) => {
      const row = worksheet.getRow(rowIndex);
      row.getCell(1).value = item[0]; // Kolom A - Nama Produk
      row.getCell(2).value = item[1]; // Kolom B - Quantity
      row.getCell(3).value = item[2]; // Kolom C - Harga
      row.commit();
      rowIndex++;
    });

    // Tambahkan total di baris terakhir menggunakan rumus Excel
    const totalRow = worksheet.getRow(rowIndex + 1);
    totalRow.getCell(1).value = "TOTAL";
    // Menggunakan rumus SUMPRODUCT untuk menghitung total (quantity * harga)
    totalRow.getCell(3).value = {
      formula: `SUMPRODUCT(B4:B${rowIndex},C4:C${rowIndex})`,
    };
    totalRow.commit();

    // Set response header untuk download
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    res.setHeader(
      "Content-Disposition",
      "attachment; filename=laporan-penjualan.xlsx"
    );

    // Simpan dan kirim file ke client
    await workbook.xlsx.write(res);
    res.end();
  } catch (error) {
    console.error("Export error:", error);
    res.status(500).send("Gagal membuat file Excel");
  }
});

// Route untuk export dengan data custom
app.post("/export-custom", async (req: Request, res: Response) => {
  try {
    const { data, filename } = req.body;

    if (!data || !Array.isArray(data)) {
      return res.status(400).send("Data harus berupa array");
    }

    const workbook = new ExcelJS.Workbook();
    const templatePath = path.join(process.cwd(), "templates/template.xlsx");

    if (!fs.existsSync(templatePath)) {
      return res.status(404).send("Template file tidak ditemukan");
    }

    await workbook.xlsx.readFile(templatePath);
    const worksheet = workbook.getWorksheet(1);

    if (!worksheet) {
      return res.status(500).send("Worksheet tidak ditemukan");
    }

    let rowIndex = 3;
    data.forEach((item: any[]) => {
      const row = worksheet.getRow(rowIndex);
      row.getCell(1).value = item[0];
      row.getCell(2).value = item[1];
      row.getCell(3).value = item[2];
      row.commit();
      rowIndex++;
    });

    // Hitung total menggunakan rumus Excel
    const totalRow = worksheet.getRow(rowIndex + 1);
    totalRow.getCell(1).value = "TOTAL";
    // Menggunakan rumus SUMPRODUCT untuk menghitung total (quantity * harga)
    totalRow.getCell(3).value = {
      formula: `SUMPRODUCT(B3:B${rowIndex},C3:C${rowIndex})`,
    };
    totalRow.commit();

    const downloadFilename = filename || "laporan-custom.xlsx";
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    res.setHeader(
      "Content-Disposition",
      `attachment; filename=${downloadFilename}`
    );

    await workbook.xlsx.write(res);
    res.end();
  } catch (error) {
    console.error("Export custom error:", error);
    res.status(500).send("Gagal membuat file Excel custom");
  }
});

// Route untuk melihat info template
app.get("/template-info", async (req: Request, res: Response) => {
  try {
    const templatePath = path.join(process.cwd(), "templates/template.xlsx");

    if (!fs.existsSync(templatePath)) {
      return res.status(404).json({ error: "Template file tidak ditemukan" });
    }

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(templatePath);
    const worksheet = workbook.getWorksheet(1);

    if (!worksheet) {
      return res.status(500).json({ error: "Worksheet tidak ditemukan" });
    }

    const info = {
      sheetName: worksheet.name,
      rowCount: worksheet.rowCount,
      columnCount: worksheet.columnCount,
      title: worksheet.getCell("A1").value,
      templatePath: templatePath,
    };

    res.json(info);
  } catch (error) {
    console.error("Template info error:", error);
    res.status(500).json({ error: "Gagal membaca info template" });
  }
});

// Route utama
app.get("/", (req: Request, res: Response) => {
  res.json({
    message: "Excel Export Server with Template",
    endpoints: {
      "GET /export": "Download Excel dengan data default",
      "POST /export-custom":
        "Download Excel dengan data custom (kirim data dalam body)",
      "GET /template-info": "Lihat informasi template",
    },
    example: {
      "POST /export-custom": {
        data: [
          ["Produk X", 5, 15000],
          ["Produk Y", 3, 25000],
        ],
        filename: "laporan-saya.xlsx",
      },
    },
  });
});

app.listen(port, () => {
  console.log(`🚀 Server running at http://localhost:${port}`);
  console.log(`📊 Endpoints available:`);
  console.log(`   GET  /export        - Download Excel dengan data default`);
  console.log(`   POST /export-custom - Download Excel dengan data custom`);
  console.log(`   GET  /template-info - Info template Excel`);
});
