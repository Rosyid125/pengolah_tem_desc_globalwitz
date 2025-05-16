const xlsx = require("xlsx");
const fs = require("fs");
const path = require("path");
const readlineSync = require("readline-sync");

// --- Fungsi Ekstraksi GSM ---
function extractGsmValue(desc) {
  if (typeof desc !== "string") return "N/A";

  let match;
  match = desc.match(/(\d[\d.,]*)\s*G\/M2/i);
  if (match) return match[1].replace(",", ".");

  match = desc.match(/(\d[\d.,]*)\s*GSM/i);
  if (match) return match[1].replace(",", ".");

  match = desc.match(/(\d[\d.,]*)\s*GR\/M2/i);
  if (match) return match[1].replace(",", ".");

  match = desc.match(/\b(WEIGHT|AVERAGE WEIGHT|BASIS WEIGHT)\s*:?\s*(\d[\d.,]*)\s*G\b/i);
  if (match) return match[2].replace(",", ".");

  match = desc.match(/(\d[\d.,]*)\s*G\s+(KIMLON|TYPE)/i);
  if (match) return match[1].replace(",", ".");

  match = desc.match(/\b(\d[\d.,]*)\s*G\b(?!\s*SM)(?!\s*\/M2)(?!\s*R\/M2)(?!\s*[A-DF-LN-QS-Z])/i);
  if (match) return match[1].replace(",", ".");

  match = desc.match(/(\d[\d.,]*)\s*GR\/YD/i);
  if (match) return match[1].replace(",", ".");

  return "N/A";
}

// --- Fungsi Ekstraksi Width ---
function extractWidthValue(desc) {
  if (typeof desc !== "string") return "N/A";

  let match;
  match = desc.match(/(\d+[\d.,]*)\s*(inch|inches|\")/i);
  if (match) {
    const inches = parseFloat(match[1].replace(",", "."));
    return (inches * 2.54).toFixed(2); // Konversi ke cm
  }

  match = desc.match(/(\d+[\d.,]*)\s*cm/i);
  if (match) return match[1].replace(",", ".");

  match = desc.match(/(\d+[\d.,]*)\s*mm/i);
  if (match) {
    const mm = parseFloat(match[1].replace(",", "."));
    return (mm / 10).toFixed(2); // Konversi ke cm
  }

  return "N/A";
}

// --- Fungsi Utama ---
async function processExcelFile(inputFilePath, outputFilePath) {
  try {
    if (!fs.existsSync(inputFilePath)) {
      console.error(`File input tidak ditemukan: ${inputFilePath}`);
      return;
    }

    const workbook = xlsx.readFile(inputFilePath);
    const sheetNames = workbook.SheetNames;
    console.log("\nSheet yang tersedia:");
    sheetNames.forEach((name, index) => {
      console.log(`${index + 1}. ${name}`);
    });

    const selectedSheetIndexes = readlineSync
      .question("\nMasukkan nomor sheet yang ingin diproses (pisahkan dengan koma jika lebih dari satu, misal: 1,3,5): ")
      .split(",")
      .map((idx) => parseInt(idx.trim()) - 1)
      .filter((idx) => idx >= 0 && idx < sheetNames.length);

    if (selectedSheetIndexes.length === 0) {
      console.error("Tidak ada sheet yang dipilih.");
      return;
    }

    selectedSheetIndexes.forEach((sheetIndex) => {
      const sheetName = sheetNames[sheetIndex];
      const worksheet = workbook.Sheets[sheetName];
      const headerRow = xlsx.utils.sheet_to_json(worksheet, { header: 1 })[0] || [];
      const descColumnName = headerRow.find((h) => h && (h.toUpperCase() === "ITEM DESC" || h.toUpperCase() === "PRODUCT DESCRIPTION(EN)"));

      if (!descColumnName) {
        console.error(`Kolom 'ITEM DESC' atau 'PRODUCT DESCRIPTION(EN)' tidak ditemukan di sheet '${sheetName}'.`);
        return;
      }

      const jsonData = xlsx.utils.sheet_to_json(worksheet);

      const processedData = jsonData.map((row) => {
        const itemDesc = row[descColumnName] || "";

        const gsm = extractGsmValue(itemDesc);
        const width = extractWidthValue(itemDesc);

        let newRow = { ...row };
        newRow["GSM"] = gsm;
        newRow["Width (cm)"] = width;

        return newRow;
      });

      const outputHeaders = Object.keys(jsonData[0] || [])
        .concat(["GSM", "Width (cm)"])
        .filter((v, i, a) => a.indexOf(v) === i);

      const newWorkbook = xlsx.utils.book_new();
      const newWorksheet = xlsx.utils.json_to_sheet(processedData, { header: outputHeaders });

      xlsx.utils.book_append_sheet(newWorkbook, newWorksheet, sheetName);
      const finalOutputPath = path.join(__dirname, `output.xlsx`);
      xlsx.writeFile(newWorkbook, finalOutputPath);
      console.log(`Sheet '${sheetName}' berhasil diproses dan disimpan sebagai ${finalOutputPath}`);
    });
  } catch (error) {
    console.error("Terjadi kesalahan saat memproses file:", error);
  }
}

const inputFileName = "input.xlsx";
const inputFilePath = path.join(__dirname, inputFileName);
processExcelFile(inputFilePath);
