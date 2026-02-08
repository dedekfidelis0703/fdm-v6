import * as XLSX from 'xlsx';
import type { FDMItemDefinition, ExtractionResult } from '@/types/fdm';

// Item definitions - same as Python script
export const itemsDefinitions: FDMItemDefinition[] = [
  // --- BAGIAN AWAL ---
  { label: "KPP", sheet: "Sheet Home", addr: "H5", mode: "Static" },
  { label: "Sektor", sheet: "Sheet Home", addr: "D8", mode: "Static" },

  // --- BAGIAN 1: Data Umum ---
  { label: "NAMA WAJIB PAJAK", sheet: "Sheet Home", addr: "H15", mode: "Static" },
  { label: "NOMOR OBJEK PAJAK", sheet: "Sheet Home", addr: "H21", mode: "Static" },
  { label: "KELURAHAN", sheet: "Sheet Home", addr: "H27", mode: "Static" },
  { label: "KECAMATAN", sheet: "Sheet Home", addr: "H29", mode: "Static" },
  { label: "KABUPATEN/KOTA", sheet: "Sheet Home", addr: "H31", mode: "Static" },
  { label: "PROVINSI", sheet: "Sheet Home", addr: "H33", mode: "Static" },

  // Rumus Penjumlahan Areal
  { label: "LUAS BUMI", sheet: "Sheet Home", mode: "Formula_LuasBumi" },

  // --- BAGIAN 2: Areal ---
  { label: "Areal Produktif", sheet: "Sheet Home", addr: "J75", mode: "Static" },
  { label: "Areal Belum Diolah", sheet: "Sheet Home", addr: "J77", mode: "Static" },
  { label: "Areal Sudah Diolah Belum Ditanami", sheet: "Sheet Home", addr: "J78", mode: "Static" },
  { label: "Areal Pembibitan", sheet: "Sheet Home", addr: "J79", mode: "Static" },
  { label: "Areal Tidak Produktif", sheet: "Sheet Home", addr: "J80", mode: "Static" },
  { label: "Areal Pengaman", sheet: "Sheet Home", addr: "J81", mode: "Static" },
  { label: "Areal Emplasemen", sheet: "Sheet Home", addr: "J82", mode: "Static" },

  { label: "Areal Produktif (Copy)", sheet: "Sheet Home", mode: "Formula_CopyProduktif" },

  // --- BAGIAN 3: NJOP & Rumus ---
  { label: "NJOP/M Areal Belum Produktif", sheet: "C.1", addr: "BK22", mode: "Static" },

  { label: "NJOP Bumi Berupa Tanah (Rp)", sheet: "Sheet Home", mode: "Formula_NJOPTanah" },

  { label: "NJOP Bumi Berupa Pengembangan Tanah (Rp)", sheet: "C.2", keyword: "Pengembangan Tanah", mode: "Dynamic_Col_G" },

  { label: "NJOP Bumi Berupa Pengembangan Tanah (Rp) (Kenaikan BIT 10.3%)", sheet: "N/A", mode: "Formula_BIT" },

  { label: "NJOP Bumi Areal Produktif (Rp)", sheet: "N/A", mode: "Formula_NJOP_Total" },

  { label: "Luas Bumi Areal Produktif (m²)", sheet: "N/A", mode: "Formula_Luas_Ref" },

  { label: "NJOP Bumi Per M2 Areal Produktif (Rp/m2)", sheet: "N/A", mode: "Formula_NJOP_PerM2" },

  // --- BAGIAN 4: FDM Kebun ABC ---
  { label: "NJOP BUMI (Rp) AREA PRODUKTIF pada A. DATA BUMI", sheet: "N/A", mode: "Formula_Final_Calc" },
  { label: "NJOP BUMI (Rp) AREA PRODUKTIF pada A. DATA BUMI (Proyeksi NDT Naik 46%)", sheet: "N/A", mode: "Formula_Proyeksi" },

  { label: "NJOP BUMI (Rp) AREAL BELUM PRODUKTIF pada A. DATA BUMI", sheet: "FDM Kebun ABC", addr: "E20", mode: "Static" },
  { label: "NJOP BUMI (Rp) AREAL BELUM PRODUKTIF pada A. DATA BUMI (Proyeksi NDT Naik 46%)", sheet: "N/A", mode: "Formula_Proyeksi_BelumProd" },

  // --- Tambahan Copy & NJOP Lainnya ---
  { label: "Areal Tidak Produktif (Copy)", sheet: "N/A", mode: "Formula_CopyTidakProduktif" },
  { label: "NJOP/M Areal Tidak Produktif", sheet: "C.1", addr: "BK62", mode: "Static" },
  { label: "NJOP BUMI (Rp) AREAL TIDAK PRODUKTIF pada A. DATA BUMI", sheet: "N/A", mode: "Formula_Calc_TidakProd" },
  { label: "NJOP BUMI (Rp) AREAL TIDAK PRODUKTIF pada A. DATA BUMI (Proyeksi NDT Naik 46%)", sheet: "N/A", mode: "Formula_Proyeksi_TidakProd" },

  { label: "Areal Pengaman (Copy)", sheet: "N/A", mode: "Formula_CopyPengaman" },
  { label: "NJOP/M Areal Pengaman", sheet: "D", addr: "L23", mode: "Static" },
  { label: "NJOP BUMI (Rp) AREAL PENGAMAN pada A. DATA BUMI", sheet: "N/A", mode: "Formula_Calc_Pengaman" },
  { label: "NJOP BUMI (Rp) AREAL PENGAMAN pada A. DATA BUMI (Proyeksi NDT Naik 46%)", sheet: "N/A", mode: "Formula_Proyeksi_Pengaman" },

  { label: "Areal Emplasemen (Copy)", sheet: "N/A", mode: "Formula_CopyEmplasemen" },
  { label: "NJOP/M Areal Emplasemen", sheet: "C.1", addr: "BK102", mode: "Static" },
  { label: "NJOP BUMI (Rp) AREAL EMPLASEMEN pada A. DATA BUMI", sheet: "N/A", mode: "Formula_Calc_Emplasemen" },
  { label: "NJOP BUMI (Rp) AREAL EMPLASEMEN pada A. DATA BUMI (Proyeksi NDT Naik 46%)", sheet: "N/A", mode: "Formula_Proyeksi_Emplasemen" },

  { label: "JUMLAH Luas (m2) pada A. DATA BUMI", sheet: "N/A", mode: "Formula_Total_Luas_Ref" },
  { label: "JUMLAH NJOP BUMI (Rp) pada A. DATA BUMI", sheet: "N/A", mode: "Formula_Total_NJOP_Sum" },

  { label: "NJOP BUMI (Rp) NJOP Bumi Per Meter Persegi pada A. DATA BUMI", sheet: "FDM Kebun ABC", addr: "E25", mode: "Static" },

  // --- BAGIAN 5: FDM Bangunan ---
  { label: "Jumlah LUAS pada B. DATA BANGUNAN", sheet: "FDM Kebun ABC", mode: "Dynamic_FDM_Bangunan_Luas" },
  { label: "Jumlah NJOP BANGUNAN pada B. DATA BANGUNAN", sheet: "N/A", mode: "Formula_Calc_Bangunan" },
  { label: "NJOP BANGUNAN PER METER PERSEGI*) pada B. DATA BANGUNAN", sheet: "FDM Kebun ABC", mode: "Dynamic_FDM_Bangunan_PerM2" },

  { label: "TOTAL NJOP (TANAH + BANGUNAN) 2025", sheet: "N/A", mode: "Formula_Grand_Total" },

  // --- KOLOM TAMBAHAN SPPT & SIMULASI ---
  { label: "SPPT 2025", sheet: "N/A", mode: "Formula_SPPT_2025" },
  { label: "SIMULASI TOTAL NJOP (TANAH + BANGUNAN) 2026 (Hanya Kenaikan BIT 10,3% + NDT Tetap)", sheet: "N/A", mode: "Formula_Simulasi_NJOP_2026" },
  { label: "SIMULASI SPPT 2026 (Hanya Kenaikan BIT 10,3% + NDT Tetap)", sheet: "N/A", mode: "Formula_Simulasi_SPPT_2026" },
  { label: "Kenaikan", sheet: "N/A", mode: "Formula_Kenaikan" },

  { label: "Persentase", sheet: "N/A", mode: "Formula_Persentase" },
  { label: "SIMULASI TOTAL NJOP (TANAH + BANGUNAN) 2026 (Kenaikan BIT 10,3% + NDT 46%)", sheet: "N/A", mode: "Formula_Simulasi_Total_2026_NDT46" },
  { label: "SIMULASI SPPT 2026 (Kenaikan BIT 10,3% + NDT 46%)", sheet: "N/A", mode: "Formula_Simulasi_SPPT_2026_NDT46" },
];

// Helper function to get column letter from index (1-based)
function getColumnLetter(colIndex: number): string {
  let result = '';
  while (colIndex > 0) {
    colIndex--;
    result = String.fromCharCode(65 + (colIndex % 26)) + result;
    colIndex = Math.floor(colIndex / 26);
  }
  return result;
}

// Smart sheet finder - same as Python
function getSheetSmart(workbook: XLSX.WorkBook, nameHint: string): XLSX.WorkSheet | null {
  if (nameHint === "N/A") return null;
  const nameHintLower = nameHint.toLowerCase();
  const sheetMap: { [key: string]: string } = {};
  
  for (const name of workbook.SheetNames) {
    sheetMap[name.toLowerCase()] = name;
  }
  
  if (workbook.Sheets[nameHint]) return workbook.Sheets[nameHint];
  if (sheetMap[nameHintLower]) return workbook.Sheets[sheetMap[nameHintLower]];
  
  for (const existingSheet of Object.keys(sheetMap)) {
    if (nameHintLower.includes("c.1") && existingSheet.includes("c.1")) 
      return workbook.Sheets[sheetMap[existingSheet]];
    if (nameHintLower.includes("c.2") && existingSheet.includes("c.2")) 
      return workbook.Sheets[sheetMap[existingSheet]];
    if (nameHintLower.includes("home") && existingSheet.includes("home")) 
      return workbook.Sheets[sheetMap[existingSheet]];
    if (nameHintLower.includes("fdm") && existingSheet.includes("fdm")) 
      return workbook.Sheets[sheetMap[existingSheet]];
    if ((nameHintLower === "d" || nameHintLower === "sheet d") && 
        (existingSheet === "d" || existingSheet === "sheet d")) 
      return workbook.Sheets[sheetMap[existingSheet]];
  }
  return null;
}

// Get cell value from sheet
function getCellValue(sheet: XLSX.WorkSheet, addr: string): any {
  const cell = sheet[addr];
  return cell ? cell.v : null;
}

// Find FDM anchor row for dynamic extraction
function findFDMAnchorRow(sheet: XLSX.WorkSheet): number | null {
  const range = XLSX.utils.decode_range(sheet['!ref'] || 'A1');
  
  for (let row = 19; row <= Math.min(149, range.e.r); row++) {
    for (let col = 0; col <= 4; col++) {
      const cellAddr = XLSX.utils.encode_cell({ r: row, c: col });
      const cell = sheet[cellAddr];
      if (cell && cell.v && typeof cell.v === 'string') {
        if (cell.v.toUpperCase().includes("NJOP BANGUNAN PER METER PERSEGI")) {
          return row + 1; // Convert to 1-based
        }
      }
    }
  }
  return null;
}

// Extract data from a single file
export async function extractFDMData(file: File): Promise<ExtractionResult> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        
        const result: ExtractionResult = {
          rowNumber: 0,
          fileName: file.name,
          data: {}
        };

        // Pre-scan FDM sheet for anchor row
        const fdmSheet = getSheetSmart(workbook, "FDM Kebun ABC");
        let fdmAnchorRow: number | null = null;
        if (fdmSheet) {
          fdmAnchorRow = findFDMAnchorRow(fdmSheet);
        }

        // Extract data for each item
        for (const item of itemsDefinitions) {
          const mode = item.mode;
          let val: any = null;

          if (mode.includes("Formula")) {
            result.data[item.label] = null; // Will be filled later with formulas
            continue;
          }

          const ws = getSheetSmart(workbook, item.sheet);

          if (!ws) {
            val = "Sheet Not Found";
          } else {
            if (mode === "Static" && item.addr) {
              try {
                val = getCellValue(ws, item.addr);
              } catch {
                val = "Error";
              }
            } else if (mode === "Dynamic_Col_G" && item.keyword) {
              val = "TIDAK DITEMUKAN";
              const keyword = item.keyword.toLowerCase();
              const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
              
              for (let row = 0; row <= Math.min(149, range.e.r); row++) {
                let found = false;
                for (let col = 0; col <= 4; col++) {
                  const cellAddr = XLSX.utils.encode_cell({ r: row, c: col });
                  const cell = ws[cellAddr];
                  if (cell && cell.v && typeof cell.v === 'string') {
                    const cellText = cell.v.toLowerCase().replace(/\s+/g, ' ');
                    if (cellText.includes(keyword)) {
                      const colGAddr = XLSX.utils.encode_cell({ r: row, c: 6 });
                      val = ws[colGAddr]?.v;
                      found = true;
                      break;
                    }
                  }
                }
                if (found) break;
              }
            } else if (mode.includes("Dynamic_FDM_Bangunan")) {
              if (fdmAnchorRow && fdmSheet) {
                if (mode === "Dynamic_FDM_Bangunan_PerM2") {
                  const addr = XLSX.utils.encode_cell({ r: fdmAnchorRow - 1, c: 4 });
                  val = fdmSheet[addr]?.v;
                } else if (mode === "Dynamic_FDM_Bangunan_NJOP") {
                  const addr = XLSX.utils.encode_cell({ r: fdmAnchorRow - 2, c: 4 });
                  val = fdmSheet[addr]?.v;
                } else if (mode === "Dynamic_FDM_Bangunan_Luas") {
                  const addr = XLSX.utils.encode_cell({ r: fdmAnchorRow - 2, c: 3 });
                  val = fdmSheet[addr]?.v;
                }
              } else {
                val = "Anchor Not Found";
              }
            }
          }

          // Clean KELURAHAN value
          if (item.label === "KELURAHAN" && val && typeof val === 'string') {
            if (val.includes("#")) {
              val = val.replace(/#\s*\d+.*$/, '').trim();
            }
          }

          result.data[item.label] = val;
        }

        resolve(result);
      } catch (error) {
        reject(error);
      }
    };

    reader.onerror = (error) => reject(error);
    reader.readAsArrayBuffer(file);
  });
}

// Generate Excel with formulas
export function generateResultExcel(results: ExtractionResult[]): XLSX.WorkBook {
  const headers = ["NO", ...itemsDefinitions.map(item => item.label)];
  
  // Create data rows
  const dataRows: any[][] = [];
  
  for (let idx = 0; idx < results.length; idx++) {
    const result = results[idx];
    const row: any[] = [idx + 1];
    
    for (const item of itemsDefinitions) {
      if (item.mode.includes("Formula")) {
        row.push(null); // Placeholder for formula
      } else {
        row.push(result.data[item.label] ?? null);
      }
    }
    
    dataRows.push(row);
  }

  // Create worksheet
  const ws = XLSX.utils.aoa_to_sheet([headers, ...dataRows]);
  
  // Add formulas for each row
  const colMap: { [key: string]: string } = {};
  for (let i = 0; i < headers.length; i++) {
    colMap[headers[i]] = getColumnLetter(i + 1);
  }

  for (let idx = 0; idx < results.length; idx++) {
    const excelRow = idx + 2;
    
    // Formula: LUAS BUMI
    const arealCols = ["Areal Produktif", "Areal Belum Diolah", "Areal Sudah Diolah Belum Ditanami", "Areal Pembibitan", "Areal Tidak Produktif", "Areal Pengaman", "Areal Emplasemen"];
    const cellsToSum = arealCols.map(c => `${colMap[c]}${excelRow}`);
    ws[`${colMap["LUAS BUMI"]}${excelRow}`] = { f: `SUM(${cellsToSum.join(",")})`, t: 'n' };

    // Formula: Areal Produktif (Copy)
    ws[`${colMap["Areal Produktif (Copy)"]}${excelRow}`] = { f: `${colMap["Areal Produktif"]}${excelRow}`, t: 'n' };

    // Formula: NJOP Bumi Berupa Tanah (Rp)
    ws[`${colMap["NJOP Bumi Berupa Tanah (Rp)"]}${excelRow}`] = { 
      f: `${colMap["Areal Produktif"]}${excelRow}*${colMap["NJOP/M Areal Belum Produktif"]}${excelRow}`, 
      t: 'n' 
    };

    // Formula: NJOP Bumi Berupa Pengembangan Tanah (Rp) (Kenaikan BIT 10.3%)
    const colAsalBIT = colMap["NJOP Bumi Berupa Pengembangan Tanah (Rp)"];
    ws[`${colMap["NJOP Bumi Berupa Pengembangan Tanah (Rp) (Kenaikan BIT 10.3%)"]}${excelRow}`] = { 
      f: `${colAsalBIT}${excelRow}+(${colAsalBIT}${excelRow}*'2. Kesimpulan'!$E$2)`, 
      t: 'n' 
    };

    // Formula: NJOP Bumi Areal Produktif (Rp)
    const colTanah = colMap["NJOP Bumi Berupa Tanah (Rp)"];
    const colPengembangan = colMap["NJOP Bumi Berupa Pengembangan Tanah (Rp)"];
    ws[`${colMap["NJOP Bumi Areal Produktif (Rp)"]}${excelRow}`] = { 
      f: `${colTanah}${excelRow}+${colPengembangan}${excelRow}`, 
      t: 'n' 
    };

    // Formula: Luas Bumi Areal Produktif (m²)
    ws[`${colMap["Luas Bumi Areal Produktif (m²)"]}${excelRow}`] = { 
      f: `${colMap["Areal Produktif"]}${excelRow}`, 
      t: 'n' 
    };

    // Formula: NJOP Bumi Per M2 Areal Produktif (Rp/m2)
    const colNJOPTotalProd = colMap["NJOP Bumi Areal Produktif (Rp)"];
    const colLuasProd = colMap["Luas Bumi Areal Produktif (m²)"];
    ws[`${colMap["NJOP Bumi Per M2 Areal Produktif (Rp/m2)"]}${excelRow}`] = { 
      f: `${colNJOPTotalProd}${excelRow}/${colLuasProd}${excelRow}`, 
      t: 'n' 
    };

    // Formula: NJOP BUMI (Rp) AREA PRODUKTIF pada A. DATA BUMI
    const colHargaM2 = colMap["NJOP Bumi Per M2 Areal Produktif (Rp/m2)"];
    ws[`${colMap["NJOP BUMI (Rp) AREA PRODUKTIF pada A. DATA BUMI"]}${excelRow}`] = { 
      f: `${colLuasProd}${excelRow}*${colHargaM2}${excelRow}`, 
      t: 'n' 
    };

    // Formula: NJOP BUMI (Rp) AREA PRODUKTIF pada A. DATA BUMI (Proyeksi NDT Naik 46%)
    const colBITNaik = colMap["NJOP Bumi Berupa Pengembangan Tanah (Rp) (Kenaikan BIT 10.3%)"];
    const colArealProd = colMap["Areal Produktif"];
    ws[`${colMap["NJOP BUMI (Rp) AREA PRODUKTIF pada A. DATA BUMI (Proyeksi NDT Naik 46%)"]}${excelRow}`] = { 
      f: `ROUND(((${colTanah}${excelRow}+${colBITNaik}${excelRow})/${colArealProd}${excelRow}),0)*${colLuasProd}${excelRow}`, 
      t: 'n' 
    };

    // Formula: NJOP BUMI (Rp) AREAL BELUM PRODUKTIF pada A. DATA BUMI (Proyeksi NDT Naik 46%)
    const colBelumProd = colMap["NJOP BUMI (Rp) AREAL BELUM PRODUKTIF pada A. DATA BUMI"];
    ws[`${colMap["NJOP BUMI (Rp) AREAL BELUM PRODUKTIF pada A. DATA BUMI (Proyeksi NDT Naik 46%)"]}${excelRow}`] = { 
      f: `${colBelumProd}${excelRow}*(1+'2. Kesimpulan'!$E$14)`, 
      t: 'n' 
    };

    // Formula: Areal Tidak Produktif (Copy)
    const colTidakProdAsal = colMap["Areal Tidak Produktif"];
    ws[`${colMap["Areal Tidak Produktif (Copy)"]}${excelRow}`] = { 
      f: `${colTidakProdAsal}${excelRow}`, 
      t: 'n' 
    };

    // Formula: NJOP BUMI (Rp) AREAL TIDAK PRODUKTIF pada A. DATA BUMI
    const colTidakProdCopy = colMap["Areal Tidak Produktif (Copy)"];
    const colNJOPMTidakProd = colMap["NJOP/M Areal Tidak Produktif"];
    ws[`${colMap["NJOP BUMI (Rp) AREAL TIDAK PRODUKTIF pada A. DATA BUMI"]}${excelRow}`] = { 
      f: `${colTidakProdCopy}${excelRow}*${colNJOPMTidakProd}${excelRow}`, 
      t: 'n' 
    };

    // Formula: NJOP BUMI (Rp) AREAL TIDAK PRODUKTIF pada A. DATA BUMI (Proyeksi NDT Naik 46%)
    const colNJOPTidakProd = colMap["NJOP BUMI (Rp) AREAL TIDAK PRODUKTIF pada A. DATA BUMI"];
    ws[`${colMap["NJOP BUMI (Rp) AREAL TIDAK PRODUKTIF pada A. DATA BUMI (Proyeksi NDT Naik 46%)"]}${excelRow}`] = { 
      f: `${colNJOPTidakProd}${excelRow}*(1+'2. Kesimpulan'!$E$14)`, 
      t: 'n' 
    };

    // Formula: Areal Pengaman (Copy)
    const colPengamanAsal = colMap["Areal Pengaman"];
    ws[`${colMap["Areal Pengaman (Copy)"]}${excelRow}`] = { 
      f: `${colPengamanAsal}${excelRow}`, 
      t: 'n' 
    };

    // Formula: NJOP BUMI (Rp) AREAL PENGAMAN pada A. DATA BUMI
    const colPengamanCopy = colMap["Areal Pengaman (Copy)"];
    const colNJOPMPengaman = colMap["NJOP/M Areal Pengaman"];
    ws[`${colMap["NJOP BUMI (Rp) AREAL PENGAMAN pada A. DATA BUMI"]}${excelRow}`] = { 
      f: `${colPengamanCopy}${excelRow}*${colNJOPMPengaman}${excelRow}`, 
      t: 'n' 
    };

    // Formula: NJOP BUMI (Rp) AREAL PENGAMAN pada A. DATA BUMI (Proyeksi NDT Naik 46%)
    const colNJOPPengaman = colMap["NJOP BUMI (Rp) AREAL PENGAMAN pada A. DATA BUMI"];
    ws[`${colMap["NJOP BUMI (Rp) AREAL PENGAMAN pada A. DATA BUMI (Proyeksi NDT Naik 46%)"]}${excelRow}`] = { 
      f: `${colNJOPPengaman}${excelRow}*(1+'2. Kesimpulan'!$E$14)`, 
      t: 'n' 
    };

    // Formula: Areal Emplasemen (Copy)
    const colEmplasemenAsal = colMap["Areal Emplasemen"];
    ws[`${colMap["Areal Emplasemen (Copy)"]}${excelRow}`] = { 
      f: `${colEmplasemenAsal}${excelRow}`, 
      t: 'n' 
    };

    // Formula: NJOP BUMI (Rp) AREAL EMPLASEMEN pada A. DATA BUMI
    const colEmplasemenCopy = colMap["Areal Emplasemen (Copy)"];
    const colNJOPMEmplasemen = colMap["NJOP/M Areal Emplasemen"];
    ws[`${colMap["NJOP BUMI (Rp) AREAL EMPLASEMEN pada A. DATA BUMI"]}${excelRow}`] = { 
      f: `${colEmplasemenCopy}${excelRow}*${colNJOPMEmplasemen}${excelRow}`, 
      t: 'n' 
    };

    // Formula: NJOP BUMI (Rp) AREAL EMPLASEMEN pada A. DATA BUMI (Proyeksi NDT Naik 46%)
    const colNJOPEmplasemen = colMap["NJOP BUMI (Rp) AREAL EMPLASEMEN pada A. DATA BUMI"];
    ws[`${colMap["NJOP BUMI (Rp) AREAL EMPLASEMEN pada A. DATA BUMI (Proyeksi NDT Naik 46%)"]}${excelRow}`] = { 
      f: `${colNJOPEmplasemen}${excelRow}*(1+'2. Kesimpulan'!$E$14)`, 
      t: 'n' 
    };

    // Formula: JUMLAH Luas (m2) pada A. DATA BUMI
    const colLuasBumiGlobal = colMap["LUAS BUMI"];
    ws[`${colMap["JUMLAH Luas (m2) pada A. DATA BUMI"]}${excelRow}`] = { 
      f: `${colLuasBumiGlobal}${excelRow}`, 
      t: 'n' 
    };

    // Formula: JUMLAH NJOP BUMI (Rp) pada A. DATA BUMI
    const njopComponents = ["NJOP BUMI (Rp) AREA PRODUKTIF pada A. DATA BUMI", "NJOP BUMI (Rp) AREAL BELUM PRODUKTIF pada A. DATA BUMI", "NJOP BUMI (Rp) AREAL TIDAK PRODUKTIF pada A. DATA BUMI", "NJOP BUMI (Rp) AREAL PENGAMAN pada A. DATA BUMI", "NJOP BUMI (Rp) AREAL EMPLASEMEN pada A. DATA BUMI"];
    const colsToSum = njopComponents.map(c => `${colMap[c]}${excelRow}`);
    ws[`${colMap["JUMLAH NJOP BUMI (Rp) pada A. DATA BUMI"]}${excelRow}`] = { 
      f: colsToSum.join("+"), 
      t: 'n' 
    };

    // Formula: Jumlah NJOP BANGUNAN pada B. DATA BANGUNAN
    const colLuasBgn = colMap["Jumlah LUAS pada B. DATA BANGUNAN"];
    const colNJOPM2Bgn = colMap["NJOP BANGUNAN PER METER PERSEGI*) pada B. DATA BANGUNAN"];
    ws[`${colMap["Jumlah NJOP BANGUNAN pada B. DATA BANGUNAN"]}${excelRow}`] = { 
      f: `${colLuasBgn}${excelRow}*${colNJOPM2Bgn}${excelRow}`, 
      t: 'n' 
    };

    // Formula: TOTAL NJOP (TANAH + BANGUNAN) 2025
    const colTotalBumi = colMap["JUMLAH NJOP BUMI (Rp) pada A. DATA BUMI"];
    const colTotalBgn = colMap["Jumlah NJOP BANGUNAN pada B. DATA BANGUNAN"];
    ws[`${colMap["TOTAL NJOP (TANAH + BANGUNAN) 2025"]}${excelRow}`] = { 
      f: `${colTotalBumi}${excelRow}+${colTotalBgn}${excelRow}`, 
      t: 'n' 
    };

    // Formula: SPPT 2025
    const colTotalNJOP25 = colMap["TOTAL NJOP (TANAH + BANGUNAN) 2025"];
    ws[`${colMap["SPPT 2025"]}${excelRow}`] = { 
      f: `((${colTotalNJOP25}${excelRow}-12000000)*40%)*0.5%`, 
      t: 'n' 
    };

    // Formula: SIMULASI TOTAL NJOP (TANAH + BANGUNAN) 2026 (Hanya Kenaikan BIT 10,3% + NDT Tetap)
    const T = `${colMap["NJOP Bumi Berupa Tanah (Rp)"]}${excelRow}`;
    const V = `${colMap["NJOP Bumi Berupa Pengembangan Tanah (Rp) (Kenaikan BIT 10.3%)"]}${excelRow}`;
    const R = `${colMap["Areal Produktif"]}${excelRow}`;
    const X = `${colMap["Luas Bumi Areal Produktif (m²)"]}${excelRow}`;
    const AB = `${colMap["NJOP BUMI (Rp) AREAL BELUM PRODUKTIF pada A. DATA BUMI"]}${excelRow}`;
    const AF = `${colMap["NJOP BUMI (Rp) AREAL TIDAK PRODUKTIF pada A. DATA BUMI"]}${excelRow}`;
    const AJ = `${colMap["NJOP BUMI (Rp) AREAL PENGAMAN pada A. DATA BUMI"]}${excelRow}`;
    const AN = `${colMap["NJOP BUMI (Rp) AREAL EMPLASEMEN pada A. DATA BUMI"]}${excelRow}`;
    const AT = `${colMap["Jumlah NJOP BANGUNAN pada B. DATA BANGUNAN"]}${excelRow}`;
    ws[`${colMap["SIMULASI TOTAL NJOP (TANAH + BANGUNAN) 2026 (Hanya Kenaikan BIT 10,3% + NDT Tetap)"]}${excelRow}`] = { 
      f: `(ROUND(((${T}+${V})/${R}),0)*${X})+${AB}+${AF}+${AJ}+${AN}+${AT}`, 
      t: 'n' 
    };

    // Formula: SIMULASI SPPT 2026 (Hanya Kenaikan BIT 10,3% + NDT Tetap)
    const colSimNJOP26 = colMap["SIMULASI TOTAL NJOP (TANAH + BANGUNAN) 2026 (Hanya Kenaikan BIT 10,3% + NDT Tetap)"];
    ws[`${colMap["SIMULASI SPPT 2026 (Hanya Kenaikan BIT 10,3% + NDT Tetap)"]}${excelRow}`] = { 
      f: `((${colSimNJOP26}${excelRow}-12000000)*40%)*0.5%`, 
      t: 'n' 
    };

    // Formula: Kenaikan
    const colSimSPPT26 = colMap["SIMULASI SPPT 2026 (Hanya Kenaikan BIT 10,3% + NDT Tetap)"];
    const colSPPT25 = colMap["SPPT 2025"];
    ws[`${colMap["Kenaikan"]}${excelRow}`] = { 
      f: `${colSimSPPT26}${excelRow}-${colSPPT25}${excelRow}`, 
      t: 'n' 
    };

    // Formula: Persentase
    const colKenaikan = colMap["Kenaikan"];
    ws[`${colMap["Persentase"]}${excelRow}`] = { 
      f: `${colKenaikan}${excelRow}/${colSPPT25}${excelRow}`, 
      t: 'n' 
    };

    // Formula: SIMULASI TOTAL NJOP (TANAH + BANGUNAN) 2026 (Kenaikan BIT 10,3% + NDT 46%)
    const AA = `${colMap["NJOP BUMI (Rp) AREA PRODUKTIF pada A. DATA BUMI (Proyeksi NDT Naik 46%)"]}${excelRow}`;
    const AC = `${colMap["NJOP BUMI (Rp) AREAL BELUM PRODUKTIF pada A. DATA BUMI (Proyeksi NDT Naik 46%)"]}${excelRow}`;
    const AG = `${colMap["NJOP BUMI (Rp) AREAL TIDAK PRODUKTIF pada A. DATA BUMI (Proyeksi NDT Naik 46%)"]}${excelRow}`;
    const AK = `${colMap["NJOP BUMI (Rp) AREAL PENGAMAN pada A. DATA BUMI (Proyeksi NDT Naik 46%)"]}${excelRow}`;
    const AO = `${colMap["NJOP BUMI (Rp) AREAL EMPLASEMEN pada A. DATA BUMI (Proyeksi NDT Naik 46%)"]}${excelRow}`;
    ws[`${colMap["SIMULASI TOTAL NJOP (TANAH + BANGUNAN) 2026 (Kenaikan BIT 10,3% + NDT 46%)"]}${excelRow}`] = { 
      f: `(${AA}+${AC}+${AG}+${AK}+${AO})+${AT}`, 
      t: 'n' 
    };

    // Formula: SIMULASI SPPT 2026 (Kenaikan BIT 10,3% + NDT 46%)
    const colSimNJOP26NDT = colMap["SIMULASI TOTAL NJOP (TANAH + BANGUNAN) 2026 (Kenaikan BIT 10,3% + NDT 46%)"];
    ws[`${colMap["SIMULASI SPPT 2026 (Kenaikan BIT 10,3% + NDT 46%)"]}${excelRow}`] = { 
      f: `((${colSimNJOP26NDT}${excelRow}-12000000)*40%)*0.5%`, 
      t: 'n' 
    };
  }

  // Set column widths
  ws['!cols'] = headers.map(() => ({ wch: 25 }));

  // Create workbook
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "1. Hasil");

  // Create Sheet 2: Kesimpulan
  const ws2 = XLSX.utils.aoa_to_sheet([]);
  
  // Set column widths for sheet 2
  ws2['!cols'] = [
    { wch: 60 }, { wch: 30 }, { wch: 25 }, { wch: 20 }, { wch: 20 }
  ];

  // Add data to sheet 2
  const kesimpulanData = [
    { cell: "E1", value: "Skenario Kenaikan BIT" },
    { cell: "E2", value: 0.103 },
    { cell: "A1", value: "Poin" },
    { cell: "B1", value: "Keterangan (BIT + 10.3% dan NDT Tetap)" },
    { cell: "C1", value: "Nilai" },
    { cell: "D1", value: "Keterangan" },
    
    { cell: "A2", value: "Simulasi Penerimaan PBB 2026" },
    { cell: "B2", value: "Perkebunan" },
    { cell: "C2", value: { f: "SUMIF('1. Hasil'!C2:C10000,\"Perkebunan\",'1. Hasil'!AY2:AY10000)" } },
    { cell: "A3", value: "Simulasi Penerimaan PBB 2026" },
    { cell: "B3", value: "Minerba" },
    { cell: "C3", value: { f: "SUMIF('1. Hasil'!C2:C10000,\"Minerba\",'1. Hasil'!AY2:AY10000)" } },
    { cell: "A4", value: "Simulasi Penerimaan PBB 2026" },
    { cell: "B4", value: "Perhutanan (HTI)" },
    { cell: "C4", value: { f: "SUMIF('1. Hasil'!C2:C10000,\"Perhutanan (HTI)\",'1. Hasil'!AY2:AY10000)" } },
    { cell: "A5", value: "Simulasi Penerimaan PBB 2026" },
    { cell: "B5", value: "Perhutanan (Hutan Alam)" },
    { cell: "C5", value: { f: "SUMIF('1. Hasil'!C2:C10000,\"Perhutanan (Hutan Alam)\",'1. Hasil'!AY2:AY10000)" } },
    { cell: "A6", value: "Simulasi Penerimaan PBB 2026" },
    { cell: "B6", value: "Sektor Lainnya" },
    { cell: "C6", value: { f: "SUMIF('1. Hasil'!C2:C10000,\"Sektor Lainnya\",'1. Hasil'!AY2:AY10000)" } },
    
    { cell: "A7", value: "Simulasi Penerimaan PBB 2026 (Collection Rate 100%)" },
    { cell: "B7", value: { f: "(COUNT('1. Hasil'!A2:A10000))&\" NOP\"" } },
    { cell: "C7", value: { f: "SUM(C2:C6)" } },
    { cell: "A8", value: "Target Penerimaan PBB 2026" },
    { cell: "C8", value: 110289165592 },
    { cell: "A9", value: "Selisih antara Simulasi (Collection Rate 100%) & Target" },
    { cell: "C9", value: { f: "C7-C8" } },
    { cell: "D9", value: { f: 'IF(C9>0,"Tercapai","Tidak Tercapai")' } },
    
    { cell: "A10", value: "Simulasi Penerimaan PBB 2026 (Collection Rate 95%)" },
    { cell: "B10", value: 0.95 },
    { cell: "C10", value: { f: "C7*B10" } },
    { cell: "A11", value: "Selisih antara Simulasi (Collection Rate 95%) Target" },
    { cell: "C11", value: { f: "C10-C8" } },
    { cell: "D11", value: { f: 'IF(C11>0,"Tercapai","Tidak Tercapai")' } },
    
    { cell: "A13", value: "Poin" },
    { cell: "B13", value: "Keterangan (BIT + 10.3% dan NDT + 46%)" },
    { cell: "C13", value: "Nilai" },
    { cell: "D13", value: "Keterangan" },
    { cell: "E13", value: "Skenario Kenaikan NDT" },
    
    { cell: "A14", value: { f: "=A2" } },
    { cell: "B14", value: { f: "=B2" } },
    { cell: "C14", value: { f: "SUMIF('1. Hasil'!C2:C10000,\"Perkebunan\",'1. Hasil'!BC2:BC10000)" } },
    { cell: "E14", value: 0.46 },
    { cell: "A15", value: { f: "=A3" } },
    { cell: "B15", value: { f: "=B3" } },
    { cell: "C15", value: { f: "SUMIF('1. Hasil'!C2:C10000,\"Minerba\",'1. Hasil'!BC2:BC10000)" } },
    { cell: "A16", value: { f: "=A4" } },
    { cell: "B16", value: { f: "=B4" } },
    { cell: "C16", value: { f: "SUMIF('1. Hasil'!C2:C10000,\"Perhutanan (HTI)\",'1. Hasil'!BC2:BC10000)" } },
    { cell: "A17", value: { f: "=A5" } },
    { cell: "B17", value: { f: "=B5" } },
    { cell: "C17", value: { f: "SUMIF('1. Hasil'!C2:C10000,\"Perhutanan (Hutan Alam)\",'1. Hasil'!BC2:BC10000)" } },
    { cell: "A18", value: { f: "=A6" } },
    { cell: "B18", value: { f: "=B6" } },
    { cell: "C18", value: { f: "SUMIF('1. Hasil'!C2:C10000,\"Sektor Lainnya\",'1. Hasil'!BC2:BC10000)" } },
    
    { cell: "A19", value: { f: "=A7" } },
    { cell: "B19", value: { f: "=B7" } },
    { cell: "C19", value: { f: "SUM(C14:C18)" } },
    { cell: "A20", value: { f: "=A8" } },
    { cell: "C20", value: { f: "=C8" } },
    { cell: "A21", value: { f: "=A9" } },
    { cell: "C21", value: { f: "C19-C20" } },
    { cell: "D21", value: { f: 'IF(C21>0,"Tercapai","Tidak Tercapai")' } },
    
    { cell: "A22", value: "Simulasi Penerimaan PBB 2026 (Collection Rate 95%)" },
    { cell: "B22", value: 0.95 },
    { cell: "C22", value: { f: "C19*B22" } },
    { cell: "A23", value: "Selisih antara Simulasi (Collection Rate 95%) Target" },
    { cell: "C23", value: { f: "C22-C20" } },
    { cell: "D23", value: { f: 'IF(C23>0,"Tercapai","Tidak Tercapai")' } }
  ];

  for (const item of kesimpulanData) {
    const addr = item.cell as string;
    if (typeof item.value === 'object' && item.value.f) {
      ws2[addr] = { f: item.value.f, t: 'n' };
    } else if (typeof item.value === 'number') {
      ws2[addr] = { v: item.value, t: 'n' };
    } else {
      ws2[addr] = { v: item.value, t: 's' };
    }
  }

  // Set range for sheet 2
  ws2['!ref'] = 'A1:E23';

  XLSX.utils.book_append_sheet(wb, ws2, "2. Kesimpulan");

  return wb;
}
