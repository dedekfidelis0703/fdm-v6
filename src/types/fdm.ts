// Type definitions for FDM Extractor

export interface FDMItemDefinition {
  label: string;
  sheet: string;
  addr?: string;
  keyword?: string;
  mode: 
    | "Static" 
    | "Dynamic_Col_G" 
    | "Dynamic_FDM_Bangunan_Luas"
    | "Dynamic_FDM_Bangunan_PerM2"
    | "Dynamic_FDM_Bangunan_NJOP"
    | "Formula_LuasBumi"
    | "Formula_CopyProduktif"
    | "Formula_NJOPTanah"
    | "Formula_BIT"
    | "Formula_NJOP_Total"
    | "Formula_Luas_Ref"
    | "Formula_NJOP_PerM2"
    | "Formula_Final_Calc"
    | "Formula_Proyeksi"
    | "Formula_Proyeksi_BelumProd"
    | "Formula_CopyTidakProduktif"
    | "Formula_Calc_TidakProd"
    | "Formula_Proyeksi_TidakProd"
    | "Formula_CopyPengaman"
    | "Formula_Calc_Pengaman"
    | "Formula_Proyeksi_Pengaman"
    | "Formula_CopyEmplasemen"
    | "Formula_Calc_Emplasemen"
    | "Formula_Proyeksi_Emplasemen"
    | "Formula_Total_Luas_Ref"
    | "Formula_Total_NJOP_Sum"
    | "Formula_Calc_Bangunan"
    | "Formula_Grand_Total"
    | "Formula_SPPT_2025"
    | "Formula_Simulasi_NJOP_2026"
    | "Formula_Simulasi_SPPT_2026"
    | "Formula_Kenaikan"
    | "Formula_Persentase"
    | "Formula_Simulasi_Total_2026_NDT46"
    | "Formula_Simulasi_SPPT_2026_NDT46";
}

export interface ExtractedData {
  [key: string]: string | number | null;
}

export interface FDMSheetData {
  [sheetName: string]: {
    [cellAddress: string]: any;
  };
}

export interface FDMWorkbook {
  sheets: { [name: string]: any[][] };
  sheetNames: string[];
}

export interface ExtractionResult {
  rowNumber: number;
  fileName: string;
  data: ExtractedData;
}
