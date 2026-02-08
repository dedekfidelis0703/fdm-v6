import { useState, useRef, useCallback } from 'react';
import { Button } from '@/components/ui/button';
import { Card, CardContent, CardHeader, CardTitle, CardDescription } from '@/components/ui/card';
import { Alert, AlertDescription } from '@/components/ui/alert';
import { Progress } from '@/components/ui/progress';
import { 
  Upload, 
  Play, 
  Download, 
  RefreshCw, 
  FileSpreadsheet, 
  CheckCircle, 
  AlertCircle,
  X,
  FileUp
} from 'lucide-react';
import { extractFDMData, generateResultExcel } from '@/utils/fdmExtractor';
import type { ExtractionResult } from '@/types/fdm';
import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';

const MAX_FILES = 50;

export function FDMExtractor() {
  const [uploadedFiles, setUploadedFiles] = useState<File[]>([]);
  const [extractionResults, setExtractionResults] = useState<ExtractionResult[] | null>(null);
  const [isExtracting, setIsExtracting] = useState(false);
  const [extractionProgress, setExtractionProgress] = useState(0);
  const [error, setError] = useState<string | null>(null);
  const [success, setSuccess] = useState<string | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  // Handle file upload
  const handleFileUpload = useCallback((event: React.ChangeEvent<HTMLInputElement>) => {
    const files = event.target.files;
    if (!files) return;

    const excelFiles = Array.from(files).filter(file => {
      const ext = file.name.toLowerCase();
      return ext.endsWith('.xlsx') || ext.endsWith('.xls') || ext.endsWith('.xlsm');
    });

    if (excelFiles.length === 0) {
      setError('Harap upload file Excel (.xlsx, .xls, .xlsm)');
      return;
    }

    if (excelFiles.length > MAX_FILES) {
      setError(`Maksimal ${MAX_FILES} file dapat diupload sekaligus. Anda memilih ${excelFiles.length} file.`);
      return;
    }

    setUploadedFiles(prev => [...prev, ...excelFiles]);
    setError(null);
    setSuccess(`${excelFiles.length} file berhasil ditambahkan. Total: ${uploadedFiles.length + excelFiles.length} file.`);
    
    // Clear success message after 3 seconds
    setTimeout(() => setSuccess(null), 3000);
  }, [uploadedFiles.length]);

  // Remove file from list
  const removeFile = useCallback((index: number) => {
    setUploadedFiles(prev => prev.filter((_, i) => i !== index));
  }, []);

  // Clear all files
  const clearAllFiles = useCallback(() => {
    setUploadedFiles([]);
    setExtractionResults(null);
    setError(null);
    setSuccess(null);
    if (fileInputRef.current) {
      fileInputRef.current.value = '';
    }
  }, []);

  // Extract data from files
  const handleExtract = useCallback(async () => {
    if (uploadedFiles.length === 0) {
      setError('Harap upload file FDM terlebih dahulu');
      return;
    }

    setIsExtracting(true);
    setExtractionProgress(0);
    setError(null);
    setSuccess(null);

    try {
      const results: ExtractionResult[] = [];
      
      for (let i = 0; i < uploadedFiles.length; i++) {
        const file = uploadedFiles[i];
        const result = await extractFDMData(file);
        result.rowNumber = i + 1;
        results.push(result);
        
        // Update progress
        setExtractionProgress(Math.round(((i + 1) / uploadedFiles.length) * 100));
      }

      setExtractionResults(results);
      
      // Auto download after extraction
      const workbook = generateResultExcel(results);
      const wbout = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
      const blob = new Blob([wbout], { type: 'application/octet-stream' });
      saveAs(blob, 'Hasil_Ekstraksi_FDM.xlsx');
      
      setSuccess('Ekstraksi berhasil! File hasil telah didownload secara otomatis.');
    } catch (err) {
      setError(`Terjadi kesalahan saat ekstraksi: ${err instanceof Error ? err.message : 'Unknown error'}`);
    } finally {
      setIsExtracting(false);
    }
  }, [uploadedFiles]);

  // Download result again
  const handleDownloadAgain = useCallback(() => {
    if (!extractionResults || extractionResults.length === 0) {
      setError('Tidak ada hasil ekstraksi yang tersedia untuk didownload');
      return;
    }

    try {
      const workbook = generateResultExcel(extractionResults);
      const wbout = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
      const blob = new Blob([wbout], { type: 'application/octet-stream' });
      saveAs(blob, 'Hasil_Ekstraksi_FDM.xlsx');
      setSuccess('File berhasil didownload ulang!');
      setTimeout(() => setSuccess(null), 3000);
    } catch (err) {
      setError(`Gagal mendownload file: ${err instanceof Error ? err.message : 'Unknown error'}`);
    }
  }, [extractionResults]);

  // Start new extraction
  const handleNewExtraction = useCallback(() => {
    clearAllFiles();
    setSuccess('Silakan upload file FDM baru untuk ekstraksi');
    setTimeout(() => setSuccess(null), 3000);
  }, [clearAllFiles]);

  // Trigger file input click
  const triggerFileInput = useCallback(() => {
    fileInputRef.current?.click();
  }, []);

  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-50 to-slate-100 py-8 px-4">
      <div className="max-w-4xl mx-auto">
        {/* Header */}
        <div className="text-center mb-8">
          <h1 className="text-4xl font-bold text-slate-800 mb-2">
            Ekstraktor FDM
          </h1>
          <p className="text-slate-600">
            Alat ekstraksi data dari file FDM (Formulir Data Masuk) ke Excel
          </p>
        </div>

        {/* Alerts */}
        {error && (
          <Alert variant="destructive" className="mb-6">
            <AlertCircle className="h-4 w-4" />
            <AlertDescription>{error}</AlertDescription>
          </Alert>
        )}

        {success && (
          <Alert className="mb-6 bg-green-50 border-green-200 text-green-800">
            <CheckCircle className="h-4 w-4 text-green-600" />
            <AlertDescription>{success}</AlertDescription>
          </Alert>
        )}

        {/* Main Card */}
        <Card className="shadow-xl border-0">
          <CardHeader className="bg-gradient-to-r from-blue-600 to-blue-700 text-white rounded-t-lg">
            <CardTitle className="text-2xl flex items-center gap-2">
              <FileSpreadsheet className="h-6 w-6" />
              Ekstraksi Data FDM
            </CardTitle>
            <CardDescription className="text-blue-100">
              Upload file FDM Anda dan ekstrak data dengan mudah
            </CardDescription>
          </CardHeader>

          <CardContent className="p-6">
            {/* Hidden file input */}
            <input
              ref={fileInputRef}
              type="file"
              accept=".xlsx,.xls,.xlsm"
              multiple
              onChange={handleFileUpload}
              className="hidden"
              max={MAX_FILES}
            />

            {/* Action Buttons */}
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4 mb-6">
              {/* Button 1: Upload FDM */}
              <Button
                onClick={triggerFileInput}
                variant="outline"
                className="h-16 text-lg font-semibold border-2 border-dashed border-blue-300 hover:border-blue-500 hover:bg-blue-50 transition-all"
                disabled={isExtracting}
              >
                <Upload className="h-5 w-5 mr-2" />
                Upload FDM
              </Button>

              {/* Button 2: Ekstrak Sekarang */}
              <Button
                onClick={handleExtract}
                className="h-16 text-lg font-semibold bg-blue-600 hover:bg-blue-700 text-white"
                disabled={isExtracting || uploadedFiles.length === 0}
              >
                <Play className="h-5 w-5 mr-2" />
                {isExtracting ? 'Mengekstrak...' : 'Ekstrak Sekarang'}
              </Button>

              {/* Button 3: Download Ulang Hasil Ekstraksi */}
              <Button
                onClick={handleDownloadAgain}
                variant="outline"
                className="h-16 text-lg font-semibold border-green-300 hover:border-green-500 hover:bg-green-50 transition-all"
                disabled={!extractionResults || isExtracting}
              >
                <Download className="h-5 w-5 mr-2" />
                Download Ulang Hasil Ekstraksi
              </Button>

              {/* Button 4: Ekstraksi FDM Lain */}
              <Button
                onClick={handleNewExtraction}
                variant="outline"
                className="h-16 text-lg font-semibold border-orange-300 hover:border-orange-500 hover:bg-orange-50 transition-all"
                disabled={isExtracting}
              >
                <RefreshCw className="h-5 w-5 mr-2" />
                Ekstraksi FDM Lain
              </Button>
            </div>

            {/* Progress Bar */}
            {isExtracting && (
              <div className="mb-6">
                <div className="flex justify-between text-sm text-slate-600 mb-2">
                  <span>Progress Ekstraksi</span>
                  <span>{extractionProgress}%</span>
                </div>
                <Progress value={extractionProgress} className="h-3" />
              </div>
            )}

            {/* File List */}
            {uploadedFiles.length > 0 && (
              <div className="bg-slate-50 rounded-lg p-4">
                <div className="flex items-center justify-between mb-3">
                  <h3 className="font-semibold text-slate-700 flex items-center gap-2">
                    <FileUp className="h-4 w-4" />
                    File yang Diupload ({uploadedFiles.length}/{MAX_FILES})
                  </h3>
                  <Button
                    variant="ghost"
                    size="sm"
                    onClick={clearAllFiles}
                    className="text-red-500 hover:text-red-700 hover:bg-red-50"
                    disabled={isExtracting}
                  >
                    <X className="h-4 w-4 mr-1" />
                    Hapus Semua
                  </Button>
                </div>

                <div className="max-h-64 overflow-y-auto space-y-2">
                  {uploadedFiles.map((file, index) => (
                    <div
                      key={`${file.name}-${index}`}
                      className="flex items-center justify-between bg-white p-3 rounded-md shadow-sm border"
                    >
                      <div className="flex items-center gap-3 overflow-hidden">
                        <FileSpreadsheet className="h-5 w-5 text-green-600 flex-shrink-0" />
                        <span className="text-sm text-slate-700 truncate" title={file.name}>
                          {file.name}
                        </span>
                        <span className="text-xs text-slate-400 flex-shrink-0">
                          ({(file.size / 1024).toFixed(1)} KB)
                        </span>
                      </div>
                      <Button
                        variant="ghost"
                        size="sm"
                        onClick={() => removeFile(index)}
                        className="text-red-400 hover:text-red-600 hover:bg-red-50 flex-shrink-0"
                        disabled={isExtracting}
                      >
                        <X className="h-4 w-4" />
                      </Button>
                    </div>
                  ))}
                </div>
              </div>
            )}

            {/* Empty State */}
            {uploadedFiles.length === 0 && !isExtracting && (
              <div className="text-center py-12 bg-slate-50 rounded-lg border-2 border-dashed border-slate-200">
                <Upload className="h-12 w-12 text-slate-300 mx-auto mb-4" />
                <p className="text-slate-500 mb-2">Belum ada file yang diupload</p>
                <p className="text-sm text-slate-400">
                  Klik "Upload FDM" untuk memilih file Excel (maksimal 50 file)
                </p>
              </div>
            )}

            {/* Results Summary */}
            {extractionResults && !isExtracting && (
              <div className="mt-6 bg-green-50 border border-green-200 rounded-lg p-4">
                <h3 className="font-semibold text-green-800 flex items-center gap-2 mb-2">
                  <CheckCircle className="h-5 w-5" />
                  Ekstraksi Berhasil!
                </h3>
                <p className="text-green-700 text-sm">
                  Berhasil mengekstrak data dari {extractionResults.length} file FDM.
                  File hasil telah didownload secara otomatis.
                </p>
                <p className="text-green-600 text-xs mt-2">
                  Jika file belum terdownload, klik "Download Ulang Hasil Ekstraksi"
                </p>
              </div>
            )}
          </CardContent>
        </Card>

        {/* Footer */}
        <div className="text-center mt-8 text-slate-500 text-sm">
          <p>Privasi Terjaga: File Anda tidak disimpan di server</p>
          <p className="mt-1">
            Didukung oleh{' '}
            <a 
              href="https://vercel.com" 
              target="_blank" 
              rel="noopener noreferrer"
              className="text-blue-600 hover:underline"
            >
              Vercel
            </a>
          </p>
        </div>
      </div>
    </div>
  );
}
