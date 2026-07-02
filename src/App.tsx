import { useState, useRef, useCallback } from 'react';
import { Upload, FileSpreadsheet, AlertCircle, XCircle } from 'lucide-react';
import { beautifyExcel } from './utils/excelBeautifier';
import { htmlToExcel } from './utils/htmlParser';
import { xmlToExcel } from './utils/xmlParser';
import { subjectExcelToExcel } from './utils/subjectExcelParser';
import ExcelJS from 'exceljs';

const MAX_FILE_SIZE = 100 * 1024 * 1024; // 100MB

function classifyFiles(files: File[]): { type: 'xml' | 'html' | 'excel'; files: File[] } | null {
  const isXml = (n: string) => n.endsWith('.xml') || n.endsWith('.roz');
  const isHtml = (n: string) => n.endsWith('.html') || n.endsWith('.htm');
  const isExcel = (n: string) => n.endsWith('.xlsx') || n.endsWith('.xls');
  if (files.some(f => isXml(f.name))) return { type: 'xml', files: files.filter(f => isXml(f.name)) };
  if (files.some(f => isHtml(f.name))) return { type: 'html', files: files.filter(f => isHtml(f.name)) };
  if (files.some(f => isExcel(f.name))) return { type: 'excel', files: files.filter(f => isExcel(f.name)) };
  return null;
}

function useFileProcessor() {
  const [loading, setLoading] = useState(false);
  const [progress, setProgress] = useState(0);
  const [progressMsg, setProgressMsg] = useState('');
  const [error, setError] = useState<string | null>(null);
  const [success, setSuccess] = useState<string | null>(null);
  const abortRef = useRef<AbortController | null>(null);

  const onProgress = useCallback((pct: number, msg: string) => {
    setProgress(pct);
    setProgressMsg(msg);
  }, []);

  const reset = useCallback(() => {
    setError(null);
    setSuccess(null);
    setProgress(0);
    setProgressMsg('');
  }, []);

  const abort = useCallback(() => {
    abortRef.current?.abort();
    abortRef.current = null;
    setLoading(false);
    setError('Proses dibatalkan.');
  }, []);

  const process = useCallback(async (files: File[]) => {
    reset();
    if (files.some(f => f.size > MAX_FILE_SIZE)) {
      setError(`File terlalu besar. Maksimal ${MAX_FILE_SIZE / 1024 / 1024}MB.`);
      return;
    }
    if (files.length === 0) return;

    const classified = classifyFiles(files);
    if (!classified) {
      setError('Harap unggah file ROZ, XML, HTML, atau Excel dari ASC Timetables.');
      return;
    }
    if (files.some(f => f.name.endsWith('.txt') || f.name.endsWith('.csv'))) {
      setError('❌ File txt/csv tidak mengandung grid jadwal (hari dan jam).');
      return;
    }

    const ctrl = new AbortController();
    abortRef.current = ctrl;
    setLoading(true);
    setProgressMsg('Memulai...');

    try {
      let blob: Blob;
      const signal = ctrl.signal;

      if (classified.type === 'xml') {
        blob = await xmlToExcel(classified.files[0], signal, onProgress);
      } else if (classified.type === 'html') {
        blob = await htmlToExcel(classified.files, signal, onProgress);
      } else {
        // excel — detect sub-type
        const buf = await classified.files[0].arrayBuffer();
        setProgressMsg('Menganalisa jenis Excel...');
        const wb = new ExcelJS.Workbook();
        await wb.xlsx.load(buf);
        let hasSubjectSheets = false;
        wb.eachSheet(s => { if (s.name.endsWith('.')) hasSubjectSheets = true; });
        blob = hasSubjectSheets
          ? await subjectExcelToExcel(buf, signal, onProgress)
          : await beautifyExcel(buf, signal, onProgress);
      }

      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `Jadwal_Rapi_${new Date().toISOString().split('T')[0]}.xlsx`;
      a.click();
      URL.revokeObjectURL(url);
      setSuccess('Jadwal berhasil diproses!');
    } catch (err: unknown) {
      if (err instanceof DOMException && err.name === 'AbortError') return;
      setError(err instanceof Error ? err.message : 'Terjadi kesalahan.');
    } finally {
      setLoading(false);
      abortRef.current = null;
    }
  }, [reset, onProgress]);

  return { loading, progress, progressMsg, error, success, process, abort };
}

function App() {
  const { loading, progress, progressMsg, error, success, process, abort } = useFileProcessor();
  const inputRef = useRef<HTMLInputElement>(null);
  const [dragOver, setDragOver] = useState(false);

  const handleFiles = useCallback((files: FileList | File[]) => process(Array.from(files)), [process]);

  const handleDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setDragOver(false);
    if (e.dataTransfer.files.length) handleFiles(e.dataTransfer.files);
  }, [handleFiles]);

  const handleDragOver = useCallback((e: React.DragEvent) => { e.preventDefault(); setDragOver(true); }, []);
  const handleDragLeave = useCallback(() => setDragOver(false), []);

  const handleKeyDown = useCallback((e: React.KeyboardEvent) => {
    if (e.key === 'Enter' || e.key === ' ') {
      e.preventDefault();
      inputRef.current?.click();
    }
  }, []);

  return (
    <div className="min-h-screen bg-slate-50 py-12 px-4 sm:px-6 lg:px-8">
      <div className="max-w-3xl mx-auto">
        <div className="text-center mb-10">
          <h1 className="text-4xl font-extrabold text-slate-900 mb-4">
            Beautify Jadwal ASC 🎨
          </h1>
          <p className="text-lg text-slate-600">
            Solusi merombak file jadwal dari ASC Timetables menjadi{' '}
            <span className="font-bold text-primary">Excel Grid yang Cantik</span>!
          </p>
        </div>

        <div className="bg-white rounded-2xl shadow-xl overflow-hidden">
          <div className="p-8 sm:p-12">
            <div className="max-w-xl mx-auto">
              <div
                role="button"
                tabIndex={0}
                aria-label="Upload file jadwal"
                className={`relative flex flex-col items-center justify-center w-full h-64 border-2 border-dashed rounded-xl cursor-pointer transition-colors duration-200 ease-in-out
                  ${loading ? 'bg-slate-50 border-slate-300 cursor-not-allowed' : ''}
                  ${dragOver ? 'bg-primary/10 border-primary' : 'bg-slate-50 border-primary/50 hover:bg-slate-100'}`}
                onDrop={handleDrop}
                onDragOver={handleDragOver}
                onDragLeave={handleDragLeave}
                onClick={() => !loading && inputRef.current?.click()}
                onKeyDown={handleKeyDown}
              >
                <div className="flex flex-col items-center justify-center pt-5 pb-6 pointer-events-none">
                  <Upload className="w-12 h-12 text-primary mb-4" />
                  <p className="mb-2 text-sm text-slate-600 text-center">
                    <span className="font-bold text-primary">Klik atau drop file di sini</span><br />
                    <span className="font-extrabold text-lg">Excel (.XLSX)</span>
                  </p>
                  <p className="text-xs text-slate-500 mt-2 text-center px-4">
                    Juga menerima .ROZ, .XML, .HTML — atau multiple file HTML
                  </p>
                </div>
                <input
                  ref={inputRef}
                  type="file"
                  className="hidden"
                  accept=".roz,.xml,.xlsx,.xls,.html,.htm"
                  multiple
                  disabled={loading}
                  onChange={e => e.target.files && handleFiles(e.target.files)}
                />
              </div>

              {loading && (
                <div className="mt-8">
                  <div className="flex items-center justify-between mb-2">
                    <span className="text-sm font-semibold text-primary">{progressMsg}</span>
                    <span className="text-sm font-semibold text-primary">{Math.round(progress)}%</span>
                  </div>
                  <div className="w-full bg-slate-200 rounded-full h-2.5">
                    <div className="bg-primary h-2.5 rounded-full transition-all duration-300 ease-out" style={{ width: `${progress}%` }} />
                  </div>
                  <button
                    onClick={abort}
                    className="mt-3 w-full flex items-center justify-center gap-2 text-sm text-red-600 hover:text-red-800 font-medium"
                  >
                    <XCircle className="w-4 h-4" /> Batalkan
                  </button>
                </div>
              )}

              {error && !loading && (
                <div className="mt-6 p-4 rounded-lg bg-red-50 border border-red-200 flex items-start">
                  <AlertCircle className="w-5 h-5 text-red-600 mt-0.5 mr-3 flex-shrink-0" />
                  <p className="text-sm text-red-700 whitespace-pre-line">{error}</p>
                </div>
              )}

              {success && !loading && (
                <div className="mt-6 p-4 rounded-lg bg-green-50 border border-green-200 flex items-start">
                  <FileSpreadsheet className="w-5 h-5 text-green-600 mt-0.5 mr-3 flex-shrink-0" />
                  <p className="text-sm text-green-700">{success}</p>
                </div>
              )}
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}

export default App;
