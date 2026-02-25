
import React, { useState, useMemo } from 'react';
import { ArrowRight, Download, RefreshCw, Database, FileText, CheckCircle2, FilePlus, Terminal, Filter } from 'lucide-react';
import Dropzone from './components/Dropzone';
import FolderSelector from './components/FolderSelector';
import StatsCard from './components/StatsCard';
import { processExcelFiles, downloadExcel, downloadMissingRecords, downloadBatchScript } from './services/excelService';
import { ProcessedResult } from './types';

const App: React.FC = () => {
  const [dbFile, setDbFile] = useState<File | null>(null);
  const [bomFile, setBomFile] = useState<File | null>(null);
  const [sourceFiles, setSourceFiles] = useState<FileList | null>(null);

  // New State for paths
  const [sourcePath, setSourcePath] = useState<string>("");
  const [targetPath, setTargetPath] = useState<string>("C:\\Tavole");

  const [isProcessing, setIsProcessing] = useState(false);
  const [result, setResult] = useState<ProcessedResult | null>(null);
  const [error, setError] = useState<string | null>(null);

  // Filter State: Key = Column Index, Value = Filter String
  const [filters, setFilters] = useState<Record<number, string>>({});

  const handleFolderSelect = (files: FileList) => {
    setSourceFiles(files);

    // Attempt to extract folder name from the first file's relative path
    if (files.length > 0) {
      const firstFile = files[0];
      if (firstFile.webkitRelativePath) {
        const pathParts = firstFile.webkitRelativePath.split('/');
        if (pathParts.length > 1) {
          const folderName = pathParts[0];
          setSourcePath(folderName);
        }
      }
    }
  };

  const handleProcess = async () => {
    if (!dbFile || !bomFile) return;

    if (sourceFiles && !sourcePath) {
      setError("Per generare lo script di copia, devi inserire il percorso 'Sorgente' testuale (es. Z:\\Disegni)");
      return;
    }

    setIsProcessing(true);
    setError(null);
    setFilters({}); // Reset filters on new process
    try {
      const res = await processExcelFiles(
        dbFile,
        bomFile,
        sourceFiles,
        sourcePath,
        targetPath
      );
      setResult(res);
    } catch (err: any) {
      let msg = 'Errore sconosciuto';
      if (err instanceof Error) {
        msg = err.message;
      } else if (typeof err === 'string') {
        msg = err;
      } else {
        msg = "Si è verificato un errore imprevisto durante l'elaborazione.";
      }
      setError(msg);
      console.error(err);
    } finally {
      setIsProcessing(false);
    }
  };

  const handleDownload = () => {
    if (result) {
      downloadExcel(result.data, result.fileName);
    }
  };

  const handleDownloadMissing = () => {
    if (result && result.stats.missing > 0) {
      downloadMissingRecords(result.data, result.fileName);
    }
  };

  const handleDownloadScript = () => {
    if (result && result.batchScript) {
      downloadBatchScript(result.batchScript, result.fileName);
    }
  };

  const reset = () => {
    setDbFile(null);
    setBomFile(null);
    setSourceFiles(null);
    setResult(null);
    setError(null);
    setFilters({});
  };

  const handleFilterChange = (colIndex: number, value: string) => {
    setFilters(prev => ({
      ...prev,
      [colIndex]: value
    }));
  };

  const handleFileFilter = (type: 'ALL' | 'PDF' | 'DWG' | 'STP') => {
    setFilters(prev => {
      const newFilters = { ...prev };
      // Reset file columns (14, 15, 16)
      delete newFilters[14];
      delete newFilters[15];
      delete newFilters[16];

      switch (type) {
        case 'PDF': newFilters[14] = 'SI'; break;
        case 'DWG': newFilters[15] = 'SI'; break;
        case 'STP': newFilters[16] = 'SI'; break;
        case 'ALL': break; // Just reset
      }
      return newFilters;
    });
  };

  // Extract unique revisions for the filter dropdown
  const uniqueRevisions = useMemo(() => {
    if (!result) return [];
    const revs = new Set<string>();
    // Start from index 1 to skip header
    for (let i = 1; i < result.data.length; i++) {
      const val = String(result.data[i][6] || '').trim();
      if (val) revs.add(val);
    }
    return Array.from(revs).sort();
  }, [result]);

  // Compute filtered rows
  const filteredRows = useMemo(() => {
    if (!result) return [];

    // Slice(1) to skip header row in data
    return result.data.slice(1).filter(row => {
      return Object.entries(filters).every(([key, filterValue]) => {
        if (!filterValue) return true;
        const colIdx = parseInt(key);
        const cellValue = String(row[colIdx] || '').toLowerCase();
        return cellValue.includes(filterValue.toLowerCase());
      });
    });
  }, [result, filters]);

  // Compute sorted file stats for display (Descending order: Left=Highest, Right=Lowest)
  const sortedFileStats = useMemo(() => {
    if (!result) return [];
    const stats = [
      {
        label: "PDF Trovati",
        value: result.stats.filesFoundDetails.pdf,
        type: 'pdf' as const,
        filter: 'PDF' as const
      },
      {
        label: "DWG Trovati",
        value: result.stats.filesFoundDetails.dwg,
        type: 'dwg' as const,
        filter: 'DWG' as const
      },
      {
        label: "STEP Trovati",
        value: result.stats.filesFoundDetails.stp,
        type: 'stp' as const,
        filter: 'STP' as const
      }
    ];
    // Sort descending by value (count)
    // If values are equal, sort alphabetically by label to ensure stability
    return stats.sort((a, b) => (b.value - a.value) || a.label.localeCompare(b.label));
  }, [result]);

  return (
    <div className="min-h-screen bg-slate-50 text-slate-900 pb-20 w-[90%] max-w-7xl mx-auto font-sans selection:bg-blue-100 selection:text-blue-900">
      {/* Header */}
      <header className="bg-white/80 backdrop-blur-md border-b border-slate-200 sticky top-0 z-50">
        <div className="w-full px-6 h-20 flex items-center justify-between">
          <div className="flex items-center gap-4">
            <div className="w-10 h-10 bg-gradient-to-br from-blue-600 to-indigo-700 rounded-xl flex items-center justify-center text-white font-bold text-xl shadow-lg shadow-blue-200">
              M
            </div>
            <div>
              <h1 className="text-xl font-extrabold bg-clip-text text-transparent bg-gradient-to-r from-blue-600 to-indigo-800 tracking-tight">
                Excel BOM Matcher Pro
              </h1>
              <p className="text-[10px] uppercase tracking-widest text-slate-400 font-bold">Industrial Design Reconciliation</p>
            </div>
          </div>
          {result && (
            <button
              onClick={reset}
              className="text-sm font-semibold text-slate-500 hover:text-blue-600 flex items-center gap-2 px-4 py-2 rounded-lg hover:bg-blue-50 transition-all active:scale-95"
            >
              <RefreshCw size={16} className="transition-transform group-hover:rotate-180" /> Nuova Analisi
            </button>
          )}
        </div>
      </header>

      <main className="w-full px-4 py-8">

        {/* Intro */}
        <div className="mb-12 text-center max-w-3xl mx-auto animate-fade-in-up">
          <h2 className="text-4xl font-black text-slate-900 mb-4 tracking-tight">Riconciliazione Distinta Base</h2>
          <p className="text-lg text-slate-500 font-medium">
            Carica i tuoi database Excel e la cartella dei disegni tecnici.
            Il sistema sincronizza i codici e prepara lo script di esportazione.
          </p>
        </div>

        <div className="grid md:grid-cols-2 gap-8 mb-8">

          {/* Left Column: Excel Files */}
          <div className="flex flex-col gap-6">
            <div className="bg-white p-5 rounded-2xl shadow-sm border border-gray-100">
              <h3 className="font-bold text-lg text-gray-800 mb-4 flex items-center gap-2">
                <Database className="w-5 h-5 text-indigo-600" />
                File Excel
              </h3>

              <div className="flex flex-col gap-4">
                <Dropzone
                  label="1. Carica File DB"
                  description="File sorgente master"
                  file={dbFile}
                  onFileSelect={setDbFile}
                  onClear={() => { setDbFile(null); setResult(null); }}
                  colorClass="border-indigo-200 hover:bg-indigo-50 hover:border-indigo-400"
                />

                <Dropzone
                  label="2. Carica Distinta Base"
                  description="File da verificare e aggiornare"
                  file={bomFile}
                  onFileSelect={setBomFile}
                  onClear={() => { setBomFile(null); setResult(null); }}
                  colorClass="border-blue-200 hover:bg-blue-50 hover:border-blue-400"
                />
              </div>
            </div>
          </div>

          {/* Right Column: File System */}
          <div className="flex flex-col gap-6">
            <div className="bg-white p-5 rounded-2xl shadow-sm border border-gray-100">
              <h3 className="font-bold text-lg text-gray-800 mb-4 flex items-center gap-2">
                <FileText className="w-5 h-5 text-purple-600" />
                Gestione File Disegni
              </h3>

              <div className="flex flex-col gap-4">
                {/* Folder Selection for Scanning */}
                <FolderSelector
                  label="3. Seleziona Cartella Disegni"
                  fileCount={sourceFiles ? sourceFiles.length : 0}
                  onFolderSelect={handleFolderSelect}
                  onClear={() => { setSourceFiles(null); }}
                />

                <div className="p-4 bg-gray-50 rounded-xl border border-gray-200">
                  <h4 className="text-sm font-semibold text-gray-700 mb-3">Configurazione Script di Copia (.bat)</h4>

                  <div className="space-y-3">
                    <div>
                      <label className="block text-xs font-medium text-gray-500 mb-1">
                        Percorso Sorgente (Drive/Locale)
                      </label>
                      <input
                        type="text"
                        value={sourcePath}
                        onChange={(e) => setSourcePath(e.target.value)}
                        placeholder="es. Z:\Progetti\Disegni Condivisi"
                        className="w-full px-3 py-2 text-sm border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 outline-none"
                      />
                      <p className="text-[10px] text-gray-400 mt-1">
                        *Il browser non fornisce il percorso completo (C:\...) per sicurezza. Inserisci il percorso completo se necessario.
                      </p>
                    </div>

                    <div>
                      <label className="block text-xs font-medium text-gray-500 mb-1">
                        Percorso Destinazione
                      </label>
                      <input
                        type="text"
                        value={targetPath}
                        onChange={(e) => setTargetPath(e.target.value)}
                        placeholder="es. C:\Tavole"
                        className="w-full px-3 py-2 text-sm border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 outline-none"
                      />
                    </div>
                  </div>
                </div>
              </div>
            </div>
          </div>

        </div>

        {/* Action Area */}
        <div className="flex flex-col items-center justify-center mb-10 gap-3">
          <button
            onClick={handleProcess}
            disabled={!dbFile || !bomFile || isProcessing}
            className={`
              relative px-12 py-5 rounded-2xl font-black text-lg shadow-xl transition-all transform hover:-translate-y-1 active:translate-y-0 active:scale-95
              ${(!dbFile || !bomFile)
                ? 'bg-slate-200 text-slate-400 cursor-not-allowed shadow-none'
                : 'bg-gradient-to-br from-blue-600 via-blue-700 to-indigo-800 text-white hover:shadow-2xl hover:shadow-blue-200 active:shadow-inner'}
            `}
          >
            {isProcessing ? (
              <span className="flex items-center gap-3 animate-pulse">
                <RefreshCw className="animate-spin" /> ELABORAZIONE IN CORSO...
              </span>
            ) : (
              <span className="flex items-center gap-3">
                <ArrowRight className="w-6 h-6" /> ANALIZZA E GENERA SCRIPT
              </span>
            )}
          </button>

          {(!sourceFiles || !sourcePath) && !result && (
            <p className="text-xs text-orange-600">
              * Per il controllo file e lo script di copia, assicurati di selezionare la cartella e inserire il percorso sorgente.
            </p>
          )}
        </div>

        {/* Error Message */}
        {error && (
          <div className="max-w-2xl mx-auto mb-8 p-4 bg-red-50 border border-red-200 rounded-lg text-red-700 text-center">
            {error}
          </div>
        )}

        {/* Results Area */}
        {result && (
          <div className="animate-fade-in-up bg-white rounded-2xl shadow-xl border border-gray-100 overflow-hidden">
            <div className="p-6 md:p-8 bg-slate-50 border-b border-gray-100">
              <h3 className="text-xl font-bold text-gray-800 mb-6 flex items-center gap-2">
                <CheckCircle2 className="text-green-500" /> Risultato Analisi
              </h3>

              {/* Logic Stats Grid */}
              <div className="grid grid-cols-1 md:grid-cols-3 gap-6 mb-8">
                <StatsCard
                  label="File Corretti"
                  value={result.stats.matches}
                  type="success"
                />
                <StatsCard
                  label="Mancanti in DatiDB"
                  value={result.stats.missing}
                  type="danger"
                />
                <StatsCard
                  label="Controlla Revisione"
                  value={result.stats.revisionMismatch}
                  type="warning"
                />
              </div>

              {/* File Breakdown Grid */}
              <h4 className="text-sm font-bold text-gray-500 tracking-wider mb-4 border-b pb-2">Riepilogo File trovati su Disco</h4>
              <div className="grid grid-cols-2 md:grid-cols-4 gap-4 mb-8">
                <StatsCard
                  label="Totale Trovati"
                  value={result.stats.filesFound}
                  type="info"
                  onClick={() => handleFileFilter('ALL')}
                />
                {sortedFileStats.map((stat) => (
                  <StatsCard
                    key={stat.type}
                    label={stat.label}
                    value={stat.value}
                    type={stat.type}
                    onClick={() => handleFileFilter(stat.filter)}
                  />
                ))}
              </div>

              <div className="grid md:grid-cols-2 gap-6">
                {/* Excel Outputs */}
                <div className="bg-blue-50 p-5 rounded-xl border border-blue-100">
                  <h4 className="font-semibold text-blue-900 mb-2">Export Excel</h4>
                  <div className="flex flex-col gap-2">
                    <button
                      onClick={handleDownload}
                      className="w-full px-4 py-2 bg-blue-600 hover:bg-blue-700 text-white rounded-lg font-medium flex items-center justify-center gap-2 transition-colors"
                    >
                      <Download size={18} /> Scarica Report Completo
                    </button>
                    {result.stats.missing > 0 && (
                      <button
                        onClick={handleDownloadMissing}
                        className="w-full px-4 py-2 bg-white border border-blue-200 text-blue-700 hover:bg-blue-100 rounded-lg font-medium flex items-center justify-center gap-2 transition-colors"
                      >
                        <FilePlus size={18} /> Scarica Solo Mancanti
                      </button>
                    )}
                  </div>
                </div>

                {/* Batch Script Output */}
                <div className="bg-purple-50 p-5 rounded-xl border border-purple-100">
                  <h4 className="font-semibold text-purple-900 mb-2">Script di Copia</h4>
                  <p className="text-xs text-purple-700 mb-3">
                    Esegui questo file .bat sul tuo computer per copiare i file da <strong>{sourcePath || '...'}</strong> a <strong>{targetPath}</strong>.
                  </p>
                  <button
                    onClick={handleDownloadScript}
                    className="w-full px-4 py-2 bg-purple-600 hover:bg-purple-700 text-white rounded-lg font-medium flex items-center justify-center gap-2 transition-colors"
                  >
                    <Terminal size={18} /> Scarica File .bat
                  </button>
                </div>
              </div>
            </div>

            {/* Simple Preview Table */}
            <div className="p-6 md:p-8">
              <div className="flex justify-between items-center mb-4">
                <h4 className="font-semibold text-gray-700 flex items-center gap-2">
                  <Filter size={16} /> Anteprima e Filtri
                </h4>
                <span className="text-sm text-gray-500">
                  Visualizzati: <strong>{filteredRows.length}</strong> / {result.stats.totalRows}
                </span>
              </div>

              <div className="overflow-x-auto border rounded-lg max-h-[600px] overflow-y-auto">
                <table className="w-full text-sm text-left text-gray-600">
                  <thead className="bg-gray-50 text-xs uppercase font-medium text-gray-500 sticky top-0 z-10 shadow-sm">
                    <tr>
                      <th className="px-4 py-3 bg-gray-50 w-[18%]">
                        <div className="mb-1">Codice</div>
                        <input
                          type="text"
                          placeholder="Filtra..."
                          className="w-full px-2 py-1 text-xs border border-gray-300 rounded font-normal normal-case"
                          value={filters[4] || ''}
                          onChange={(e) => handleFilterChange(4, e.target.value)}
                        />
                      </th>
                      <th className="px-4 py-3 bg-gray-50 w-[18%]">
                        <div className="mb-1">Conf</div>
                        <input
                          type="text"
                          placeholder="Filtra..."
                          className="w-full px-2 py-1 text-xs border border-gray-300 rounded font-normal normal-case"
                          value={filters[5] || ''}
                          onChange={(e) => handleFilterChange(5, e.target.value)}
                        />
                      </th>
                      <th className="px-4 py-3 bg-gray-50 w-[8%]">
                        <div className="mb-1">Rev</div>
                        <select
                          className="w-full px-1 py-1 text-xs border border-gray-300 rounded font-normal normal-case"
                          value={filters[6] || ''}
                          onChange={(e) => handleFilterChange(6, e.target.value)}
                        >
                          <option value="">Tutti</option>
                          {uniqueRevisions.map(rev => (
                            <option key={rev} value={rev}>{rev}</option>
                          ))}
                        </select>
                      </th>
                      <th className="px-4 py-3 bg-gray-50 w-[8%]">
                        <div className="mb-1">PDF</div>
                        <select
                          className="w-full px-1 py-1 text-xs border border-gray-300 rounded font-normal normal-case"
                          value={filters[14] || ''}
                          onChange={(e) => handleFilterChange(14, e.target.value)}
                        >
                          <option value="">Tutti</option>
                          <option value="SI">SI</option>
                          <option value="NO">NO</option>
                        </select>
                      </th>
                      <th className="px-4 py-3 bg-gray-50 w-[8%]">
                        <div className="mb-1">DWG</div>
                        <select
                          className="w-full px-1 py-1 text-xs border border-gray-300 rounded font-normal normal-case"
                          value={filters[15] || ''}
                          onChange={(e) => handleFilterChange(15, e.target.value)}
                        >
                          <option value="">Tutti</option>
                          <option value="SI">SI</option>
                          <option value="NO">NO</option>
                        </select>
                      </th>
                      <th className="px-4 py-3 bg-gray-50 w-[8%]">
                        <div className="mb-1">STP</div>
                        <select
                          className="w-full px-1 py-1 text-xs border border-gray-300 rounded font-normal normal-case"
                          value={filters[16] || ''}
                          onChange={(e) => handleFilterChange(16, e.target.value)}
                        >
                          <option value="">Tutti</option>
                          <option value="SI">SI</option>
                          <option value="NO">NO</option>
                        </select>
                      </th>
                      <th className="px-4 py-3 bg-gray-50 w-[16%]">
                        <div className="mb-1">Stato</div>
                        <select
                          className="w-full px-1 py-1 text-xs border border-gray-300 rounded font-normal normal-case"
                          value={filters[12] || ''}
                          onChange={(e) => handleFilterChange(12, e.target.value)}
                        >
                          <option value="">Tutti</option>
                          <option value="OK">OK</option>
                          <option value="Aggiungi">Da Aggiungere</option>
                        </select>
                      </th>
                      <th className="px-4 py-3 bg-gray-50 w-[16%]">
                        <div className="mb-1">Note</div>
                        <select
                          className="w-full px-1 py-1 text-xs border border-gray-300 rounded font-normal normal-case"
                          value={filters[13] || ''}
                          onChange={(e) => handleFilterChange(13, e.target.value)}
                        >
                          <option value="">Tutti</option>
                          <option value="Verificare revisione">Verificare Revisione</option>
                        </select>
                      </th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-gray-100">
                    {filteredRows.map((row, idx) => (
                      <tr key={idx} className="hover:bg-gray-50">
                        <td className="px-4 py-2 font-medium">{String(row[4] || '-')}</td>
                        <td className="px-4 py-2">{String(row[5] || '-')}</td>
                        <td className="px-4 py-2">{String(row[6] || '-')}</td>
                        <td className={`px-4 py-2 ${row[14] === 'SI' ? 'text-green-600 font-bold' : 'text-gray-400'}`}>{String(row[14] || '')}</td>
                        <td className={`px-4 py-2 ${row[15] === 'SI' ? 'text-green-600 font-bold' : 'text-gray-400'}`}>{String(row[15] || '')}</td>
                        <td className={`px-4 py-2 ${row[16] === 'SI' ? 'text-green-600 font-bold' : 'text-gray-400'}`}>{String(row[16] || '')}</td>
                        <td className={`px-4 py-2 ${row[12] === 'OK' ? 'text-green-600' : 'text-red-600'}`}>
                          {String(row[12] || '')}
                        </td>
                        <td className="px-4 py-2 text-orange-700 font-medium">{String(row[13] || '')}</td>
                      </tr>
                    ))}
                    {filteredRows.length === 0 && (
                      <tr>
                        <td colSpan={8} className="px-4 py-8 text-center text-gray-500 italic">
                          Nessun risultato corrisponde ai filtri selezionati.
                        </td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        )}
      </main>
    </div>
  );
};

export default App;
