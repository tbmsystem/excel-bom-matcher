
// Versione Pro - Build per GitHub Actions
import React, { useState, useMemo, useEffect } from 'react';
import { ArrowRight, Download, RefreshCw, Database, FileText, CheckCircle2, FilePlus, Terminal, Filter, X, Layers } from 'lucide-react';
import Dropzone from './components/Dropzone';
import FolderSelector from './components/FolderSelector';
import StatsCard from './components/StatsCard';
import { processExcelFiles, downloadExcel, downloadMissingRecords, downloadBatchScript } from './services/excelService';
import { ProcessedResult } from './types';

const App: React.FC = () => {
  const [dbFile, setDbFile] = useState<File | null>(null);
  const [bomFile, setBomFile] = useState<File | null>(null);
  const [sourceFiles, setSourceFiles] = useState<FileList | null>(null);

  // New State for paths with persistence
  const [sourcePath, setSourcePath] = useState<string>(() => localStorage.getItem('tbm_source_path') || "");
  const [targetPath, setTargetPath] = useState<string>(() => localStorage.getItem('tbm_target_path') || "C:\\Tavole");

  // Persist paths to localStorage
  useEffect(() => {
    localStorage.setItem('tbm_source_path', sourcePath);
  }, [sourcePath]);

  useEffect(() => {
    localStorage.setItem('tbm_target_path', targetPath);
  }, [targetPath]);

  const [isProcessing, setIsProcessing] = useState(false);
  const [hideDuplicates, setHideDuplicates] = useState(false);
  const [result, setResult] = useState<ProcessedResult | null>(null);
  const [error, setError] = useState<string | null>(null);

  // Filter State: Key = Column Index, Value = Filter String
  const [filters, setFilters] = useState<Record<number, string>>({});

  const handleFolderSelect = (files: FileList) => {
    setSourceFiles(files);
    // Non sovrascriviamo sourcePath in automatico perché il browser non fornisce il percorso assoluto.
    // L'utente deve inserirlo manualmente nel campo "Percorso Sorgente".
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

  const handleDownload = async () => {
    if (result) {
      await downloadExcel(result.data, result.fileName);
    }
  };

  const handleDownloadMissing = async () => {
    if (result && result.stats.missing > 0) {
      await downloadMissingRecords(result.data, result.fileName);
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

    // Filter by criteria first
    const items = result.data.slice(1).filter(row => {
      return Object.entries(filters).every(([key, filterValue]) => {
        if (!filterValue) return true;
        const colIdx = parseInt(key);
        const cellValue = String((row as any)[colIdx] || '').toLowerCase();
        return cellValue.includes((filterValue as string).toLowerCase());
      });
    });

    // Handle duplicates if requested
    if (hideDuplicates) {
      const seen = new Set<string>();
      return items.filter(row => {
        // Unique key based on Codice (4), Config (5), Rev (6)
        const key = `${row[4]}|${row[5]}|${row[17]}|${row[6]}`; // Added Descrizione (17) to key as well for better accuracy
        if (seen.has(key)) return false;
        seen.add(key);
        return true;
      });
    }

    return items;
  }, [result, filters, hideDuplicates]);

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
    <div className="min-h-screen bg-slate-50/30 text-slate-800 font-sans w-full py-[5vh]">
      <div className="max-w-[85%] mx-auto bg-white min-h-[90vh] shadow-xl rounded-2xl overflow-hidden border border-slate-200 flex flex-col">
        <header className="border-b border-slate-200 w-full bg-white sticky top-0 z-50">
          <div className="w-full px-8 h-16 flex items-center justify-between">
            <div className="flex items-center gap-3">
              <div className="w-8 h-8 bg-blue-600 rounded flex items-center justify-center text-white font-bold text-lg">
                M
              </div>
              <h1 className="text-lg font-semibold text-blue-600">
                Elaborazione Distinta Base ed Esportazione Disegni Tecnici
              </h1>
            </div>
            {result && (
              <button
                onClick={reset}
                className="text-sm font-medium text-slate-400 hover:text-blue-600 flex items-center gap-2 transition-colors"
              >
                <RefreshCw size={14} /> Nuova Elaborazione
              </button>
            )}
          </div>
        </header>

        <main className="flex-1 w-full px-8 py-10">

          {/* Intro - Versione Originale */}
          <div className="mb-12 text-center w-full mx-auto">
            <h2 className="text-3xl font-bold text-slate-800 mb-4">Gestione Distinta Base Progetto</h2>
            <p className="text-slate-500 font-medium max-w-3xl mx-auto">
              Carica i file Excel e la cartella locale dei disegni. Il sistema verificherà i dati e genererà
              uno script per copiare i file trovati (.pdf, .dwg, .stp).
            </p>
          </div>

          <div className="grid md:grid-cols-2 gap-8 mb-12">

            {/* Left Column: Excel Files */}
            <div className="bg-white p-6 rounded-xl shadow-sm border border-slate-200">
              <h3 className="font-bold text-base text-slate-700 mb-6 flex items-center gap-2">
                <Database className="w-4 h-4 text-indigo-500" />
                File Excel
              </h3>

              <div className="flex flex-col gap-4">
                <Dropzone
                  label="1. Carica File DB"
                  description="File sorgente master"
                  file={dbFile}
                  onFileSelect={setDbFile}
                  onClear={() => { setDbFile(null); setResult(null); }}
                  colorClass="border-indigo-50 hover:bg-indigo-50/20 hover:border-indigo-200"
                />

                <Dropzone
                  label="2. Carica Distinta Base"
                  description="File da verificare e aggiornare"
                  file={bomFile}
                  onFileSelect={setBomFile}
                  onClear={() => { setBomFile(null); setResult(null); }}
                  colorClass="border-blue-50 hover:bg-blue-50/20 hover:border-blue-200"
                />
              </div>
            </div>

            {/* Right Column: File System */}
            <div className="bg-white p-6 rounded-xl shadow-sm border border-slate-200">
              <h3 className="font-bold text-base text-slate-700 mb-6 flex items-center gap-2">
                <FileText className="w-4 h-4 text-purple-500" />
                Gestione File Disegni
              </h3>

              <div className="flex flex-col gap-4">
                <FolderSelector
                  label="3. Seleziona Cartella Disegni"
                  fileCount={sourceFiles ? sourceFiles.length : 0}
                  onFolderSelect={handleFolderSelect}
                  onClear={() => { setSourceFiles(null); }}
                />

                <div className="p-4 bg-slate-50/50 rounded-lg border border-slate-200">
                  <h4 className="text-xs font-bold text-slate-500 mb-3 uppercase tracking-wider">Configurazione Script di Copia (.bat)</h4>

                  <div className="space-y-3">
                    <div>
                      <label className="block text-xs font-medium text-gray-400 mb-1">
                        Percorso Sorgente (Drive/Locale)
                      </label>
                      <input
                        type="text"
                        value={sourcePath}
                        onChange={(e) => setSourcePath(e.target.value)}
                        placeholder="es. Z:\Progetti\Disegni Condivisi"
                        className="w-full px-3 py-2 text-sm border border-slate-200 rounded-md focus:ring-1 focus:ring-blue-400 outline-none transition-all placeholder:text-slate-300"
                      />
                      <p className="text-[10px] text-gray-400 mt-1 italic">
                        * Il browser non fornisce il percorso completo (C:\...) per sicurezza. Inserisci il percorso completo se necessario.
                      </p>
                    </div>

                    <div>
                      <label className="block text-xs font-medium text-gray-400 mb-1">
                        Percorso Destinazione
                      </label>
                      <input
                        type="text"
                        value={targetPath}
                        onChange={(e) => setTargetPath(e.target.value)}
                        placeholder="es. C:\Tavole"
                        className="w-full px-3 py-2 text-sm border border-slate-100 rounded-md focus:ring-1 focus:ring-blue-400 outline-none transition-all"
                      />
                    </div>
                  </div>
                </div>
              </div>
            </div>

          </div>

          {/* Action Area - Versione Originale */}
          <div className="flex flex-col items-center justify-center mb-10 gap-4">
            <button
              onClick={handleProcess}
              disabled={!dbFile || !bomFile || isProcessing}
              className={`
              px-10 py-4 rounded-lg font-bold text-base transition-all
              ${(!dbFile || !bomFile)
                  ? 'bg-slate-100 text-slate-300 cursor-not-allowed border border-slate-200'
                  : 'bg-slate-600 text-white hover:bg-slate-700 active:scale-95 shadow-md'}
            `}
            >
              {isProcessing ? (
                <span className="flex items-center gap-2">
                  <RefreshCw size={18} className="animate-spin" /> Elaborazione...
                </span>
              ) : (
                "Analizza Excel e Genera Script"
              )}
            </button>

            <p className="text-[11px] text-orange-600 font-medium">
              * Per il controllo file e lo script di copia, assicurati di selezionare la cartella e inserire il percorso sorgente.
            </p>
          </div>

          {/* Error Message */}
          {error && (
            <div className="max-w-2xl mx-auto mb-8 p-4 bg-red-50 border border-red-100 rounded-lg text-red-700 text-center">
              {error}
            </div>
          )}

          {/* Results Area */}
          {result && (
            <div className="bg-white rounded-xl shadow-lg border border-slate-200 overflow-hidden">
              <div className="p-8 bg-white border-b border-slate-200">
                <h3 className="text-xl font-bold text-slate-800 mb-8 flex items-center gap-2">
                  <CheckCircle2 className="text-green-500" /> Risultato Analisi
                </h3>

                {/* Logic Stats Grid */}
                <div className="grid grid-cols-1 md:grid-cols-3 gap-6 mb-10">
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
                <h4 className="text-xs font-bold text-slate-400 uppercase tracking-widest mb-6 border-b border-slate-100 pb-3">Riepilogo File trovati su Disco</h4>
                <div className="grid grid-cols-2 md:grid-cols-4 gap-4 mb-10">
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

                <div className="grid md:grid-cols-2 gap-8">
                  {/* Excel Outputs */}
                  <div className="bg-blue-50/50 p-6 rounded-xl border border-blue-100">
                    <h4 className="font-bold text-blue-900 mb-4 text-sm uppercase">Export Excel</h4>
                    <div className="flex flex-col gap-3">
                      <button
                        onClick={handleDownload}
                        className="w-full px-4 py-3 bg-blue-600 hover:bg-blue-700 text-white rounded-lg font-bold flex items-center justify-center gap-2 transition-all shadow-sm"
                      >
                        <Download size={18} /> Scarica Report Completo
                      </button>
                      {result.stats.missing > 0 && (
                        <button
                          onClick={handleDownloadMissing}
                          className="w-full px-4 py-3 bg-white border border-blue-200 text-blue-600 hover:bg-blue-50 rounded-lg font-bold flex items-center justify-center gap-2 transition-all shadow-sm"
                        >
                          <FilePlus size={18} /> Scarica Solo Mancanti
                        </button>
                      )}
                    </div>
                  </div>

                  {/* Batch Script Output */}
                  <div className="bg-purple-50/50 p-6 rounded-xl border border-purple-100">
                    <h4 className="font-bold text-purple-900 mb-4 text-sm uppercase">Script di Copia</h4>
                    <p className="text-xs text-purple-700 mb-4">
                      Esegui questo file <strong>.bat</strong> sul tuo computer per copiare i file da <strong>{sourcePath || '...'}</strong> a <strong>{targetPath}</strong>.
                    </p>
                    <button
                      onClick={handleDownloadScript}
                      className="w-full px-4 py-3 bg-purple-600 hover:bg-purple-700 text-white rounded-lg font-bold flex items-center justify-center gap-2 transition-all shadow-sm"
                    >
                      <Terminal size={18} /> Scarica File .bat
                    </button>
                  </div>
                </div>
              </div>

              {/* Preview Table */}
              <div className="p-8">
                <div className="flex justify-between items-center mb-6">
                  <div className="flex items-center gap-4">
                    <h4 className="font-bold text-slate-700 flex items-center gap-2 text-sm uppercase">
                      <Filter size={16} /> Anteprima e Filtri
                    </h4>

                    <button
                      onClick={() => setHideDuplicates(!hideDuplicates)}
                      className={`
                        flex items-center gap-2 px-3 py-1.5 rounded-full text-[10px] font-bold uppercase transition-all
                        ${hideDuplicates
                          ? 'bg-blue-600 text-white shadow-md ring-2 ring-blue-100'
                          : 'bg-slate-100 text-slate-400 hover:bg-slate-200'}
                      `}
                    >
                      {hideDuplicates ? <CheckCircle2 size={12} /> : <Layers size={12} />}
                      {hideDuplicates ? 'Duplicati Nascosti' : 'Nascondi Duplicati'}
                    </button>
                  </div>
                  <div className="flex items-center gap-2">
                    <span className="text-xs font-medium text-slate-400">
                      Visualizzati: <strong className="text-blue-600">{filteredRows.length}</strong> / {result.stats.totalRows}
                    </span>
                  </div>
                </div>

                <div className="overflow-x-auto border border-slate-200 rounded-xl overflow-hidden w-full bg-white shadow-sm font-medium">
                  <table className="w-full text-[11px] text-left text-slate-600 border-collapse table-fixed">
                    <thead className="bg-slate-100/80 text-[10px] uppercase font-bold text-slate-500 sticky top-0 z-10 border-b border-slate-200">
                      <tr>
                        <th className="px-4 py-4 w-[20%] text-left">
                          <div className="mb-2">Descrizione</div>
                          <div className="relative group/filter">
                            <input
                              type="text"
                              placeholder="Filtra..."
                              className="w-full pl-2 pr-6 py-1.5 text-xs border border-slate-100 rounded bg-white/50 font-normal normal-case outline-none focus:border-blue-200 transition-colors"
                              value={filters[17] || ''}
                              onChange={(e) => handleFilterChange(17, e.target.value)}
                            />
                            {filters[17] && (
                              <button
                                onClick={() => handleFilterChange(17, '')}
                                className="absolute right-1 top-1/2 -translate-y-1/2 text-slate-300 hover:text-red-400 p-0.5"
                              >
                                <X size={10} />
                              </button>
                            )}
                          </div>
                        </th>
                        <th className="px-4 py-4 w-[12%]">
                          <div className="mb-2">Codice</div>
                          <div className="relative">
                            <input
                              type="text"
                              placeholder="Filtra..."
                              className="w-full pl-2 pr-6 py-1.5 text-xs border border-slate-100 rounded bg-white/50 font-normal normal-case outline-none focus:border-blue-200 transition-colors"
                              value={filters[4] || ''}
                              onChange={(e) => handleFilterChange(4, e.target.value)}
                            />
                            {filters[4] && (
                              <button onClick={() => handleFilterChange(4, '')} className="absolute right-1 top-1/2 -translate-y-1/2 text-slate-300 hover:text-red-400 p-0.5">
                                <X size={10} />
                              </button>
                            )}
                          </div>
                        </th>
                        <th className="px-4 py-4 w-[10%]">
                          <div className="mb-2">Conf</div>
                          <div className="relative">
                            <input
                              type="text"
                              placeholder="Filtra..."
                              className="w-full pl-2 pr-6 py-1.5 text-xs border border-slate-100 rounded bg-white/50 font-normal normal-case outline-none focus:border-blue-200 transition-colors"
                              value={filters[5] || ''}
                              onChange={(e) => handleFilterChange(5, e.target.value)}
                            />
                            {filters[5] && (
                              <button onClick={() => handleFilterChange(5, '')} className="absolute right-1 top-1/2 -translate-y-1/2 text-slate-300 hover:text-red-400 p-0.5">
                                <X size={10} />
                              </button>
                            )}
                          </div>
                        </th>
                        <th className="px-4 py-4 w-[8%]">
                          <div className="mb-2 text-center">Rev</div>
                          <div className="relative">
                            <select
                              className="w-full pl-1 pr-4 py-1.5 text-xs border border-slate-100 rounded bg-white/50 font-normal normal-case outline-none focus:border-blue-200 transition-colors appearance-none"
                              value={filters[6] || ''}
                              onChange={(e) => handleFilterChange(6, e.target.value)}
                            >
                              <option value="">Tutti</option>
                              {uniqueRevisions.map(rev => (
                                <option key={rev} value={rev}>{rev}</option>
                              ))}
                            </select>
                            {filters[6] && (
                              <button onClick={() => handleFilterChange(6, '')} className="absolute right-0.5 top-1/2 -translate-y-1/2 text-slate-300 hover:text-red-400 p-0.5">
                                <X size={10} />
                              </button>
                            )}
                          </div>
                        </th>
                        <th className="px-4 py-4 w-[8%]">
                          <div className="mb-2 text-center">PDF</div>
                          <div className="relative">
                            <select
                              className="w-full pl-1 pr-4 py-1.5 text-xs border border-slate-100 rounded bg-white/50 font-normal normal-case outline-none focus:border-blue-200 transition-colors appearance-none"
                              value={filters[14] || ''}
                              onChange={(e) => handleFilterChange(14, e.target.value)}
                            >
                              <option value="">Tutti</option>
                              <option value="SI">SI</option>
                              <option value="NO">NO</option>
                            </select>
                            {filters[14] && (
                              <button onClick={() => handleFilterChange(14, '')} className="absolute right-0.5 top-1/2 -translate-y-1/2 text-slate-300 hover:text-red-400 p-0.5">
                                <X size={10} />
                              </button>
                            )}
                          </div>
                        </th>
                        <th className="px-4 py-4 w-[8%]">
                          <div className="mb-2 text-center">DWG</div>
                          <div className="relative">
                            <select
                              className="w-full pl-1 pr-4 py-1.5 text-xs border border-slate-100 rounded bg-white/50 font-normal normal-case outline-none focus:border-blue-200 transition-colors appearance-none"
                              value={filters[15] || ''}
                              onChange={(e) => handleFilterChange(15, e.target.value)}
                            >
                              <option value="">Tutti</option>
                              <option value="SI">SI</option>
                              <option value="NO">NO</option>
                            </select>
                            {filters[15] && (
                              <button onClick={() => handleFilterChange(15, '')} className="absolute right-0.5 top-1/2 -translate-y-1/2 text-slate-300 hover:text-red-400 p-0.5">
                                <X size={10} />
                              </button>
                            )}
                          </div>
                        </th>
                        <th className="px-4 py-4 w-[8%]">
                          <div className="mb-2 text-center">STP</div>
                          <div className="relative">
                            <select
                              className="w-full pl-1 pr-4 py-1.5 text-xs border border-slate-100 rounded bg-white/50 font-normal normal-case outline-none focus:border-blue-200 transition-colors appearance-none"
                              value={filters[16] || ''}
                              onChange={(e) => handleFilterChange(16, e.target.value)}
                            >
                              <option value="">Tutti</option>
                              <option value="SI">SI</option>
                              <option value="NO">NO</option>
                            </select>
                            {filters[16] && (
                              <button onClick={() => handleFilterChange(16, '')} className="absolute right-0.5 top-1/2 -translate-y-1/2 text-slate-300 hover:text-red-400 p-0.5">
                                <X size={10} />
                              </button>
                            )}
                          </div>
                        </th>
                        <th className="px-4 py-4 w-[11%]">
                          <div className="mb-2">Stato</div>
                          <div className="relative">
                            <select
                              className="w-full pl-1 pr-4 py-1.5 text-xs border border-slate-100 rounded bg-white/50 font-normal normal-case outline-none focus:border-blue-200 transition-colors appearance-none"
                              value={filters[12] || ''}
                              onChange={(e) => handleFilterChange(12, e.target.value)}
                            >
                              <option value="">Tutti</option>
                              <option value="OK">OK</option>
                              <option value="Aggiungi">Da Aggiungere</option>
                            </select>
                            {filters[12] && (
                              <button onClick={() => handleFilterChange(12, '')} className="absolute right-0.5 top-1/2 -translate-y-1/2 text-slate-300 hover:text-red-400 p-0.5">
                                <X size={10} />
                              </button>
                            )}
                          </div>
                        </th>
                        <th className="px-4 py-4 w-[15%]">
                          <div className="mb-2">Note</div>
                          <div className="relative">
                            <select
                              className="w-full pl-1 pr-4 py-1.5 text-xs border border-slate-100 rounded bg-white/50 font-normal normal-case outline-none focus:border-blue-200 transition-colors appearance-none"
                              value={filters[13] || ''}
                              onChange={(e) => handleFilterChange(13, e.target.value)}
                            >
                              <option value="">Tutti</option>
                              <option value="Verificare revisione">Verificare Revisione</option>
                            </select>
                            {filters[13] && (
                              <button onClick={() => handleFilterChange(13, '')} className="absolute right-0.5 top-1/2 -translate-y-1/2 text-slate-300 hover:text-red-400 p-0.5">
                                <X size={10} />
                              </button>
                            )}
                          </div>
                        </th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100">
                      {filteredRows.map((row, idx) => (
                        <tr key={idx} className="hover:bg-slate-50/50 transition-colors">
                          <td className="px-4 py-2 truncate text-slate-500 italic">{String(row[17] || '-')}</td>
                          <td className="px-4 py-2 font-semibold text-slate-700 truncate">{String(row[4] || '-')}</td>
                          <td className="px-4 py-2 truncate">{String(row[5] || '-')}</td>
                          <td className="px-4 py-2 text-center">{String(row[6] || '-')}</td>
                          <td className={`px-4 py-2 text-center font-bold ${row[14] === 'SI' ? 'text-green-500' : 'text-slate-200'}`}>{String(row[14] || '')}</td>
                          <td className={`px-4 py-2 text-center font-bold ${row[15] === 'SI' ? 'text-green-500' : 'text-slate-200'}`}>{String(row[15] || '')}</td>
                          <td className={`px-4 py-2 text-center font-bold ${row[16] === 'SI' ? 'text-green-500' : 'text-slate-200'}`}>{String(row[16] || '')}</td>
                          <td className={`px-4 py-2 font-medium ${row[12] === 'OK' ? 'text-green-500' : 'text-red-400'}`}>
                            {String(row[12] || '')}
                          </td>
                          <td className="px-4 py-2 text-orange-400 font-medium italic truncate">{String(row[13] || '')}</td>
                        </tr>
                      ))}
                      {filteredRows.length === 0 && (
                        <tr>
                          <td colSpan={9} className="px-4 py-12 text-center text-slate-300 italic">
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
    </div>
  );
};

export default App;
