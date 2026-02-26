import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import { ProcessedResult, ProcessStats } from '../types';

// Helper to normalize cell data to string for key comparison
const normalize = (val: any): string => {
  if (val === null || val === undefined) return '';
  return String(val).trim().toUpperCase();
};

export const processExcelFiles = async (
  dbFile: File,
  bomFile: File,
  sourceFiles: FileList | null,
  sourcePathStr: string,
  targetPathStr: string
): Promise<ProcessedResult> => {
  // Normalize paths to ensure consistent concatenation throughout the service
  const cleanSourcePath = (sourcePathStr || '').replace(/[\\/]+$/, '');
  const cleanTargetPath = (targetPathStr || 'C:\\Tavole').replace(/[\\/]+$/, '');

  // 1. Read files
  const dbData = await readFile(dbFile);
  const bomData = await readFile(bomFile);

  // Validate Data
  if (!dbData || dbData.length === 0) throw new Error("Il file DB è vuoto o non valido.");
  if (!bomData || bomData.length === 0) throw new Error("Il file Distinta Base è vuoto o non valido.");

  // 2. Index the DB Data for fast lookup
  const dbMap = new Map<string, any[]>();
  const dbPartialMap = new Set<string>();

  // Start from row 1 (assuming row 0 is header)
  for (let i = 1; i < dbData.length; i++) {
    const row = dbData[i];
    if (!row || row.length < 4) continue;

    // DB Columns: B=1, C=2, D=3
    const code = normalize(row[1]);
    const config = normalize(row[2]);
    const rev = normalize(row[3]);

    if (!code) continue;

    const fullKey = `${code}|${config}|${rev}`;
    const partialKey = `${code}|${config}`;

    dbMap.set(fullKey, row);
    dbPartialMap.add(partialKey);
  }

  // 3. Index Source Files (if provided)
  const availableFiles = new Map<string, string>(); // Key: Normalized Name, Value: Real Name
  const availableDwgFiles: string[] = []; // Store raw names of all DWGs for wildcard search

  // Filter for specific extensions
  const allowedExtensions = new Set(['.PDF', '.DWG', '.STP', '.STEP']);

  if (sourceFiles) {
    for (let i = 0; i < sourceFiles.length; i++) {
      const file = sourceFiles[i];
      if (file && file.name) {
        const nameUpper = file.name.toUpperCase();
        const lastDotIndex = nameUpper.lastIndexOf('.');
        if (lastDotIndex !== -1) {
          const ext = nameUpper.substring(lastDotIndex);
          if (allowedExtensions.has(ext)) {
            // Map the UPPERCASE filename to the REAL filename (to preserve case for copying)
            availableFiles.set(nameUpper, file.name);

            if (ext === '.DWG') {
              availableDwgFiles.push(file.name);
            }
          }
        }
      }
    }
  }

  // 4. Process BOM Data
  let matches = 0;
  let missing = 0;
  let revisionMismatch = 0;

  let batchScriptLines: string[] = [];
  const uniqueBatchCommands = new Set<string>();
  const uniqueFilesFound = new Set<string>();

  // Specific counters per type (using Sets to count unique files only)
  const uniquePDFs = new Set<string>();
  const uniqueDWGs = new Set<string>();
  const uniqueSTPs = new Set<string>();

  // Initialize Batch Script Header
  batchScriptLines.push('@echo off');
  batchScriptLines.push('chcp 65001 > nul'); // Set encoding to UTF-8
  batchScriptLines.push(`echo Avvio copia file da "${cleanSourcePath}" a "${cleanTargetPath}"...`);

  // Create Main Target Folder
  batchScriptLines.push(`if not exist "${cleanTargetPath}" mkdir "${cleanTargetPath}"`);

  // Create Subfolders for Extensions
  batchScriptLines.push(`if not exist "${cleanTargetPath}\\PDF" mkdir "${cleanTargetPath}\\PDF"`);
  batchScriptLines.push(`if not exist "${cleanTargetPath}\\DWG" mkdir "${cleanTargetPath}\\DWG"`);
  batchScriptLines.push(`if not exist "${cleanTargetPath}\\STEP" mkdir "${cleanTargetPath}\\STEP"`);

  batchScriptLines.push('');

  // Clone BOM data to avoid mutating original
  const newBomData = bomData.map(row => [...(row || [])]);

  // Add Headers for columns 7-11 and rename E1, F1, G1 as requested
  if (newBomData.length > 0) {
    newBomData[0][4] = 'Codice';          // E1
    newBomData[0][5] = 'Configurazione';  // F1
    newBomData[0][6] = 'Revisione';       // G1

    newBomData[0][7] = 'Fornitore';
    newBomData[0][8] = 'Ordinato';
    newBomData[0][9] = 'Cons. Disegni';
    newBomData[0][10] = 'Codice STP';
    newBomData[0][11] = 'File Comuni';

    // Status columns for internal use
    newBomData[0][12] = 'Stato DB';
    newBomData[0][13] = 'Note';
    newBomData[0][14] = 'File PDF';
    newBomData[0][15] = 'File DWG';
    newBomData[0][16] = 'File STEP';
    newBomData[0][17] = 'Descrizione';
  }

  for (let i = 1; i < newBomData.length; i++) {
    const row = newBomData[i];
    while (row.length < 18) row.push(undefined);

    const code = normalize(row[4]);
    const config = normalize(row[5]);
    const rev = normalize(row[6]);

    if (!code) continue;

    const fullKey = `${code}|${config}|${rev}`;
    const partialKey = `${code}|${config}`;
    const dbRow = dbMap.get(fullKey);

    if (dbRow) {
      matches++;
      row[7] = dbRow[4];
      row[8] = dbRow[5];
      row[9] = dbRow[6];
      row[10] = dbRow[7];
      row[11] = dbRow[8];
      row[12] = 'OK';
      row[13] = '';
      row[17] = dbRow[0]; // Descrizione dalla colonna A del DB
    } else {
      missing++;
      row[12] = 'Aggiungi a Dati DB';
      if (dbPartialMap.has(partialKey)) {
        revisionMismatch++;
        row[13] = 'Verificare revisione';
      } else {
        row[13] = '';
      }
    }

    // File Logic
    let baseFileName = '';
    const rawCode = String(row[4] || '').trim();
    const rawConfig = String(row[5] || '').trim();
    const rawRev = String(row[6] || '').trim();

    if (normalize(rawCode) === normalize(rawConfig)) {
      baseFileName = `${rawCode}${rawRev}`;
    } else {
      baseFileName = `${rawCode}_${rawConfig}${rawRev}`;
    }
    baseFileName = baseFileName.replace(/[<>:"/\\|?*]/g, '');

    const checkAndLogFile = (exts: string[], colIndex: number) => {
      let foundFilesList: string[] = [];

      exts.forEach(ext => {
        const extUpper = ext.toUpperCase();
        if (extUpper === 'DWG') {
          const baseNameUpper = baseFileName.toUpperCase();
          let dwgMatches = availableDwgFiles.filter(f => f.toUpperCase().startsWith(baseNameUpper));
          if (dwgMatches.length === 0 && /^\d$/.test(rawRev)) {
            const paddedRev = rawRev.padStart(2, '0');
            let paddedFileName = (normalize(rawCode) === normalize(rawConfig)) ? `${rawCode}${paddedRev}` : `${rawCode}_${rawConfig}${paddedRev}`;
            dwgMatches = availableDwgFiles.filter(f => f.toUpperCase().startsWith(paddedFileName.toUpperCase()));
          }
          foundFilesList.push(...dwgMatches);
        } else {
          const targetNameUpper = `${baseFileName}.${ext}`.toUpperCase();
          let foundRealName = availableFiles.get(targetNameUpper);
          if (!foundRealName && /^\d$/.test(rawRev)) {
            const paddedRev = rawRev.padStart(2, '0');
            let paddedFileName = (normalize(rawCode) === normalize(rawConfig)) ? `${rawCode}${paddedRev}` : `${rawCode}_${rawConfig}${paddedRev}`;
            foundRealName = availableFiles.get(`${paddedFileName}.${ext}`.toUpperCase());
          }
          if (foundRealName) foundFilesList.push(foundRealName);
        }
      });

      if (foundFilesList.length > 0) {
        row[colIndex] = 'SI';
        new Set(foundFilesList).forEach(realName => {
          uniqueFilesFound.add(realName);
          const nameUpper = realName.toUpperCase();
          let subFolder = 'STEP';

          if (nameUpper.endsWith('.PDF')) {
            uniquePDFs.add(realName);
            subFolder = 'PDF';
          } else if (nameUpper.endsWith('.DWG')) {
            uniqueDWGs.add(realName);
            subFolder = 'DWG';
          } else if (nameUpper.endsWith('.STP') || nameUpper.endsWith('.STEP')) {
            uniqueSTPs.add(realName);
            subFolder = 'STEP';
          }

          uniqueBatchCommands.add(`if exist "${cleanSourcePath}\\${realName}" copy "${cleanSourcePath}\\${realName}" "${cleanTargetPath}\\${subFolder}\\${realName}" > nul`);
        });
      } else {
        row[colIndex] = 'NO';
      }
    };

    if (baseFileName) {
      checkAndLogFile(['pdf'], 14);
      checkAndLogFile(['dwg'], 15);
      checkAndLogFile(['stp', 'step'], 16);
    }
  }

  batchScriptLines.push(...uniqueBatchCommands);
  batchScriptLines.push('');
  batchScriptLines.push('echo Operazione completata.');
  batchScriptLines.push('pause');

  const stats: ProcessStats = {
    totalRows: newBomData.length - 1,
    matches,
    missing,
    revisionMismatch,
    filesFound: uniqueFilesFound.size,
    filesFoundDetails: { pdf: uniquePDFs.size, dwg: uniqueDWGs.size, stp: uniqueSTPs.size }
  };

  return {
    data: newBomData,
    fileName: bomFile.name.replace(/\.[^/.]+$/, "") + "_COMPILATO.xlsx",
    stats,
    batchScript: batchScriptLines.join('\r\n')
  };
};

const readFile = (file: File): Promise<any[][]> => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const data = e.target?.result;
      if (!data) return reject(new Error("Il file è vuoto"));
      try {
        const workbook = XLSX.read(data, { type: 'array' });
        if (!workbook.SheetNames.length) return reject(new Error("Nessun foglio trovato"));
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 }) as any[][];
        resolve(jsonData);
      } catch (err) {
        reject(new Error(err instanceof Error ? err.message : "Errore parsing Excel"));
      }
    };
    reader.onerror = () => reject(new Error("Errore lettura file."));
    reader.readAsArrayBuffer(file);
  });
};

export const downloadExcel = async (data: any[][], fileName: string, exportOnly12Cols: boolean = true) => {
  const dataToExport = exportOnly12Cols ? data.map(row => row.slice(0, 12)) : data;

  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Risultato');

  // Add data
  dataToExport.forEach((row) => {
    worksheet.addRow(row);
  });

  const greyColor = 'FFD3D3D3';

  // Apply Styling
  worksheet.eachRow((row, rowNumber) => {
    row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      if (colNumber > 12) return;

      // Base Border
      cell.border = {
        top: { style: 'thin', color: { argb: 'FFD3D3D3' } },
        left: { style: 'thin', color: { argb: 'FFD3D3D3' } },
        bottom: { style: 'thin', color: { argb: 'FFD3D3D3' } },
        right: { style: 'thin', color: { argb: 'FFD3D3D3' } }
      };

      cell.font = { size: 10, name: 'Calibri' };
      cell.alignment = { vertical: 'middle' };

      // Header specific
      if (rowNumber === 1) {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFD3D3D3' }
        };
        cell.font = { bold: true, size: 11, name: 'Calibri' };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
      }
    });
  });

  // Auto-filter
  worksheet.autoFilter = {
    from: 'A1',
    to: {
      row: 1,
      column: 12
    }
  };

  // Column widths
  const widths = [10, 30, 8, 45, 15, 20, 10, 20, 15, 15, 15, 15];
  worksheet.columns = widths.map(w => ({ width: w }));

  // Save to buffer and download
  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = window.URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = fileName;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  window.URL.revokeObjectURL(url);
};

export const downloadMissingRecords = async (data: any[][], fileName: string) => {
  const header = data[0];
  const filteredRows = data.slice(1).filter(row => row[12] === 'Aggiungi a Dati DB');
  if (filteredRows.length === 0) return 0;
  const newData = [header, ...filteredRows];
  const cleanName = fileName.replace('_COMPILATO.xlsx', '').replace('.xlsx', '').replace('.xls', '');
  await downloadExcel(newData, `DA_AGGIUNGERE_${cleanName}.xlsx`, true);
  return filteredRows.length;
};

export const downloadBatchScript = (content: string, originalFileName: string) => {
  const blob = new Blob([content], { type: 'text/plain;charset=utf-8' });
  const url = window.URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  const cleanName = originalFileName.replace('_COMPILATO', '').replace('.xlsx', '').replace('.xls', '');
  a.download = `COPIA_FILE_${cleanName}.bat`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  window.URL.revokeObjectURL(url);
};
