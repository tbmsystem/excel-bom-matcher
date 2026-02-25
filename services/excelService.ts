
import * as XLSX from 'xlsx';
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
  batchScriptLines.push(`echo Avvio copia file da "${sourcePathStr}" a "${targetPathStr}"...`);
  
  // Create Main Target Folder
  batchScriptLines.push(`if not exist "${targetPathStr}" mkdir "${targetPathStr}"`);
  
  // Create Subfolders for Extensions
  batchScriptLines.push(`if not exist "${targetPathStr}\\PDF" mkdir "${targetPathStr}\\PDF"`);
  batchScriptLines.push(`if not exist "${targetPathStr}\\DWG" mkdir "${targetPathStr}\\DWG"`);
  batchScriptLines.push(`if not exist "${targetPathStr}\\STEP" mkdir "${targetPathStr}\\STEP"`);
  
  batchScriptLines.push('');

  // Clone BOM data to avoid mutating original if passed by ref (though readFile creates new)
  // We use map to ensure we have a deep enough copy for the rows we modify
  const newBomData = bomData.map(row => [...(row || [])]);

  // Add Headers for new columns if it's the first row
  if (newBomData.length > 0) {
    newBomData[0][12] = 'Stato DB';
    newBomData[0][13] = 'Note';
    newBomData[0][14] = 'File PDF';
    newBomData[0][15] = 'File DWG';
    newBomData[0][16] = 'File STEP';
  }

  for (let i = 1; i < newBomData.length; i++) {
    const row = newBomData[i];
    
    // Ensure row has enough columns up to Q (Index 16)
    while (row.length < 17) {
      row.push(undefined);
    }

    // BOM Columns: E=4, F=5, G=6
    const code = normalize(row[4]);
    const config = normalize(row[5]);
    const rev = normalize(row[6]);

    if (!code) continue; 

    // --- DB Reconciliation Logic ---
    const fullKey = `${code}|${config}|${rev}`;
    const partialKey = `${code}|${config}`;
    const dbRow = dbMap.get(fullKey);

    if (dbRow) {
      matches++;
      // Copy DB Cols E,F,G,H,I -> BOM H,I,J,K,L (Indices 7-11)
      row[7] = dbRow[4];
      row[8] = dbRow[5];
      row[9] = dbRow[6];
      row[10] = dbRow[7];
      row[11] = dbRow[8];
      row[12] = 'OK';
      row[13] = ''; 
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

    // --- File Logic ---
    // Rule:
    // If Code == Config: Filename = Code + Rev
    // Else: Filename = Code + "_" + Config + Rev
    // Example: BA001218_RP000005-01A (Code_ConfigRev)
    
    let baseFileName = '';
    
    // Use raw values (trimmed) to preserve case sensitivity if needed, though we usually compare UPPER
    const rawCode = String(row[4] || '').trim();
    const rawConfig = String(row[5] || '').trim();
    const rawRev = String(row[6] || '').trim();

    if (normalize(rawCode) === normalize(rawConfig)) {
      // Case: BA102262A
      baseFileName = `${rawCode}${rawRev}`;
    } else {
      // Case: BA001218_RP000005-01A
      // Unire con '_' le 2 colonne codice e configurazione, poi aggiungere revisione
      baseFileName = `${rawCode}_${rawConfig}${rawRev}`;
    }

    // Sanitize filename (remove illegal chars just in case)
    baseFileName = baseFileName.replace(/[<>:"/\\|?*]/g, '');

    const checkAndLogFile = (ext: string, colIndex: number) => {
      const extUpper = ext.toUpperCase();
      let foundFilesList: string[] = [];

      // SPECIAL LOGIC FOR DWG: Wildcard Match
      if (extUpper === 'DWG') {
        // Find ALL files that start with the baseFileName
        // This captures "BA...-01.dwg", "BA...-01 DISTINTA.dwg", etc.
        const baseNameUpper = baseFileName.toUpperCase();
        
        // 1. Strict startsWith check on available DWGs
        foundFilesList = availableDwgFiles.filter(f => f.toUpperCase().startsWith(baseNameUpper));
        
        // 2. Padding logic for DWG if no matches found with standard rev
        if (foundFilesList.length === 0 && /^\d$/.test(rawRev)) {
           const paddedRev = rawRev.padStart(2, '0');
           let paddedFileName = '';
           if (normalize(rawCode) === normalize(rawConfig)) {
              paddedFileName = `${rawCode}${paddedRev}`;
           } else {
              paddedFileName = `${rawCode}_${rawConfig}${paddedRev}`;
           }
           const paddedBaseUpper = paddedFileName.toUpperCase();
           foundFilesList = availableDwgFiles.filter(f => f.toUpperCase().startsWith(paddedBaseUpper));
        }

      } else {
        // STANDARD LOGIC FOR PDF/STP: Exact Match (with existing padding logic)
        const targetNameUpper = `${baseFileName}.${ext}`.toUpperCase();
        let foundRealName = availableFiles.get(targetNameUpper);

        // Special check: if revision is single digit "1" but file has "01" (padding)
        if (!foundRealName && /^\d$/.test(rawRev)) {
          // Try constructing with padded revision
          const paddedRev = rawRev.padStart(2, '0');
          let paddedFileName = '';
          if (normalize(rawCode) === normalize(rawConfig)) {
            paddedFileName = `${rawCode}${paddedRev}`;
          } else {
            paddedFileName = `${rawCode}_${rawConfig}${paddedRev}`;
          }
          const paddedNameUpper = `${paddedFileName}.${ext}`.toUpperCase();
          foundRealName = availableFiles.get(paddedNameUpper);
        }

        if (foundRealName) {
          foundFilesList.push(foundRealName);
        }
      }
      
      // Process Found Files
      if (foundFilesList.length > 0) {
        row[colIndex] = 'SI';
        
        foundFilesList.forEach(realName => {
           uniqueFilesFound.add(realName);

           // Count specific extensions
           if (extUpper === 'PDF') uniquePDFs.add(realName);
           else if (extUpper === 'DWG') uniqueDWGs.add(realName);
           else if (extUpper === 'STP' || extUpper === 'STEP') uniqueSTPs.add(realName);
           
           // Determine subfolder based on extension
           let subFolder = 'OTHER';
           if (extUpper === 'PDF') subFolder = 'PDF';
           else if (extUpper === 'DWG') subFolder = 'DWG';
           else if (extUpper === 'STP' || extUpper === 'STEP') subFolder = 'STEP';

           // Add to batch script using the REAL filename found on disk
           uniqueBatchCommands.add(`if exist "${sourcePathStr}\\${realName}" copy "${sourcePathStr}\\${realName}" "${targetPathStr}\\${subFolder}\\${realName}" > nul`);
        });

      } else {
        row[colIndex] = 'NO';
      }
    };

    if (baseFileName) {
      // PDF (Col O / 14)
      checkAndLogFile('pdf', 14);
      // DWG (Col P / 15)
      checkAndLogFile('dwg', 15);
      // STP (Col Q / 16)
      checkAndLogFile('stp', 16); 
       
       // Fallback for STP to check .step if .stp fails
       if (row[16] === 'NO') {
          const targetNameUpperStep = `${baseFileName}.STEP`.toUpperCase();
          if (availableFiles.has(targetNameUpperStep)) {
             const realName = availableFiles.get(targetNameUpperStep);
             row[16] = 'SI';
             uniqueFilesFound.add(realName!);
             uniqueSTPs.add(realName!);
             // Destination: STEP subfolder
             uniqueBatchCommands.add(`if exist "${sourcePathStr}\\${realName}" copy "${sourcePathStr}\\${realName}" "${targetPathStr}\\STEP\\${realName}" > nul`);
          }
       }
    }
  }

  // Add unique commands to the script
  batchScriptLines.push(...uniqueBatchCommands);

  batchScriptLines.push('');
  batchScriptLines.push('echo Operazione completata.');
  batchScriptLines.push('pause');

  const stats: ProcessStats = {
    totalRows: newBomData.length - 1, // Exclude header
    matches,
    missing,
    revisionMismatch,
    filesFound: uniqueFilesFound.size,
    filesFoundDetails: {
      pdf: uniquePDFs.size,
      dwg: uniqueDWGs.size,
      stp: uniqueSTPs.size
    }
  };

  return {
    data: newBomData,
    fileName: `PROCESSED_${bomFile.name}`,
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
        if (!workbook.SheetNames.length) {
            return reject(new Error("Nessun foglio trovato nel file Excel"));
        }
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 }) as any[][];
        resolve(jsonData);
      } catch (err) {
        const errorMsg = err instanceof Error ? err.message : "Errore parsing Excel";
        reject(new Error(errorMsg));
      }
    };
    reader.onerror = () => reject(new Error("Errore durante la lettura del file (I/O Error)."));
    reader.readAsArrayBuffer(file);
  });
};

export const downloadExcel = (data: any[][], fileName: string) => {
  const worksheet = XLSX.utils.aoa_to_sheet(data);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Risultato");
  XLSX.writeFile(workbook, fileName);
};

export const downloadMissingRecords = (data: any[][], fileName: string) => {
  const header = data[0];
  const filteredRows = data.slice(1).filter(row => row[12] === 'Aggiungi a Dati DB');
  
  if (filteredRows.length === 0) return 0;

  const newData = [header, ...filteredRows];
  const cleanName = fileName.replace('PROCESSED_', '');
  downloadExcel(newData, `DA_AGGIUNGERE_${cleanName}`);
  return filteredRows.length;
};

export const downloadBatchScript = (content: string, originalFileName: string) => {
  const blob = new Blob([content], { type: 'text/plain;charset=utf-8' });
  const url = window.URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  const cleanName = originalFileName.replace('PROCESSED_', '').replace('.xlsx', '').replace('.xls', '');
  a.download = `COPIA_FILE_${cleanName}.bat`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  window.URL.revokeObjectURL(url);
};
