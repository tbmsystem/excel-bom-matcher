
export interface ProcessStats {
  totalRows: number;
  matches: number;
  missing: number;
  revisionMismatch: number;
  filesFound: number;
  filesFoundDetails: {
    pdf: number;
    dwg: number;
    stp: number;
  };
}

export interface ProcessedResult {
  data: any[][];
  fileName: string;
  stats: ProcessStats;
  batchScript: string;
}

export enum FileType {
  DB = 'DB',
  BOM = 'BOM'
}

// Represents a row in the excel file (array of cells)
export type ExcelRow = (string | number | null | undefined)[];