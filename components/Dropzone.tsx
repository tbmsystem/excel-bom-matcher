import React, { useCallback } from 'react';
import { Upload, FileSpreadsheet, CheckCircle, X } from 'lucide-react';
import { clsx } from 'clsx';

interface DropzoneProps {
  label: string;
  description: string;
  file: File | null;
  onFileSelect: (file: File) => void;
  onClear: () => void;
  colorClass: string;
}

const Dropzone: React.FC<DropzoneProps> = ({ 
  label, 
  description, 
  file, 
  onFileSelect, 
  onClear,
  colorClass 
}) => {
  const handleDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      onFileSelect(e.dataTransfer.files[0]);
    }
  }, [onFileSelect]);

  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
  };

  const handleChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files.length > 0) {
      onFileSelect(e.target.files[0]);
    }
  };

  if (file) {
    return (
      <div className={clsx("relative p-6 rounded-xl border-2 border-solid transition-all bg-white shadow-sm flex items-center gap-4", colorClass)}>
        <div className="p-3 rounded-full bg-green-100 text-green-600">
          <CheckCircle size={24} />
        </div>
        <div className="flex-1 overflow-hidden">
          <p className="font-semibold text-gray-900 truncate">{file.name}</p>
          <p className="text-sm text-gray-500">{(file.size / 1024).toFixed(1)} KB</p>
        </div>
        <button 
          onClick={onClear}
          className="p-2 hover:bg-red-50 text-gray-400 hover:text-red-500 rounded-full transition-colors"
        >
          <X size={20} />
        </button>
      </div>
    );
  }

  return (
    <label 
      onDrop={handleDrop}
      onDragOver={handleDragOver}
      className={clsx(
        "cursor-pointer group relative flex flex-col items-center justify-center p-8 rounded-xl border-2 border-dashed transition-all hover:bg-opacity-50",
        colorClass
      )}
    >
      <input type="file" className="hidden" accept=".xlsx, .xls" onChange={handleChange} />
      
      <div className="mb-4 p-4 rounded-full bg-white shadow-sm group-hover:scale-110 transition-transform duration-200">
        <FileSpreadsheet className="w-8 h-8 text-gray-400 group-hover:text-gray-600" />
      </div>
      
      <h3 className="text-lg font-semibold text-gray-900 mb-1">{label}</h3>
      <p className="text-sm text-gray-500 text-center max-w-[200px]">{description}</p>
    </label>
  );
};

export default Dropzone;