
import React, { useRef } from 'react';
import { FolderSearch, CheckCircle, X } from 'lucide-react';
import { clsx } from 'clsx';

interface FolderSelectorProps {
  label: string;
  fileCount: number;
  onFolderSelect: (files: FileList) => void;
  onClear: () => void;
}

const FolderSelector: React.FC<FolderSelectorProps> = ({
  label,
  fileCount,
  onFolderSelect,
  onClear
}) => {
  const inputRef = useRef<HTMLInputElement>(null);

  const handleClick = () => {
    inputRef.current?.click();
  };

  const handleChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files.length > 0) {
      onFolderSelect(e.target.files);
    }
  };

  if (fileCount > 0) {
    return (
      <div className="relative p-6 rounded-xl border-2 border-solid border-purple-300 bg-white shadow-md flex items-center gap-4">
        <div className="p-3 rounded-full bg-green-200 text-green-700">
          <CheckCircle size={24} />
        </div>
        <div className="flex-1 overflow-hidden">
          <p className="font-semibold text-gray-900">Cartella Selezionata</p>
          <p className="text-sm text-gray-500">{fileCount} file trovati per la scansione</p>
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
    <div
      onClick={handleClick}
      className={clsx(
        "cursor-pointer group relative flex flex-col items-center justify-center p-8 rounded-xl border-2 border-dashed transition-all bg-slate-50/50",
        "border-purple-300 hover:bg-purple-100/50 hover:border-purple-500"
      )}
    >
      <input
        ref={inputRef}
        type="file"
        className="hidden"
        // @ts-ignore - directory attributes are non-standard but supported
        webkitdirectory=""
        multiple
        onChange={handleChange}
      />

      <div className="mb-4 p-4 rounded-full bg-white shadow-sm group-hover:scale-110 transition-transform duration-200">
        <FolderSearch className="w-8 h-8 text-purple-400 group-hover:text-purple-600" />
      </div>

      <h3 className="text-lg font-semibold text-gray-900 mb-1">{label}</h3>
      <p className="text-sm text-gray-500 text-center max-w-[250px]">
        Clicca per selezionare la cartella locale contenente i disegni (.pdf, .dwg, .stp)
      </p>
    </div>
  );
};

export default FolderSelector;
