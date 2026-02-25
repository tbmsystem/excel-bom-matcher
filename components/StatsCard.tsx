
import React from 'react';
import { CheckCircle2, AlertTriangle, FileWarning, FolderCheck } from 'lucide-react';

interface StatsCardProps {
  label: string;
  value: number | React.ReactNode;
  type: 'success' | 'warning' | 'danger' | 'info' | 'pdf' | 'dwg' | 'stp';
  onClick?: () => void;
}

// Custom component to render a file icon with the extension label inside
const FileTypeIcon = ({ ext }: { ext: string }) => (
  <svg 
    xmlns="http://www.w3.org/2000/svg" 
    width="24" 
    height="24" 
    viewBox="0 0 24 24" 
    fill="none" 
    stroke="currentColor" 
    strokeWidth="2" 
    strokeLinecap="round" 
    strokeLinejoin="round" 
    className="w-5 h-5"
  >
    <path d="M15 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V7Z" />
    <path d="M14 2v4a2 2 0 0 0 2 2h4" />
    <text 
      x="12" 
      y="20.5" 
      textAnchor="middle" 
      fontSize="6.5" 
      fontFamily="sans-serif"
      fontWeight="900"
      stroke="none"
      fill="currentColor"
    >
      {ext}
    </text>
  </svg>
);

const StatsCard: React.FC<StatsCardProps> = ({ label, value, type, onClick }) => {
  const styles = {
    success: 'bg-green-50 border-green-200 text-green-700',
    warning: 'bg-yellow-50 border-yellow-200 text-yellow-700',
    danger: 'bg-red-50 border-red-200 text-red-700',
    info: 'bg-purple-50 border-purple-200 text-purple-700',
    pdf: 'bg-rose-50 border-rose-200 text-rose-700',
    dwg: 'bg-blue-50 border-blue-200 text-blue-700',
    stp: 'bg-emerald-50 border-emerald-200 text-emerald-700',
  };

  const icons = {
    success: <CheckCircle2 className="w-5 h-5" />,
    warning: <FileWarning className="w-5 h-5" />,
    danger: <AlertTriangle className="w-5 h-5" />,
    info: <FolderCheck className="w-5 h-5" />,
    pdf: <FileTypeIcon ext="PDF" />,
    dwg: <FileTypeIcon ext="DWG" />,
    stp: <FileTypeIcon ext="STP" />,
  };

  const hoverStyles = onClick 
    ? 'cursor-pointer hover:ring-2 hover:ring-offset-1 hover:ring-opacity-50' 
    : '';
    
  // Add specific ring colors matching the card type
  const ringColors = {
    success: 'hover:ring-green-400',
    warning: 'hover:ring-yellow-400',
    danger: 'hover:ring-red-400',
    info: 'hover:ring-purple-400',
    pdf: 'hover:ring-rose-400',
    dwg: 'hover:ring-blue-400',
    stp: 'hover:ring-emerald-400',
  };

  return (
    <div 
      onClick={onClick}
      className={`p-4 rounded-xl border-2 shadow-sm flex flex-col gap-2 transition-all hover:scale-105 ${styles[type]} ${hoverStyles} ${onClick ? ringColors[type] : ''}`}
    >
      <div className="flex items-center gap-2 opacity-80">
        {icons[type]}
        <span className="font-semibold text-xs uppercase tracking-wider">{label}</span>
      </div>
      <div className="text-3xl font-bold">
        {value}
      </div>
    </div>
  );
};

export default StatsCard;
