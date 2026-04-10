import React, { useState, useRef, useEffect, useCallback, useMemo } from 'react';
import * as XLSX from 'xlsx';
import { HyperFormula } from 'hyperformula';
import { 
  Bold, Italic, Underline, AlignLeft, AlignCenter, AlignRight, 
  Download, Upload, Plus, Trash2, Printer, Check, Undo, Redo, Search, Save
} from 'lucide-react';

const colToAlpha = (col) => {
  let alpha = '';
  while (col >= 0) {
    alpha = String.fromCharCode((col % 26) + 65) + alpha;
    col = Math.floor(col / 26) - 1;
  }
  return alpha;
};

export default function ExcelEditor() {
  const hf = useRef(null);
  
  const [sheets, setSheets] = useState([{ id: 0, name: 'Sheet1', rows: 100, cols: 40, merges: [] }]);
  const [activeSheetId, setActiveSheetId] = useState(0);
  const [fileName, setFileName] = useState("Book1");
  const [formats, setFormats] = useState({});
  const [activeCell, setActiveCell] = useState({ r: 0, c: 0 });
  const [editMode, setEditMode] = useState(false);
  const [editValue, setEditValue] = useState("");
  const [isDirty, setIsDirty] = useState(false);
  const [history, setHistory] = useState([]);
  const [historyIndex, setHistoryIndex] = useState(-1);
  const [showSearch, setShowSearch] = useState(false);
  const [searchTerm, setSearchTerm] = useState("");
  const [lastSaved, setLastSaved] = useState(null);
  const [contextMenu, setContextMenu] = useState(null);
  const [colWidths, setColWidths] = useState({});
  const [resizing, setResizing] = useState(null);
  const [zoom, setZoom] = useState(1.0);

  const inputRef = useRef(null);
  const [tick, setTick] = useState(0);

  // Initialize HyperFormula
  if (!hf.current) {
    hf.current = HyperFormula.buildEmpty({ licenseKey: 'gpl-v3' });
    hf.current.addSheet('Sheet1');
    hf.current.setCellContents({ sheet: 0, row: 0, col: 0 }, [Array(40).fill("")]); // pre-warm
  }

  const currentSheet = useMemo(() => sheets.find(s => s.id === activeSheetId), [sheets, activeSheetId]);

  // Merge lookup for rendering
  const mergeInfo = useMemo(() => {
    const lookup = {};
    const hidden = new Set();
    currentSheet?.merges?.forEach(m => {
      const { s, e } = m;
      lookup[`${s.r}_${s.c}`] = { 
        rowspan: e.r - s.r + 1, 
        colspan: e.c - s.c + 1 
      };
      for (let r = s.r; r <= e.r; r++) {
        for (let c = s.c; c <= e.c; c++) {
          if (r === s.r && c === s.c) continue;
          hidden.add(`${r}_${c}`);
        }
      }
    });
    return { lookup, hidden };
  }, [currentSheet]);

  // Load state from LocalStorage on mount
  useEffect(() => {
    const saved = localStorage.getItem('excel_editor_state');
    if (saved) {
      try {
        const { sheets: s, formats: f, data: d, colWidths: cw } = JSON.parse(saved);
        setSheets(s);
        setFormats(f);
        setColWidths(cw || {});
        hf.current = HyperFormula.buildEmpty({ licenseKey: 'gpl-v3' });
        s.forEach((sheet, i) => {
          hf.current.addSheet(sheet.name);
          if (d[sheet.name]) {
            hf.current.setCellContents({ sheet: i, row: 0, col: 0 }, d[sheet.name]);
          }
        });
        setTick(t => t + 1);
      } catch (e) {
        console.error("Failed to load saved state", e);
      }
    }
  }, []);

  // Auto-save logic
  useEffect(() => {
    if (!isDirty) return;
    const timer = setInterval(() => {
      saveToLocalStorage();
    }, 30000);
    return () => clearInterval(timer);
  }, [isDirty, sheets, formats]);

  const saveToLocalStorage = useCallback(() => {
    const workbookData = {};
    sheets.forEach((s, i) => {
      const sheetData = [];
      for (let r = 0; r < s.rows; r++) {
        const row = [];
        for (let c = 0; c < s.cols; c++) {
          const formula = hf.current.getCellFormula({ sheet: i, row: r, col: c });
          const val = hf.current.getCellValue({ sheet: i, row: r, col: c });
          row.push(formula ? "=" + formula : (val !== null && val !== undefined ? (val.value ?? val) : ""));
        }
        sheetData.push(row);
      }
      workbookData[s.name] = sheetData;
    });

    localStorage.setItem('excel_editor_state', JSON.stringify({
      sheets,
      formats,
      colWidths,
      data: workbookData
    }));
    setIsDirty(false);
    setLastSaved(new Date().toLocaleTimeString());
  }, [sheets, formats, colWidths]);

  const pushHistory = useCallback(() => {
    const workbookData = {};
    sheets.forEach((s, i) => {
      const sheetData = [];
      for (let r = 0; r < s.rows; r++) {
        const row = [];
        for (let c = 0; c < s.cols; c++) {
          const formula = hf.current.getCellFormula({ sheet: i, row: r, col: c });
          const val = hf.current.getCellValue({ sheet: i, row: r, col: c });
          row.push(formula ? "=" + formula : (val !== null && val !== undefined ? (val.value ?? val) : ""));
        }
        sheetData.push(row);
      }
      workbookData[s.name] = sheetData;
    });

    const snapshot = { sheets, formats, colWidths, data: workbookData };
    setHistory(prev => {
      const newHist = prev.slice(0, historyIndex + 1);
      newHist.push(JSON.stringify(snapshot));
      if (newHist.length > 50) newHist.shift();
      return newHist;
    });
    setHistoryIndex(prev => Math.min(prev + 1, 49));
    setIsDirty(true);
  }, [sheets, formats, colWidths, historyIndex]);

  const undo = () => {
    if (historyIndex <= 0) return;
    const prevSnapshot = JSON.parse(history[historyIndex - 1]);
    applySnapshot(prevSnapshot);
    setHistoryIndex(historyIndex - 1);
  };

  const redo = () => {
    if (historyIndex >= history.length - 1) return;
    const nextSnapshot = JSON.parse(history[historyIndex + 1]);
    applySnapshot(nextSnapshot);
    setHistoryIndex(historyIndex + 1);
  };

  const applySnapshot = (snapshot) => {
    const { sheets: s, formats: f, data: d, colWidths: cw } = snapshot;
    setSheets(s);
    setFormats(f);
    setColWidths(cw || {});
    hf.current = HyperFormula.buildEmpty({ licenseKey: 'gpl-v3' });
    s.forEach((sheet, i) => {
      hf.current.addSheet(sheet.name);
      if (d[sheet.name]) {
        hf.current.setCellContents({ sheet: i, row: 0, col: 0 }, d[sheet.name]);
      }
    });
    setTick(t => t + 1);
  };

  // Keyboard navigation
  useEffect(() => {
    const onKeyDown = (e) => {
      if (editMode) return;
      if (!activeCell || !currentSheet) return;
      let { r, c } = activeCell;
      
      switch(e.key) {
        case 'ArrowUp': r = Math.max(0, r-1); break;
        case 'ArrowDown': r = Math.min(currentSheet.rows-1, r+1); break;
        case 'ArrowLeft': c = Math.max(0, c-1); break;
        case 'ArrowRight':
        case 'Tab': 
           e.preventDefault();
           c = Math.min(currentSheet.cols-1, c+1); 
           break;
        case 'Enter':
           e.preventDefault();
           startEdit(r, c);
           return;
        case 'Delete':
        case 'Backspace':
           hf.current.setCellContents({ sheet: activeSheetId, row: r, col: c }, [[""]]);
           setTick(t=>t+1);
           setIsDirty(true);
           return;
        default:
           if (e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) {
             setEditMode(true);
             setEditValue(e.key);
           }
           return;
      }
      setActiveCell({ r, c });
    };
    window.addEventListener('keydown', onKeyDown);
    return () => window.removeEventListener('keydown', onKeyDown);
  }, [editMode, activeCell, currentSheet, activeSheetId]);

  useEffect(() => {
    if (editMode && inputRef.current) {
      inputRef.current.focus();
    }
  }, [editMode]);

  const startEdit = (r, c) => {
    const formula = hf.current.getCellFormula({ sheet: activeSheetId, row: r, col: c });
    const val = hf.current.getCellValue({ sheet: activeSheetId, row: r, col: c });
    let initVal = "";
    if (formula) initVal = "=" + formula;
    else if (val !== null && val !== undefined) initVal = String(val.value ?? val);
    
    setActiveCell({ r, c });
    setEditMode(true);
    setEditValue(initVal);
  };

  const commitEdit = () => {
    if (!editMode || !activeCell) return;
    hf.current.setCellContents({ sheet: activeSheetId, row: activeCell.r, col: activeCell.c }, [[editValue]]);
    setEditMode(false);
    setTick(t=>t+1);
    setIsDirty(true);
  };

  const handleInputKeyDown = (e) => {
    if (e.key === 'Enter') {
       commitEdit();
       setActiveCell(prev => ({ r: Math.min(currentSheet.rows-1, prev.r + 1), c: prev.c }));
    } else if (e.key === 'Escape') {
       setEditMode(false);
    }
  };

  const handleContextMenu = (e, r, c) => {
    e.preventDefault();
    setActiveCell({ r, c });
    setContextMenu({ x: e.pageX, y: e.pageY });
  };

  const gridAction = (type) => {
    const { r, c } = activeCell;
    const newSheets = [...sheets];
    const sheet = newSheets.find(s => s.id === activeSheetId);
    
    // In a production app, we'd use HyperFormula's internal move/insert methods
    // For this simple grid, we manipulate the metadata and hf will adjust
    switch (type) {
      case 'insert-row':
        hf.current.addRows(activeSheetId, [r, 1]);
        sheet.rows += 1;
        break;
      case 'delete-row':
        hf.current.removeRows(activeSheetId, [r, 1]);
        sheet.rows -= 1;
        break;
      case 'insert-col':
        hf.current.addColumns(activeSheetId, [c, 1]);
        sheet.cols += 1;
        break;
      case 'delete-col':
        hf.current.removeColumns(activeSheetId, [c, 1]);
        sheet.cols -= 1;
        break;
    }
    setSheets(newSheets);
    setTick(t => t + 1);
    setIsDirty(true);
    setContextMenu(null);
    pushHistory();
  };

  const startResizing = (e, c) => {
    e.preventDefault();
    setResizing({ index: c, startX: e.pageX, startWidth: colWidths[c] || 96 }); // default 24rem * 4 (96px)
  };

  useEffect(() => {
    const handleMouseMove = (e) => {
      if (!resizing) return;
      const delta = e.pageX - resizing.startX;
      setColWidths(prev => ({ ...prev, [resizing.index]: Math.max(40, resizing.startWidth + delta) }));
    };
    const handleMouseUp = () => {
      if (resizing) {
        setResizing(null);
        setIsDirty(true);
        pushHistory();
      }
    };
    if (resizing) {
      window.addEventListener('mousemove', handleMouseMove);
      window.addEventListener('mouseup', handleMouseUp);
    }
    return () => {
      window.removeEventListener('mousemove', handleMouseMove);
      window.removeEventListener('mouseup', handleMouseUp);
    };
  }, [resizing]);

  const handleFormatChange = (key, valToggle) => {
    if (!activeCell) return;
    setFormats(prev => {
      const formatKey = `${activeSheetId}_${activeCell.r}_${activeCell.c}`;
      if (key === 'clear') {
        const { [formatKey]: _, ...rest } = prev;
        return rest;
      }
      const currentFormat = prev[formatKey] || {};
      const newFormat = { ...currentFormat };
      if (valToggle === undefined) {
        newFormat[key] = !newFormat[key];
      } else {
        newFormat[key] = valToggle;
      }
      return { ...prev, [formatKey]: newFormat };
    });
    setIsDirty(true);
    pushHistory();
  };

  const importFile = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (e) => {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array', cellFormula: true });
      
      hf.current = HyperFormula.buildEmpty({ licenseKey: 'gpl-v3' });
      const newSheets = [];
      const newFormats = {};
      
      workbook.SheetNames.forEach((name, i) => {
        hf.current.addSheet(name);
        const ws = workbook.Sheets[name];
        const range = ws['!ref'] ? XLSX.utils.decode_range(ws['!ref']) : {s:{r:0,c:0}, e:{r:39,c:19}};
        const merges = ws['!merges'] || [];
        
        let initialData = [];
        for (let R = 0; R <= range.e.r; R++) {
           let row = [];
           for (let C = 0; C <= range.e.c; C++) {
             const cellAddr = XLSX.utils.encode_cell({r: R, c: C});
             const cell = ws[cellAddr];
             if (cell) {
               row.push(cell.f ? "=" + cell.f : (cell.v !== undefined ? cell.v : ""));
             } else {
               row.push("");
             }
           }
           initialData.push(row);
        }
        hf.current.setCellContents({ sheet: i, row: 0, col: 0 }, initialData);
        newSheets.push({ 
          id: i, 
          name, 
          rows: Math.max(40, range.e.r + 5), 
          cols: Math.max(20, range.e.c + 5),
          merges: merges.map(m => ({ s: m.s, e: m.e }))
        });
      });
      
      setSheets(newSheets);
      setActiveSheetId(0);
      setFormats(newFormats);
      setFileName(file.name.replace(/\.[^/.]+$/, ""));
      setTick(t => t + 1);
      setIsDirty(true);
    };
    reader.readAsArrayBuffer(file);
  };

  const exportExport = (type) => {
    const wb = XLSX.utils.book_new();
    sheets.forEach(sheetInfo => {
      const sheetData = [];
      for (let r = 0; r < sheetInfo.rows; r++) {
        const rowData = [];
        for (let c = 0; c < sheetInfo.cols; c++) {
          const val = hf.current.getCellValue({ sheet: sheetInfo.id, row: r, col: c });
          const formula = hf.current.getCellFormula({ sheet: sheetInfo.id, row: r, col: c });
          if (formula) rowData.push({ f: formula, v: val && val.value !== undefined ? val.value : val });
          else rowData.push(val === null || val === undefined ? "" : (val.value !== undefined ? val.value : val));
        }
        sheetData.push(rowData);
      }
      const ws = XLSX.utils.aoa_to_sheet(sheetData);
      ws['!merges'] = sheetInfo.merges;
      XLSX.utils.book_append_sheet(wb, ws, sheetInfo.name);
    });
    XLSX.writeFile(wb, `workbook.${type}`);
  };

  const printPdf = () => {
    window.print();
  };

  const addSheet = () => {
    const name = `Sheet${sheets.length + 1}`;
    hf.current.addSheet(name);
    setSheets(prev => [...prev, { id: prev.length, name, rows: 100, cols: 40, merges: [] }]);
    setActiveSheetId(sheets.length);
    setIsDirty(true);
  };

  const removeSheet = (id) => {
    if (sheets.length <= 1) {
      alert("Cannot delete the only sheet in the workbook.");
      return;
    }
    if (window.confirm(`Are you sure you want to delete "${sheets.find(s=>s.id === id).name}"? This cannot be undone.`)) {
      setSheets(prev => prev.filter(s => s.id !== id));
      if (activeSheetId === id) {
        setActiveSheetId(sheets.find(s => s.id !== id).id);
      }
      setIsDirty(true);
      pushHistory();
    }
  };

  const activeFormat = activeCell ? formats[`${activeSheetId}_${activeCell.r}_${activeCell.c}`] || {} : {};

  // Formula bar value sync
  let formulaBarValue = "";
  if (editMode) {
     formulaBarValue = editValue;
  } else if (activeCell) {
     const f = hf.current.getCellFormula({ sheet: activeSheetId, row: activeCell.r, col: activeCell.c });
     const v = hf.current.getCellValue({ sheet: activeSheetId, row: activeCell.r, col: activeCell.c });
     if (f) formulaBarValue = "=" + f;
     else if (v !== null && v !== undefined) formulaBarValue = String(v.value ?? v);
  }
  const executeSearch = () => {
    if (!searchTerm) return;
    const term = searchTerm.toLowerCase();
    for (let r = 0; r < currentSheet.rows; r++) {
      for (let c = 0; c < currentSheet.cols; c++) {
        let val = hf.current.getCellValue({ sheet: activeSheetId, row: r, col: c });
        if (val && typeof val === 'object' && val.value !== undefined) val = String(val.value);
        if (val !== null && val !== undefined && String(val).toLowerCase().includes(term)) {
           setActiveCell({r, c});
           return;
        }
      }
    }
  };

  return (
    <div className="flex flex-col h-screen w-screen bg-[#f3f2f1] font-sans text-[13px] overflow-hidden print:h-auto print:w-auto print:overflow-visible print:block" onClick={() => setContextMenu(null)}>
      {/* Title Bar */}
      <div className="flex items-center justify-between bg-[#107c41] text-white h-[40px] px-3 shrink-0">
         <div className="flex items-center gap-4">
            <div className="flex bg-white/20 p-1 rounded-sm shadow-sm">
               <Save size={16} className={isDirty ? "animate-pulse text-[#ffea00]" : "text-white"}/>
            </div>
            <div className="flex items-center gap-2">
               <button onClick={undo} disabled={historyIndex <= 0} className="p-1 hover:bg-white/20 rounded-sm disabled:opacity-30" title="Undo"><Undo size={14}/></button>
               <button onClick={redo} disabled={historyIndex >= history.length - 1} className="p-1 hover:bg-white/20 rounded-sm disabled:opacity-30" title="Redo"><Redo size={14}/></button>
            </div>
            <div className="w-px h-5 bg-white/20"></div>
            <span className="font-semibold text-[14px] tracking-tight ml-2">{fileName} - Excel Clone</span>
         </div>
         {lastSaved && <div className="text-[11px] opacity-90 font-medium">Saved: {lastSaved}</div>}
      </div>

      {/* Ribbon Tabs */}
      <div className="flex items-center bg-[#f3f2f1] pt-2 px-2 gap-2 shrink-0 border-b border-gray-300">
         <div className="px-3 py-1 text-white bg-[#107c41] font-semibold cursor-pointer rounded-t-sm hover:bg-[#0c5e31] transition-colors leading-none tracking-wide text-xs">File</div>
         <div className="px-3 py-1 text-[#107c41] bg-white border border-b-0 border-gray-300 font-bold cursor-pointer rounded-t-sm relative top-[1px] leading-none tracking-wide text-xs">Home</div>
      </div>

      {/* Ribbon Panel (Home) */}
      <div className="flex items-start bg-white h-24 shrink-0 border-b border-gray-300 shadow-sm z-30 px-3 py-2 gap-6 select-none">
         
         {/* Font Group */}
         <div className="flex flex-col border-r border-gray-300 pr-6 h-full relative group/ribbontool">
            <div className="flex gap-1 mb-2">
               <div className="relative group cursor-pointer" title="Font Family">
                 <select 
                    className="border border-gray-300 rounded px-1 min-h-[24px] text-xs w-36 focus:outline-none focus:border-[#107c41] bg-white hover:bg-gray-50"
                    value={activeFormat.fontFamily || 'Calibri'}
                    onChange={e => handleFormatChange('fontFamily', e.target.value)}
                 >
                    <option value="Arial">Arial</option>
                    <option value="Calibri">Calibri</option>
                    <option value="'Courier New'">Courier New</option>
                    <option value="Tahoma">Tahoma</option>
                    <option value="'Times New Roman'">Times New Roman</option>
                    <option value="Verdana">Verdana</option>
                 </select>
               </div>
               <div className="relative group cursor-pointer" title="Font Size">
                 <select 
                    className="border border-gray-300 rounded px-1 min-h-[24px] text-xs w-16 focus:outline-none focus:border-[#107c41] bg-white hover:bg-gray-50 text-center"
                    value={activeFormat.fontSize || 13}
                    onChange={e => handleFormatChange('fontSize', parseInt(e.target.value))}
                 >
                    {[8,9,10,11,12,13,14,16,18,20,24,28,36,48,72].map(s => (
                       <option key={s} value={s}>{s}</option>
                    ))}
                 </select>
               </div>
               <button onClick={() => handleFormatChange('clear', true)} className="ml-2 px-2 h-6 text-gray-500 hover:bg-gray-100 hover:text-red-500 border border-transparent rounded transition-colors" title="Clear Formatting">
                  <Trash2 size={14}/>
               </button>
            </div>
            <div className="flex items-center gap-1 px-1">
               <button onClick={() => handleFormatChange('bold')} className={`p-1 hover:bg-[#d2e0cc] border border-transparent rounded ${activeFormat.bold?'bg-[#c3d5bb] border-gray-400 font-bold text-black':''}`} title="Bold (Ctrl+B)"><Bold size={14} className="stroke-[3]"/></button>
               <button onClick={() => handleFormatChange('italic')} className={`p-1 hover:bg-[#d2e0cc] border border-transparent rounded ${activeFormat.italic?'bg-[#c3d5bb] border-gray-400 text-black':''}`} title="Italic (Ctrl+I)"><Italic size={14}/></button>
               <button onClick={() => handleFormatChange('underline')} className={`p-1 hover:bg-[#d2e0cc] border border-transparent rounded ${activeFormat.underline?'bg-[#c3d5bb] border-gray-400 text-black':''}`} title="Underline (Ctrl+U)"><Underline size={14}/></button>
               <div className="w-px h-5 bg-gray-300 mx-2 shrink-0"></div>
               <div className="relative group/pick flex items-center justify-center">
                  <input type="color" className="absolute opacity-0 w-8 h-8 cursor-pointer z-10" value={activeFormat.bg || '#ffffff'} onChange={(e) => handleFormatChange('bg', e.target.value)} title="Fill Color"/>
                  <button className="p-1 px-1.5 hover:bg-[#d2e0cc] border border-transparent rounded text-gray-700 flex flex-col items-center">
                     <div className="w-4 h-3 border border-gray-400" style={{backgroundColor: activeFormat.bg || '#fff'}}></div>
                     <div className="w-4 h-1 mt-0.5 bg-yellow-400"></div>
                  </button>
               </div>
               <div className="relative group/pick flex items-center justify-center">
                  <input type="color" className="absolute opacity-0 w-8 h-8 cursor-pointer z-10" value={activeFormat.color || '#000000'} onChange={(e) => handleFormatChange('color', e.target.value)} title="Font Color"/>
                  <button className="p-1 px-1.5 hover:bg-[#d2e0cc] border border-transparent rounded flex flex-col items-center text-red-600 font-bold" style={{lineHeight:'12px'}}>
                     <span className="text-[14px] font-serif" style={{color: activeFormat.color || '#000'}}>A</span>
                     <div className="w-4 h-1 mt-0.5" style={{backgroundColor: activeFormat.color || '#000'}}></div>
                  </button>
               </div>
            </div>
            <div className="text-[11px] text-gray-400 text-center mt-auto w-full font-medium tracking-tight">Font</div>
         </div>

         {/* Alignment Group */}
         <div className="flex flex-col border-r border-gray-300 pr-6 h-full relative">
             <div className="flex flex-col gap-1 mt-1">
               <div className="flex gap-1 px-1 opacity-50 pointer-events-none" title="Vertical Alignment">
                 <button className="p-1 border border-transparent rounded"><AlignLeft size={14} className="rotate-90"/></button>
                 <button className="p-1 border border-transparent rounded"><AlignCenter size={14} className="rotate-90"/></button>
                 <button className="p-1 border border-transparent rounded"><AlignRight size={14} className="rotate-90"/></button>
               </div>
               <div className="flex gap-1 px-1">
                 <button onClick={() => handleFormatChange('align', 'left')} className={`p-1 hover:bg-[#d2e0cc] border border-transparent rounded ${activeFormat.align==='left'?'bg-[#c3d5bb] border-gray-400':''}`} title="Align Left"><AlignLeft size={14} /></button>
                 <button onClick={() => handleFormatChange('align', 'center')} className={`p-1 hover:bg-[#d2e0cc] border border-transparent rounded ${activeFormat.align==='center'?'bg-[#c3d5bb] border-gray-400':''}`} title="Center"><AlignCenter size={14} /></button>
                 <button onClick={() => handleFormatChange('align', 'right')} className={`p-1 hover:bg-[#d2e0cc] border border-transparent rounded ${activeFormat.align==='right'?'bg-[#c3d5bb] border-gray-400':''}`} title="Align Right"><AlignRight size={14} /></button>
               </div>
             </div>
             <div className="text-[11px] text-gray-400 text-center mt-auto w-full font-medium tracking-tight">Alignment</div>
         </div>

         {/* Data/Export Group */}
         <div className="flex flex-col pr-6 h-full border-r border-gray-300">
            <div className="grid grid-cols-2 gap-x-2 gap-y-2 mt-1">
               <label className="flex items-center gap-2 cursor-pointer hover:bg-gray-100 border border-gray-300 px-3 py-1 rounded text-gray-800 text-xs font-semibold bg-gray-50 w-24 justify-center shadow-sm transition-colors">
                  <Upload size={14} className="text-[#107c41]" /> Import
                  <input type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={importFile} />
               </label>
               <button onClick={() => exportExport('xlsx')} className="flex items-center gap-2 hover:bg-[#d2e0cc] hover:border-[#107c41] hover:text-[#107c41] border border-gray-300 px-3 py-1 rounded text-gray-800 text-xs font-semibold bg-white w-24 justify-center shadow-sm transition-all">
                  <Download size={14} className="text-[#107c41]" /> Export
               </button>
               <button onClick={printPdf} className="flex items-center gap-2 hover:bg-[#d2e0cc] hover:border-[#107c41] hover:text-[#107c41] border border-gray-300 px-3 py-1 rounded text-gray-800 text-xs font-semibold bg-white w-full justify-center shadow-sm transition-all col-span-2">
                  <Printer size={14} className="text-[#107c41]" /> Print PDF
               </button>
            </div>
            <div className="text-[11px] text-gray-400 text-center mt-auto w-full font-medium tracking-tight">I/O</div>
         </div>
         
         <div className="flex-1"></div>
         
         {/* Search Area */}
         <div className="flex flex-col mt-1 pr-6 relative w-48">
            <div className="flex items-center border border-gray-300 rounded overflow-hidden shadow-inner bg-white">
              <div className="px-2 text-gray-400 bg-gray-50 border-r border-gray-200 py-1.5"><Search size={12}/></div>
              <input 
                className="flex-1 text-xs px-2 outline-none py-1.5 font-medium placeholder:text-gray-400 placeholder:font-normal text-gray-800" 
                placeholder="Find in sheet... (Enter to search)" 
                value={searchTerm} 
                onChange={e => setSearchTerm(e.target.value)} 
                onKeyDown={e => e.key === 'Enter' && executeSearch()}
              />
            </div>
         </div>
      </div>

      {/* Formula Bar */}
      <div className="flex items-center gap-2 bg-white h-[32px] shrink-0 border-b border-[#c0c0c0] shadow-sm z-20 px-2 py-1 select-none">
         <div className="w-[85px] bg-white border border-[#c0c0c0] h-full flex items-center justify-center font-bold text-gray-700 text-xs shadow-inner shrink-0" style={{fontFamily: 'Calibri, sans-serif'}}>
            {activeCell ? `${colToAlpha(activeCell.c)}${activeCell.r + 1}` : ''}
         </div>
         <div className="text-gray-400 font-serif italic text-lg px-2 shrink-0 border-r border-[#e6e6e6] pr-4 flex items-center pb-1 shadow-sm">ƒx</div>
         <input 
            className="flex-1 h-full bg-white border border-transparent focus:border-[#107c41] outline-none px-2 text-[13px] font-sans -ml-1 placeholder:text-gray-300 transition-colors"
            placeholder={editMode ? "Enter value or formula..." : ""}
            value={formulaBarValue}
            onChange={(e) => {
               if (editMode) {
                  setEditValue(e.target.value);
               } else if (activeCell) {
                  startEdit(activeCell.r, activeCell.c);
                  setEditValue(e.target.value);
               }
            }}
            onKeyDown={(e) => { if(e.key === 'Enter') commitEdit(); }}
         />
         {editMode && (
            <div className="flex gap-1 ml-2 border-l pl-2 border-gray-200">
               <button onClick={() => { setEditMode(false); setEditValue(""); }} className="text-red-500 hover:bg-red-50 p-1 rounded-sm"><Trash2 size={16}/></button>
               <button onClick={commitEdit} className="text-[#107c41] hover:bg-[#d2e0cc] p-1 rounded-sm"><Check size={16}/></button>
            </div>
         )}
      </div>

      {/* Grid Container */}
      <div className="flex-1 min-h-0 overflow-auto relative bg-[#e6e6e6] print:bg-white print:overflow-visible custom-scrollbar">
         <table className="border-collapse table-fixed bg-white m-0 outline-none print:m-0 print:border-none print:shadow-none" style={{ minWidth: 'max-content' }}>
            <thead>
               <tr style={{ height: Math.max(22, 22 * zoom) }} className="select-none">
                  {/* Top Left Corner */}
                  <th className="w-[42px] bg-[#f3f2f1] border-r border-b border-[#c0c0c0] sticky top-0 left-0 z-30 print:hidden shadow-[inset_-1px_-1px_0_0_#9ca3af]"></th>
                  {/* Column Headers */}
                  {Array.from({length: currentSheet.cols}).map((_, c) => {
                    const isColActive = activeCell?.c === c && !editMode;
                    return (
                      <th 
                        key={c} 
                        className={`bg-[#f3f2f1] border-r border-b border-[#c0c0c0] font-normal text-gray-700 sticky top-0 z-20 print:hidden text-[11px] shadow-[inset_-1px_-1px_0_0_#c0c0c0] relative hover:bg-[#e6e6e6]
                          ${isColActive ? 'bg-[#d2e0cc] text-[#107c41] border-b-[#107c41] font-semibold' : ''}
                        `}
                        style={{ width: (colWidths[c] || 110) * zoom }}
                      >
                        {colToAlpha(c)}
                        <div onMouseDown={(e) => startResizing(e, c)} className="absolute right-0 top-0 bottom-0 w-2 cursor-col-resize hover:bg-[#107c41] opacity-50 z-20"></div>
                      </th>
                    );
                  })}
               </tr>
            </thead>
            <tbody>
               {Array.from({length: currentSheet.rows}).map((_, r) => (
                 <tr key={r} style={{ height: Math.max(20, 20 * zoom) }}>
                    {/* Row Headers */}
                    <td className={`w-[42px] bg-[#f3f2f1] border-r border-b border-[#c0c0c0] text-center text-gray-700 sticky left-0 z-20 print:hidden text-[11px] font-normal shadow-[inset_-1px_-1px_0_0_#c0c0c0] select-none
                       ${activeCell?.r === r && !editMode ? 'bg-[#d2e0cc] text-[#107c41] border-r-[#107c41] font-semibold' : ''}
                    `}>
                      {r+1}
                    </td>
                    {/* Main Cells */}
                    {Array.from({length: currentSheet.cols}).map((_, c) => {
                       const cellKey = `${r}_${c}`;
                       if (mergeInfo.hidden.has(cellKey)) return null;

                       const isActive = activeCell?.r === r && activeCell?.c === c;
                       const fKey = `${activeSheetId}_${r}_${c}`;
                       const f = formats[fKey] || {};
                       const merge = mergeInfo.lookup[cellKey];
                       
                       let val = hf.current.getCellValue({ sheet: activeSheetId, row: r, col: c });
                       if (val && typeof val === 'object' && val.value !== undefined) val = val.value;
                       const displayValue = val !== null && val !== undefined ? String(val) : "";

                       const cellFontFamily = f.fontFamily || 'Calibri, sans-serif';
                       const cellFontSize = Math.max(8, (f.fontSize || 13) * zoom);

                       return (
                         <td 
                           key={c}
                           rowSpan={merge?.rowspan}
                           colSpan={merge?.colspan}
                           onClick={(e) => { e.stopPropagation(); !editMode && setActiveCell({r,c}); }}
                           onDoubleClick={() => startEdit(r,c)}
                           onContextMenu={(e) => handleContextMenu(e, r, c)}
                           className={`relative border-r border-b border-[#d4d4d4] px-1 cursor-cell truncate print:border-gray-500 print:break-inside-avoid
                             ${isActive && !editMode ? 'outline outline-[2px] -outline-offset-[2px] outline-[#107c41] z-[5]' : ''}
                           `}
                           style={{
                             width: (colWidths[c] || 110) * zoom,
                             fontWeight: f.bold ? 'bold' : 'normal',
                             fontStyle: f.italic ? 'italic' : 'normal',
                             textDecoration: f.underline ? 'underline' : 'none',
                             textAlign: f.align || 'left',
                             color: f.color || '#000000',
                             backgroundColor: f.bg || '#ffffff',
                             fontFamily: cellFontFamily,
                             fontSize: cellFontSize + 'px',
                             verticalAlign: 'bottom'
                           }}
                         >
                           {isActive && editMode ? (
                             <input
                               ref={inputRef}
                               className="absolute inset-0 w-full h-full outline outline-[2px] -outline-offset-[2px] outline-[#107c41] px-1 m-0 block z-10 box-border border-none bg-white/90"
                               style={{
                                    fontWeight: f.bold ? 'bold' : 'normal',
                                    fontStyle: f.italic ? 'italic' : 'normal',
                                    textDecoration: f.underline ? 'underline' : 'none',
                                    textAlign: f.align || 'left',
                                    color: f.color || '#000000',
                                    backgroundColor: f.bg || '#ffffff',
                                    fontFamily: cellFontFamily,
                                    fontSize: cellFontSize + 'px',
                               }}
                               value={editValue}
                               onChange={e => setEditValue(e.target.value)}
                               onKeyDown={handleInputKeyDown}
                               onBlur={commitEdit}
                             />
                           ) : (
                              <div className="overflow-hidden pointer-events-none truncate select-none h-full" style={{lineHeight: Math.max(20, 20 * zoom) + 'px'}}>{displayValue}</div>
                           )}
                           {isActive && !editMode && (
                             <div className="absolute -bottom-[3px] -right-[3px] w-[6px] h-[6px] bg-[#107c41] cursor-crosshair z-10 border border-white"></div>
                           )}
                         </td>
                       )
                    })}
                 </tr>
               ))}
            </tbody>
         </table>
      </div>

      {/* Footer System: Tabs & Status */}
      <div className="flex flex-col shrink-0 bg-[#f3f2f1] border-t border-[#c0c0c0] print:hidden shadow-[0_-1px_5px_rgba(0,0,0,0.05)] z-30 select-none">
         {/* Sheet Tabs Scroll Area */}
         <div className="flex items-center h-[30px] bg-[#f3f2f1] overflow-hidden">
            <div className="flex items-center overflow-x-auto no-scrollbar scroll-smooth h-full">
               {sheets.map(s => (
                  <div 
                     key={s.id}
                     onClick={() => { commitEdit(); setActiveSheetId(s.id); }}
                     onDoubleClick={() => {
                        const newName = prompt("Rename Sheet", s.name);
                        if (newName) {
                           setSheets(prev => prev.map(sh => sh.id === s.id ? { ...sh, name: newName } : sh));
                           setIsDirty(true);
                           pushHistory();
                        }
                     }}
                     className={`px-4 h-full cursor-pointer min-w-max text-[12px] flex items-center justify-center gap-1 group/tab relative border-r border-[#c0c0c0] font-sans
                       ${activeSheetId === s.id ? 'bg-white font-bold text-[#107c41] shadow-[0_3px_0_0_#107c41_inset] pt-[1px]' : 'text-gray-600 hover:bg-[#e6e6e6]'}`}
                  >
                     {s.name}
                     {sheets.length > 1 && (
                        <button 
                           onClick={(e) => { e.stopPropagation(); removeSheet(s.id); }} 
                           className="opacity-0 group-hover/tab:opacity-100 hover:bg-gray-200 rounded-full p-0.5 ml-1 transition-opacity text-gray-500"
                           title="Delete sheet"
                        >
                           <Trash2 size={10} />
                        </button>
                     )}
                  </div>
               ))}
               <button onClick={addSheet} className="mx-2 p-1 hover:bg-[#d2e0cc] rounded-full text-[#107c41] transition-colors"><Plus size={14}/></button>
            </div>
         </div>
         {/* Status Bar */}
         <div className="flex items-center h-[24px] bg-[#107c41] text-white px-4 shadow-[inset_0_1px_0_rgba(255,255,255,0.1)] justify-between">
            <div className="text-[11px] font-medium tracking-wide">READY</div>
            <div className="flex items-center gap-2 h-full">
               <button onClick={() => setZoom(prev => Math.max(0.2, prev - 0.1))} className="hover:bg-[#0c5e31] p-0.5 rounded leading-none border border-transparent hover:border-white/20 transition-all">-</button>
               <input 
                 type="range" min="0.2" max="2.0" step="0.1" value={zoom} 
                 onChange={(e) => setZoom(parseFloat(e.target.value))}
                 className="w-24 h-1 bg-green-900 rounded-sm appearance-none cursor-pointer accent-white"
               />
               <button onClick={() => setZoom(prev => Math.min(2.0, prev + 0.1))} className="hover:bg-[#0c5e31] p-0.5 rounded leading-none border border-transparent hover:border-white/20 transition-all">+</button>
               <div className="text-[11px] font-medium w-10 text-right pr-2 select-none border-l border-white/20 pl-2 h-[14px] flex items-center justify-end">{Math.round(zoom * 100)}%</div>
            </div>
         </div>
      </div>
      {/* Context Menu Dialog */}
      {contextMenu && (
        <div 
          className="fixed bg-white border border-gray-300 shadow-xl rounded py-1 z-[100] min-w-[160px] text-gray-700"
          style={{ left: contextMenu.x, top: contextMenu.y }}
          onClick={(e) => e.stopPropagation()}
        >
          <button className="w-full text-left px-4 py-1.5 hover:bg-green-50 hover:text-green-700 flex items-center gap-2" onClick={() => gridAction('insert-row')}><Plus size={14}/> Insert Row Above</button>
          <button className="w-full text-left px-4 py-1.5 hover:bg-green-50 hover:text-green-700 flex items-center gap-2 border-b border-gray-100 mb-1" onClick={() => gridAction('delete-row')}><Trash2 size={14}/> Delete Row</button>
          
          <button className="w-full text-left px-4 py-1.5 hover:bg-green-50 hover:text-green-700 flex items-center gap-2" onClick={() => gridAction('insert-col')}><Plus size={14}/> Insert Column Left</button>
          <button className="w-full text-left px-4 py-1.5 hover:bg-green-50 hover:text-green-700 flex items-center gap-2" onClick={() => gridAction('delete-col')}><Trash2 size={14}/> Delete Column</button>
        </div>
      )}

      <style>{`
        .custom-scrollbar::-webkit-scrollbar { width: 14px; height: 14px; }
        .custom-scrollbar::-webkit-scrollbar-track { background: #f1f1f1; }
        .custom-scrollbar::-webkit-scrollbar-thumb { 
          background: #ccc; 
          border: 3px solid #f1f1f1; 
          border-radius: 10px;
        }
        .custom-scrollbar::-webkit-scrollbar-thumb:hover { background: #aaa; }
        
        @media print {
          body { background: white !important; margin: 0; padding: 0; }
          .print\\:hidden { display: none !important; }
          .flex-1 { overflow: visible !important; height: auto !important; position: static !important; padding: 0 !important; margin: 0 !important; }
          table { border-collapse: collapse !important; width: 100% !important; margin: 0 !important; border: 1px solid #000 !important; shadow: none !important; ring: 0 !important; }
          th, td { border: 1px solid #000 !important; padding: 4px !important; }
          
          /* Hide grid headers for clean report */
          thead, .sticky { display: none !important; }
          tbody td:first-child { display: none !important; }
          
          * { -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; }
        }
      `}</style>
    </div>
  );
}

