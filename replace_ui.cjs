const fs = require('fs');
const content = fs.readFileSync('src/ExcelEditor.jsx', 'utf8');

const returnRegex = /  return \([\s\S]*?\n  }\);\n\n  return \(/;
// Since there's only one top-level return in the component, we can split safely by looking for `  return (`
const lines = content.split('\n');

let startIndex = -1;
let endIndex = -1;

for (let i = 0; i < lines.length; i++) {
  if (lines[i].trim() === 'return (' && startIndex === -1) {
    startIndex = i;
  }
  if (lines[i].includes('{/* Context Menu Dialog */}') && startIndex !== -1) {
    endIndex = i - 2; // previous </div> line
    break;
  }
}

if (startIndex !== -1 && endIndex !== -1) {
  const newReturn = `  return (
    <div className="flex flex-col h-screen w-screen bg-[#f3f2f1] font-sans text-[13px] overflow-hidden" onClick={() => setContextMenu(null)}>
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
            <span className="font-semibold text-[14px] tracking-tight ml-2">Book1 - Excel Clone</span>
         </div>
         {lastSaved && <div className="text-[11px] opacity-90 font-medium">Saved: {lastSaved}</div>}
      </div>

      {/* Ribbon Tabs */}
      <div className="flex items-center bg-[#f3f2f1] pt-1.5 px-2 gap-1 shrink-0 border-b border-gray-300">
         <div className="px-4 py-1.5 text-white bg-[#107c41] font-semibold cursor-pointer rounded-t-sm hover:bg-[#0c5e31] transition-colors leading-none tracking-wide text-xs">File</div>
         <div className="px-4 py-1.5 text-[#107c41] bg-white border border-b-0 border-gray-300 font-bold cursor-pointer rounded-t-sm relative top-[1px] leading-none tracking-wide text-xs">Home</div>
         <div className="px-4 py-1.5 text-gray-700 font-medium cursor-pointer rounded-t-sm hover:bg-gray-200 transition-colors leading-none tracking-wide text-xs">Insert</div>
         <div className="px-4 py-1.5 text-gray-700 font-medium cursor-pointer rounded-t-sm hover:bg-gray-200 transition-colors leading-none tracking-wide text-xs">Formulas</div>
         <div className="px-4 py-1.5 text-gray-700 font-medium cursor-pointer rounded-t-sm hover:bg-gray-200 transition-colors leading-none tracking-wide text-xs">View</div>
      </div>

      {/* Ribbon Panel (Home) */}
      <div className="flex items-start bg-white h-[90px] shrink-0 border-b border-gray-300 shadow-sm z-30 px-3 py-2 gap-4 select-none">
         
         {/* Font Group */}
         <div className="flex flex-col border-r border-[#c0c0c0] pr-4 h-full relative group/ribbontool">
            <div className="flex gap-1 mb-1">
               <div className="relative group cursor-pointer" title="Font Family">
                 <select 
                    className="border border-[#c0c0c0] rounded px-1.5 h-6 text-xs w-[140px] focus:outline-none hover:bg-gray-100 hover:border-[#107c41] appearance-none"
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
                    className="border border-[#c0c0c0] rounded px-1.5 h-6 text-xs w-[52px] focus:outline-none hover:bg-gray-100 hover:border-[#107c41] appearance-none text-center"
                    value={activeFormat.fontSize || 13}
                    onChange={e => handleFormatChange('fontSize', parseInt(e.target.value))}
                 >
                    {[8,9,10,11,12,13,14,16,18,20,24,28,36,48,72].map(s => (
                       <option key={s} value={s}>{s}</option>
                    ))}
                 </select>
               </div>
               <button onClick={() => handleFormatChange('clear', true)} className="ml-1 px-1.5 h-6 text-gray-500 hover:bg-[#e6e6e6] border border-transparent hover:border-[#c0c0c0] rounded" title="Clear Formatting">
                  <Trash2 size={12}/>
               </button>
            </div>
            <div className="flex items-center gap-0.5 px-0.5">
               <button onClick={() => handleFormatChange('bold')} className={\`p-[3px] hover:bg-[#d2e0cc] border border-transparent rounded-sm \${activeFormat.bold?'bg-[#c3d5bb] border-[#c0c0c0] font-bold text-black':''}\`} title="Bold (Ctrl+B)"><Bold size={14} className="stroke-[3]"/></button>
               <button onClick={() => handleFormatChange('italic')} className={\`p-[3px] hover:bg-[#d2e0cc] border border-transparent rounded-sm \${activeFormat.italic?'bg-[#c3d5bb] border-[#c0c0c0] text-black':''}\`} title="Italic (Ctrl+I)"><Italic size={14}/></button>
               <button onClick={() => handleFormatChange('underline')} className={\`p-[3px] hover:bg-[#d2e0cc] border border-transparent rounded-sm \${activeFormat.underline?'bg-[#c3d5bb] border-[#c0c0c0] text-black':''}\`} title="Underline (Ctrl+U)"><Underline size={14}/></button>
               <div className="w-px h-[18px] bg-[#c0c0c0] mx-1.5 shrink-0"></div>
               <div className="relative group/pick">
                  <input type="color" className="absolute opacity-0 w-[24px] h-[24px] cursor-pointer z-10" value={activeFormat.bg || '#ffffff'} onChange={(e) => handleFormatChange('bg', e.target.value)} title="Fill Color"/>
                  <button className="p-0.5 px-[3px] hover:bg-[#d2e0cc] border border-transparent rounded text-gray-700 flex flex-col items-center">
                     <div className="w-[14px] h-[10px] border border-gray-400" style={{backgroundColor: activeFormat.bg || '#fff'}}></div>
                     <div className="w-[14px] h-[4px] mt-[1px] bg-yellow-400"></div>
                  </button>
               </div>
               <div className="relative group/pick">
                  <input type="color" className="absolute opacity-0 w-[24px] h-[24px] cursor-pointer z-10" value={activeFormat.color || '#000000'} onChange={(e) => handleFormatChange('color', e.target.value)} title="Font Color"/>
                  <button className="p-0.5 px-[3px] hover:bg-[#d2e0cc] border border-transparent rounded flex flex-col items-center text-red-600 font-bold" style={{lineHeight:'12px'}}>
                     <span className="text-[12px] font-serif" style={{color: activeFormat.color || '#000'}}>A</span>
                     <div className="w-[14px] h-[4px] mt-[2px]" style={{backgroundColor: activeFormat.color || '#000'}}></div>
                  </button>
               </div>
            </div>
            <div className="text-[11px] text-gray-500 text-center mt-auto w-full font-medium tracking-tight">Font</div>
         </div>

         {/* Alignment Group */}
         <div className="flex flex-col border-r border-[#c0c0c0] pr-4 h-full relative">
             <div className="flex flex-col gap-1 mt-0">
               <div className="flex gap-0.5 px-0.5 opacity-50 pointer-events-none" title="Vertical Alignment">
                 {/* Dummy vertical alignment buttons for Excel feel */}
                 <button className="p-[3px] border border-transparent rounded-sm"><AlignLeft size={14} className="rotate-90"/></button>
                 <button className="p-[3px] border border-transparent rounded-sm"><AlignCenter size={14} className="rotate-90"/></button>
                 <button className="p-[3px] border border-transparent rounded-sm"><AlignRight size={14} className="rotate-90"/></button>
               </div>
               <div className="flex gap-0.5 px-0.5">
                 <button onClick={() => handleFormatChange('align', 'left')} className={\`p-[3px] hover:bg-[#d2e0cc] border border-transparent rounded-sm \${activeFormat.align==='left'?'bg-[#c3d5bb] border-[#c0c0c0]':''}\`} title="Align Left"><AlignLeft size={14} /></button>
                 <button onClick={() => handleFormatChange('align', 'center')} className={\`p-[3px] hover:bg-[#d2e0cc] border border-transparent rounded-sm \${activeFormat.align==='center'?'bg-[#c3d5bb] border-[#c0c0c0]':''}\`} title="Center"><AlignCenter size={14} /></button>
                 <button onClick={() => handleFormatChange('align', 'right')} className={\`p-[3px] hover:bg-[#d2e0cc] border border-transparent rounded-sm \${activeFormat.align==='right'?'bg-[#c3d5bb] border-[#c0c0c0]':''}\`} title="Align Right"><AlignRight size={14} /></button>
               </div>
             </div>
             <div className="text-[11px] text-gray-500 text-center mt-auto w-full font-medium tracking-tight">Alignment</div>
         </div>

         {/* Data/Export Group */}
         <div className="flex flex-col pr-4 h-full">
            <div className="grid grid-cols-2 gap-x-2 gap-y-1 mt-0.5">
               <label className="flex items-center gap-1.5 cursor-pointer hover:bg-[#e6e6e6] border border-[#c0c0c0] px-2 py-1 rounded-sm text-gray-800 text-[11px] font-semibold bg-gray-50 w-[95px] justify-center shadow-sm">
                  <Upload size={12} className="text-[#107c41]" /> Import
                  <input type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={importFile} />
               </label>
               <button onClick={() => exportExport('xlsx')} className="flex items-center gap-1.5 hover:bg-[#d2e0cc] hover:border-[#107c41] hover:text-[#107c41] border border-[#c0c0c0] px-2 py-1 rounded-sm text-gray-800 text-[11px] font-semibold bg-white w-[95px] justify-center transition-all shadow-sm">
                  <Download size={12} className="text-[#107c41]" /> Export
               </button>
               <button onClick={printPdf} className="flex items-center gap-1.5 hover:bg-[#d2e0cc] hover:border-[#107c41] hover:text-[#107c41] border border-[#c0c0c0] px-2 py-1 rounded-sm text-gray-800 text-[11px] font-semibold bg-white w-[95px] justify-center transition-all shadow-sm col-span-2">
                  <Printer size={12} className="text-[#107c41]" /> Print PDF
               </button>
            </div>
            <div className="text-[11px] text-gray-500 text-center mt-auto w-full font-medium tracking-tight">I/O</div>
         </div>
         <div className="flex-1"></div>
         
         {/* Search Area */}
         <div className="flex flex-col mt-1 pr-6 relative w-48">
            <div className="flex items-center border border-gray-300 rounded overflow-hidden shadow-inner bg-white">
              <div className="px-2 text-gray-400 bg-gray-50 border-r border-gray-200 py-1.5"><Search size={12}/></div>
              <input 
                className="flex-1 text-xs px-2 outline-none py-1.5 font-medium placeholder:text-gray-400 placeholder:font-normal text-gray-800" 
                placeholder="Find in sheet..." 
                value={searchTerm} 
                onChange={e => setSearchTerm(e.target.value)} 
              />
            </div>
         </div>
      </div>

      {/* Formula Bar */}
      <div className="flex items-center gap-2 bg-white h-[32px] shrink-0 border-b border-[#c0c0c0] shadow-sm z-20 px-2 py-1 select-none">
         <div className="w-[85px] bg-white border border-[#c0c0c0] h-full flex items-center justify-center font-bold text-gray-700 text-xs shadow-inner shrink-0" style={{fontFamily: 'Calibri, sans-serif'}}>
            {activeCell ? \`\${colToAlpha(activeCell.c)}\${activeCell.r + 1}\` : ''}
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
                        className={\`bg-[#f3f2f1] border-r border-b border-[#c0c0c0] font-normal text-gray-700 sticky top-0 z-20 print:hidden text-[11px] shadow-[inset_-1px_-1px_0_0_#c0c0c0] relative hover:bg-[#e6e6e6]
                          \${isColActive ? 'bg-[#d2e0cc] text-[#107c41] border-b-[#107c41] font-semibold' : ''}
                        \`}
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
                    <td className={\`w-[42px] bg-[#f3f2f1] border-r border-b border-[#c0c0c0] text-center text-gray-700 sticky left-0 z-20 print:hidden text-[11px] font-normal shadow-[inset_-1px_-1px_0_0_#c0c0c0] select-none
                       \${activeCell?.r === r && !editMode ? 'bg-[#d2e0cc] text-[#107c41] border-r-[#107c41] font-semibold' : ''}
                    \`}>
                      {r+1}
                    </td>
                    {/* Main Cells */}
                    {Array.from({length: currentSheet.cols}).map((_, c) => {
                       const cellKey = \`\${r}_\${c}\`;
                       if (mergeInfo.hidden.has(cellKey)) return null;

                       const isActive = activeCell?.r === r && activeCell?.c === c;
                       const fKey = \`\${activeSheetId}_\${r}_\${c}\`;
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
                           className={\`relative border-r border-b border-[#d4d4d4] px-1 cursor-cell truncate print:border-gray-500 print:break-inside-avoid
                             \${isActive && !editMode ? 'outline outline-[2px] -outline-offset-[2px] outline-[#107c41] z-[5]' : ''}
                           \`}
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
                     className={\`px-4 h-full cursor-pointer min-w-max text-[12px] flex items-center justify-center gap-1 group/tab relative border-r border-[#c0c0c0] font-sans
                       \${activeSheetId === s.id ? 'bg-white font-bold text-[#107c41] shadow-[0_3px_0_0_#107c41_inset] pt-[1px]' : 'text-gray-600 hover:bg-[#e6e6e6]'}\`}
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
      </div>
  );`;
  
  lines.splice(startIndex, endIndex - startIndex + 1, newReturn);
  fs.writeFileSync('src/ExcelEditor.jsx', lines.join('\n'));
  console.log('Successfully replaced layout block');
} else {
  console.log('Could not find start or end index', startIndex, endIndex);
}
