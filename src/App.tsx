import React, { useState, useRef, useEffect } from 'react';
import { motion, AnimatePresence } from 'motion/react';
import { 
  Users, 
  UserPlus, 
  Shuffle, 
  Trash2, 
  Copy, 
  Check, 
  Settings2, 
  RotateCcw,
  LayoutGrid,
  List,
  FileUp,
  FileDown,
  Download,
  ChevronRight
} from 'lucide-react';
import confetti from 'canvas-confetti';
import * as XLSX from 'xlsx';

interface Person {
  name: string;
  employeeId: string;
}

interface GroupedData {
  id: number;
  members: Person[];
}

interface HistoryItem {
  id: string;
  timestamp: number;
  mode: 'count' | 'per';
  configValue: number;
  operator: string;
  groups: GroupedData[];
  totalCount: number;
}

export default function App() {
  const operators = ['顏小琳', '王小明', '李小華'];
  const [people, setPeople] = useState<Person[]>([]);
  const [manualInput, setManualInput] = useState('');
  const [groupCount, setGroupCount] = useState(2);
  const [peoplePerGroup, setPeoplePerGroup] = useState(4);
  const [mode, setMode] = useState<'count' | 'per'>('count');
  const [selectedOperator, setSelectedOperator] = useState(operators[0]);
  const [groups, setGroups] = useState<GroupedData[]>([]);
  const [history, setHistory] = useState<HistoryItem[]>(() => {
    const saved = localStorage.getItem('grouping_history');
    return saved ? JSON.parse(saved) : [];
  });
  const [isGrouping, setIsGrouping] = useState(false);
  const [copied, setCopied] = useState(false);
  const [activeTab, setActiveTab] = useState<'current' | 'history'>('current');
  
  const fileInputRef = useRef<HTMLInputElement>(null);

  // Sync history to localStorage
  useEffect(() => {
    localStorage.setItem('grouping_history', JSON.stringify(history));
  }, [history]);

  const totalNames = people.length;

  const handleGroup = () => {
    if (totalNames === 0) return;
    
    setIsGrouping(true);
    
    setTimeout(() => {
      const shuffled = [...people].sort(() => Math.random() - 0.5);
      const newGroups: GroupedData[] = [];
      
      let calculatedGroupCount = groupCount;
      if (mode === 'per') {
        calculatedGroupCount = Math.ceil(totalNames / peoplePerGroup);
      }

      for (let i = 0; i < (calculatedGroupCount || 1); i++) {
        newGroups.push({ id: i + 1, members: [] });
      }

      shuffled.forEach((person, index) => {
        newGroups[index % (calculatedGroupCount || 1)].members.push(person);
      });

      const filteredGroups = newGroups.filter(g => g.members.length > 0);
      setGroups(filteredGroups);
      
      // Save to history
      const newHistoryItem: HistoryItem = {
        id: crypto.randomUUID(),
        timestamp: Date.now(),
        mode: mode,
        configValue: mode === 'count' ? groupCount : peoplePerGroup,
        operator: selectedOperator,
        groups: filteredGroups,
        totalCount: totalNames
      };
      setHistory(prev => [newHistoryItem, ...prev].slice(0, 50)); // Keep last 50 items

      setIsGrouping(false);
      
      confetti({
        particleCount: 100,
        spread: 70,
        origin: { y: 0.6 },
        colors: ['#74d9b6', '#f06292', '#fbc02d', '#ba68c8', '#64b5f6']
      });
    }, 600);
  };

  const clearAll = () => {
    setPeople([]);
    setManualInput('');
    setGroups([]);
  };

  const handleManualInputAdd = () => {
    const lines = manualInput.split('\n').filter(l => l.trim() !== '');
    const newPeople: Person[] = lines.map(line => {
      // Split by comma, space, or tab
      const parts = line.split(/[, \t]+/).filter(p => p !== '');
      return {
        name: parts[0] || 'Unknown',
        employeeId: parts[1] || '---'
      };
    });
    setPeople(prev => [...prev, ...newPeople]);
    setManualInput('');
  };

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (event) => {
      const dataStr = event.target?.result;
      const workbook = XLSX.read(dataStr, { type: 'binary' });
      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];
      const data = XLSX.utils.sheet_to_json<any>(worksheet);

      const parsedPeople: Person[] = data.map(item => {
        // Try to find columns by common names
        const name = String(item['姓名'] || item['Name'] || Object.values(item)[0] || '');
        const id = String(item['員工編號'] || item['ID'] || item['Employee ID'] || Object.values(item)[1] || '');
        return { name, employeeId: id };
      }).filter(p => p.name !== '');

      setPeople(prev => [...prev, ...parsedPeople]);
    };
    reader.readAsBinaryString(file);
    if (fileInputRef.current) fileInputRef.current.value = '';
  };

  const downloadTemplate = () => {
    const ws = XLSX.utils.json_to_sheet([
      { '姓名': '王小明', '員工編號': 'A001' },
      { '姓名': '李小華', '員工編號': 'A002' },
      { '姓名': '張大衛', '員工編號': 'A003' }
    ]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Template');
    XLSX.writeFile(wb, '匯入範例.xlsx');
  };

  const exportResults = (format: 'xlsx' | 'csv' | 'txt') => {
    if (groups.length === 0) return;

    const exportData = groups.flatMap(group => 
      group.members.map(member => ({
        '姓名': member.name,
        '員工編號': member.employeeId,
        '組別': `第 ${group.id} 組`
      }))
    );

    if (format === 'xlsx') {
      const ws = XLSX.utils.json_to_sheet(exportData);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, '分組結果');
      XLSX.writeFile(wb, '分組結果.xlsx');
    } else if (format === 'csv') {
      const ws = XLSX.utils.json_to_sheet(exportData);
      const csv = XLSX.utils.sheet_to_csv(ws);
      const blob = new Blob(["\ufeff" + csv], { type: 'text/csv;charset=utf-8;' });
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.setAttribute('download', '分組結果.csv');
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
    } else if (format === 'txt') {
      const text = groups
        .map(g => `第 ${g.id} 組:\n${g.members.map(m => `${m.name} (${m.employeeId})`).join('\n')}`)
        .join('\n\n');
      const blob = new Blob([text], { type: 'text/plain;charset=utf-8;' });
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.setAttribute('download', '分組結果.txt');
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
    }
  };

  const copyTSVForSheets = async () => {
    if (groups.length === 0) return;

    // TSV (Tab Separated Values) is the best format for direct pasting into spreadsheet cells
    const header = "姓名\t員工編號\t組別\n";
    const body = groups.flatMap(group => 
      group.members.map(member => `${member.name}\t${member.employeeId}\t第 ${group.id} 組`)
    ).join('\n');
      
    try {
      await navigator.clipboard.writeText(header + body);
      setCopied(true);
      setTimeout(() => setCopied(false), 2000);
    } catch (err) {
      console.error('Failed to copy TSV!', err);
    }
  };

  const deleteHistoryItem = (id: string, e: React.MouseEvent) => {
    e.stopPropagation();
    setHistory(prev => prev.filter(item => item.id !== id));
  };

  const clearHistory = () => {
    if (confirm('確定要清除所有分組歷程嗎？')) {
      setHistory([]);
    }
  };

  const restoreFromHistory = (item: HistoryItem) => {
    setGroups(item.groups);
    setActiveTab('current');
    
    // Add a small feedback effect
    confetti({
      particleCount: 50,
      spread: 40,
      origin: { y: 0.8 },
      colors: ['#e3f2fd', '#74d9b6']
    });
  };

  const formatTime = (ts: number) => {
    return new Intl.DateTimeFormat('zh-TW', {
      month: 'numeric',
      day: 'numeric',
      hour: '2-digit',
      minute: '2-digit'
    }).format(ts);
  };

  const copyResults = async () => {
    if (groups.length === 0) return;
    
    const text = groups
      .map(g => `第 ${g.id} 組:\n${g.members.map(m => `${m.name} (${m.employeeId})`).join('\n')}`)
      .join('\n\n');
      
    try {
      await navigator.clipboard.writeText(text);
      setCopied(true);
      setTimeout(() => setCopied(false), 2000);
    } catch (err) {
      console.error('Failed to copy!', err);
    }
  };

  return (
    <div className="min-h-screen bg-[#fafafa] py-12 px-4 sm:px-6 lg:px-8">
      <div className="max-w-6xl mx-auto">
        {/* Header */}
        <header className="mb-12 text-center">
          <motion.div
            initial={{ opacity: 0, y: -20 }}
            animate={{ opacity: 1, y: 0 }}
            className="inline-flex items-center justify-center p-3 bg-macaron-mint-deep rounded-[2rem] shadow-lg shadow-macaron-mint/50 mb-6"
          >
            <Users className="w-8 h-8 text-white" />
          </motion.div>
          <motion.h1
            initial={{ opacity: 0, y: -10 }}
            animate={{ opacity: 1, y: 0 }}
            className="text-4xl font-bold font-display text-slate-900 tracking-tight"
          >
            隨機分組工具
          </motion.h1>
          <motion.p
            initial={{ opacity: 0, y: -5 }}
            animate={{ opacity: 1, y: 0 }}
            className="mt-3 text-lg text-slate-500"
          >
            支援 Excel 兩欄位（姓名、員編）匯入匯出
          </motion.p>
        </header>

        <div className="grid grid-cols-1 lg:grid-cols-12 gap-8 items-start">
          {/* Left Column */}
          <div className="lg:col-span-5 space-y-6">
            <section className="bg-white rounded-[2.5rem] p-8 shadow-sm border border-slate-100">
              <div className="flex items-center justify-between mb-6">
                <label className="flex items-center gap-2 font-black text-slate-800 text-lg uppercase tracking-tight">
                  <UserPlus className="w-6 h-6 text-macaron-blue-deep" />
                  成員清單 ({totalNames})
                </label>
                {totalNames > 0 && (
                  <button onClick={clearAll} className="p-2 hover:bg-red-50 text-red-400 rounded-xl transition-colors">
                    <Trash2 className="w-5 h-5" />
                  </button>
                )}
              </div>

              {/* Members List Scroll Area */}
              <div className="bg-slate-50/50 rounded-3xl p-4 mb-6 max-h-72 overflow-y-auto border-2 border-dashed border-slate-100">
                {people.length === 0 ? (
                  <div className="py-16 text-center text-slate-300 font-medium italic">
                    請匯入成員清單...
                  </div>
                ) : (
                  <div className="space-y-2">
                    {people.map((p, i) => (
                      <motion.div 
                        initial={{ opacity: 0, x: -10 }}
                        animate={{ opacity: 1, x: 0 }}
                        key={i} 
                        className="flex items-center justify-between px-4 py-3 bg-white rounded-2xl border border-slate-100 shadow-sm transition-all hover:border-macaron-blue"
                      >
                        <span className="font-bold text-slate-700">{p.name}</span>
                        <span className="text-[10px] font-black bg-slate-100 px-2 py-1 rounded-lg text-slate-400 uppercase tracking-tighter">ID: {p.employeeId}</span>
                      </motion.div>
                    ))}
                  </div>
                )}
              </div>

              <div className="space-y-3">
                <textarea
                  value={manualInput}
                  onChange={(e) => setManualInput(e.target.value)}
                  placeholder="手動輸入：姓名 員工編號"
                  className="w-full h-24 p-4 rounded-3xl bg-slate-50 border-none focus:ring-4 focus:ring-macaron-blue/20 transition-all text-sm resize-none"
                />
                <button
                  onClick={handleManualInputAdd}
                  disabled={!manualInput.trim()}
                  className="w-full py-3 bg-macaron-blue-deep/10 hover:bg-macaron-blue-deep/20 text-macaron-blue-deep rounded-2xl text-sm font-black transition-all disabled:opacity-30 disabled:cursor-not-allowed uppercase"
                >
                  添加成員
                </button>
              </div>
            </section>

            {/* Config Section */}
            <section className="bg-white rounded-[2.5rem] p-8 shadow-sm border border-slate-100">
              <div className="flex items-center gap-2 mb-8 font-black text-slate-800 text-lg uppercase tracking-tight">
                <Settings2 className="w-6 h-6 text-macaron-pink-deep" />
                分組參數
              </div>

              <div className="space-y-8">
                {/* Operator Selection */}
                <div className="space-y-3">
                  <span className="text-slate-500 font-extrabold text-sm uppercase tracking-widest block px-1">選擇執行人員</span>
                  <div className="grid grid-cols-3 gap-2">
                    {operators.map(op => (
                      <button
                        key={op}
                        onClick={() => setSelectedOperator(op)}
                        className={`py-2 px-1 rounded-xl text-[10px] font-black border-2 transition-all ${
                          selectedOperator === op 
                          ? 'border-macaron-pink-deep bg-macaron-pink text-macaron-pink-deep' 
                          : 'border-slate-100 bg-slate-50 text-slate-400 hover:border-slate-200'
                        }`}
                      >
                        {op}
                      </button>
                    ))}
                  </div>
                </div>

                <div className="flex p-1.5 bg-slate-100 rounded-3xl">
                  <button
                    onClick={() => setMode('count')}
                    className={`flex-1 flex items-center justify-center gap-2 py-4 rounded-[1.25rem] text-sm font-black transition-all ${
                      mode === 'count' ? 'bg-white text-macaron-blue-deep shadow-sm' : 'text-slate-400'
                    }`}
                  >
                    按組數
                  </button>
                  <button
                    onClick={() => setMode('per')}
                    className={`flex-1 flex items-center justify-center gap-2 py-4 rounded-[1.25rem] text-sm font-black transition-all ${
                      mode === 'per' ? 'bg-white text-macaron-blue-deep shadow-sm' : 'text-slate-400'
                    }`}
                  >
                    按人數
                  </button>
                </div>

                <div className="px-2">
                  <div className="flex justify-between items-end mb-4">
                    <span className="text-slate-500 font-extrabold text-sm uppercase tracking-widest">{mode === 'count' ? '分成幾組?' : '每組幾人?'}</span>
                    <span className="text-macaron-blue-deep text-5xl font-black font-display leading-none">
                      {mode === 'count' ? groupCount : peoplePerGroup}
                    </span>
                  </div>
                  <input
                    type="range"
                    min={mode === 'count' ? "2" : "1"}
                    max={Math.max(2, totalNames)}
                    value={mode === 'count' ? groupCount : peoplePerGroup}
                    onChange={(e) => mode === 'count' ? setGroupCount(Number(e.target.value)) : setPeoplePerGroup(Number(e.target.value))}
                    className="w-full h-3 bg-slate-100 rounded-full appearance-none cursor-pointer accent-macaron-blue-deep"
                  />
                </div>

                <button
                  onClick={handleGroup}
                  disabled={totalNames === 0 || isGrouping}
                  className="w-full py-6 bg-macaron-pink-deep hover:bg-macaron-pink-deep/90 disabled:bg-slate-100 disabled:text-slate-300 text-white rounded-[2rem] font-black text-lg flex items-center justify-center gap-4 transition-all shadow-xl shadow-macaron-pink/40 active:scale-95 uppercase tracking-wide"
                >
                  {isGrouping ? (
                    <motion.div animate={{ rotate: 360 }} transition={{ duration: 1, repeat: Infinity, ease: "linear" }}>
                      <RotateCcw className="w-6 h-6" />
                    </motion.div>
                  ) : (
                    <>
                      <Shuffle className="w-6 h-6" />
                      立即分組
                    </>
                  )}
                </button>
              </div>
            </section>
          </div>

          {/* Right Column */}
          <div className="lg:col-span-7">
            <div className="bg-white rounded-[3rem] p-1 shadow-sm border border-slate-100 flex flex-col h-full min-h-[720px]">
              {/* Tab Header */}
              <div className="flex p-4 gap-2 mb-2 bg-slate-50/50 rounded-t-[2.8rem] border-b border-slate-100">
                <button
                  onClick={() => setActiveTab('current')}
                  className={`flex-1 py-3 rounded-2xl flex items-center justify-center gap-2 font-black text-xs uppercase tracking-widest transition-all ${
                    activeTab === 'current' 
                    ? 'bg-white text-macaron-blue-deep shadow-sm' 
                    : 'text-slate-400 hover:text-slate-600'
                  }`}
                >
                  <Shuffle className="w-4 h-4" />
                  本次結果
                </button>
                <button
                  onClick={() => setActiveTab('history')}
                  className={`flex-1 py-3 rounded-2xl flex items-center justify-center gap-2 font-black text-xs uppercase tracking-widest transition-all ${
                    activeTab === 'history' 
                    ? 'bg-white text-macaron-lavender-deep shadow-sm' 
                    : 'text-slate-400 hover:text-slate-600'
                  }`}
                >
                  <RotateCcw className="w-4 h-4" />
                  分組歷程 {history.length > 0 && `(${history.length})`}
                </button>
              </div>

              <div className="p-8 flex-1 flex flex-col">
                <AnimatePresence mode="wait">
                  {activeTab === 'current' ? (
                    <motion.div
                      key="current-tab"
                      initial={{ opacity: 0, x: -10 }}
                      animate={{ opacity: 1, x: 0 }}
                      exit={{ opacity: 0, x: 10 }}
                      className="flex-1 flex flex-col"
                    >
                      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-6 mb-10">
                        <h2 className="text-3xl font-black font-display text-slate-800 tracking-tight flex items-center gap-4">
                          分組結果
                          {groups.length > 0 && (
                            <span className="bg-macaron-blue text-macaron-blue-deep text-xs py-1.5 px-4 rounded-full font-black uppercase tracking-widest border-2 border-white shadow-sm">
                              {groups.length} Groups
                            </span>
                          )}
                        </h2>
                        
                        {groups.length > 0 && (
                          <div className="flex flex-wrap items-center gap-3">
                            <button
                              onClick={copyResults}
                              className="flex items-center gap-2 text-[10px] font-black px-4 py-3 bg-slate-50 hover:bg-macaron-blue-deep/10 transition-all rounded-[1rem] border border-slate-100 text-slate-500 hover:text-macaron-blue-deep uppercase"
                            >
                              {copied ? <Check className="w-3.5 h-3.5 text-green-500" /> : <Copy className="w-3.5 h-3.5" />}
                              複製文字
                            </button>
                            
                            <div className="h-6 w-px bg-slate-100 mx-1 hidden sm:block" />
                            
                            <div className="flex items-center bg-slate-50 rounded-[1rem] p-1 border border-slate-100">
                              <button
                                onClick={() => exportResults('txt')}
                                className="px-3 py-2 text-[10px] font-black text-slate-400 hover:text-slate-600 transition-colors uppercase"
                              >
                                TXT
                              </button>
                              <button
                                onClick={() => exportResults('csv')}
                                className="px-3 py-2 text-[10px] font-black text-slate-400 hover:text-slate-600 transition-colors uppercase"
                              >
                                CSV
                              </button>
                              <button
                                onClick={() => exportResults('xlsx')}
                                className="flex items-center gap-2 px-4 py-2 bg-white text-macaron-mint-deep text-[10px] font-black rounded-[0.75rem] shadow-sm border border-slate-100 hover:bg-macaron-mint transition-all uppercase"
                              >
                                <FileDown className="w-3 h-3" />
                                Excel 匯出
                              </button>
                            </div>

                            <button
                              onClick={async () => {
                                await copyTSVForSheets();
                                window.open('https://sheets.new', '_blank');
                              }}
                              className="flex items-center gap-2 text-[10px] font-black px-5 py-3 bg-[#0F9D58] text-white hover:scale-105 transition-all rounded-[1rem] border-2 border-white shadow-lg shadow-green-200 uppercase"
                            >
                              <LayoutGrid className="w-3.5 h-3.5" />
                              Google 試算表
                            </button>
                          </div>
                        )}
                      </div>

                      <div className="flex-1">
                        {groups.length > 0 ? (
                          <div className="grid grid-cols-1 sm:grid-cols-2 gap-6">
                            {groups.map((group, index) => (
                              <motion.div
                                key={group.id}
                                initial={{ opacity: 0, scale: 0.9, y: 20 }}
                                animate={{ opacity: 1, scale: 1, y: 0 }}
                                transition={{ delay: index * 0.05 }}
                                className={`p-6 rounded-[2.5rem] border-4 border-white shadow-lg transition-all ${
                                  [
                                    'bg-macaron-blue',
                                    'bg-macaron-pink',
                                    'bg-macaron-mint',
                                    'bg-macaron-yellow',
                                    'bg-macaron-lavender'
                                  ][index % 5]
                                }`}
                              >
                                <div className="flex items-center justify-between mb-5">
                                  <span className="text-[10px] font-black text-slate-700/50 uppercase tracking-[0.2em] italic">Team Group 0{group.id}</span>
                                  <span className="bg-white/50 px-3 py-1 rounded-full text-[10px] text-slate-700 font-black">{group.members.length} MEMBERS</span>
                                </div>
                                <ul className="space-y-2.5">
                                  {group.members.map((member, mIdx) => (
                                    <motion.li
                                      key={mIdx}
                                      initial={{ opacity: 0, x: -10 }}
                                      animate={{ opacity: 1, x: 0 }}
                                      transition={{ delay: index * 0.05 + mIdx * 0.03 }}
                                      className="flex items-center justify-between bg-white/40 backdrop-blur-md p-4 rounded-2xl shadow-sm border border-white/50"
                                    >
                                      <span className="font-black text-slate-800 text-sm">{member.name}</span>
                                      <span className="text-[9px] font-bold text-slate-500/80 uppercase tracking-widest">{member.employeeId}</span>
                                    </motion.li>
                                  ))}
                                </ul>
                              </motion.div>
                            ))}
                          </div>
                        ) : (
                          <div className="h-full flex flex-col items-center justify-center text-center opacity-40 grayscale">
                            <div className="w-32 h-32 bg-slate-100 rounded-[2.5rem] flex items-center justify-center mb-8 rotate-12">
                              <Users className="w-12 h-12 text-slate-400" />
                            </div>
                            <p className="font-black text-2xl text-slate-400 font-display uppercase tracking-widest">No Active Groups</p>
                          </div>
                        )}
                      </div>
                    </motion.div>
                  ) : (
                    <motion.div
                      key="history-tab"
                      initial={{ opacity: 0, x: 10 }}
                      animate={{ opacity: 1, x: 0 }}
                      exit={{ opacity: 0, x: -10 }}
                      className="flex-1 flex flex-col"
                    >
                      <div className="flex items-center justify-between mb-8">
                        <h2 className="text-3xl font-black font-display text-slate-800 tracking-tight">分組歷程</h2>
                        {history.length > 0 && (
                          <button
                            onClick={clearHistory}
                            className="text-[10px] font-black px-4 py-2 bg-red-50 text-red-400 rounded-xl hover:bg-red-100 transition-all uppercase tracking-widest"
                          >
                            清除全部歷程
                          </button>
                        )}
                      </div>

                      <div className="flex-1 space-y-4">
                        {history.length > 0 ? (
                          history.map((item) => (
                            <motion.div
                              key={item.id}
                              onClick={() => restoreFromHistory(item)}
                              className="group p-5 bg-slate-50 rounded-[2rem] border border-slate-100 flex items-center gap-6 cursor-pointer hover:border-macaron-lavender hover:bg-macaron-lavender/10 transition-all"
                            >
                              <div className="w-12 h-12 rounded-2xl bg-white flex items-center justify-center shadow-sm text-macaron-lavender-deep font-black">
                                {item.groups.length}
                              </div>
                              <div className="flex-1">
                                <div className="text-sm font-black text-slate-700 mb-0.5">
                                  {item.totalCount} 位成員 • {item.groups.length} 組
                                </div>
                                <div className="text-[10px] text-slate-400 font-bold uppercase tracking-wide flex flex-wrap gap-2">
                                  <span>{formatTime(item.timestamp)}</span>
                                  <span className="text-macaron-pink-deep font-black">• 執行人員: {item.operator}</span>
                                  <span>• {item.mode === 'count' ? '按組數' : '按人數'} ({item.configValue})</span>
                                </div>
                              </div>
                              <div className="flex items-center gap-2 opacity-0 group-hover:opacity-100 transition-all scale-90 group-hover:scale-100">
                                <button
                                  onClick={(e) => deleteHistoryItem(item.id, e)}
                                  className="p-2 text-slate-300 hover:text-red-400 transition-colors"
                                >
                                  <Trash2 className="w-4 h-4" />
                                </button>
                                <div className="w-8 h-8 rounded-full bg-macaron-lavender-deep text-white flex items-center justify-center">
                                  <ChevronRight className="w-4 h-4" />
                                </div>
                              </div>
                            </motion.div>
                          ))
                        ) : (
                          <div className="h-full flex flex-col items-center justify-center text-center py-20 opacity-30">
                            <RotateCcw className="w-16 h-16 text-slate-300 mb-6" />
                            <p className="font-black text-slate-400 uppercase tracking-widest">目前尚無歷史紀錄</p>
                          </div>
                        )}
                      </div>
                    </motion.div>
                  )}
                </AnimatePresence>
              </div>
            </div>
          </div>
        </div>
      </div>

      {/* Action Control Panel */}
      <div className="fixed bottom-10 left-10 z-50 flex flex-col gap-4">
        <input
          type="file"
          ref={fileInputRef}
          onChange={handleFileUpload}
          accept=".xlsx,.xls,.csv"
          className="hidden"
        />
        
        <motion.button
          whileHover={{ scale: 1.05, x: 10 }}
          whileTap={{ scale: 0.95 }}
          onClick={downloadTemplate}
          className="flex items-center gap-3 bg-white text-slate-500 px-6 py-4 rounded-[1.5rem] shadow-xl border border-slate-100 text-xs font-black uppercase tracking-widest hover:bg-macaron-yellow-deep/10 group transition-all"
        >
          <div className="w-8 h-8 rounded-xl bg-macaron-yellow-deep/20 flex items-center justify-center group-hover:bg-macaron-yellow-deep/50 transition-colors">
            <Download className="w-4 h-4 text-macaron-yellow-deep" />
          </div>
          下載 EXCEL 範例 (姓名/員編)
        </motion.button>

        <motion.button
          whileHover={{ scale: 1.05, x: 10 }}
          whileTap={{ scale: 0.95 }}
          onClick={() => fileInputRef.current?.click()}
          className="flex items-center gap-4 bg-slate-900 text-white px-8 py-5 rounded-[2rem] shadow-[0_20px_50px_rgba(0,0,0,0.15)] font-black uppercase tracking-widest hover:bg-macaron-mint-deep transition-all group"
        >
          <div className="w-12 h-12 rounded-[1.25rem] bg-white/10 flex items-center justify-center group-hover:bg-white/20 transition-all">
            <FileUp className="w-6 h-6 text-white" />
          </div>
          匯入成員 EXCEL
        </motion.button>
      </div>

      <footer className="mt-32 text-center text-slate-300 text-[10px] font-black uppercase tracking-[0.4em] pb-12">
        <p>Random Group Pro v2.0 • Data Security First • Happy Coding</p>
      </footer>
    </div>
  );
}
