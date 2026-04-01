import React, { useState, useMemo } from 'react';
import { Upload, FileSpreadsheet, Users, AlertTriangle, Cloud, BarChart3, Search, Download, CheckCircle2 } from 'lucide-react';
import { cn } from './lib/utils';
import * as XLSX from 'xlsx';
import WordCloud from './components/WordCloud';
import { BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer, PieChart, Pie, Cell } from 'recharts';

// Types
interface Student {
  id: string;
  grade: string;
  class: string;
  number: string;
  name: string;
  score?: number;
  riskLevel?: string;
  descriptive?: string;
  category?: string;
  responses?: Record<string, any>;
  uniqueCode?: string;
}

const COLORS = ['#10b981', '#f59e0b', '#ef4444', '#8b5cf6'];

export default function App() {
  const [students, setStudents] = useState<Student[]>([]);
  const [rosterLoaded, setRosterLoaded] = useState(false);
  const [resultsLoaded, setResultsLoaded] = useState(false);
  const [activeTab, setActiveTab] = useState<'overview' | 'students' | 'wordcloud'>('overview');
  const [searchTerm, setSearchTerm] = useState('');

  // --- Excel Parsing Logic ---
  const handleRosterUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (event) => {
      try {
        const data = new Uint8Array(event.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        
        let parsedStudents: Student[] = [];
        const existingCodes = new Set<string>();
        
        const generateUniqueCode = () => {
          let code;
          do {
            // Generate 6-character alphanumeric code
            code = Math.random().toString(36).substring(2, 8).toUpperCase();
          } while (existingCodes.has(code));
          existingCodes.add(code);
          return code;
        };
        
        workbook.SheetNames.forEach(sheetName => {
          const sheet = workbook.Sheets[sheetName];
          const json = XLSX.utils.sheet_to_json<any>(sheet);
          
          json.forEach(row => {
            const name = row['이름'] || row['성명'] || row['Name'] || row['name'];
            const grade = row['학년'] || row['Grade'] || '';
            const cls = row['반'] || row['Class'] || sheetName.replace(/[^0-9]/g, '');
            const num = row['번호'] || row['Number'] || '';
            const id = row['학번'] || row['ID'] || `${grade}${cls.toString().padStart(2, '0')}${num.toString().padStart(2, '0')}`;
            
            if (name) {
              parsedStudents.push({
                id: id.toString(),
                grade: grade.toString(),
                class: cls.toString(),
                number: num.toString(),
                name: name.toString(),
                uniqueCode: generateUniqueCode()
              });
            }
          });
        });
        
        setStudents(parsedStudents);
        setRosterLoaded(true);
      } catch (err) {
        console.error("Failed to parse roster:", err);
        alert("명렬표 파일을 읽는 중 오류가 발생했습니다.");
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const handleResultsUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file || students.length === 0) {
      alert("먼저 명렬표를 업로드해주세요.");
      return;
    }

    const reader = new FileReader();
    reader.onload = (event) => {
      try {
        const data = new Uint8Array(event.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        
        let updatedStudents = [...students];
        
        workbook.SheetNames.forEach(sheetName => {
          const sheet = workbook.Sheets[sheetName];
          const json = XLSX.utils.sheet_to_json<any>(sheet);
          
          json.forEach(row => {
            const name = row['이름'] || row['성명'];
            const id = row['학번']?.toString();
            
            const studentIndex = updatedStudents.findIndex(s => 
              (id && s.id === id) || 
              (name && s.name === name.toString())
            );
            
            if (studentIndex !== -1) {
              let calculatedScore = 0;
              let descriptiveTexts: string[] = [];
              let responses: Record<string, any> = {};

              Object.keys(row).forEach(key => {
                const value = row[key];
                if (value === undefined || value === null || value === '') return;
                
                responses[key] = value;

                if (['학번', '이름', '성명', '학년', '반', '번호', '총점', '점수', '위험군', '판정', 'Score', 'Risk'].includes(key)) return;

                const isDescriptive = key.includes('?') || key.includes('시간이다');
                
                if (isDescriptive) {
                  descriptiveTexts.push(String(value));
                } else {
                  const numValue = parseFloat(value);
                  if (!isNaN(numValue)) {
                    calculatedScore += numValue;
                  }
                }
              });

              const score = row['총점'] !== undefined ? parseFloat(row['총점']) : 
                            row['점수'] !== undefined ? parseFloat(row['점수']) : 
                            row['Score'] !== undefined ? parseFloat(row['Score']) : calculatedScore;
                            
              const riskLevel = row['위험군'] || row['판정'] || row['Risk'] || '';
              const descriptive = descriptiveTexts.join(' ');
              
              updatedStudents[studentIndex] = {
                ...updatedStudents[studentIndex],
                score: isNaN(score) ? undefined : score,
                riskLevel: riskLevel.toString(),
                descriptive: descriptive || updatedStudents[studentIndex].descriptive,
                responses
              };
            }
          });
        });
        
        // Categorize based on score if riskLevel is empty
        updatedStudents = updatedStudents.map(s => {
          if (s.score !== undefined) {
             if (s.score >= 80) s.category = '고위험군';
             else if (s.score >= 60) s.category = '관심군';
             else if (s.score >= 40) s.category = '주의군';
             else s.category = '일반군';
             
             if (!s.riskLevel) s.riskLevel = s.category;
          }
          return s;
        });
        
        setStudents(updatedStudents);
        setResultsLoaded(true);
      } catch (err) {
        console.error("Failed to parse results:", err);
        alert("결과 파일을 읽는 중 오류가 발생했습니다.");
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const exportRosterByClass = () => {
    if (students.length === 0) return;
    
    const workbook = XLSX.utils.book_new();
    
    const grouped = students.reduce((acc, student) => {
      const key = student.grade && student.class ? `${student.grade}학년 ${student.class}반` : '미분류';
      if (!acc[key]) acc[key] = [];
      acc[key].push(student);
      return acc;
    }, {} as Record<string, Student[]>);
    
    Object.keys(grouped).sort().forEach(key => {
      const classStudents = grouped[key].sort((a, b) => {
         const numA = parseInt(a.number) || 0;
         const numB = parseInt(b.number) || 0;
         return numA - numB;
      });
      
      const exportData = classStudents.map(s => ({
        '학번': s.id,
        '학년': s.grade,
        '반': s.class,
        '번호': s.number,
        '이름': s.name,
        '고유코드(무의미철자)': s.uniqueCode || ''
      }));
      
      const worksheet = XLSX.utils.json_to_sheet(exportData);
      // Sheet names cannot exceed 31 characters
      XLSX.utils.book_append_sheet(workbook, worksheet, key.substring(0, 31));
    });
    
    XLSX.writeFile(workbook, "학생명렬표_고유코드.xlsx");
  };

  const exportToExcel = () => {
    const exportData = students.map(s => {
      const base: Record<string, any> = {
        '학번': s.id,
        '학년': s.grade,
        '반': s.class,
        '번호': s.number,
        '이름': s.name,
        '총점': s.score ?? '',
        '분류': s.category ?? '',
        '서술형 전체응답': s.descriptive ?? ''
      };
      
      if (s.responses) {
        Object.keys(s.responses).forEach(key => {
          if (base[key] === undefined) {
            base[key] = s.responses![key];
          }
        });
      }
      
      return base;
    });
    const worksheet = XLSX.utils.json_to_sheet(exportData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "분석결과");
    XLSX.writeFile(workbook, "정서행동특성검사_분석결과.xlsx");
  };

  // --- Data Processing ---
  const stats = useMemo(() => {
    const total = students.filter(s => s.score !== undefined).length;
    const highRisk = students.filter(s => s.category === '고위험군').length;
    const interest = students.filter(s => s.category === '관심군').length;
    const caution = students.filter(s => s.category === '주의군').length;
    const normal = students.filter(s => s.category === '일반군').length;

    return { total, highRisk, interest, caution, normal };
  }, [students]);

  const chartData = useMemo(() => {
    return [
      { name: '일반군', value: stats.normal },
      { name: '주의군', value: stats.caution },
      { name: '관심군', value: stats.interest },
      { name: '고위험군', value: stats.highRisk },
    ];
  }, [stats]);

  const wordCloudData = useMemo(() => {
    const text = students.map(s => s.descriptive).filter(Boolean).join(' ');
    const words = text.split(/[\s,.]+/);
    const wordCount: Record<string, number> = {};
    const stopWords = ['그리고', '그래서', '하지만', '그런데', '이', '그', '저', '것', '수', '등', '및', '또는', '있는', '없는', '합니다', '했습니다', '입니다', '없습니다', '너무', '많이', '조금', '잘', '안', '못', '나는', '나를', '나의', '내가', '진짜', '정말', '시간이다', '시간'];
    
    words.forEach(word => {
      let cleanWord = word.replace(/(은|는|이|가|을|를|에|에게|에서|로|으로|과|와|의|도|만|까지|부터|다|요)$/, '');
      if (cleanWord.length > 1 && !stopWords.includes(cleanWord)) {
        wordCount[cleanWord] = (wordCount[cleanWord] || 0) + 1;
      }
    });
    
    return Object.entries(wordCount)
      .map(([text, value]) => ({ text, value }))
      .sort((a, b) => b.value - a.value)
      .slice(0, 50);
  }, [students]);

  const filteredStudents = useMemo(() => {
    if (!searchTerm) return students;
    return students.filter(s => 
      s.name.includes(searchTerm) || 
      s.id.includes(searchTerm) ||
      (s.riskLevel && s.riskLevel.includes(searchTerm))
    );
  }, [students, searchTerm]);

  // --- UI Render ---
  return (
    <div className="min-h-screen bg-[#f5f5f4] text-[#1a1a1a] font-sans selection:bg-blue-200">
      {/* Header */}
      <header className="bg-white border-b border-slate-200 sticky top-0 z-10">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 h-16 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="w-8 h-8 bg-blue-600 rounded-md flex items-center justify-center text-white font-bold">
              <BarChart3 size={20} />
            </div>
            <h1 className="text-xl font-bold tracking-tight">정서행동특성검사 분석기</h1>
          </div>
          <div className="flex gap-4">
            <label className={cn(
              "flex items-center gap-2 px-4 py-2 rounded-md text-sm font-medium cursor-pointer transition-colors",
              rosterLoaded ? "bg-emerald-50 text-emerald-700 border border-emerald-200" : "bg-white border border-slate-300 hover:bg-slate-50"
            )}>
              {rosterLoaded ? <CheckCircle2 size={16} /> : <Users size={16} />}
              {rosterLoaded ? "명렬표 로드됨" : "1. 명렬표 업로드"}
              <input type="file" accept=".xlsx, .xls" className="hidden" onChange={handleRosterUpload} />
            </label>

            {rosterLoaded && (
              <button 
                onClick={exportRosterByClass}
                className="flex items-center gap-2 px-4 py-2 bg-indigo-50 text-indigo-700 border border-indigo-200 rounded-md text-sm font-medium hover:bg-indigo-100 transition-colors shadow-sm"
                title="반별로 고유코드가 포함된 명렬표를 다운로드합니다"
              >
                <Download size={16} />
                고유코드 명렬표
              </button>
            )}
            
            <label className={cn(
              "flex items-center gap-2 px-4 py-2 rounded-md text-sm font-medium cursor-pointer transition-colors",
              !rosterLoaded ? "opacity-50 cursor-not-allowed bg-slate-100 border border-slate-200 text-slate-400" :
              resultsLoaded ? "bg-emerald-50 text-emerald-700 border border-emerald-200" : "bg-blue-600 text-white hover:bg-blue-700 shadow-sm"
            )}>
              {resultsLoaded ? <CheckCircle2 size={16} /> : <FileSpreadsheet size={16} />}
              {resultsLoaded ? "결과 로드됨" : "2. 검사결과 업로드"}
              <input type="file" accept=".xlsx, .xls" className="hidden" onChange={handleResultsUpload} disabled={!rosterLoaded} />
            </label>

            {resultsLoaded && (
              <button 
                onClick={exportToExcel}
                className="flex items-center gap-2 px-4 py-2 bg-slate-900 text-white rounded-md text-sm font-medium hover:bg-slate-800 transition-colors shadow-sm"
              >
                <Download size={16} />
                결과 다운로드
              </button>
            )}
          </div>
        </div>
      </header>

      <main className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-8">
        {!rosterLoaded && !resultsLoaded ? (
          <div className="flex flex-col items-center justify-center h-[60vh] text-center">
            <div className="w-24 h-24 bg-blue-50 rounded-full flex items-center justify-center mb-6">
              <Upload size={48} className="text-blue-500" />
            </div>
            <h2 className="text-2xl font-bold mb-2">데이터를 업로드해주세요</h2>
            <p className="text-slate-500 max-w-md mb-8">
              우측 상단의 버튼을 통해 학생 명렬표(엑셀)를 먼저 업로드한 후, 검사 결과(엑셀)를 업로드하면 자동으로 데이터를 매칭하여 분석합니다.
            </p>
            <div className="grid grid-cols-2 gap-4 text-sm text-left w-full max-w-2xl">
              <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-sm">
                <h3 className="font-bold flex items-center gap-2 mb-2"><Users size={16} className="text-blue-500"/> 명렬표 권장 양식</h3>
                <p className="text-slate-500 mb-2">시트 탭으로 반을 구분하거나, 열에 반 정보를 포함하세요.</p>
                <code className="bg-slate-100 px-2 py-1 rounded text-xs text-slate-700 block">학번 | 학년 | 반 | 번호 | 이름</code>
              </div>
              <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-sm">
                <h3 className="font-bold flex items-center gap-2 mb-2"><FileSpreadsheet size={16} className="text-emerald-500"/> 결과표 권장 양식</h3>
                <p className="text-slate-500 mb-2">이름 또는 학번을 기준으로 매칭됩니다.</p>
                <code className="bg-slate-100 px-2 py-1 rounded text-xs text-slate-700 block">학번 | 이름 | 총점 | 위험군 | 서술형</code>
              </div>
            </div>
          </div>
        ) : (
          <div className="space-y-6">
            {/* Navigation Tabs */}
            <div className="flex space-x-1 bg-slate-200/50 p-1 rounded-lg w-fit">
              <button
                onClick={() => setActiveTab('overview')}
                className={cn("px-4 py-2 text-sm font-medium rounded-md transition-all", activeTab === 'overview' ? "bg-white text-slate-900 shadow-sm" : "text-slate-600 hover:text-slate-900")}
              >
                대시보드 요약
              </button>
              <button
                onClick={() => setActiveTab('students')}
                className={cn("px-4 py-2 text-sm font-medium rounded-md transition-all", activeTab === 'students' ? "bg-white text-slate-900 shadow-sm" : "text-slate-600 hover:text-slate-900")}
              >
                학생 명단 및 선별
              </button>
              <button
                onClick={() => setActiveTab('wordcloud')}
                className={cn("px-4 py-2 text-sm font-medium rounded-md transition-all", activeTab === 'wordcloud' ? "bg-white text-slate-900 shadow-sm" : "text-slate-600 hover:text-slate-900")}
              >
                서술형 워드클라우드
              </button>
            </div>

            {/* Tab Content: Overview */}
            {activeTab === 'overview' && (
              <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
                {/* Stats Cards */}
                <div className="col-span-1 md:col-span-3 grid grid-cols-2 md:grid-cols-5 gap-4">
                  <div className="bg-white p-6 rounded-2xl border border-slate-200 shadow-sm flex flex-col justify-between">
                    <span className="text-sm font-medium text-slate-500 uppercase tracking-wider">총 검사 인원</span>
                    <span className="text-4xl font-light mt-2">{stats.total}</span>
                  </div>
                  <div className="bg-white p-6 rounded-2xl border border-slate-200 shadow-sm flex flex-col justify-between border-b-4 border-b-emerald-500">
                    <span className="text-sm font-medium text-slate-500 uppercase tracking-wider">일반군</span>
                    <span className="text-4xl font-light mt-2 text-emerald-600">{stats.normal}</span>
                  </div>
                  <div className="bg-white p-6 rounded-2xl border border-slate-200 shadow-sm flex flex-col justify-between border-b-4 border-b-amber-500">
                    <span className="text-sm font-medium text-slate-500 uppercase tracking-wider">주의군</span>
                    <span className="text-4xl font-light mt-2 text-amber-600">{stats.caution}</span>
                  </div>
                  <div className="bg-white p-6 rounded-2xl border border-slate-200 shadow-sm flex flex-col justify-between border-b-4 border-b-orange-500">
                    <span className="text-sm font-medium text-slate-500 uppercase tracking-wider">관심군</span>
                    <span className="text-4xl font-light mt-2 text-orange-600">{stats.interest}</span>
                  </div>
                  <div className="bg-white p-6 rounded-2xl border border-slate-200 shadow-sm flex flex-col justify-between border-b-4 border-b-red-500">
                    <span className="text-sm font-medium text-slate-500 uppercase tracking-wider">고위험군</span>
                    <span className="text-4xl font-light mt-2 text-red-600">{stats.highRisk}</span>
                  </div>
                </div>

                {/* Charts */}
                <div className="col-span-1 md:col-span-2 bg-white p-6 rounded-2xl border border-slate-200 shadow-sm min-h-[400px]">
                  <h3 className="text-lg font-bold mb-6">위험군 분포 현황</h3>
                  <ResponsiveContainer width="100%" height={300}>
                    <BarChart data={chartData} margin={{ top: 20, right: 30, left: 20, bottom: 5 }}>
                      <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
                      <XAxis dataKey="name" axisLine={false} tickLine={false} />
                      <YAxis axisLine={false} tickLine={false} />
                      <Tooltip cursor={{fill: '#f1f5f9'}} contentStyle={{borderRadius: '8px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)'}} />
                      <Bar dataKey="value" radius={[4, 4, 0, 0]}>
                        {chartData.map((entry, index) => (
                          <Cell key={`cell-${index}`} fill={COLORS[index % COLORS.length]} />
                        ))}
                      </Bar>
                    </BarChart>
                  </ResponsiveContainer>
                </div>

                <div className="col-span-1 bg-white p-6 rounded-2xl border border-slate-200 shadow-sm min-h-[400px]">
                  <h3 className="text-lg font-bold mb-6">비율</h3>
                  <ResponsiveContainer width="100%" height={300}>
                    <PieChart>
                      <Pie
                        data={chartData}
                        cx="50%"
                        cy="50%"
                        innerRadius={60}
                        outerRadius={100}
                        paddingAngle={5}
                        dataKey="value"
                      >
                        {chartData.map((entry, index) => (
                          <Cell key={`cell-${index}`} fill={COLORS[index % COLORS.length]} />
                        ))}
                      </Pie>
                      <Tooltip contentStyle={{borderRadius: '8px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)'}} />
                      <Legend verticalAlign="bottom" height={36}/>
                    </PieChart>
                  </ResponsiveContainer>
                </div>
              </div>
            )}

            {/* Tab Content: Students List */}
            {activeTab === 'students' && (
              <div className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden flex flex-col h-[calc(100vh-200px)]">
                <div className="p-4 border-b border-slate-200 flex justify-between items-center bg-slate-50/50">
                  <div className="relative w-72">
                    <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" size={18} />
                    <input 
                      type="text" 
                      placeholder="이름, 학번, 위험군 검색..." 
                      className="w-full pl-10 pr-4 py-2 rounded-lg border border-slate-300 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent text-sm"
                      value={searchTerm}
                      onChange={(e) => setSearchTerm(e.target.value)}
                    />
                  </div>
                  <div className="text-sm text-slate-500">
                    총 {filteredStudents.length}명
                  </div>
                </div>
                
                <div className="overflow-auto flex-1">
                  <table className="w-full text-left text-sm">
                    <thead className="bg-slate-50 sticky top-0 z-10 shadow-sm">
                      <tr>
                        <th className="px-6 py-3 font-medium text-slate-500">학번</th>
                        <th className="px-6 py-3 font-medium text-slate-500">이름</th>
                        <th className="px-6 py-3 font-medium text-slate-500">고유코드</th>
                        <th className="px-6 py-3 font-medium text-slate-500">총점</th>
                        <th className="px-6 py-3 font-medium text-slate-500">분류/위험군</th>
                        <th className="px-6 py-3 font-medium text-slate-500">서술형 의견</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-200">
                      {filteredStudents.map((student, i) => (
                        <tr key={i} className="hover:bg-slate-50 transition-colors group">
                          <td className="px-6 py-4 font-mono text-slate-600">{student.id}</td>
                          <td className="px-6 py-4 font-medium">{student.name}</td>
                          <td className="px-6 py-4 font-mono text-indigo-600 font-medium">{student.uniqueCode}</td>
                          <td className="px-6 py-4">{student.score !== undefined ? student.score : '-'}</td>
                          <td className="px-6 py-4">
                            {student.category && (
                              <span className={cn(
                                "inline-flex items-center px-2.5 py-0.5 rounded-full text-xs font-medium",
                                student.category === '고위험군' ? "bg-red-100 text-red-800" :
                                student.category === '관심군' ? "bg-orange-100 text-orange-800" :
                                student.category === '주의군' ? "bg-amber-100 text-amber-800" :
                                "bg-emerald-100 text-emerald-800"
                              )}>
                                {student.category}
                              </span>
                            )}
                          </td>
                          <td className="px-6 py-4 text-slate-600 max-w-md truncate group-hover:whitespace-normal group-hover:break-words transition-all">
                            {student.descriptive || '-'}
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            )}

            {/* Tab Content: Word Cloud */}
            {activeTab === 'wordcloud' && (
              <div className="bg-white p-6 rounded-2xl border border-slate-200 shadow-sm h-[calc(100vh-200px)] flex flex-col">
                <div className="flex items-center gap-2 mb-6">
                  <Cloud className="text-blue-500" />
                  <h3 className="text-lg font-bold">서술형 문항 주요 키워드</h3>
                </div>
                <div className="flex-1 relative bg-slate-50 rounded-xl border border-slate-100 overflow-hidden">
                  {wordCloudData.length > 0 ? (
                    <WordCloud words={wordCloudData} width={800} height={500} />
                  ) : (
                    <div className="absolute inset-0 flex items-center justify-center text-slate-400">
                      서술형 데이터가 충분하지 않습니다.
                    </div>
                  )}
                </div>
              </div>
            )}
          </div>
        )}
      </main>
    </div>
  );
}
