import React, { useState, useMemo, useEffect } from 'react';
import { Upload, FileSpreadsheet, Users, AlertTriangle, Cloud, BarChart3, Search, Download, CheckCircle2, Sparkles, Loader2 } from 'lucide-react';
import { cn } from './lib/utils';
import * as XLSX from 'xlsx';
import WordCloud from './components/WordCloud';
import { BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer, PieChart, Pie, Cell } from 'recharts';
import { GoogleGenAI } from '@google/genai';
import Markdown from 'react-markdown';

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

const parseScaleValue = (val: any): number | null => {
  if (val === undefined || val === null || val === '') return null;
  const strVal = String(val).replace(/\s+/g, '');
  if (strVal === '전혀아니다') return 0;
  if (strVal === '조금그렇다') return 1;
  if (strVal === '그렇다') return 2;
  if (strVal === '매우그렇다') return 3;
  
  const num = parseFloat(String(val).trim());
  return isNaN(num) ? null : num;
};

export default function App() {
  const [students, setStudents] = useState<Student[]>([]);
  const [rosterLoaded, setRosterLoaded] = useState(false);
  const [resultsLoaded, setResultsLoaded] = useState(false);
  const [activeTab, setActiveTab] = useState<'overview' | 'students' | 'wordcloud' | 'unmatched' | 'questions'>('overview');
  const [searchTerm, setSearchTerm] = useState('');
  const [descriptiveQuestions, setDescriptiveQuestions] = useState<string[]>([]);
  const [selectedQuestion, setSelectedQuestion] = useState<string>('all');
  const [aiAnalysis, setAiAnalysis] = useState<string>('');
  const [isAnalyzing, setIsAnalyzing] = useState(false);
  const [analysisError, setAnalysisError] = useState('');
  const [unmatchedResults, setUnmatchedResults] = useState<any[]>([]);

  useEffect(() => {
    setAiAnalysis('');
    setAnalysisError('');
  }, [selectedQuestion]);

  const analyzeWithAI = async () => {
    if (!process.env.GEMINI_API_KEY) {
      setAnalysisError('Gemini API 키가 설정되지 않았습니다.');
      return;
    }
    
    setIsAnalyzing(true);
    setAnalysisError('');
    
    try {
      const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY });
      
      let questionText = selectedQuestion === 'all' ? '전체 서술형 문항' : selectedQuestion;
      let responsesText = '';
      
      if (selectedQuestion === 'all') {
        responsesText = students.map(s => s.descriptive).filter(Boolean).join('\n- ');
      } else {
        responsesText = students.map(s => s.responses?.[selectedQuestion]).filter(Boolean).join('\n- ');
      }
      
      if (!responsesText.trim()) {
        setAnalysisError('분석할 응답 데이터가 없습니다.');
        setIsAnalyzing(false);
        return;
      }

      const prompt = `다음은 정서행동특성검사에서 학생들의 서술형 응답입니다.
문항: ${questionText}

학생들의 응답:
- ${responsesText}

위 응답들을 바탕으로 다음 사항들을 분석해주세요:
1. 학생들의 전반적인 정서/행동 특성 및 주요 관심사 요약
2. 긍정적인 측면과 부정적인/우려되는 측면
3. 교사가 학급 운영이나 상담 시 주의 깊게 살펴봐야 할 점 및 지도 조언

전문적이고 따뜻한 교사의 시각에서 분석 결과를 작성해주세요.`;

      const response = await ai.models.generateContent({
        model: 'gemini-3-flash-preview',
        contents: prompt,
      });
      
      setAiAnalysis(response.text || '분석 결과를 생성하지 못했습니다.');
    } catch (err) {
      console.error('AI Analysis error:', err);
      setAnalysisError('AI 분석 중 오류가 발생했습니다. 다시 시도해주세요.');
    } finally {
      setIsAnalyzing(false);
    }
  };

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
            // Generate 5-character alphanumeric code
            code = Math.random().toString(36).substring(2, 7).toUpperCase();
          } while (existingCodes.has(code));
          existingCodes.add(code);
          return code;
        };
        
        workbook.SheetNames.forEach(sheetName => {
          const sheet = workbook.Sheets[sheetName];
          const json = XLSX.utils.sheet_to_json<any>(sheet);
          
          json.forEach(row => {
            // 띄어쓰기 무시하고 키워드 포함 여부로 컬럼 찾기
            const findKey = (keywords: string[]) => Object.keys(row).find(k => keywords.some(kw => k.replace(/\s+/g, '').toLowerCase().includes(kw)));
            
            const nameKey = findKey(['이름', '성명', 'name']);
            const gradeKey = findKey(['학년', 'grade']);
            const clsKey = findKey(['반', 'class']);
            const numKey = findKey(['번호', 'number']);
            const idKey = findKey(['학번', 'id']);
            const codeKey = findKey(['고유코드', '코드', 'code']);

            const name = nameKey ? row[nameKey] : '';
            const grade = gradeKey ? row[gradeKey] : '';
            const cls = clsKey ? row[clsKey] : sheetName.replace(/[^0-9]/g, '');
            const num = numKey ? row[numKey] : '';
            const id = idKey ? row[idKey] : `${grade}${cls.toString().padStart(2, '0')}${num.toString().padStart(2, '0')}`;
            const existingCode = codeKey ? row[codeKey]?.toString().trim().toUpperCase() : '';
            
            if (name) {
              const finalCode = existingCode || generateUniqueCode();
              if (existingCode) existingCodes.add(existingCode);
              
              parsedStudents.push({
                id: id.toString(),
                grade: grade.toString(),
                class: cls.toString(),
                number: num.toString(),
                name: name.toString(),
                uniqueCode: finalCode
              });
            }
          });
        });
        
        if (parsedStudents.length === 0) {
          alert("명렬표에서 학생 데이터를 찾을 수 없습니다.\n엑셀 파일의 첫 번째 줄(헤더)에 '이름' 또는 '성명' 열이 있는지 확인해주세요.");
          e.target.value = ''; // input 초기화
          setRosterLoaded(false);
          return;
        }

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
        let matchCount = 0;
        let unmatched: any[] = [];
        const descKeys = new Set<string>();
        
        workbook.SheetNames.forEach(sheetName => {
          const sheet = workbook.Sheets[sheetName];
          const json = XLSX.utils.sheet_to_json<any>(sheet);
          
          json.forEach(row => {
            const findKey = (keywords: string[]) => Object.keys(row).find(k => keywords.some(kw => k.replace(/\s+/g, '').toLowerCase().includes(kw)));
            
            const nameKey = findKey(['이름', '성명', 'name']);
            const idKey = findKey(['학번', 'id']);
            const totalKey = findKey(['총점', '점수', 'score']);
            const riskKey = findKey(['위험군', '판정', 'risk']);
            const codeKey = findKey(['고유코드', '코드', 'code']);

            const name = nameKey ? row[nameKey] : '';
            const id = idKey ? row[idKey]?.toString() : '';
            const code = codeKey ? row[codeKey]?.toString().trim().toUpperCase() : '';
            
            const studentIndex = updatedStudents.findIndex(s => 
              (code && s.uniqueCode === code) ||
              (id && s.id === id) || 
              (name && s.name === name.toString())
            );
            
            if (studentIndex !== -1) {
              matchCount++;
              let calculatedScore = 0;
              let descriptiveTexts: string[] = [];
              let responses: Record<string, any> = {};

              Object.keys(row).forEach(key => {
                const value = row[key];
                if (value === undefined || value === null || value === '') return;
                
                responses[key] = value;

                const normalizedKey = key.replace(/\s+/g, '').toLowerCase();
                if (['학번', '이름', '성명', '학년', '반', '번호', '총점', '점수', '위험군', '판정', 'score', 'risk', '고유코드'].some(k => normalizedKey.includes(k))) return;

                const isDescriptive = key.includes('?') || key.includes('시간이다');
                
                if (isDescriptive) {
                  descKeys.add(key);
                  descriptiveTexts.push(String(value));
                } else {
                  const numValue = parseScaleValue(value);
                  if (numValue !== null) {
                    calculatedScore += numValue;
                  }
                }
              });

              const score = totalKey && row[totalKey] !== undefined ? parseFloat(row[totalKey]) : calculatedScore;
                            
              const riskLevel = riskKey ? row[riskKey] : '';
              const descriptive = descriptiveTexts.join(' ');
              
              updatedStudents[studentIndex] = {
                ...updatedStudents[studentIndex],
                score: isNaN(score) ? undefined : score,
                riskLevel: riskLevel.toString(),
                descriptive: descriptive || updatedStudents[studentIndex].descriptive,
                responses
              };
            } else {
              // Unmatched case
              if (name || id || code) {
                unmatched.push(row);
              }
            }
          });
        });
        
        if (matchCount === 0 && unmatched.length > 0) {
          alert(`명렬표와 일치하는 학생을 한 명도 찾을 수 없습니다.\n매칭되지 않은 데이터 ${unmatched.length}건이 발견되었습니다.`);
        } else if (matchCount === 0) {
          alert("업로드한 검사결과에서 유효한 데이터를 찾을 수 없습니다.\n파일을 다시 확인해주세요.");
          e.target.value = '';
          return;
        }

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
        setUnmatchedResults(unmatched);
        setDescriptiveQuestions(Array.from(descKeys));
        setResultsLoaded(true);
      } catch (err) {
        console.error("Failed to parse results:", err);
        alert("결과 파일을 읽는 중 오류가 발생했습니다.");
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const exportRosterByClass = () => {
    try {
      if (students.length === 0) {
        alert("다운로드할 명단이 없습니다.");
        return;
      }
      
      const workbook = XLSX.utils.book_new();
      
      const grouped = students.reduce((acc, student) => {
        // 엑셀 시트 이름에 사용할 수 없는 특수문자 제거
        const safeGrade = (student.grade || '').toString().replace(/[\[\]\/*?:\\]/g, '');
        const safeClass = (student.class || '').toString().replace(/[\[\]\/*?:\\]/g, '');
        
        let key = safeGrade && safeClass ? `${safeGrade}학년 ${safeClass}반` : 
                  safeClass ? `${safeClass}반` : 
                  safeGrade ? `${safeGrade}학년` : '미분류';
                  
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
          '고유코드': s.uniqueCode || ''
        }));
        
        const worksheet = XLSX.utils.json_to_sheet(exportData);
        // 엑셀 시트 이름은 최대 31자 제한
        XLSX.utils.book_append_sheet(workbook, worksheet, key.substring(0, 31));
      });
      
      XLSX.writeFile(workbook, "학생명렬표_고유코드.xlsx");
    } catch (error) {
      console.error("Export error:", error);
      alert("엑셀 파일 생성 중 오류가 발생했습니다. 콘솔 창을 확인해주세요.");
    }
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
    let text = '';
    if (selectedQuestion === 'all') {
      text = students.map(s => s.descriptive).filter(Boolean).join(' ');
    } else {
      text = students.map(s => s.responses?.[selectedQuestion]).filter(Boolean).join(' ');
    }
    
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
  }, [students, selectedQuestion]);

  const filteredStudents = useMemo(() => {
    if (!searchTerm) return students;
    return students.filter(s => 
      s.name.includes(searchTerm) || 
      s.id.includes(searchTerm) ||
      (s.riskLevel && s.riskLevel.includes(searchTerm))
    );
  }, [students, searchTerm]);

  const questionStats = useMemo(() => {
    if (!resultsLoaded) return [];
    
    const submittedStudents = students.filter(s => s.responses && Object.keys(s.responses).length > 0);
    const totalSubmitted = submittedStudents.length;
    
    if (totalSubmitted === 0) return [];
    
    const allQuestions = new Set<string>();
    submittedStudents.forEach(s => {
      Object.keys(s.responses || {}).forEach(q => allQuestions.add(q));
    });
    
    return Array.from(allQuestions).map(q => {
      const isDescriptive = q.includes('?') || q.includes('시간이다');
      let respondedCount = 0;
      let sum = 0;
      const missingStudents: string[] = [];
      
      submittedStudents.forEach(s => {
        const val = s.responses?.[q];
        if (val !== undefined && val !== null && val !== '') {
          respondedCount++;
          if (!isDescriptive) {
            const num = parseScaleValue(val);
            if (num !== null) sum += num;
          }
        } else {
          missingStudents.push(`${s.grade}-${s.class}-${s.number} ${s.name}`);
        }
      });
      
      return {
        question: q,
        isDescriptive,
        respondedCount,
        totalSubmitted,
        missingStudents,
        average: isDescriptive || respondedCount === 0 ? null : (sum / respondedCount).toFixed(2)
      };
    });
  }, [students, resultsLoaded]);

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
                <p className="text-slate-500 mb-2">고유코드, 이름 또는 학번을 기준으로 매칭됩니다.</p>
                <code className="bg-slate-100 px-2 py-1 rounded text-xs text-slate-700 block">고유코드 | 학번 | 이름 | 총점 | 위험군 | 서술형</code>
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
              <button
                onClick={() => setActiveTab('questions')}
                className={cn("px-4 py-2 text-sm font-medium rounded-md transition-all", activeTab === 'questions' ? "bg-white text-slate-900 shadow-sm" : "text-slate-600 hover:text-slate-900")}
              >
                문항별 입력 현황
              </button>
              {unmatchedResults.length > 0 && (
                <button
                  onClick={() => setActiveTab('unmatched')}
                  className={cn("px-4 py-2 text-sm font-medium rounded-md transition-all flex items-center gap-2", activeTab === 'unmatched' ? "bg-red-50 text-red-700 shadow-sm border border-red-100" : "text-red-500 hover:text-red-700 hover:bg-red-50/50")}
                >
                  <AlertTriangle size={16} />
                  미매칭 데이터 ({unmatchedResults.length})
                </button>
              )}
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
              <div className="flex flex-col gap-6">
                <div className="bg-white p-6 rounded-2xl border border-slate-200 shadow-sm h-[500px] flex flex-col">
                  <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4 mb-6">
                    <div className="flex items-center gap-2">
                      <Cloud className="text-blue-500" />
                      <h3 className="text-lg font-bold">서술형 문항 주요 키워드</h3>
                    </div>
                    {descriptiveQuestions.length > 0 && (
                      <select
                        value={selectedQuestion}
                        onChange={(e) => setSelectedQuestion(e.target.value)}
                        className="border border-slate-300 rounded-lg px-4 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500 max-w-full sm:max-w-md bg-white shadow-sm"
                      >
                        <option value="all">전체 문항 통합</option>
                        {descriptiveQuestions.map((q, idx) => (
                          <option key={idx} value={q}>{q}</option>
                        ))}
                      </select>
                    )}
                  </div>
                  <div className="flex-1 relative bg-slate-50 rounded-xl border border-slate-100 overflow-hidden">
                    {wordCloudData.length > 0 ? (
                      <WordCloud words={wordCloudData} width={800} height={400} />
                    ) : (
                      <div className="absolute inset-0 flex items-center justify-center text-slate-400">
                        서술형 데이터가 충분하지 않습니다.
                      </div>
                    )}
                  </div>
                </div>

                {/* AI Analysis Section */}
                <div className="bg-indigo-50 p-6 rounded-2xl border border-indigo-100 shadow-sm">
                  <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4 mb-6">
                    <div className="flex items-center gap-2">
                      <Sparkles className="text-indigo-600" />
                      <h3 className="text-lg font-bold text-indigo-900">AI 문항 응답 분석</h3>
                    </div>
                    <button
                      onClick={analyzeWithAI}
                      disabled={isAnalyzing || wordCloudData.length === 0}
                      className={cn(
                        "flex items-center gap-2 px-4 py-2 rounded-lg text-sm font-medium transition-colors shadow-sm",
                        isAnalyzing || wordCloudData.length === 0
                          ? "bg-indigo-200 text-indigo-400 cursor-not-allowed"
                          : "bg-indigo-600 text-white hover:bg-indigo-700"
                      )}
                    >
                      {isAnalyzing ? (
                        <>
                          <Loader2 size={16} className="animate-spin" />
                          분석 중...
                        </>
                      ) : (
                        <>
                          <Sparkles size={16} />
                          AI 분석 시작
                        </>
                      )}
                    </button>
                  </div>

                  {analysisError && (
                    <div className="p-4 bg-red-50 text-red-700 rounded-lg border border-red-200 mb-4 text-sm">
                      {analysisError}
                    </div>
                  )}

                  {aiAnalysis ? (
                    <div className="bg-white p-6 rounded-xl border border-indigo-100 shadow-sm prose prose-indigo max-w-none text-slate-700 text-sm sm:text-base">
                      <div className="markdown-body">
                        <Markdown>{aiAnalysis}</Markdown>
                      </div>
                    </div>
                  ) : (
                    !isAnalyzing && !analysisError && (
                      <div className="flex flex-col items-center justify-center py-12 text-center text-indigo-400">
                        <Sparkles size={32} className="mb-3 opacity-50" />
                        <p>우측 상단의 버튼을 눌러 학생들의 응답을 AI로 분석해보세요.</p>
                        <p className="text-sm mt-1 opacity-75">선택된 문항에 대한 전반적인 특성, 긍정/우려되는 점, 지도 조언을 제공합니다.</p>
                      </div>
                    )
                  )}
                </div>
              </div>
            )}
            {/* Tab Content: Unmatched */}
            {activeTab === 'unmatched' && unmatchedResults.length > 0 && (
              <div className="bg-white rounded-2xl border border-red-200 shadow-sm overflow-hidden flex flex-col h-[calc(100vh-200px)]">
                <div className="p-4 border-b border-red-100 flex justify-between items-center bg-red-50/50">
                  <div className="flex items-center gap-2 text-red-700">
                    <AlertTriangle size={20} />
                    <h3 className="font-bold">명렬표와 매칭되지 않은 결과 데이터</h3>
                  </div>
                  <div className="text-sm font-medium text-red-600 bg-red-100 px-3 py-1 rounded-full">
                    총 {unmatchedResults.length}건
                  </div>
                </div>
                
                <div className="p-4 bg-red-50/30 border-b border-red-100 text-sm text-red-600">
                  아래 데이터는 업로드된 결과표에 존재하지만, 명렬표의 <strong>고유코드, 이름, 학번</strong> 중 어느 것과도 일치하지 않아 분석에서 제외된 항목들입니다. 학생이 코드를 잘못 입력했거나 명렬표에 없는 학생일 수 있습니다.
                </div>

                <div className="overflow-auto flex-1">
                  <table className="w-full text-left text-sm">
                    <thead className="bg-slate-50 sticky top-0 z-10 shadow-sm">
                      <tr>
                        <th className="px-6 py-3 font-medium text-slate-500">제출된 고유코드</th>
                        <th className="px-6 py-3 font-medium text-slate-500">제출된 학번</th>
                        <th className="px-6 py-3 font-medium text-slate-500">제출된 이름</th>
                        <th className="px-6 py-3 font-medium text-slate-500">원시 데이터 (전체)</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-200">
                      {unmatchedResults.map((row, i) => {
                        const findKey = (keywords: string[]) => Object.keys(row).find(k => keywords.some(kw => k.replace(/\s+/g, '').toLowerCase().includes(kw)));
                        const nameKey = findKey(['이름', '성명', 'name']);
                        const idKey = findKey(['학번', 'id']);
                        const codeKey = findKey(['고유코드', '코드', 'code']);
                        
                        return (
                          <tr key={i} className="hover:bg-red-50/30 transition-colors">
                            <td className="px-6 py-4 font-mono text-slate-600">{codeKey ? row[codeKey] : '-'}</td>
                            <td className="px-6 py-4 font-mono text-slate-600">{idKey ? row[idKey] : '-'}</td>
                            <td className="px-6 py-4 font-medium">{nameKey ? row[nameKey] : '-'}</td>
                            <td className="px-6 py-4">
                              <details className="cursor-pointer text-slate-500">
                                <summary className="hover:text-slate-800 font-medium">데이터 보기</summary>
                                <pre className="mt-2 p-3 bg-slate-100 rounded-md text-xs overflow-x-auto max-w-md">
                                  {JSON.stringify(row, null, 2)}
                                </pre>
                              </details>
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </div>
            )}

            {/* Tab Content: Questions */}
            {activeTab === 'questions' && (
              <div className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden flex flex-col h-[calc(100vh-200px)]">
                <div className="p-4 border-b border-slate-100 flex justify-between items-center bg-slate-50">
                  <div className="flex items-center gap-2">
                    <CheckCircle2 className="text-blue-500" size={20} />
                    <h3 className="font-bold">문항별 입력 현황</h3>
                  </div>
                  <div className="text-sm text-slate-500">
                    총 제출 인원: {students.filter(s => s.responses && Object.keys(s.responses).length > 0).length}명
                  </div>
                </div>
                
                <div className="overflow-auto flex-1 p-0">
                  <table className="w-full text-left text-sm">
                    <thead className="bg-slate-50 sticky top-0 z-10 shadow-sm">
                      <tr>
                        <th className="px-6 py-3 font-medium text-slate-500 w-1/2">문항</th>
                        <th className="px-6 py-3 font-medium text-slate-500 text-center">유형</th>
                        <th className="px-6 py-3 font-medium text-slate-500 text-center">응답률</th>
                        <th className="px-6 py-3 font-medium text-slate-500 text-center">평균 점수</th>
                        <th className="px-6 py-3 font-medium text-slate-500">미응답 학생</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-200">
                      {questionStats.map((stat, i) => (
                        <tr key={i} className="hover:bg-slate-50 transition-colors">
                          <td className="px-6 py-4 font-medium text-slate-800">{stat.question}</td>
                          <td className="px-6 py-4 text-center">
                            <span className={cn("px-2 py-1 rounded-full text-xs font-medium", stat.isDescriptive ? "bg-blue-50 text-blue-600" : "bg-emerald-50 text-emerald-600")}>
                              {stat.isDescriptive ? '서술형' : '척도 (0~3)'}
                            </span>
                          </td>
                          <td className="px-6 py-4 text-center">
                            <div className="flex flex-col items-center">
                              <span className={cn("font-bold", stat.respondedCount === stat.totalSubmitted ? "text-emerald-600" : "text-amber-600")}>
                                {stat.respondedCount} / {stat.totalSubmitted}
                              </span>
                              <span className="text-xs text-slate-400">
                                ({Math.round((stat.respondedCount / stat.totalSubmitted) * 100)}%)
                              </span>
                            </div>
                          </td>
                          <td className="px-6 py-4 text-center font-mono text-slate-600">
                            {stat.average !== null ? stat.average : '-'}
                          </td>
                          <td className="px-6 py-4">
                            {stat.missingStudents.length === 0 ? (
                              <span className="text-emerald-500 text-xs flex items-center gap-1"><CheckCircle2 size={12}/> 전원 응답</span>
                            ) : (
                              <details className="cursor-pointer text-amber-600 text-xs">
                                <summary className="hover:text-amber-800 font-medium outline-none">
                                  미응답 {stat.missingStudents.length}명 보기
                                </summary>
                                <div className="mt-2 p-2 bg-amber-50 rounded border border-amber-100 max-h-32 overflow-y-auto">
                                  {stat.missingStudents.join(', ')}
                                </div>
                              </details>
                            )}
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            )}
          </div>
        )}
      </main>
    </div>
  );
}
