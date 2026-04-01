import * as XLSX from 'xlsx';

export interface Student {
  id: string;
  grade: string;
  class: string;
  number: string;
  name: string;
  score?: number;
  riskLevel?: string;
  descriptive?: string;
  category?: string;
}

export const parseRoster = async (file: File): Promise<Student[]> => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        
        let students: Student[] = [];
        
        // Iterate through all sheets
        workbook.SheetNames.forEach(sheetName => {
          const sheet = workbook.Sheets[sheetName];
          const json = XLSX.utils.sheet_to_json<any>(sheet);
          
          json.forEach(row => {
            // Try to find common column names
            const name = row['이름'] || row['성명'] || row['Name'] || row['name'];
            const grade = row['학년'] || row['Grade'] || '';
            const cls = row['반'] || row['Class'] || sheetName.replace(/[^0-9]/g, ''); // Fallback to sheet name
            const num = row['번호'] || row['Number'] || '';
            const id = row['학번'] || row['ID'] || `${grade}${cls.toString().padStart(2, '0')}${num.toString().padStart(2, '0')}`;
            
            if (name) {
              students.push({
                id: id.toString(),
                grade: grade.toString(),
                class: cls.toString(),
                number: num.toString(),
                name: name.toString()
              });
            }
          });
        });
        
        resolve(students);
      } catch (err) {
        reject(err);
      }
    };
    reader.readAsArrayBuffer(file);
  });
};

export const parseResults = async (file: File, existingStudents: Student[]): Promise<Student[]> => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        
        let updatedStudents = [...existingStudents];
        
        workbook.SheetNames.forEach(sheetName => {
          const sheet = workbook.Sheets[sheetName];
          const json = XLSX.utils.sheet_to_json<any>(sheet);
          
          json.forEach(row => {
            const name = row['이름'] || row['성명'];
            const id = row['학번'];
            
            // Find student by ID or Name
            const studentIndex = updatedStudents.findIndex(s => 
              (id && s.id === id.toString()) || 
              (name && s.name === name.toString())
            );
            
            if (studentIndex !== -1) {
              const score = parseFloat(row['총점'] || row['점수'] || row['Score'] || '0');
              const riskLevel = row['위험군'] || row['판정'] || row['Risk'] || (score >= 60 ? '관심군' : score >= 80 ? '위험군' : '정상');
              const descriptive = row['서술형'] || row['기타의견'] || row['주관식'] || row['의견'] || '';
              
              updatedStudents[studentIndex] = {
                ...updatedStudents[studentIndex],
                score: isNaN(score) ? 0 : score,
                riskLevel: riskLevel.toString(),
                descriptive: descriptive.toString()
              };
            }
          });
        });
        
        // Simple clustering based on score ranges if category is not provided
        updatedStudents = updatedStudents.map(s => {
          if (s.score !== undefined && !s.category) {
            if (s.score >= 80) s.category = '고위험 (전문기관 연계 필요)';
            else if (s.score >= 60) s.category = '관심군 (지속 관찰 필요)';
            else if (s.score >= 40) s.category = '주의군 (상담 권장)';
            else s.category = '일반군 (정상)';
          }
          return s;
        });
        
        resolve(updatedStudents);
      } catch (err) {
        reject(err);
      }
    };
    reader.readAsArrayBuffer(file);
  });
};

export const extractWords = (students: Student[]): { text: string; value: number }[] => {
  const text = students
    .map(s => s.descriptive)
    .filter(Boolean)
    .join(' ');
    
  // Simple Korean word extraction (very basic, splits by space and removes common particles)
  const words = text.split(/[\s,.]+/);
  const wordCount: Record<string, number> = {};
  
  const stopWords = ['그리고', '그래서', '하지만', '그런데', '이', '그', '저', '것', '수', '등', '및', '또는', '있는', '없는', '합니다', '했습니다', '입니다', '없습니다'];
  
  words.forEach(word => {
    // Remove basic Korean particles
    let cleanWord = word.replace(/(은|는|이|가|을|를|에|에게|에서|로|으로|과|와|의)$/, '');
    
    if (cleanWord.length > 1 && !stopWords.includes(cleanWord)) {
      wordCount[cleanWord] = (wordCount[cleanWord] || 0) + 1;
    }
  });
  
  return Object.entries(wordCount)
    .map(([text, value]) => ({ text, value }))
    .sort((a, b) => b.value - a.value)
    .slice(0, 50); // Top 50 words
};
