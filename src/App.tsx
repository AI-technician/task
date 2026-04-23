/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import { useState, useEffect, useRef } from 'react';
import { 
  Plus, 
  Trash2, 
  CheckCircle2, 
  Circle, 
  CloudSun, 
  Search, 
  Download, 
  FileUp, 
  ExternalLink,
  X,
  Newspaper
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import * as XLSX from 'xlsx';

interface Todo {
  id: string;
  text: string;
  completed: boolean;
  createdAt: number;
}

interface NewsKeyword {
  id: string;
  word: string;
}

export default function App() {
  const [todos, setTodos] = useState<Todo[]>([]);
  const [todoInput, setTodoInput] = useState('');
  const [keywords, setKeywords] = useState<NewsKeyword[]>([]);
  const [keywordInput, setKeywordInput] = useState('');
  const fileInputRef = useRef<HTMLInputElement>(null);

  // Load data from LocalStorage on mount
  useEffect(() => {
    const savedTodos = localStorage.getItem('dash_todos');
    const savedKeywords = localStorage.getItem('dash_keywords');
    if (savedTodos) setTodos(JSON.parse(savedTodos));
    if (savedKeywords) setKeywords(JSON.parse(savedKeywords));
  }, []);

  // Save data to LocalStorage whenever it changes
  useEffect(() => {
    localStorage.setItem('dash_todos', JSON.stringify(todos));
  }, [todos]);

  useEffect(() => {
    localStorage.setItem('dash_keywords', JSON.stringify(keywords));
  }, [keywords]);

  // To-Do Handlers
  const addTodo = (e: React.FormEvent) => {
    e.preventDefault();
    if (!todoInput.trim()) return;
    const newTodo: Todo = {
      id: crypto.randomUUID(),
      text: todoInput.trim(),
      completed: false,
      createdAt: Date.now(),
    };
    setTodos([newTodo, ...todos]);
    setTodoInput('');
  };

  const toggleTodo = (id: string) => {
    setTodos(todos.map(t => t.id === id ? { ...t, completed: !t.completed } : t));
  };

  const deleteTodo = (id: string) => {
    setTodos(todos.filter(t => t.id !== id));
  };

  // Keyword Handlers
  const addKeyword = (e: React.FormEvent) => {
    e.preventDefault();
    if (!keywordInput.trim()) return;
    if (keywords.some(k => k.word === keywordInput.trim())) return;
    
    const newKeyword: NewsKeyword = {
      id: crypto.randomUUID(),
      word: keywordInput.trim(),
    };
    setKeywords([...keywords, newKeyword]);
    setKeywordInput('');
  };

  const removeKeyword = (id: string) => {
    setKeywords(keywords.filter(k => k.id !== id));
  };

  const openGoogleNews = (word: string) => {
    window.open(`https://www.google.com/search?q=${encodeURIComponent(word)}&tbm=nws`, '_blank');
  };

  const openGoogleWeather = () => {
    window.open('https://www.google.com/search?q=weather', '_blank');
  };

  // Excel Handlers
  const exportToExcel = () => {
    const data = todos.map(t => ({
      '내용': t.text,
      '상태': t.completed ? '완료' : '진행중',
      '생성일': new Date(t.createdAt).toLocaleString()
    }));
    
    const worksheet = XLSX.utils.json_to_sheet(data);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, '할일목록');
    XLSX.writeFile(workbook, `할일목록_${new Date().toISOString().split('T')[0]}.xlsx`);
  };

  const handleImportClick = () => {
    fileInputRef.current?.click();
  };

  const importFromExcel = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      const bstr = evt.target?.result;
      const wb = XLSX.read(bstr, { type: 'binary' });
      const wsname = wb.SheetNames[0];
      const ws = wb.Sheets[wsname];
      const data = XLSX.utils.sheet_to_json(ws) as any[];

      const importedTodos: Todo[] = data.map(item => ({
        id: crypto.randomUUID(),
        text: item['내용'] || item['Content'] || '내용 없음',
        completed: (item['상태'] || item['Status']) === '완료',
        createdAt: Date.now()
      }));

      setTodos([...importedTodos, ...todos]);
      if (e.target) e.target.value = '';
    };
    reader.readAsBinaryString(file);
  };

  return (
    <div className="min-h-screen bg-bg-dark text-text-dark font-sans p-6 md:p-10">
      <div className="max-w-6xl mx-auto space-y-6">
        
        {/* Header Section */}
        <header className="flex flex-col md:flex-row md:items-center justify-between gap-4 pb-6 border-bottom border-border-dark">
          <div>
            <h1 className="text-3xl font-bold tracking-tight text-text-dark">QuickDash One-Click</h1>
            <p className="text-muted-dark text-sm mt-1">
              {new Date().toLocaleDateString('ko-KR', { year: 'numeric', month: 'long', day: 'numeric', weekday: 'long' })}
            </p>
          </div>
          
          <div className="flex items-center gap-3">
            {/* Weather Widget as defined in theme */}
            <button 
              onClick={openGoogleWeather}
              className="group flex flex-col items-center justify-center bg-card-dark px-6 py-4 rounded-2xl border border-border-dark hover:border-muted-dark transition-all active:scale-95 cursor-pointer text-center"
              id="weather-btn"
            >
              <CloudSun className="w-8 h-8 text-amber-400 mb-1" />
              <span className="font-bold text-sm text-text-dark">Google 날씨</span>
              <span className="text-[11px] text-muted-dark">실시간 기상 뉴스 확인</span>
            </button>
          </div>
        </header>

        {/* Main Content Grid: 1fr 340px layout pattern */}
        <div className="grid grid-cols-1 md:grid-cols-[1fr_340px] gap-6">
          
          {/* Left Section: To-Do (Primary) */}
          <section className="space-y-4">
            <div className="bg-card-dark p-6 md:p-8 rounded-2xl border border-border-dark min-h-[500px] flex flex-col h-full">
              <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 mb-6">
                <div className="flex items-center gap-2">
                  <h2 className="text-xl font-bold text-text-dark">오늘의 할 일</h2>
                  <span className="bg-bg-dark text-muted-dark text-xs font-bold px-2 py-0.5 rounded-full border border-border-dark">
                    {todos.length}
                  </span>
                </div>
                
                <div className="flex items-center gap-2">
                  <button 
                    onClick={handleImportClick}
                    className="flex items-center gap-2 px-3 py-1.5 text-xs font-semibold text-muted-dark hover:text-text-dark border border-border-dark hover:border-muted-dark rounded-lg transition-all cursor-pointer bg-bg-dark"
                  >
                    <FileUp className="w-3.5 h-3.5" />
                    <span>불러오기</span>
                  </button>
                  <input 
                    type="file" 
                    ref={fileInputRef} 
                    onChange={importFromExcel} 
                    className="hidden" 
                    accept=".xlsx, .xls"
                  />
                  <button 
                    onClick={exportToExcel}
                    className="flex items-center gap-2 px-3 py-1.5 text-xs font-semibold text-muted-dark hover:text-text-dark border border-border-dark hover:border-muted-dark rounded-lg transition-all cursor-pointer bg-bg-dark"
                  >
                    <Download className="w-3.5 h-3.5" />
                    <span>내보내기</span>
                  </button>
                </div>
              </div>

              <form onSubmit={addTodo} className="flex gap-2 mb-8">
                <input
                  type="text"
                  value={todoInput}
                  onChange={(e) => setTodoInput(e.target.value)}
                  placeholder="할 일을 입력하고 Enter를 누르세요..."
                  className="flex-1 bg-bg-dark border border-border-dark focus:border-accent-dark rounded-lg px-4 py-3 text-sm text-text-dark outline-none transition-all"
                  id="todo-input"
                />
                <button 
                  type="submit"
                  className="bg-accent-dark text-white px-5 py-3 rounded-lg hover:bg-accent-hover-dark font-bold text-sm transition-all shadow-sm cursor-pointer whitespace-nowrap"
                >
                  추가
                </button>
              </form>

              <div className="flex-1 space-y-2 overflow-y-auto pr-2 custom-scrollbar">
                <AnimatePresence initial={false}>
                  {todos.length === 0 ? (
                    <div className="flex flex-col items-center justify-center py-20 text-muted-dark/40">
                      <CheckCircle2 className="w-12 h-12 mb-2" />
                      <p className="text-sm">목록이 비어 있습니다.</p>
                    </div>
                  ) : (
                    todos.map((todo) => (
                      <motion.div
                        key={todo.id}
                        initial={{ opacity: 0, y: 5 }}
                        animate={{ opacity: 1, y: 0 }}
                        exit={{ opacity: 0, x: -10 }}
                        className={`group flex items-center gap-3 p-3.5 rounded-xl border transition-all bg-bg-dark border-border-dark ${
                          todo.completed ? 'opacity-50' : 'hover:border-muted-dark'
                        }`}
                      >
                        <button 
                          onClick={() => toggleTodo(todo.id)}
                          className="flex-shrink-0 cursor-pointer"
                        >
                          {todo.completed ? (
                            <CheckCircle2 className="w-5 h-5 text-success-dark" />
                          ) : (
                            <div className="w-5 h-5 rounded-sm border border-border-dark group-hover:border-muted-dark transition-colors" />
                          )}
                        </button>
                        
                        <span className={`flex-1 text-sm transition-all ${
                          todo.completed ? 'text-muted-dark line-through' : 'text-text-dark font-medium'
                        }`}>
                          {todo.text}
                        </span>

                        <button 
                          onClick={() => deleteTodo(todo.id)}
                          className="opacity-0 group-hover:opacity-100 p-1.5 text-red-400 hover:text-red-500 hover:bg-red-500/10 rounded-md transition-all cursor-pointer"
                        >
                          <Trash2 className="w-4 h-4" />
                        </button>
                      </motion.div>
                    ))
                  )}
                </AnimatePresence>
              </div>
            </div>
          </section>

          {/* Right Section: News Panel (Sidebar) */}
          <section className="space-y-6">
            <div className="bg-card-dark p-6 rounded-2xl border border-border-dark flex flex-col h-full">
              <div className="flex items-center gap-2 mb-6">
                <h2 className="text-lg font-bold text-text-dark">뉴스 키워드</h2>
              </div>
              
              <form onSubmit={addKeyword} className="flex gap-2 mb-6">
                <input
                  type="text"
                  value={keywordInput}
                  onChange={(e) => setKeywordInput(e.target.value)}
                  placeholder="관심 키워드..."
                  className="flex-1 bg-bg-dark border border-border-dark focus:border-accent-dark rounded-lg px-3 py-2.5 text-xs text-text-dark outline-none transition-all"
                  id="keyword-input"
                />
                <button 
                  type="submit"
                  className="bg-accent-dark text-white px-3 py-2 rounded-lg hover:bg-accent-hover-dark font-bold text-xs transition-all cursor-pointer"
                >
                  +
                </button>
              </form>

              <div className="flex flex-wrap gap-2 mb-6">
                <AnimatePresence>
                  {keywords.length === 0 && (
                    <p className="text-muted-dark text-[11px] italic">등록된 키워드가 없습니다.</p>
                  )}
                  {keywords.map(k => (
                    <motion.div
                      key={k.id}
                      initial={{ opacity: 0, scale: 0.95 }}
                      animate={{ opacity: 1, scale: 1 }}
                      exit={{ opacity: 0, scale: 0.9 }}
                      className="group flex items-center gap-1.5 bg-bg-dark hover:bg-border-dark px-3 py-1.5 border border-border-dark rounded-full cursor-pointer transition-colors"
                      onClick={() => openGoogleNews(k.word)}
                    >
                      <span className="text-xs font-medium text-accent-dark">{k.word}</span>
                      <button 
                        onClick={(e) => {
                          e.stopPropagation();
                          removeKeyword(k.id);
                        }}
                        className="p-0.5 rounded-full hover:bg-bg-dark text-muted-dark hover:text-red-400 transition-all ml-1"
                      >
                        <X className="w-3 h-3" />
                      </button>
                    </motion.div>
                  ))}
                </AnimatePresence>
              </div>

              <div className="mt-auto pt-6 border-t border-border-dark">
                <div className="bg-bg-dark border border-border-dark p-4 rounded-xl">
                  <h3 className="text-xs font-bold text-text-dark mb-1">스마트 팁!</h3>
                  <p className="text-muted-dark text-[11px] leading-relaxed">
                    키워드를 클릭하면 즉시 구글 뉴스로 이동합니다. 브라우저 설정에 상관없이 닫았다가 다시 열어도 그대로 유지됩니다.
                  </p>
                </div>
              </div>
            </div>
          </section>

        </div>
        
        <footer className="text-center py-8 text-muted-dark text-[11px] font-medium tracking-tight">
          <p>© 2026 DashTool. Elegant Dark Edition.</p>
        </footer>
      </div>
    </div>
  );
}
