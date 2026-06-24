import React, { useState, useRef } from 'react';
import { motion, AnimatePresence } from 'motion/react';
import { 
  Lock, 
  User, 
  UploadCloud, 
  FileSpreadsheet, 
  Download, 
  CheckCircle2, 
  LogOut, 
  Building2, 
  Receipt, 
  TrendingUp, 
  Copy, 
  Check, 
  FileText, 
  ChevronRight, 
  ChevronDown, 
  AlertTriangle,
  Info
} from 'lucide-react';
import { processExcelData, formatProcessedOutput } from './utils/excelProcessor';
import { exportToWord, exportToPDF } from './utils/reportExporter';
import { ProcessedData, UserSession } from './types';

export default function App() {
  // Session State
  const [session, setSession] = useState<UserSession>({
    username: '',
    isAuthenticated: false,
  });
  
  // Login Form States
  const [loginUser, setLoginUser] = useState('');
  const [loginPass, setLoginPass] = useState('');
  const [loginError, setLoginError] = useState('');
  
  // File Uploader & Processing States
  const [dragActive, setDragActive] = useState(false);
  const [fileName, setFileName] = useState('');
  const [fileSize, setFileSize] = useState('');
  const [isProcessing, setIsProcessing] = useState(false);
  const [processedData, setProcessedData] = useState<ProcessedData | null>(null);
  const [copied, setCopied] = useState(false);
  const [processingError, setProcessingError] = useState('');
  
  // Interactive Accordion States (Stores expanded/collapsed)
  const [expandedStores, setExpandedStores] = useState<Record<string, boolean>>({});

  const fileInputRef = useRef<HTMLInputElement>(null);

  // Default Login Handler
  const handleLogin = (e: React.FormEvent) => {
    e.preventDefault();
    if (!loginUser.trim() || !loginPass.trim()) {
      setLoginError('Por favor, preencha todos os campos.');
      return;
    }
    
    // Validando credenciais (admin / admin123)
    if (loginUser.toLowerCase() === 'admin' && loginPass === 'admin123') {
      setSession({
        username: 'Administrador',
        isAuthenticated: true,
      });
      setLoginError('');
    } else {
      setLoginError('Usuário ou senha incorretos. (Dica: admin / admin123)');
    }
  };

  const handleLogout = () => {
    setSession({ username: '', isAuthenticated: false });
    setProcessedData(null);
    setFileName('');
    setFileSize('');
    setLoginUser('');
    setLoginPass('');
    setProcessingError('');
  };

  // Drag & Drop Handlers
  const handleDrag = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    if (e.type === "dragenter" || e.type === "dragover") {
      setDragActive(true);
    } else if (e.type === "dragleave") {
      setDragActive(false);
    }
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setDragActive(false);

    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      const file = e.dataTransfer.files[0];
      validateAndProcessFile(file);
    }
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
      validateAndProcessFile(e.target.files[0]);
    }
  };

  const validateAndProcessFile = (file: File) => {
    const extension = file.name.split('.').pop()?.toLowerCase();
    if (extension !== 'xlsx' && extension !== 'xls') {
      setProcessingError('Por favor, faça o upload de um arquivo válido do Excel (.xlsx ou .xls)');
      setProcessedData(null);
      return;
    }

    setProcessingError('');
    setFileName(file.name);
    // Formatar tamanho do arquivo
    const sizeInKB = (file.size / 1024).toFixed(1);
    setFileSize(`${sizeInKB} KB`);
    setIsProcessing(true);

    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const buffer = e.target?.result as ArrayBuffer;
        const results = processExcelData(buffer);
        setProcessedData(results);
        
        // Expandir todas as lojas por padrão ao carregar
        const initialExpandState: Record<string, boolean> = {};
        Object.keys(results).forEach((loja) => {
          initialExpandState[loja] = true;
        });
        setExpandedStores(initialExpandState);
      } catch (err) {
        console.error(err);
        setProcessingError('Ocorreu um erro ao processar o arquivo. Verifique se o formato está correto.');
      } finally {
        setIsProcessing(false);
      }
    };
    reader.onerror = () => {
      setProcessingError('Erro ao ler o arquivo.');
      setIsProcessing(false);
    };
    reader.readAsArrayBuffer(file);
  };

  const triggerFileInput = () => {
    fileInputRef.current?.click();
  };

  const toggleStore = (loja: string) => {
    setExpandedStores(prev => ({
      ...prev,
      [loja]: !prev[loja]
    }));
  };

  // Métricas auxiliares para o Dashboard
  const getStats = () => {
    if (!processedData) return { lojasCount: 0, notasCount: 0, produtosCount: 0 };
    
    const lojas = Object.keys(processedData);
    const lojasCount = lojas.length;
    let notasCount = 0;
    let produtosCount = 0;
    
    for (const loja of lojas) {
      const notasMap = processedData[loja];
      const notas = Object.keys(notasMap);
      notasCount += notas.length;
      
      for (const nota of notas) {
        const fornecedoresMap = notasMap[nota];
        const fornecedores = Object.keys(fornecedoresMap);
        
        for (const fornecedor of fornecedores) {
          const produtos = fornecedoresMap[fornecedor];
          produtosCount += produtos.length;
        }
      }
    }
    
    return { lojasCount, notasCount, produtosCount };
  };

  const stats = getStats();
  const rawMarkdown = processedData ? formatProcessedOutput(processedData) : '';

  const copyToClipboard = () => {
    navigator.clipboard.writeText(rawMarkdown);
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  };

  return (
    <div id="app-container" className="min-h-screen bg-slate-50 text-slate-800 font-sans selection:bg-blue-100 antialiased">
      <AnimatePresence mode="wait">
        
        {/* TELA DE LOGIN */}
        {!session.isAuthenticated ? (
          <motion.div 
            key="login"
            initial={{ opacity: 0, y: 15 }}
            animate={{ opacity: 1, y: 0 }}
            exit={{ opacity: 0, y: -15 }}
            transition={{ duration: 0.3, ease: 'easeOut' }}
            className="flex min-h-screen items-center justify-center px-4 py-12 sm:px-6 lg:px-8 bg-gradient-to-tr from-slate-100 to-slate-200"
            id="login-screen"
          >
            <div className="w-full max-w-md space-y-8 bg-white p-8 rounded-2xl shadow-xl border border-slate-100">
              <div className="text-center">
                <div className="mx-auto flex h-14 w-14 items-center justify-center rounded-2xl bg-blue-50 text-blue-600 mb-4">
                  <Lock className="h-7 w-7" id="lock-icon" />
                </div>
                <h2 className="text-3xl font-bold tracking-tight text-slate-900 font-sans">
                  Controle de Custos
                </h2>
                <p className="mt-2 text-sm text-slate-500">
                  Faça login para analisar planilhas e gerar relatórios
                </p>
              </div>

              <form className="mt-8 space-y-6" onSubmit={handleLogin} id="login-form">
                {loginError && (
                  <motion.div 
                    initial={{ opacity: 0, scale: 0.95 }}
                    animate={{ opacity: 1, scale: 1 }}
                    className="p-3.5 text-sm text-red-700 bg-red-50 border border-red-100 rounded-lg flex items-center gap-2"
                    id="login-error-alert"
                  >
                    <AlertTriangle className="h-4.5 w-4.5 shrink-0" />
                    <span>{loginError}</span>
                  </motion.div>
                )}

                <div className="space-y-4 rounded-md">
                  <div>
                    <label htmlFor="username" className="block text-sm font-medium text-slate-700 mb-1.5">
                      Usuário
                    </label>
                    <div className="relative">
                      <span className="absolute inset-y-0 left-0 flex items-center pl-3 text-slate-400">
                        <User className="h-5 w-5" />
                      </span>
                      <input
                        id="username"
                        name="username"
                        type="text"
                        autoComplete="username"
                        value={loginUser}
                        onChange={(e) => setLoginUser(e.target.value)}
                        className="block w-full rounded-xl border border-slate-200 py-3 pl-10 pr-3 text-slate-900 placeholder:text-slate-400 focus:border-blue-500 focus:outline-none focus:ring-2 focus:ring-blue-100 text-sm transition-all"
                        placeholder="admin"
                      />
                    </div>
                  </div>

                  <div>
                    <label htmlFor="password" className="block text-sm font-medium text-slate-700 mb-1.5">
                      Senha
                    </label>
                    <div className="relative">
                      <span className="absolute inset-y-0 left-0 flex items-center pl-3 text-slate-400">
                        <Lock className="h-5 w-5" />
                      </span>
                      <input
                        id="password"
                        name="password"
                        type="password"
                        autoComplete="current-password"
                        value={loginPass}
                        onChange={(e) => setLoginPass(e.target.value)}
                        className="block w-full rounded-xl border border-slate-200 py-3 pl-10 pr-3 text-slate-900 placeholder:text-slate-400 focus:border-blue-500 focus:outline-none focus:ring-2 focus:ring-blue-100 text-sm transition-all"
                        placeholder="••••••••"
                      />
                    </div>
                  </div>
                </div>

                <div className="bg-amber-50 rounded-xl p-4 border border-amber-100 text-xs text-amber-800 space-y-1">
                  <p className="font-semibold flex items-center gap-1.5">
                    <Info className="h-3.5 w-3.5" /> Credenciais para acesso:
                  </p>
                  <p>Usuário: <code className="bg-amber-100 px-1 py-0.5 rounded font-mono font-medium">admin</code></p>
                  <p>Senha: <code className="bg-amber-100 px-1 py-0.5 rounded font-mono font-medium">admin123</code></p>
                </div>

                <div>
                  <button
                    type="submit"
                    id="btn-login-submit"
                    className="group relative flex w-full justify-center rounded-xl bg-blue-600 px-4 py-3 text-sm font-semibold text-white hover:bg-blue-700 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:ring-offset-2 shadow-lg shadow-blue-500/20 transition-all cursor-pointer"
                  >
                    Entrar no Sistema
                  </button>
                </div>
              </form>
            </div>
          </motion.div>
        ) : (
          
          /* TELA ANÁLISE DE CUSTOS */
          <motion.div 
            key="dashboard"
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            className="flex flex-col min-h-screen"
            id="cost-analysis-screen"
          >
            {/* Header da Aplicação */}
            <header className="sticky top-0 z-10 bg-white border-b border-slate-100 shadow-sm">
              <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 h-16 flex items-center justify-between">
                <div className="flex items-center gap-2.5">
                  <div className="h-9 w-9 rounded-lg bg-blue-600 flex items-center justify-center text-white font-bold">
                    <FileSpreadsheet className="h-5 w-5" />
                  </div>
                  <div>
                    <h1 className="text-lg font-bold text-slate-900 tracking-tight">Análise de Custos</h1>
                    <p className="text-xs text-slate-400 font-medium">Desenvolvido por Manus AI</p>
                  </div>
                </div>

                <div className="flex items-center gap-4">
                  <div className="hidden md:flex items-center gap-2 px-3 py-1.5 bg-slate-100 rounded-full text-xs font-medium text-slate-600">
                    <span className="w-2 h-2 rounded-full bg-green-500 animate-pulse"></span>
                    {session.username}
                  </div>
                  <button
                    onClick={handleLogout}
                    id="btn-logout"
                    className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg border border-slate-200 text-xs font-medium text-slate-600 hover:text-red-600 hover:bg-red-50 hover:border-red-100 transition-all cursor-pointer"
                    title="Sair do sistema"
                  >
                    <LogOut className="h-4 w-4" />
                    <span>Sair</span>
                  </button>
                </div>
              </div>
            </header>

            {/* Conteúdo Principal */}
            <main className="flex-1 max-w-7xl w-full mx-auto px-4 sm:px-6 lg:px-8 py-8 space-y-8">
              
              {/* Título da Seção */}
              <div className="space-y-1">
                <h2 className="text-2xl font-bold tracking-tight text-slate-950">Processador de Relatório de Custos Excel</h2>
                <p className="text-sm text-slate-500">
                  Faça o upload do seu arquivo Excel para filtrar variações de custo superiores a 25% por loja, nota e fornecedor.
                </p>
              </div>

              {/* Área de Upload */}
              <section className="bg-white p-6 rounded-2xl border border-slate-100 shadow-sm space-y-4">
                <div 
                  className={`relative border-2 border-dashed rounded-xl p-8 text-center transition-all cursor-pointer ${
                    dragActive 
                      ? "border-blue-500 bg-blue-50/40" 
                      : "border-slate-200 hover:border-blue-400 hover:bg-slate-50/50"
                  }`}
                  onDragEnter={handleDrag}
                  onDragOver={handleDrag}
                  onDragLeave={handleDrag}
                  onDrop={handleDrop}
                  onClick={triggerFileInput}
                  id="drop-zone"
                >
                  <input 
                    ref={fileInputRef}
                    type="file" 
                    id="excel-file-input"
                    className="hidden" 
                    accept=".xlsx, .xls"
                    onChange={handleFileChange}
                  />

                  <div className="flex flex-col items-center space-y-4">
                    <div className="p-4 rounded-full bg-blue-50 text-blue-600">
                      <UploadCloud className="h-8 w-8" />
                    </div>
                    <div>
                      <p className="font-semibold text-slate-800">
                        Clique para buscar ou arraste seu arquivo Excel aqui
                      </p>
                      <p className="text-xs text-slate-400 mt-1">
                        Suporta apenas planilhas Excel (.xlsx, .xls)
                      </p>
                    </div>
                  </div>
                </div>

                {processingError && (
                  <div className="p-3.5 bg-red-50 border border-red-100 text-red-700 text-sm rounded-xl flex items-center gap-2">
                    <AlertTriangle className="h-4.5 w-4.5 shrink-0" />
                    <span>{processingError}</span>
                  </div>
                )}

                {fileName && !processingError && (
                  <div className="p-4 bg-blue-50/50 border border-blue-100 rounded-xl flex flex-col sm:flex-row sm:items-center sm:justify-between gap-3">
                    <div className="flex items-center gap-3">
                      <div className="h-10 w-10 bg-blue-100 rounded-lg flex items-center justify-center text-blue-600 shrink-0">
                        <FileSpreadsheet className="h-5.5 w-5.5" />
                      </div>
                      <div className="min-w-0">
                        <p className="text-sm font-semibold text-slate-900 truncate">{fileName}</p>
                        <p className="text-xs text-slate-500">Tamanho: {fileSize}</p>
                      </div>
                    </div>
                    <div className="flex items-center gap-1 text-green-700 text-xs font-semibold bg-green-50 px-2.5 py-1 rounded-full border border-green-100 shrink-0 w-fit">
                      <CheckCircle2 className="h-3.5 w-3.5" />
                      <span>Arquivo pronto</span>
                    </div>
                  </div>
                )}
              </section>

              {/* Loader */}
              {isProcessing && (
                <div className="flex flex-col items-center justify-center py-12 space-y-4">
                  <div className="w-12 h-12 border-4 border-blue-200 border-t-blue-600 rounded-full animate-spin"></div>
                  <p className="text-sm font-medium text-slate-500">Processando e interpretando os dados da planilha...</p>
                </div>
              )}

              {/* RESULTADOS DA ANÁLISE */}
              {processedData && !isProcessing && (
                <motion.div 
                  initial={{ opacity: 0, y: 10 }}
                  animate={{ opacity: 1, y: 0 }}
                  transition={{ duration: 0.3 }}
                  className="space-y-6"
                >
                  
                  {/* Dashboard de Métricas Rápidas */}
                  <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
                    <div className="bg-white p-5 rounded-xl border border-slate-100 shadow-sm flex items-center gap-4">
                      <div className="p-3 bg-indigo-50 rounded-lg text-indigo-600">
                        <Building2 className="h-6 w-6" />
                      </div>
                      <div>
                        <p className="text-xs font-medium text-slate-400 uppercase tracking-wider">Lojas Identificadas</p>
                        <p className="text-2xl font-bold text-slate-900 mt-0.5">{stats.lojasCount}</p>
                      </div>
                    </div>

                    <div className="bg-white p-5 rounded-xl border border-slate-100 shadow-sm flex items-center gap-4">
                      <div className="p-3 bg-violet-50 rounded-lg text-violet-600">
                        <Receipt className="h-6 w-6" />
                      </div>
                      <div>
                        <p className="text-xs font-medium text-slate-400 uppercase tracking-wider">Notas Fiscais</p>
                        <p className="text-2xl font-bold text-slate-900 mt-0.5">{stats.notasCount}</p>
                      </div>
                    </div>

                    <div className="bg-white p-5 rounded-xl border border-slate-100 shadow-sm flex items-center gap-4">
                      <div className="p-3 bg-rose-50 rounded-lg text-rose-600">
                        <TrendingUp className="h-6 w-6" />
                      </div>
                      <div>
                        <p className="text-xs font-medium text-slate-400 uppercase tracking-wider">Produtos com Variação &gt; 25%</p>
                        <p className="text-2xl font-bold text-slate-900 mt-0.5">{stats.produtosCount}</p>
                      </div>
                    </div>
                  </div>

                  {/* Painel de Visualização e Downloads */}
                  <div className="bg-white rounded-2xl border border-slate-100 shadow-sm overflow-hidden">
                    
                    {/* Header do Painel com Ações de Exportação */}
                    <div className="bg-slate-50 border-b border-slate-100 px-6 py-4 flex flex-col md:flex-row md:items-center md:justify-between gap-4">
                      <div>
                        <h3 className="font-bold text-slate-900 text-base">Resultados Filtrados</h3>
                        <p className="text-xs text-slate-500 mt-0.5">Clique nas lojas para detalhar ou selecione uma opção de exportação abaixo</p>
                      </div>

                      <div className="flex flex-wrap items-center gap-2">
                        {/* Download Word */}
                        <button
                          onClick={() => exportToWord(processedData)}
                          id="btn-download-word"
                          className="flex items-center gap-1.5 px-3 py-2 bg-blue-600 text-white rounded-lg text-xs font-semibold hover:bg-blue-700 transition-all cursor-pointer shadow-sm shadow-blue-500/10"
                        >
                          <Download className="h-3.5 w-3.5" />
                          <span>Baixar Word (.docx)</span>
                        </button>

                        {/* Download PDF */}
                        <button
                          onClick={() => exportToPDF(processedData)}
                          id="btn-download-pdf"
                          className="flex items-center gap-1.5 px-3 py-2 bg-slate-900 text-white rounded-lg text-xs font-semibold hover:bg-slate-800 transition-all cursor-pointer shadow-sm shadow-slate-900/10"
                        >
                          <FileText className="h-3.5 w-3.5" />
                          <span>Baixar PDF (.pdf)</span>
                        </button>

                        {/* Copiar Markdown */}
                        <button
                          onClick={copyToClipboard}
                          id="btn-copy-markdown"
                          className="flex items-center gap-1.5 px-3 py-2 border border-slate-200 text-slate-700 rounded-lg text-xs font-semibold hover:bg-slate-100 transition-all cursor-pointer"
                        >
                          {copied ? (
                            <>
                              <Check className="h-3.5 w-3.5 text-green-600" />
                              <span className="text-green-600">Copiado!</span>
                            </>
                          ) : (
                            <>
                              <Copy className="h-3.5 w-3.5" />
                              <span>Copiar Markdown</span>
                            </>
                          )}
                        </button>
                      </div>
                    </div>

                    {/* Exibição Interativa do Relatório */}
                    <div className="p-6 space-y-6">
                      {stats.produtosCount === 0 ? (
                        <div className="text-center py-12 space-y-3">
                          <Info className="mx-auto h-12 w-12 text-slate-300" />
                          <h4 className="font-bold text-slate-800">Nenhum registro crítico encontrado</h4>
                          <p className="text-sm text-slate-400 max-w-sm mx-auto">
                            O arquivo Excel foi processado perfeitamente, mas nenhum produto apresentou variação de custo superior a 25%.
                          </p>
                        </div>
                      ) : (
                        <div className="space-y-4" id="report-view-container">
                          {Object.keys(processedData).map((lojaName) => {
                            const isExpanded = expandedStores[lojaName] !== false;
                            const notas = processedData[lojaName];
                            
                            // Contagem de produtos desta loja
                            let storeProdsCount = 0;
                            Object.values(notas).forEach((forns) => {
                              Object.values(forns).forEach((prods) => {
                                storeProdsCount += prods.length;
                              });
                            });

                            return (
                              <div key={lojaName} className="border border-slate-100 rounded-xl overflow-hidden shadow-sm">
                                <button
                                  onClick={() => toggleStore(lojaName)}
                                  className="w-full flex items-center justify-between p-4 bg-slate-50/50 hover:bg-slate-50 transition-all text-left"
                                >
                                  <div className="flex items-center gap-3">
                                    <div className="p-1.5 bg-blue-50 text-blue-600 rounded">
                                      <Building2 className="h-4 w-4" />
                                    </div>
                                    <div>
                                      <span className="font-bold text-slate-900 text-sm md:text-base">{lojaName}</span>
                                      <span className="ml-2.5 text-xs text-slate-400 font-medium">({storeProdsCount} itens críticos)</span>
                                    </div>
                                  </div>
                                  <div>
                                    {isExpanded ? (
                                      <ChevronDown className="h-5 w-5 text-slate-400" />
                                    ) : (
                                      <ChevronRight className="h-5 w-5 text-slate-400" />
                                    )}
                                  </div>
                                </button>

                                <AnimatePresence initial={false}>
                                  {isExpanded && (
                                    <motion.div
                                      initial={{ height: 0, opacity: 0 }}
                                      animate={{ height: "auto", opacity: 1 }}
                                      exit={{ height: 0, opacity: 0 }}
                                      transition={{ duration: 0.2 }}
                                      className="overflow-hidden bg-white border-t border-slate-100"
                                    >
                                      <div className="p-4 space-y-6">
                                        {Object.keys(notas).sort((a, b) => parseInt(a, 10) - parseInt(b, 10)).map((notaNum) => {
                                          const fornecedores = notas[notaNum];
                                          return (
                                            <div key={notaNum} className="space-y-4">
                                              {Object.keys(fornecedores).sort().map((fornecedorName) => {
                                                const produtos = fornecedores[fornecedorName];
                                                return (
                                                  <div key={fornecedorName} className="space-y-2.5">
                                                    {/* Tag Nota e Fornecedor */}
                                                    <div className="flex flex-wrap items-center gap-2 text-xs font-semibold text-slate-600 bg-slate-100/75 py-1.5 px-3 rounded-lg w-fit border border-slate-200/50">
                                                      <span className="text-blue-700">NOTA: {notaNum}</span>
                                                      <span className="text-slate-300">|</span>
                                                      <span className="text-slate-700">FORNECEDOR: {fornecedorName}</span>
                                                    </div>

                                                    {/* Lista de Produtos */}
                                                    <div className="grid grid-cols-1 gap-2 pl-2">
                                                      {produtos.map((prod, idx) => {
                                                        const isPriceIncrease = prod.variacao > 0;
                                                        return (
                                                          <div 
                                                            key={`${prod.codigo}-${idx}`}
                                                            className="flex flex-col md:flex-row md:items-center justify-between p-3.5 bg-slate-50/50 hover:bg-slate-50 rounded-xl border border-slate-100 transition-all gap-3"
                                                          >
                                                            <div className="flex items-start gap-3 min-w-0">
                                                              <span className="text-xs font-mono font-bold bg-slate-200 text-slate-700 px-2 py-0.5 rounded mt-0.5">
                                                                {prod.codigo}
                                                              </span>
                                                              <span className="text-sm font-semibold text-slate-800 truncate">
                                                                {prod.nome}
                                                              </span>
                                                            </div>

                                                            <div className="flex flex-wrap items-center justify-between md:justify-end gap-4 md:gap-6 shrink-0">
                                                              <div className="flex items-center gap-4 text-xs">
                                                                <div>
                                                                  <p className="text-slate-400 font-medium">Anterior</p>
                                                                  <p className="font-semibold text-slate-700">R$ {prod.custo_anter.toFixed(2)}</p>
                                                                </div>
                                                                <ChevronRight className="h-4 w-4 text-slate-300 hidden sm:block" />
                                                                <div>
                                                                  <p className="text-slate-400 font-medium">Atual</p>
                                                                  <p className="font-bold text-slate-950">R$ {prod.custo_atual.toFixed(2)}</p>
                                                                </div>
                                                              </div>

                                                              <div className={`px-2.5 py-1 rounded-full text-xs font-bold shrink-0 ${
                                                                isPriceIncrease 
                                                                  ? "bg-red-50 text-red-700 border border-red-100" 
                                                                  : "bg-green-50 text-green-700 border border-green-100"
                                                              }`}>
                                                                {isPriceIncrease ? "+" : ""}{prod.variacao.toFixed(2)}%
                                                              </div>
                                                            </div>
                                                          </div>
                                                        );
                                                      })}
                                                    </div>
                                                  </div>
                                                );
                                              })}
                                            </div>
                                          );
                                        })}
                                      </div>
                                    </motion.div>
                                  )}
                                </AnimatePresence>
                              </div>
                            );
                          })}
                        </div>
                      )}
                    </div>
                  </div>

                  {/* Aba Informativa Adicional - Visualizador Markdown */}
                  <div className="bg-white rounded-2xl border border-slate-100 shadow-sm p-6 space-y-4">
                    <div className="flex items-center justify-between">
                      <div>
                        <h3 className="font-bold text-slate-900 text-base">Visualização do Relatório Bruto (Markdown)</h3>
                        <p className="text-xs text-slate-500 mt-0.5">Formato texto copiado para a área de transferência</p>
                      </div>
                      <button
                        onClick={copyToClipboard}
                        className="p-2 border border-slate-100 text-slate-400 hover:text-blue-600 hover:bg-slate-50 rounded-lg transition-all cursor-pointer"
                        title="Copiar texto"
                      >
                        {copied ? <Check className="h-5 w-5 text-green-600" /> : <Copy className="h-5 w-5" />}
                      </button>
                    </div>

                    <div className="bg-slate-950 text-slate-200 rounded-xl p-5 font-mono text-xs overflow-x-auto max-h-72 border border-slate-900">
                      <pre className="whitespace-pre-wrap">{rawMarkdown}</pre>
                    </div>
                  </div>
                </motion.div>
              )}

              {/* Dica do Desenvolvedor no rodapé */}
              {!processedData && !isProcessing && (
                <div className="text-center py-20 max-w-md mx-auto space-y-3">
                  <div className="w-14 h-14 mx-auto rounded-full bg-slate-100 text-slate-400 flex items-center justify-center">
                    <FileSpreadsheet className="h-7 w-7" />
                  </div>
                  <h3 className="font-bold text-slate-900">Nenhum arquivo analisado</h3>
                  <p className="text-sm text-slate-500">
                    Insira uma planilha do Excel no uploader acima contendo o relatório de custos para gerar a análise automaticamente.
                  </p>
                </div>
              )}
            </main>

            {/* Rodapé da aplicação */}
            <footer className="mt-auto bg-white border-t border-slate-100 py-6">
              <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 flex flex-col sm:flex-row items-center justify-between gap-4 text-xs text-slate-400">
                <p>© 2026 Controle de Custos. Todos os direitos reservados.</p>
                <p>Desenvolvido por Manus AI</p>
              </div>
            </footer>
          </motion.div>
        )}
      </AnimatePresence>
    </div>
  );
}
