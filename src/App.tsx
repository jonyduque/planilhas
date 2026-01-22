import React, { useState, useEffect } from 'react';
import { Upload, FileSpreadsheet, Download, RefreshCw, AlertCircle, CheckCircle, Loader2 } from 'lucide-react';
import './App.css';

// Declaração de tipos globais para as bibliotecas carregadas via CDN
declare global {
  interface Window {
    XLSX: any;
    ExcelJS: any;
  }
}

// URL para a biblioteca SheetJS (XLSX) - Usada para LER (excelente compatibilidade)
const XLSX_CDN = "https://cdn.sheetjs.com/xlsx-0.20.1/package/dist/xlsx.full.min.js";

// URL para a biblioteca ExcelJS - Usada para ESCREVER (suporte a Tabelas e Estilos)
const EXCELJS_CDN = "https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.4.0/exceljs.min.js";

interface ColMap {
  processo: number;
  localizadores: number;
  inclusao: number;
  ultimoEvento: number;
}

export default function App() {
  const [isLibLoaded, setIsLibLoaded] = useState<boolean>(false);
  const [file, setFile] = useState<File | null>(null);
  const [processedData, setProcessedData] = useState<any[][] | null>(null);
  const [headers, setHeaders] = useState<string[]>([]); // Armazenar cabeçalhos separadamente para o ExcelJS
  const [isLoading, setIsLoading] = useState<boolean>(false);
  const [isDownloading, setIsDownloading] = useState<boolean>(false);
  const [error, setError] = useState<string>("");
  const [isDragging, setIsDragging] = useState<boolean>(false);

  // Carregar as bibliotecas dinamicamente
  useEffect(() => {
    const loadLibraries = async () => {
      // Função helper para carregar script
      const loadScript = (src: string, checkGlobal: keyof Window) => {
        return new Promise<void>((resolve) => {
          if (window[checkGlobal]) {
            resolve();
            return;
          }
          const script = document.createElement('script');
          script.src = src;
          script.onload = () => resolve();
          document.body.appendChild(script);
        });
      };

      await Promise.all([
        loadScript(XLSX_CDN, 'XLSX'),
        loadScript(EXCELJS_CDN, 'ExcelJS')
      ]);

      setIsLibLoaded(true);
    };

    loadLibraries();
  }, []);

  // Função auxiliar para decodificar HTML Entities
  const decodeHTMLEntities = (text: any): string => {
    if (!text || typeof text !== 'string') return String(text || "");
    const parser = new DOMParser();
    const doc = parser.parseFromString(`<!doctype html><body>${text}`, 'text/html');
    return doc.body.textContent || "";
  };

  // Função auxiliar para converter string de data PT-BR
  const parseBrazilianDate = (dateStr: any): Date | string => {
    if (!dateStr || typeof dateStr !== 'string') return dateStr;
    try {
      const [datePart, timePart] = dateStr.split(' ');
      if (!datePart) return dateStr;

      const [day, month, year] = datePart.split('/').map(Number);

      let hours = 0, minutes = 0, seconds = 0;
      if (timePart) {
        [hours, minutes, seconds] = timePart.split(':').map(Number);
      }

      if (!day || !month || !year) return dateStr;
      return new Date(year, month - 1, day, hours, minutes, seconds || 0);
    } catch (e) {
      return dateStr;
    }
  };

  const handleFileSelection = (selectedFile: File) => {
    if (selectedFile) {
      const validTypes = ['.xlsx', '.xls', '.csv'];
      const isExtensionValid = validTypes.some(ext => selectedFile.name.toLowerCase().endsWith(ext));

      if (!isExtensionValid) {
          setError("Por favor, envie apenas arquivos Excel (.xlsx, .xls) ou CSV.");
          return;
      }

      setFile(selectedFile);
      setError("");
      setProcessedData(null);
      setHeaders([]);
    }
  };

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files.length > 0) {
      handleFileSelection(e.target.files[0]);
    }
  };

  const handleDragOver = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    setIsDragging(true);
  };

  const handleDragLeave = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();

    // Evita que o estado de "dragging" seja removido ao passar por cima dos filhos (ícone, texto)
    if (e.relatedTarget && e.currentTarget.contains(e.relatedTarget as Node)) {
      return;
    }

    setIsDragging(false);
  };

  const handleDrop = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    setIsDragging(false);
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      handleFileSelection(e.dataTransfer.files[0]);
      const fileInput = document.getElementById('fileInput') as HTMLInputElement;
      if (fileInput) fileInput.value = '';
    }
  };

  const processFile = async () => {
    if (!file || !window.XLSX) return;
    setIsLoading(true);
    setError("");

    const reader = new FileReader();

    reader.onload = (e: ProgressEvent<FileReader>) => {
      try {
        const data = e.target?.result;
        if (!data) throw new Error("Erro ao ler dados do arquivo.");

        const workbook = window.XLSX.read(data, { type: 'binary' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        // Forçando o tipo any[] aqui pois sheet_to_json retorna array de objetos ou arrays dependendo da config
        const jsonData: any[][] = window.XLSX.utils.sheet_to_json(worksheet, { header: 1 });

        if (jsonData.length < 2) throw new Error("Arquivo vazio ou sem linhas suficientes.");

        // 1. Excluir a primeira linha
        const cleanData = jsonData.slice(1);
        if (cleanData.length === 0) throw new Error("Não há dados após remover cabeçalho.");

        const currentHeaders: any[] = cleanData[0];
        const rows = cleanData.slice(1);

        const colMap: ColMap = {
          processo: currentHeaders.findIndex((h: any) => h && h.toString().toLowerCase().includes("número processo")),
          localizadores: currentHeaders.findIndex((h: any) => h && h.toString().toLowerCase().includes("localizadores")),
          inclusao: currentHeaders.findIndex((h: any) => h && h.toString().toLowerCase().includes("inclusão no localizador")),
          ultimoEvento: currentHeaders.findIndex((h: any) => h && h.toString().toLowerCase().includes("último evento")),
        };

        if (colMap.localizadores === -1) setError("Aviso: Coluna 'Localizadores' não encontrada.");

        // Adicionar novo cabeçalho
        currentHeaders.push("Localizadores do Gabinete");

        let gabCountTotal = 0;

        const processedRows = rows.map((row: any[]) => {
          // Garante que row é um array
          if (!Array.isArray(row)) return row;

          while (row.length < currentHeaders.length - 1) row.push("");

          // Limpar Processo
          if (colMap.processo !== -1 && row[colMap.processo]) {
            row[colMap.processo] = String(row[colMap.processo]).trim();
          }

          // Processar Localizadores
          let gCount = 0;
          if (colMap.localizadores !== -1 && row[colMap.localizadores]) {
            let locText = String(row[colMap.localizadores]);
            locText = decodeHTMLEntities(locText);

            const matches = locText.match(/\(G\)/g);
            gCount = matches ? matches.length : 0;
            gabCountTotal += gCount;

            locText = locText.replace(/\(Principal\)/gi, '').trim();

            // CORREÇÃO: Substituir " - " por quebra de linha SOMENTE se a próxima palavra for maiúscula ou emoji
            locText = locText.replace(/\s+-\s+([^\s]+)/g, (_: string, nextWord: string) => {
              const isAllUpperCaseOrSymbol = nextWord === nextWord.toUpperCase();

              if (isAllUpperCaseOrSymbol) {
                return '\n' + nextWord;
              } else {
                return ' - ' + nextWord;
              }
            });

            row[colMap.localizadores] = locText;
          }

          // Converter Data Inclusão
          if (colMap.inclusao !== -1 && row[colMap.inclusao]) {
            row[colMap.inclusao] = parseBrazilianDate(row[colMap.inclusao]);
          }

          // Converter Data Último Evento
          if (colMap.ultimoEvento !== -1 && row[colMap.ultimoEvento]) {
            row[colMap.ultimoEvento] = parseBrazilianDate(row[colMap.ultimoEvento]);
          }

          row.push(gCount);
          return row;
        });

        // Converte currentHeaders para string[] explicitamente para o estado
        setHeaders(currentHeaders.map(String));
        setProcessedData(processedRows);

      } catch (err: any) {
        console.error(err);
        setError("Erro ao processar: " + (err.message || String(err)));
      } finally {
        setIsLoading(false);
      }
    };

    reader.readAsBinaryString(file);
  };

  const downloadFile = async () => {
    if (!processedData || !window.ExcelJS) return;

    setIsDownloading(true);

    try {
      // Simular um pequeno delay para que o usuário veja o estado de carregamento
      await new Promise(resolve => setTimeout(resolve, 800));

      // Criar Workbook ExcelJS
      const workbook = new window.ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet('Dados Processados');

      // Configurar colunas (definir largura e chaves)
      const columnsConfig = headers.map(header => ({
        name: header,
        filterButton: true,
      }));

      // Adicionar Tabela do Excel (ListObject)
      worksheet.addTable({
        name: 'TabelaProcessos',
        ref: 'A1',
        headerRow: true,
        totalsRow: false,
        style: {
          theme: 'TableStyleMedium2', // Estilo azul padrão do Excel (Listrado)
          showRowStripes: true,
        },
        columns: columnsConfig,
        rows: processedData, // Dados vão direto na tabela
      });

      // AJUSTES DE ESTILO APÓS CRIAR A TABELA

      // Encontrar índices das colunas importantes
      const locIndex = headers.findIndex(h => h.toString().toLowerCase().includes("localizadores"));
      const ultimoEventoIndex = headers.findIndex(h => h.toString().toLowerCase().includes("último evento"));

      // Ajustar larguras e alinhamentos
      worksheet.columns.forEach((col: any, index: number) => {
        // Largura padrão
        col.width = 20;

        // Se for Localizadores, deixa mais largo e ativa quebra de linha
        if (index === locIndex) {
          col.width = 60;
          col.alignment = { wrapText: true, vertical: 'top', horizontal: 'left' };
        }
        // Se for Último Evento, formata como data curta
        else if (index === ultimoEventoIndex) {
          col.width = 18;
          col.numFmt = 'dd/mm/yyyy'; // Formato de data curta (pt-BR)
          col.alignment = { vertical: 'top', horizontal: 'center' };
        }
        else {
            // Alinhamento padrão para outras colunas
            col.alignment = { vertical: 'top', horizontal: 'left' };
        }
      });

      // Gerar Buffer
      const buffer = await workbook.xlsx.writeBuffer();

      // Download Blob
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url = window.URL.createObjectURL(blob);
      const anchor = document.createElement('a');
      anchor.href = url;
      anchor.download = `Planilha_Judicial_Formatada_${new Date().toISOString().slice(0,10)}.xlsx`;
      anchor.click();
      window.URL.revokeObjectURL(url);

    } catch (err: any) {
      console.error(err);
      setError("Erro ao gerar o arquivo Excel: " + (err.message || String(err)));
    } finally {
        setIsDownloading(false);
    }
  };

  return (
    <div className="app-container">
      <div className="main-card-wrapper">

        {/* Header */}
        <header className="header">
          <h1 className="header-title">
            <FileSpreadsheet className="header-icon" />
            Processador de Planilhas
          </h1>
        </header>

        {/* Main Card */}
        <div className="card">

          {/* Upload Section com Drag & Drop */}
          <div className="upload-section">
            <div
              onDragOver={handleDragOver}
              onDragLeave={handleDragLeave}
              onDrop={handleDrop}
              className={`dropzone ${isDragging ? 'dragging' : ''}`}
            >
              <input
                type="file"
                accept=".xlsx, .xls, .csv"
                onChange={handleFileUpload}
                className="hidden"
                id="fileInput"
              />
              <label htmlFor="fileInput">
                <div className={`icon-circle ${isDragging ? 'dragging' : 'default'}`}>
                  <Upload className="icon-lg" />
                </div>
                <h3 className="drop-title">
                  {file ? file.name : (isDragging ? "Solte o arquivo aqui!" : "Clique ou arraste sua planilha aqui")}
                </h3>
                <p className="drop-subtitle">
                  Suporta .XLSX, .XLS e .CSV
                </p>
              </label>
            </div>

            {/* Actions */}
            {file && !processedData && (
              <div className="actions-container">
                <button
                  onClick={processFile}
                  disabled={isLoading || !isLibLoaded}
                  className="btn-primary"
                >
                  {isLoading ? (
                    <>
                      <Loader2 />
                      Processando...
                    </>
                  ) : (
                    <>
                      <RefreshCw />
                      Iniciar Processamento
                    </>
                  )}
                </button>
              </div>
            )}
          </div>

          {/* Error Message */}
          {error && (
            <div className="error-box">
              <AlertCircle className="error-icon" />
              <p>{error}</p>
            </div>
          )}

          {/* Results Section Simplificada */}
          {processedData && (
            <div className="results-section">
              <div className="results-content">
                <div className="success-icon-bg">
                  <CheckCircle className="icon-xl" />
                </div>
                <div>
                  <h3 className="results-title">Processamento Concluído!</h3>
                  <div className="results-text">
                  </div>
                </div>

                <button
                  onClick={downloadFile}
                  disabled={isDownloading}
                  className={`btn-download ${isDownloading ? 'loading' : 'success'}`}
                >
                  {isDownloading ? (
                    <>
                        <Loader2 className="w-6 h-6 spin" />
                        Gerando Arquivo...
                    </>
                  ) : (
                    <>
                        <Download className="w-6 h-6" />
                        Baixar Nova Planilha
                    </>
                  )}
                </button>
              </div>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}
