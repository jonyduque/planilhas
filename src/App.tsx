import React, { useState } from 'react';
import { Upload, FileSpreadsheet, Download, RefreshCw, AlertCircle, CheckCircle, Loader2 } from 'lucide-react';
import './App.css';
import { processRows } from './logic/processor';



import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';

export default function App() {
  const [file, setFile] = useState<File | null>(null);
  const [processedData, setProcessedData] = useState<any[][] | null>(null);
  const [headers, setHeaders] = useState<string[]>([]); // Armazenar cabeçalhos separadamente para o ExcelJS
  const [isLoading, setIsLoading] = useState<boolean>(false);
  const [isDownloading, setIsDownloading] = useState<boolean>(false);
  const [error, setError] = useState<string>("");
  const [isDragging, setIsDragging] = useState<boolean>(false);

  // Carregar as bibliotecas (agora via bundle, então efeito removido)


// Função helper para carregar script

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
    if (!file) return;
    setIsLoading(true);
    setError("");

    const reader = new FileReader();

    reader.onload = (e: ProgressEvent<FileReader>) => {
      try {
        const data = e.target?.result;
        if (!data) throw new Error("Erro ao ler dados do arquivo.");

        const workbook = XLSX.read(data, { type: 'binary' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        // Forçando o tipo any[] aqui pois sheet_to_json retorna array de objetos ou arrays dependendo da config
        const jsonData: any[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

        const result = processRows(jsonData);

        if (result.error) throw new Error(result.error);

        // Converte currentHeaders para string[] explicitamente para o estado
        setHeaders(result.headers);
        setProcessedData(result.data);

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
    if (!processedData) return;

    setIsDownloading(true);

    try {
      // Simular um pequeno delay para que o usuário veja o estado de carregamento
      await new Promise(resolve => setTimeout(resolve, 800));

      // Criar Workbook ExcelJS
      const workbook = new ExcelJS.Workbook();
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
      const feitoIndex = headers.findIndex(h => h.toString().toLowerCase() === "feito");

      // Ajustar larguras e alinhamentos
      worksheet.columns.forEach((col: any, index: number) => {
        // Largura padrão
        col.width = 20;

        // Configuração de alinhamento
        let horizontalAlign: 'left' | 'center' | 'right' = 'left';

        // Se for Localizadores, deixa mais largo
        if (index === locIndex) {
          col.width = 60;
          horizontalAlign = 'left';
        }
        // Se for Último Evento, formata como data curta
        else if (index === ultimoEventoIndex) {
          col.width = 18;
          col.numFmt = 'dd/mm/yyyy'; // Formato de data curta (pt-BR)
          horizontalAlign = 'center';
        }
        // Se for Feito, adiciona validação de dados
        else if (index === feitoIndex) {
          col.width = 15;
          horizontalAlign = 'center';
          col.dataValidation = {
            type: 'list',
            allowBlank: true,
            formulae: ['"Sim,Não"'] // Dropdown com opções
          };
        }

        col.alignment = {
          wrapText: true,
          vertical: 'middle',
          horizontal: horizontalAlign
        };
      });

      // Aplica Formatação Condicional
      // Se a coluna Feito existir
      if (feitoIndex !== -1) {
        // Obter a letra da coluna Feito (ExcelJS colunas são 1-based, mas index é 0-based)
        // Uma forma segura de pegar a letra é usar worksheet.getColumn(feitoIndex + 1).letter
        const feitoColLetter = worksheet.getColumn(feitoIndex + 1).letter;

        // Intervalo da tabela (começa na linha 2, até o final)
        // Podemos usar A2:Z<rowCount> (ajustando Z para ultima coluna)
        const lastColLetter = worksheet.getColumn(headers.length).letter;
        const totalRows = processedData.length + 1; // +1 do header

        worksheet.addConditionalFormatting({
          ref: `A2:${lastColLetter}${totalRows}`,
          rules: [
            {
              type: 'expression',
              priority: 1,
              formulae: [`=$${feitoColLetter}2="Sim"`], // $Col2 fixa a coluna, 2 é relativo a linha inicial
              style: {
                fill: {
                  type: 'pattern',
                  pattern: 'solid',
                  bgColor: { argb: 'FFC6EFCE' } // Green background
                },
                font: {
                  color: { argb: 'FF006100' } // Dark Green text
                },
                border: {
                  top: { style: 'thin', color: { argb: 'FF006100' } },
                  left: { style: 'thin', color: { argb: 'FF006100' } },
                  bottom: { style: 'thin', color: { argb: 'FF006100' } },
                  right: { style: 'thin', color: { argb: 'FF006100' } }
                }
              }
            }
          ]
        });
      }

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
                  disabled={isLoading}
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
