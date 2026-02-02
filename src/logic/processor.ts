export interface ProcessResult {
  headers: string[];
  data: any[][];
  error?: string;
}

interface ColMap {
  processo: number;
  localizadores: number;
  inclusao: number;
  ultimoEvento: number;
}

// Função auxiliar para decodificar HTML Entities
const decodeHTMLEntities = (text: any): string => {
  if (!text || typeof text !== 'string') return String(text || "");
  // Em ambiente Node/Testes, DOMParser não existe nativamente sem jsdom.
  // Criaremos uma implementação simples ou usaremos uma lib se necessário.
  // Mas como estamos com jsdom configurado no Vitest, deve funcionar.
  // Fallback seguro caso DOMParser não esteja disponível:
  if (typeof DOMParser === 'undefined') {
    return text.replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>');
  }

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

    // Validação básica
    if (!day || !month || !year) return dateStr;

    let hours = 0, minutes = 0, seconds = 0;
    if (timePart) {
      [hours, minutes, seconds] = timePart.split(':').map(Number);
    }

    return new Date(year, month - 1, day, hours, minutes, seconds || 0);
  } catch (e) {
    return dateStr;
  }
};

export const processRows = (rawData: any[][]): ProcessResult => {
  try {
    if (rawData.length < 2) return { headers: [], data: [], error: "Arquivo vazio ou sem linhas suficientes." };

    // 1. Excluir a primeira linha (cabeçalho técnico/falso)
    // Assumindo que o primeiro array é cabeçalho técnico e o segundo é o real, conforme lógica original
    // Na lógica original:
    // const jsonData: any[][] = ...
    // const cleanData = jsonData.slice(1);
    // const currentHeaders = cleanData[0];
    // const rows = cleanData.slice(1);

    // Então rawData[0] é ignorado. rawData[1] são os headers originais.

    const cleanData = rawData.slice(1);
    if (cleanData.length === 0) return { headers: [], data: [], error: "Não há dados após remover primeira linha." };

    const currentHeaders: any[] = cleanData[0];
    const rows = cleanData.slice(1);

    const colMap: ColMap = {
      processo: currentHeaders.findIndex((h: any) => h && h.toString().toLowerCase().includes("número processo")),
      localizadores: currentHeaders.findIndex((h: any) => h && h.toString().toLowerCase().includes("localizadores")),
      inclusao: currentHeaders.findIndex((h: any) => h && h.toString().toLowerCase().includes("inclusão no localizador")),
      ultimoEvento: currentHeaders.findIndex((h: any) => h && h.toString().toLowerCase().includes("último evento")),
    };

    // Nota: A lógica original apenas avisa, não retorna erro, mas define headers
    const newHeaders = [...currentHeaders, "Localizadores do Gabinete", "Dígito", "Feito"].map(String);

    const processedRows = rows.map((row: any[]) => {
      // Garante que row é um array
      if (!Array.isArray(row)) return row;

      // Preenche colunas vazias até o header original
      while (row.length < currentHeaders.length) row.push("");

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

        locText = locText.replace(/\(Principal\)/gi, '').trim();

        // Regex para quebra de linha
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

      // Extrair Dígito
      let digito = "";
      if (colMap.processo !== -1 && row[colMap.processo]) {
        const procStr = String(row[colMap.processo]).trim();
        if (procStr.length >= 7) {
          digito = procStr.charAt(6);
        }
      }

      // Adicionar novas colunas
      row.push(gCount);
      row.push(digito);
      row.push("FALSO");

      return row;
    });

    return {
      headers: newHeaders,
      data: processedRows
    };

  } catch (err: any) {
    return { headers: [], data: [], error: err.message || String(err) };
  }
};
