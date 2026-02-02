import { describe, it, expect } from 'vitest';
import { processRows } from './processor';

describe('processRows', () => {
  it('should process a valid file correctly', () => {
    // Row 0: Technical/ignored header
    // Row 1: Real headers
    // Row 2: Data
    const rawData = [
      ["Ignore", "Me"],
      ["Número Processo", "Other", "Localizadores", "Inclusão no Localizador", "Último Evento"],
      ["1234567890", "Data", "Loc A - TESTE (G)", "01/01/2023 10:00", "02/02/2023 11:00"]
    ];

    const result = processRows(rawData);

    expect(result.error).toBeUndefined();

    // Verify Column Order
    const procIndex = result.headers.indexOf("Número Processo");
    expect(procIndex).toBeGreaterThan(-1);
    expect(result.headers[procIndex + 1]).toBe("Dígito");
    expect(result.headers[procIndex + 2]).toBe("Feito");

    const locIndexHeader = result.headers.indexOf("Localizadores");
    expect(locIndexHeader).toBeGreaterThan(-1);
    expect(result.headers[locIndexHeader + 1]).toBe("Localizadores do Gabinete");

    const processedRow = result.data[0];

    // Check Dígito (index 6 of "1234567890" is "7")
    const digitoIndex = result.headers.indexOf("Dígito");
    expect(processedRow[digitoIndex]).toBe("7");

    // Check Feito
    const feitoIndex = result.headers.indexOf("Feito");
    expect(processedRow[feitoIndex]).toBe("FALSO");

    // Check Localizadores logic (G count)
    const gCountIndex = result.headers.indexOf("Localizadores do Gabinete");
    expect(processedRow[gCountIndex]).toBe(1);

    // Check Date parsing
    const inclusaoIndex = result.headers.indexOf("Inclusão no Localizador");
    expect(processedRow[inclusaoIndex]).toBeInstanceOf(Date);
  });

  it('should handle short process numbers for Dígito', () => {
    const rawData = [
      ["Ignored"],
      ["Número Processo"],
      ["123456"] // Length 6, no 7th char
    ];

    const result = processRows(rawData);
    const digitoIndex = result.headers.indexOf("Dígito");
    expect(result.data[0][digitoIndex]).toBe("");
  });

  it('should format Localizadores text with line breaks for Uppercase/Symbols', () => {
    // "Loc A - TESTE" -> "Loc A\nTESTE" (TESTE is all upper)
    // "Loc B - Normal" -> "Loc B - Normal" (Normal is Mixed)
    const rawData = [
      ["Ignored"],
      ["Localizadores"],
      ["Loc A - TESTE"],
      ["Loc B - Normal"]
    ];

    const result = processRows(rawData);
    // col index 0 in processedRow (since we only have Localizadores from original map)
    const locIndex = 0;

    // processedRow 0
    expect(result.data[0][locIndex]).toBe("Loc A\nTESTE");

    // processedRow 1
    expect(result.data[1][locIndex]).toBe("Loc B - Normal");
  });

  it('should return error for empty file', () => {
    const result = processRows([]);
    expect(result.error).toBeDefined();
  });
});
