import * as XLSX from 'xlsx';
import { ProcessedData, ProductInfo } from '../types';

/**
 * Processa o arquivo Excel para extrair produtos com variação de custo > 25%.
 * Baseado rigorosamente no algoritmo fornecido na especificação do usuário.
 */
export function processExcelData(arrayBuffer: ArrayBuffer, variationThreshold: number = 25): ProcessedData {
  const data = new Uint8Array(arrayBuffer);
  const workbook = XLSX.read(data, { type: 'array' });
  
  // Processar a primeira aba da planilha
  const firstSheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[firstSheetName];
  if (!worksheet) {
    return {};
  }
  
  // Converter a planilha para um array de arrays (linhas e colunas brutas)
  const rows = XLSX.utils.sheet_to_json<any[]>(worksheet, { header: 1 });
  
  const results: ProcessedData = {};
  
  let currentLoja: string | null = null;
  let currentFornecedorContext = "N/A";
  let currentNotaContext = "N/A";
  
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    // O índice da coluna % de variação é 29, portanto a linha deve ter pelo menos 30 elementos
    if (!row || row.length < 30) {
      continue;
    }
    
    const col0Val = row[0];
    const col0Str = col0Val !== undefined && col0Val !== null ? String(col0Val).trim() : "";
    
    // 1. Identificar Troca de Loja
    // Formato esperado: "1-SAO GERALDO", "2-CONSELHEIRO", etc.
    if (col0Str.includes("-")) {
      const parts = col0Str.split("-");
      if (parts.length > 1 && /^\d+$/.test(parts[0].trim())) {
        const code = parts[0].trim();
        const restName = parts.slice(1).join("-").trim();
        currentLoja = `${code} - ${restName}`;
        continue;
      }
    }
    
    // 2. Identificar Contexto de Fornecedor
    // Coluna 8 é o fornecedor (0-indexed)
    const potentialFornVal = row[8];
    const potentialFornStr = potentialFornVal !== undefined && potentialFornVal !== null ? String(potentialFornVal).trim() : "";
    
    if (
      (col0Val === undefined || col0Val === null || String(col0Val).trim() === "") &&
      potentialFornStr !== "" &&
      !potentialFornStr.toLowerCase().includes("total") &&
      potentialFornStr !== "*Fornecedor" &&
      potentialFornStr.toLowerCase() !== "nan"
    ) {
      currentFornecedorContext = potentialFornStr;
    }
    
    // 3. Identificar Contexto de Nota
    // Coluna 17 é o número da nota (0-indexed)
    const potentialNotaVal = row[17];
    if (potentialNotaVal !== undefined && potentialNotaVal !== null) {
      const potentialNotaStr = String(potentialNotaVal).trim();
      const cleanNota = potentialNotaStr.replace(".0", "");
      if (/^\d+$/.test(cleanNota) && cleanNota !== "") {
        currentNotaContext = String(parseInt(cleanNota, 10));
      }
    }
    
    // 4. Processar Linha de Produto
    // Se col0Str for apenas dígitos, trata-se de um código de produto válido
    if (/^\d+$/.test(col0Str)) {
      try {
        const codigo = parseInt(col0Str, 10);
        
        // Coluna 2 é o nome do produto (0-indexed)
        const nomeVal = row[2];
        const nome = nomeVal !== undefined && nomeVal !== null ? String(nomeVal).trim() : "Produto Sem Nome";
        
        // Determinar Fornecedor específico da linha ou usar o do contexto
        let fornecedor = currentFornecedorContext;
        if (row[8] !== undefined && row[8] !== null) {
          const fornRowStr = String(row[8]).trim();
          if (
            fornRowStr !== "" && 
            fornRowStr.toLowerCase() !== "nan" && 
            !fornRowStr.toLowerCase().includes("deposito") && 
            !fornRowStr.toLowerCase().includes("fornecedor")
          ) {
            fornecedor = fornRowStr;
          }
        }
        
        // Determinar Nota específica da linha ou usar a do contexto
        let nota = currentNotaContext;
        if (row[17] !== undefined && row[17] !== null) {
          const notaRowStr = String(row[17]).trim().replace(".0", "");
          if (/^\d+$/.test(notaRowStr) && notaRowStr !== "") {
            nota = String(parseInt(notaRowStr, 10));
          }
        }
        
        // Extrair custos e variação
        // Coluna 24 é custo anterior, 26 é custo atual, 29 é variação (0-indexed)
        const custoAnterVal = row[24];
        const custoAtualVal = row[26];
        const variacaoVal = row[29];
        
        if (variacaoVal !== undefined && variacaoVal !== null) {
          const varFloat = parseFloat(String(variacaoVal));
          if (!isNaN(varFloat)) {
            // Critério de variação: maior que o limite definido para mais ou para menos
            if (Math.abs(varFloat) > variationThreshold) {
              const custoAnter = custoAnterVal !== undefined && custoAnterVal !== null && !isNaN(Number(custoAnterVal)) 
                ? parseFloat(String(custoAnterVal)) 
                : 0.0;
              const custoAtual = custoAtualVal !== undefined && custoAtualVal !== null && !isNaN(Number(custoAtualVal)) 
                ? parseFloat(String(custoAtualVal)) 
                : 0.0;
              
              const prodInfo: ProductInfo = {
                codigo,
                nome,
                custo_anter: custoAnter,
                custo_atual: custoAtual,
                variacao: varFloat
              };
              
              if (currentLoja && nota && fornecedor) {
                if (!results[currentLoja]) {
                  results[currentLoja] = {};
                }
                if (!results[currentLoja][nota]) {
                  results[currentLoja][nota] = {};
                }
                if (!results[currentLoja][nota][fornecedor]) {
                  results[currentLoja][nota][fornecedor] = [];
                }
                
                // Evitar duplicidade exata
                const alreadyExists = results[currentLoja][nota][fornecedor].some(
                  p => p.codigo === prodInfo.codigo && 
                       p.custo_anter === prodInfo.custo_anter && 
                       p.custo_atual === prodInfo.custo_atual &&
                       p.variacao === prodInfo.variacao
                );
                
                if (!alreadyExists) {
                  results[currentLoja][nota][fornecedor].push(prodInfo);
                }
              }
            }
          }
        }
      } catch (err) {
        // Ignora erros individuais de linha e prossegue
      }
    }
  }
  
  return results;
}

/**
 * Formata os dados processados em uma string Markdown correspondente à do Streamlit.
 */
export function formatProcessedOutput(processedData: ProcessedData, variationThreshold: number = 25): string {
  const outputLines: string[] = [];
  
  const lojas = Object.keys(processedData);
  if (lojas.length === 0) {
    return `Nenhum produto com variação superior a ${variationThreshold}% foi encontrado.`;
  }
  
  // Ordenar lojas pelo número inicial para uma apresentação consistente
  const sortedLojas = [...lojas].sort((a, b) => {
    const numA = a.includes(" - ") ? parseInt(a.split(" - ")[0], 10) : 999;
    const numB = b.includes(" - ") ? parseInt(b.split(" - ")[0], 10) : 999;
    const cleanNumA = isNaN(numA) ? 999 : numA;
    const cleanNumB = isNaN(numB) ? 999 : numB;
    return cleanNumA - cleanNumB;
  });
  
  for (const loja of sortedLojas) {
    outputLines.push(`# ${loja}`);
    outputLines.push(""); // Linha em branco para espaçamento
    
    const notasDaLoja = processedData[loja];
    const sortedNotas = Object.keys(notasDaLoja).sort((a, b) => parseInt(a, 10) - parseInt(b, 10));
    
    for (const nota of sortedNotas) {
      const fornecedoresDaNota = notasDaLoja[nota];
      const sortedFornecedores = Object.keys(fornecedoresDaNota).sort();
      
      for (const fornecedor of sortedFornecedores) {
        outputLines.push(`NOTA: ${nota} FORNECEDOR: ${fornecedor}`);
        for (const prod of fornecedoresDaNota[fornecedor]) {
          const line = `${prod.codigo} - ${prod.nome} - CUSTO ANTERIOR R$${prod.custo_anter.toFixed(2)} - CUSTO ATUAL R$${prod.custo_atual.toFixed(2)} ( ${prod.variacao.toFixed(2)}% )`;
          outputLines.push(line);
        }
        outputLines.push(""); // Espaço após bloco de produtos da mesma nota/fornecedor
      }
    }
    
    outputLines.push("------------------------------------------------------------------------------------------------------------");
    outputLines.push(""); // Espaço após o separador de loja
  }
  
  return outputLines.join("\n");
}
