import { Document, Packer, Paragraph, TextRun, AlignmentType } from "docx";
import { jsPDF } from "jspdf";
import { ProcessedData } from "../types";

/**
 * Triggers a file download in the browser.
 */
function downloadBlob(blob: Blob, filename: string) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

/**
 * Exporta os dados processados para um arquivo Word (.docx) formatado e estilizado.
 */
export function exportToWord(processedData: ProcessedData, filename = "relatorio_variacao_custo.docx") {
  const sectionsChildren: Paragraph[] = [];
  
  // Título Principal do Relatório
  sectionsChildren.push(
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 200, after: 150 },
      children: [
        new TextRun({
          text: "RELATÓRIO DE VARIAÇÃO DE CUSTOS",
          bold: true,
          size: 32, // 16pt
          color: "1F2937", // Slate 800
        }),
      ],
    })
  );
  
  // Subtítulo com Metadata
  sectionsChildren.push(
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 600 },
      children: [
        new TextRun({
          text: `Gerado em: ${new Date().toLocaleDateString("pt-BR")} às ${new Date().toLocaleTimeString("pt-BR")}`,
          italics: true,
          size: 20, // 10pt
          color: "4B5563", // Gray 600
        }),
      ],
    })
  );

  const sortedLojas = Object.keys(processedData).sort((a, b) => {
    const numA = a.includes(" - ") ? parseInt(a.split(" - ")[0], 10) : 999;
    const numB = b.includes(" - ") ? parseInt(b.split(" - ")[0], 10) : 999;
    const cleanNumA = isNaN(numA) ? 999 : numA;
    const cleanNumB = isNaN(numB) ? 999 : numB;
    return cleanNumA - cleanNumB;
  });

  if (sortedLojas.length === 0) {
    sectionsChildren.push(
      new Paragraph({
        spacing: { before: 200 },
        children: [
          new TextRun({
            text: "Nenhum produto com variação superior a 25% foi encontrado.",
            size: 24,
            color: "6B7280",
          }),
        ],
      })
    );
  }

  for (const loja of sortedLojas) {
    // Título da Loja (Heading 1)
    sectionsChildren.push(
      new Paragraph({
        spacing: { before: 400, after: 200 },
        children: [
          new TextRun({
            text: loja,
            bold: true,
            size: 26, // 13pt
            color: "1D4ED8", // Navy Blue
          }),
        ],
      })
    );

    const notasDaLoja = processedData[loja];
    const sortedNotas = Object.keys(notasDaLoja).sort((a, b) => parseInt(a, 10) - parseInt(b, 10));

    for (const nota of sortedNotas) {
      const fornecedoresDaNota = notasDaLoja[nota];
      const sortedFornecedores = Object.keys(fornecedoresDaNota).sort();

      for (const fornecedor of sortedFornecedores) {
        // Cabeçalho da Nota e Fornecedor
        sectionsChildren.push(
          new Paragraph({
            spacing: { before: 200, after: 100 },
            children: [
              new TextRun({
                text: `NOTA: ${nota}    FORNECEDOR: ${fornecedor}`,
                bold: true,
                size: 20, // 10pt
                color: "374151", // Gray 700
              }),
            ],
          })
        );

        // Listar os produtos
        for (const prod of fornecedoresDaNota[fornecedor]) {
          const isIncrease = prod.variacao > 0;
          const varSignal = isIncrease ? "+" : "";
          const colorHex = isIncrease ? "DC2626" : "16A34A"; // Vermelho para alta, verde para baixa

          sectionsChildren.push(
            new Paragraph({
              spacing: { before: 60, after: 60 },
              indent: { left: 360 }, // Recuo de lista
              children: [
                new TextRun({
                  text: `${prod.codigo} - ${prod.nome} - `,
                  size: 19,
                  color: "1F2937",
                }),
                new TextRun({
                  text: `CUSTO ANTERIOR R$${prod.custo_anter.toFixed(2)} - CUSTO ATUAL R$${prod.custo_atual.toFixed(2)} `,
                  size: 19,
                  color: "4B5563",
                }),
                new TextRun({
                  text: `( ${varSignal}${prod.variacao.toFixed(2)}% )`,
                  bold: true,
                  size: 19,
                  color: colorHex,
                }),
              ],
            })
          );
        }
        
        // Espaço sutil após bloco de produtos
        sectionsChildren.push(
          new Paragraph({
            spacing: { after: 100 },
            children: [],
          })
        );
      }
    }
    
    // Divisor de Loja
    sectionsChildren.push(
      new Paragraph({
        spacing: { before: 200, after: 300 },
        children: [
          new TextRun({
            text: "------------------------------------------------------------------------------------------------------------------------",
            color: "E5E7EB",
            size: 14,
          }),
        ],
      })
    );
  }

  const doc = new Document({
    sections: [{
      properties: {},
      children: sectionsChildren,
    }],
  });

  Packer.toBlob(doc).then((blob) => {
    downloadBlob(blob, filename);
  });
}

/**
 * Exporta os dados processados para um arquivo PDF (.pdf) formatado profissionalmente.
 */
export function exportToPDF(processedData: ProcessedData, filename = "relatorio_variacao_custo.pdf") {
  const doc = new jsPDF({
    orientation: "portrait",
    unit: "mm",
    format: "a4",
  });
  
  let y = 20;
  const pageHeight = doc.internal.pageSize.height;
  const margin = 20;

  // Função interna para paginação segura
  const checkPageBreak = (neededSpace: number) => {
    if (y + neededSpace > pageHeight - margin) {
      doc.addPage();
      y = 20;
      return true;
    }
    return false;
  };

  // Banner Azul-Slate no Topo da Primeira Página
  doc.setFillColor(31, 41, 55); // Slate 800
  doc.rect(0, 0, 210, 38, "F");
  
  doc.setTextColor(255, 255, 255);
  doc.setFont("Helvetica", "bold");
  doc.setFontSize(16);
  doc.text("RELATÓRIO DE VARIAÇÃO DE CUSTOS EXCEL", 105, 18, { align: "center" });
  
  doc.setFont("Helvetica", "normal");
  doc.setFontSize(10);
  doc.text(`Gerado em: ${new Date().toLocaleDateString("pt-BR")} às ${new Date().toLocaleTimeString("pt-BR")}`, 105, 28, { align: "center" });

  y = 50;

  const sortedLojas = Object.keys(processedData).sort((a, b) => {
    const numA = a.includes(" - ") ? parseInt(a.split(" - ")[0], 10) : 999;
    const numB = b.includes(" - ") ? parseInt(b.split(" - ")[0], 10) : 999;
    const cleanNumA = isNaN(numA) ? 999 : numA;
    const cleanNumB = isNaN(numB) ? 999 : numB;
    return cleanNumA - cleanNumB;
  });

  if (sortedLojas.length === 0) {
    doc.setFont("Helvetica", "normal");
    doc.setFontSize(11);
    doc.setTextColor(107, 114, 128);
    doc.text("Nenhum produto com variação superior a 25% foi encontrado.", 20, y);
    doc.save(filename);
    return;
  }

  for (const loja of sortedLojas) {
    checkPageBreak(25);
    
    // Cabeçalho da Loja com fundo cinza claro
    doc.setFillColor(243, 244, 246); // Gray 100
    doc.rect(15, y - 5, 180, 8, "F");
    
    doc.setFont("Helvetica", "bold");
    doc.setFontSize(11);
    doc.setTextColor(29, 78, 216); // Royal Blue 700
    doc.text(loja, 20, y);
    y += 10;

    const notasDaLoja = processedData[loja];
    const sortedNotas = Object.keys(notasDaLoja).sort((a, b) => parseInt(a, 10) - parseInt(b, 10));

    for (const nota of sortedNotas) {
      const fornecedoresDaNota = notasDaLoja[nota];
      const sortedFornecedores = Object.keys(fornecedoresDaNota).sort();

      for (const fornecedor of sortedFornecedores) {
        checkPageBreak(18);
        
        doc.setFont("Helvetica", "bold");
        doc.setFontSize(9);
        doc.setTextColor(55, 65, 81); // Slate 700
        doc.text(`NOTA: ${nota}    FORNECEDOR: ${fornecedor}`, 20, y);
        y += 6;

        for (const prod of fornecedoresDaNota[fornecedor]) {
          checkPageBreak(10);
          
          doc.setFont("Helvetica", "normal");
          doc.setFontSize(8.5);
          doc.setTextColor(17, 24, 39); // Gray 900
          
          // Formatando código e nome com limite de caracteres
          const codAndName = `${prod.codigo} - ${prod.nome}`;
          let truncatedName = codAndName;
          if (truncatedName.length > 52) {
            truncatedName = truncatedName.substring(0, 49) + "...";
          }
          doc.text(`• ${truncatedName}`, 22, y);

          // Custos formatados
          const costText = `Ant: R$ ${prod.custo_anter.toFixed(2)}  →  Atu: R$ ${prod.custo_atual.toFixed(2)}`;
          doc.setTextColor(75, 85, 99); // Gray 600
          doc.text(costText, 118, y);

          // Percentual de Variação com cor condicional
          const varText = `${prod.variacao > 0 ? "+" : ""}${prod.variacao.toFixed(2)}%`;
          if (prod.variacao > 0) {
            doc.setTextColor(220, 38, 38); // Red 600
          } else {
            doc.setTextColor(22, 163, 74); // Green 600
          }
          doc.setFont("Helvetica", "bold");
          doc.text(varText, 185, y, { align: "right" });
          
          // Reset para texto normal
          doc.setFont("Helvetica", "normal");
          y += 5.5;
        }
        
        y += 3; // Pequeno respiro após o bloco
      }
    }
    
    checkPageBreak(12);
    // Linha divisória fina cinza entre lojas
    doc.setDrawColor(229, 231, 235); // Gray 200
    doc.line(15, y, 195, y);
    y += 8;
  }

  // Rodapé dinâmico em todas as páginas
  const pageCount = doc.internal.pages.length;
  for (let i = 1; i <= pageCount; i++) {
    doc.setPage(i);
    doc.setFont("Helvetica", "normal");
    doc.setFontSize(8);
    doc.setTextColor(156, 163, 175); // Gray 400
    doc.text(`Página ${i} de ${pageCount}`, 105, 287, { align: "center" });
    doc.text("Análise de Variação de Custos", 20, 287);
    doc.text("Desenvolvido por Manus AI", 190, 287, { align: "right" });
  }

  doc.save(filename);
}
