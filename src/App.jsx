import React, { useState, useEffect } from 'react';
import { systemPrompt } from './config/prompts.jsx';
import * as XLSX from "xlsx";

//docx js
import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  HeadingLevel,
  Table,
  TableRow,
  TableCell,
  WidthType
} from "docx";

import { saveAs } from "file-saver";


// Definição das 10 novas competências avaliadas
const COMPETENCIES = [
  { id: 'research', label: 'Pesquisa com Usuário', desc: 'Condução de pesquisas, entrevistas, testes de usabilidade e síntese de dados.' },
  { id: 'ui', label: 'Design de Interfaces', desc: 'Domínio visual, tipografia, cores, grids e consistência estética.' },
  { id: 'tools', label: 'Ferramentas de Design', desc: 'Domínio de softwares como Figma, Sketch, Adobe XD e prototipação.' },
  { id: 'metrics', label: 'Análise, Métricas e Dados de UX', desc: 'Capacidade de coletar, analisar e interpretar dados quantitativos e qualitativos para decisões de design.' },
  { id: 'architecture', label: 'Arquitetura de Informação', desc: 'Mapeamento de fluxos de usuário, sitemaps, taxonomia e navegação.' },
  { id: 'documentation', label: 'Documentação & Entregas', desc: 'Especificação técnica, organização de handoff para desenvolvimento e UI kits.' },
  { id: 'feedback_iter', label: 'Feedback & Iteração', desc: 'Coleta de métricas e melhoria contínua a partir de dados reais do usuário.' },
  { id: 'communication', label: 'Comunicação & Colaboração', desc: 'Alinhamento com stakeholders, escuta ativa e trabalho em equipe.' },
  { id: 'receive_feedback', label: 'Receber Feedback', desc: 'Maturidade profissional para escutar e implementar críticas construtivas.' },
  { id: 'problem_solving', label: 'Resolução de Problemas', desc: 'Pensamento crítico e análise estratégica de gargalos e restrições de projetos.' },
  { id: 'autonomy', label: 'Autonomia', desc: 'Proatividade, gerenciamento do próprio tempo e capacidade de entrega ponta a ponta.' },
  { id: 'presentation', label: 'Apresentar', desc: 'Defesa fundamentada de soluções, storytelling e facilitação de dinâmicas.' }
];

export default function App() {
  // Estado padrão limpo para o formulário
  const [profile, setProfile] = useState({
    name: '',
    role: '',
    experience: '',
    date: new Date().toISOString().split('T')[0]
  });

  // Inicialização neutra com nota 3 (Intermediário) para todas as 12 competências
  const [scores, setScores] = useState({
    research: 3,
    ui: 3,
    tools: 3,
    metrics: 3,
    architecture: 3,
    documentation: 3,
    feedback_iter: 3,
    communication: 3,
    receive_feedback: 3,
    problem_solving: 3,
    autonomy: 3,
    presentation: 3
  });

  const [discursive, setDiscursive] = useState({
    systems: '',
    challenge: '',
    success: '',
    gaps: ''
  });

  // Estado da geração do relatório
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);
  const [report, setReport] = useState("");
  const [activeTab, setActiveTab] = useState('input'); // 'input' ou 'report'

  // Estados específicos para controle do Upload de Planilhas
  const [uploadedFile, setUploadedFile] = useState(null);
  const [multipleProfessionals, setMultipleProfessionals] = useState([]);
  const [selectedMappingIndex, setSelectedMappingIndex] = useState(null);
  const [customToast, setCustomToast] = useState(null);

  /* Carregar dinamicamente a biblioteca XLSX (SheetJS) por CDN para leitura de planilhas
  useEffect(() => {
    const script = document.createElement('script');
    script.src = "https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js";
    script.async = true;
    document.head.appendChild(script);
    return () => {
      const existing = document.querySelector(`script[src="${script.src}"]`);
      if (existing) {
        existing.remove();
      }
    };
  }, []);*/

  // Mostrar mensagens customizadas elegantes na UI (evitando alert)
  const showToast = (title, message, type = "success") => {
    setCustomToast({ title, message, type });
    setTimeout(() => {
      setCustomToast(null);
    }, 5000);
  };

  // Alterar notas individuais
  const handleScoreChange = (id, val) => {
    setScores(prev => ({ ...prev, [id]: parseInt(val) }));
  };

  // Alterar textos individuais
  const handleDiscursiveChange = (id, val) => {
    setDiscursive(prev => ({ ...prev, [id]: val }));
  };

  // Processar leitura de arquivo Excel / CSV para carregar dados
  const handleFileChange = (e) => {
    const file = e.target.files[0];
    if (!file) return;

    if (!XLSX) {
      setError("A biblioteca de processamento de planilhas ainda está carregando no navegador. Aguarde 3 segundos e tente novamente.");
      return;
    }

    setUploadedFile(file);
    const reader = new FileReader();
    
    reader.onload = (evt) => {
      try {
        const dataBinary = evt.target.result;
        const workbook = XLSX.read(dataBinary, { type: 'binary' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        
        // Converter em matriz de dados por posição de coluna.
        // Importante: as competências quantitativas usam colunas fixas da planilha
        // para evitar erro de leitura por cabeçalhos parecidos.
        const rowsArray = XLSX.utils.sheet_to_json(worksheet, {
          header: 1,
          defval: ""
        });

        if (rowsArray.length <= 1) {
          setError("A planilha importada parece estar vazia.");
          return;
        }

        const headers = rowsArray[0];
        const dataRows = rowsArray.slice(1).filter((rowArray) =>
          rowArray.some((cell) => String(cell).trim() !== "")
        );

        if (dataRows.length === 0) {
          setError("A planilha importada não possui respostas preenchidas.");
          return;
        }

        const columnToIndex = (col) => {
          let index = 0;
          const normalizedCol = col.toUpperCase().trim();

          for (let i = 0; i < normalizedCol.length; i++) {
            index = index * 26 + normalizedCol.charCodeAt(i) - 64;
          }

          return index - 1;
        };

        const getCellByColumn = (rowArray, col) => {
          return rowArray[columnToIndex(col)];
        };

        const toScore = (value) => {
          const parsed = parseInt(value, 10);
          return Math.min(5, Math.max(1, Number.isNaN(parsed) ? 3 : parsed));
        };

        const normalizeText = (text) =>
          String(text || "")
            .toLowerCase()
            .normalize("NFD")
            .replace(/[\u0300-\u036f]/g, "")
            .trim();

        const parsedList = dataRows.map((rowArray) => {
          const findByHeader = (candidates) => {
            const foundIndex = headers.findIndex((header) => {
              const normalizedHeader = normalizeText(header);
              return candidates.some((candidate) => normalizedHeader.includes(normalizeText(candidate)));
            });

            return foundIndex >= 0 ? rowArray[foundIndex] : "";
          };

          const name =
            findByHeader(["nome", "profissional", "participante", "user", "name"]) ||
            "Profissional Sem Nome";

          const role =
            findByHeader(["cargo", "funcao", "função", "titulo", "título", "role", "position"]) ||
            "Designer de Produto";

          const experience =
            findByHeader(["experiencia", "experiência", "tempo", "anos", "senioridade"]) ||
            "Experiência não informada";

          // Mapeamento oficial por coluna da planilha:
          const research = toScore(getCellByColumn(rowArray, "M"));
          const ui = toScore(getCellByColumn(rowArray, "O"));
          const architecture = toScore(getCellByColumn(rowArray, "Q"));
          const tools = toScore(getCellByColumn(rowArray, "S"));
          const metrics = toScore(getCellByColumn(rowArray, "U"));
          const documentation = toScore(getCellByColumn(rowArray, "W"));
          const feedback_iter = toScore(getCellByColumn(rowArray, "Y"));
          const communication = toScore(getCellByColumn(rowArray, "AC"));
          const receive_feedback = toScore(getCellByColumn(rowArray, "AE"));
          const problem_solving = toScore(getCellByColumn(rowArray, "AG"));
          const autonomy = toScore(getCellByColumn(rowArray, "AI"));
          const presentation = toScore(getCellByColumn(rowArray, "AK"));

          // Respostas discursivas por coluna fixa:
          // AL = Expectativa
          // AM = Habilidades a desenvolver
          // AN = Desafios atuais
          // AO = Gaps percebidos
          const expectationText =
            getCellByColumn(rowArray, "AM") ||
            findByHeader(["principal expectativa", "expectativa ao participar", "programa de mentoria"]) ||
            "Expectativa não informada.";

          const skillsText =
            getCellByColumn(rowArray, "AN") ||
            findByHeader(["habilidades", "competencias especificas", "competências específicas", "desenvolver ou aprimorar"]) ||
            "Habilidades a desenvolver não informadas.";

          const challengesText =
            getCellByColumn(rowArray, "AO") ||
            findByHeader(["desafios", "projetos de ux", "caso ja atue", "caso já atue"]) ||
            "Desafios atuais não informados.";

          const gapsText =
            getCellByColumn(rowArray, "AP") ||
            findByHeader(["gaps", "conhecimento", "praticas", "práticas", "trabalho de ux"]) ||
            "Gaps percebidos não informados.";

          return {
            name,
            role,
            experience,
            date: new Date().toISOString().split('T')[0],
            scores: {
              research,
              ui,
              architecture,
              tools,
              metrics,
              documentation,
              feedback_iter,
              communication,
              receive_feedback,
              problem_solving,
              autonomy,
              presentation
            },
            discursive: {
              expectation: expectationText,
              skills: skillsText,
              challenges: challengesText,
              gaps: gapsText
          }
          };
        });
        console.log("PARSED LIST", parsedList);  
        setMultipleProfessionals(parsedList);
        setSelectedMappingIndex(0); // foca no primeiro por padrão
        
        applySelectedProfessional(parsedList[0]);
        showToast("Planilha Importada", `Sucesso ao ler o documento. ${parsedList.length} profissional(is) mapeado(s).`, "success");

      } catch (err) {
        console.error(err);
        setError("Ocorreu um erro ao decodificar a planilha. Verifique se o formato está correto (.xlsx ou .csv).");
      }
    };

    reader.readAsBinaryString(file);
  };

  // Aplica o profissional selecionado da planilha aos inputs ativos do app
  const applySelectedProfessional = (prof) => {
    setProfile({
      name: prof.name,
      role: prof.role,
      experience: prof.experience,
      date: prof.date
    });
    setScores({ ...prof.scores });
    setDiscursive({ ...prof.discursive });
  };

  // Manipular troca manual de profissional na lista da planilha importada
  const handleSelectProfessionalFromList = (idx) => {
    setSelectedMappingIndex(idx);
    applySelectedProfessional(multipleProfessionals[idx]);
    showToast("Profissional Alterado", `Carregados os dados de ${multipleProfessionals[idx].name}`, "info");
  };

  // Lógica de cálculo de coordenadas para o Gráfico Radar SVG (Decágono de 10 competências)
  const size = 320;
  const center = size / 2;
  const radius = 95;
  const levels = [1, 2, 3, 4, 5];

  // Ângulo baseado no comprimento dinâmico das competências (10 posições)
  const getCoordinates = (index, value) => {
    const angle = (index * 2 * Math.PI) / COMPETENCIES.length - Math.PI / 2;
    const r = (value / 5) * radius;
    const x = center + r * Math.cos(angle);
    const y = center + r * Math.sin(angle);
    return { x, y };
  };

  // Gerar o polígono principal para a autoavaliação do designer
  const points = COMPETENCIES.map((comp, i) => {
    const score = scores[comp.id] || 1;
    const { x, y } = getCoordinates(i, score);
    return `${x},${y}`;
  }).join(' ');

 // Requisição para a API interna da Vercel
const fetchGeminiReport = async (payload) => {
  let retries = 1;
  let delay = 1000;

  for (let i = 0; i < retries; i++) {
    try {
      const response = await fetch('/api/gemini', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({ payload })
      });

      const data = await response.json();

      if (response.ok) {
        return data;
      }

      if (response.status === 429) {
        throw new Error("Limite temporário da API atingido. Aguarde 1 minuto e tente novamente.");

      }

      throw new Error(
        data?.error?.message ||
        `Erro na API interna (Código: ${response.status})`
      );

    } catch (err) {
      if (i === retries - 1) throw err;

      await new Promise(res => setTimeout(res, delay));
      delay *= 2;
    }
  }
};

  // Disparar geração do relatório executivo
  const generateReport = async () => {
    if (!profile.name || !profile.role) {
      setError("Por favor, preencha o Nome e o Cargo do avaliado antes de gerar o relatório executivo.");
      return;
    }

    setLoading(true);
    setError(null);
    setReport("");
    setActiveTab('report');

        const userPrompt = `DADOS DO PROFISSIONAL AVALIADO:
- Nome: ${profile.name}
- Cargo Atual: ${profile.role}
- Experiência: ${profile.experience || "Não informada"}
- Data da Avaliação: ${profile.date}

NOTAS AUTOAVALIADAS DE 1 A 5 (Mapeando o Radar de 12 Pontos):
1. Pesquisa com Usuário: ${scores.research} / 5
2. Design de Interfaces: ${scores.ui} / 5
3. Ferramentas de Design: ${scores.tools} / 5
4. Análise, Métricas e Dados de UX: ${scores.metrics} / 5
5. Arquitetura de Informação: ${scores.architecture} / 5
6. Documentação & Entregas: ${scores.documentation} / 5
7. Feedback & Iteração: ${scores.feedback_iter} / 5
8. Comunicação & Colaboração em Equipe: ${scores.communication} / 5
9. Receber Feedback: ${scores.receive_feedback} / 5
10. Resolução de Problemas: ${scores.problem_solving} / 5
11. Autonomia: ${scores.autonomy} / 5
12. Apresentar: ${scores.presentation} / 5

RESPOSTAS QUALITATIVAS DISCURSIVAS:
1. Qual é sua principal expectativa ao participar desse programa de mentoria?
"${discursive.expectation || "Sem resposta preenchida."}"

2. Quais habilidades ou competências específicas você gostaria de desenvolver ou aprimorar?
"${discursive.skills || "Sem resposta preenchida."}"

3. Quais desafios você tem enfrentado atualmente em seus projetos de UX?
"${discursive.challenges || "Sem resposta preenchida."}"

4. Com base em suas atividades recentes, identifique os principais gaps de conhecimento ou práticas que você percebe em seu trabalho de UX.
"${discursive.gaps || "Sem resposta preenchida."}"`;

    const payload = {
      contents: [{ parts: [{ text: userPrompt }] }],
      systemInstruction: { parts: [{ text: systemPrompt }] }
    };

    try {
  const data = await fetchGeminiReport(payload);

  if (!data || !data.candidates) {
    throw new Error(
      data?.error?.message ||
      "A API Gemini não retornou conteúdo. Aguarde alguns minutos e tente novamente."
    );
  }

  const generatedText = data.candidates[0]?.content?.parts?.[0]?.text;

  if (generatedText) {
    setReport(generatedText);
  } else {
    throw new Error("Resposta inválida recebida da inteligência artificial.");
  }

} catch (err) {
  setError(
    err.message ||
    "Não foi possível gerar o diagnóstico neste momento."
  );
} finally {
  setLoading(false);
}
};

  //Auxiliar para exportar em docx
  const handleExportDocx = async () => {
  if (!report) {
    setError("Nenhum relatório disponível para exportar.");
    return;
  }

  const isTableLine = (line) => line.trim().startsWith("|") && line.trim().endsWith("|");

  const isSeparatorLine = (line) => {
    const cleaned = line
      .trim()
      .replace(/\|/g, "")
      .replace(/:/g, "")
      .replace(/-/g, "")
      .trim();

    return cleaned === "";
  };

  const splitTableRow = (line) =>
    line.trim().split("|").map((cell) => cell.trim()).filter(Boolean);

  const children = [
    new Paragraph({ text: "Relatório Executivo UX/UI", heading: HeadingLevel.TITLE }),
    new Paragraph(`Profissional: ${profile.name || "Não informado"}`),
    new Paragraph(`Cargo: ${profile.role || "Não informado"}`),
    new Paragraph(`Data: ${profile.date || "Não informada"}`),
    new Paragraph("")
  ];

  const lines = report.split("\n");
  let tableRows = [];

  const flushTable = () => {
    if (!tableRows.length) return;

    children.push(
      new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        rows: tableRows.map((row, rowIndex) =>
          new TableRow({
            children: row.map((cell) =>
              new TableCell({
                children: [
                  new Paragraph({
                    children: [
                      new TextRun({
                        text: cell.replace(/\*\*/g, ""),
                        bold: rowIndex === 0,
                        size: 20
                      })
                    ]
                  })
                ]
              })
            )
          })
        )
      })
    );

    children.push(new Paragraph(""));
    tableRows = [];
  };

  lines.forEach((line) => {
    const trimmed = line.trim();

    if (!trimmed || trimmed === "---") {
      flushTable();
      children.push(new Paragraph(""));
      return;
    }

    if (isTableLine(trimmed)) {
      if (!isSeparatorLine(trimmed)) {
        tableRows.push(splitTableRow(trimmed));
      }
      return;
    }

    flushTable();

    const cleanLine = trimmed.replace(/^#{1,3}\s*/, "").replace(/\*\*/g, "");

    if (trimmed.startsWith("# ")) {
      children.push(new Paragraph({ text: cleanLine, heading: HeadingLevel.HEADING_1 }));
      return;
    }

    if (trimmed.startsWith("## ")) {
      children.push(new Paragraph({ text: cleanLine, heading: HeadingLevel.HEADING_2 }));
      return;
    }

    if (trimmed.startsWith("### ")) {
      children.push(new Paragraph({ text: cleanLine, heading: HeadingLevel.HEADING_3 }));
      return;
    }

    children.push(
      new Paragraph({
        children: [new TextRun({ text: cleanLine, size: 22 })]
      })
    );
  });

  flushTable();

  const safeName = (profile.name || "profissional")
    .toLowerCase()
    .replace(/\s+/g, "-")
    .replace(/[^\w-]/g, "");

  const doc = new Document({
    sections: [{ children }]
  });

  const blob = await Packer.toBlob(doc);
  saveAs(blob, `relatorio-${safeName}.docx`);
};

  // Auxiliar para copiar o relatório gerado
  const handleCopyToClipboard = () => {
    const tempTextArea = document.createElement('textarea');
    tempTextArea.value = report;
    document.body.appendChild(tempTextArea);
    tempTextArea.select();
    try {
      document.execCommand('copy');
      showToast("Copiado!", "Relatório executivo copiado para a área de transferência.", "success");
    } catch (err) {
      console.error('Erro ao copiar', err);
    }
    document.body.removeChild(tempTextArea);
  };

  // Parser de Markdown simples para renderizar as tabelas, parágrafos e estilos de texto na UI
  const renderMarkdown = (text) => {
    if (!text) return null;
    
    const lines = text.split('\n');
    let inTable = false;
    let tableHeaders = [];
    let tableRows = [];

    return lines.map((line, idx) => {
      const trimmed = line.trim();

      if (inTable && (!trimmed.startsWith('|') || trimmed.includes('---'))) {
        if (trimmed.includes('---')) return null;
        const headers = tableHeaders;
        const rows = tableRows;
        inTable = false;
        tableHeaders = [];
        tableRows = [];
        return (
          <div key={`table-wrapper-${idx}`} className="overflow-x-auto my-6 border border-slate-200 dark:border-slate-700 rounded-lg shadow-sm">
            <table className="min-w-full divide-y divide-slate-200 dark:divide-slate-700 text-sm">
              <thead className="bg-slate-50 dark:bg-slate-800 text-slate-800 dark:text-slate-200 font-semibold text-left">
                <tr>
                  {headers.map((h, i) => (
                    <th key={i} className="px-4 py-3 border-b border-slate-200 dark:border-slate-700">{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100 dark:divide-slate-800 bg-white dark:bg-slate-900 text-slate-700 dark:text-slate-300">
                {rows.map((row, rIdx) => (
                  <tr key={rIdx} className="hover:bg-slate-50 dark:hover:bg-slate-800/50 transition">
                    {row.map((cell, cIdx) => (
                      <td key={cIdx} className="px-4 py-3 align-top">{cell}</td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        );
      }

      if (trimmed.startsWith('|')) {
        const cells = trimmed.split('|').map(c => c.trim()).filter((c, i, arr) => i > 0 && i < arr.length - 1);
        if (trimmed.includes('---') || cells.every(c => c.match(/^[:-\s]+$/))) {
          return null;
        }
        if (!inTable) {
          inTable = true;
          tableHeaders = cells;
          return null;
        } else {
          tableRows.push(cells);
          return null;
        }
      }

      if (trimmed.startsWith('# ')) {
        return <h1 key={idx} className="text-2xl font-bold text-slate-900 dark:text-white mt-8 mb-4 border-b border-slate-200 dark:border-slate-800 pb-2">{trimmed.replace('# ', '')}</h1>;
      }
      if (trimmed.startsWith('## ')) {
        return <h2 key={idx} className="text-xl font-bold text-slate-800 dark:text-slate-100 mt-6 mb-3">{trimmed.replace('## ', '')}</h2>;
      }
      if (trimmed.startsWith('### ')) {
        return <h3 key={idx} className="text-lg font-semibold text-slate-800 dark:text-slate-200 mt-5 mb-2">{trimmed.replace('### ', '')}</h3>;
      }

      if (trimmed.startsWith('- ') || trimmed.startsWith('* ')) {
        return (
          <ul key={idx} className="list-disc list-inside ml-4 my-2 text-slate-700 dark:text-slate-300">
            <li>{trimmed.substring(2)}</li>
          </ul>
        );
      }
      if (/^\d+\.\s/.test(trimmed)) {
        return (
          <ol key={idx} className="list-decimal list-inside ml-4 my-2 text-slate-700 dark:text-slate-300">
            <li>{trimmed.replace(/^\d+\.\s/, '')}</li>
          </ol>
        );
      }

      if (trimmed === '') {
        return <div key={idx} className="h-2"></div>;
      }

      return (
        <p key={idx} className="my-3 text-slate-700 dark:text-slate-300 leading-relaxed text-[15px]">
          {trimmed}
        </p>
      );
    });
  };

  return (
    <div className="min-h-screen bg-slate-50 dark:bg-slate-950 text-slate-950 dark:text-slate-50 font-sans">
      
      {/* Toast flutuante para mensagens de feedback */}
      {customToast && (
        <div className="fixed bottom-6 left-6 max-w-sm bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 p-4 rounded-xl shadow-2xl z-50 flex items-start gap-3 animate-bounce">
          <div className="bg-indigo-50 dark:bg-indigo-950/40 p-1.5 rounded-lg text-indigo-600">
            <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={2.5} stroke="currentColor" className="w-5 h-5">
              <path strokeLinecap="round" strokeLinejoin="round" d="M9 12.75L11.25 15 15 9.75M21 12c0 1.268-.63 2.39-1.593 3.068a3.745 3.745 0 01-1.043 3.296 3.745 3.745 0 01-3.296 1.043A3.745 3.745 0 0112 21c-1.268 0-2.39-.63-3.068-1.593a3.746 3.746 0 01-3.296-1.043 3.745 3.745 0 01-1.043-3.296A3.745 3.745 0 013 12c0-1.268.63-2.39 1.593-3.068a3.745 3.745 0 011.043-3.296 3.746 3.746 0 013.296-1.043A3.746 3.746 0 0112 3c1.268 0 2.39.63 3.068 1.593a3.746 3.746 0 013.296 1.043 3.746 3.746 0 011.043 3.296A3.745 3.745 0 0121 12z" />
            </svg>
          </div>
          <div>
            <p className="font-bold text-xs text-slate-900 dark:text-white">{customToast.title}</p>
            <p className="text-[11px] text-slate-500 dark:text-slate-400 mt-0.5">{customToast.message}</p>
          </div>
        </div>
      )}

      {/* Header */}
      <header className="bg-white dark:bg-slate-900 border-b border-slate-200 dark:border-slate-800 py-4 px-6 sticky top-0 z-30 shadow-sm">
        <div className="max-w-7xl mx-auto flex flex-col sm:flex-row justify-between items-center gap-4">
          <div className="flex items-center gap-3">
            <div className="bg-indigo-600 text-white p-2.5 rounded-xl shadow-md">
              <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={2} stroke="currentColor" className="w-6 h-6">
                <path strokeLinecap="round" strokeLinejoin="round" d="M9 12.75L11.25 15 15 9.75M21 12c0 1.268-.63 2.39-1.593 3.068a3.745 3.745 0 01-1.043 3.296 3.745 3.745 0 01-3.296 1.043A3.745 3.745 0 0112 21c-1.268 0-2.39-.63-3.068-1.593a3.746 3.746 0 01-3.296-1.043 3.745 3.745 0 01-1.043-3.296A3.745 3.745 0 013 12c0-1.268.63-2.39 1.593-3.068a3.745 3.745 0 011.043-3.296 3.746 3.746 0 013.296-1.043A3.746 3.746 0 0112 3c1.268 0 2.39.63 3.068 1.593a3.746 3.746 0 013.296 1.043 3.746 3.746 0 011.043 3.296A3.745 3.745 0 0121 12z" />
              </svg>
            </div>
            <div>
              <h1 className="font-bold text-lg text-slate-900 dark:text-white leading-tight">Autodiagnóstico UX/UI</h1>
              <p className="text-xs text-slate-500 dark:text-slate-400">Plataforma Avançada de Mapeamento Técnico & Comportamental</p>
            </div>
          </div>
          
          <div className="flex items-center gap-2">
            <button
              onClick={() => setActiveTab('input')}
              className={`px-4 py-2 rounded-lg text-sm font-semibold transition ${
                activeTab === 'input' 
                  ? 'bg-indigo-50 dark:bg-indigo-950/40 text-indigo-600 dark:text-indigo-400 border border-indigo-200 dark:border-indigo-900/50' 
                  : 'text-slate-600 dark:text-slate-300 hover:bg-slate-100 dark:hover:bg-slate-800'
              }`}
            >
              Mapeamento de Dados
            </button>
            <button
              onClick={() => { if (report) setActiveTab('report'); }}
              disabled={!report}
              className={`px-4 py-2 rounded-lg text-sm font-semibold transition flex items-center gap-2 ${
                !report ? 'opacity-50 cursor-not-allowed text-slate-400' :
                activeTab === 'report' 
                  ? 'bg-indigo-50 dark:bg-indigo-950/40 text-indigo-600 dark:text-indigo-400 border border-indigo-200 dark:border-indigo-900/50' 
                  : 'text-slate-600 dark:text-slate-300 hover:bg-slate-100 dark:hover:bg-slate-800'
              }`}
            >
              <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={2} stroke="currentColor" className="w-4 h-4">
                <path strokeLinecap="round" strokeLinejoin="round" d="M19.5 14.25v-2.625a3.375 3.375 0 00-3.375-3.375h-1.5A1.125 1.125 0 0113.5 7.125v-1.5a3.375 3.375 0 00-3.375-3.375H8.25m0 12.75h7.5m-7.5 3H12M10.5 2.25H5.625c-.621 0-1.125.504-1.125 1.125v17.25c0 .621.504 1.125 1.125 1.125h12.75c.621 0 1.125-.504 1.125-1.125V11.25a9 9 0 00-9-9z" />
              </svg>
              Ver Relatório Gerado
            </button>
          </div>
        </div>
      </header>

      <main className="max-w-7xl mx-auto px-4 sm:px-6 py-8">
        {activeTab === 'input' ? (
          <div className="grid grid-cols-1 lg:grid-cols-12 gap-8">
            
            {/* Esquerda: Formulários de Preenchimento e área de Upload */}
            <div className="lg:col-span-7 space-y-8">
              
              {/* ÁREA DE UPLOAD REAL DE PLANILHA (XLSX / CSV) */}
              <div className="bg-gradient-to-br from-indigo-950 to-slate-900 border border-indigo-500/30 rounded-2xl p-6 text-white shadow-xl relative overflow-hidden">
                <div className="absolute top-0 right-0 w-32 h-32 bg-indigo-600/10 rounded-full blur-2xl"></div>
                
                <div className="flex items-center gap-2 mb-4">
                  <span className="text-[10px] font-bold uppercase tracking-wider text-indigo-300 bg-indigo-800/40 px-2.5 py-1 rounded-md">
                    Processamento Inteligente XLS/CSV
                  </span>
                </div>
                
                <h3 className="font-bold text-lg mb-2">Importar planilha de avaliação técnica</h3>
                <p className="text-slate-300 text-xs mb-4">
                  Suba uma planilha contendo as colunas cadastradas de notas (1 a 5) e as respostas textuais dos candidatos. O sistema fará o mapeamento completo na memória local de forma privada.
                </p>

                <div className="border-2 border-dashed border-indigo-400/30 hover:border-indigo-400/70 rounded-xl p-5 text-center transition bg-indigo-950/40 relative">
                  <input
                    type="file"
                    accept=".xlsx, .xls, .csv"
                    onChange={handleFileChange}
                    className="absolute inset-0 w-full h-full opacity-0 cursor-pointer"
                  />
                  <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={1.5} stroke="currentColor" className="w-8 h-8 mx-auto text-indigo-300 mb-2">
                    <path strokeLinecap="round" strokeLinejoin="round" d="M12 16.5V9.75m0 0l3 3m-3-3l-3 3M6.75 19.5a4.5 4.5 0 01-1.41-8.775 5.25 5.25 0 0110.233-2.33 3 3 0 013.758 3.848A3.752 3.752 0 0118 19.5H6.75z" />
                  </svg>
                  <p className="text-xs font-semibold">Arraste a planilha aqui ou clique para selecionar</p>
                  <p className="text-[10px] text-slate-400 mt-1">Compatível com Excel (.xlsx, .xls) ou delimitado por vírgulas (.csv)</p>
                </div>

                {/* Seletor quando múltiplos profissionais forem encontrados na Planilha */}
                {multipleProfessionals.length > 0 && (
                  <div className="mt-5 pt-4 border-t border-indigo-500/20">
                    <p className="text-xs font-bold text-indigo-300 mb-2.5 flex items-center gap-1.5">
                      <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={2} stroke="currentColor" className="w-3.5 h-3.5">
                        <path strokeLinecap="round" strokeLinejoin="round" d="M15 19.128a9.38 9.38 0 002.625.372 9.337 9.337 0 004.121-.952 4.125 4.125 0 00-7.533-2.493M15 19.128v-.003c0-1.113-.285-2.16-.786-3.07M15 19.128v.109A11.386 11.386 0 0110.081 12H15m-1.5-1.5H1.5A1.5 1.5 0 000 12v3c0 .828.672 1.5 1.5 1.5h1.318a4.502 4.502 0 0010.364 0H15" />
                      </svg>
                      Selecione o profissional para o cockpit ({multipleProfessionals.length} encontrados):
                    </p>
                    <div className="flex gap-2 overflow-x-auto pb-2 scrollbar-thin">
                      {multipleProfessionals.map((prof, idx) => (
                        <button
                          key={idx}
                          onClick={() => handleSelectProfessionalFromList(idx)}
                          className={`flex-shrink-0 text-left px-3 py-2 rounded-lg text-xs transition border ${
                            selectedMappingIndex === idx
                              ? 'bg-indigo-600 border-white font-semibold shadow-sm'
                              : 'bg-slate-800 border-slate-700 hover:bg-slate-700 text-slate-300'
                          }`}
                        >
                          <p className="truncate max-w-[140px] font-semibold">{prof.name}</p>
                          <span className="text-[9px] text-indigo-200 block truncate max-w-[140px] mt-0.5">{prof.role}</span>
                        </button>
                      ))}
                    </div>
                  </div>
                )}
              </div>

              {/* Informações Profissionais */}
              <div className="bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 rounded-2xl p-6 shadow-sm space-y-4">
                <div className="border-b border-slate-100 dark:border-slate-800 pb-3">
                  <h3 className="font-bold text-slate-900 dark:text-white text-lg">1. Identificação do Profissional</h3>
                  <p className="text-slate-500 dark:text-slate-400 text-xs">Informações cadastrais e cronograma da devolutiva executiva.</p>
                </div>
                
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                  <div>
                    <label className="block text-xs font-semibold text-slate-700 dark:text-slate-300 mb-1.5">Nome Completo</label>
                    <input
                      type="text"
                      placeholder="Ex: Mariana Mendonça"
                      value={profile.name}
                      onChange={(e) => setProfile(prev => ({ ...prev, name: e.target.value }))}
                      className="w-full bg-slate-50 dark:bg-slate-950 border border-slate-200 dark:border-slate-800 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500"
                    />
                  </div>
                  <div>
                    <label className="block text-xs font-semibold text-slate-700 dark:text-slate-300 mb-1.5">Cargo / Título Atual</label>
                    <input
                      type="text"
                      placeholder="Ex: Designer de Produto Sênior"
                      value={profile.role}
                      onChange={(e) => setProfile(prev => ({ ...prev, role: e.target.value }))}
                      className="w-full bg-slate-50 dark:bg-slate-950 border border-slate-200 dark:border-slate-800 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500"
                    />
                  </div>
                  <div>
                    <label className="block text-xs font-semibold text-slate-700 dark:text-slate-300 mb-1.5">Tempo & Contexto de Experiência</label>
                    <input
                      type="text"
                      placeholder="Ex: 5 anos em e-commerce B2B"
                      value={profile.experience}
                      onChange={(e) => setProfile(prev => ({ ...prev, experience: e.target.value }))}
                      className="w-full bg-slate-50 dark:bg-slate-950 border border-slate-200 dark:border-slate-800 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500"
                    />
                  </div>
                  <div>
                    <label className="block text-xs font-semibold text-slate-700 dark:text-slate-300 mb-1.5">Data da Avaliação <span className="text-red-500">*</span></label>
                    <input
                      type="date"
                      value={profile.date}
                      onChange={(e) => setProfile(prev => ({ ...prev, date: e.target.value }))}
                      className="w-full bg-slate-50 dark:bg-slate-950 border border-indigo-200 dark:border-indigo-900 focus:border-indigo-500 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500 font-semibold text-indigo-600 dark:text-indigo-400"
                    />
                  </div>
                </div>
              </div>

              {/* Avaliação Quantitativa - 12 Competências */}
              <div className="bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 rounded-2xl p-6 shadow-sm space-y-5">
                <div className="border-b border-slate-100 dark:border-slate-800 pb-3">
                  <h3 className="font-bold text-slate-900 dark:text-white text-lg">2. Autoavaliação Técnica & Comportamental (1 a 5)</h3>
                  <p className="text-slate-500 dark:text-slate-400 text-xs">Arraste para calibrar cada competência avaliada. O gráfico radar se ajustará em tempo real.</p>
                </div>

                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                  {COMPETENCIES.map((comp) => (
                    <div key={comp.id} className="bg-slate-50 dark:bg-slate-950/40 p-3.5 rounded-xl border border-slate-100 dark:border-slate-800/60 flex flex-col justify-between">
                      <div className="flex justify-between items-start mb-2">
                        <div>
                          <span className="font-semibold text-xs text-slate-800 dark:text-slate-200">{comp.label}</span>
                          <p className="text-[10px] text-slate-400 dark:text-slate-500 line-clamp-1">{comp.desc}</p>
                        </div>
                        <span className="bg-indigo-100 dark:bg-indigo-950/80 text-indigo-700 dark:text-indigo-300 font-bold px-2 py-0.5 rounded text-xs ml-2">
                          {scores[comp.id] || 1}
                        </span>
                      </div>
                      <input
                        type="range"
                        min="1"
                        max="5"
                        step="1"
                        value={scores[comp.id] || 1}
                        onChange={(e) => handleScoreChange(comp.id, e.target.value)}
                        className="w-full accent-indigo-600 cursor-pointer h-1.5 bg-slate-200 dark:bg-slate-800 rounded-lg"
                      />
                    </div>
                  ))}
                </div>
              </div>

              {/* Respostas Qualitativas */}
              <div className="bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 rounded-2xl p-6 shadow-sm space-y-5">
                <div className="border-b border-slate-100 dark:border-slate-800 pb-3">
                  <h3 className="font-bold text-slate-900 dark:text-white text-lg">3. Respostas Discursivas (Análise de Projeto e Postura)</h3>
                  <p className="text-slate-500 dark:text-slate-400 text-xs">Forneça as respostas literais para que o mentor cruze as justificativas com as notas do radar.</p>
                </div>

               <div className="space-y-4">
                <div>
                  <label className="block text-xs font-semibold text-slate-700 dark:text-slate-300 mb-1.5">
                    Qual é sua principal expectativa ao participar desse programa de mentoria?
                  </label>
                  <textarea
                    rows="3"
                    placeholder="Descreva sua expectativa principal com a mentoria..."
                    value={discursive.expectation}
                    onChange={(e) => handleDiscursiveChange('expectation', e.target.value)}
                    className="w-full bg-slate-50 dark:bg-slate-950 border border-slate-200 dark:border-slate-800 rounded-lg p-3 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500"
                  />
                </div>

                <div>
                  <label className="block text-xs font-semibold text-slate-700 dark:text-slate-300 mb-1.5">
                    Quais habilidades ou competências específicas você gostaria de desenvolver ou aprimorar?
                  </label>
                  <textarea
                    rows="3"
                    placeholder="Descreva habilidades ou competências que deseja desenvolver..."
                    value={discursive.skills}
                    onChange={(e) => handleDiscursiveChange('skills', e.target.value)}
                    className="w-full bg-slate-50 dark:bg-slate-950 border border-slate-200 dark:border-slate-800 rounded-lg p-3 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500"
                  />
                </div>

                <div>
                  <label className="block text-xs font-semibold text-slate-700 dark:text-slate-300 mb-1.5">
                    Quais desafios você tem enfrentado atualmente em seus projetos de UX?
                  </label>
                  <textarea
                    rows="3"
                    placeholder="Descreva os principais desafios enfrentados em projetos de UX..."
                    value={discursive.challenges}
                    onChange={(e) => handleDiscursiveChange('challenges', e.target.value)}
                    className="w-full bg-slate-50 dark:bg-slate-950 border border-slate-200 dark:border-slate-800 rounded-lg p-3 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500"
                  />
                </div>

                <div>
                  <label className="block text-xs font-semibold text-slate-700 dark:text-slate-300 mb-1.5">
                    Com base em suas atividades recentes, identifique os principais gaps de conhecimento ou práticas que você percebe em seu trabalho de UX.
                  </label>
                  <textarea
                    rows="3"
                    placeholder="Descreva gaps de conhecimento, práticas ou maturidade profissional percebidos..."
                    value={discursive.gaps}
                    onChange={(e) => handleDiscursiveChange('gaps', e.target.value)}
                    className="w-full bg-slate-50 dark:bg-slate-950 border border-slate-200 dark:border-slate-800 rounded-lg p-3 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500"
                  />
                </div>
              </div>
                            </div>

              {/* Botão de Geração */}
              <button
                onClick={generateReport}
                disabled={loading}
                className="w-full bg-indigo-600 hover:bg-indigo-700 disabled:bg-slate-300 dark:disabled:bg-slate-800 text-white font-bold py-4 px-6 rounded-2xl shadow-lg transition flex items-center justify-center gap-3 text-base"
              >
                {loading ? (
                  <>
                    <svg className="animate-spin h-5 w-5 text-white" fill="none" viewBox="0 0 24 24">
                      <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4" />
                      <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z" />
                    </svg>
                    Mentor Analisando & Redigindo Parecer...
                  </>
                ) : (
                  <>
                    <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={2.5} stroke="currentColor" className="w-5 h-5">
                      <path strokeLinecap="round" strokeLinejoin="round" d="M9.813 15.904L9 21l8.982-8.982M18 13.612V6a2.25 2.25 0 00-2.25-2.25H9.13a2.25 2.25 0 00-2.25 2.25v12a2.25 2.25 0 002.25 2.25h1.384m7.462-7.962L18 13.612m0 0l-3.04-3.04M19.5 8.25L12 15.75" />
                    </svg>
                    Gerar Diagnóstico Executivo
                  </>
                )}
              </button>
            </div>

            {/* Direita: Radar Chart Decagonal e Visualizador Lateral */}
            <div className="lg:col-span-5 space-y-6">
              <div className="bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 rounded-2xl p-6 shadow-sm sticky top-28">
                <div className="border-b border-slate-100 dark:border-slate-800 pb-3 mb-6">
                  <h3 className="font-bold text-slate-900 dark:text-white text-base">Radar de Competências (12 Eixos)</h3>
                  <p className="text-slate-500 dark:text-slate-400 text-xs">Mapeamento multidimensional de {profile.name || "Profissional"}.</p>
                </div>

                <div className="flex justify-center items-center">
                  <svg width={size} height={size} className="overflow-visible">
                    {/* Linhas do Grid Radial Decagonal */}
                    {levels.map((level) => {
                      const levelPoints = COMPETENCIES.map((_, idx) => {
                        const { x, y } = getCoordinates(idx, level);
                        return `${x},${y}`;
                      }).join(' ');

                      return (
                        <polygon
                          key={`level-${level}`}
                          points={levelPoints}
                          fill="none"
                          stroke="rgba(99, 102, 241, 0.15)"
                          strokeWidth="1"
                        />
                      );
                    })}

                    {/* Linhas dos Eixos Radiais */}
                    {COMPETENCIES.map((_, idx) => {
                      const maxCoord = getCoordinates(idx, 5);
                      return (
                        <line
                          key={`axis-${idx}`}
                          x1={center}
                          y1={center}
                          x2={maxCoord.x}
                          y2={maxCoord.y}
                          stroke="rgba(148, 163, 184, 0.25)"
                          strokeWidth="1.2"
                          strokeDasharray="2,2"
                        />
                      );
                    })}

                    {/* Área do Polígono das Notas do Candidato */}
                    <polygon
                      points={points}
                      fill="rgba(99, 102, 241, 0.2)"
                      stroke="rgba(99, 102, 241, 0.85)"
                      strokeWidth="2.5"
                    />

                    {/* Pontos de Interseção */}
                    {COMPETENCIES.map((comp, idx) => {
                      const score = scores[comp.id] || 1;
                      const { x, y } = getCoordinates(idx, score);
                      return (
                        <circle
                          key={`dot-${idx}`}
                          cx={x}
                          cy={y}
                          r="5"
                          className="fill-indigo-600 dark:fill-indigo-400 stroke-white dark:stroke-slate-900"
                          strokeWidth="1.5"
                        />
                      );
                    })}

                    {/* Legendas Externas */}
                    {COMPETENCIES.map((comp, idx) => {
                      const labelPos = getCoordinates(idx, 5.6);
                      const isLeft = labelPos.x < center;
                      const isCenter = Math.abs(labelPos.x - center) < 10;
                      const textAnchor = isCenter ? "middle" : isLeft ? "end" : "start";

                      return (
                        <text
                          key={`label-${idx}`}
                          x={labelPos.x}
                          y={labelPos.y + 3}
                          className="text-[9px] font-bold fill-slate-700 dark:fill-slate-300"
                          textAnchor={textAnchor}
                        >
                          {comp.label} ({scores[comp.id]})
                        </text>
                      );
                    })}
                  </svg>
                </div>

                <div className="mt-6 bg-indigo-50/50 dark:bg-indigo-950/20 rounded-xl p-4 border border-indigo-100/50 dark:border-indigo-950">
                  <h4 className="font-semibold text-xs text-indigo-900 dark:text-indigo-400 mb-1.5 flex items-center gap-1.5">
                    <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={2} stroke="currentColor" className="w-4 h-4">
                      <path strokeLinecap="round" strokeLinejoin="round" d="M11.25 11.25l.041-.02a.75.75 0 111.063 1.062L10.83 13.84a.75.75 0 01-1.06 0l-2.06-2.06a.75.75 0 011.06-1.06l1.53 1.53 2.45-2.45z" />
                    </svg>
                    Diagnóstico Multidimensional Ativo
                  </h4>
                  <p className="text-[11px] text-indigo-950 dark:text-indigo-300 leading-relaxed">
                    Com doze competências estruturadas, você obtém uma teia muito mais precisa e fidedigna. O mentor cruzará esses dados para indicar a prontidão de transição em Y (Especialista vs. Gestão).
                  </p>
                </div>
              </div>
            </div>

          </div>
        ) : (
          <div className="bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 rounded-3xl p-6 sm:p-8 shadow-sm space-y-6">
            
            {/* Controles do Relatório */}
            <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4 pb-4 border-b border-slate-200 dark:border-slate-800">
              <div>
                <span className="text-[10px] font-bold uppercase tracking-wider text-emerald-600 dark:text-emerald-400 bg-emerald-50 dark:bg-emerald-950/50 px-2.5 py-1 rounded-md">Status: Diagnóstico Concluído</span>
                <h2 className="text-xl font-bold text-slate-900 dark:text-white mt-2">Relatório Executivo de Mentoria UX/UI</h2>
                <p className="text-xs text-slate-500 dark:text-slate-400">Processado com base no cruzamento qualitativo-quantitativo de {profile.name}.</p>
              </div>

              <div className="flex items-center gap-2 w-full sm:w-auto">
                  <button
                    onClick={handleCopyToClipboard}
                    className="flex-1 sm:flex-initial bg-slate-100 hover:bg-slate-200 dark:bg-slate-800 dark:hover:bg-slate-700 text-slate-700 dark:text-slate-200 px-4 py-2 rounded-xl text-xs font-semibold transition flex items-center justify-center gap-1.5"
                  >
                    Copiar Markdown
                  </button>

                  <button
                    onClick={() => window.print()}
                    className="flex-1 sm:flex-initial bg-slate-100 hover:bg-slate-200 dark:bg-slate-800 dark:hover:bg-slate-700 text-slate-700 dark:text-slate-200 px-4 py-2 rounded-xl text-xs font-semibold transition flex items-center justify-center gap-1.5"
                  >
                    Imprimir / PDF
                  </button>

                  <button
                    onClick={handleExportDocx}
                    className="flex-1 sm:flex-initial bg-slate-100 hover:bg-slate-200 dark:bg-slate-800 dark:hover:bg-slate-700 text-slate-700 dark:text-slate-200 px-4 py-2 rounded-xl text-xs font-semibold transition flex items-center justify-center gap-1.5"
                  >
                    Exportar DOCX
                  </button>

              </div>
            </div>

            {/* Radar dentro do relatório para PDF */}
<div className="bg-slate-50 dark:bg-slate-950/40 border border-slate-200 dark:border-slate-800 rounded-2xl p-6 mb-8">

  <div className="mb-5">
    <h3 className="font-bold text-slate-900 dark:text-white text-lg">
      Gráfico Radar de Competências
    </h3>

    <p className="text-xs text-slate-500 dark:text-slate-400 mt-1">
      Visualização quantitativa das 12 competências avaliadas.
    </p>
  </div>

  <div className="flex justify-center">
    <svg width={420} height={420} className="overflow-visible">

      {/* Grid */}
      {levels.map((level) => {
        const levelPoints = COMPETENCIES.map((_, idx) => {
          const { x, y } = getCoordinates(idx, level);
          return `${x},${y}`;
        }).join(' ');

        return (
          <polygon
            key={`report-level-${level}`}
            points={levelPoints}
            fill="none"
            stroke="rgba(99, 102, 241, 0.15)"
            strokeWidth="1"
          />
        );
      })}

      {/* Eixos */}
      {COMPETENCIES.map((_, idx) => {
        const maxCoord = getCoordinates(idx, 5);

        return (
          <line
            key={`report-axis-${idx}`}
            x1={center}
            y1={center}
            x2={maxCoord.x}
            y2={maxCoord.y}
            stroke="rgba(148, 163, 184, 0.25)"
            strokeWidth="1.2"
            strokeDasharray="2,2"
          />
        );
      })}

      {/* Polígono */}
      <polygon
        points={points}
        fill="rgba(99, 102, 241, 0.2)"
        stroke="rgba(99, 102, 241, 0.85)"
        strokeWidth="2.5"
      />

      {/* Pontos */}
      {COMPETENCIES.map((comp, idx) => {
        const score = scores[comp.id] || 1;
        const { x, y } = getCoordinates(idx, score);

        return (
          <circle
            key={`report-dot-${idx}`}
            cx={x}
            cy={y}
            r="5"
            className="fill-indigo-600 dark:fill-indigo-400 stroke-white dark:stroke-slate-900"
            strokeWidth="1.5"
          />
        );
      })}

      {/* Labels */}
      {COMPETENCIES.map((comp, idx) => {
        const labelPos = getCoordinates(idx, 5.6);

        const isLeft = labelPos.x < center;
        const isCenter = Math.abs(labelPos.x - center) < 10;

        const textAnchor = isCenter
          ? "middle"
          : isLeft
          ? "end"
          : "start";

        return (
          <text
            key={`report-label-${idx}`}
            x={labelPos.x}
            y={labelPos.y + 3}
            className="text-[9px] font-bold fill-slate-700 dark:fill-slate-300"
            textAnchor={textAnchor}
          >
            {comp.label} ({scores[comp.id]})
          </text>
        );
      })}
    </svg>
  </div>
</div>

{/* Renderização do Texto */}
<div className="prose prose-slate dark:prose-invert max-w-none">
              {report ? (
                renderMarkdown(report)
              ) : (
                <div className="text-center py-12 text-slate-500 dark:text-slate-400">
                  Nenhum relatório gerado ainda. Retorne à aba de dados para preencher as informações e clicar em "Gerar".
                </div>
              )}
            </div>

            <div className="pt-6 border-t border-slate-100 dark:border-slate-800 flex flex-col sm:flex-row justify-between text-xs text-slate-400 gap-2">
              <span>Avaliação gerada pelo Mentor de Carreira de Alta Senioridade.</span>
              <span>Data de Referência da Sessão: {profile.date}</span>
            </div>
          </div>
        )}

        {/* Estado de carregamento */}
        {loading && (
          <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-50 flex items-center justify-center p-4">
            <div className="bg-white dark:bg-slate-900 rounded-3xl p-8 max-w-md w-full border border-slate-200 dark:border-slate-800 text-center space-y-4 shadow-2xl">
              <div className="relative w-20 h-20 mx-auto flex items-center justify-center">
                <div className="absolute inset-0 rounded-full border-4 border-indigo-100 dark:border-indigo-950"></div>
                <div className="absolute inset-0 rounded-full border-4 border-indigo-600 border-t-transparent animate-spin"></div>
                <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={2} stroke="currentColor" className="w-8 h-8 text-indigo-600 dark:text-indigo-400">
                  <path strokeLinecap="round" strokeLinejoin="round" d="M12 18a3.75 3.75 0 00.495-7.467 5.99 5.99 0 00-1.925 3.546 5.974 5.974 0 01-2.133-1A3.75 3.75 0 0012 18z" />
                  <path strokeLinecap="round" strokeLinejoin="round" d="M19.5 12a7.5 7.5 0 11-15 0 7.5 7.5 0 0115 0z" />
                </svg>
              </div>
              <h3 className="font-bold text-slate-950 dark:text-white text-lg">Mapeando Maturidade...</h3>
              <p className="text-slate-500 dark:text-slate-400 text-xs text-center">
                O mentor sênior está decodificando as 10 notas quantitativas e cruzando dados qualitativos para estruturar as tabelas executivas e o plano de ação personalizado.
              </p>
            </div>
          </div>
        )}

        

        {/* Notificação de Erro */}
        {error && (
          <div className="fixed bottom-6 right-6 max-w-sm bg-red-50 dark:bg-red-950/40 border border-red-200 dark:border-red-900/50 p-4 rounded-xl shadow-lg z-50 flex gap-3">
            <div className="text-red-500 mt-0.5">
              <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={2} stroke="currentColor" className="w-5 h-5">
                <path strokeLinecap="round" strokeLinejoin="round" d="M12 9v3.75m9-.75a9 9 0 11-18 0 9 9 0 0118 0zm-9 3.75h.008v.008H12v-.008z" />
              </svg>
            </div>
            <div>
              <p className="font-bold text-xs text-red-800 dark:text-red-300">Aviso do Sistema</p>
              <p className="text-[10px] text-red-600 dark:text-red-400 mt-1">{error}</p>
              <button 
                onClick={() => setError(null)} 
                className="text-[10px] underline font-bold text-red-800 dark:text-red-300 mt-2 block"
              >
                Dispensar
              </button>
            </div>
          </div>
        )}

      </main>

    </div>
  );
}
