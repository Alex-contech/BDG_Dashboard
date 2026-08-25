/* ============================================================
   pdf.js — BDG · Controle de Cotações (bdg-cotacoes-staging)
   Fase Extra · Etapa 3 — Geração automática de PDF

   DECISÃO REGISTRADA — biblioteca de geração de PDF no navegador:
   o projeto já carrega jsPDF + jspdf-autotable via CDN no index.html
   (usados hoje pelo botão "Exportar PDF" da listagem, em cotacoes.js
   → document.getElementById('btnExportPdf')). Este módulo reaproveita
   o mesmo par (window.jspdf.jsPDF + doc.autoTable).

   v2 — layout redesenhado a partir dos modelos reais fornecidos pelo
   gestor (cotações SDS_MAO_IA/IM em PDF, aéreas e marítimas): capa com
   faixa azul + "CONFIDENCIAL", bloco CUSTOMER/DATE, tabela de rota em
   caixa (estilo "PORT OF LOADING / PORT OF DISCHARGE..."), tabela de
   custo no formato Kg/Rate/Cost/Service + ALL IN, bloco PER HAWB /
   Total Collect, PAYMENT TERM / SPREAD e as Condições Gerais completas
   (texto institucional real, não mais placeholder), com paginação
   automática quando o texto jurídico não cabe numa página.

   Responsabilidade única deste arquivo: a partir de uma cotação já
   salva no Firebase (o mesmo objeto `d` usado em renderTable), gerar
   um PDF de UMA cotação, replicando o mapeamento campo a campo
   documentado na Análise Técnica (Parte 6 — "Do Excel ao PDF").

   O que NUNCA entra neste PDF (Análise Técnica, Parte 6 — "O que NUNCA
   aparece no PDF"): tarifa de custo (AP), margem (GAP%/GAP$) e o bloco
   contábil interno (Due origin, Due 302D, Ops Result MAO, P/S DT) —
   este módulo só recebe e imprime os campos já comerciais/finais salvos
   na cotação, nunca os dados internos de margem.
   ============================================================ */

const NAVY = [10, 20, 64];
const INK = [30, 35, 60];
const GREY = [110, 118, 150];

/* Modais marítimos usam a nomenclatura de porto/contêiner; os demais
   (AIR, SEA/AIR) usam a nomenclatura de aeroporto/HAWB — mesma divisão
   já usada na Análise Técnica, Parte 7.3 vs. as demais partes. */
function isOceanModal(modal) {
  return /OCEAN/i.test(String(modal || ''));
}

/* ---------- Condições Gerais — texto institucional real -----------
   Consolidado a partir dos modelos de cotação SDS fornecidos (blocos
   AIR IMPORT e OCEAN IMPORT), condensado para caber no documento
   gerado pelo app, preservando as cláusulas essenciais de cada família.
   PAYMENT TERM e SPREAD continuam editáveis por cotação (campos
   d.paymentTerm / d.spread); quando não informados, usa-se o padrão
   observado nos modelos analisados. */
const CONDICOES_GERAIS_AIR = [
  'Após o aceite pelo cliente, as condições apresentadas têm validade conforme prazo informado na proposta. A partir do aceite do serviço, a alocação é submetida para prévia aprovação da companhia aérea.',
  'Valores informados nesta proposta devem ser considerados como VATOS (Validity At Time Of Shipment) e estão sujeitos a alteração caso os dados da carga recepcionada para embarque sejam diferentes dos informados nesta cotação.',
  'Devem ser consultados previamente casos em que o peso ultrapasse 3.000 kg, o comprimento por volume ultrapasse 3.000 mm, o volume ultrapasse 7 m³, a altura ultrapasse 2.000 mm, ou a carga não seja empilhável/remontada.',
  'Cargas classificadas como bagagem, bebidas alcoólicas, produto químico (perigoso ou não) e derivados alimentícios ou animais exigem detalhamento prévio para verificação de aceitação na origem e no destino, podendo gerar custo adicional.',
  'Seguindo a regulamentação IATA, o frete é aplicado sobre o que for maior entre o peso bruto e o peso volumétrico (1 m³ = 167 kg no transporte aéreo).',
  'O transit time é o tempo estimado do transporte internacional (voo), podendo ser alterado sem aviso prévio pela companhia aérea, sem responsabilidade da SAMSUNG SDS por atrasos decorrentes de fatores fora de seu controle.',
  'Toda carga deve ser entregue em embalagem apropriada, de acordo com a Instrução Normativa 32/15 quanto ao tratamento de embalagens de madeira.',
  'Taxas de origem, despachante e inland são de responsabilidade do exportador, exceto quando o Incoterm for EXW.',
  'A responsabilidade da SAMSUNG SDS é limitada aos valores das Convenções Internacionais aplicáveis, conforme informado no Airway Bill (AWB) e/ou Bill of Lading (BL). A SAMSUNG SDS atua como Agente de Cargas.',
  'Eventuais valores em moeda estrangeira serão pagos em Reais, mediante conversão pela taxa PTAX de venda do dia anterior ao faturamento, publicada pelo Banco Central do Brasil, acrescida do spread informado nesta cotação.',
  'Aplicam-se a esta proposta todos os acordos internacionais e a legislação de comércio exterior correlata, bem como o documento "Termos e Condições Comerciais" vigente, parte integrante desta proposta.',
];

const CONDICOES_GERAIS_OCEAN = [
  'Após o aceite pelo cliente, as condições apresentadas têm validade conforme prazo de embarque informado na proposta. A partir do aceite do serviço, a alocação é submetida para prévia aprovação do armador.',
  'Valores informados nesta proposta devem ser considerados como VATOS (Validity At Time Of Shipment) e estão sujeitos a alteração caso os dados da carga recepcionada para embarque sejam diferentes dos informados nesta cotação.',
  'Para embarques LCL, os valores informados são válidos para cargas com peso e dimensões regulares em contêiner 20\' dry, sem restrições; a necessidade de contêiner 40\' dry ou HC deve ser informada previamente para consulta de adicionais.',
  'Devem ser consultados previamente casos em que o peso ultrapasse 3.000 kg, o comprimento por volume ultrapasse 3.000 mm, o volume ultrapasse 7 m³, a altura ultrapasse 2.000 mm, ou a carga não seja empilhável/remontada.',
  'O transit time é o tempo estimado informado pelo armador, podendo sofrer alteração em razão de problemas climáticos, greves, atrasos de conexão ou mudança de rota, sem responsabilidade da SAMSUNG SDS.',
  'Toda carga deve ser entregue em embalagem apropriada para embarque marítimo, de acordo com a Instrução Normativa 32/15 quanto ao tratamento de embalagens de madeira.',
  'Taxas de origem, despachante e inland são de responsabilidade do exportador, exceto quando o Incoterm for EXW. Taxas de destino são cobradas pelo valor vigente na data de atracação do navio, sujeito a alteração sem aviso prévio.',
  'Bookings estão sujeitos à disponibilidade de espaço e equipamento. O shipper é responsável por providenciar o VGM (Verified Gross Mass), sendo toda informação contida no VGM de inteira responsabilidade do embarcador.',
  'A responsabilidade da SAMSUNG SDS é limitada aos valores das Convenções Internacionais aplicáveis, conforme informado no Bill of Lading (BL). A SAMSUNG SDS atua como Agente de Cargas.',
  'Eventuais valores em moeda estrangeira serão pagos em Reais, mediante conversão pela taxa PTAX de venda do dia anterior ao faturamento, publicada pelo Banco Central do Brasil, acrescida do spread informado nesta cotação.',
  'Aplicam-se a esta proposta todos os acordos internacionais e a legislação de comércio exterior correlata, bem como o documento "Termos e Condições Comerciais" vigente, parte integrante desta proposta.',
];

function fmt(n) {
  if (n === undefined || n === null || !Number.isFinite(n)) return '—';
  return n.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function fmtMoney(n) {
  if (n === undefined || n === null || !Number.isFinite(n)) return '—';
  return 'USD ' + fmt(n);
}

function fmtDate(v) {
  if (!v) return '—';
  const s = String(v).slice(0, 10);
  const m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  return m ? `${m[3]}/${m[2]}/${m[1]}` : s;
}

/* ---------- cabeçalho de página (capa) ---------- */
function drawHeader(doc, d, pageWidth) {
  doc.setTextColor(140, 148, 180);
  doc.setFont('helvetica', 'normal'); doc.setFontSize(7.5);
  doc.text('CONFIDENCIAL', 24, 20);
  doc.setTextColor(...NAVY);
  doc.setFont('helvetica', 'bold'); doc.setFontSize(12);
  doc.text('SAMSUNG SDS', pageWidth - 24, 18, { align: 'right' });
  doc.setFont('helvetica', 'normal'); doc.setFontSize(7.5);
  doc.setTextColor(...GREY);
  doc.text('Realize your vision', pageWidth - 24, 26, { align: 'right' });

  doc.setDrawColor(...NAVY);
  doc.setLineWidth(0.6);
  doc.line(24, 32, pageWidth - 24, 32);

  doc.setFillColor(...NAVY);
  doc.rect(0, 40, pageWidth, 56, 'F');
  doc.setTextColor(255, 255, 255);
  doc.setFont('helvetica', 'bold'); doc.setFontSize(14);
  doc.text(`QUOTATION SDS ${d.ref || '—'}`, 24, 62);
  doc.setFont('helvetica', 'normal'); doc.setFontSize(10);
  doc.setTextColor(180, 200, 255);
  const servicoLabel = isOceanModal(d.modal) ? 'OCEAN IMPORT/EXPORT' : (d.modal === 'SEA/AIR' ? 'SEA & AIR' : 'AIR IMPORT/EXPORT');
  doc.text(`${d.type || 'Import'} · ${servicoLabel}`, 24, 78);

  doc.setTextColor(...INK);
  doc.setFont('helvetica', 'bold'); doc.setFontSize(9);
  doc.text('CUSTOMER:', 24, 108);
  doc.text('DATE:', pageWidth / 2, 108);
  doc.setFont('helvetica', 'normal');
  doc.text(String(d.client || '—'), 90, 108);
  doc.text(fmtDate(d.week), pageWidth / 2 + 40, 108);

  return 122; // y inicial disponível após o cabeçalho
}

function drawFooter(doc, pageWidth, pageHeight) {
  doc.setFontSize(7.5);
  doc.setTextColor(...GREY);
  doc.setFont('helvetica', 'normal');
  doc.text('BDG · Controle de Cotações — Samsung SDS Logistics', 24, pageHeight - 16);
  doc.text(`Pg. ${doc.internal.getNumberOfPages()}`, pageWidth - 24, pageHeight - 16, { align: 'right' });
}

/* tabela em caixa, estilo "label / valor", igual aos modelos reais
   (PORT OF LOADING / PORT OF DISCHARGE / FREQUENCY...) */
function drawBoxTable(doc, startY, rows, pageWidth) {
  doc.autoTable({
    startY,
    theme: 'grid',
    body: rows,
    styles: { fontSize: 9, cellPadding: 6, textColor: INK, lineColor: [200, 205, 225], lineWidth: 0.5 },
    columnStyles: { 0: { fontStyle: 'bold', fillColor: [235, 238, 250], cellWidth: pageWidth * 0.32 } },
    margin: { left: 24, right: 24 },
  });
  return doc.lastAutoTable.finalY;
}

function drawSectionTitle(doc, y, text) {
  doc.setFont('helvetica', 'bold'); doc.setFontSize(10.5);
  doc.setTextColor(...NAVY);
  doc.text(text, 24, y);
  return y + 8;
}

/* ----- um "bloco" de serviço (Kg/Rate/Cost + Per HAWB) -----
   Fase Extra (pedido do usuário) — a mesma referência pode ter mais de
   um serviço fechado ao mesmo tempo (ex.: AIR STD + AIR EXP na
   SDS_MAO_IA_07497, Análise Técnica Parte 7.1). Cada bloco recebe seu
   próprio título de serviço e sua própria tabela de custo/Per HAWB,
   exatamente como no modelo real (páginas 2 e 3 do 07497) — só que
   aqui dentro do mesmo fluxo de página, sem duplicar a capa/rota. */
function drawBlocoServico(doc, d, bloco, pageWidth, y, pageHeight, tituloServico) {
  const ocean = isOceanModal(d.modal);
  const origemMaisMia = (bloco.originCharges || 0) + (bloco.miaCharges || 0);

  y = drawSectionTitle(doc, y, tituloServico);
  const custoRows = ocean ? [
    ['Ocean/Air Freight', fmtMoney(bloco.custoAF)],
    ['Destination Charges (THC/DOC/Handling etc.)', fmtMoney(origemMaisMia)],
    ['Handling Fee', fmtMoney(bloco.handling)],
    ['ALL IN', fmtMoney(bloco.allIn)],
  ] : [
    ['Chargeable Weight (Kg)', fmt(bloco.pesoTarifavel)],
    ['Rate (USD/Kg)', fmt(bloco.rateAR)],
    ['Cost — A/F', fmtMoney(bloco.custoAF)],
    ['FCA/EXW/Origin + MIA Charges', fmtMoney(origemMaisMia)],
    ['Handling Fee', fmtMoney(bloco.handling)],
    ['ALL IN', fmtMoney(bloco.allIn)],
  ];
  doc.autoTable({
    startY: y,
    theme: 'grid',
    body: custoRows,
    styles: { fontSize: 9, cellPadding: 6, textColor: INK, lineColor: [200, 205, 225], lineWidth: 0.5 },
    columnStyles: { 0: { cellWidth: pageWidth * 0.55 }, 1: { halign: 'right', fontStyle: 'bold' } },
    didParseCell: (data) => {
      if (data.row.index === custoRows.length - 1) {
        data.cell.styles.fillColor = [20, 40, 160];
        data.cell.styles.textColor = 255;
        data.cell.styles.fontStyle = 'bold';
      }
    },
    margin: { left: 24, right: 24 },
  });
  y = doc.lastAutoTable.finalY + 20;

  /* Per HAWB / Total Collect / Handling Information — Análise Técnica,
     Parte 5.1: Weight charge = A/F; Total other charges due to Agent =
     FCA/EXW/Origin + MIA charges; Total Collect = soma das duas. */
  y = drawSectionTitle(doc, y, ocean ? 'Total Collect' : 'Per HAWB');
  const hawbRows = [
    ['Weight/Freight Charge', fmtMoney(bloco.custoAF)],
    ['Total Other Charges Due to Agent', fmtMoney(origemMaisMia)],
    ['Total Collect', fmtMoney(bloco.totalCollect)],
    ['Handling Information', fmtMoney(bloco.handling)],
  ];
  doc.autoTable({
    startY: y,
    theme: 'grid',
    body: hawbRows,
    styles: { fontSize: 9, cellPadding: 6, textColor: INK, lineColor: [200, 205, 225], lineWidth: 0.5 },
    columnStyles: { 0: { cellWidth: pageWidth * 0.55 }, 1: { halign: 'right' } },
    margin: { left: 24, right: 24 },
  });
  return doc.lastAutoTable.finalY + 20;
}

/* ---------- gerador principal ----------
   @param {object} d  Cotação já salva (mesmo formato usado em RAW/renderTable).
                       d.blocosExtras (opcional): array de serviços adicionais
                       fechados sob a mesma referência (Bloco 1 = os próprios
                       campos de topo de d; Blocos 2/3 = d.blocosExtras[0/1]).
   @returns {jsPDF}    Instância do documento, já pronta para .save() ou .output() */
export function generatePdfCotacao(d) {
  d = d || {};
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ orientation: 'portrait', unit: 'pt', format: 'a4' });
  const pageWidth = doc.internal.pageSize.getWidth();
  const pageHeight = doc.internal.pageSize.getHeight();
  const ocean = isOceanModal(d.modal);

  let y = drawHeader(doc, d, pageWidth);

  /* ----- Tabela de rota (dados operacionais — nunca vêm do motor de
     cálculo, são digitados/salvos junto com o cadastro da cotação) -----
     Rótulos seguem a nomenclatura dos modelos reais: marítimo → PORT OF
     LOADING/DISCHARGE; aéreo → AIRPORT OF LOADING/DISCHARGE. Comum a
     todos os blocos de serviço desta referência (mesma rota, mesmo
     cliente — só o serviço/custo muda de bloco para bloco). */
  const routeRows = [
    [ocean ? 'PORT OF LOADING' : 'AIRPORT OF LOADING', d.origin || '—'],
    [ocean ? 'PORT OF DISCHARGE' : 'AIRPORT OF DISCHARGE', d.dest || '—'],
    ['INCOTERM', d.incoterm || '—'],
    ['TRANSIT TIME', d.transitTime || '—'],
    ['AGENTE (ORIGEM)', d.agente || '—'],
  ];
  y = drawBoxTable(doc, y + 14, routeRows, pageWidth) + 20;

  /* ----- Blocos de serviço -----
     Bloco 1 = os próprios campos salvos no topo de `d` (compatibilidade
     com toda cotação já existente, criada antes desta funcionalidade).
     Blocos 2/3, se houver, vêm de d.blocosExtras — cada um em sua
     própria página, igual ao padrão real observado no 07497 (AIR STD
     na página 2, AIR EXP na página 3), sob a mesma referência/capa. */
  const modalLabel = d.modal === 'SEA/AIR' ? 'SEA/AIR' : (d.modal || '');
  const blocos = [d, ...(Array.isArray(d.blocosExtras) ? d.blocosExtras : [])];
  blocos.forEach((bloco, i) => {
    if (i > 0) { doc.addPage(); y = 40; }
    /* Auditoria Geral, Bloco 7 — aqui havia um ternário cujos dois ramos
       devolviam exatamente o mesmo texto: `ocean ? 'Tabela de Custo' :
       'Tabela de Custo'`. A condição não decidia nada. Colapsado para a
       constante única, sem mudança de comportamento.
       Se a intenção original era um título diferente para o modal marítimo,
       isso nunca chegou a funcionar — e continua valendo a pena decidir qual
       seria esse título, mas é decisão de negócio, não de código.
       O `|| 'Tabela de Custo'` da linha seguinte também saiu: tituloServico
       já não pode ser vazio, porque o `||` acima garante o mesmo padrão. */
    const tituloServico = `${modalLabel}${bloco.servico ? ' ' + bloco.servico : ''}`.trim() || 'Tabela de Custo';
    y = drawBlocoServico(doc, d, bloco, pageWidth, y, pageHeight, tituloServico);
  });

  if (d.remarks) {
    doc.setFont('helvetica', 'bold'); doc.setFontSize(9.5);
    doc.setTextColor(...NAVY);
    doc.text('Observações', 24, y);
    doc.setFont('helvetica', 'normal'); doc.setFontSize(9);
    doc.setTextColor(...INK);
    const remarksLines = doc.splitTextToSize(String(d.remarks), pageWidth - 48);
    doc.text(remarksLines, 24, y + 14);
    y += 14 + remarksLines.length * 11 + 10;
  }

  /* ----- Payment term / Spread ----- */
  doc.setFont('helvetica', 'bold'); doc.setFontSize(9);
  doc.setTextColor(...INK);
  doc.text(`PAYMENT TERM: ${d.paymentTerm || '30 dias'}`, 24, y);
  doc.text(`SPREAD (PTAX): ${d.spread || '+ 3%'}`, pageWidth / 2 + 20, y);
  y += 20;

  /* ----- Condições gerais (texto institucional real, por família de modal) ----- */
  const condicoes = ocean ? CONDICOES_GERAIS_OCEAN : CONDICOES_GERAIS_AIR;
  if (y > pageHeight - 140) { doc.addPage(); y = 40; }
  doc.setFont('helvetica', 'bold'); doc.setFontSize(10);
  doc.setTextColor(...NAVY);
  doc.text('Condições Gerais', 24, y);
  y += 14;
  doc.setFont('helvetica', 'normal'); doc.setFontSize(8);
  doc.setTextColor(...INK);
  condicoes.forEach((linha) => {
    const lines = doc.splitTextToSize('• ' + linha, pageWidth - 48);
    if (y + lines.length * 10 > pageHeight - 40) { doc.addPage(); y = 40; }
    doc.text(lines, 24, y);
    y += lines.length * 10 + 5;
  });

  /* ----- rodapé em todas as páginas ----- */
  const totalPages = doc.internal.getNumberOfPages();
  for (let p = 1; p <= totalPages; p++) {
    doc.setPage(p);
    drawFooter(doc, pageWidth, pageHeight);
  }

  return doc;
}

/* ---------- atalho para salvar direto (uso no botão "Gerar PDF") ----------
   Fase Extra · Etapa 3 — qualquer erro na geração (ex.: biblioteca jsPDF que
   não carregou por bloqueio de rede/ad-blocker, ou um campo em formato
   inesperado) fica visível para o vendedor em vez de falhar em silêncio —
   antes disso, um clique sem efeito nenhum era indistinguível de "não fiz
   nada". */
export function downloadPdfCotacao(d) {
  try {
    if (!window.jspdf || !window.jspdf.jsPDF) {
      throw new Error('Biblioteca jsPDF não carregou (verifique a conexão ou um bloqueador de anúncios/scripts).');
    }
    const doc = generatePdfCotacao(d);
    const stamp = new Date().toISOString().slice(0, 10);
    doc.save(`Cotacao_${(d && d.ref) || 'BDG'}_${stamp}.pdf`);
  } catch (err) {
    console.error('Falha ao gerar PDF da cotação:', err);
    alert('Não foi possível gerar o PDF: ' + err.message);
  }
}
