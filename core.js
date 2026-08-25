/* eslint-disable security/detect-object-injection --
   Decisão herdada de utils.js: STATUS_FROM_TEXT[s] usa apenas chaves de um
   objeto de sinônimos fixo definido neste próprio arquivo — falso positivo
   já conhecido da regra para esse padrão. */
/* ==========================================================================
   core.js — BDG · Controle de Cotações (Samsung SDS Logistics)
   ==========================================================================
   Consolidação da Auditoria Geral de 01/08/2026 (Bloco 4, Alternativa B).

   Reúne, sem alterar uma linha de lógica, os dois módulos de função pura do
   projeto — nenhum dos dois toca DOM, rede ou Firebase, e ambos já eram
   cobertos pelo mesmo runner de testes:

     utils.js      384 linhas   datas, filtros, KPIs, sparkline, validações
     calculos.js   216 linhas   motor de cálculo aéreo e marítimo

   Efeito colateral aceito e registrado na auditoria: auth.js importa daqui
   apenas checkPasswordRules, e passa a baixar o motor de cálculo junto —
   cerca de 5 KB extras, sem impacto perceptível.

   Todos os export originais foram preservados com o mesmo nome, então os
   importadores só precisam trocar o caminho do módulo.
   ========================================================================== */

/* ==========================================================================
   PARTE 1 de 2 — FUNÇÕES PURAS DE APOIO (antes: utils.js)
   ========================================================================== */
/* ============================================================
   utils.js — BDG · Controle de Cotações (bdg-cotacoes-staging)
   Fase 3 · Qualidade de Código e Testes — Etapa 1 (modularização parcial)

   Funções PURAS (sem acesso a DOM, sem efeitos colaterais) extraídas dos
   dois <script type="module"> do index.html, para permitir cobertura de
   testes automatizados (Jest) e lint isolado (ESLint).

   IMPORTANTE — escopo desta etapa: este arquivo cobre apenas as funções
   puras já previstas no roteiro da Fase 3 (toISODate, validação de senha,
   normalizadores de importação). A modularização completa do app em
   auth.js / cotacoes.js / ui.js (que ainda dependem fortemente do DOM e
   do SDK do Firebase) é um passo seguinte, não incluído aqui.
   ============================================================ */

/* ---------- constantes compartilhadas (única fonte de verdade) ---------- */
export const TYPE_VALUES = ['Import', 'Export', 'Cabotagem', 'Doméstico', 'Desembaraço', 'Puxada de Container'];

/* Fase Extra · Etapa 1 — tipos de serviço identificados na engenharia
   reversa da planilha (Análise Técnica, Parte 7): STD e EXP são dois
   blocos de cálculo independentes dentro do mesmo modal aéreo; S&A (Sea
   & Air) usa o mesmo motor do aéreo aplicado a um serviço combinado.
   Campo opcional — cotações marítimas/rodoviárias podem deixar em branco. */
export const SERVICE_VALUES = ['STD', 'EXP', 'S&A'];

export const STATUS_FROM_TEXT = {
  'APPROVED': 'APPROVED', 'APROVADA': 'APPROVED', 'APROVADO': 'APPROVED',
  'LOST': 'LOST', 'PERDIDA': 'LOST', 'PERDIDO': 'LOST',
  'NO RESPONSE': 'NO RESPONSE', 'SEM RESPOSTA': 'NO RESPONSE',
  'UNDER REVIEW': 'UNDER REVIEW', 'EM REVISAO': 'UNDER REVIEW', 'EM REVISÃO': 'UNDER REVIEW',
  'N/D': 'N/D', 'NAO DEFINIDO': 'N/D', 'NÃO DEFINIDO': 'N/D',
};

/* ---------- datas (Fase 2 — padronização ISO 8601) ----------
   Normaliza qualquer valor de entrada para AAAA-MM-DD e valida contra o
   calendário real (ex.: 2026-02-30 é rejeitado). Retorna null se inválida. */
export function toISODate(v) {
  if (v === undefined || v === null || v === '') return null;
  const s = String(v).trim();
  const m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!m) return null;
  const y = +m[1], mo = +m[2], d = +m[3];
  if (mo < 1 || mo > 12 || d < 1 || d > 31) return null;
  const test = new Date(Date.UTC(y, mo - 1, d));
  if (test.getUTCFullYear() !== y || test.getUTCMonth() !== mo - 1 || test.getUTCDate() !== d) return null;
  return m[1] + '-' + m[2] + '-' + m[3];
}

/* ---------- validação de senha (regras puras, sem tocar no DOM) ----------
   O index.html mantém uma função validatePassword() de mesma assinatura que
   chama esta e só cuida da parte visual (marcar ✓ nas regras na tela). */
export function checkPasswordRules(pwd, confirm) {
  pwd = pwd || '';
  confirm = confirm || '';
  return {
    len: pwd.length >= 8 && pwd.length <= 10,
    charset: /^[A-Za-z0-9]*$/.test(pwd) && pwd.length > 0,
    letters: (pwd.match(/[A-Za-z]/g) || []).length >= 3,
    numbers: (pwd.match(/\d/g) || []).length >= 3,
    repeat: !/(.)\1{5,}/.test(pwd),
    match: pwd.length > 0 && pwd === confirm,
  };
}

/* ---------- normalizadores de importação em massa (planilha Excel) ---------- */
export function normalizeTypeImport(v) {
  const s = String(v || 'Import').trim();
  const found = TYPE_VALUES.find(t => t.toLowerCase() === s.toLowerCase());
  return found || s || 'Import';
}

export function normalizeStatusImport(v) {
  if (!v) return 'UNDER REVIEW';
  const s = String(v).trim().toUpperCase();
  return STATUS_FROM_TEXT[s] || 'UNDER REVIEW';
}

/* Converte formatos de entrada (Date do Excel, número serial, DD/MM/AAAA,
   AAAA-MM-DD) para o texto AAAA-MM-DD. A validação real de calendário é
   sempre feita por toISODate() antes de retornar.
   `parseSerial` é injetado pelo chamador (no index.html: XLSX.SSF.parse_date_code)
   para não acoplar este módulo puro à biblioteca SheetJS/XLSX do navegador —
   isso é o que permite testar esta função no Jest sem carregar o XLSX. */
export function normalizeDateImport(v, parseSerial) {
  if (v === undefined || v === null || v === '') return null;
  if (v instanceof Date) {
    const y = v.getUTCFullYear(), m = v.getUTCMonth() + 1, day = v.getUTCDate();
    if (!y || isNaN(y)) return null;
    return toISODate(y + '-' + String(m).padStart(2, '0') + '-' + String(day).padStart(2, '0'));
  }
  if (typeof v === 'number') {
    const d = parseSerial ? parseSerial(v) : null;
    if (d) return toISODate(d.y + '-' + String(d.m).padStart(2, '0') + '-' + String(d.d).padStart(2, '0'));
    return null;
  }
  const s = String(v).trim();
  if (!s) return null;
  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return toISODate(m[1] + '-' + String(m[2]).padStart(2, '0') + '-' + String(m[3]).padStart(2, '0'));
  // Formato brasileiro DD/MM/AAAA (ou com "-"): padrão adotado para
  // desambiguar, já que a base de usuários é brasileira (corrigido na Fase 2).
  m = s.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{4})$/);
  if (m) return toISODate(m[3] + '-' + String(m[2]).padStart(2, '0') + '-' + String(m[1]).padStart(2, '0'));
  return null;
}

/* ---------- filtros do painel (Fase 3 · Etapa 2 de modularização) ----------
   Antes vivia como getFiltered() dentro do cotacoes.js, fechando sobre as
   variáveis de módulo RAW/state (closure) — não dava pra testar isoladamente.
   Aqui vira uma função pura: recebe as linhas e os filtros como parâmetros,
   devolve um novo array, sem tocar em nada fora dela. */
export function filterQuotations(rows, filters) {
  rows = rows || [];
  filters = filters || {};
  const q = String(filters.search || '').trim().toUpperCase();
  return rows.filter((d) => {
    if (filters.month && d.month !== filters.month) return false;
    if (filters.client && d.client !== filters.client) return false;
    if (filters.seller && d.seller !== filters.seller) return false;
    if (filters.status && d.status !== filters.status) return false;
    if (filters.modal && d.modal !== filters.modal) return false;
    if (filters.type && d.type !== filters.type) return false;
    if (q) {
      const hay = [d.ref, d.client, d.origin, d.dest, d.seller, d.remarks].join(' ').toUpperCase();
      if (!hay.includes(q)) return false;
    }
    return true;
  });
}

/* ---------- KPIs do painel (Fase 3 · Etapa 2 de modularização) ----------
   Antes calculado inline dentro de render() no cotacoes.js, misturado com
   escrita no DOM. Aqui separa o CÁLCULO (puro, testável) da EXIBIÇÃO — o
   cotacoes.js chama computeKpis(data) e só escreve o resultado nos elementos
   da tela, sem repetir a lógica de contagem/porcentagem.
   Em caso de empate no "cliente destaque", desempata pela ordem de primeira
   aparição no array (mesmo comportamento que Object.entries já tinha, já
   que a ordem de inserção de chaves em objetos JS segue a primeira vez que
   a chave aparece). */
/* ---------- campos numéricos comerciais (Fase Extra · Etapa 1) ----------
   O vendedor digita valores como "1.720,18" (padrão BR) ou "1720.18"
   (padrão internacional/colado de outra planilha) — parseDecimalBR aceita
   os dois formatos e devolve um Number ou null (vazio/ inválido nunca
   lança erro, só devolve null, já que estes campos são opcionais na
   versão beta — ver "Regra da versão beta" no plano da Fase Extra).
   formatDecimalBR faz o caminho inverso, para reexibir o valor já salvo
   no formulário de edição no mesmo padrão em que o vendedor digitou. */
export function parseDecimalBR(v) {
  if (v === undefined || v === null) return null;
  let s = String(v).trim();
  if (!s) return null;
  s = s.replace(/[^\d,.\-]/g, ''); // remove "USD", espaços, símbolos de moeda
  if (!s || s === '-') return null;
  const hasComma = s.includes(',');
  const hasDot = s.includes('.');
  if (hasComma && hasDot) {
    // Padrão BR com milhar: "1.720,18" → ponto é milhar, vírgula é decimal
    s = s.replace(/\./g, '').replace(',', '.');
  } else if (hasComma) {
    // Só vírgula: é o separador decimal ("7,53" → 7.53)
    s = s.replace(',', '.');
  }
  // Só ponto (ou nenhum separador): já está no formato aceito por Number()
  const n = Number(s);
  return Number.isFinite(n) ? n : null;
}

export function formatDecimalBR(n, decimals = 2) {
  if (n === undefined || n === null || !Number.isFinite(n)) return '';
  return n.toLocaleString('pt-BR', { minimumFractionDigits: decimals, maximumFractionDigits: decimals });
}

/* ---------- sparkline (Fase 4 · Etapa 2 — mini-gráfico de tendência nos KPIs) ----------
   Função pura: recebe uma série de números (ex.: contagem por mês) e devolve
   a string já pronta para o atributo "points" de uma <polyline> SVG,
   normalizada para caber em width×height. Não toca DOM nem Chart.js — só
   faz a conta; quem desenha é o cotacoes.js (mesmo padrão de computeKpis:
   cálculo puro e testável separado da renderização). */
export function sparklinePoints(values, width = 64, height = 22, padding = 2) {
  const vals = (values || []).map((v) => (Number.isFinite(v) ? v : 0));
  if (vals.length === 0) return '';
  if (vals.length === 1) {
    const y = height / 2;
    return `${padding},${y} ${width - padding},${y}`;
  }
  const min = Math.min(...vals);
  const max = Math.max(...vals);
  const range = max - min || 1;
  const stepX = (width - padding * 2) / (vals.length - 1);
  return vals
    .map((v, i) => {
      const x = padding + i * stepX;
      const y = height - padding - ((v - min) / range) * (height - padding * 2);
      return `${x.toFixed(1)},${y.toFixed(1)}`;
    })
    .join(' ');
}

/* Especificação de melhorias, item 2.1 — percentual de cada categoria em
   relação ao TOTAL GERAL (não ao total já filtrado). totalGeral é opcional
   e por padrão igual a rows.length (mesmo comportamento de antes desta
   mudança) — quem já chama computeKpis(data) sem o segundo argumento
   continua funcionando exatamente como antes. Quando o chamador passa o
   total geral real (ex.: RAW.length em cotacoes.js), os campos *PctTotal
   passam a refletir "quanto isto representa do total geral", distinto de
   *Rate (que já existia e é relativo ao próprio filtro). */
export function computeKpis(rows, totalGeral) {
  rows = rows || [];
  const total = rows.length;
  const base = Number.isFinite(totalGeral) && totalGeral > 0 ? totalGeral : total;
  const approved = rows.filter((d) => d.status === 'APPROVED').length;
  const lost = rows.filter((d) => d.status === 'LOST').length;
  const review = rows.filter((d) => d.status === 'UNDER REVIEW' || d.status === 'NO RESPONSE').length;

  const clientCounts = {};
  rows.forEach((d) => { if (d.client) clientCounts[d.client] = (clientCounts[d.client] || 0) + 1; });
  const clientsInView = Object.keys(clientCounts).length;
  const topEntry = Object.entries(clientCounts).sort((a, b) => b[1] - a[1])[0];

  const pctOfTotal = (part) => (base > 0 ? (part / base) * 100 : null);

  return {
    total,
    approved,
    lost,
    review,
    clientsInView,
    approvedRate: total ? (approved / total) * 100 : null,
    lostRate: total ? (lost / total) * 100 : null,
    topClient: topEntry ? { name: topEntry[0], count: topEntry[1] } : null,
    // Percentual de cada categoria em relação ao total geral (não ao filtro).
    totalPctTotal: pctOfTotal(total),
    approvedPctTotal: pctOfTotal(approved),
    lostPctTotal: pctOfTotal(lost),
    reviewPctTotal: pctOfTotal(review),
  };
}

/* ---------- estatísticas por vendedor (Especificação de melhorias, item 5.1) ----------
   Antes calculado implicitamente só dentro da configuração dos datasets
   do Chart.js (chartSeller/chartSellerRate, cotacoes.js), sem nenhuma
   estrutura de dados própria e testável. Extraído para função pura para
   alimentar a nova tabela complementar aos gráficos, na view Desempenho
   — mesmo padrão de computeKpis: cálculo puro e testado separado da
   renderização (DOM/Chart.js ficam em cotacoes.js). Devolve um array
   ordenado por taxa de conversão (maior primeiro), para o vendedor de
   melhor desempenho aparecer em destaque por padrão. */
export function computeSellerStats(rows) {
  rows = rows || [];
  const bySeller = {};
  rows.forEach((d) => {
    const seller = d.seller || 'N/D';
    if (!bySeller[seller]) bySeller[seller] = { seller, total: 0, approved: 0, lost: 0, review: 0 };
    const entry = bySeller[seller];
    entry.total += 1;
    if (d.status === 'APPROVED') entry.approved += 1;
    else if (d.status === 'LOST') entry.lost += 1;
    else if (d.status === 'UNDER REVIEW' || d.status === 'NO RESPONSE') entry.review += 1;
  });

  return Object.values(bySeller)
    .map((entry) => ({
      ...entry,
      approvedRate: entry.total ? (entry.approved / entry.total) * 100 : null,
      lostRate: entry.total ? (entry.lost / entry.total) * 100 : null,
    }))
    .sort((a, b) => (b.approvedRate ?? -1) - (a.approvedRate ?? -1));
}

/* ============================================================
   PERÍODOS GENÉRICOS PARA COMPARAÇÃO (Especificação de melhorias, item 5.2)
   ============================================================
   A view "Comparar" hoje só sabe comparar Mês A vs Mês B (hard-coded).
   As funções abaixo generalizam isso: getPeriodKey() recebe uma data ISO
   (AAAA-MM-DD) e uma granularidade, e devolve uma chave de agrupamento —
   "Mês" deixa de ser um caminho de código separado e vira só mais um
   caso de granularidade, ao lado de Ano/Semestre/Trimestre/Bimestre/
   Semana. Adicionar uma granularidade nova no futuro é adicionar um
   `case` aqui, não duplicar a lógica de comparação inteira. */
export const PERIOD_GRANULARITIES = ['year', 'semester', 'quarter', 'bimonth', 'month', 'week'];

/* Número da semana ISO 8601 (semanas começam na segunda-feira; a semana
   1 é a que contém a primeira quinta-feira do ano — definição padrão
   internacional, evita a ambiguidade de "semana do mês" ou "semana
   corrida" que dependeria de fuso/convenção local). */
function isoWeekParts(dateObjUtc) {
  const d = new Date(Date.UTC(dateObjUtc.getUTCFullYear(), dateObjUtc.getUTCMonth(), dateObjUtc.getUTCDate()));
  const dayNum = (d.getUTCDay() + 6) % 7; // segunda=0 ... domingo=6
  d.setUTCDate(d.getUTCDate() - dayNum + 3); // quinta-feira da mesma semana
  const isoYear = d.getUTCFullYear();
  const firstThursday = new Date(Date.UTC(isoYear, 0, 4));
  const firstThursdayDayNum = (firstThursday.getUTCDay() + 6) % 7;
  firstThursday.setUTCDate(firstThursday.getUTCDate() - firstThursdayDayNum + 3);
  const weekNum = 1 + Math.round((d - firstThursday) / (7 * 24 * 3600 * 1000));
  return { isoYear, weekNum };
}

/* Recebe uma data ISO (AAAA-MM-DD, já validada por toISODate) e devolve
   a chave de agrupamento para a granularidade pedida. Retorna null para
   entrada inválida/vazia — nunca lança erro (mesmo princípio de
   robustez de toISODate/computeKpis: dado ausente não quebra o cálculo,
   só fica fora do agrupamento). */
export function getPeriodKey(dateISO, granularity) {
  const iso = toISODate(dateISO);
  if (!iso) return null;
  const [y, m, d] = iso.split('-').map(Number);
  switch (granularity) {
    case 'year': return String(y);
    case 'semester': return `${y}-S${m <= 6 ? 1 : 2}`;
    case 'quarter': return `${y}-T${Math.ceil(m / 3)}`;
    case 'bimonth': return `${y}-B${Math.ceil(m / 2)}`;
    case 'month': return `${y}-${String(m).padStart(2, '0')}`;
    case 'week': {
      const { isoYear, weekNum } = isoWeekParts(new Date(Date.UTC(y, m - 1, d)));
      return `${isoYear}-W${String(weekNum).padStart(2, '0')}`;
    }
    default: return null;
  }
}

const PT_MONTHS_SHORT = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez'];

/* Rótulo legível em pt-BR para uma chave de período, dado a granularidade
   que a gerou (a mesma passada a getPeriodKey — não é redetectada a
   partir do formato da chave, para não ambiguar "2026-06" mês vs uma
   futura chave que comece igual). */
export function periodLabel(key, granularity) {
  if (!key) return 'Sem data';
  switch (granularity) {
    case 'year': return key;
    case 'semester': {
      const [y, s] = key.split('-S');
      return `${s}º sem./${y.slice(2)}`;
    }
    case 'quarter': {
      const [y, t] = key.split('-T');
      return `${t}º tri./${y.slice(2)}`;
    }
    case 'bimonth': {
      const [y, b] = key.split('-B');
      return `${b}º bim./${y.slice(2)}`;
    }
    case 'month': {
      const [y, m] = key.split('-');
      return `${PT_MONTHS_SHORT[Number(m) - 1]}/${y.slice(2)}`;
    }
    case 'week': {
      const [y, w] = key.split('-W');
      return `Sem.${w}/${y.slice(2)}`;
    }
    default: return key;
  }
}

/* ---------- ranking de clientes com valores monetários (Especificação de melhorias, item 6) ----------
   Antes calculado implicitamente só como contagem (clientCounts), usado
   apenas para alimentar chartClients — sem valor monetário nenhum e sem
   estrutura própria testável. Devolve um array ordenado por volume
   (mesmo critério do ranking do gráfico, para que gráfico e tabela nunca
   contem histórias diferentes), com contagem, ALL IN somado e ticket
   médio por cliente — usados na nova tabela ao lado do gráfico de pizza
   reduzido. */
export function computeClientRanking(rows) {
  rows = rows || [];
  const byClient = {};
  rows.forEach((d) => {
    if (!d.client) return;
    if (!byClient[d.client]) byClient[d.client] = { client: d.client, count: 0, totalAllIn: 0 };
    byClient[d.client].count += 1;
    byClient[d.client].totalAllIn += Number(d.allIn) || 0;
  });

  return Object.values(byClient)
    .map((entry) => ({
      ...entry,
      avgTicket: entry.count ? entry.totalAllIn / entry.count : 0,
    }))
    .sort((a, b) => b.count - a.count);
}

/* ==========================================================================
   PARTE 2 de 2 — MOTOR DE CÁLCULO (antes: calculos.js)
   ========================================================================== */
/* ============================================================
   calculos.js — BDG · Controle de Cotações (bdg-cotacoes-staging)
   Fase Extra · Etapa 2 — Motor de cálculo (engenharia reversa da planilha)

   Funções PURAS (sem DOM, sem Firebase, sem efeitos colaterais) que
   reproduzem a lógica de cálculo identificada na Análise Técnica de
   engenharia reversa da planilha (7 cotações, 9 planilhas, 100% dos
   blocos conferidos batendo com o PDF entregue ao cliente).

   Como recomendado diretamente pela análise técnica (Parte 7.4 /
   Conclusões), existem DUAS famílias de cálculo, tratadas aqui como
   duas funções separadas, sem tentar generalizar num único motor:

     - calcularCotacaoAerea()     → motor AP → GAP → AR por Kg
                                     (cotações Aéreo / S&A — mesmo layout
                                     de bloco, ver Análise Técnica 7.2)
     - calcularCotacaoMaritima()  → motor por contêiner (FCL) ou por
                                     W/M — peso ou cubagem, o maior (LCL)

   Nenhuma tela consome este arquivo ainda — essa integração é tarefa
   da Etapa 3. Aqui a única responsabilidade é o cálculo, validado
   célula a célula contra os casos reais já conferidos na Análise
   Técnica (ver calculos.test.js).
   ============================================================ */

/* ---------- motor aéreo / S&A (Análise Técnica, Partes 2, 3, 4 e 5) ----------

   Parâmetros de entrada — NUNCA como constantes fixas no código (ver
   Análise Técnica, Parte 4, "Ponto de atenção para a Etapa 2" e
   Recomendação 2 das Conclusões): os valores fixos de custo em Miami
   e a margem comercial mudam por rota/parceiro/negociação, então
   entram sempre como parâmetro, nunca hard-coded.

   @param {number} input.pesoTarifavel     Peso tarifável em Kg (chargeable weight)
   @param {number[]} input.tarifasPorTrecho Tarifas de custo por trecho, USD/Kg (AP).
                                             Ex.: [6.15, 1.20] para HKG×MIA + MIA×MAO.
                                             Um único valor é aceito para serviços de
                                             um trecho só (ex.: AIR STD, Análise Técnica 7.1).
   @param {number} input.gapPercentual     Margem comercial (ex.: 0.025 para 2,5%)
   @param {number} [input.custosFixosOrigem=0]
                                             Soma dos custos fixos por operação/parceiro
                                             em Miami (N96:N102 na planilha — pickup local,
                                             handling, consolidação etc.), configurável por
                                             rota/parceiro, nunca fixo no código.
   @param {Array<{tarifa:number, peso:number, minimo?:number}>} [input.taxasMiami=[]]
                                             Tabela de taxas aeroportuárias/fixas por Kg
                                             (Análise Técnica, Parte 4 — Q95:U104). Cada
                                             taxa só entra no cálculo se estiver ativa
                                             naquela cotação (peso/tarifa informados).
   @param {number} [input.handlingFee=0]    Taxa fixa por HAWB (J101 na planilha)
   @param {boolean} [input.aplicarMinimos=false]
                                             DECISÃO REGISTRADA (roteiro de execução,
                                             Etapa 2, passo 7): nas planilhas analisadas os
                                             "mínimos" por taxa (coluna MIN) aparecem hoje
                                             apenas como referência de conferência manual do
                                             vendedor — a fórmula em produção (T104 =
                                             SUM(T97:T103)) NÃO aplica Math.max(calculado,
                                             mínimo) automaticamente. Por isso o padrão desta
                                             função é `false` (replica o comportamento real
                                             hoje em produção). Setar `true` só depois de
                                             confirmar com o time de operações que a trava
                                             automática deve passar a valer (ver Análise
                                             Técnica, Parte 4, e Recomendação 4 das Conclusões).

   @returns {object} apTotal, gapValor, ar, af, miaCharges, allIn, totalCollect,
                      pesoCargaCharge, totalOutrosEncargos, handlingInformation
                      — nomenclatura e fórmulas espelham exatamente as células
                      C98/C100/C101/J99/N106/J102/J107/J105/J106/J109 da planilha
                      (Análise Técnica, Partes 3 e 5). */
export function calcularCotacaoAerea(input) {
  const pesoTarifavel = input.pesoTarifavel || 0;
  const tarifasPorTrecho = Array.isArray(input.tarifasPorTrecho)
    ? input.tarifasPorTrecho
    : [input.tarifasPorTrecho || 0];
  const gapPercentual = input.gapPercentual || 0;
  const custosFixosOrigem = input.custosFixosOrigem || 0;
  const taxasMiami = input.taxasMiami || [];
  const handlingFee = input.handlingFee || 0;
  const aplicarMinimos = input.aplicarMinimos || false;

  // AP → GAP → AR (Análise Técnica, Parte 3)
  const apTotal = tarifasPorTrecho.reduce((acc, t) => acc + (t || 0), 0); // C98
  const gapValor = apTotal * gapPercentual; // C100
  const ar = apTotal + gapValor; // C101 — é o "Rate" impresso no PDF

  // Cost (A/F) — J99 = AR (venda) × peso tarifável
  const af = ar * pesoTarifavel;

  // Tabela de taxas fixas de Miami (Análise Técnica, Parte 4)
  const totalTaxasMiami = taxasMiami.reduce((acc, taxa) => {
    const calculado = (taxa.tarifa || 0) * (taxa.peso || 0); // T = S×R
    const valor = aplicarMinimos && taxa.minimo
      ? Math.max(calculado, taxa.minimo)
      : calculado;
    return acc + valor;
  }, 0); // T104

  // FCA/EXW/Origin + MIA charges — N106 = custos fixos + taxas de Miami
  const miaCharges = custosFixosOrigem + totalTaxasMiami; // J97 / N106

  // Consolidação da cotação (Análise Técnica, Parte 5.1)
  const allIn = miaCharges + af + handlingFee; // J102 = SUM(J97:J101)
  const pesoCargaCharge = af; // J105 — Weight charge (per HAWB)
  const totalOutrosEncargos = miaCharges; // J106 — Total other charges due to Agent
  const totalCollect = totalOutrosEncargos + pesoCargaCharge; // J107
  const handlingInformation = handlingFee; // J109

  return {
    apTotal,
    gapValor,
    ar,
    af,
    miaCharges,
    allIn,
    totalCollect,
    pesoCargaCharge,
    totalOutrosEncargos,
    handlingInformation,
  };
}

/* ---------- motor marítimo — FCL e LCL (Análise Técnica, Parte 7.3) ----------

   Diferente do motor aéreo, não existe tarifa por Kg: os custos são
   somados em três blocos fixos (Origin / Freight / Destination), por
   contêiner (FCL) ou por W/M (LCL) — mas o mesmo princípio custo →
   margem → venda se aplica em ambos (Análise Técnica, Parte 7.3, e
   Recomendação 1 das Conclusões: função separada da aérea, sem tentar
   unificar as duas).

   @param {'FCL'|'LCL'} input.tipo
   @param {number} [input.custosOrigem=0]    Origin Charges (ex.: FCA Charges)
   @param {number} [input.custosDestino=0]   Destination Charges (THC, DOC Fee,
                                              Drop Off, Handling...)
   @param {number} [input.custosFrete=0]     Freight Charges (Ocean Freight) —
                                              usado apenas quando tipo === 'FCL'
                                              (valor já fechado por contêiner).
   @param {number} [input.pesoToneladas=0]   Peso da carga em toneladas — usado
                                              apenas quando tipo === 'LCL'.
   @param {number} [input.cubagemM3=0]       Cubagem da carga em m³ — usado
                                              apenas quando tipo === 'LCL'.
   @param {number} [input.tarifaPorWM=0]     Tarifa de frete por W/M (USD/ton
                                              ou USD/m³) — usado apenas em LCL.
   @param {Array<{tarifa:number, minimo?:number}>} [input.taxasPortuarias=[]]
                                              Taxas portuárias rateadas por W/M
                                              (Unstuffing, Siscarga, BL Fee, Cargo
                                              Split, Capatazia, ISPS, LWS,
                                              Fumigation, TRS...) — usado em LCL.
   @param {number} [input.gapPercentual=0]   Margem comercial (mesmo princípio
                                              AP/GAP/AR do motor aéreo, aplicada
                                              aqui sobre o custo total do contêiner
                                              ou do W/M).
   @param {boolean} [input.aplicarMinimos=false]
                                              Mesma decisão registrada do motor
                                              aéreo (ver calcularCotacaoAerea):
                                              os mínimos por taxa portuária não
                                              são aplicados automaticamente por
                                              padrão, replicando o comportamento
                                              hoje observado nas planilhas.

   @returns {object} tipo, wm (null em FCL), apTotal (custo total), gapValor,
                      ar (venda total), custosOrigem, custosFrete, custosDestino */
export function calcularCotacaoMaritima(input) {
  const tipo = input.tipo === 'LCL' ? 'LCL' : 'FCL';
  const custosOrigem = input.custosOrigem || 0;
  const custosDestino = input.custosDestino || 0;
  const gapPercentual = input.gapPercentual || 0;
  const aplicarMinimos = input.aplicarMinimos || false;

  let custosFrete;
  let wm = null;

  if (tipo === 'LCL') {
    const pesoToneladas = input.pesoToneladas || 0;
    const cubagemM3 = input.cubagemM3 || 0;
    // W/M — "whichever is greater" entre peso (ton) e cubagem (m³)
    // (Análise Técnica, Parte 7.3 e Glossário)
    wm = Math.max(pesoToneladas, cubagemM3);

    const tarifaPorWM = input.tarifaPorWM || 0;
    const freteBase = tarifaPorWM * wm;

    const taxasPortuarias = input.taxasPortuarias || [];
    const totalTaxasPortuarias = taxasPortuarias.reduce((acc, taxa) => {
      const calculado = (taxa.tarifa || 0) * wm;
      const valor = aplicarMinimos && taxa.minimo
        ? Math.max(calculado, taxa.minimo)
        : calculado;
      return acc + valor;
    }, 0);

    custosFrete = freteBase + totalTaxasPortuarias;
  } else {
    // FCL — valor já fechado por contêiner, somado diretamente
    // (sem tarifa por Kg ou W/M — Análise Técnica, Parte 7.3)
    custosFrete = input.custosFrete || 0;
  }

  // Mesmo princípio custo → margem → venda do motor aéreo, aplicado
  // sobre o custo total do contêiner/W-M (Análise Técnica, Parte 7.3:
  // "cada aba termina com... AP (custo total), GAP (margem em valor) e % (margem percentual)")
  const apTotal = custosOrigem + custosFrete + custosDestino;
  const gapValor = apTotal * gapPercentual;
  const ar = apTotal + gapValor;

  return {
    tipo,
    wm,
    apTotal,
    gapValor,
    ar,
    custosOrigem,
    custosFrete,
    custosDestino,
  };
}
