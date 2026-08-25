/* eslint-disable security/detect-object-injection --
   Decisão registrada (Fase Extra · Etapa 3): as chaves usadas em acesso
   dinâmico a objeto neste arquivo (ex.: taxasPadrao[idx][campo],
   configListas[key], CAMPOS_CALCULAVEIS[id]) vêm sempre de constantes
   internas do próprio código — ids de campo fixos, nomes de configuração
   definidos aqui mesmo — nunca de entrada bruta e não validada do usuário.
   É o mesmo tipo de falso positivo já registrado para utils.js na Fase 3;
   deixar a regra ativa aqui só gera ruído sem reduzir risco real. */
/* ==========================================================================
   cotacoes.js — BDG · Controle de Cotações (Samsung SDS Logistics)
   ==========================================================================
   Fase 3 — Qualidade de Código e Testes | Etapa 2 de modularização.

   Módulo de Cotações, Gráficos e Conexão. Extraído do segundo bloco
   <script type="module"> que antes vivia embutido no index.html, logo após
   o auth.js (etapa 1).

   Responsabilidades deste módulo:
   - Sincronização em tempo real das cotações com o Firebase Realtime Database
   - Monitoramento reativo da conexão (.info/connected) e reconexão automática
     por backoff exponencial, com botão manual preservado (Fase 2)
   - Filtros, ordenação, busca e KPIs do painel
   - Gráficos (Chart.js): volume mensal, status, vendedores, clientes, modais
   - Adicionar/editar/excluir cotação, upload e recorte de foto
   - Importação em massa via planilha Excel e exportação Excel/PDF
   - Comparativo mensal
   - Datas padronizadas em ISO 8601 (toISODate — Fase 2)

   Por que virou um arquivo próprio (em vez de <script> inline no HTML):
   - Reduz o index.html em ~1070 linhas, na mesma linha do que foi feito com
     auth.js na etapa 1 — o HTML volta a concentrar só marcação/estilo.
   - Passa a ser enxergado pelo ESLint como um arquivo JS normal, entrando
     automaticamente na cobertura do "npm run lint" já configurado na Fase 3.

   Não modularizado ainda (proposital, para reduzir risco nesta etapa):
   - Este módulo continua com a lógica de dados (Firebase) e de interface
     (DOM, gráficos) juntas, como já estava no index.html. Separar em
     cotacoes.js (dados) + ui.js (renderização) é um passo futuro possível,
     mas exige desacoplar estado compartilhado (RAW, state, instâncias dos
     gráficos) entre os dois arquivos — maior risco de regressão, avaliado
     como não prioritário frente ao ganho já obtido com esta extração.

   Depende de globais carregados via <script src> no <head> do index.html
   (não são imports ES porque essas bibliotecas não publicam módulos ES
   utilizáveis diretamente por CDN neste projeto): Chart (Chart.js), XLSX
   (SheetJS) e window.jspdf (jsPDF + autotable).

   Depende também de funções expostas em "window" por outros arquivos:
   - window.escapeHtml, window.openPhotoCropper (definidos no <script> comum
     do index.html, fora dos módulos — ver comentário lá para o porquê)
   - window.logActivity, window.getCurrentUser (expostos por auth.js)
   - window.markLoadingStep, window.showAppLoading, window.hideAppLoadingImmediately
     (definidos no <script> comum do index.html, tela de splash/carregamento)
   ========================================================================== */

import { getApp } from "https://www.gstatic.com/firebasejs/12.16.0/firebase-app.js";
import { getDatabase, ref as dbRef, set as dbSet, push as dbPush, remove as dbRemove, onValue, goOffline as dbGoOffline, goOnline as dbGoOnline } from "https://www.gstatic.com/firebasejs/12.16.0/firebase-database.js";
import { getAuth } from "https://www.gstatic.com/firebasejs/12.16.0/firebase-auth.js";
/* Fase 3 (modularização): funções puras de data, importação em massa,
   filtros e KPIs vêm de utils.js, com cobertura de testes automatizados
   (Jest) — etapa 2 acrescentou filterQuotations e computeKpis. */
/* Fase Extra · Etapa 3 — motor de cálculo (validado na Etapa 2).
   calcularCotacaoAerea() cobre o modal aéreo/S&A (mesmo motor, Análise Técnica
   7.2); o motor marítimo (calcularCotacaoMaritima) fica para uma iteração
   seguinte da tela de criação, já que os modais atuais do cadastro (AIR,
   OCEAN, OCEAN/LCL, SEA/AIR) usam predominantemente o motor por Kg.

   As funções puras e o motor de cálculo vinham de utils.js e calculos.js,
   unificados em core.js na Auditoria Geral de 01/08 (Bloco 4). */
import {
  TYPE_VALUES, toISODate, normalizeTypeImport, normalizeStatusImport, normalizeDateImport,
  filterQuotations, computeKpis, parseDecimalBR, formatDecimalBR, sparklinePoints, computeSellerStats,
  getPeriodKey, periodLabel, PERIOD_GRANULARITIES, computeClientRanking,
  calcularCotacaoAerea, calcularCotacaoMaritima
} from './core.js?v=20260802a';
import { downloadPdfCotacao } from './pdf.js?v=20260802a';
// Especificação de melhorias, item 3 — gate de manutenção POR VIEW (view
// "Adicionar"), reaproveitando o mesmo mecanismo de autorização (role
// admin) do gate global já existente em maintenance.js.
import { checkViewMaintenanceGate } from './maintenance.js?v=20260802a';

/* Base inicial de cotações — usada UMA única vez, para semear o Realtime
   Database quando o nó ainda está vazio (ver onValue mais abaixo). Depois
   disso o banco é a fonte de verdade e este array é sobrescrito na
   primeira leitura.
   Os marcadores DATA_MARKER que existiam aqui serviam ao antigo botão
   "Baixar HTML atualizado", removido na Auditoria Geral de 01/08
   (Bloco 2) — nada mais os referencia. */
let RAW = [{"ref": "MAO_IM_01022", "client": "MAGTEC", "seller": "ROGER", "type": "Import", "modal": "OCEAN", "incoterm": "", "origin": "SHEKOU", "dest": "MAO", "week": null, "month": null, "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_02108", "client": "TUTIPLAST", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "HONG KONG", "dest": "MAO", "week": null, "month": null, "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IM_02113", "client": "BEMOL", "seller": "ROGER", "type": "Import", "modal": "OCEAN", "incoterm": "FOB", "origin": "PTY", "dest": "MAO", "week": "2026-03-03 00:00:00", "month": "2026-03", "status": "LOST", "remarks": "VALOR SUPERIOR AOS CONCORRENTES"}, {"ref": "MAO_EA_03114", "client": "TUTIPLAST", "seller": "ROGER", "type": "Export", "modal": "AIR", "incoterm": "DDP", "origin": "MANAUS", "dest": "SHENZHEN", "week": "2026-03-04 00:00:00", "month": "2026-03", "status": "NO RESPONSE", "remarks": "CLIENTE NÃO RESPONDEU O E-MAIL"}, {"ref": "MAO_IL_03119", "client": "ADATA", "seller": "ROGER", "type": "Import", "modal": "OCEAN", "incoterm": "EXW", "origin": "SHENZHEN", "dest": "MAO", "week": "2026-03-04 00:00:00", "month": "2026-03", "status": "NO RESPONSE", "remarks": "CLIENTE NÃO RESPONDEU O E-MAIL"}, {"ref": "MAO_IL_03120", "client": "ADATA", "seller": "ROGER", "type": "Import", "modal": "OCEAN", "incoterm": "EXW", "origin": "SHENZHEN", "dest": "MAO", "week": "2026-03-04 00:00:00", "month": "2026-03", "status": "LOST", "remarks": ""}, {"ref": "MAO_IM_03122", "client": "MAGTEC", "seller": "ROGER", "type": "Import", "modal": "OCEAN", "incoterm": "FOB", "origin": "QINGDAO", "dest": "", "week": "2026-03-03 00:00:00", "month": "2026-03", "status": "LOST", "remarks": "VALIDADE DA PROPOSTA NÃO ATENDE AO CLIENTE"}, {"ref": "MAO_IA_03124", "client": "SALCOMP", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "NANJING/JIANGSU", "dest": "", "week": "2026-03-06 00:00:00", "month": "2026-03", "status": "APPROVED", "remarks": "N/A"}, {"ref": "MAIO_IA_3124", "client": "SALCOMP", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "SHANGHAI", "dest": "MAO", "week": "2026-04-02 00:00:00", "month": "2026-04", "status": "APPROVED", "remarks": ""}, {"ref": "MAIO_IA_3125", "client": "ELECTROLUX MAGTEC", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "SUSHOU", "dest": "MAO", "week": "2026-03-25 00:00:00", "month": "2026-03", "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IL_03129", "client": "TUTIPLAST", "seller": "ROGER", "type": "Import", "modal": "OCEAN", "incoterm": "EXW", "origin": "SHENZHEN", "dest": "MAO", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "VCP_IA_03136", "client": "ADATA", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "", "dest": "", "week": null, "month": null, "status": "LOST", "remarks": "VALIDADE DA PROPOSTA NÃO ATENDE AO CLIENTE"}, {"ref": "MAO_IA_03141", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "", "dest": "", "week": null, "month": null, "status": "APPROVED", "remarks": "TARIFÁRIO / BID"}, {"ref": "MAO_IA_03142", "client": "SALCOMP", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "HONG KONG", "dest": "MAO", "week": "2026-03-11 00:00:00", "month": "2026-03", "status": "APPROVED", "remarks": "TARIFÁRIO / BID"}, {"ref": "MAO_IA_03149", "client": "BEMOL", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "MIA", "dest": "MAO", "week": "2026-03-13 00:00:00", "month": "2026-03", "status": "APPROVED", "remarks": "N/A"}, {"ref": "MAO_IA_03155", "client": "TUTIPLAST", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "SHENZHEN", "dest": "MAO", "week": "2026-03-18 00:00:00", "month": "2026-03", "status": "NO RESPONSE", "remarks": "CLIENTE NÃO RESPONDEU O E-MAIL"}, {"ref": "MAO_IA_03160", "client": "TUTIPLAST", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "SHENZHEN", "dest": "MAO", "week": "2026-03-19 00:00:00", "month": "2026-03", "status": "NO RESPONSE", "remarks": "CLIENTE NÃO RESPONDEU O E-MAIL"}, {"ref": "MAO_IM_03163", "client": "MULTIFORTE", "seller": "ROGER", "type": "Import", "modal": "OCEAN", "incoterm": "FOB", "origin": "DALIAN", "dest": "MAO", "week": "2026-03-20 00:00:00", "month": "2026-03", "status": "NO RESPONSE", "remarks": "CLIENTE NÃO RESPONDEU O E-MAIL"}, {"ref": "MAO_IA_03164", "client": "SX", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "PVG", "dest": "MAO", "week": null, "month": null, "status": "APPROVED", "remarks": "TARIFÁRIO / BID"}, {"ref": "MAO_IA_03165", "client": "SALCOMP", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "HONG KONG", "dest": "MAO", "week": null, "month": null, "status": "APPROVED", "remarks": "TARIFÁRIO / BID"}, {"ref": "MAO_IA_03167 ", "client": "ADATA", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "SUSHOU", "dest": "MAO", "week": null, "month": null, "status": "APPROVED", "remarks": "TARIFÁRIO / BID"}, {"ref": "MAO_IM_03168", "client": "BEMOL", "seller": "ROGER", "type": "Import", "modal": "OCEAN", "incoterm": "FOB", "origin": "CHINA", "dest": "MAO", "week": "2026-03-20 00:00:00", "month": "2026-03", "status": "NO RESPONSE", "remarks": "CLIENTE NÃO RESPONDEU O E-MAIL"}, {"ref": "MAO_IA_03171", "client": "TUTIPLAST", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "SHENZHEN", "dest": "MAO", "week": "2026-03-24 00:00:00", "month": "2026-03", "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IA_03172", "client": "TUTIPLAST", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "XIAMEN", "dest": "MAO", "week": "2026-03-24 00:00:00", "month": "2026-03", "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IA_03173", "client": "BEMOL", "seller": "ROGER", "type": "Import", "modal": "OCEAN", "incoterm": "FOB", "origin": "CHINA", "dest": "MAO", "week": null, "month": null, "status": "NO RESPONSE", "remarks": "FALTA ENVIAR PARA O CLIENTE"}, {"ref": "MAO_IA_03175", "client": "TUTIPLAST", "seller": "ROGER", "type": "Import", "modal": "OCEAN", "incoterm": "EXW", "origin": "SHENZHEN", "dest": "MAO", "week": "2026-03-24 00:00:00", "month": "2026-03", "status": "LOST", "remarks": "VALOR SUPERIOR AOS CONCORRENTES"}, {"ref": "MAO_IA_03180", "client": "TUTIPLAST", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "XIAMEN", "dest": "MAO", "week": "2026-03-24 00:00:00", "month": "2026-03", "status": "LOST", "remarks": "VALOR SUPERIOR AOS CONCORRENTES"}, {"ref": "MAO_IA_03181", "client": "TUTIPLAST", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "SHENZHEN", "dest": "MAO", "week": "2026-03-31 00:00:00", "month": "2026-03", "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_03182", "client": "TUTIPLAST", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "SHENZHEN", "dest": "MAO", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAIO_IA_3183", "client": "SALCOMP", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "HONG KONG", "dest": "MAO", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IM_03185", "client": "TUTIPLAST", "seller": "ROGER", "type": "Import", "modal": "OCEAN", "incoterm": "FCA", "origin": "ANTWERP", "dest": "MAO", "week": "2026-04-01 00:00:00", "month": "2026-04", "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IA_03187", "client": "SALCOMP", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "HONG KONG", "dest": "MAO", "week": "2026-03-30 00:00:00", "month": "2026-03", "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IA_03188", "client": "SALCOMP", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "HONG KONG", "dest": "MAO", "week": "2026-03-30 00:00:00", "month": "2026-03", "status": "LOST", "remarks": "VALOR SUPERIOR AOS CONCORRENTES"}, {"ref": "MAO_IA_03189", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "FUZHOU", "dest": "MAO", "week": "2026-04-02 00:00:00", "month": "2026-04", "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_03190", "client": "ADATA", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "TBC", "origin": "KOREA", "dest": "MAO", "week": "2026-03-30 00:00:00", "month": "2026-03", "status": "APPROVED", "remarks": "TARIFÁRIO / BID"}, {"ref": "MAO_IA_03191", "client": "ADATA", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "TBC", "origin": "VIATNÃ", "dest": "MAO", "week": "2026-03-30 00:00:00", "month": "2026-03", "status": "APPROVED", "remarks": "TARIFÁRIO / BID"}, {"ref": "MAO_IA_03193", "client": "TUTIPLAST", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "HONG KONG", "dest": "MAO", "week": "2026-03-31 00:00:00", "month": "2026-03", "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IA_03194", "client": "BEMOL", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "SHANGHAI", "dest": "MAO", "week": "2026-03-31 00:00:00", "month": "2026-03", "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_03196", "client": "SX LED", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "NINGBO", "dest": "MAO", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IM_03197", "client": "TUTIPLAST", "seller": "ROGER", "type": "Import", "modal": "OCEAN", "incoterm": "FOB", "origin": "BUSAN", "dest": "MAO", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "VCP_IA_04198", "client": "ADATA", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "HONG KONG", "dest": "VCP", "week": "2026-03-31 00:00:00", "month": "2026-03", "status": "APPROVED", "remarks": "TARIFÁRIO / BID"}, {"ref": "VCP_IA_04199", "client": "ADATA", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "HONG KONG", "dest": "VCP", "week": "2026-03-31 00:00:00", "month": "2026-03", "status": "APPROVED", "remarks": "TARIFÁRIO / BID"}, {"ref": "VCP_IA_04200", "client": "ADATA", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "HONG KONG", "dest": "VCP", "week": "2026-03-31 00:00:00", "month": "2026-03", "status": "APPROVED", "remarks": "TARIFÁRIO / BID"}, {"ref": "MAO_IA_04200", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "HONG KONG", "dest": "MAO", "week": null, "month": null, "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_04201", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "SHANGHAI", "dest": "MAO", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IA_04202", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "HONG KONG", "dest": "MAO", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IA_04203", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "TIANJIN", "dest": "MAO", "week": null, "month": null, "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_04204", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "TIANJIN", "dest": "MAO", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IA_04205", "client": "SALCOMP", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "HONG KONG", "dest": "MAO", "week": "2026-04-01 00:00:00", "month": "2026-04", "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_04205", "client": "SALCOMP", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "HONG KONG", "dest": "MAO", "week": null, "month": null, "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_04206", "client": "SALCOMP", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "HONG KONG", "dest": "MAO", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IM_04207", "client": "BEMOL", "seller": "ROGER", "type": "Import", "modal": "OCEAN", "incoterm": "EXW", "origin": "PANAMA", "dest": "MAO", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IM_04207", "client": "BEMOL", "seller": "ROGER", "type": "Import", "modal": "OCEAN", "incoterm": "FOB", "origin": "", "dest": "", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IA_04208", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "HONG KONG", "dest": "MAO", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IA_04209", "client": "ELOX TRADE", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "HONG KONG", "dest": "MAO", "week": null, "month": null, "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_04210", "client": "BEMOL", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "MIA", "dest": "MAO", "week": "2026-04-02 00:00:00", "month": "2026-04", "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_04210", "client": "BEMOL", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "MIAMI", "dest": "MAO", "week": null, "month": null, "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_04211", "client": "BEMOL", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "MIA", "dest": "MAO", "week": null, "month": null, "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_04211", "client": "BEMOL", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "MIAMI", "dest": "MAO", "week": null, "month": null, "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_04212", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "FUZHOU", "dest": "MAO", "week": "2026-04-02 00:00:00", "month": "2026-04", "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_04212", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "FUZHOU", "dest": "MAO", "week": null, "month": null, "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_04213", "client": "ADATA", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "", "dest": "MAO", "week": null, "month": null, "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_04214", "client": "ADATA", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "SUSHOU", "dest": "MAO", "week": "2026-04-02 00:00:00", "month": "2026-04", "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_04214", "client": "ADATA", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "SHANGHAI", "dest": "MAO", "week": null, "month": null, "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_04215", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "HONG KONG", "dest": "MAO", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IA_04216", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "TIANJIN", "dest": "MAO", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IA_04217", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "HAI PHONG", "dest": "MAO", "week": null, "month": null, "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_04218", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "INCHEON", "dest": "MAO", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IA_04219", "client": "BEMOL", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "HONG KONG", "dest": "MAO", "week": null, "month": null, "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_04220", "client": "BEMOL", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "NINGBO", "dest": "MAO", "week": "2026-04-08 00:00:00", "month": "2026-04", "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_04220", "client": "BEMOL", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "NINGBO", "dest": "MAO", "week": null, "month": null, "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IM_04221", "client": "ELOX TRADE", "seller": "ROGER", "type": "Import", "modal": "OCEAN", "incoterm": "FOB", "origin": "", "dest": "", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IA_04222", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "JIANGSU", "dest": "MAO", "week": null, "month": null, "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IM_04223", "client": "N/D", "seller": "ROGER", "type": "Import", "modal": "OCEAN", "incoterm": "", "origin": "", "dest": "", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IM_04223", "client": "ELOX TRADE", "seller": "ROGER", "type": "Import", "modal": "OCEAN", "incoterm": "FOB", "origin": "", "dest": "", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IA_04224", "client": "GERTEC", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "HONG KONG", "dest": "MAO", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IA_04225", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "HONG KONG", "dest": "MAO", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IA_04226", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "FUZHOU", "dest": "MAO", "week": null, "month": null, "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_04227", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "HONG KONG", "dest": "MAO", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IA_04228", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "", "origin": "", "dest": "", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IA_04229", "client": "SALCOMP", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "HONG KONG", "dest": "MAO", "week": null, "month": null, "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_04230", "client": "SALCOMP", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "SHENZHEN", "dest": "MAO", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IA_04231", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "SHANGHAI", "dest": "MAO", "week": null, "month": null, "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_04232", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "FUZHOU", "dest": "MAO", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IA_04233", "client": "SALCOMP", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "", "origin": "", "dest": "", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IA_04234", "client": "GERTEC", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "", "origin": "", "dest": "", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IA_04235", "client": "GERTEC", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "", "origin": "", "dest": "", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IA_04236", "client": "GERTEC", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "", "origin": "", "dest": "", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IA_04237", "client": "TUTIPLAST", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FOB", "origin": "XIAMEN", "dest": "MAO", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IA_04238", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "HONG KONG", "dest": "MAO", "week": null, "month": null, "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_04239", "client": "BEMOL", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "NINGBO", "dest": "MAO", "week": null, "month": null, "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IM_04240", "client": "TUTIPLAST", "seller": "ROGER", "type": "Import", "modal": "OCEAN", "incoterm": "FOB", "origin": "DALIAN", "dest": "MAO", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IM_04241", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "OCEAN", "incoterm": "FOB", "origin": "YANTIAN", "dest": "MAO", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IA_04242", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "FUZHOU", "dest": "MAO", "week": null, "month": null, "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_04243", "client": "BEMOL", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "NINGBO", "dest": "MAO", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IA_04244", "client": "BEMOL", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "", "origin": "", "dest": "", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IA_04247", "client": "SALCOMP", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "HONG KONG", "dest": "MAO", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IA_04248", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "GUANGDONG", "dest": "MAO", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IA_04252", "client": "ELOX TRADE", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "MIAMI", "dest": "MAO", "week": null, "month": null, "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IM_04253", "client": "ELOX TRADE", "seller": "ROGER", "type": "Import", "modal": "OCEAN", "incoterm": "FOB", "origin": "", "dest": "", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IM_04254", "client": "ELOX TRADE", "seller": "ROGER", "type": "Import", "modal": "OCEAN", "incoterm": "FOB", "origin": "QINGDAO", "dest": "MAO", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IA_04255", "client": "ELOX TRADE", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "GUANGDONG", "dest": "MAO", "week": null, "month": null, "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_04256", "client": "TUTIPLAST", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "XIAMEN", "dest": "MAO", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IA_04257", "client": "TUTIPLAST", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "XIAMEN", "dest": "MAO", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IA_04258", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "SHANGHAI", "dest": "MAO", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IA_04259", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "SHENZHEN", "dest": "MAO", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IA_04260", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "SHENZHEN", "dest": "MAO", "week": null, "month": null, "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_04261", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "FUZHOU", "dest": "MAO", "week": null, "month": null, "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_04262", "client": "BEMOL", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "HONG KONG", "dest": "MAO", "week": null, "month": null, "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_04263", "client": "HARMAN", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FOB", "origin": "YANTIAN", "dest": "MAO", "week": null, "month": null, "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_04264", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "MIA", "dest": "MAO", "week": null, "month": null, "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_04265", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "HONG KONG", "dest": "MAO", "week": null, "month": null, "status": "LOST", "remarks": ""}, {"ref": "MAO_IA_04266", "client": "ELOX TRADE", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "GUANGDONG", "dest": "MAO", "week": null, "month": null, "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_04267", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "PHU THO,PROVINCE", "dest": "MAO", "week": null, "month": null, "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_04268", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "HONG KONG", "dest": "MAO", "week": null, "month": null, "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_04269", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "SHANGHAI", "dest": "MAO", "week": null, "month": null, "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_04270", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "INCHEON", "dest": "MAO", "week": null, "month": null, "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_04271", "client": "TENNECO", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "FARIDABAD", "dest": "VCP", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IM_04272", "client": "ELOX TRADE", "seller": "ROGER", "type": "Import", "modal": "OCEAN", "incoterm": "FOB", "origin": "DALIAN", "dest": "MAO", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IA_04273", "client": "SALCOMP", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "HONG KONG", "dest": "MAO", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IA_04274", "client": "SALCOMP", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "HONG KONG", "dest": "MAO", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IA_04275", "client": "SALCOMP", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "HONG KONG", "dest": "MAO", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IA_04276", "client": "HARMAN", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "SHANGHAI", "dest": "MAO", "week": null, "month": null, "status": "LOST", "remarks": ""}, {"ref": "MAO_IA_04277", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "HONG KONG", "dest": "MAO", "week": null, "month": null, "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_04278", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "ZHONGSHAN", "dest": "MAO", "week": null, "month": null, "status": "LOST", "remarks": ""}, {"ref": "MAO_IA_04279", "client": "SALCOMP", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "HONG KONG", "dest": "MAO", "week": null, "month": null, "status": "LOST", "remarks": ""}, {"ref": "MAO_IA_04280", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "SHANGHAI", "dest": "MAO", "week": null, "month": null, "status": "LOST", "remarks": ""}, {"ref": "MAO_IA_04281", "client": "SX LIGHTING", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "GUANGDONG", "dest": "MAO", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IA_04282", "client": "HARMAN", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "GUANGDONG", "dest": "MAO", "week": null, "month": null, "status": "LOST", "remarks": ""}, {"ref": "MAO_IA_04283", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "INCHEON", "dest": "MAO", "week": null, "month": null, "status": "LOST", "remarks": ""}, {"ref": "MAO_IA_04284", "client": "ELOX TRADE", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "GUANGDONG", "dest": "MAO", "week": null, "month": null, "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_04285", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "FUZHOU", "dest": "MAO", "week": null, "month": null, "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_04300", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "FUZHOU", "dest": "MAO", "week": null, "month": null, "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_04301", "client": "BEMOL", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "NINGBO", "dest": "MAO", "week": null, "month": null, "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IM_05288", "client": "ELOX TRADE", "seller": "ROGER", "type": "Import", "modal": "OCEAN", "incoterm": "FOB", "origin": "NINGBO", "dest": "MAO", "week": "2026-05-05 00:00:00", "month": "2026-05", "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_05289", "client": "ADATA", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "TPE", "dest": "VCP", "week": "2026-05-05 00:00:00", "month": "2026-05", "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_05290", "client": "ELOX TRADE", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "GUANGDONG", "dest": "MAO", "week": null, "month": null, "status": "LOST", "remarks": ""}, {"ref": "MAO_IA_05291", "client": "TUTIPLAST", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "XIAMEN", "dest": "MAO", "week": null, "month": null, "status": "LOST", "remarks": ""}, {"ref": "MAO_IM_05292", "client": "ELOX TRADE", "seller": "ROGER", "type": "Import", "modal": "OCEAN", "incoterm": "FOB", "origin": "YANTIAN", "dest": "MAO", "week": null, "month": null, "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_05293", "client": "BEMOL", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "HONG KONG", "dest": "MAO", "week": "2026-05-06 00:00:00", "month": "2026-05", "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_05294", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "INCHEON", "dest": "MAO", "week": null, "month": null, "status": "LOST", "remarks": ""}, {"ref": "MAO_IA_05295", "client": "BEMOL", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "MIAMI", "dest": "MAO", "week": null, "month": null, "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IL_05296", "client": "MERCO FITNESS", "seller": "ROGER", "type": "Import", "modal": "OCEAN", "incoterm": "", "origin": "NINGBO", "dest": "MAO", "week": "2026-05-07 00:00:00", "month": "2026-05", "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IL_05297", "client": "ELOX TRADE", "seller": "ROGER", "type": "Import", "modal": "OCEAN", "incoterm": "FOB", "origin": "SHENZHEN", "dest": "MAO", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IA_05298", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "HONG KONG", "dest": "MAO", "week": "2026-05-07", "month": "2026-05", "status": "LOST", "remarks": ""}, {"ref": "MAO_IA_05299", "client": "TUTIPLAST", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "GUANGDONG", "dest": "MAO", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IA_05302", "client": "ELOX TRADE", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "GUANGDONG", "dest": "MAO", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IA_05303", "client": "SALCOMP", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "HONG KONG", "dest": "MAO", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IA_05304", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "TIANJIN", "dest": "MAO", "week": "2026-05-07", "month": "2026-05", "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_05304", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "TIANJIN", "dest": "MAO", "week": "2026-05-07", "month": "2026-05", "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_05304", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "TIANJIN", "dest": "MAO", "week": "2026-05-07", "month": "2026-05", "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_05304", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "TIANJIN", "dest": "MAO", "week": "2026-05-07", "month": "2026-05", "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_05304", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "TIANJIN", "dest": "MAO", "week": "2026-05-07", "month": "2026-05", "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_05304", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "TIANJIN", "dest": "MAO", "week": "2026-05-07", "month": "2026-05", "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IM_05305", "client": "TUTIPLAST", "seller": "ROGER", "type": "Import", "modal": "OCEAN", "incoterm": "FOB", "origin": "NINGBO", "dest": "MAO", "week": "2026-05-08 00:00:00", "month": "2026-05", "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IA_05306", "client": "HARMAN", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "GUANGZHOU", "dest": "MAO", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IA_05307", "client": "HARMAN", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "GUANGZHOU", "dest": "MAO", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": " MAO_IL_05308", "client": "TUTIPLAST", "seller": "ROGER", "type": "Import", "modal": "OCEAN/LCL", "incoterm": "", "origin": "GUANGZHOU", "dest": "MAO", "week": null, "month": null, "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_05309", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "SHANGHAI", "dest": "MAO", "week": "2026-05-11 00:00:00", "month": "2026-05", "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_05310", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "GUANGDONG", "dest": "MAO", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IM_05311", "client": "ARMOR IIMAK", "seller": "ROGER", "type": "Import", "modal": "OCEAN", "incoterm": "", "origin": "SHENZEN", "dest": "MAO", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IA_05312", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "HONG KONG", "dest": "", "week": null, "month": null, "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_05313", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "FUZHOU", "dest": "MAO", "week": "2026-05-07", "month": "2026-05", "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_05314", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "FUZHOU", "dest": "MAO", "week": "2026-05-13 00:00:00", "month": "2026-05", "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_05315", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "NOI BAI", "dest": "MAO", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IA_05316", "client": "SATBRAS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "SHENZEN", "dest": "MAO", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IA_05317", "client": "ADATA", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "SHANGHAI", "dest": "MAO", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IA_05318", "client": "SX LIGHTING", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "GUANGDONG", "dest": "MAO", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IA_05319", "client": "MAGTEC", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "ZHENGZHOU", "dest": "MAO", "week": "2026-05-12 00:00:00", "month": "2026-05", "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_05320", "client": "MAGTEC", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "ZHENGZHOU", "dest": "MAO", "week": "2026-05-12 00:00:00", "month": "2026-05", "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_05321", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "SHANGHAI", "dest": "MAO", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IA_05323", "client": "HARMAN", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "", "origin": "", "dest": "", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IL_05324", "client": "NIPPON SEIKI", "seller": "ROGER", "type": "Import", "modal": "OCEAN/LCL", "incoterm": "", "origin": "MIAMI", "dest": "MAO", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IA_05325", "client": "ELOX TRADE", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "GUANGDONG", "dest": "MAO", "week": "2026-05-15 00:00:00", "month": "2026-05", "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_05326", "client": "SALCOMP", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "HONG KONG", "dest": "MAO", "week": null, "month": null, "status": "LOST", "remarks": ""}, {"ref": "MAO_IA_05327", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "SHANGAI", "dest": "MAO", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IM_05328", "client": "BEONTAG", "seller": "JARDEL", "type": "Import", "modal": "OCEAN", "incoterm": "", "origin": "SHANGHAI", "dest": "MAO", "week": "2026-05-18 00:00:00", "month": "2026-05", "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_05329", "client": "TUTIPLAST", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "XIAMEN", "dest": "MAO", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IA_05330", "client": "TUTIPLAST", "seller": "ROGER", "type": "Import", "modal": "OCEAN/LCL", "incoterm": "EXW", "origin": "GUANGDONG", "dest": "MAO", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IA_05331", "client": "BEMOL", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "HONG KONG", "dest": "MAO", "week": "2026-05-19 00:00:00", "month": "2026-05", "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_05332", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "NAGOYA", "dest": "MAO", "week": "2026-05-14 00:00:00", "month": "2026-05", "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_05333", "client": "TUTIPLAST", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "GUANGDONG", "dest": "MAO", "week": "2026-05-14 00:00:00", "month": "2026-05", "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_05334", "client": "ELOX TRADE", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "GUANGONG", "dest": "MAO", "week": "2026-05-18 00:00:00", "month": "2026-05", "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_05335", "client": "SATBRAS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "SHENZHEN", "dest": "MAO", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IM_05337", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "OCEAN", "incoterm": "", "origin": "YANTIAN", "dest": "MAO", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IA_05338", "client": "BEMOL", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "MIAMI", "dest": "MAO", "week": "2026-05-21 00:00:00", "month": "2026-05", "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_05339", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "SHENZEN", "dest": "MAO", "week": "2026-05-19 00:00:00", "month": "2026-05", "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_05340", "client": "SALCOMP", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "HONG KONG", "dest": "MAO", "week": null, "month": null, "status": "LOST", "remarks": ""}, {"ref": "MAO_IA_05341", "client": "ELOX TRADE", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "HONG KONG", "dest": "MAO", "week": "2026-05-19 00:00:00", "month": "2026-05", "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_05342", "client": "ADATA", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "INCHEON", "dest": "MAO", "week": "2026-05-19 00:00:00", "month": "2026-05", "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_05343", "client": "ADATA", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "HAN", "dest": "MAO", "week": "2026-05-21 00:00:00", "month": "2026-05", "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_05344", "client": "BEMOL", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "HONG KONG", "dest": "MAO", "week": "2026-05-20 00:00:00", "month": "2026-05", "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_05345", "client": "BEMOL", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "", "origin": "FLORIDA", "dest": "MAO", "week": "2026-05-20 00:00:00", "month": "2026-05", "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_05346", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "Xangai", "dest": "MAO", "week": "2026-05-20 00:00:00", "month": "2026-05", "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_05347", "client": "SALCOMP", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "HONG KONG", "dest": "MAO", "week": null, "month": null, "status": "LOST", "remarks": ""}, {"ref": "MAO_IA_05348", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "HAI PHONG", "dest": "MAO", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IM_05349", "client": "BEMOL", "seller": "ROGER", "type": "Import", "modal": "OCEAN", "incoterm": "", "origin": "SHANGHAI", "dest": "MAO", "week": "2026-05-21 00:00:00", "month": "2026-05", "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_05350", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "HONG KONG", "dest": "MAO", "week": "2026-05-25 00:00:00", "month": "2026-05", "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_05351", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "TIANJIN", "dest": "MAO", "week": "2026-06-08 00:00:00", "month": "2026-06", "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IL_05352", "client": "BEONTAG", "seller": "ROGER", "type": "Import", "modal": "OCEAN/LCL", "incoterm": "", "origin": "", "dest": "MAO", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IM_05353", "client": "HB MAQUINAS", "seller": "JARDEL", "type": "Import", "modal": "OCEAN", "incoterm": "", "origin": "QINGDAO", "dest": "MAO", "week": "2026-05-25 00:00:00", "month": "2026-05", "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IA_05354", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "HONG KONG", "dest": "MAO", "week": "2026-05-20 00:00:00", "month": "2026-05", "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IM_05355", "client": "MERCO FITNESS", "seller": "JARDEL", "type": "Import", "modal": "OCEAN", "incoterm": "FOB", "origin": "NINGBO", "dest": "MAO", "week": "2026-05-25 00:00:00", "month": "2026-05", "status": "LOST", "remarks": ""}, {"ref": "MAO_IM_05356", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "HONG KONG", "dest": "MAO", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IA_05357", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "SHENZEN", "dest": "MAO", "week": "2026-05-25 00:00:00", "month": "2026-05", "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IA_05358", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "HONG KONG", "dest": "MAO", "week": "2026-05-26 00:00:00", "month": "2026-05", "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_05359", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "SHENZEN", "dest": "MAO", "week": "2026-05-26 00:00:00", "month": "2026-05", "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_05360", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "SHENZEN", "dest": "MAO", "week": "2026-05-26 00:00:00", "month": "2026-05", "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_05361", "client": "TUTIPLAST", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "GUANGDONG", "dest": "MAO", "week": "2026-05-26 00:00:00", "month": "2026-05", "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_05362", "client": "TUTIPLAST", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "SHENZEN", "dest": "MAO", "week": "2026-05-25 00:00:00", "month": "2026-05", "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IL_05363", "client": "MERCO FITNESS", "seller": "JARDEL", "type": "Import", "modal": "OCEAN/LCL", "incoterm": "FOB", "origin": "XIAMEN", "dest": "MAO", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IA_05364", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "HONG KONG", "dest": "MAO", "week": "2026-05-26 00:00:00", "month": "2026-05", "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_05365", "client": "BEONTAG", "seller": "JARDEL", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "SHANGHAI", "dest": "MAO", "week": "2026-05-27 00:00:00", "month": "2026-05", "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IA_05366", "client": "MERCO FITNESS", "seller": "JARDEL", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "Shenzhen", "dest": "MAO", "week": "2026-05-29 00:00:00", "month": "2026-05", "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_05367", "client": "SALCOMP", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "SHANGHAI", "dest": "MAO", "week": "2026-05-26 00:00:00", "month": "2026-05", "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IM_05368", "client": "BEONTAG", "seller": "JARDEL", "type": "Import", "modal": "OCEAN", "incoterm": "FOB", "origin": "SHANGHAI", "dest": "MAO", "week": "2026-05-27 00:00:00", "month": "2026-05", "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IA_05370", "client": "HARMAN", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "SHANGAI", "dest": "MAO", "week": "2026-06-01 00:00:00", "month": "2026-06", "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_05371", "client": "MAGTEC", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "", "origin": "PORTLAND", "dest": "MAO", "week": "2026-05-28 00:00:00", "month": "2026-05", "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_05372", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "SHENZHEN", "dest": "MAO", "week": "2026-05-22 00:00:00", "month": "2026-05", "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IL_05373", "client": "HB MAQUINAS", "seller": "JARDEL", "type": "Import", "modal": "OCEAN/LCL", "incoterm": "", "origin": "XIAMEN", "dest": "MAO", "week": "2026-06-01 00:00:00", "month": "2026-06", "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IL_05373", "client": "HB MAQUINAS", "seller": "JARDEL", "type": "Import", "modal": "OCEAN/LCL", "incoterm": "FOB", "origin": "XIAMEN", "dest": "MAO", "week": "2026-05-28 00:00:00", "month": "2026-05", "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_05374", "client": "MERCO FITNESS", "seller": "JARDEL", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "SHANGHAI", "dest": "MAO", "week": "2026-05-29 00:00:00", "month": "2026-05", "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IA_06375", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "HONG KONG", "dest": "MAO", "week": "2026-05-29 00:00:00", "month": "2026-05", "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_06376", "client": "SALCOMP", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "HONG KONG", "dest": "MAO", "week": "2026-05-29 00:00:00", "month": "2026-05", "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_06377", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "HONG KONG", "dest": "MAO", "week": "2026-06-01 00:00:00", "month": "2026-06", "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_06378", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "SHENZHEN", "dest": "MAO", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IM_06379", "client": "HB MAQUINAS", "seller": "JARDEL", "type": "Import", "modal": "OCEAN", "incoterm": "", "origin": "QINGDAO", "dest": "MAO", "week": null, "month": null, "status": "LOST", "remarks": ""}, {"ref": "MAO_IA_06380", "client": "TUTIPLAST", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "GUANDDONG", "dest": "MAO", "week": "2026-06-01 00:00:00", "month": "2026-06", "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IA_06381", "client": "TUTIPLAST", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "SHENZEN", "dest": "MAO", "week": "2026-06-01 00:00:00", "month": "2026-06", "status": "LOST", "remarks": ""}, {"ref": "MAO_IA_06382", "client": "TUTIPLAST", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "XIAMEN", "dest": "MAO", "week": "2026-06-01 00:00:00", "month": "2026-06", "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IA_06383", "client": "ELOX TRADE", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "FLORIDA", "dest": "MAO", "week": "2026-06-01 00:00:00", "month": "2026-06", "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IM_06384", "client": "MERCO FITNESS", "seller": "JARDEL", "type": "Import", "modal": "OCEAN", "incoterm": "FOB", "origin": "NINGBO", "dest": "MAO", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IM_06385", "client": "MINEBEA", "seller": "JARDEL", "type": "N/D", "modal": "OCEAN", "incoterm": "", "origin": "", "dest": "", "week": null, "month": null, "status": "UNDER REVIEW", "remarks": ""}, {"ref": "MAO_IA_06387", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "HONG KONG", "dest": "MAO", "week": null, "month": null, "status": "APPROVED", "remarks": ""}, {"ref": " MAO_IA_06388", "client": "SALCOMP", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "HONG KONG", "dest": "MAO", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IA_06389", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "SHANGHAI", "dest": "MAO", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IA_06390", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "", "origin": "SHANGHAI", "dest": "MAO", "week": null, "month": null, "status": "APPROVED", "remarks": ""}, {"ref": "MAO_ROD_06391", "client": "VERDE BRASIL", "seller": "JARDEL", "type": "N/D", "modal": "N/D", "incoterm": "", "origin": "", "dest": "", "week": null, "month": null, "status": "UNDER REVIEW", "remarks": ""}, {"ref": "MAO_IA_06392", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "SHANGHAI", "dest": "MAO", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IA_06393", "client": "ADATA", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "CGO", "dest": "MAO", "week": null, "month": null, "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_06394", "client": "TUTIPLAST", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "GUANGDONG", "dest": "MAO", "week": "2026-06-08 00:00:00", "month": "2026-06", "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_06395", "client": "BEMOL", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "MIA", "dest": "MAO", "week": null, "month": null, "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_06396", "client": "BEMOL", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "MIA", "dest": "MAO", "week": null, "month": null, "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_06397", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "VIETNÃ", "dest": "MAO", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IA_06398", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "INCHEON", "dest": "MAO", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IM_06399", "client": "ELOX TRADE", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FOB", "origin": "NINGBO", "dest": "MAO", "week": "2026-06-11 00:00:00", "month": "2026-06", "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_06400", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "HONG KONG", "dest": "MAO", "week": null, "month": null, "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IM_06401", "client": "BEMOL", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FOB", "origin": "NINGBO", "dest": "MAO", "week": null, "month": null, "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_06402", "client": "MERCO FITNESS", "seller": "JARDEL", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "XIAMEN", "dest": "MAO", "week": null, "month": null, "status": "LOST", "remarks": ""}, {"ref": "MAO_IA_06403", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "Zhejiang", "dest": "MAO", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IM_06404", "client": "BEMOL", "seller": "ROGER", "type": "Import", "modal": "OCEAN", "incoterm": "FOB", "origin": "NINGBO", "dest": "MAO", "week": null, "month": null, "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IM_06405", "client": "BEMOL", "seller": "ROGER", "type": "Import", "modal": "OCEAN", "incoterm": "FOB", "origin": "QINGDAO", "dest": "MAO", "week": null, "month": null, "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IM_06406", "client": "ELOX TRADE", "seller": "ROGER", "type": "Import", "modal": "OCEAN", "incoterm": "FOB", "origin": "YANTIAN", "dest": "MAO", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IA_06407", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "GUANGDONG", "dest": "MAO", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IA_06408", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "FUZHOU", "dest": "MAO", "week": null, "month": null, "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_06409", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "SHANGHAI", "dest": "MAO", "week": null, "month": null, "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_06410", "client": "TUTIPLAST", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "SHENZHEN", "dest": "MAO", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IA_06411", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "SHENZHEN", "dest": "MAO", "week": "2026-06-11 00:00:00", "month": "2026-06", "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IA_06412", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "INCHEON", "dest": "MAO", "week": "2026-06-11 00:00:00", "month": "2026-06", "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IA_06413", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "INCHEON", "dest": "MAO", "week": "2026-06-11 00:00:00", "month": "2026-06", "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IM_06415", "client": "BEMOL", "seller": "ROGER", "type": "Import", "modal": "OCEAN", "incoterm": "FOB", "origin": "NINGBO", "dest": "MAO", "week": "2026-06-11 00:00:00", "month": "2026-06", "status": "LOST", "remarks": ""}, {"ref": "MAO_IA_06416", "client": "BEMOL", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "HONG KONG", "dest": "MAO", "week": null, "month": null, "status": "LOST", "remarks": ""}, {"ref": "MAO_IA_06417", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "INCHEON", "dest": "MAO", "week": null, "month": null, "status": "LOST", "remarks": ""}, {"ref": "MAO_IM_06419", "client": "BEONTAG", "seller": "JARDEL", "type": "Import", "modal": "OCEAN", "incoterm": "FOB", "origin": "XIAMEN", "dest": "MAO", "week": null, "month": null, "status": "LOST", "remarks": ""}, {"ref": "MAO_IM_06420", "client": "BEONTAG", "seller": "JARDEL", "type": "Import", "modal": "OCEAN", "incoterm": "FOB", "origin": "SHANGHAI", "dest": "MAO", "week": "2026-06-12 00:00:00", "month": "2026-06", "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IM_06421", "client": "BEONTAG", "seller": "JARDEL", "type": "Import", "modal": "OCEAN", "incoterm": "FOB", "origin": "XIAMEN", "dest": "MAO", "week": "2026-06-12 00:00:00", "month": "2026-06", "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IM_06422", "client": "BEONTAG", "seller": "JARDEL", "type": "Import", "modal": "OCEAN", "incoterm": "FOB", "origin": "NINGBO", "dest": "MAO", "week": "2026-06-12 00:00:00", "month": "2026-06", "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IM_06423", "client": "BEONTAG", "seller": "JARDEL", "type": "Import", "modal": "OCEAN", "incoterm": "FOB", "origin": "QINGDAO", "dest": "MAO", "week": null, "month": null, "status": "LOST", "remarks": ""}, {"ref": "MAO_IM_06424", "client": "BEONTAG", "seller": "JARDEL", "type": "Import", "modal": "OCEAN", "incoterm": "EXW", "origin": "FRANÇA", "dest": "MAO", "week": "2026-06-12 00:00:00", "month": "2026-06", "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IA_06425", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "SHENZHEN", "dest": "MAO", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IM_06426", "client": "MERCO FITNESS", "seller": "JARDEL", "type": "Import", "modal": "OCEAN", "incoterm": "FOB", "origin": "XIAMEN", "dest": "MAO", "week": null, "month": null, "status": "LOST", "remarks": ""}, {"ref": "MAO_IA_06427", "client": "HARMAN", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "PEQUIN", "dest": "MAO", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IM_06428", "client": "HB MAQUINAS", "seller": "JARDEL", "type": "Import", "modal": "OCEAN", "incoterm": "FOB", "origin": "QINGDAO", "dest": "MAO", "week": "2026-06-15 00:00:00", "month": "2026-06", "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IA_06429", "client": "TUTIPLAST", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "SHENZHEN", "dest": "MAO", "week": null, "month": null, "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_06430", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "Zhengzhou Xinzheng", "dest": "MAO", "week": null, "month": null, "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_06430", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "HONG KONG", "dest": "MAO", "week": "2026-06-15 00:00:00", "month": "2026-06", "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IA_06431", "client": "TUTIPLAST", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "PEQUIN", "dest": "MAO", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IA_06432", "client": "ELOX TRADE", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "Ezhou Huahu", "dest": "MAO", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IA_06433", "client": "ADATA", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "INCHEON", "dest": "MAO", "week": null, "month": null, "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IM_06434", "client": "ELOX TRADE", "seller": "ROGER", "type": "Import", "modal": "OCEAN", "incoterm": "FOB", "origin": "NINGBO", "dest": "MAO", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IA_06436", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "SEA/AIR", "incoterm": "FCA", "origin": "YANTIAN", "dest": "MAO", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IM_06437", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "OCEAN", "incoterm": "FOB", "origin": "QINGDAO", "dest": "MAO", "week": "2026-06-16 00:00:00", "month": "2026-06", "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IA_06438", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "HONG KONG", "dest": "MAO", "week": null, "month": null, "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_06439", "client": "TUTIPLAST", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "PEQUIN", "dest": "MAO", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IM_06440", "client": "ELOX TRADE", "seller": "ROGER", "type": "Import", "modal": "OCEAN", "incoterm": "FOB", "origin": "SHANGHAI", "dest": "MAO", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IA_06441", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "ZHENGZHOU", "dest": "MAO", "week": "2026-06-16 00:00:00", "month": "2026-06", "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_06442", "client": "HARMAN", "seller": "ROGER", "type": "Import", "modal": "SEA/AIR", "incoterm": "EXW", "origin": "YANTIAN", "dest": "MAO", "week": "2026-06-16 00:00:00", "month": "2026-06", "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IA_06443", "client": "ADATA", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "COREIA", "dest": "VCP", "week": "2026-06-16 00:00:00", "month": "2026-06", "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_06444", "client": "TUTIPLAST", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "*", "origin": "*", "dest": "*", "week": null, "month": null, "status": "LOST", "remarks": ""}, {"ref": " MAO_IA_06445", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "Zhengzhou Xinzheng", "dest": "MAO", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IA_06446", "client": "SALCOMP", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "CHINA", "dest": "MAO", "week": null, "month": null, "status": "LOST", "remarks": ""}, {"ref": "MAO_IA_06447", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "HONG KONG", "dest": "MAO", "week": "2026-06-17 00:00:00", "month": "2026-06", "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IA_06448", "client": "ELOX TRADE", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "XIAMEN", "dest": "MAO", "week": null, "month": null, "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_06449", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "HONG KONG", "dest": "MAO", "week": null, "month": null, "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_06450", "client": "SALCOMP", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "Ezhou Huahu", "dest": "MAO", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IA_06451", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "", "origin": "SHENZHEN", "dest": "MAO", "week": "2026-06-18 00:00:00", "month": "2026-06", "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IM_06452", "client": "ELOX TRADE", "seller": "ROGER", "type": "Import", "modal": "OCEAN", "incoterm": "FOB", "origin": "YANTIAN", "dest": "MAO", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IM_06453", "client": "ELOX TRADE", "seller": "ROGER", "type": "Import", "modal": "OCEAN", "incoterm": "EXW", "origin": "YANTIAN", "dest": "MAO", "week": null, "month": null, "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IM_06454", "client": "ELOX TRADE", "seller": "ROGER", "type": "Import", "modal": "OCEAN", "incoterm": "EXW", "origin": "YANTIAN", "dest": "MAO", "week": null, "month": null, "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IM_06455", "client": "ADATA", "seller": "ROGER", "type": "Import", "modal": "OCEAN", "incoterm": "EXW", "origin": "SHENZHEN", "dest": "MAO", "week": "2026-06-19 00:00:00", "month": "2026-06", "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IL_06456", "client": "BEONTAG", "seller": "JARDEL", "type": "Import", "modal": "AIR", "incoterm": "FOB", "origin": "SHANGHAI", "dest": "MAO", "week": "2026-06-22 00:00:00", "month": "2026-06", "status": "LOST", "remarks": ""}, {"ref": "MAO_IA_06457", "client": "TUTIPLAST", "seller": "N/D", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "SUZHOU", "dest": "MAO", "week": "2026-06-22 00:00:00", "month": "2026-06", "status": "UNDER REVIEW", "remarks": ""}, {"ref": "MAO_IA_06458", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "NOI BAI", "dest": "MAO", "week": "2026-06-23 00:00:00", "month": "2026-06", "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IM_06459", "client": "BEMOL", "seller": "ROGER", "type": "Import", "modal": "OCEAN", "incoterm": "FOB", "origin": "YANTIAN", "dest": "MAO", "week": "2026-07-01 00:00:00", "month": "2026-07", "status": "UNDER REVIEW", "remarks": ""}, {"ref": "MAO_IM_06460", "client": "BEMOL", "seller": "ROGER", "type": "Import", "modal": "OCEAN", "incoterm": "FOB", "origin": "SHENZHEN", "dest": "MAO", "week": "2026-06-23 00:00:00", "month": "2026-06", "status": "LOST", "remarks": ""}, {"ref": " MAO_IM_06461", "client": "BEMOL", "seller": "ROGER", "type": "Import", "modal": "OCEAN", "incoterm": "FOB", "origin": "NINGBO", "dest": "MAO", "week": "2026-06-23 00:00:00", "month": "2026-06", "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IA_06462", "client": "TUTIPLAST", "seller": "ROGER", "type": "Import", "modal": "SEA/AIR", "incoterm": "EXW", "origin": "GUANGDONG", "dest": "MAO", "week": "2026-06-23 00:00:00", "month": "2026-06", "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IM_06463", "client": "BEMOL", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FOB", "origin": "NINGBO", "dest": "MAO", "week": "2026-06-23 00:00:00", "month": "2026-06", "status": "LOST", "remarks": ""}, {"ref": "MAO_IM_06464", "client": "ELOX TRADE", "seller": "N/D", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "HONG KONG", "dest": "MAO", "week": "2026-06-24 00:00:00", "month": "2026-06", "status": "UNDER REVIEW", "remarks": ""}, {"ref": "MAO_IA_06464", "client": "ELOX TRADE", "seller": "ROGER", "type": "Import", "modal": "OCEAN", "incoterm": "", "origin": "NINGBO", "dest": "MAO", "week": "2026-06-24 00:00:00", "month": "2026-06", "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IM_06465", "client": "BEMOL", "seller": "ROGER", "type": "Import", "modal": "OCEAN", "incoterm": "EXW", "origin": "MANZANILLO", "dest": "MAO", "week": "2026-06-24 00:00:00", "month": "2026-06", "status": "UNDER REVIEW", "remarks": ""}, {"ref": "MAO_IA_06465", "client": "ELOX TRADE", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "WENZHOU", "dest": "MAO", "week": "2026-06-24 00:00:00", "month": "2026-06", "status": "LOST", "remarks": ""}, {"ref": "VCP_IA_06466", "client": "ADATA", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "JAPAN", "dest": "MAO", "week": "2026-06-24 00:00:00", "month": "2026-06", "status": "UNDER REVIEW", "remarks": ""}, {"ref": "MAO_IA_06467", "client": "BEONTAG", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "ALEMANHA", "dest": "MAO", "week": null, "month": null, "status": "UNDER REVIEW", "remarks": ""}, {"ref": "MAO_IA_06468", "client": "VERDE BRASIL", "seller": "JARDEL", "type": "Import", "modal": "AIR", "incoterm": "", "origin": "GUANGZHOU", "dest": "MAO", "week": "2026-06-24 00:00:00", "month": "2026-06", "status": "UNDER REVIEW", "remarks": ""}, {"ref": "MAO_IA_06469", "client": "MERCO FITNESS", "seller": "JARDEL", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "XIAMEN", "dest": "MAO", "week": "2026-06-24 00:00:00", "month": "2026-06", "status": "UNDER REVIEW", "remarks": ""}, {"ref": "MAO_IM_06470", "client": "BEONTAG", "seller": "N/D", "type": "Import", "modal": "OCEAN", "incoterm": "FOB", "origin": "NINGBO", "dest": "MAO", "week": "2026-06-25 00:00:00", "month": "2026-06", "status": "LOST", "remarks": ""}, {"ref": "MAO_IM_06471", "client": "BEONTAG", "seller": "JARDEL", "type": "Import", "modal": "OCEAN", "incoterm": "FOB", "origin": "QINGDAO", "dest": "MAO", "week": "2026-06-25 00:00:00", "month": "2026-06", "status": "UNDER REVIEW", "remarks": ""}, {"ref": "MAO_IM_06471", "client": "BEONTAG", "seller": "N/D", "type": "Import", "modal": "OCEAN", "incoterm": "FOB", "origin": "QINGDAO", "dest": "MAO", "week": "2026-06-26 00:00:00", "month": "2026-06", "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IM_06472", "client": "BEONTAG", "seller": "N/D", "type": "Import", "modal": "OCEAN", "incoterm": "FOB", "origin": "XIAMEN", "dest": "MAO", "week": "2026-06-25 00:00:00", "month": "2026-06", "status": "LOST", "remarks": ""}, {"ref": "MAO_IM_06473", "client": "ELOX TRADE", "seller": "N/D", "type": "Import", "modal": "OCEAN", "incoterm": "", "origin": "YANTIAN", "dest": "MAO", "week": "2026-06-26 00:00:00", "month": "2026-06", "status": "N/D", "remarks": ""}, {"ref": "MAO_IM_06474", "client": "ALFAIA TAIWAN", "seller": "N/D", "type": "Import", "modal": "OCEAN", "incoterm": "FOB", "origin": "TAIWAN", "dest": "MAO", "week": "2026-06-25 00:00:00", "month": "2026-06", "status": "NO RESPONSE", "remarks": ""}, {"ref": "SSZ_IM_06475", "client": "ADATA", "seller": "ROGER", "type": "Import", "modal": "OCEAN", "incoterm": "", "origin": "SHENZHEN", "dest": "MAO", "week": "2026-06-29 00:00:00", "month": "2026-06", "status": "UNDER REVIEW", "remarks": ""}, {"ref": "MAO_IM_06475", "client": "ADATA", "seller": "ROGER", "type": "Import", "modal": "OCEAN", "incoterm": "EXW", "origin": "SHENZHEN", "dest": "SANTOS", "week": "2026-06-26 00:00:00", "month": "2026-06", "status": "UNDER REVIEW", "remarks": ""}, {"ref": "MAO_IA_06476", "client": "BEMOL", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "", "origin": "", "dest": "", "week": "2026-07-01 00:00:00", "month": "2026-07", "status": "UNDER REVIEW", "remarks": ""}, {"ref": "MAO_IL_06478", "client": "BEONTAG", "seller": "N/D", "type": "Import", "modal": "OCEAN", "incoterm": "FOB", "origin": "TAIWAN", "dest": "MAO", "week": "2026-06-26 00:00:00", "month": "2026-06", "status": "NO RESPONSE", "remarks": ""}, {"ref": "MAO_IM_06480", "client": "BEMOL", "seller": "ROGER", "type": "Import", "modal": "OCEAN", "incoterm": "EXW", "origin": "MANZANILLO", "dest": "", "week": "2026-06-26 00:00:00", "month": "2026-06", "status": "UNDER REVIEW", "remarks": ""}, {"ref": "MAO_IM_06481", "client": "ELOX TRADE", "seller": "ROGER", "type": "Import", "modal": "OCEAN", "incoterm": "", "origin": "SHENZHEN", "dest": "MAO", "week": "2026-06-26 00:00:00", "month": "2026-06", "status": "UNDER REVIEW", "remarks": ""}, {"ref": "MAO_CAB_06482", "client": "TUPAMOTOS", "seller": "ROGER", "type": "Import", "modal": "OCEAN", "incoterm": "", "origin": "SALVADOR", "dest": "MAO", "week": "2026-07-01 00:00:00", "month": "2026-07", "status": "UNDER REVIEW", "remarks": ""}, {"ref": "MAO_IA_06483", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "SHANGHAI", "dest": "MAO", "week": "2026-06-30 00:00:00", "month": "2026-06", "status": "UNDER REVIEW", "remarks": ""}, {"ref": "MAO_IM_06484", "client": "BEMOL", "seller": "ROGER", "type": "Import", "modal": "OCEAN", "incoterm": "", "origin": "SHANGHAI", "dest": "MAO", "week": "2026-06-29 00:00:00", "month": "2026-06", "status": "UNDER REVIEW", "remarks": ""}, {"ref": "MAO_IA_06485", "client": "INVENTUS", "seller": "N/D", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "HONG KONG", "dest": "MAO", "week": "2026-06-30 00:00:00", "month": "2026-06", "status": "N/D", "remarks": ""}, {"ref": "MAO_IA_06486", "client": "BEONTAG", "seller": "N/D", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "ALEMANHA", "dest": "BRASIL", "week": null, "month": null, "status": "LOST", "remarks": ""}, {"ref": "MAO_ROD_06487", "client": "VALGROUP", "seller": "JARDEL", "type": "Import", "modal": "TRUCK", "incoterm": "FTL", "origin": "SÃO PAULO", "dest": "MAO", "week": "2026-07-01 00:00:00", "month": "2026-07", "status": "UNDER REVIEW", "remarks": ""}, {"ref": "MAO_IA_07488", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "SHANGHAI", "dest": "MAO", "week": "2026-07-01 00:00:00", "month": "2026-07", "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_07489", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "N/D", "incoterm": "", "origin": "", "dest": "", "week": "2026-07-01 00:00:00", "month": "2026-07", "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_07490", "client": "TUTIPLAST", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "Wuhu Wanli Airport", "dest": "MAO", "week": "2026-07-01 00:00:00", "month": "2026-07", "status": "APPROVED", "remarks": ""}, {"ref": "VCP_IA_07491", "client": "ADATA", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "JAPAN", "dest": "SÃO PAULO", "week": "2026-07-01 00:00:00", "month": "2026-07", "status": "UNDER REVIEW", "remarks": ""}, {"ref": "MAO_IA_07492", "client": "ADATA", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "TAOYUAN", "dest": "SÃO PAULO", "week": "2026-07-01 00:00:00", "month": "2026-07", "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_07493", "client": "INVENTUS", "seller": "N/D", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "SHANGHAI", "dest": "MAO", "week": "2026-07-03 00:00:00", "month": "2026-07", "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_07495", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "HONG KONG", "dest": "MAO", "week": "2026-07-02 00:00:00", "month": "2026-07", "status": "UNDER REVIEW", "remarks": ""}, {"ref": "MAO_IA_07494", "client": "BEMOL", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "HONG KONG", "dest": "MAO", "week": "2026-07-02 00:00:00", "month": "2026-07", "status": "UNDER REVIEW", "remarks": ""}, {"ref": "MAO_IA_07496", "client": "BEMOL", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "HONG KONG", "dest": "MAO", "week": "2026-07-02 00:00:00", "month": "2026-07", "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_07497", "client": "TUTIPLAST", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "SHANGHAI", "dest": "MAO", "week": "2026-07-02 00:00:00", "month": "2026-07", "status": "APPROVED", "remarks": ""}, {"ref": " MAO_IA_07498", "client": "ELOX TRADE", "seller": "N/D", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "CHINA", "dest": "MAO", "week": "2026-07-03 00:00:00", "month": "2026-07", "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_07498", "client": "ELOX TRADE", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "DONGGUAN", "dest": "MAO", "week": "2026-07-03 00:00:00", "month": "2026-07", "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IM_07499", "client": "ELOX TRADE", "seller": "ROGER", "type": "Import", "modal": "OCEAN", "incoterm": "EXW", "origin": "HONG KONG", "dest": "MAO", "week": "2026-07-02 00:00:00", "month": "2026-07", "status": "UNDER REVIEW", "remarks": ""}, {"ref": " MAO_IM_07500", "client": "BEONTAG", "seller": "N/D", "type": "Import", "modal": "OCEAN", "incoterm": "FCL", "origin": "CNSHA", "dest": "MAO", "week": "2026-07-03 00:00:00", "month": "2026-07", "status": "UNDER REVIEW", "remarks": ""}, {"ref": " MAO_IM_07501", "client": "BEONTAG", "seller": "N/D", "type": "Import", "modal": "OCEAN", "incoterm": "FCL", "origin": "CNSHA", "dest": "MAO", "week": "2026-07-03 00:00:00", "month": "2026-07", "status": "UNDER REVIEW", "remarks": ""}, {"ref": "MAO_IA_07502", "client": "HARMAN", "seller": "N/D", "type": "Import", "modal": "AIR", "incoterm": "", "origin": "", "dest": "", "week": null, "month": null, "status": "N/D", "remarks": ""}, {"ref": "MAO_IA_07503", "client": "BEMOL", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "MIAMI", "dest": "MAO", "week": "2026-07-04 00:00:00", "month": "2026-07", "status": "APPROVED", "remarks": ""}, {"ref": "MAO_IA_07504", "client": "INVENTUS", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "FCA", "origin": "Xangai Pudong", "dest": "MAO", "week": "2026-07-04 00:00:00", "month": "2026-07", "status": "UNDER REVIEW", "remarks": ""}, {"ref": "MAO_IA_09477", "client": "OMNI", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "MIA", "dest": "MAO", "week": "2026-03-11 00:00:00", "month": "2026-03", "status": "NO RESPONSE", "remarks": "CLIENTE NÃO RESPONDEU O E-MAIL"}, {"ref": "MAO_IA_07507", "client": "SYNTECH", "seller": "ROGER", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "HONG KONG", "dest": "MAO", "week": "2026-07-06 00:00:00", "month": "2026-07", "status": "UNDER REVIEW", "remarks": ""}, {"ref": "MAO_IA_07506", "client": "BEONTAG", "seller": "JARDEL", "type": "Import", "modal": "AIR", "incoterm": "EXW", "origin": "ALEMANHA", "dest": "BRASIL", "week": "2026-07-06 00:00:00", "month": "2026-07", "status": "UNDER REVIEW", "remarks": ""}];

/* ===== FIREBASE REALTIME DATABASE ===== */
const firebaseConfig = {
  apiKey: "AIzaSyBTRzLDGEsZWXDGvEksycxemUmmlBh3cLw",
  authDomain: "bdg-sds.firebaseapp.com",
  databaseURL: "https://bdg-sds-default-rtdb.firebaseio.com",
  projectId: "bdg-sds",
  storageBucket: "bdg-sds.firebasestorage.app",
  messagingSenderId: "30015192050",
  appId: "1:30015192050:web:7a9787ae83ab987fac496b"
};
const app = getApp();
const db = getDatabase(app, firebaseConfig.databaseURL);
const cotacoesRef = dbRef(db, 'cotacoes');
const newlyAddedKeys = new Set();
let seeded = false;

/* ===== Fase Extra · Etapa 3 — Camada 1: Configurações do sistema (Plano
   Fase Extra, Seção 3.1). Versão enxuta para esta rodada: um único nó
   compartilhado com os valores padrão hoje "fixos" na planilha (custos
   fixos de Miami e a taxa de screening mais usada), editáveis sem alterar
   código. Quando um parceiro reajusta uma tarifa, atualiza aqui uma vez —
   cotações novas continuam podendo "usar padrão" ou "sobrescrever nesta
   cotação" (Seção 3.3 do plano), sem exigir uma tela de administração
   separada nesta primeira versão. */
const configRef = dbRef(db, 'configuracoes/taxasMiami');
let configPadraoMiami = {};
onValue(configRef, (snapshot) => { configPadraoMiami = snapshot.val() || {}; });

function setSyncPill(status){
  const el = document.getElementById('pillSync');
  if(!el) return;
  const map = {
    ok: {text:'☁ Nuvem conectada', cls:'ok'},
    seeding: {text:'☁ Enviando base inicial...', cls:'wait'},
    error: {text:'⚠ Sem conexão com a nuvem', cls:'error'},
    reconnecting: {text:'☁ Reconectando...', cls:'wait'}
  };
  const m = map[status] || map.ok;
  el.textContent = m.text;
  el.className = 'pill sync-' + m.cls;
}

onValue(cotacoesRef, (snapshot) => {
  const val = snapshot.val();
  if (val) {
    RAW = Object.entries(val).map(([key, v]) => ({ ...v, _key: key, _new: newlyAddedKeys.has(key) }));
    setSyncPill('ok');
  } else if (!seeded) {
    seeded = true;
    setSyncPill('seeding');
    const seedObj = {};
    RAW.forEach(item => {
      const newKey = dbPush(cotacoesRef).key;
      const clean = { ...item }; delete clean._new;
      seedObj[newKey] = clean;
    });
    dbSet(cotacoesRef, seedObj).then(()=>setSyncPill('ok')).catch(()=>setSyncPill('error'));
    return; // wait for the resulting onValue fire with real data
  }
  recomputeMeta(); render();
  if(window.markLoadingStep) window.markLoadingStep('data');
}, (err) => { setSyncPill('error'); console.error(err); });

/* ---------- RECONEXÃO MANUAL/AUTOMÁTICA (sem deslogar) ----------
   Em apps instalados como PWA (sem barra de navegador), não existe um
   "atualizar" nativo — por isso este botão força uma nova conexão com
   o Firebase sem precisar sair da conta e entrar de novo. A sessão do
   usuário nunca é tocada aqui: só a conexão com o banco é reiniciada. */
let rtdbConnected = true;
let manualReconnecting = false;

/* ---------- FASE 2: reconexão automática por backoff exponencial ----------
   Antes, a reconexão só acontecia com o clique manual no botão, ou nos
   eventos 'online'/'visibilitychange'. Isso deixa o app "morto" se a queda
   de conexão acontecer com o app parado em segundo plano sem nenhum desses
   dois eventos disparando (comum em Wi-Fi instável sem trocar de rede).
   Agora, sempre que '.info/connected' ficar false, um timer tenta reconectar
   sozinho em intervalos crescentes (1s → 2s → 4s → 8s → 16s → 30s, teto),
   sem qualquer ação do usuário. O timer é cancelado assim que a conexão
   volta, e reiniciado do zero (1s) na próxima queda. */
let backoffMs = 1000;
const BACKOFF_MAX_MS = 30000;
let backoffTimer = null;

function clearBackoff(){
  if(backoffTimer){ clearTimeout(backoffTimer); backoffTimer = null; }
  backoffMs = 1000;
}
function backoffTick(){
  backoffTimer = null;
  if(!rtdbConnected){
    // Fase 2 (correção da regressão de autenticação): se houver um
    // login/cadastro/redefinição de senha em andamento (window.__authBusy),
    // adia esta tentativa em vez de mexer na conexão agora — evita que o
    // goOffline/goOnline do Realtime Database e o refresh de token
    // concorram com a chamada de auth em andamento e produzam
    // "auth/network-request-failed". A tentativa não é perdida: só
    // reagendada logo em seguida, sem avançar o backoff.
    if(window.__authBusy){
      backoffTimer = setTimeout(backoffTick, 800);
      return;
    }
    attemptReconnect(false);
    backoffMs = Math.min(backoffMs * 2, BACKOFF_MAX_MS);
    scheduleAutoReconnect();
  }
}
function scheduleAutoReconnect(){
  if(backoffTimer) return; // já agendado
  backoffTimer = setTimeout(backoffTick, backoffMs);
}

onValue(dbRef(db, '.info/connected'), (snap) => {
  rtdbConnected = snap.val() === true;
  if(!manualReconnecting) setSyncPill(rtdbConnected ? 'ok' : 'error');
  if(rtdbConnected){
    clearBackoff();
  }else{
    scheduleAutoReconnect();
  }
});

function attemptReconnect(isManual){
  if(manualReconnecting) return;
  // Fase 2 (correção da regressão de autenticação): nunca mexe na conexão
  // do Realtime Database (goOffline/goOnline) enquanto um login, cadastro
  // ou redefinição de senha está em andamento — isso é o que gerava
  // "auth/network-request-failed" de forma intermitente, por concorrência
  // entre a chamada de auth e o ciclo de reconexão automática. Uma
  // tentativa manual (clique no botão) durante esse período simplesmente
  // aguarda um instante e tenta de novo.
  if(window.__authBusy){
    if(isManual) setTimeout(function(){ attemptReconnect(true); }, 800);
    return;
  }
  manualReconnecting = true;
  setSyncPill('reconnecting');
  const btn = document.getElementById('btnReconnect');
  const label = document.getElementById('btnReconnectLabel');
  if(btn){ btn.disabled = true; btn.classList.add('spinning'); }
  if(label) label.textContent = 'Reconectando...';
  try{ dbGoOffline(db); }catch{}
  setTimeout(function(){
    try{ dbGoOnline(db); }catch{}
    // Renova o token de autenticação em segundo plano, sem deslogar — mas
    // SOMENTE quando a reconexão foi disparada manualmente (clique no
    // botão). Forçar esse refresh a cada tentativa automática (a cada
    // 1s/2s/4s.../30s enquanto offline) é o que causava a regressão de
    // rede: o refresh forçado falha imediatamente sem internet e, em
    // conexões instáveis, chega a concorrer com um login em andamento.
    // A própria SDK do Firebase Auth já renova o token sozinha quando
    // necessário, então esse refresh manual é só um reforço opcional.
    if(isManual){
      var authInst = getAuth(app);
      if(authInst.currentUser){ authInst.currentUser.getIdToken(true).catch(function(){}); }
    }
    setTimeout(function(){
      manualReconnecting = false;
      if(btn){ btn.disabled = false; btn.classList.remove('spinning'); }
      if(label) label.textContent = 'Atualizar conexão';
      if(rtdbConnected){
        setSyncPill('ok');
      }else{
        setSyncPill('error');
        // Ainda offline após a tentativa leve: só quando a ação foi manual
        // (clique no botão) recarrega a página — a sessão continua salva
        // (persistência local), então isso nunca desloga o usuário.
        if(isManual) location.reload();
      }
    }, 3200);
  }, 400);
}
document.getElementById('btnReconnect')?.addEventListener('click', function(){ attemptReconnect(true); });
window.addEventListener('online', function(){ if(!rtdbConnected) attemptReconnect(false); });
document.addEventListener('visibilitychange', function(){
  if(document.visibilityState === 'visible' && !rtdbConnected) attemptReconnect(false);
});

const STATUS_COLORS = { 'APPROVED':'#22D3A6','LOST':'#FF6B6B','NO RESPONSE':'#8C9AE8','UNDER REVIEW':'#FFC24B','N/D':'#4A5388' };
const STATUS_LABEL_PT = { 'APPROVED':'Aprovada','LOST':'Perdida','NO RESPONSE':'Sem resposta','UNDER REVIEW':'Em revisão','N/D':'Não definido' };
const PT_MONTHS = ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];
function monthLabel(m){
  if(!m){ return 'Sem data'; }
  const p = m.split('-');
  return PT_MONTHS[parseInt(p[1],10)-1] + '/' + p[0].slice(2);
}

/* ===== FASE 2: padronização de datas em ISO 8601 (AAAA-MM-DD) =====
   Todo o app grava e exibe datas nesse formato. Registros antigos da base
   seed ainda trazem "AAAA-MM-DD HH:MM:SS" (um resquício de quando os dados
   vieram de planilha); esta função normaliza qualquer valor de entrada para
   somente a parte de data, e retorna null se o valor não for uma data ISO
   válida (evita gravar/exibir lixo como "*" ou strings soltas). */
let months=[], clients=[], sellers=[], statuses=[], modals=[], types=[];
let configListas = {}; // Fase Extra · Etapa 3 — Camada 1 (declarada cedo: render() já lê configListas.clientesFrequentes)
let state = { search:'', month:'', client:'', seller:'', status:'', modal:'', type:'', sortKey:'week', sortDir:'desc' };

function uniqueSorted(arr){ return [...new Set(arr.filter(Boolean))].sort(); }

function populateSelect(id, values, labelFn, keepFirst){
  const sel = document.getElementById(id);
  const startAt = keepFirst ? 1 : 0;
  while(sel.options.length > startAt) sel.remove(startAt);
  values.forEach(v=>{
    const opt = document.createElement('option');
    opt.value = v; opt.textContent = labelFn ? labelFn(v) : v;
    sel.appendChild(opt);
  });
}

function recomputeMeta(){
  months = uniqueSorted(RAW.map(d=>d.month));
  clients = uniqueSorted(RAW.map(d=>d.client).concat(configListas.clientesFrequentes || []));
  sellers = uniqueSorted(RAW.map(d=>d.seller));
  statuses = uniqueSorted(RAW.map(d=>d.status));
  modals = uniqueSorted(RAW.map(d=>d.modal));
  types = uniqueSorted(RAW.map(d=>d.type));

  populateSelect('fMonth', months, monthLabel, true);
  populateSelect('fClient', clients, null, true);
  populateSelect('fSeller', sellers, null, true);
  populateSelect('fStatus', statuses, s=>STATUS_LABEL_PT[s]||s, true);
  populateSelect('fModal', modals, null, true);
  populateSelect('fType', types, null, true);

  ['fMonth','fClient','fSeller','fStatus','fModal','fType'].forEach(id=>{
    const map = {fMonth:'month',fClient:'client',fSeller:'seller',fStatus:'status',fModal:'modal',fType:'type'};
    const key = map[id]; const sel = document.getElementById(id);
    if([...sel.options].some(o=>o.value===state[key])) sel.value = state[key]; else { sel.value=''; state[key]=''; }
  });

  const dl = document.getElementById('clientsList');
  dl.innerHTML = clients.map(c=>`<option value="${window.escapeHtml(c)}">`).join('');

  rebuildChips();
  rebuildCompareSelects();
  document.getElementById('pillTotal').textContent = RAW.length;
}

function rebuildChips(){
  const chipRow = document.getElementById('statusChips');
  chipRow.innerHTML='';
  statuses.forEach(s=>{
    const chip = document.createElement('div');
    chip.className='chip'; chip.dataset.status = s;
    chip.textContent = (STATUS_LABEL_PT[s]||s) + ' · ' + RAW.filter(d=>d.status===s).length;
    chip.addEventListener('click', ()=>{
      state.status = (state.status === s) ? '' : s;
      document.getElementById('fStatus').value = state.status;
      render();
    });
    chipRow.appendChild(chip);
  });
  [...chipRow.children].forEach(c=>c.classList.toggle('active', c.dataset.status===state.status));
}

/* ============================================================
   COMPARATIVO GENÉRICO POR PERÍODO (Especificação de melhorias, item 5.2)
   ============================================================
   compareGranularity é o estado único que decide o que os seletores A/B
   mostram — trocar de granularidade é só recalcular o array de chaves
   disponíveis (getPeriodKey/periodLabel, utils.js) e repopular os
   selects, sem duplicar lógica por período. Estado inicial: "Mês" já
   era o padrão; ao adotar Ano/Mês como default pedido na especificação,
   isso agora é só um valor inicial da variável, não um caminho de
   código diferente. */
let compareGranularity = 'month';

function availablePeriodKeys(granularity){
  const keys = new Set();
  RAW.forEach(d => { const k = getPeriodKey(d.week, granularity); if(k) keys.add(k); });
  return [...keys].sort();
}

function rebuildCompareSelects(){
  const a = document.getElementById('cMonthA'), b = document.getElementById('cMonthB');
  const prevA = a.value, prevB = b.value;
  const keys = availablePeriodKeys(compareGranularity);
  const optsHtml = keys.map(k=>`<option value="${k}">${periodLabel(k, compareGranularity)}</option>`).join('');
  a.innerHTML = optsHtml;
  b.innerHTML = optsHtml;
  if(keys.length){
    // Estado inicial padrão (Especificação de melhorias, item 5.2): ao
    // trocar de granularidade sem seleção anterior válida, compara os
    // dois últimos períodos disponíveis — com granularidade "month" (o
    // padrão ao carregar a tela), isso reproduz "comparando os meses
    // dentro do ano corrente" pedido como estado inicial.
    a.value = keys.includes(prevA) ? prevA : keys[Math.max(0,keys.length-2)];
    b.value = keys.includes(prevB) ? prevB : keys[keys.length-1];
  }
}

function getFiltered(){
  // Fase 3 (etapa 2): a lógica de filtro em si virou uma função pura e
  // testada (filterQuotations, em utils.js) — aqui só repassamos o estado
  // atual do módulo (RAW/state) pra ela.
  return filterQuotations(RAW, state);
}
// Fase 5 (Redesign Premium) · Etapa 2, revisão — corrige o bug de "nenhuma
// rota para exibir" mesmo com o painel filtrado mostrando resultados.
// Exposta diretamente em window (não via callback registrado, que teria
// o mesmo problema de timing que está sendo corrigido: se cotacoes.js
// tentasse "empurrar" o registro antes de route-map.js existir, o
// registro se perderia). Assim, o widget de mapa pode ler
// window.getFilteredQuotations a qualquer momento, sempre que for
// expandido — nunca depende de quem carregou primeiro.
window.getFilteredQuotations = getFiltered;

// Especificação de melhorias, item 9 — fecha o ciclo entre "explorar no
// mapa" e "agir na lista": clicar em "ver todas na lista" (route-map.js,
// modo location) aplica o código do ponto como filtro de busca (mesmo
// campo #fSearch que já filtra por origem/destino em filterQuotations,
// utils.js), navega para a tabela e fecha o mapa. Exposta diretamente em
// window pelo mesmo motivo de getFilteredQuotations acima: route-map.js
// não deve depender de qual dos dois arquivos terminou de carregar primeiro.
window.applyRouteFilterAndGoToTable = function(code){
  if(!code) return;
  state.search = code;
  const searchInput = document.getElementById('fSearch');
  if(searchInput) searchInput.value = code;
  render();
  goToView('table');
  window.routeMapCollapse?.();
};

/* ===== ICONS drawn on canvas (plugin) ===== */
function drawPlaneIcon(ctx,cx,cy,s,color){
  ctx.save(); ctx.translate(cx,cy); ctx.rotate(-Math.PI/4);
  ctx.beginPath(); ctx.moveTo(0,-s); ctx.lineTo(s*0.62,s*0.6); ctx.lineTo(0,s*0.22); ctx.lineTo(-s*0.62,s*0.6); ctx.closePath();
  ctx.fillStyle=color; ctx.fill(); ctx.restore();
}
function drawShipIcon(ctx,cx,cy,s,color){
  ctx.save(); ctx.translate(cx,cy);
  ctx.beginPath(); ctx.moveTo(-s,s*0.3); ctx.lineTo(s,s*0.3); ctx.lineTo(s*0.62,s*0.9); ctx.lineTo(-s*0.62,s*0.9); ctx.closePath();
  ctx.fillStyle=color; ctx.fill();
  ctx.fillRect(-s*0.22,-s*0.15,s*0.44,s*0.45);
  ctx.beginPath(); ctx.moveTo(0,-s*0.15); ctx.lineTo(0,-s*0.95); ctx.strokeStyle=color; ctx.lineWidth=1.6; ctx.stroke();
  ctx.restore();
}
function drawTruckIcon(ctx,cx,cy,s,color){
  ctx.save(); ctx.translate(cx,cy);
  ctx.fillStyle=color;
  ctx.fillRect(-s,-s*0.45,s*1.05,s*0.65);
  ctx.beginPath(); ctx.moveTo(s*0.08,-s*0.1); ctx.lineTo(s*0.72,-s*0.1); ctx.lineTo(s*0.92,s*0.22); ctx.lineTo(s*0.08,s*0.22); ctx.closePath(); ctx.fill();
  ctx.beginPath(); ctx.arc(-s*0.5,s*0.28,s*0.16,0,Math.PI*2); ctx.fill();
  ctx.beginPath(); ctx.arc(s*0.5,s*0.28,s*0.16,0,Math.PI*2); ctx.fill();
  ctx.restore();
}
function drawModalIcon(ctx,label,cx,cy,s,color){
  const m = (label||'').toUpperCase();
  if(m.includes('AIR')) drawPlaneIcon(ctx,cx,cy,s,color);
  else if(m.includes('OCEAN') || m.includes('SEA')) drawShipIcon(ctx,cx,cy,s,color);
  else if(m.includes('TRUCK')) drawTruckIcon(ctx,cx,cy,s,color);
  else { ctx.beginPath(); ctx.arc(cx,cy,s*0.4,0,Math.PI*2); ctx.fillStyle=color; ctx.fill(); }
}
const modalIconPlugin = {
  id:'modalIcons',
  afterDatasetsDraw(chart){
    const meta = chart.getDatasetMeta(0);
    const {ctx} = chart;
    meta.data.forEach((bar,i)=>{
      const label = chart.data.labels[i];
      const color = chart.data.datasets[0].backgroundColor[i] || '#00AEEF';
      drawModalIcon(ctx, label, bar.x, bar.y-16, 9, color);
    });
  }
};

/* ===== FASE 4 · ETAPA 2 (UX/UI) — gradiente sutil para os gráficos de
   Desempenho e Faturamento, mantendo Chart.js (só o preenchimento muda
   de cor sólida para gradiente vertical na mesma paleta já em uso). ===== */
function hexToRgba(hex, alpha){
  const h = hex.replace('#','');
  const r = parseInt(h.substring(0,2),16), g = parseInt(h.substring(2,4),16), b = parseInt(h.substring(4,6),16);
  return `rgba(${r},${g},${b},${alpha})`;
}
/* O 5º parâmetro (Auditoria Geral 01/08, Bloco 5) escolhe a direção do
   degradê. Num gráfico de barras o degradê precisa correr no MESMO eixo em
   que a barra cresce: numa barra vertical, de cima para baixo; numa barra
   horizontal (indexAxis:'y'), da esquerda para a direita. Aplicar o degradê
   vertical numa barra horizontal faz todas as barras terminarem na mesma
   tonalidade, independentemente do valor — que é justamente o que dá a
   aparência "chapada" reclamada no gráfico de desempenho. */
function gradientFill(canvasId, colorHex, opacityTop, opacityBottom, direcao = 'vertical'){
  const canvas = document.getElementById(canvasId);
  if(!canvas) return colorHex;
  const ctx = canvas.getContext('2d');
  const g = direcao === 'horizontal'
    ? ctx.createLinearGradient(0, 0, canvas.width || 480, 0)
    : ctx.createLinearGradient(0, 0, 0, canvas.height || 280);
  g.addColorStop(0, hexToRgba(colorHex, opacityTop ?? 0.92));
  g.addColorStop(1, hexToRgba(colorHex, opacityBottom ?? 0.22));
  return g;
}

/* ===== CHART SETUP ===== */
Chart.defaults.font.family = "'Inter', sans-serif";
Chart.defaults.color = '#B7C0E8';
function baseGridOpts(){
  return {
    x:{ grid:{color:'rgba(140,158,235,0.08)'}, ticks:{color:'#8C97D4', font:{size:11}} },
    y:{ grid:{color:'rgba(140,158,235,0.08)'}, ticks:{color:'#8C97D4', font:{size:11}}, beginAtZero:true }
  };
}
let chartMonth, chartStatus, chartSeller, chartSellerRate, chartClients, chartModal, chartType, chartCompare;
let chartExecEvolution, chartExecShare; // Versão de demonstração executiva (Fase 5)
let chartFatxEvol; // Versão de demonstração executiva (Fase 5) — Faturamento

function buildCharts(){
  chartMonth = new Chart(document.getElementById('chartMonth'), {
    type:'line', data:{labels:[], datasets:[{label:'Cotações', data:[], borderColor:'#00AEEF', backgroundColor:'rgba(0,174,239,0.18)', fill:true, tension:0.35, pointRadius:4, pointBackgroundColor:'#00AEEF', pointBorderColor:'#0A1440', pointBorderWidth:2, borderWidth:2.5}]},
    options:{responsive:true, maintainAspectRatio:false, plugins:{legend:{display:false}}, scales:baseGridOpts()}
  });
  chartStatus = new Chart(document.getElementById('chartStatus'), {
    type:'doughnut', data:{labels:[], datasets:[{data:[], backgroundColor:[], borderColor:'#0E1A4E', borderWidth:3}]},
    options:{responsive:true, maintainAspectRatio:false, cutout:'68%', plugins:{legend:{position:'bottom', labels:{boxWidth:10,boxHeight:10,padding:14,font:{size:11.5}}}}}
  });
  /* Auditoria Geral 01/08 (Bloco 5, seção 6.2) — "gráfico amontoado".
     Era uma barra VERTICAL EMPILHADA (stacked nos dois eixos), com os
     vendedores no eixo X. Com dois vendedores na base, os seis status
     ficavam comprimidos em duas colunas altas: não era problema de cor nem
     de degradê, era o formato não comportar os dados.

     Agora é barra HORIZONTAL AGRUPADA (indexAxis:'y', sem empilhamento) —
     cada status ganha barra própria e os rótulos ficam legíveis. É a mesma
     configuração do chartSellerRate logo ao lado, que a auditoria registra
     como a que já funciona bem nesta view. */
  chartSeller = new Chart(document.getElementById('chartSeller'), {
    type:'bar', data:{labels:[], datasets:[]},
    options:{
      responsive:true, maintainAspectRatio:false, indexAxis:'y',
      plugins:{legend:{position:'bottom', labels:{boxWidth:10,boxHeight:10,padding:14,font:{size:11.5}}}},
      scales:{
        x:{...baseGridOpts().x, stacked:false, beginAtZero:true, ticks:{...baseGridOpts().x.ticks, precision:0}},
        y:{...baseGridOpts().y, stacked:false}
      }
    }
  });
  chartSellerRate = new Chart(document.getElementById('chartSellerRate'), {
    type:'bar', data:{labels:[], datasets:[{label:'Taxa de aprovação (%)', data:[], backgroundColor:['#00AEEF','#2643D6'], borderRadius:8, barThickness:60}]},
    options:{responsive:true, maintainAspectRatio:false, indexAxis:'y', plugins:{legend:{display:false}}, scales:{x:{...baseGridOpts().x, max:100}, y:baseGridOpts().y}}
  });
  chartClients = new Chart(document.getElementById('chartClients'), {
    type:'bar', data:{labels:[], datasets:[{label:'Cotações', data:[], backgroundColor:'#2643D6', borderRadius:6, barThickness:14}]},
    options:{responsive:true, maintainAspectRatio:false, indexAxis:'y', plugins:{legend:{display:false}}, scales:{x:baseGridOpts().x, y:{...baseGridOpts().y, beginAtZero:true}}}
  });
  chartModal = new Chart(document.getElementById('chartModal'), {
    type:'bar', data:{labels:[], datasets:[{label:'Cotações', data:[], backgroundColor:[], borderRadius:6, barThickness:34}]},
    options:{responsive:true, maintainAspectRatio:false, layout:{padding:{top:22}}, plugins:{legend:{display:false}}, scales:{x:baseGridOpts().x, y:{...baseGridOpts().y, beginAtZero:true}}},
    plugins:[modalIconPlugin]
  });
  chartType = new Chart(document.getElementById('chartType'), {
    type:'doughnut', data:{labels:[], datasets:[{data:[], backgroundColor:['#1428A0','#00AEEF','#8C9AE8'], borderColor:'#0E1A4E', borderWidth:3}]},
    options:{responsive:true, maintainAspectRatio:false, cutout:'60%', plugins:{legend:{position:'bottom', labels:{boxWidth:10,boxHeight:10,font:{size:11.5}}}}}
  });
  chartCompare = new Chart(document.getElementById('chartCompare'), {
    type:'bar', data:{labels:[], datasets:[]},
    options:{responsive:true, maintainAspectRatio:false, plugins:{legend:{position:'bottom', labels:{boxWidth:10,boxHeight:10,font:{size:11.5}}}}, scales:baseGridOpts()}
  });
  // chartFaturamento saiu na Auditoria Geral de 01/08 (Bloco 2), junto com
  // renderFaturamento() e o canvas oculto que o hospedava: era um Chart.js
  // instanciado a cada carregamento para um gráfico que nenhuma view exibia.
  // Versão de demonstração executiva (Fase 5) — dashboard financeiro
  // baseado nos três valores cadastrados manualmente (ALL IN, Total
  // Collect, Handling Fee). Mesmo padrão de configuração dos demais
  // gráficos Chart.js do arquivo — nenhuma biblioteca nova.
  chartExecEvolution = new Chart(document.getElementById('chartExecEvolution'), {
    type:'line',
    data:{labels:[], datasets:[
      {label:'ALL IN', data:[], borderColor:'#00AEEF', backgroundColor:gradientFill('chartExecEvolution','#00AEEF',0.32,0.02), fill:true, tension:0.35, pointRadius:3, borderWidth:2.5},
      {label:'Total Collect', data:[], borderColor:'#8C9AE8', backgroundColor:'rgba(140,154,232,0.06)', fill:true, tension:0.35, pointRadius:3, borderWidth:2},
      {label:'Handling Fee', data:[], borderColor:'#22D3A6', backgroundColor:'rgba(34,211,166,0.06)', fill:true, tension:0.35, pointRadius:3, borderWidth:2},
    ]},
    options:{responsive:true, maintainAspectRatio:false, plugins:{legend:{position:'bottom', labels:{boxWidth:10,boxHeight:10,padding:14,font:{size:11.5}}}}, scales:baseGridOpts()}
  });
  chartExecShare = new Chart(document.getElementById('chartExecShare'), {
    type:'doughnut',
    data:{labels:['ALL IN','Total Collect','Handling Fee'], datasets:[{data:[0,0,0], backgroundColor:['#00AEEF','#8C9AE8','#22D3A6'], borderColor:'#0E1A4E', borderWidth:3}]},
    options:{responsive:true, maintainAspectRatio:false, cutout:'68%', plugins:{legend:{position:'bottom', labels:{boxWidth:10,boxHeight:10,padding:14,font:{size:11.5}}}}}
  });
  // Versão de demonstração executiva (Fase 5) — view Faturamento.
  // O gráfico de pizza chartFatxDist que existia aqui foi removido na
  // Auditoria Geral de 01/08 (Bloco 2): duplicava exatamente os três
  // valores já mostrados pelo chartExecShare do Dashboard.
  chartFatxEvol = new Chart(document.getElementById('chartFatxEvol'), {
    type:'bar',
    data:{labels:[], datasets:[
      {label:'ALL IN', data:[], backgroundColor:'#00AEEF', borderRadius:6},
      {label:'Total Collect', data:[], backgroundColor:'#8C9AE8', borderRadius:6},
      {label:'Handling Fee', data:[], backgroundColor:'#22D3A6', borderRadius:6},
    ]},
    options:{responsive:true, maintainAspectRatio:false, plugins:{legend:{position:'bottom', labels:{boxWidth:10,boxHeight:10,font:{size:11.5}}}}, scales:baseGridOpts()}
  });
}

function statusClass(s){ return (s||'N/D').replace(/\s+/g,'-'); }

// Fase 4 · Etapa 2 (UX/UI) — desenha a mini-tendência de um KPI: escreve os
// pontos calculados por sparklinePoints() (utils.js) na <polyline> do card.
// Só formata/escreve no DOM, mesma divisão de responsabilidade do resto do
// arquivo (cálculo puro fora, DOM aqui).
function setKpiSparkline(id, series){
  const el = document.getElementById(id);
  if(!el) return;
  el.setAttribute('points', sparklinePoints(series, 64, 22));
  /* Auditoria Geral 01/08 (Bloco 5) — marca o ÚLTIMO ponto da série com um
     círculo. Sem ele, a sparkline é uma linha sem começo nem fim declarados e
     o olho não sabe qual ponta é "agora". O círculo fica num <circle> irmão
     da polyline, com id previsível (mesmo id + 'Dot'). */
  const dot = document.getElementById(id + 'Dot');
  if(dot){
    const pts = sparklinePoints(series, 64, 22).trim().split(/\s+/);
    const ultimo = pts[pts.length - 1];
    if(ultimo && ultimo.includes(',')){
      const [x, y] = ultimo.split(',');
      /* O ponto é um <line> de comprimento ZERO com stroke-linecap="round",
         não um <circle>. Motivo: a sparkline usa preserveAspectRatio="none"
         num viewBox 64x22 esticado até a largura do card, o que deformaria
         qualquer círculo numa elipse achatada, com achatamento variando
         conforme a largura da tela. Um traço de comprimento zero com ponta
         redonda e vector-effect="non-scaling-stroke" ignora a escala do
         viewBox e sai sempre redondo, em qualquer largura. */
      dot.setAttribute('x1', x); dot.setAttribute('x2', x);
      dot.setAttribute('y1', y); dot.setAttribute('y2', y);
      dot.style.display = '';
    } else {
      dot.style.display = 'none';
    }
  }
}

/* Auditoria Geral 01/08 (Bloco 5) — indicador de tendência do card de KPI.
   A sparkline mostra o formato do movimento, mas não o tamanho dele: dois
   cards com curvas parecidas podiam representar +3% e +40%. Esta função
   compara o último período fechado com o anterior e escreve a variação num
   selo ao lado do valor.

   Compara os DOIS últimos pontos da mesma série já usada na sparkline, então
   não há fonte de dado nova nem risco de divergir do gráfico ao lado.
   Sem série suficiente, ou com base zero (variação percentual indefinida),
   o selo simplesmente não aparece — nunca mostra "Infinity%" nem "NaN". */
function setKpiTrend(id, series, maiorEhMelhor = true){
  const el = document.getElementById(id);
  if(!el) return;
  const limpa = (series || []).filter(v => typeof v === 'number' && isFinite(v));
  if(limpa.length < 2){ el.style.display = 'none'; return; }
  const atual = limpa[limpa.length - 1];
  const anterior = limpa[limpa.length - 2];
  if(anterior === 0){ el.style.display = 'none'; return; }
  const variacao = ((atual - anterior) / Math.abs(anterior)) * 100;
  if(!isFinite(variacao)){ el.style.display = 'none'; return; }
  const subiu = variacao > 0;
  const estavel = Math.abs(variacao) < 0.05;
  el.style.display = '';
  el.textContent = (estavel ? '→ ' : (subiu ? '▲ ' : '▼ ')) +
    Math.abs(variacao).toFixed(Math.abs(variacao) >= 10 ? 0 : 1) + '%';
  /* "Bom" e "ruim" dependem do indicador: crescer em Aprovadas é bom, crescer
     em Perdidas é ruim. Por isso maiorEhMelhor é parâmetro, não suposição. */
  el.className = 'dash-kpi-trend' + (estavel ? ' is-flat' : (subiu === maiorEhMelhor ? ' is-good' : ' is-bad'));
  el.title = 'Comparado ao período anterior da mesma série mostrada abaixo';
}

/* ============================================================
   VERSÃO DE DEMONSTRAÇÃO EXECUTIVA — Fase 5, ajuste temporário
   ============================================================
   Dashboard executivo baseado exclusivamente nos três valores cadastrados
   manualmente no modal de Editar Cotação (d.allIn, d.totalCollect,
   d.handling — já lidos/gravados por readCommercialFields/fillCommercial
   Fields, ver essas funções mais abaixo neste arquivo).

   Correção (Especificação de melhorias, item 2.2): "Receita Total"
   (definida como ALL IN − Total Collect) foi removida deste dashboard —
   mantém-se apenas All-in, Handling Fee e Total Collect, como pedido.

   Todos os valores numéricos ausentes (cotação sem nenhum dos três campos
   preenchidos ainda) contam como 0 nos agregados, nunca quebram o cálculo
   nem geram NaN — mesmo princípio de robustez já usado em computeKpis()
   (utils.js) para arrays vazios. */
function renderExecutiveDashboard(data){
  const n = data.length;
  let sumAllIn = 0, sumTotalCollect = 0, sumHandling = 0;
  const porMes = {}; // 'YYYY-MM' -> { allIn, totalCollect, handling }

  data.forEach(d=>{
    const allIn = Number(d.allIn) || 0;
    const totalCollect = Number(d.totalCollect) || 0;
    const handling = Number(d.handling) || 0;
    sumAllIn += allIn;
    sumTotalCollect += totalCollect;
    sumHandling += handling;

    if(d.month){
      if(!porMes[d.month]) porMes[d.month] = { allIn:0, totalCollect:0, handling:0 };
      porMes[d.month].allIn += allIn;
      porMes[d.month].totalCollect += totalCollect;
      porMes[d.month].handling += handling;
    }
  });

  const fmt = (v) => 'USD ' + v.toLocaleString('en-US', { minimumFractionDigits:0, maximumFractionDigits:0 });
  const pct = (part, whole) => whole > 0 ? (part/whole*100) : 0;

  document.getElementById('execAllInValue').textContent = fmt(sumAllIn);
  document.getElementById('execTotalCollectValue').textContent = fmt(sumTotalCollect);
  document.getElementById('execHandlingValue').textContent = fmt(sumHandling);

  const allInSharePct = pct(sumAllIn, sumAllIn + sumTotalCollect + sumHandling);
  const totalCollectShareOfAllIn = pct(sumTotalCollect, sumAllIn);
  const handlingShareOfAllIn = pct(sumHandling, sumAllIn);
  document.getElementById('execAllInShare').textContent = allInSharePct.toFixed(1)+'% do total combinado';
  document.getElementById('execTotalCollectShare').textContent = totalCollectShareOfAllIn.toFixed(1)+'% do ALL IN';
  document.getElementById('execHandlingShare').textContent = handlingShareOfAllIn.toFixed(1)+'% do ALL IN';

  // barras de progresso: cada uma relativa ao maior valor entre os três,
  // para que a barra mais alta sempre chegue a 100% (leitura visual rápida)
  const maxVal = Math.max(sumAllIn, sumTotalCollect, sumHandling, 1);
  document.getElementById('execAllInBar').style.width = Math.min(100, pct(sumAllIn, maxVal)) + '%';
  document.getElementById('execTotalCollectBar').style.width = Math.min(100, pct(sumTotalCollect, maxVal)) + '%';
  document.getElementById('execHandlingBar').style.width = Math.min(100, pct(sumHandling, maxVal)) + '%';

  // crescimento/variação: compara os dois últimos meses com dados
  // (mesmo padrão de "sem dado suficiente = —" já usado em kpiApprovedSub)
  const mesesComDado = Object.keys(porMes).sort();
  function trendBadge(elId, key){
    const el = document.getElementById(elId);
    if(mesesComDado.length < 2){ el.textContent='—'; el.className='exec-kpi-badge'; return; }
    const atual = porMes[mesesComDado[mesesComDado.length-1]][key] || 0;
    const anterior = porMes[mesesComDado[mesesComDado.length-2]][key] || 0;
    if(anterior === 0){ el.textContent='—'; el.className='exec-kpi-badge'; return; }
    const variacao = ((atual-anterior)/anterior)*100;
    const up = variacao >= 0;
    el.textContent = (up?'▲ ':'▼ ') + Math.abs(variacao).toFixed(1) + '%';
    el.className = 'exec-kpi-badge ' + (up ? 'is-up' : 'is-down');
  }
  trendBadge('execAllInTrend', 'allIn');
  trendBadge('execTotalCollectTrend', 'totalCollect');
  trendBadge('execHandlingTrend', 'handling');

  // gráfico de evolução mensal (usa a lista global "months", já ordenada,
  // mesma fonte usada pelos demais gráficos deste arquivo)
  if(chartExecEvolution){
    chartExecEvolution.data.labels = months.map(monthLabel);
    chartExecEvolution.data.datasets[0].data = months.map(m => porMes[m] ? porMes[m].allIn : 0);
    chartExecEvolution.data.datasets[1].data = months.map(m => porMes[m] ? porMes[m].totalCollect : 0);
    chartExecEvolution.data.datasets[2].data = months.map(m => porMes[m] ? porMes[m].handling : 0);
    chartExecEvolution.update();
  }
  if(chartExecShare){
    chartExecShare.data.datasets[0].data = [sumAllIn, sumTotalCollect, sumHandling];
    chartExecShare.update();
  }

  return { sumAllIn, sumTotalCollect, sumHandling, n };
}

/* ============================================================
   VERSÃO DE DEMONSTRAÇÃO EXECUTIVA — Fase 5, ajuste temporário
   ============================================================
   View "Faturamento executivo": usa TODO o histórico cadastrado (RAW),
   não apenas o subconjunto filtrado pelo painel — decisão deliberada
   para que esta tela sirva como visão financeira ampla na apresentação,
   com filtro de período PRÓPRIO (fatxPeriodFilter), independente dos
   filtros globais da tela Geral. Mesma fonte de dados (d.allIn,
   d.totalCollect, d.handling) usada em renderExecutiveDashboard().

   Correção (Especificação de melhorias, item 2.2): "Receita total"
   (ALL IN − Total Collect) foi removida. Os KPIs que antes usavam esse
   valor como base (Ticket médio, Crescimento mensal, Média por período)
   agora usam ALL IN puro como base — decisão de produto registrada
   aqui a pedido do usuário, para não deixar esses três indicadores sem
   nenhuma base de cálculo. */
function filtrarPorPeriodo(rows, periodo){
  if(periodo === 'all') return rows;
  const dias = parseInt(periodo, 10);
  const limite = new Date();
  limite.setDate(limite.getDate() - dias);
  const limiteISO = toISODate(limite.toISOString().slice(0,10));
  return rows.filter(d => d.week && d.week >= limiteISO);
}

function renderFaturamentoExecutivo(){
  const grid = document.getElementById('fatxSummaryGrid');
  if(!grid) return; // segurança: view ainda não presente nesta versão do HTML

  const periodoEl = document.getElementById('fatxPeriodFilter');
  const periodo = periodoEl ? periodoEl.value : 'all';
  const todasComData = RAW.filter(d => d.week); // histórico completo, não o filtro do painel
  const rows = filtrarPorPeriodo(todasComData, periodo);

  let sumAllIn = 0, sumTotalCollect = 0, sumHandling = 0;
  const porMes = {};
  rows.forEach(d=>{
    const allIn = Number(d.allIn) || 0;
    const totalCollect = Number(d.totalCollect) || 0;
    const handling = Number(d.handling) || 0;
    sumAllIn += allIn; sumTotalCollect += totalCollect; sumHandling += handling;
    if(d.month){
      if(!porMes[d.month]) porMes[d.month] = { allIn:0, totalCollect:0, handling:0, count:0 };
      porMes[d.month].allIn += allIn; porMes[d.month].totalCollect += totalCollect; porMes[d.month].handling += handling;
      porMes[d.month].count++;
    }
  });

  const fmt = (v) => 'USD ' + v.toLocaleString('en-US', { minimumFractionDigits:0, maximumFractionDigits:0 });

  document.getElementById('fatxAllIn').textContent = fmt(sumAllIn);
  document.getElementById('fatxTotalCollect').textContent = fmt(sumTotalCollect);
  document.getElementById('fatxHandling').textContent = fmt(sumHandling);

  const mesesOrdenados = Object.keys(porMes).sort();
  if(chartFatxEvol){
    chartFatxEvol.data.labels = mesesOrdenados.map(monthLabel);
    chartFatxEvol.data.datasets[0].data = mesesOrdenados.map(m=>+porMes[m].allIn.toFixed(2));
    chartFatxEvol.data.datasets[1].data = mesesOrdenados.map(m=>+porMes[m].totalCollect.toFixed(2));
    chartFatxEvol.data.datasets[2].data = mesesOrdenados.map(m=>+porMes[m].handling.toFixed(2));
    chartFatxEvol.update();
  }

  // KPIs — base de cálculo trocada de "Receita" (ALL IN − Total Collect)
  // para ALL IN puro, já que o card de Receita foi removido (ver nota acima).
  const n = rows.length;
  const ticketMedio = n > 0 ? sumAllIn / n : 0;
  document.getElementById('fatxKpiTicket').textContent = fmt(ticketMedio);
  document.getElementById('fatxKpiCount').textContent = n;
  const growthEl = document.getElementById('fatxKpiGrowth');
  if(mesesOrdenados.length >= 2){
    const atual = porMes[mesesOrdenados[mesesOrdenados.length-1]].allIn;
    const anterior = porMes[mesesOrdenados[mesesOrdenados.length-2]].allIn;
    if(anterior !== 0){
      const variacao = ((atual-anterior)/Math.abs(anterior))*100;
      growthEl.textContent = (variacao>=0?'▲ ':'▼ ') + Math.abs(variacao).toFixed(1) + '%';
      growthEl.style.color = variacao>=0 ? '#22D3A6' : '#FF6B6B';
    } else { growthEl.textContent = '—'; growthEl.style.color = ''; }
  } else { growthEl.textContent = '—'; growthEl.style.color = ''; }
  const mediaPeriodo = mesesOrdenados.length > 0 ? sumAllIn / mesesOrdenados.length : 0;
  document.getElementById('fatxKpiAvgPeriod').textContent = fmt(mediaPeriodo);

  // histórico (tabela)
  const tbody = document.getElementById('fatxTbody');
  if(tbody){
    const ordenadas = [...rows].sort((a,b)=> (b.week||'').localeCompare(a.week||''));
    tbody.innerHTML = ordenadas.length ? ordenadas.slice(0,200).map(d=>{
      const allIn = Number(d.allIn)||0, totalCollect = Number(d.totalCollect)||0, handling = Number(d.handling)||0;
      const total = allIn + totalCollect + handling;
      return `<tr>
        <td class="mono">${window.escapeHtml(toISODate(d.week)||'—')}</td>
        <td>${window.escapeHtml(d.ref||'—')}</td>
        <td>${previewFmtMoney(allIn)}</td>
        <td>${previewFmtMoney(totalCollect)}</td>
        <td>${previewFmtMoney(handling)}</td>
        <td><b>${previewFmtMoney(total)}</b></td>
      </tr>`;
    }).join('') : `<tr><td colspan="6">${emptyStateHtml({ icon:'🧾', title:'Nenhum registro no período selecionado', sub:'Cadastre valores de ALL IN, Total Collect e Handling Fee ao editar uma cotação.', compact:true })}</td></tr>`;
  }
}

/* ============================================================
   TABELA DE DESEMPENHO POR VENDEDOR (Especificação de melhorias, item 5.1)
   ============================================================
   Complementar aos gráficos chartSeller/chartSellerRate — mesmos dados
   (computeSellerStats, utils.js), exibidos como tabela ordenável por
   coluna. O estado de ordenação (sellerStatsSortKey/Dir) persiste entre
   chamadas de render() (mesmo padrão já usado pela tabela principal,
   ver objeto `state`), para reordenar não perder a escolha do usuário
   a cada atualização de filtro. */
let sellerStatsSortKey = 'approvedRate'; // computeSellerStats já ordena por isso por padrão
let sellerStatsSortDir = 'desc';
let sellerStatsCache = []; // guarda o último cálculo para reordenar sem recalcular ao clicar no cabeçalho

function renderSellerStatsTable(data){
  const tbody = document.getElementById('sellerStatsTbody');
  if(!tbody) return; // segurança: view ainda não presente nesta versão do HTML
  sellerStatsCache = computeSellerStats(data);
  paintSellerStatsTable();
}

function paintSellerStatsTable(){
  const tbody = document.getElementById('sellerStatsTbody');
  if(!tbody) return;

  const sorted = [...sellerStatsCache].sort((a,b)=>{
    let av = a[sellerStatsSortKey], bv = b[sellerStatsSortKey];
    if(av === null || av === undefined) av = sellerStatsSortKey === 'seller' ? '' : -1;
    if(bv === null || bv === undefined) bv = sellerStatsSortKey === 'seller' ? '' : -1;
    if(av < bv) return sellerStatsSortDir==='asc' ? -1 : 1;
    if(av > bv) return sellerStatsSortDir==='asc' ? 1 : -1;
    return 0;
  });

  if(sorted.length === 0){
    tbody.innerHTML = `<tr><td colspan="6">${emptyStateHtml({ icon:'📈', title:'Nenhum registro no período selecionado', sub:'Ajuste os filtros do painel para ver o desempenho por vendedor.', compact:true })}</td></tr>`;
    return;
  }

  tbody.innerHTML = sorted.map(s=>{
    const rate = s.approvedRate !== null ? s.approvedRate : 0;
    return `
    <tr>
      <td><span class="seller-tag">${window.escapeHtml(s.seller)}</span></td>
      <td>${s.total}</td>
      <td><span class="stamp APPROVED">${s.approved}</span></td>
      <td><span class="stamp LOST">${s.lost}</span></td>
      <td><span class="stamp UNDER-REVIEW">${s.review}</span></td>
      <td>
        <div class="seller-rate-cell">
          <div class="seller-rate-bar"><i style="width:${Math.max(0,Math.min(100,rate))}%"></i></div>
          <span class="seller-rate-value">${s.approvedRate !== null ? s.approvedRate.toFixed(1)+'%' : '—'}</span>
        </div>
      </td>
    </tr>`;
  }).join('');
}

document.querySelectorAll('#sellerStatsTable thead th[data-seller-sort]').forEach(th=>{
  th.addEventListener('click', ()=>{
    const key = th.dataset.sellerSort;
    if(sellerStatsSortKey === key){
      sellerStatsSortDir = sellerStatsSortDir === 'asc' ? 'desc' : 'asc';
    } else {
      sellerStatsSortKey = key;
      sellerStatsSortDir = key === 'seller' ? 'asc' : 'desc'; // texto começa crescente, número começa decrescente (maior primeiro)
    }
    document.querySelectorAll('#sellerStatsTable thead th[data-seller-sort]').forEach(other=>{
      other.classList.toggle('sort-active', other === th);
      const arrow = other.querySelector('.sort-arrow');
      if(arrow) arrow.remove();
    });
    const arrowSpan = document.createElement('span');
    arrowSpan.className = 'sort-arrow';
    arrowSpan.textContent = sellerStatsSortDir === 'asc' ? '▲' : '▼';
    th.appendChild(arrowSpan);
    paintSellerStatsTable();
  });
});

/* ============================================================
   TABELA DE RANKING DE CLIENTES (Especificação de melhorias, item 6)
   ============================================================
   Ao lado do gráfico de pizza reduzido (chartClients) — mesmos 10
   clientes, na mesma ordem, com os valores monetários que explicam a
   posição no ranking. Formatação de moeda simples (USD, sem casas
   decimais) — mesmo padrão já usado em renderExecutiveDashboard/
   renderFaturamentoExecutivo, para consistência visual entre views. */
function renderClientRankingTable(topClients){
  const tbody = document.getElementById('clientRankingTbody');
  if(!tbody) return; // segurança: view ainda não presente nesta versão do HTML
  const fmtUsd = (v) => 'USD ' + v.toLocaleString('en-US', { minimumFractionDigits:0, maximumFractionDigits:0 });

  if(topClients.length === 0){
    tbody.innerHTML = `<tr><td colspan="4">${emptyStateHtml({ icon:'🌐', title:'Nenhum cliente no período selecionado', sub:'Ajuste os filtros do painel para ver o ranking.', compact:true })}</td></tr>`;
    return;
  }

  tbody.innerHTML = topClients.map(c=>`
    <tr>
      <td>${window.escapeHtml(c.client)}</td>
      <td>${c.count}</td>
      <td>${fmtUsd(c.totalAllIn)}</td>
      <td>${fmtUsd(c.avgTicket)}</td>
    </tr>`).join('');
}

function render(){
  const data = getFiltered();
  // Fase 5 (Redesign Premium) · Etapa 2 — widget flutuante de mapa de rotas:
  // repassa o conjunto já filtrado para o modo "overview" do widget. Chamada
  // aditiva e segura — se route-map.js não tiver carregado por algum motivo,
  // window.routeMapSetOverview simplesmente não existe e o encadeamento
  // opcional (?.) evita qualquer erro, sem afetar o restante de render().
  window.routeMapSetOverview?.(data);
  // Versão de demonstração executiva (Fase 5) — dashboard financeiro
  // baseado em ALL IN / Total Collect / Handling Fee cadastrados
  // manualmente. Ver função para detalhes; roda sempre que render() roda
  // (mesmo ciclo de atualização de todo o resto da tela).
  renderExecutiveDashboard(data);
  // Fase 3 (etapa 2): cálculo dos KPIs extraído para função pura e testada
  // (computeKpis, em utils.js) — aqui só formata e escreve no DOM.
  // Especificação de melhorias, item 2.1 — RAW.length como segundo
  // argumento: computeKpis agora também devolve *PctTotal (percentual de
  // cada categoria em relação ao TOTAL GERAL, não ao filtro atual).
  const kpis = computeKpis(data, RAW.length);
  // "Clientes ativos" não tem *PctTotal vindo de computeKpis (o total
  // geral de clientes únicos exige RAW inteiro, fora do escopo de uma
  // função que recebe só o conjunto já filtrado) — calculado aqui,
  // mesmo padrão de "cálculo ao lado" já usado para clientCounts abaixo.
  const totalClientesGeral = new Set(RAW.map(d=>d.client).filter(Boolean)).size;
  const clientsPctTotal = totalClientesGeral > 0 ? (kpis.clientsInView / totalClientesGeral) * 100 : null;

  // Fase 6 (Refatoração UI/UX) — os cards de contagem (Cotações
  // filtradas/Aprovadas/Perdidas/Em análise/Clientes) voltaram a ser
  // exibidos na tela, lado a lado com os cards financeiros, a pedido
  // do usuário (resgate do dashboard antigo). setTextSafe() continua
  // protegendo contra elemento ausente por segurança (ex.: se uma
  // versão futura remover algum desses elementos de novo, render()
  // não quebra no meio da execução).
  function setTextSafe(id, value){ const el = document.getElementById(id); if(el) el.textContent = value; }
  // Especificação de melhorias, item 2.1 — cada card ganha o trio de
  // métricas pedido: valor absoluto (já era exibido antes), % do total
  // geral (novo) e valor relativo ao filtro (taxa/contagem, já existia
  // para Aprovadas/Perdidas; agora também para Filtradas/Em análise/
  // Clientes). Formato compacto separado por "·", mesmo separador visual
  // já usado em outros pontos do app (ex. "Base: 355 registros").
  function pctTxt(pct){ return pct !== null ? pct.toFixed(1)+'% do total' : '—'; }

  setTextSafe('kpiTotal', kpis.total);
  setTextSafe('kpiTotalSub', pctTxt(kpis.totalPctTotal) + ' · de ' + RAW.length + ' no total');
  setTextSafe('kpiApproved', kpis.approved);
  setTextSafe('kpiApprovedSub', pctTxt(kpis.approvedPctTotal) + (kpis.approvedRate !== null ? ' · '+kpis.approvedRate.toFixed(1)+'% de conversão' : ''));
  setTextSafe('kpiLost', kpis.lost);
  setTextSafe('kpiLostSub', pctTxt(kpis.lostPctTotal) + (kpis.lostRate !== null ? ' · '+kpis.lostRate.toFixed(1)+'% de perda' : ''));
  setTextSafe('kpiReview', kpis.review);
  setTextSafe('kpiReviewSub', pctTxt(kpis.reviewPctTotal) + ' · under review + sem resposta');
  setTextSafe('kpiClients', kpis.clientsInView);
  setTextSafe('kpiClientsSub', pctTxt(clientsPctTotal));
  setTextSafe('kpiTopClient', kpis.topClient ? 'top: '+kpis.topClient.name+' ('+kpis.topClient.count+')' : '—');

  // Usado logo abaixo pelo gráfico "Top 10 clientes" (chartClients). Fica de
  // fora do computeKpis() de propósito: computeKpis devolve só o resumo
  // (cliente destaque), não o mapa completo de contagens — não é
  // responsabilidade de uma função de KPI alimentar gráfico nenhum.
  const clientCounts = {};
  data.forEach(d=>{ if(d.client) clientCounts[d.client] = (clientCounts[d.client]||0)+1; });

  const monthCounts = {}; months.forEach(m=>monthCounts[m]=0);
  data.forEach(d=>{ if(d.month) monthCounts[d.month] = (monthCounts[d.month]||0)+1; });
  chartMonth.data.labels = months.map(monthLabel);
  chartMonth.data.datasets[0].data = months.map(m=>monthCounts[m]||0);
  chartMonth.update();

  // Fase 4 · Etapa 2 (UX/UI) — sparklines nos cards de KPI: tendência dos
  // últimos até 6 meses. Reaproveita o monthCounts acima (mesmos dados do
  // gráfico "Cotações por mês") em vez de carregar algo novo do Firebase.
  const sparkMonths = months.slice(-6);
  const monthApproved = {}, monthLost = {}, monthReview = {}, monthClientSets = {};
  sparkMonths.forEach(m=>{ monthApproved[m]=0; monthLost[m]=0; monthReview[m]=0; monthClientSets[m]=new Set(); });
  data.forEach(d=>{
    if(!d.month || !(d.month in monthApproved)) return;
    if(d.status==='APPROVED') monthApproved[d.month]++;
    else if(d.status==='LOST') monthLost[d.month]++;
    else if(d.status==='UNDER REVIEW' || d.status==='NO RESPONSE') monthReview[d.month]++;
    if(d.client) monthClientSets[d.month].add(d.client);
  });
  const serieTotal    = sparkMonths.map(m=>monthCounts[m]||0);
  const serieApproved = sparkMonths.map(m=>monthApproved[m]||0);
  const serieLost     = sparkMonths.map(m=>monthLost[m]||0);
  const serieReview   = sparkMonths.map(m=>monthReview[m]||0);
  const serieClients  = sparkMonths.map(m=>monthClientSets[m].size);
  setKpiSparkline('kpiTotalSpark', serieTotal);
  setKpiSparkline('kpiApprovedSpark', serieApproved);
  setKpiSparkline('kpiLostSpark', serieLost);
  setKpiSparkline('kpiReviewSpark', serieReview);
  setKpiSparkline('kpiClientsSpark', serieClients);
  /* Auditoria Geral 01/08 (Bloco 5) — selo de variação sobre a MESMA série da
     sparkline. O 2º argumento diz se crescer é bom: em Perdidas e Em análise,
     crescer é ruim, e o selo fica vermelho em vez de verde. */
  setKpiTrend('kpiTotalTrend', serieTotal, true);
  setKpiTrend('kpiApprovedTrend', serieApproved, true);
  setKpiTrend('kpiLostTrend', serieLost, false);
  setKpiTrend('kpiReviewTrend', serieReview, false);
  setKpiTrend('kpiClientsTrend', serieClients, true);

  const statusCounts = {}; statuses.forEach(s=>statusCounts[s]=0);
  data.forEach(d=>{ statusCounts[d.status] = (statusCounts[d.status]||0)+1; });
  chartStatus.data.labels = statuses.map(s=>STATUS_LABEL_PT[s]||s);
  chartStatus.data.datasets[0].data = statuses.map(s=>statusCounts[s]);
  chartStatus.data.datasets[0].backgroundColor = statuses.map(s=>STATUS_COLORS[s]||'#4A5388');
  chartStatus.update();

  const sellerStatusData = {};
  sellers.forEach(sel=>{ sellerStatusData[sel] = {}; statuses.forEach(s=>sellerStatusData[sel][s]=0); });
  data.forEach(d=>{ if(sellerStatusData[d.seller]) sellerStatusData[d.seller][d.status] = (sellerStatusData[d.seller][d.status]||0)+1; });
  chartSeller.data.labels = sellers;
  /* FASE 4 · ETAPA 2 (UX/UI) — gradiente sutil por status, no lugar da cor
     sólida. Direção 'horizontal' desde o Bloco 5 da Auditoria Geral de
     01/08: o gráfico virou barra horizontal (indexAxis:'y'), e o degradê
     precisa correr no mesmo eixo em que a barra cresce. */
  chartSeller.data.datasets = statuses.map(s=>({
    label: STATUS_LABEL_PT[s]||s,
    data: sellers.map(sel=>sellerStatusData[sel][s]||0),
    backgroundColor: gradientFill('chartSeller', STATUS_COLORS[s]||'#4A5388', 0.92, 0.35, 'horizontal'),
    borderRadius: 4,
    borderSkipped: false,
  }));
  chartSeller.update();

  const sellerRate = sellers.map(sel=>{
    const subset = data.filter(d=>d.seller===sel);
    const app = subset.filter(d=>d.status==='APPROVED').length;
    return subset.length ? +(app/subset.length*100).toFixed(1) : 0;
  });
  chartSellerRate.data.labels = sellers;
  chartSellerRate.data.datasets[0].data = sellerRate;
  chartSellerRate.update();

  // Especificação de melhorias, item 5.1 — tabela complementar aos 2
  // gráficos acima (não substitui nenhum dos dois), mesmos dados
  // (data, já filtrado), calculados por computeSellerStats (utils.js).
  renderSellerStatsTable(data);

  // Especificação de melhorias, item 6 — gráfico e tabela usam a MESMA
  // fonte de dados/ordenação (computeClientRanking, utils.js), para
  // nunca contarem histórias diferentes sobre a posição de cada cliente.
  const clientRanking = computeClientRanking(data);
  const topClients = clientRanking.slice(0, 10);
  chartClients.data.labels = topClients.map(c=>c.client);
  chartClients.data.datasets[0].data = topClients.map(c=>c.count);
  chartClients.update();
  renderClientRankingTable(topClients);

  const modalCounts = {}; data.forEach(d=>{ modalCounts[d.modal] = (modalCounts[d.modal]||0)+1; });
  const modalEntries = Object.entries(modalCounts).sort((a,b)=>b[1]-a[1]);
  const modalPalette = ['#00AEEF','#2643D6','#5FD4FF','#8C9AE8','#22D3A6'];
  chartModal.data.labels = modalEntries.map(e=>e[0]);
  chartModal.data.datasets[0].data = modalEntries.map(e=>e[1]);
  chartModal.data.datasets[0].backgroundColor = modalEntries.map((_,i)=>modalPalette[i % modalPalette.length]);
  chartModal.update();

  const typeCounts = {}; data.forEach(d=>{ typeCounts[d.type] = (typeCounts[d.type]||0)+1; });
  const typeEntries = Object.entries(typeCounts).sort((a,b)=>b[1]-a[1]);
  chartType.data.labels = typeEntries.map(e=>e[0]);
  chartType.data.datasets[0].data = typeEntries.map(e=>e[1]);
  chartType.update();

  renderTable(data);
  [...document.getElementById('statusChips').children].forEach(c=>c.classList.toggle('active', c.dataset.status===state.status));
  renderCompare();
  renderFaturamentoExecutivo(); // Versão de demonstração executiva (Fase 5) — usa RAW completo, não o filtro do painel
}

/* REMOVIDO (Auditoria Geral 01/08, Bloco 2) — a view "Faturamento" da Fase
   Extra (renderFaturamento + coletarLinhasFaturamento, ~99 linhas) foi
   substituída pela versão executiva da Fase 5 (renderFaturamentoExecutivo,
   acima), que é a que o HTML de hoje realmente exibe. O código antigo
   continuava sendo chamado a cada render, mas saía na primeira linha por
   early-return — o elemento #fatKpiGrid que ele procurava não existe mais
   no HTML desde a Fase 5. Era código morto com custo de leitura, na mesma
   categoria de dashboard.js. */

function renderTable(data){
  let sorted = [...data];
  const key = state.sortKey;
  sorted.sort((a,b)=>{
    let av, bv;
    if(key==='route'){ av=(a.origin||'')+'>'+(a.dest||''); bv=(b.origin||'')+'>'+(b.dest||''); }
    else { av = a[key] || ''; bv = b[key] || ''; }
    if(key==='week'){ av = toISODate(av) || '0000-00-00'; bv = toISODate(bv) || '0000-00-00'; }
    if(av < bv) return state.sortDir==='asc' ? -1 : 1;
    if(av > bv) return state.sortDir==='asc' ? 1 : -1;
    return 0;
  });
  const tbody = document.getElementById('tbody');
  tbody.innerHTML='';
  document.getElementById('tableCount').textContent = data.length + ' linhas';
  if(sorted.length===0){
    tbody.innerHTML = `<tr class="empty-row"><td colspan="12">${emptyStateHtml({
      icon:'🔍', title:'Nenhum registro encontrado',
      sub:'Ajuste os filtros acima ou limpe-os para ver todas as cotações.',
      actionHtml:'<button type="button" class="btn-ghost" id="btnEmptyLimparFiltros" style="padding:6px 12px;font-size:11.5px;border-radius:8px;cursor:pointer;">Limpar filtros</button>',
      compact:true
    })}</td></tr>`;
    document.getElementById('btnEmptyLimparFiltros')?.addEventListener('click', () => document.getElementById('btnReset')?.click());
    return;
  }
  const frag = document.createDocumentFragment();
  sorted.slice(0,400).forEach(d=>{
    const tr = document.createElement('tr');
    if(d._new) tr.classList.add('is-new');
    const photoCell = d.photo
      ? `<img src="${d.photo}" class="row-photo-thumb" style="width:34px;height:34px;object-fit:cover;border-radius:6px;cursor:zoom-in;border:1px solid var(--line);" onclick="openPhotoLightbox('${d.photo}')">`
      : '<span style="color:#5C68A8;">—</span>';
    const editBtn = d._key
      ? `<button class="btn-ghost" style="padding:5px 10px;font-size:11px;border:1px solid var(--line);border-radius:7px;cursor:pointer;color:var(--cyan-300);" data-edit-key="${d._key}">✎ Editar</button>`
      : '<span style="color:#5C68A8;font-size:11px;">—</span>';
    const deleteBtn = d._key
      ? `<button class="btn-ghost" style="padding:5px 9px;font-size:11px;border:1px solid rgba(255,107,107,0.4);border-radius:7px;cursor:pointer;color:var(--lost);margin-left:6px;" data-delete-key="${d._key}" title="Excluir cotação">🗑</button>`
      : '';
    // Fase Extra · Etapa 3 — gera o PDF da cotação sob demanda, a partir dos
    // dados já salvos (inversão de fluxo BDG → PDF, Análise Técnica/Plano Fase Extra)
    const pdfBtn = `<button class="btn-ghost" style="padding:5px 10px;font-size:11px;border:1px solid var(--line);border-radius:7px;cursor:pointer;color:var(--cyan-300);margin-left:6px;" data-pdf-key="${d._key || ''}" title="Gerar PDF desta cotação">⬇ PDF</button>`;
    if(d._key){ tr.style.cursor = 'pointer'; tr.dataset.rowKey = d._key; }
    tr.innerHTML = `
      <td class="mono">${window.escapeHtml(d.ref||'—')}${d._new?'<span class="new-badge">NOVO</span>':''}${d._key?' <span style="color:var(--cyan-300);font-size:10px;" title="Toque na linha para editar">✎</span>':''}</td>
      <td>${window.escapeHtml(d.client||'')}</td>
      <td><span class="seller-tag">${window.escapeHtml(d.seller||'')}</span></td>
      <td>${window.escapeHtml(d.type||'')}</td>
      <td>${window.escapeHtml(d.modal||'')}</td>
      <td>${window.escapeHtml(d.incoterm||'—')}</td>
      <td class="route">${window.escapeHtml(d.origin||'?')} → ${window.escapeHtml(d.dest||'?')}</td>
      <td class="mono">${window.escapeHtml(toISODate(d.week)||'—')}</td>
      <td><span class="stamp ${statusClass(d.status)}">${window.escapeHtml(STATUS_LABEL_PT[d.status]||d.status||'')}</span></td>
      <td style="white-space:normal;max-width:260px;color:#8C97D4;font-size:11.5px;">${window.escapeHtml(d.remarks||'')}</td>
      <td>${photoCell}</td>
      <td>${editBtn}${deleteBtn}${pdfBtn}</td>`;
    frag.appendChild(tr);
  });
  tbody.appendChild(frag);
  tbody.querySelectorAll('[data-edit-key]').forEach(btn=>{
    btn.addEventListener('click', (e)=>{
      e.stopPropagation();
      const row = RAW.find(r=>r._key===btn.dataset.editKey);
      if(row) openEditModal(row, btn.closest('tr'));
    });
  });
  tbody.querySelectorAll('[data-delete-key]').forEach(btn=>{
    btn.addEventListener('click', (e)=>{
      e.stopPropagation();
      deleteQuotation(btn.dataset.deleteKey);
    });
  });
  // Fase Extra · Etapa 3 — "Gerar PDF" sob demanda, direto da linha da listagem
  tbody.querySelectorAll('[data-pdf-key]').forEach(btn=>{
    btn.addEventListener('click', (e)=>{
      e.stopPropagation();
      const row = RAW.find(r=>r._key===btn.dataset.pdfKey);
      if(row) downloadPdfCotacao(row);
    });
  });
  tbody.querySelectorAll('tr[data-row-key]').forEach(rowEl=>{
    rowEl.addEventListener('click', (e)=>{
      if(e.target.closest('.row-photo-thumb')) return; // não abre editar ao ver a foto
      const row = RAW.find(r=>r._key===rowEl.dataset.rowKey);
      if(row) openEditModal(row, rowEl);
    });
  });
}

/* ===== EXPORTAÇÃO: EXCEL (.xlsx) E PDF ===== */
function exportRowsPlain(){
  // Usa os mesmos filtros ativos na tela de Cotações
  return getFiltered().map(d => ({
    'Referência': d.ref || '',
    'Cliente': d.client || '',
    'Vendedor': d.seller || '',
    'Operação': d.type || '',
    'Modal': d.modal || '',
    'Incoterm': d.incoterm || '',
    'Origem': d.origin || '',
    'Destino': d.dest || '',
    'Data': toISODate(d.week) || '',
    'Status': STATUS_LABEL_PT[d.status] || d.status || '',
    'Observações': d.remarks || ''
  }));
}

document.getElementById('btnExportExcel').addEventListener('click', ()=>{
  const rows = exportRowsPlain();
  if(rows.length===0){ alert('Nenhum registro para exportar com os filtros atuais.'); return; }
  const ws = XLSX.utils.json_to_sheet(rows);
  ws['!cols'] = [{wch:16},{wch:22},{wch:10},{wch:14},{wch:10},{wch:10},{wch:16},{wch:16},{wch:12},{wch:14},{wch:34}];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Cotações BDG');
  const stamp = new Date().toISOString().slice(0,10);
  XLSX.writeFile(wb, `Cotacoes_BDG_${stamp}.xlsx`);
});

document.getElementById('btnExportPdf').addEventListener('click', ()=>{
  const rows = exportRowsPlain();
  if(rows.length===0){ alert('Nenhum registro para exportar com os filtros atuais.'); return; }
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ orientation:'landscape', unit:'pt', format:'a4' });
  const stamp = new Date().toLocaleDateString('pt-BR');

  doc.setFillColor(10,20,64);
  doc.rect(0,0, doc.internal.pageSize.getWidth(), 56, 'F');
  doc.setTextColor(255,255,255);
  doc.setFont('helvetica','bold'); doc.setFontSize(15);
  doc.text('BDG · Controle de Cotações — Samsung SDS Logistics', 24, 26);
  doc.setFont('helvetica','normal'); doc.setFontSize(9.5);
  doc.setTextColor(180,200,255);
  doc.text(`Exportado em ${stamp} · ${rows.length} registro(s) com os filtros ativos aplicados`, 24, 42);

  doc.autoTable({
    startY: 70,
    head: [['Referência','Cliente','Vendedor','Operação','Modal','Incoterm','Origem','Destino','Data','Status','Observações']],
    body: rows.map(r => Object.values(r)),
    styles: { fontSize: 7.5, cellPadding: 4, textColor: [20,25,50] },
    headStyles: { fillColor: [20,40,160], textColor: 255, fontStyle:'bold' },
    alternateRowStyles: { fillColor: [235,240,255] },
    margin: { left: 20, right: 20 },
  });
  doc.save(`Cotacoes_BDG_${new Date().toISOString().slice(0,10)}.pdf`);
});


let pendingPhotoBase64 = null;

// Nota de limpeza (débito técnico da auditoria de 29/07/2026, seção 6.2):
// compressImageToBase64() e withTimeout() foram removidas daqui — ficaram
// órfãs desde que o recorte de foto passou a comprimir a imagem
// internamente (window.openPhotoCropper, ver abaixo). O equivalente em
// auth.js já havia sido removido antes; esta é a mesma limpeza aplicada
// aqui, sem alteração de comportamento (nenhuma das duas era chamada em
// lugar nenhum do arquivo).

document.getElementById('nPhoto').addEventListener('change', (e) => {
  const file = e.target.files[0];
  if(!file){ return; }
  // Abre o recorte antes de confirmar — a compressão final já acontece dentro
  // do próprio recorte (window.openPhotoCropper), então não precisa mais do
  // compressImageToBase64/withTimeout para este fluxo.
  window.openPhotoCropper(file, {
    targetSize: 900,
    quality: 0.72,
    round: false,
    onConfirm: (dataUrl) => {
      pendingPhotoBase64 = dataUrl;
      document.getElementById('nPhotoPreview').src = dataUrl;
      document.getElementById('nPhotoPreviewWrap').style.display = 'flex';
    }
  });
  document.getElementById('nPhoto').value = '';
});

document.getElementById('nPhotoRemove').addEventListener('click', () => {
  pendingPhotoBase64 = null;
  document.getElementById('nPhoto').value = '';
  document.getElementById('nPhotoPreviewWrap').style.display = 'none';
});

window.openPhotoLightbox = function(src){
  document.getElementById('photoLightboxImg').src = src;
  document.getElementById('photoLightbox').style.display = 'flex';
}
document.getElementById('photoLightbox').addEventListener('click', () => {
  document.getElementById('photoLightbox').style.display = 'none';
});

/* ===== EDITAR COTAÇÃO ===== */
let editingKey = null;
let editingPhotoBase64 = null; // undefined=sem alteração, null=removida, string=nova foto

/* ===== EXCLUIR COTAÇÃO =====
   Usada tanto pelo botão 🗑 na linha da tabela quanto pelo botão dentro do
   modal de edição. Retorna true se a exclusão foi confirmada e concluída. */
async function deleteQuotation(key){
  const entry = RAW.find(r=>r._key===key);
  const label = entry ? (entry.ref || entry.client || key) : key;
  if(!confirm('Excluir a cotação "'+label+'"? Essa ação não pode ser desfeita.')) return false;
  try{
    await dbRemove(dbRef(db, 'cotacoes/'+key));
    const cu = window.getCurrentUser ? window.getCurrentUser() : {uid:null, data:null};
    if(window.logActivity){
      window.logActivity(cu.uid, cu.data?.username || 'Usuário', 'Excluiu a cotação ' + label);
    }
    return true;
  }catch(err){
    alert('Não foi possível excluir a cotação: ' + err.message);
    return false;
  }
}

/* ===== FASE 4 · ETAPA 1 (UX/UI) — feedback visual padronizado dos botões =====
   setBtnBusy() entra em estado "salvando..." (reaproveita o spinner via
   .btn:disabled já existente no CSS); setBtnDone() finaliza com um pulso de
   sucesso ou erro (classes .btn-success/.btn-error) antes de voltar ao
   rótulo original. Usado por Salvar tabela, Atualizar cotação e Excluir
   cotação — os três pontos citados no plano de evolução. */
function setBtnBusy(btn, busyLabel){
  if(!btn) return;
  if(btn.dataset.origLabel === undefined) btn.dataset.origLabel = btn.textContent;
  btn.disabled = true;
  btn.classList.remove('btn-success','btn-error');
  btn.textContent = busyLabel || 'Salvando...';
}
function setBtnDone(btn, opts){
  if(!btn) return;
  const { ok = true, okLabel, errorLabel, holdMs = 1200 } = opts || {};
  const original = btn.dataset.origLabel !== undefined ? btn.dataset.origLabel : btn.textContent;
  btn.disabled = false;
  btn.classList.add(ok ? 'btn-success' : 'btn-error');
  btn.textContent = ok ? (okLabel || '✓ Salvo') : (errorLabel || '✕ Falhou');
  setTimeout(() => {
    btn.classList.remove('btn-success','btn-error');
    btn.textContent = original;
  }, holdMs);
}

/* ===== FASE 4 · ETAPA 1 (UX/UI) — componente único de "estado vazio" =====
   Gera o HTML padrão (ícone + frase curta + ação sugerida opcional) usado
   pela tabela principal, ranking de clientes, tabela de margem e listas de
   Config, no lugar de textos explicativos isolados em cada tela. */
function emptyStateHtml({ icon = '🗂️', title, sub, actionHtml = '', compact = false } = {}){
  return `<div class="empty-state${compact ? ' compact' : ''}">
    <div class="es-icon">${icon}</div>
    ${title ? `<div class="es-title">${window.escapeHtml(title)}</div>` : ''}
    ${sub ? `<div class="es-sub">${window.escapeHtml(sub)}</div>` : ''}
    ${actionHtml ? `<div class="es-action">${actionHtml}</div>` : ''}
  </div>`;
}

/* ============================================================
   PAINEL DE VISUALIZAÇÃO/EDIÇÃO DE COTAÇÃO (Especificação de melhorias,
   item 4.1)
   ============================================================
   Substitui o antigo modal flutuante por um painel que abre logo abaixo
   da lista (#quotePanel, dentro de view-table — ver index.html), com
   dois modos:
     - 'view' (padrão ao abrir): mostra um resumo somente-leitura
       (#quotePanelSummary), formulário de edição escondido.
     - 'edit' (só após clicar em "Editar"): mostra o formulário real
       (mesmos ids de campo do modal antigo — nenhuma lógica de leitura/
       gravação foi alterada), com o resumo escondido.

   Detecção de "clicar fora durante edição": um listener de click no
   document, ativo só quando o modo é 'edit', compara os valores atuais
   dos campos com o snapshot tirado no momento em que "Editar" foi
   clicado. Se houver diferença E o clique for fora do painel e fora da
   linha que o abriu, abre o modal de confirmação (#discardChangesModal)
   em vez de fechar direto — do contrário, fecha sem perguntar (nada
   mudou, não há o que perder). */

const QUOTE_PANEL_FIELD_IDS = [
  'eClient','eSeller','eType','eModal','eIncoterm','eOrigin','eDest','eDate',
  'eRef','eStatus','eRemarks','eAllIn','eTotalCollect','eHandling',
  'ePeso','eRateAR','eCustoAF','eOriginCharges','eMiaCharges','eTransitTime',
  'eServico','eAgente','ePaymentTerm','eSpread'
];

let quotePanelMode = 'view'; // 'view' | 'edit'
let quotePanelSnapshot = null; // { [fieldId]: value } no momento em que "Editar" foi clicado
let quotePanelOpenerRowEl = null; // <tr> que abriu o painel — clique nela não conta como "fora"

function snapshotQuotePanelFields(){
  const snap = {};
  QUOTE_PANEL_FIELD_IDS.forEach((id) => {
    const el = document.getElementById(id);
    if(el) snap[id] = el.value;
  });
  snap.__photo = editingPhotoBase64; // undefined = não mexeu, null = removida, string = nova
  return snap;
}

function quotePanelHasUnsavedChanges(){
  if(!quotePanelSnapshot) return false;
  return QUOTE_PANEL_FIELD_IDS.some((id) => {
    const el = document.getElementById(id);
    return el && el.value !== quotePanelSnapshot[id];
  }) || editingPhotoBase64 !== quotePanelSnapshot.__photo;
}

/* Resumo somente-leitura (modo 'view'): reaproveita os mesmos dados já
   carregados nos campos do formulário (não lê RAW de novo), então nunca
   diverge do que o formulário mostraria ao entrar em edição. */
function renderQuotePanelSummary(d){
  const summary = document.getElementById('quotePanelSummary');
  if(!summary) return;
  const rows = [
    ['Cliente', d.client || '—'],
    ['Vendedor', d.seller || '—'],
    ['Operação', d.type || '—'],
    ['Modal', labelModalPt(d.modal)],
    ['Incoterm', d.incoterm || '—'],
    ['Rota', (d.origin || '?') + ' → ' + (d.dest || '?')],
    ['Data', toISODate(d.week) || '—'],
    ['Status', STATUS_LABEL_PT[d.status] || d.status || '—'],
  ];
  const financeiro = [
    ['ALL IN', d.allIn != null ? formatDecimalBR(Number(d.allIn)) : '—'],
    ['Total Collect', d.totalCollect != null ? formatDecimalBR(Number(d.totalCollect)) : '—'],
    ['Handling Fee', d.handling != null ? formatDecimalBR(Number(d.handling)) : '—'],
  ];
  const rowsHtml = rows.map(([label, value]) => `
    <div class="quote-summary-row"><span>${window.escapeHtml(label)}</span><b>${window.escapeHtml(String(value))}</b></div>
  `).join('');
  const finHtml = financeiro.map(([label, value]) => `
    <div class="quote-summary-row"><span>${window.escapeHtml(label)}</span><b>${window.escapeHtml(String(value))}</b></div>
  `).join('');
  const remarksHtml = d.remarks
    ? `<div class="quote-summary-remarks"><span>Observações</span><p>${window.escapeHtml(d.remarks)}</p></div>`
    : '';
  const photoHtml = d.photo
    ? `<div class="quote-summary-photo"><img src="${d.photo}" alt="Anexo da cotação" onclick="openPhotoLightbox('${d.photo}')"></div>`
    : '';
  summary.innerHTML = `
    <div class="quote-summary-grid">${rowsHtml}</div>
    <div class="quote-summary-grid quote-summary-grid-fin">${finHtml}</div>
    ${remarksHtml}
    ${photoHtml}
  `;
}
function labelModalPt(modal){
  const map = { AIR: 'Aéreo', OCEAN: 'Marítimo (FCL)', 'OCEAN/LCL': 'Marítimo (LCL)', 'SEA/AIR': 'Sea & Air', TRUCK: 'Rodoviário' };
  return map[modal] || modal || '—';
}

function setQuotePanelMode(mode){
  quotePanelMode = mode;
  const panel = document.getElementById('quotePanel');
  const summary = document.getElementById('quotePanelSummary');
  const editArea = document.getElementById('quotePanelEditArea');
  const badge = document.getElementById('quotePanelModeBadge');
  const btnEnterEdit = document.getElementById('btnEnterEditMode');
  if(!panel) return;
  panel.classList.toggle('is-edit-mode', mode === 'edit');
  if(summary) summary.style.display = mode === 'edit' ? 'none' : '';
  if(editArea) editArea.style.display = mode === 'edit' ? '' : 'none';
  if(badge) badge.textContent = mode === 'edit' ? 'Editando' : 'Visualizando';
  if(btnEnterEdit) btnEnterEdit.style.display = mode === 'edit' ? 'none' : '';
  if(mode === 'edit'){
    quotePanelSnapshot = snapshotQuotePanelFields();
    if(window.resetEditStepper) window.resetEditStepper();
  } else {
    quotePanelSnapshot = null;
  }
}

/* ===== FASE 4 · ETAPA 3 (UX/UI) — stepper do painel "Editar cotação" =====
   Comportamento de acordeão (só uma seção aberta por vez) + barra de
   progresso de 4 segmentos que acende conforme o vendedor visita cada
   bloco (Dados gerais / Rota / Valores / Anexos). Não depende de nenhum
   id de campo — só observa os <details class="edit-section"> do painel. */
(function setupEditStepper(){
  const editArea = document.getElementById('quotePanelEditArea');
  if(!editArea) return;
  const sections = Array.from(editArea.querySelectorAll('.edit-section'));
  const segs = Array.from(editArea.querySelectorAll('.edit-progress-seg'));
  function paintProgress(){
    sections.forEach((sec, i) => {
      if(segs[i]) segs[i].classList.toggle('is-visited', sec.dataset.visited === '1' || sec.open);
    });
  }
  sections.forEach((sec) => {
    sec.addEventListener('toggle', () => {
      if(sec.open){
        sec.dataset.visited = '1';
        // acordeão: fecha as demais seções ao abrir uma nova
        sections.forEach((other) => { if(other !== sec) other.open = false; });
      }
      paintProgress();
    });
  });
  // ao entrar em modo edição (setQuotePanelMode('edit')), reinicia
  // o "visitado" e volta sempre para a primeira seção.
  window.resetEditStepper = function resetEditStepper(){
    sections.forEach((sec, i) => {
      sec.dataset.visited = i === 0 ? '1' : '0';
      sec.open = i === 0;
    });
    paintProgress();
  };
})();

function openEditModal(d, rowEl){
  editingKey = d._key;
  editingPhotoBase64 = undefined;
  quotePanelOpenerRowEl = rowEl || null;
  // Fase 5 (Redesign Premium) · Etapa 2 — widget de mapa: entra em modo
  // "contextual" com a rota desta cotação específica. Chamada aditiva e
  // segura (encadeamento opcional) — nunca bloqueia a abertura do painel.
  window.routeMapSetContext?.(d);
  document.getElementById('editRefLabel').textContent = d.ref || '';
  document.getElementById('eClient').value = d.client || '';
  document.getElementById('eSeller').value = d.seller || 'ROGER';
  document.getElementById('eType').value = d.type || 'Import';
  document.getElementById('eModal').value = d.modal || 'AIR';
  document.getElementById('eIncoterm').value = d.incoterm || '';
  document.getElementById('eOrigin').value = d.origin || '';
  document.getElementById('eDest').value = d.dest || '';
  document.getElementById('eDate').value = toISODate(d.week) || '';
  document.getElementById('eRef').value = d.ref || '';
  document.getElementById('eStatus').value = d.status || 'APPROVED';
  document.getElementById('eRemarks').value = d.remarks || '';
  fillCommercialFields('e', d); // Fase Extra · Etapa 1 (campos comerciais opcionais)
  document.getElementById('ePhoto').value = '';
  const wrap = document.getElementById('ePhotoPreviewWrap');
  if(d.photo){
    document.getElementById('ePhotoPreview').src = d.photo;
    wrap.style.display = 'flex';
  } else {
    wrap.style.display = 'none';
  }
  // Painel sempre abre em modo VISUALIZAÇÃO — nunca direto em edição
  // (critério de aceite da Especificação de melhorias, item 4.1).
  renderQuotePanelSummary(d);
  setQuotePanelMode('view');
  document.getElementById('quotePanel').style.display = '';
  if(quotePanelOpenerRowEl) quotePanelOpenerRowEl.classList.add('is-panel-open');
  // Gaveta: rola até o painel ficar visível, sem forçar o topo da página.
  requestAnimationFrame(() => {
    document.getElementById('quotePanel').scrollIntoView({ behavior:'smooth', block:'nearest' });
  });
}
function closeEditModalForce(){
  document.getElementById('quotePanel').style.display = 'none';
  if(quotePanelOpenerRowEl) quotePanelOpenerRowEl.classList.remove('is-panel-open');
  quotePanelOpenerRowEl = null;
  quotePanelSnapshot = null;
  quotePanelMode = 'view';
  // Fase 5 (Redesign Premium) · Etapa 2 — widget de mapa: sai do modo
  // contextual e volta para overview. Chamada aditiva e segura.
  window.routeMapClearContext?.();
  editingKey = null;
  editingPhotoBase64 = undefined;
}
/* Ponto único de saída do painel: se estiver em modo edição com
   alterações não salvas, abre a confirmação em vez de fechar direto.
   Todo botão/ação que "sai" do painel (✕, Cancelar, clique fora) passa
   por aqui — não só o clique fora, para não deixar nenhum atalho que
   perca edição silenciosamente. */
function closeEditModal(){
  if(quotePanelMode === 'edit' && quotePanelHasUnsavedChanges()){
    document.getElementById('discardChangesModal').style.display = 'flex';
    return;
  }
  closeEditModalForce();
}
document.getElementById('btnCloseEdit').addEventListener('click', closeEditModal);
document.getElementById('btnCancelEdit').addEventListener('click', closeEditModal);
document.getElementById('btnEnterEditMode').addEventListener('click', () => setQuotePanelMode('edit'));
document.getElementById('btnDeleteQuotation').addEventListener('click', async ()=>{
  if(!editingKey) return;
  const btn = document.getElementById('btnDeleteQuotation');
  setBtnBusy(btn, 'Excluindo...');
  const ok = await deleteQuotation(editingKey);
  if(ok){
    setBtnDone(btn, { ok:true, okLabel:'✓ Excluída' });
    closeEditModalForce();
  } else {
    setBtnDone(btn, { ok:false, errorLabel:'✕ Não foi possível' });
  }
});
// Modal de confirmação de saída (Especificação de melhorias, item 4.1):
// "Sair sem salvar" descarta de fato; "Continuar editando" só fecha o
// aviso, o painel permanece aberto e editável exatamente como estava.
document.getElementById('btnDiscardChangesCancel').addEventListener('click', () => {
  document.getElementById('discardChangesModal').style.display = 'none';
  closeEditModalForce();
});
document.getElementById('btnDiscardChangesStay').addEventListener('click', () => {
  document.getElementById('discardChangesModal').style.display = 'none';
});
// Clique fora do painel (mas dentro da view-table, ou em qualquer outro
// lugar da página) durante edição: mesma checagem de closeEditModal(),
// nunca um caminho separado que pule a confirmação. Clique na própria
// linha que abriu o painel não conta como "fora" (evita reabrir/fechar
// em loop ao clicar de novo na mesma linha).
document.addEventListener('click', (e) => {
  const panel = document.getElementById('quotePanel');
  if(!panel || panel.style.display === 'none') return;
  if(panel.contains(e.target)) return;
  /* CORREÇÃO (Auditoria Geral 01/08, Bloco 3) — o botão "Continuar editando"
     não funcionava, e o usuário ficava preso num laço.

     #discardChangesModal é IRMÃO de #quotePanel no HTML, não filho — então a
     checagem panel.contains(e.target) acima não o cobre. A sequência era:
       1. o clique em "Continuar editando" escondia o modal de confirmação;
       2. o mesmo clique continuava subindo (bubbling) até o document;
       3. este ouvinte via painel aberto + clique fora dele + modo 'edit';
       4. chamava closeEditModal(), que detectava alteração não salva;
       5. o modal de confirmação reabria na hora.
     "Sair sem salvar" escapava por acidente: closeEditModalForce() zera o
     display do painel antes, e a checagem da linha acima já retornava.

     A guarda abaixo trata o modal como parte do painel para efeito de
     "clique fora". Preferida a e.stopPropagation() nos dois botões porque
     cobre também o clique no fundo escuro do próprio modal.

     ATENÇÃO à ausência de checagem de display aqui: quando este ouvinte
     roda, o handler do botão JÁ escondeu o modal (target antes de bubble),
     então testar style.display !== 'none' faria a guarda nunca valer — o
     que mantinha o bug intacto. O que importa é de ONDE o clique saiu, não
     o estado atual do modal. */
  const discard = document.getElementById('discardChangesModal');
  if(discard && discard.contains(e.target)) return;
  if(quotePanelOpenerRowEl && quotePanelOpenerRowEl.contains(e.target)) return;
  if(quotePanelMode === 'edit'){
    closeEditModal();
  }
  // modo 'view': clicar fora não fecha sozinho — o painel de visualização
  // não tem risco de perda de dado, então só fecha via ✕/Cancelar
  // explícitos, evitando fechamentos acidentais ao rolar/clicar na lista.
});
// Fase Extra · Etapa 3 — "Gerar PDF" também disponível dentro do painel de
// edição (não só na linha da listagem), para ficar visível assim que o
// vendedor abre a cotação já salva.
document.getElementById('btnGerarPdfModal').addEventListener('click', ()=>{
  if(!editingKey) return;
  const row = RAW.find(r=>r._key===editingKey);
  if(row) downloadPdfCotacao(row);
});
document.getElementById('ePhoto').addEventListener('change', (e)=>{
  const file = e.target.files[0];
  if(!file) return;
  window.openPhotoCropper(file, {
    targetSize: 900,
    quality: 0.72,
    round: false,
    onConfirm: (dataUrl) => {
      editingPhotoBase64 = dataUrl;
      document.getElementById('ePhotoPreview').src = dataUrl;
      document.getElementById('ePhotoPreviewWrap').style.display = 'flex';
    }
  });
  document.getElementById('ePhoto').value = '';
});
document.getElementById('ePhotoRemove').addEventListener('click', ()=>{
  editingPhotoBase64 = null;
  document.getElementById('ePhoto').value = '';
  document.getElementById('ePhotoPreviewWrap').style.display = 'none';
});

document.getElementById('btnUpdateQuotation').addEventListener('click', async ()=>{
  if(!editingKey) return;
  const client = document.getElementById('eClient').value.trim().toUpperCase();
  if(!client){ document.getElementById('eClient').focus(); return; }
  const rawEDate = document.getElementById('eDate').value;
  const date = rawEDate ? toISODate(rawEDate) : null; // Fase 2: valida/normaliza para ISO 8601 antes de gravar
  if(rawEDate && !date){ alert('Data inválida. Use o seletor de data (formato AAAA-MM-DD).'); document.getElementById('eDate').focus(); return; }
  const original = RAW.find(r=>r._key===editingKey) || {};
  const entry = {
    ref: document.getElementById('eRef').value.trim() || original.ref || ('MAN_'+Date.now().toString().slice(-6)),
    client, seller: document.getElementById('eSeller').value,
    type: document.getElementById('eType').value,
    modal: document.getElementById('eModal').value,
    incoterm: document.getElementById('eIncoterm').value,
    origin: document.getElementById('eOrigin').value.trim().toUpperCase(),
    dest: document.getElementById('eDest').value.trim().toUpperCase(),
    week: date || null,
    month: date ? date.slice(0,7) : null,
    status: document.getElementById('eStatus').value,
    remarks: document.getElementById('eRemarks').value.trim(),
    ...readCommercialFields('e'), // Fase Extra · Etapa 1 (campos comerciais opcionais)
    // O painel "Editar" não tem motor de cálculo automático — qualquer edição aqui é,
    // por definição, manual. Preserva a marcação já existente da criação da cotação.
    manualOverrides: original.manualOverrides || {}
  };
  // Correção: como este salvamento é um dbSet() (sobrescreve o nó inteiro), sem o código
  // abaixo os dados de custo/margem (view Faturamento) e os blocos extras (STD/EXP...)
  // de uma cotação criada no motor automático eram apagados silenciosamente ao editar
  // qualquer campo básico aqui neste painel simplificado.
  if (original.custoTotal != null) entry.custoTotal = original.custoTotal;
  if (original.margemUsd != null) entry.margemUsd = original.margemUsd;
  if (original.margemPercentual != null) entry.margemPercentual = original.margemPercentual;
  if (original.blocosExtras) entry.blocosExtras = original.blocosExtras;
  if(editingPhotoBase64 === undefined){ if(original.photo) entry.photo = original.photo; }
  else if(editingPhotoBase64 !== null){ entry.photo = editingPhotoBase64; }
  // se editingPhotoBase64 === null (removida), simplesmente não incluímos 'photo' no objeto

  const btn = document.getElementById('btnUpdateQuotation');
  setBtnBusy(btn, 'Salvando...');
  try{
    await dbSet(dbRef(db, 'cotacoes/'+editingKey), entry);
    const cu = window.getCurrentUser ? window.getCurrentUser() : {uid:null, data:null};
    if(window.logActivity){
      window.logActivity(cu.uid, cu.data?.username || 'Usuário', 'Editou a cotação ' + (entry.ref || editingKey));
    }
    const toast = document.getElementById('editToastMsg');
    toast.textContent = '✓ Cotação atualizada';
    toast.classList.add('show');
    setTimeout(()=>toast.classList.remove('show'), 2500);
    setBtnDone(btn, { ok:true, okLabel:'✓ Atualizada', holdMs:700 });
    // Salvar sempre volta o painel para modo visualização com os dados
    // recém-salvos (não fecha o painel) — o vendedor vê o resultado
    // imediatamente, sem precisar reabrir a cotação para conferir.
    quotePanelSnapshot = null;
    renderQuotePanelSummary(entry);
    setQuotePanelMode('view');
  }catch(err){
    setBtnDone(btn, { ok:false, errorLabel:'✕ Erro ao salvar' });
    alert('Não foi possível salvar a alteração: ' + err.message);
  }
});

/* ===== COMPARATIVO MENSAL ===== */
function renderCompare(){
  const mA = document.getElementById('cMonthA').value;
  const mB = document.getElementById('cMonthB').value;
  if(!mA || !mB) return;
  // Especificação de melhorias, item 5.2 — generalizado de "d.month===mX"
  // (só mês) para getPeriodKey(d.week, compareGranularity)===mX, que
  // funciona para qualquer uma das 6 granularidades sem duplicar lógica.
  const dataA = RAW.filter(d=>getPeriodKey(d.week, compareGranularity)===mA);
  const dataB = RAW.filter(d=>getPeriodKey(d.week, compareGranularity)===mB);

  function stats(arr){
    const total = arr.length;
    const approved = arr.filter(d=>d.status==='APPROVED').length;
    const lost = arr.filter(d=>d.status==='LOST').length;
    const wait = arr.filter(d=>d.status==='NO RESPONSE').length;
    const review = arr.filter(d=>d.status==='UNDER REVIEW').length;
    const rate = total ? (approved/total*100) : 0;
    const cc = {}; arr.forEach(d=>cc[d.client]=(cc[d.client]||0)+1);
    const top = Object.entries(cc).sort((a,b)=>b[1]-a[1])[0];
    // Especificação de melhorias, item 5.2 — valores monetários, além das
    // de quantidade e porcentagem já existentes. Mesmos três campos
    // (ALL IN/Total Collect/Handling) já usados em renderExecutiveDashboard/
    // renderFaturamentoExecutivo, para consistência de fonte no app inteiro.
    let sumAllIn = 0, sumTotalCollect = 0, sumHandling = 0;
    arr.forEach(d=>{
      sumAllIn += Number(d.allIn) || 0;
      sumTotalCollect += Number(d.totalCollect) || 0;
      sumHandling += Number(d.handling) || 0;
    });
    return {total,approved,lost,wait,review,rate,top: top?window.escapeHtml(top[0])+' ('+top[1]+')':'—', sumAllIn, sumTotalCollect, sumHandling};
  }
  const sA = stats(dataA), sB = stats(dataB);
  const fmtUsd = (v) => 'USD ' + v.toLocaleString('en-US', { minimumFractionDigits:0, maximumFractionDigits:0 });

  // Cards dos meses, lado a lado, sempre iguais (sem delta embutido)
  function colHtml(label, s){
    return `
      <h4>${label}</h4>
      <div class="compare-row"><span>Total de cotações</span><b>${s.total}</b></div>
      <div class="compare-row"><span>Aprovadas</span><b>${s.approved}</b></div>
      <div class="compare-row"><span>Perdidas</span><b>${s.lost}</b></div>
      <div class="compare-row"><span>Sem resposta</span><b>${s.wait}</b></div>
      <div class="compare-row"><span>Em revisão</span><b>${s.review}</b></div>
      <div class="compare-row"><span>Taxa de aprovação</span><b>${s.rate.toFixed(1)}%</b></div>
      <div class="compare-row"><span>Cliente destaque</span><b>${s.top}</b></div>
      <div class="compare-row compare-row-money"><span>ALL IN</span><b>${fmtUsd(s.sumAllIn)}</b></div>
      <div class="compare-row compare-row-money"><span>Total Collect</span><b>${fmtUsd(s.sumTotalCollect)}</b></div>
      <div class="compare-row compare-row-money"><span>Handling Fee</span><b>${fmtUsd(s.sumHandling)}</b></div>`;
  }
  document.getElementById('compareColA').innerHTML = colHtml(periodLabel(mA, compareGranularity), sA);
  document.getElementById('compareColB').innerHTML = colHtml(periodLabel(mB, compareGranularity), sB);

  // Resumo comparativo: número absoluto + variação percentual, numa única lista
  function deltaBadge(vA, vB, higherIsBetter=true, isPercentPoint=false){
    const diff = vB - vA;
    if(Math.abs(diff) < 0.05) return '<span class="cs-delta flat">— igual</span>';
    const good = higherIsBetter ? diff>0 : diff<0;
    const arrow = diff>0 ? '▲' : '▼';
    const pct = vA !== 0 ? (diff/Math.abs(vA)*100) : (vB!==0 ? 100 : 0);
    const numTxt = isPercentPoint ? Math.abs(diff).toFixed(1)+' p.p.' : Math.abs(diff).toFixed(diff%1===0?0:1);
    const pctTxt = isPercentPoint ? '' : ` · ${pct>=0?'+':''}${pct.toFixed(0)}%`;
    return `<span class="cs-delta ${good?'up':'down'}">${arrow} ${numTxt}${pctTxt}</span>`;
  }
  const summaryRows = [
    ['Total de cotações', sA.total, sB.total, deltaBadge(sA.total, sB.total)],
    ['Aprovadas', sA.approved, sB.approved, deltaBadge(sA.approved, sB.approved)],
    ['Perdidas', sA.lost, sB.lost, deltaBadge(sA.lost, sB.lost, false)],
    ['Sem resposta', sA.wait, sB.wait, deltaBadge(sA.wait, sB.wait, false)],
    ['Em revisão', sA.review, sB.review, deltaBadge(sA.review, sB.review, false)],
    ['Taxa de aprovação', sA.rate.toFixed(1)+'%', sB.rate.toFixed(1)+'%', deltaBadge(sA.rate, sB.rate, true, true)],
    ['ALL IN', fmtUsd(sA.sumAllIn), fmtUsd(sB.sumAllIn), deltaBadge(sA.sumAllIn, sB.sumAllIn)],
    ['Total Collect', fmtUsd(sA.sumTotalCollect), fmtUsd(sB.sumTotalCollect), deltaBadge(sA.sumTotalCollect, sB.sumTotalCollect)],
    ['Handling Fee', fmtUsd(sA.sumHandling), fmtUsd(sB.sumHandling), deltaBadge(sA.sumHandling, sB.sumHandling)],
  ];
  document.getElementById('compareSummaryList').innerHTML = summaryRows.map(([label,a,b,delta])=>`
    <div class="compare-summary-row">
      <span class="cs-label">${label}</span>
      <span class="cs-vals">${a} <span class="cs-arrow">→</span> ${b}</span>
      ${delta}
    </div>`).join('');

  chartCompare.data.labels = statuses.map(s=>STATUS_LABEL_PT[s]||s);
  chartCompare.data.datasets = [
    { label: periodLabel(mA, compareGranularity), data: statuses.map(s=>dataA.filter(d=>d.status===s).length), backgroundColor:'#2643D6', borderRadius:6 },
    { label: periodLabel(mB, compareGranularity), data: statuses.map(s=>dataB.filter(d=>d.status===s).length), backgroundColor:'#00AEEF', borderRadius:6 }
  ];
  // Especificação de melhorias, item 5.2 — legibilidade do eixo X quando
  // a granularidade for "Semana" (até 52+ rótulos num ano): rotaciona os
  // rótulos só quando a contagem de períodos disponíveis excede ~15, para
  // não poluir visualmente comparações mais curtas (ex.: 6 semestres).
  const totalPeriodosDisponiveis = availablePeriodKeys(compareGranularity).length;
  Object.assign(chartCompare.options.scales.x.ticks, totalPeriodosDisponiveis > 15 ? { maxRotation:60, minRotation:60 } : { maxRotation:0, minRotation:0 });
  chartCompare.update();
}

/* ===== FASE EXTRA · ETAPA 1 — CAMPOS COMERCIAIS DA COTAÇÃO FECHADA =====
   Mesmo bloco de campos existe duas vezes na tela (formulário "+ Adicionar"
   com prefixo "n" e modal "Editar" com prefixo "e"). Para não duplicar a
   leitura/gravação/limpeza desses 11 campos em cada um dos dois lugares,
   as três funções abaixo recebem o prefixo e operam sobre os elementos
   correspondentes — mesmo padrão de reaproveitamento já usado no restante
   deste arquivo (ex.: getFiltered/render compartilhados entre views).

   Todos os campos são opcionais nesta etapa (ver "Regra da versão beta" no
   plano da Fase Extra): o motor de cálculo automático só chega na Etapa 2,
   então aqui só registramos o que o vendedor digitar manualmente. Campos
   numéricos usam parseDecimalBR/formatDecimalBR (utils.js) para aceitar
   tanto "1.720,18" (BR) quanto "1720.18" (internacional). */
const COMMERCIAL_NUMERIC_FIELDS = [
  ['Peso', 'pesoTarifavel'], ['RateAR', 'rateAR'], ['CustoAF', 'custoAF'],
  ['OriginCharges', 'originCharges'], ['MiaCharges', 'miaCharges'], ['Handling', 'handling'],
  ['AllIn', 'allIn'], ['TotalCollect', 'totalCollect'],
];
const COMMERCIAL_TEXT_FIELDS = [
  ['TransitTime', 'transitTime'], ['Agente', 'agente'],
  // Usados pelo gerador de PDF (pdf.js: d.paymentTerm/d.spread). Continuam
  // preenchidos manualmente no formulário — a importação de PDF que também
  // os preenchia saiu na Auditoria Geral de 01/08 (Bloco 2).
  ['PaymentTerm', 'paymentTerm'], ['Spread', 'spread'],
];

function readCommercialFields(prefix) {
  const out = {};
  COMMERCIAL_NUMERIC_FIELDS.forEach(([suffix, key]) => {
    out[key] = parseDecimalBR(document.getElementById(prefix + suffix).value);
  });
  COMMERCIAL_TEXT_FIELDS.forEach(([suffix, key]) => {
    out[key] = document.getElementById(prefix + suffix).value.trim();
  });
  out.servico = document.getElementById(prefix + 'Servico').value;
  // Bloco 1 do formulário "Adicionar cotação" (prefix 'n') é o único que passa pelo motor
  // de cálculo — é aqui que custo (AP) e margem (GAP) são gravados junto com o resto,
  // para alimentar a view de Faturamento. O modal de edição rápida (prefix 'e') não tem
  // motor, então nada é sobrescrito nele.
  if (prefix === 'n') {
    const calc = ULTIMO_CALCULO[''] || {};
    out.custoTotal = calc.apTotal != null ? calc.apTotal : null;
    out.margemUsd = calc.gapValor != null ? calc.gapValor : null;
    out.margemPercentual = calc.gapPercentual != null ? +(calc.gapPercentual * 100).toFixed(4) : null;
  }
  return out;
}

function fillCommercialFields(prefix, d) {
  d = d || {};
  COMMERCIAL_NUMERIC_FIELDS.forEach(([suffix, key]) => {
    document.getElementById(prefix + suffix).value = formatDecimalBR(d[key]);
  });
  COMMERCIAL_TEXT_FIELDS.forEach(([suffix, key]) => {
    document.getElementById(prefix + suffix).value = d[key] || '';
  });
  document.getElementById(prefix + 'Servico').value = d.servico || '';
}

function clearCommercialFields(prefix) {
  fillCommercialFields(prefix, {});
}

/* ============================================================
   Fase Extra · Etapa 3 (v2) — modal como pivô + até 3 blocos por
   cotação + pré-visualização do PDF em tempo real.

   Cada "bloco" representa um serviço fechado dentro da MESMA
   referência (ex.: AIR STD + AIR EXP na SDS_MAO_IA_07497 — Análise
   Técnica, Parte 7.1). O Bloco 1 usa os IDs originais sem sufixo
   (nPeso, nRateAR...) para não quebrar o modelo de dados já salvo
   no Firebase; os Blocos 2 e 3 (opcionais) usam sufixo _b2/_b3 e são
   gravados à parte, em entry.blocosExtras, no momento de salvar.
   Todos os três blocos usam exatamente os mesmos campos e o mesmo
   motor de cálculo — só o sufixo do id muda.
   ============================================================ */
const BLOCOS_SUFIXOS = ['']; // sempre começa só com o Bloco 1
// Guarda o último resultado de custo (AP) e margem (GAP) calculado por bloco — o motor
// calcula isso a cada tecla, mas até agora nenhum desses dois valores era salvo no Firebase
// junto com a cotação (só o lado de venda: Rate/AR, A/F, ALL IN...). É a partir daqui que a
// view "Faturamento" consegue saber quanto cada cotação realmente deu de lucro.
let ULTIMO_CALCULO = {};
const MAX_BLOCOS = 3;

function labelModal(modal) {
  return { AIR: 'AIR', 'OCEAN': 'OCEAN (FCL)', 'OCEAN/LCL': 'OCEAN (LCL)', 'SEA/AIR': 'SEA & AIR', TRUCK: 'TRUCK' }[modal] || modal || '—';
}
function ehMaritimoModal(modal) { return modal === 'OCEAN' || modal === 'OCEAN/LCL'; }

/* ---------- template de um bloco de valores comerciais ---------- */
function blocoFieldsHTML(sfx, comPadrao) {
  const padraoBtns = comPadrao ? `
      <button type="button" id="btnUsarPadraoCustosFixos${sfx}" class="btn-ghost calc-mini-btn" title="Preencher com o valor padrão salvo em Configurações">↺ padrão</button>
      <button type="button" id="btnSalvarPadraoCustosFixos${sfx}" class="btn-ghost calc-mini-btn" title="Salvar o valor atual como novo padrão do sistema">💾 salvar padrão</button>` : '';
  const padraoBtnsTaxa = comPadrao ? `
      <button type="button" id="btnUsarPadraoTaxaMiami${sfx}" class="btn-ghost calc-mini-btn" title="Preencher com o valor padrão salvo em Configurações">↺ padrão</button>
      <button type="button" id="btnSalvarPadraoTaxaMiami${sfx}" class="btn-ghost calc-mini-btn" title="Salvar o valor atual como novo padrão do sistema">💾 salvar padrão</button>` : '';
  return `
    <div class="add-grid">
      <div class="field" id="blocoCalcAereoTitulo${sfx}" style="grid-column:1/-1;">
        <label class="calc-subtitle" style="margin:0;">💡 Cálculo automático (opcional) — motor AP→GAP→AR</label>
      </div>
      <div id="blocoCalcAereo${sfx}" style="display:contents;">
        <div class="field"><label>AP — trecho 1 (USD/Kg)</label><input type="text" inputmode="decimal" id="nApTrecho1${sfx}" placeholder="Ex: 6,15"></div>
        <div class="field"><label>AP — trecho 2 (USD/Kg, opcional)</label><input type="text" inputmode="decimal" id="nApTrecho2${sfx}" placeholder="Ex: 1,20"></div>
        <div class="field"><label>GAP — margem (%)</label><input type="text" inputmode="decimal" id="nGapPercentual${sfx}" placeholder="Ex: 2,5"></div>
        <div class="field">
          <label>Custos fixos em Miami (USD)${padraoBtns}</label>
          <input type="text" inputmode="decimal" id="nCustosFixosMiami${sfx}" placeholder="Ex: 249,00">
        </div>
        <div class="field">
          <label>Taxa Miami ativa (USD/Kg)${padraoBtnsTaxa}</label>
          <input type="text" inputmode="decimal" id="nTaxaMiami${sfx}" placeholder="Ex: 0,60 (screening)">
        </div>
      </div>
      <div class="field" id="blocoCalcMaritimoTitulo${sfx}" style="grid-column:1/-1; display:none;">
        <label class="calc-subtitle" style="margin:0;">💡 Cálculo automático (opcional) — motor por contêiner/W-M</label>
      </div>
      <div id="blocoCalcMaritimo${sfx}" style="display:none;">
        <div class="field"><label>Origin Charges (USD)</label><input type="text" inputmode="decimal" id="nMarCustosOrigem${sfx}" placeholder="Ex: 350,00"></div>
        <div class="field"><label>Freight Charges — FCL (USD/contêiner)</label><input type="text" inputmode="decimal" id="nMarCustosFrete${sfx}" placeholder="Ex: 2.100,00"></div>
        <div class="field"><label>Destination Charges (USD)</label><input type="text" inputmode="decimal" id="nMarCustosDestino${sfx}" placeholder="Ex: 480,00"></div>
        <div class="field"><label>Peso — LCL (toneladas)</label><input type="text" inputmode="decimal" id="nMarPesoTon${sfx}" placeholder="Ex: 1,85"></div>
        <div class="field"><label>Cubagem — LCL (m³)</label><input type="text" inputmode="decimal" id="nMarCubagem${sfx}" placeholder="Ex: 2,40"></div>
        <div class="field"><label>Tarifa por W/M — LCL (USD)</label><input type="text" inputmode="decimal" id="nMarTarifaWM${sfx}" placeholder="Ex: 65,00"></div>
        <div class="field"><label>Taxas portuárias — LCL, total (USD)</label><input type="text" inputmode="decimal" id="nMarTaxasPortuarias${sfx}" placeholder="Ex: 210,00"></div>
        <div class="field"><label>GAP — margem (%)</label><input type="text" inputmode="decimal" id="nMarGapPercentual${sfx}" placeholder="Ex: 8"></div>
      </div>
      <div class="calc-subtitle" style="grid-column:1/-1;">Valores finais desta seção</div>
      <div class="field"><label>Peso tarifável (Kg)</label><input type="text" inputmode="decimal" id="nPeso${sfx}" placeholder="Ex: 228,33"></div>
      <div class="field"><label>Rate — AR (USD/Kg)</label><input type="text" inputmode="decimal" id="nRateAR${sfx}" placeholder="Ex: 7,53"></div>
      <div class="field"><label>A/F (USD)</label><input type="text" inputmode="decimal" id="nCustoAF${sfx}" placeholder="Ex: 1.720,18"></div>
      <div class="field"><label>FCA/EXW/Origin Charges (USD)</label><input type="text" inputmode="decimal" id="nOriginCharges${sfx}" placeholder="Ex: 249,00"></div>
      <div class="field"><label>MIA Charges (USD)</label><input type="text" inputmode="decimal" id="nMiaCharges${sfx}" placeholder="Ex: 136,99"></div>
      <div class="field"><label>Handling Fee (USD)</label><input type="text" inputmode="decimal" id="nHandling${sfx}" placeholder="Ex: 90,00"></div>
      <div class="field"><label>ALL IN (USD)</label><input type="text" inputmode="decimal" id="nAllIn${sfx}" placeholder="Ex: 2.196,18"></div>
      <div class="field"><label>Total Collect (USD)</label><input type="text" inputmode="decimal" id="nTotalCollect${sfx}" placeholder="Ex: 2.106,18"></div>
      <div class="field"><label>Transit Time</label><input type="text" id="nTransitTime${sfx}" placeholder="Ex: 5-7 dias"></div>
      <div class="field"><label>Serviço</label>
        <select id="nServico${sfx}"><option value="">—</option><option value="STD">STD</option><option value="EXP">EXP</option><option value="S&amp;A">S&amp;A (Sea &amp; Air)</option></select>
      </div>
      <div class="field"><label>Agente (origem)</label><input type="text" id="nAgente${sfx}" placeholder="Ex: parceiro em Miami"></div>
      <div class="field"><label>Payment Term</label><input type="text" id="nPaymentTerm${sfx}" placeholder="Ex: 30 dd"></div>
      <div class="field"><label>Spread (PTAX)</label><input type="text" id="nSpread${sfx}" placeholder="Ex: + 3%"></div>
      <div class="margem-preview" id="margemPreview${sfx}" style="grid-column:1/-1;">
        <span class="mp-icon">💹</span>
        <span class="mp-label">Margem estimada nesta seção</span>
        <span class="mp-value" id="margemPreviewValue${sfx}">preencha AP/GAP acima para calcular</span>
      </div>
    </div>`;
}
function blocoCardHTML(sfx, indice) {
  const removivel = sfx !== '';
  const headBtn = removivel ? `<button type="button" class="bloco-remove" data-remover-bloco="${sfx}">✕ remover</button>` : '';
  return `
    <div class="bloco-card" id="blocoCard${sfx}" data-sfx="${sfx}">
      <div class="bloco-card-head">
        <h4>Bloco ${indice} <span class="bloco-badge">serviço ${indice}</span></h4>
        ${headBtn}
      </div>
      ${blocoFieldsHTML(sfx, sfx === '')}
    </div>`;
}

function renderBlocosContainer() {
  const cont = document.getElementById('blocosContainer');
  if (!cont) return;
  cont.innerHTML = BLOCOS_SUFIXOS.map((sfx, i) => blocoCardHTML(sfx, i + 1)).join('');
  BLOCOS_SUFIXOS.forEach(sfx => wireBlocoEvents(sfx));
  document.getElementById('btnAddBloco').style.display = BLOCOS_SUFIXOS.length >= MAX_BLOCOS ? 'none' : '';
  atualizarCamposPorModal();
  renderPreview();
}

document.getElementById('btnAddBloco')?.addEventListener('click', () => {
  if (BLOCOS_SUFIXOS.length >= MAX_BLOCOS) return;
  const sfx = '_b' + (BLOCOS_SUFIXOS.length + 1);
  BLOCOS_SUFIXOS.push(sfx);
  renderBlocosContainer();
});
function removerBloco(sfx) {
  const idx = BLOCOS_SUFIXOS.indexOf(sfx);
  if (idx <= 0) return; // Bloco 1 nunca é removível
  BLOCOS_SUFIXOS.splice(idx, 1);
  delete ULTIMO_CALCULO[sfx];
  renderBlocosContainer();
}

/* ===== cálculo em tempo real, agora por sufixo de bloco =====
   Mesmo motor já usado no Bloco 1 (Análise Técnica, Partes 3-5 e
   7.3), apenas parametrizado para funcionar em qualquer bloco. */
function recalcularCotacaoAerea(sfx) {
  sfx = sfx || '';
  const modal = document.getElementById('nModal').value;
  if (modal !== 'AIR' && modal !== 'SEA/AIR') { delete ULTIMO_CALCULO[sfx]; atualizarMargemPreview(sfx); return; }
  const g = (id) => document.getElementById(id + sfx);
  const peso = parseDecimalBR(g('nPeso')?.value);
  const trecho1 = parseDecimalBR(g('nApTrecho1')?.value);
  const trecho2 = parseDecimalBR(g('nApTrecho2')?.value);
  const gapPct = parseDecimalBR(g('nGapPercentual')?.value);
  const custosFixos = parseDecimalBR(g('nCustosFixosMiami')?.value);
  const taxaMiami = parseDecimalBR(g('nTaxaMiami')?.value);
  const handling = parseDecimalBR(g('nHandling')?.value);
  if (peso === null || trecho1 === null) { delete ULTIMO_CALCULO[sfx]; atualizarMargemPreview(sfx); return; }
  const r = calcularCotacaoAerea({
    pesoTarifavel: peso,
    tarifasPorTrecho: trecho2 !== null ? [trecho1, trecho2] : [trecho1],
    gapPercentual: (gapPct || 0) / 100,
    custosFixosOrigem: custosFixos || 0,
    taxasMiami: taxaMiami !== null ? [{ tarifa: taxaMiami, peso }] : [],
    handlingFee: handling || 0,
  });
  setValorCalculado('nRateAR' + sfx, r.ar);
  setValorCalculado('nCustoAF' + sfx, r.af);
  setValorCalculado('nMiaCharges' + sfx, r.miaCharges);
  setValorCalculado('nAllIn' + sfx, r.allIn);
  setValorCalculado('nTotalCollect' + sfx, r.totalCollect);
  ULTIMO_CALCULO[sfx] = { apTotal: r.apTotal, gapValor: r.gapValor, gapPercentual: (gapPct || 0) / 100 };
  atualizarMargemPreview(sfx);
}
function recalcularCotacaoMaritima(sfx) {
  sfx = sfx || '';
  const modal = document.getElementById('nModal').value;
  if (modal !== 'OCEAN' && modal !== 'OCEAN/LCL') { delete ULTIMO_CALCULO[sfx]; atualizarMargemPreview(sfx); return; }
  const tipo = modal === 'OCEAN/LCL' ? 'LCL' : 'FCL';
  const g = (id) => document.getElementById(id + sfx);
  const custosOrigem = parseDecimalBR(g('nMarCustosOrigem')?.value) || 0;
  let custosDestino = parseDecimalBR(g('nMarCustosDestino')?.value) || 0;
  const gapPct = parseDecimalBR(g('nMarGapPercentual')?.value);
  const input = { tipo, custosOrigem, gapPercentual: (gapPct || 0) / 100 };
  if (tipo === 'FCL') {
    const custosFrete = parseDecimalBR(g('nMarCustosFrete')?.value);
    if (custosFrete === null) { delete ULTIMO_CALCULO[sfx]; atualizarMargemPreview(sfx); return; }
    input.custosFrete = custosFrete;
    input.custosDestino = custosDestino;
  } else {
    const pesoToneladas = parseDecimalBR(g('nMarPesoTon')?.value);
    const cubagemM3 = parseDecimalBR(g('nMarCubagem')?.value);
    const tarifaPorWM = parseDecimalBR(g('nMarTarifaWM')?.value);
    if (tarifaPorWM === null || (pesoToneladas === null && cubagemM3 === null)) { delete ULTIMO_CALCULO[sfx]; atualizarMargemPreview(sfx); return; }
    const taxasPortuariasTotal = parseDecimalBR(g('nMarTaxasPortuarias')?.value) || 0;
    input.pesoToneladas = pesoToneladas || 0;
    input.cubagemM3 = cubagemM3 || 0;
    input.tarifaPorWM = tarifaPorWM;
    input.custosDestino = custosDestino + taxasPortuariasTotal;
  }
  const r = calcularCotacaoMaritima(input);
  setValorCalculado('nOriginCharges' + sfx, r.custosOrigem);
  setValorCalculado('nMiaCharges' + sfx, r.custosFrete);
  setValorCalculado('nAllIn' + sfx, r.ar);
  setValorCalculado('nTotalCollect' + sfx, r.ar);
  ULTIMO_CALCULO[sfx] = { apTotal: r.apTotal, gapValor: r.gapValor, gapPercentual: (gapPct || 0) / 100 };
  atualizarMargemPreview(sfx);
}

/* atualiza o badge "Margem estimada nesta seção" (💹) — não salva nada sozinho,
   só dá feedback visual imediato do GAP calculado, no mesmo espírito do cálculo
   em tempo real do resto do motor. */
function atualizarMargemPreview(sfx) {
  const el = document.getElementById('margemPreview' + sfx);
  const valEl = document.getElementById('margemPreviewValue' + sfx);
  if (!el || !valEl) return;
  const calc = ULTIMO_CALCULO[sfx];
  el.classList.remove('mp-good', 'mp-low', 'mp-neg');
  if (!calc || calc.apTotal == null || calc.gapValor == null) {
    valEl.textContent = 'preencha AP/GAP acima para calcular';
    return;
  }
  const pct = calc.gapPercentual != null ? calc.gapPercentual * 100 : (calc.apTotal > 0 ? (calc.gapValor / calc.apTotal) * 100 : 0);
  valEl.textContent = previewFmtMoney(calc.gapValor) + '  ·  ' + pct.toFixed(1) + '%';
  if (calc.gapValor < 0) el.classList.add('mp-neg');
  else if (pct < 3) el.classList.add('mp-low');
  else el.classList.add('mp-good');
}

/* mostra só os campos do motor do modal escolhido, em TODOS os blocos ativos */
function atualizarCamposPorModal() {
  const modal = document.getElementById('nModal').value;
  const ehMaritimo = ehMaritimoModal(modal);
  const lcl = modal === 'OCEAN/LCL';
  const stepValores = document.getElementById('stepValores');
  const stepValoresVazio = document.getElementById('stepValoresVazio');
  if (stepValores) stepValores.style.display = modal ? '' : 'none';
  if (stepValoresVazio) stepValoresVazio.style.display = modal ? 'none' : '';
  const lbl = document.getElementById('stepValoresModalLabel');
  if (lbl) lbl.textContent = modal ? '· ' + labelModal(modal) : '';

  BLOCOS_SUFIXOS.forEach(sfx => {
    const aereoTitulo = document.getElementById('blocoCalcAereoTitulo' + sfx);
    const aereoBloco = document.getElementById('blocoCalcAereo' + sfx);
    const marTitulo = document.getElementById('blocoCalcMaritimoTitulo' + sfx);
    const marBloco = document.getElementById('blocoCalcMaritimo' + sfx);
    if (aereoTitulo) aereoTitulo.style.display = ehMaritimo ? 'none' : '';
    if (aereoBloco) aereoBloco.style.display = ehMaritimo ? 'none' : 'contents';
    if (marTitulo) marTitulo.style.display = ehMaritimo ? '' : 'none';
    if (marBloco) marBloco.style.display = ehMaritimo ? 'contents' : 'none';
    const freteEl = document.getElementById('nMarCustosFrete' + sfx);
    if (freteEl) freteEl.closest('.field').style.display = (ehMaritimo && !lcl) ? '' : 'none';
    ['nMarPesoTon', 'nMarCubagem', 'nMarTarifaWM', 'nMarTaxasPortuarias'].forEach(id => {
      const el = document.getElementById(id + sfx); if (el) el.closest('.field').style.display = (ehMaritimo && lcl) ? '' : 'none';
    });
    // rótulo "Peso tarifável" não se aplica ao marítimo FCL — some junto com o motor por Kg
    const pesoEl = document.getElementById('nPeso' + sfx);
    if (pesoEl) pesoEl.closest('.field').style.display = (ehMaritimo && !lcl) ? 'none' : '';
    const rateEl = document.getElementById('nRateAR' + sfx);
    if (rateEl) rateEl.closest('.field').style.display = ehMaritimo ? 'none' : '';
  });
}

/* ===== rastreabilidade de valor manual vs. calculado (por bloco) ===== */
const CAMPOS_CALCULAVEIS_BASE = ['nRateAR', 'nCustoAF', 'nOriginCharges', 'nMiaCharges', 'nAllIn', 'nTotalCollect'];
function setValorCalculado(id, valor) {
  const el = document.getElementById(id);
  if (!el) return;
  el.value = formatDecimalBR(valor);
  el.dataset.manual = 'false';
  el.classList.remove('valor-manual');
  el.title = 'Calculado automaticamente pelo motor de cotação';
  renderPreview();
}
function wireCampoManual(id) {
  const el = document.getElementById(id);
  if (!el || el.dataset.manualWired) return;
  el.dataset.manualWired = '1';
  el.addEventListener('input', () => {
    el.dataset.manual = 'true';
    el.classList.add('valor-manual');
    el.title = 'Valor inserido manualmente (sobrescreve o cálculo automático)';
  });
}
function lerSobrescritasManuais(sfx) {
  sfx = sfx || '';
  const out = {};
  CAMPOS_CALCULAVEIS_BASE.forEach(id => {
    const key = id.replace('n', '').replace(/^./, c => c.toLowerCase());
    out[key] = document.getElementById(id + sfx)?.dataset.manual === 'true';
  });
  return out;
}
function limparSobrescritasManuais(sfx) {
  sfx = sfx || '';
  CAMPOS_CALCULAVEIS_BASE.forEach(id => {
    const el = document.getElementById(id + sfx);
    if (!el) return;
    delete el.dataset.manual;
    el.classList.remove('valor-manual');
    el.title = '';
  });
}

/* ===== conecta eventos de um bloco recém-criado ===== */
function wireBlocoEvents(sfx) {
  ['nPeso', 'nApTrecho1', 'nApTrecho2', 'nGapPercentual', 'nCustosFixosMiami', 'nTaxaMiami', 'nHandling'].forEach(id => {
    document.getElementById(id + sfx)?.addEventListener('input', () => recalcularCotacaoAerea(sfx));
  });
  ['nMarCustosOrigem', 'nMarCustosFrete', 'nMarCustosDestino', 'nMarPesoTon', 'nMarCubagem', 'nMarTarifaWM', 'nMarTaxasPortuarias', 'nMarGapPercentual'].forEach(id => {
    document.getElementById(id + sfx)?.addEventListener('input', () => recalcularCotacaoMaritima(sfx));
  });
  CAMPOS_CALCULAVEIS_BASE.forEach(id => wireCampoManual(id + sfx));
  // qualquer campo do bloco também atualiza a pré-visualização do PDF
  document.getElementById('blocoCard' + sfx)?.querySelectorAll('input, select').forEach(el => {
    el.addEventListener('input', renderPreview);
    el.addEventListener('change', renderPreview);
  });
  document.querySelector(`[data-remover-bloco="${sfx}"]`)?.addEventListener('click', () => removerBloco(sfx));
  if (sfx === '') {
    document.getElementById('btnUsarPadraoCustosFixos')?.addEventListener('click', () => {
      const el = document.getElementById('nCustosFixosMiami');
      if (el && configPadraoMiami.custosFixosMiami != null) { el.value = formatDecimalBR(configPadraoMiami.custosFixosMiami); recalcularCotacaoAerea(''); }
    });
    document.getElementById('btnSalvarPadraoCustosFixos')?.addEventListener('click', () => {
      const valor = parseDecimalBR(document.getElementById('nCustosFixosMiami')?.value);
      if (valor === null) { alert('Preencha o campo antes de salvar como padrão.'); return; }
      dbSet(dbRef(db, 'configuracoes/taxasMiami/custosFixosMiami'), valor);
    });
    document.getElementById('btnUsarPadraoTaxaMiami')?.addEventListener('click', () => {
      const el = document.getElementById('nTaxaMiami');
      if (el && configPadraoMiami.taxaMiami != null) { el.value = formatDecimalBR(configPadraoMiami.taxaMiami); recalcularCotacaoAerea(''); }
    });
    document.getElementById('btnSalvarPadraoTaxaMiami')?.addEventListener('click', () => {
      const valor = parseDecimalBR(document.getElementById('nTaxaMiami')?.value);
      if (valor === null) { alert('Preencha o campo antes de salvar como padrão.'); return; }
      dbSet(dbRef(db, 'configuracoes/taxasMiami/taxaMiami'), valor);
    });
  }
}

/* ===== Passo 1 — seletor de Modal como pivô da tela =====
   Clicar num cartão marca o <select id="nModal"> (mantido oculto só
   para não duplicar a leitura em nenhum outro trecho do código) e
   dispara o mesmo fluxo de sempre: mostrar/ocultar campos + recalcular.
   Chama a mesma função abaixo diretamente — em vez de simular um evento
   de 'change' — porque o global Event não está na lista de globais do
   ESLint deste projeto (lint-test.yml acusava "'Event' is not defined"
   em cotacoes.js#L1273); chamar a função direto resolve sem precisar
   editar a configuração do ESLint. */
function aoTrocarModal() {
  atualizarCamposPorModal();
  BLOCOS_SUFIXOS.forEach(sfx => { recalcularCotacaoAerea(sfx); recalcularCotacaoMaritima(sfx); });
  renderPreview();
}
function selecionarModalCard(opt) {
  document.querySelectorAll('.modal-opt').forEach(o => o.classList.remove('active'));
  opt.classList.add('active');
  const nModalEl = document.getElementById('nModal');
  nModalEl.value = opt.dataset.modal;
  aoTrocarModal();
}
document.querySelectorAll('.modal-opt').forEach(opt => {
  opt.addEventListener('click', () => selecionarModalCard(opt));
  opt.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); selecionarModalCard(opt); }
  });
});
document.getElementById('nModal')?.addEventListener('change', aoTrocarModal);
// estado inicial: destaca no seletor visual o modal já selecionado no <select> oculto
document.querySelector(`.modal-opt[data-modal="${document.getElementById('nModal')?.value}"]`)?.classList.add('active');
renderBlocosContainer(); // desenha o Bloco 1 e aplica o estado inicial da tela

/* ===== Pré-visualização do PDF — espelho visual do formulário =====
   Não lê nem grava nada; é só um resumo em HTML do que o pdf.js vai
   efetivamente gerar (mesmo mapeamento de campos, Análise Técnica
   Parte 6), atualizado a cada tecla. */
['nClient', 'nRef', 'nDate', 'nOrigin', 'nDest', 'nIncoterm', 'nType'].forEach(id => {
  document.getElementById(id)?.addEventListener('input', renderPreview);
});
function previewFmtMoney(v) {
  if (v === null || v === undefined || v === '' || isNaN(v)) return '—';
  return 'USD ' + Number(v).toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}
function renderPreview() {
  const el = document.getElementById('pdfPreview');
  if (!el) return;
  const modal = document.getElementById('nModal')?.value;
  const client = document.getElementById('nClient')?.value.trim();
  if (!modal || !client) {
    el.innerHTML = '<div class="pp-empty">Escolha o modal e o cliente para começar a ver o PDF se formando aqui.</div>';
    return;
  }
  const ref = document.getElementById('nRef')?.value.trim() || 'MAO_XX_00000';
  const dateVal = document.getElementById('nDate')?.value;
  const dateFmt = dateVal ? dateVal.split('-').reverse().join('/') : '—';
  const ocean = ehMaritimoModal(modal);
  const servicoLabel = ocean ? 'OCEAN IMPORT/EXPORT' : (modal === 'SEA/AIR' ? 'SEA & AIR' : 'AIR IMPORT/EXPORT');
  const type = document.getElementById('nType')?.value || 'Import';
  const origin = document.getElementById('nOrigin')?.value.trim() || '—';
  const dest = document.getElementById('nDest')?.value.trim() || '—';
  const incoterm = document.getElementById('nIncoterm')?.value || '—';

  const blocosHtml = BLOCOS_SUFIXOS.map((sfx, i) => {
    const g = (id) => document.getElementById(id + sfx)?.value;
    const servico = g('nServico') || '';
    const tituloServico = (modal === 'SEA/AIR' ? 'SEA/AIR' : modal) + (servico ? ' ' + servico : '');
    const peso = parseDecimalBR(g('nPeso'));
    const rate = parseDecimalBR(g('nRateAR'));
    const af = parseDecimalBR(g('nCustoAF'));
    const origemMia = (parseDecimalBR(g('nOriginCharges')) || 0) + (parseDecimalBR(g('nMiaCharges')) || 0);
    const handling = parseDecimalBR(g('nHandling'));
    const allIn = parseDecimalBR(g('nAllIn'));
    const totalCollect = parseDecimalBR(g('nTotalCollect'));
    const linhasAereo = `
      <tr><td>Chargeable Weight (Kg)</td><td>${peso !== null ? peso.toLocaleString('pt-BR') : '—'}</td></tr>
      <tr><td>Rate (USD/Kg)</td><td>${rate !== null ? rate.toLocaleString('pt-BR', { minimumFractionDigits: 2 }) : '—'}</td></tr>
      <tr><td>Cost — A/F</td><td>${previewFmtMoney(af)}</td></tr>
      <tr><td>FCA/EXW/Origin + MIA Charges</td><td>${previewFmtMoney(origemMia)}</td></tr>
      <tr><td>Handling Fee</td><td>${previewFmtMoney(handling)}</td></tr>
      <tr><td>ALL IN</td><td>${previewFmtMoney(allIn)}</td></tr>`;
    const linhasMaritimo = `
      <tr><td>Ocean/Air Freight</td><td>${previewFmtMoney(af)}</td></tr>
      <tr><td>Destination Charges</td><td>${previewFmtMoney(origemMia)}</td></tr>
      <tr><td>Handling Fee</td><td>${previewFmtMoney(handling)}</td></tr>
      <tr><td>ALL IN</td><td>${previewFmtMoney(allIn)}</td></tr>`;
    return `
      <div class="pp-block-title">${window.escapeHtml(tituloServico)}${BLOCOS_SUFIXOS.length > 1 ? ' <span style="font-weight:400;color:#6E76A0;">· bloco ' + (i + 1) + '</span>' : ''}</div>
      <table class="pp-table pp-cost"><tbody>${ocean ? linhasMaritimo : linhasAereo}</tbody></table>
      <table class="pp-table"><tbody>
        <tr><td>Total Collect</td><td style="text-align:right;">${previewFmtMoney(totalCollect)}</td></tr>
      </tbody></table>`;
  }).join('<div style="height:6px;"></div>');

  el.innerHTML = `
    <div style="display:flex; justify-content:space-between; align-items:flex-start;">
      <span class="pp-conf">CONFIDENCIAL</span>
      <span class="pp-brand">SAMSUNG SDS<small>Realize your vision</small></span>
    </div>
    <hr>
    <div class="pp-banner">
      <div class="pp-ref">QUOTATION SDS ${window.escapeHtml(ref)}</div>
      <div class="pp-sub">${window.escapeHtml(type)} · ${servicoLabel}</div>
    </div>
    <div class="pp-meta">
      <div><b>CUSTOMER</b>${window.escapeHtml(client)}</div>
      <div><b>DATE</b>${dateFmt}</div>
    </div>
    <table class="pp-table"><tbody>
      <tr><td>${ocean ? 'PORT' : 'AIRPORT'} OF LOADING</td><td>${window.escapeHtml(origin)}</td></tr>
      <tr><td>${ocean ? 'PORT' : 'AIRPORT'} OF DISCHARGE</td><td>${window.escapeHtml(dest)}</td></tr>
      <tr><td>INCOTERM</td><td>${window.escapeHtml(incoterm)}</td></tr>
    </tbody></table>
    ${blocosHtml}
  `;
}
/* ============================================================
   Fase Extra · Etapa 3 — Camada 1 completa: tela "⚙️ Configurações"
   (Plano Fase Extra, Seção 3.1). Guarda, sem precisar alterar código,
   os valores que hoje ficam "fixos" espalhados pela planilha:
   - tabela de taxas e custos padrão (pickup, delivery, screening,
     consolidação, docs, ISC, AES...), com tarifa e mínimo por linha
     (Análise Técnica, Parte 4 — mesma estrutura de colunas
     tarifa/peso/mínimo da tabela Q95:U104);
   - listas simples reaproveitadas em toda a tela de cotação: agentes
     de origem, aeroportos, portos, parceiros de Miami, tipos de
     handling, rotas frequentes e clientes frequentes.
   Nenhuma dessas listas é obrigatória — são atalhos/referência; o
   vendedor sempre pode digitar um valor novo à mão em qualquer campo
   (mesma regra da versão beta já aplicada ao resto do módulo). */

// ---- 1) Tabela de taxas e custos padrão ----
let taxasPadrao = []; // [{ nome, tarifa, minimo }]
const taxasPadraoRef = dbRef(db, 'configuracoes/taxasPadrao');
onValue(taxasPadraoRef, (snapshot) => {
  const val = snapshot.val();
  taxasPadrao = Array.isArray(val) ? val : (val ? Object.values(val) : []);
  if (taxasPadrao.length === 0) {
    // Sementes iniciais — mesmos nomes de taxa identificados na Análise
    // Técnica, Parte 4 (Q95:U104), só para não começar a tabela vazia.
    taxasPadrao = [
      { nome: 'ARPT PICK UP', tarifa: 0.06, minimo: 45 },
      { nome: 'ARPT DEL', tarifa: 0.06, minimo: 45 },
      { nome: 'screening', tarifa: 0.60, minimo: 20 },
      { nome: 'CONSOL', tarifa: 0.10, minimo: 45 },
    ];
  }
  renderTabelaTaxasPadrao();
});

function renderTabelaTaxasPadrao() {
  const wrap = document.getElementById('tabelaTaxasPadrao');
  if (!wrap) return;
  wrap.innerHTML = `
    <div style="display:grid; grid-template-columns:2fr 1fr 1fr 32px; gap:8px; font-size:11px; color:var(--ink-500); margin-bottom:4px;">
      <span>Nome da taxa</span><span>Tarifa (USD)</span><span>Mínimo (USD, opcional)</span><span></span>
    </div>
    ${taxasPadrao.map((t, i) => `
      <div style="display:grid; grid-template-columns:2fr 1fr 1fr 32px; gap:8px; margin-bottom:6px;">
        <input type="text" data-taxa-idx="${i}" data-taxa-campo="nome" value="${window.escapeHtml(t.nome || '')}" placeholder="Ex: pickup">
        <input type="text" inputmode="decimal" data-taxa-idx="${i}" data-taxa-campo="tarifa" value="${t.tarifa != null ? formatDecimalBR(t.tarifa) : ''}" placeholder="Ex: 0,06">
        <input type="text" inputmode="decimal" data-taxa-idx="${i}" data-taxa-campo="minimo" value="${t.minimo != null ? formatDecimalBR(t.minimo) : ''}" placeholder="Ex: 45,00">
        <button type="button" class="btn-ghost" data-remover-taxa-idx="${i}" style="cursor:pointer; border-radius:6px;">×</button>
      </div>`).join('')}
  `;
  wrap.querySelectorAll('input[data-taxa-idx]').forEach(inp => {
    inp.addEventListener('input', () => {
      const idx = Number(inp.dataset.taxaIdx);
      const campo = inp.dataset.taxaCampo;
      if (!taxasPadrao[idx]) return;
      taxasPadrao[idx][campo] = campo === 'nome' ? inp.value : (parseDecimalBR(inp.value) ?? undefined);
    });
  });
  wrap.querySelectorAll('button[data-remover-taxa-idx]').forEach(btn => {
    btn.addEventListener('click', () => {
      taxasPadrao.splice(Number(btn.dataset.removerTaxaIdx), 1);
      renderTabelaTaxasPadrao();
    });
  });
}
document.getElementById('btnAddTaxaPadrao')?.addEventListener('click', () => {
  taxasPadrao.push({ nome: '', tarifa: null, minimo: null });
  renderTabelaTaxasPadrao();
});
document.getElementById('btnSalvarTaxasPadrao')?.addEventListener('click', () => {
  const btn = document.getElementById('btnSalvarTaxasPadrao');
  const limpo = taxasPadrao.filter(t => t.nome && t.nome.trim());
  setBtnBusy(btn, '💾 Salvando...');
  dbSet(taxasPadraoRef, limpo).then(() => {
    setBtnDone(btn, { ok:true, okLabel:'✓ Tabela salva' });
    const toast = document.getElementById('toastMsg');
    toast.textContent = '✓ Tabela de taxas salva';
    toast.classList.add('show');
    setTimeout(() => toast.classList.remove('show'), 3000);
  }).catch(err => {
    setBtnDone(btn, { ok:false, errorLabel:'✕ Não salvou', holdMs:2200 });
    alert('Não foi possível salvar: ' + err.message);
  });
});

// ---- 2) Listas simples reaproveitadas na tela de cotação ----
const CONFIG_LISTAS = [
  { key: 'agentesOrigem', label: 'Agentes de origem', placeholder: 'Ex: Samsung SDS Xangai' },
  { key: 'aeroportos', label: 'Aeroportos', placeholder: 'Ex: MIA — Miami Intl' },
  { key: 'portos', label: 'Portos', placeholder: 'Ex: Santos (BR)' },
  { key: 'parceirosMiami', label: 'Parceiros em Miami', placeholder: 'Ex: XYZ Logistics' },
  { key: 'tiposHandling', label: 'Tipos de handling', placeholder: 'Ex: Handling padrão' },
  { key: 'rotas', label: 'Rotas frequentes', placeholder: 'Ex: HKG–MIA–MAO' },
  { key: 'clientesFrequentes', label: 'Empresas / clientes frequentes', placeholder: 'Ex: INVENTUS POWER' },
];
// configListas já declarado no topo do arquivo (evita TDZ com render())
const configListasRef = dbRef(db, 'configuracoes/listas');
onValue(configListasRef, (snapshot) => {
  configListas = snapshot.val() || {};
  renderConfigListasContainer();
  render(); // Fase Extra · Etapa 3 — clientes frequentes também entram no autocomplete do campo "Cliente"
});

function renderConfigListasContainer() {
  const container = document.getElementById('configListasContainer');
  if (!container) return;
  container.innerHTML = CONFIG_LISTAS.map(({ key, label, placeholder }) => `
    <div>
      <label style="display:block; font-size:12px; color:var(--cyan-300); margin-bottom:6px;">${label}</label>
      <div style="display:flex; gap:6px; margin-bottom:8px;">
        <input type="text" id="configListaInput-${key}" placeholder="${placeholder}" style="flex:1;">
        <button type="button" class="btn-ghost" data-add-lista="${key}" style="padding:6px 10px; border-radius:8px; cursor:pointer;">＋</button>
      </div>
      <div id="configListaChips-${key}" style="display:flex; flex-wrap:wrap; gap:6px;"></div>
    </div>
  `).join('');

  CONFIG_LISTAS.forEach(({ key }) => renderChipsLista(key));

  container.querySelectorAll('button[data-add-lista]').forEach(btn => {
    btn.addEventListener('click', () => adicionarItemLista(btn.dataset.addLista));
  });
  container.querySelectorAll('input[id^="configListaInput-"]').forEach(inp => {
    inp.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') { e.preventDefault(); adicionarItemLista(inp.id.replace('configListaInput-', '')); }
    });
  });
}
function renderChipsLista(key) {
  const wrap = document.getElementById('configListaChips-' + key);
  if (!wrap) return;
  const itens = configListas[key] || [];
  wrap.innerHTML = itens.map((item, i) => `
    <span style="display:inline-flex; align-items:center; gap:5px; padding:4px 8px; border-radius:999px; background:rgba(255,255,255,0.06); border:1px solid var(--line); font-size:12px;">
      ${window.escapeHtml(item)}
      <button type="button" data-remover-lista="${key}" data-remover-idx="${i}" style="background:none; border:none; color:var(--ink-500); cursor:pointer; font-size:13px; line-height:1;">×</button>
    </span>`).join('') || emptyStateHtml({ icon:'📄', title:'Nenhum item ainda', sub:'Use o campo acima para adicionar o primeiro.', compact:true });
  wrap.querySelectorAll('button[data-remover-lista]').forEach(btn => {
    btn.addEventListener('click', () => {
      const k = btn.dataset.removerLista;
      const idx = Number(btn.dataset.removerIdx);
      const nova = (configListas[k] || []).slice();
      nova.splice(idx, 1);
      dbSet(dbRef(db, 'configuracoes/listas/' + k), nova);
    });
  });
}
function adicionarItemLista(key) {
  const input = document.getElementById('configListaInput-' + key);
  const valor = input?.value.trim();
  if (!valor) return;
  const atual = configListas[key] || [];
  if (atual.some(v => v.toLowerCase() === valor.toLowerCase())) { input.value = ''; return; } // evita duplicado
  dbSet(dbRef(db, 'configuracoes/listas/' + key), [...atual, valor]).then(() => { input.value = ''; });
}

/* ===== leitura de um bloco extra (Bloco 2/3) para gravação =====
   Mesmo shape de readCommercialFields('n'), mas para qualquer sufixo
   de bloco — usado só para os blocos ALÉM do Bloco 1 (que continua
   gravado nos campos de topo da cotação, para não quebrar a listagem/
   gráficos/PDF de uma cotação já existentes). */
function readBlocoComercial(sfx) {
  const out = {};
  COMMERCIAL_NUMERIC_FIELDS.forEach(([suffix, key]) => {
    out[key] = parseDecimalBR(document.getElementById('n' + suffix + sfx)?.value);
  });
  COMMERCIAL_TEXT_FIELDS.forEach(([suffix, key]) => {
    out[key] = document.getElementById('n' + suffix + sfx)?.value.trim() || '';
  });
  out.servico = document.getElementById('nServico' + sfx)?.value || '';
  out.manualOverrides = lerSobrescritasManuais(sfx);
  const calc = ULTIMO_CALCULO[sfx] || {};
  out.custoTotal = calc.apTotal != null ? calc.apTotal : null;
  out.margemUsd = calc.gapValor != null ? calc.gapValor : null;
  out.margemPercentual = calc.gapPercentual != null ? +(calc.gapPercentual * 100).toFixed(4) : null;
  return out;
}

/* ===== ADD QUOTATION ===== */
document.getElementById('btnAdd').addEventListener('click', ()=>{
  const client = document.getElementById('nClient').value.trim().toUpperCase();
  if(!client){ document.getElementById('nClient').focus(); return; }
  const rawDate = document.getElementById('nDate').value; // já vem em AAAA-MM-DD do <input type=date>
  const date = rawDate ? toISODate(rawDate) : null; // Fase 2: valida/normaliza para ISO 8601 antes de gravar
  if(rawDate && !date){ alert('Data inválida. Use o seletor de data (formato AAAA-MM-DD).'); document.getElementById('nDate').focus(); return; }
  // Blocos 2 e 3 (opcionais): mesma referência, mesmo PDF final, seções
  // extras — pedido do usuário, mesmo padrão da SDS_MAO_IA_07497 (STD+EXP).
  const blocosExtras = BLOCOS_SUFIXOS.slice(1).map(sfx => readBlocoComercial(sfx));
  const entry = {
    ref: document.getElementById('nRef').value.trim() || ('MAN_'+Date.now().toString().slice(-6)),
    client, seller: document.getElementById('nSeller').value,
    type: document.getElementById('nType').value,
    modal: document.getElementById('nModal').value,
    incoterm: document.getElementById('nIncoterm').value,
    origin: document.getElementById('nOrigin').value.trim().toUpperCase(),
    dest: document.getElementById('nDest').value.trim().toUpperCase(),
    week: date || null,
    month: date ? date.slice(0,7) : null,
    status: document.getElementById('nStatus').value,
    remarks: document.getElementById('nRemarks').value.trim(),
    ...readCommercialFields('n'), // Fase Extra · Etapa 1 (campos comerciais opcionais — Bloco 1)
    manualOverrides: lerSobrescritasManuais(''), // Fase Extra · Etapa 3 (rastreabilidade — Seção 7, Bloco 1)
    blocosExtras, // Blocos 2/3 — cada um vira uma seção própria no mesmo PDF (pdf.js)
  };
  if(pendingPhotoBase64) entry.photo = pendingPhotoBase64;
  const btn = document.getElementById('btnAdd');
  btn.disabled = true; btn.textContent = 'Salvando...';
  dbPush(cotacoesRef, entry).then((newRef)=>{
    newlyAddedKeys.add(newRef.key);
    const toast = document.getElementById('toastMsg');
    toast.textContent = '✓ Cotação salva na nuvem';
    toast.classList.add('show');
    setTimeout(()=>toast.classList.remove('show'), 3000);
    ['nClient','nOrigin','nRef','nRemarks'].forEach(id=>document.getElementById(id).value='');
    document.getElementById('nDate').value='';
    document.getElementById('nPhoto').value='';
    document.getElementById('nPhotoPreviewWrap').style.display='none';
    clearCommercialFields('n'); // Fase Extra · Etapa 1
    ['nApTrecho1', 'nApTrecho2', 'nGapPercentual', 'nCustosFixosMiami', 'nTaxaMiami',
     'nMarCustosOrigem', 'nMarCustosFrete', 'nMarCustosDestino', 'nMarPesoTon', 'nMarCubagem', 'nMarTarifaWM', 'nMarTaxasPortuarias', 'nMarGapPercentual'].forEach(id => {
      const el = document.getElementById(id); if (el) el.value = '';
    }); // Fase Extra · Etapa 3 — limpa também os campos do cálculo automático (aéreo e marítimo)
    limparSobrescritasManuais(''); // reseta a sinalização de "valor manual" do Bloco 1
    BLOCOS_SUFIXOS.length = 1; BLOCOS_SUFIXOS[0] = ''; // volta a ter só o Bloco 1 na tela
    ULTIMO_CALCULO = {}; // limpa custo/margem calculados — próxima cotação começa do zero
    renderBlocosContainer();
    pendingPhotoBase64 = null;
  }).catch(err=>{
    alert('Não foi possível salvar na nuvem agora: ' + err.message + '\nVerifique sua conexão ou as regras do Firebase.');
  }).finally(()=>{
    btn.disabled = false; btn.textContent = '＋ Adicionar cotação';
  });
});

/* REMOVIDO (Auditoria Geral 01/08, Bloco 2) — handler do botão "Baixar HTML
   atualizado" (#btnDownload).

   Ele vinha da época em que o app era um HTML único com o array RAW embutido
   no próprio arquivo: bastava reescrever o literal entre os marcadores
   DATA_MARKER e salvar a página inteira. Depois da modularização (Fase 3), o
   RAW passou a viver neste módulo, e o .replace() continuou rodando sobre
   document.documentElement.outerHTML — que não contém marcador nenhum. A
   regex nunca casava, então o arquivo baixado saía sem dado algum e ainda
   apontava para ./cotacoes.js por caminho relativo, não abrindo fora do
   servidor. Estava quebrado desde a modularização, sem ninguém notar.

   Substituto: a view Cotações já exporta os mesmos dados, respeitando os
   filtros ativos, por duas vias testadas — btnExportExcel e btnExportPdf,
   ambas alimentadas por exportRowsPlain(). */

/* ===== IMPORTAÇÃO EM MASSA POR PLANILHA (EXCEL) ===== */
const TEMPLATE_HEADERS = ['Referência','Cliente','Vendedor','Operação','Modal','Incoterm','Origem','Destino','Data','Status','Observações'];
document.getElementById('btnDownloadTemplate').addEventListener('click', ()=>{
  const exampleRow = {
    'Referência': 'MAO_IA_00000',
    'Cliente': 'NOME DO CLIENTE (exemplo — apague esta linha)',
    'Vendedor': 'ROGER',
    'Operação': 'Import',
    'Modal': 'AIR',
    'Incoterm': 'EXW',
    'Origem': 'SHANGHAI',
    'Destino': 'MAO',
    'Data': '2026-07-15',
    'Status': 'APPROVED',
    'Observações': ''
  };
  const ws = XLSX.utils.json_to_sheet([exampleRow], {header: TEMPLATE_HEADERS});
  ws['!cols'] = [{wch:22},{wch:26},{wch:10},{wch:16},{wch:11},{wch:10},{wch:16},{wch:16},{wch:12},{wch:14},{wch:34}];

  const instructions = [
    ['Como usar este modelo'],
    ['1. Preencha uma linha por cotação, sem alterar os nomes das colunas na aba "Cotações".'],
    ['2. Apague a linha de exemplo antes de importar (ou apenas sobrescreva os valores dela).'],
    ['3. Volte ao painel e clique em "Importar planilha preenchida" para enviar tudo de uma vez.'],
    [],
    ['Campo', 'Valores aceitos'],
    ['Referência', 'Livre (opcional — se deixar em branco, o painel gera uma automaticamente)'],
    ['Cliente', 'Livre — campo obrigatório (linhas sem cliente são ignoradas na importação)'],
    ['Vendedor', 'ROGER ou JARDEL'],
    ['Operação', TYPE_VALUES.join(', ')],
    ['Modal', 'AIR, OCEAN, OCEAN/LCL, SEA/AIR ou TRUCK'],
    ['Incoterm', 'EXW, FCA, CPT, CIP, DAP, DPU, DDP, FOB, CFR ou CIF (opcional)'],
    ['Origem', 'Livre'],
    ['Destino', 'Livre'],
    ['Data', 'Formato AAAA-MM-DD, ex: 2026-07-15 (opcional)'],
    ['Status', 'APPROVED, LOST, NO RESPONSE ou UNDER REVIEW'],
    ['Observações', 'Livre (opcional)']
  ];
  const wsInfo = XLSX.utils.aoa_to_sheet(instructions);
  wsInfo['!cols'] = [{wch:16},{wch:78}];

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Cotações');
  XLSX.utils.book_append_sheet(wb, wsInfo, 'Instruções');
  XLSX.writeFile(wb, 'Modelo_Importacao_Cotacoes_BDG.xlsx');
});

document.getElementById('btnImportExcel').addEventListener('click', ()=>{
  document.getElementById('importExcelInput').click();
});

document.getElementById('importExcelInput').addEventListener('change', async (e)=>{
  const file = e.target.files[0];
  if(!file) return;
  const btnImport = document.getElementById('btnImportExcel');
  const originalLabel = btnImport.textContent;
  btnImport.disabled = true; btnImport.textContent = 'Lendo planilha...';
  try{
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, {type:'array', cellDates:true});
    const sheetName = workbook.SheetNames.find(n => /cota/i.test(n)) || workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, {defval:''});

    const entries = [];
    let skipped = 0;
    rows.forEach((row, idx)=>{
      const client = String(row['Cliente']||'').trim().toUpperCase();
      if(!client || client.startsWith('NOME DO CLIENTE')){ skipped++; return; }
      const date = normalizeDateImport(row['Data'], XLSX.SSF.parse_date_code);
      entries.push({
        ref: String(row['Referência']||'').trim() || ('MAN_'+Date.now().toString().slice(-6)+'_'+idx),
        client,
        seller: String(row['Vendedor']||'ROGER').trim().toUpperCase() || 'ROGER',
        type: normalizeTypeImport(row['Operação']),
        modal: String(row['Modal']||'AIR').trim().toUpperCase(),
        incoterm: String(row['Incoterm']||'').trim().toUpperCase(),
        origin: String(row['Origem']||'').trim().toUpperCase(),
        dest: String(row['Destino']||'').trim().toUpperCase(),
        week: date,
        month: date ? date.slice(0,7) : null,
        status: normalizeStatusImport(row['Status']),
        remarks: String(row['Observações']||'').trim()
      });
    });

    if(entries.length === 0){
      alert('Nenhuma linha válida encontrada na planilha. Confira se a coluna "Cliente" está preenchida e se a linha de exemplo foi apagada/substituída.');
      return;
    }
    const confirmMsg = entries.length + (entries.length===1 ? ' cotação será importada' : ' cotações serão importadas') + ' para o painel'
      + (skipped ? ' (' + skipped + ' linha(s) sem cliente foram ignoradas)' : '') + '. Continuar?';
    if(!confirm(confirmMsg)) return;

    btnImport.textContent = 'Importando 0/'+entries.length+'...';
    let imported = 0;
    for(const entry of entries){
      const newRef = await dbPush(cotacoesRef, entry);
      newlyAddedKeys.add(newRef.key);
      imported++;
      btnImport.textContent = 'Importando '+imported+'/'+entries.length+'...';
    }

    const cu = window.getCurrentUser ? window.getCurrentUser() : {uid:null, data:null};
    if(window.logActivity){
      window.logActivity(cu.uid, cu.data?.username || 'Usuário', 'Importou '+entries.length+' cotações via planilha Excel');
    }

    const toast = document.getElementById('importToastMsg');
    toast.textContent = '✓ '+entries.length+' cotações importadas';
    toast.classList.add('show');
    setTimeout(()=>toast.classList.remove('show'), 3500);
  }catch(err){
    alert('Não foi possível importar a planilha: ' + err.message + '\nVerifique se o arquivo é o modelo baixado no painel (.xlsx) e tente novamente.');
  }finally{
    btnImport.disabled = false; btnImport.textContent = originalLabel;
    document.getElementById('importExcelInput').value = '';
  }
});

/* ===== EVENTS ===== */
document.getElementById('fSearch').addEventListener('input', e=>{ state.search = e.target.value; render(); });
document.getElementById('fMonth').addEventListener('change', e=>{ state.month = e.target.value; render(); });
document.getElementById('fClient').addEventListener('change', e=>{ state.client = e.target.value; render(); });
document.getElementById('fSeller').addEventListener('change', e=>{ state.seller = e.target.value; render(); });
document.getElementById('fStatus').addEventListener('change', e=>{ state.status = e.target.value; render(); });
document.getElementById('fModal').addEventListener('change', e=>{ state.modal = e.target.value; render(); });
document.getElementById('fType').addEventListener('change', e=>{ state.type = e.target.value; render(); });
// Versão de demonstração executiva (Fase 5) — filtro de período próprio da
// view Faturamento, independente dos filtros globais acima (state.*).
document.getElementById('fatxPeriodFilter')?.addEventListener('change', ()=>{ renderFaturamentoExecutivo(); });
document.getElementById('btnReset').addEventListener('click', ()=>{
  state = { search:'', month:'', client:'', seller:'', status:'', modal:'', type:'', sortKey:'week', sortDir:'desc' };
  document.getElementById('fSearch').value='';
  ['fMonth','fClient','fSeller','fStatus','fModal','fType'].forEach(id=>document.getElementById(id).value='');
  render();
});
document.querySelectorAll('thead th').forEach(th=>{
  th.addEventListener('click', ()=>{
    const key = th.dataset.key;
    if(state.sortKey===key){ state.sortDir = state.sortDir==='asc' ? 'desc':'asc'; }
    else { state.sortKey = key; state.sortDir='asc'; }
    render();
  });
});
document.getElementById('cMonthA').addEventListener('change', renderCompare);
document.getElementById('cMonthB').addEventListener('change', renderCompare);
// Especificação de melhorias, item 5.2 — trocar a granularidade repopula
// os dois selects de período (novo conjunto de chaves disponíveis para
// a granularidade escolhida) e já re-renderiza a comparação em seguida.
document.getElementById('cGranularity').addEventListener('change', (e)=>{
  // Validação defensiva: só aceita uma das granularidades que utils.js
  // realmente sabe calcular (PERIOD_GRANULARITIES é a mesma lista usada
  // por getPeriodKey) — protege contra um <option> adicionado no HTML
  // no futuro sem o case correspondente em getPeriodKey.
  if(!PERIOD_GRANULARITIES.includes(e.target.value)) return;
  compareGranularity = e.target.value;
  rebuildCompareSelects();
  renderCompare();
});

buildCharts();
recomputeMeta();
render();

/* ============================================================
   SELETOR DE GRÁFICOS DOS INDICADORES (Especificação de melhorias,
   item 2.3)
   ============================================================
   As 4 instâncias Chart.js já existem (criadas por buildCharts() acima)
   — esta função só controla qual <canvas> fica visível. Como o Chart.js
   calcula dimensões a partir do elemento pai, um canvas dentro de um
   .kpi-chart-pane com display:none não tem largura/altura reais no
   momento em que os dados são atualizados por render(); por isso, ao
   TROCAR de aba, chamamos .resize() na instância que acabou de ficar
   visível — sem isso, o gráfico poderia renderizar cortado/vazio na
   primeira troca. Persistência em localStorage: pequeno detalhe de
   conveniência, mesmo padrão informal de "lembrar preferência de sessão"
   que os filtros do painel já têm. */
const KPI_CHART_TAB_STORAGE_KEY = 'bdg:kpiChartTab';
const KPI_CHART_INSTANCES = {
  evolucao: () => chartExecEvolution,
  participacao: () => chartExecShare,
  volume: () => chartMonth,
  status: () => chartStatus,
};
function setupKpiChartTabs(){
  const tabsRoot = document.getElementById('kpiChartTabs');
  if(!tabsRoot) return;
  const tabs = Array.from(tabsRoot.querySelectorAll('[data-chart-tab]'));
  const panes = Array.from(document.querySelectorAll('[data-chart-pane]'));

  function activate(tabName, { persist = true } = {}){
    if(!KPI_CHART_INSTANCES[tabName]) return;
    tabs.forEach(t => {
      const isActive = t.dataset.chartTab === tabName;
      t.classList.toggle('active', isActive);
      t.setAttribute('aria-selected', String(isActive));
    });
    panes.forEach(p => p.classList.toggle('is-active', p.dataset.chartPane === tabName));
    const chart = KPI_CHART_INSTANCES[tabName]();
    // resize() depois que o pane ficou visível (display:block já aplicado
    // acima) — sem isso, o Chart.js mede o canvas ainda invisível.
    requestAnimationFrame(() => chart?.resize());
    if(persist){
      try{ localStorage.setItem(KPI_CHART_TAB_STORAGE_KEY, tabName); }catch{ /* localStorage indisponível — não é crítico, só não persiste a preferência */ }
    }
  }

  tabs.forEach(tab => {
    tab.addEventListener('click', () => activate(tab.dataset.chartTab));
  });

  let initialTab = 'evolucao';
  try{
    const saved = localStorage.getItem(KPI_CHART_TAB_STORAGE_KEY);
    if(saved && KPI_CHART_INSTANCES[saved]) initialTab = saved;
  }catch{ /* localStorage indisponível — usa o padrão 'evolucao' */ }
  activate(initialTab, { persist:false });
}
setupKpiChartTabs();

/* ===== NAVEGAÇÃO ENTRE VIEWS (menu lateral) ===== */
const VIEWS_WITH_FILTERS = new Set(['overview','table','performance','clients','faturamento']);
// FASE 4 · ETAPA 1 (UX/UI) — ordem visual da navegação (usada só para decidir
// a direção do deslize entre telas; não é usada para nenhuma outra lógica).
const NAV_ORDER = ['overview','add','table','performance','compare','clients','faturamento','users','profile','config'];
let currentView = 'overview';

/* ===== FASE 6 (UX/UI) — indicador "blob" líquido preso ao ícone =====
   Antes o blob esticava para cobrir a largura/altura inteira do
   .nav-item ativo — na prática virava um retângulo atrás de ícone +
   rótulo. Agora ele tem tamanho fixo (mesmo tamanho do ícone, ver CSS)
   e se posiciona centralizado exatamente sobre o .nav-icon ativo,
   medindo a posição real do ícone (getBoundingClientRect) em vez de
   um valor fixo — funciona em qualquer breakpoint sem duplicar números.
   O resultado: ao trocar de view, o ícone clicado "nasce" de dentro da
   mesma massa de cor em vez de só ganhar uma caixa por trás. */
function updateNavBlob(instant) {
  const nav = document.getElementById('sideNav');
  const blob = document.getElementById('navBlob');
  const active = nav?.querySelector('.nav-item.active');
  const icon = active?.querySelector('.nav-icon');
  if (!nav || !blob || !active || !icon) return;
  if (instant) blob.style.transition = 'none';
  const navRect = nav.getBoundingClientRect();
  const iconRect = icon.getBoundingClientRect();
  const centerY = (iconRect.top - navRect.top) + iconRect.height / 2;
  blob.style.transform = `translate(-50%, calc(${centerY}px - 50%))`;
  blob.classList.add('is-ready');
  if (instant) {
    /* eslint-disable-next-line sonarjs/void-use --
       Falso positivo. Ler offsetHeight e descartar o valor é o idioma padrão
       para forçar um reflow síncrono; o `void` existe justamente para deixar
       explícito que o valor é descartado de propósito. Sem ele, a expressão
       solta seria acusada por no-unused-expressions, e atribuir a uma
       variável seria acusada por no-unused-vars. */
    void blob.offsetHeight; // força reflow antes de reativar a transição
    blob.style.transition = '';
  }
}

let suppressHashUpdate = false; // true durante navegação originada do próprio hash (evita pushState duplicado/loop)

function goToView(view){
  const fromIdx = NAV_ORDER.indexOf(currentView);
  const toIdx = NAV_ORDER.indexOf(view);
  document.body.classList.toggle('dir-back', toIdx > -1 && fromIdx > -1 && toIdx < fromIdx);
  currentView = view;
  document.querySelectorAll('.nav-item').forEach(b=>b.classList.toggle('active', b.dataset.view===view));
  document.querySelectorAll('.view').forEach(v=>v.classList.toggle('active', v.id==='view-'+view));
  const filtersBlock = document.getElementById('filtersBlock');
  filtersBlock.style.display = VIEWS_WITH_FILTERS.has(view) ? '' : 'none';
  window.scrollTo({top:0, behavior:'smooth'});
  updateNavBlob();
  // Especificação de melhorias, item 3 — verifica o gate de manutenção
  // POR VIEW só ao entrar na view "add" (a única com esse mecanismo hoje).
  // Chamada assíncrona e "fire and forget" de propósito: a view já troca
  // de exibição normalmente acima; se estiver bloqueada, o banner é
  // injetado dentro dela um instante depois, sem travar a navegação.
  if(view === 'add'){
    const cu = window.getCurrentUser ? window.getCurrentUser() : { uid:null, data:null };
    checkViewMaintenanceGate('add', 'view-add', cu.data ? { uid: cu.uid, ...cu.data } : null).catch((err) => {
      console.error('[cotacoes] Erro ao verificar gate de manutenção da view Adicionar:', err);
    });
  }
  // Especificação de melhorias, item 1.2 — mantém a URL sincronizada com a
  // view atual (hash routing, sem reload), para o botão "voltar" do
  // navegador funcionar entre views e para colar/recarregar a URL abrir
  // na view certa. pushState (não replaceState) para cada troca ficar no
  // histórico do navegador — suppressHashUpdate evita loop quando é o
  // próprio hash (popstate ou rota inicial) que está chamando goToView().
  if(!suppressHashUpdate && location.hash !== '#'+view){
    history.pushState({ bdgView: view }, '', '#'+view);
  }
}
document.querySelectorAll('.nav-item').forEach(btn=>{
  btn.addEventListener('click', ()=>goToView(btn.dataset.view));
});
updateNavBlob(true);
window.addEventListener('resize', () => updateNavBlob(true));
window.addEventListener('load', () => updateNavBlob(true));

/* ============================================================
   ROTEAMENTO POR URL (Especificação de melhorias, item 1.2)
   ============================================================
   Hash routing (#table, #faturamento, ...), não path routing: o app é
   100% estático (Firebase Hosting, sem backend próprio — ver auditoria
   técnica), então rotas por pathname exigiriam configurar reescrita de
   rota no firebase.json, o que é uma mudança de infraestrutura fora do
   escopo desta correção. Hash funciona com zero configuração de servidor
   adicional.

   Sequência de boot: o hash é lido AQUI (no carregamento do módulo,
   antes de qualquer resposta do Firebase) e só GUARDADO como intenção
   (pendingRouteFromHash) — nunca aplicado direto, porque uma view como
   "add" não deve renderizar antes de auth.js confirmar que o usuário
   está aprovado. window.applyPendingRouteFromHash() é chamada por
   auth.js exatamente nos dois pontos (onNormalAccess/onBypass) em que
   o app hoje já libera #appRoot — ver nota de integração no final deste
   bloco. Se o hash não corresponder a nenhuma view conhecida, cai em
   silêncio para 'overview' (colar uma URL errada não deveria mostrar
   erro nenhum ao usuário). */
function readViewFromHash(){
  const raw = (location.hash || '').replace(/^#/, '').trim();
  return NAV_ORDER.includes(raw) ? raw : null;
}
let pendingRouteFromHash = readViewFromHash(); // capturado agora, aplicado só após a autenticação resolver

window.applyPendingRouteFromHash = function applyPendingRouteFromHash(){
  const target = pendingRouteFromHash || readViewFromHash() || 'overview';
  pendingRouteFromHash = null;
  suppressHashUpdate = true;
  goToView(target);
  suppressHashUpdate = false;
  // Garante que a URL reflita a view final mesmo quando o hash original
  // era inválido/ausente e caiu no padrão 'overview'.
  if(location.hash !== '#'+target) history.replaceState({ bdgView: target }, '', '#'+target);
};

// Botão "voltar"/"avançar" do navegador: navega para a view do hash sem
// adicionar nova entrada no histórico (senão cada "voltar" empurraria
// uma entrada nova, invertendo o próprio botão voltar).
window.addEventListener('popstate', () => {
  const view = readViewFromHash();
  if(!view) return;
  suppressHashUpdate = true;
  goToView(view);
  suppressHashUpdate = false;
});
// Edição manual do hash na barra de endereço (ex.: apagar "#table" e
// digitar "#faturamento" sem dar Enter numa URL nova) nem sempre dispara
// popstate — hashchange cobre esse caso também, com a mesma lógica.
window.addEventListener('hashchange', () => {
  const view = readViewFromHash();
  if(!view || view === currentView) return;
  suppressHashUpdate = true;
  goToView(view);
  suppressHashUpdate = false;
});

/* ===== SERVICE WORKER (cache offline, sem prompt de instalação) ===== */
if ('serviceWorker' in navigator) {
  window.addEventListener('load', () => {
    navigator.serviceWorker.register('sw.js').catch(err => console.warn('SW falhou:', err));
  });
}
