/* eslint-disable security/detect-object-injection --
   Decisão registrada (Especificação de melhorias, item 3): os acessos
   dinâmicos a objeto neste arquivo (viewGateState[viewName]) usam sempre
   viewName vindo de uma chamada interna já validada (checkViewMaintenance
   Gate/onClickViewGateBypass, nunca de entrada bruta do usuário). Mesmo
   padrão de falso positivo já registrado em
   route-map.js/utils.js/cotacoes.js/geo-data.js. */
/* ==========================================================================
   maintenance.js — BDG · Controle de Cotações (Samsung SDS Logistics)
   ==========================================================================
   Artifact 05/08 — Refatoração UI/UX (Fase 6) — Tela Global de Manutenção

   RESPONSABILIDADE: decidir se o app deve mostrar a Tela Global de
   Manutenção (bloqueando o uso) ou liberar o acesso normalmente, com uma
   exceção para usuários Administradores ("Acessar mesmo assim").

   DECISÃO DE ARQUITETURA — POR QUE NÃO É UM MECANISMO NOVO:
   O projeto BDG já possui um sistema de papéis (roles) funcional e testado
   em produção, implementado em auth.js: cada usuário tem um registro em
   users/{uid} no Realtime Database, com os campos:
     - status: 'pending' | 'approved' | 'rejected'
     - role:   'admin' | 'user'
   Esse é exatamente o "Método 2 (Firestore/RTDB) — users/{uid}.role" que a
   especificação desta fase propõe como alternativa ao Custom Claims.
   Reimplementar isso como um mecanismo paralelo duplicaria lógica já
   validada (a Fase 1 do projeto já testou esse fluxo com aprovação real de
   usuários) e criaria duas fontes de verdade sobre "quem é admin" — um
   risco real de inconsistência. Este arquivo, portanto, LÊ o mesmo campo
   role já existente, não cria um novo.

   SOBRE CUSTOM CLAIMS (Método 1): não foi adotado nesta fase porque exige
   Firebase Admin SDK rodando em um backend (Cloud Functions ou servidor
   próprio) para definir a claim — o projeto não tem hoje nenhum backend
   próprio (é 100% client-side + Realtime Database, hospedado como site
   estático no Firebase Hosting). Adotar Custom Claims exigiria criar e
   manter Cloud Functions, o que sai do escopo de "zero infraestrutura
   nova" desta fase. Fica registrado como opção válida para uma fase futura
   se o projeto vier a precisar de um backend por outro motivo.

   COMO ESTE MÓDULO SE CONECTA AO RESTANTE DO APP:
   auth.js já expõe, ao final de onAuthStateChanged, o momento exato em que
   o usuário está autenticado E aprovado (status === 'approved'), antes de
   liberar #appRoot. Este módulo escuta esse mesmo evento via uma função
   window.checkMaintenanceGate(userData), que auth.js deve chamar naquele
   ponto (uma única linha nova em auth.js — ver comentário de integração
   ao final deste arquivo). Este módulo NUNCA inicializa uma segunda
   instância do Firebase: reaproveita getApps()/getApp(), exatamente como
   auth.js já faz, para não gerar o erro "Firebase App named '[DEFAULT]'
   already exists".
   ========================================================================== */

import { getApps, getApp } from "https://www.gstatic.com/firebasejs/12.16.0/firebase-app.js";
import { getDatabase, ref as dbRef, get as dbGet } from "https://www.gstatic.com/firebasejs/12.16.0/firebase-database.js";

/* CORREÇÃO CRÍTICA — nó errado no Realtime Database ----------------------
   Este arquivo lia/gravava em "config/maintenanceMode" e
   "config/viewMaintenance/{view}", mas as regras de segurança do projeto
   só concedem permissão para o nó "configuracoes" (com "es") — "config"
   nunca teve regra nenhuma, então TODA leitura nesses caminhos era negada
   com PERMISSION_DENIED pelo Realtime Database, sempre, para qualquer
   usuário. O fallback de segurança (nunca bloquear em caso de erro)
   evitava travar o app por causa disso, mas o gate de manutenção nunca
   funcionava de verdade. Corrigido para "configuracoes/maintenanceMode" e
   "configuracoes/viewMaintenance/{view}", batendo com as regras reais:
     "configuracoes": { ".read": "auth != null", ".write": "..." } */

/* ---------- flag de manutenção ----------
   Em produção, o ideal é este valor vir do Realtime Database (ex.:
   configuracoes/maintenanceMode: true/false), para o admin poder ativar/desativar
   sem precisar editar e reimplantar o código. Ficou como constante local
   nesta entrega para ser testável imediatamente; a variante que lê do
   banco está pronta e comentada logo abaixo — trocar é uma linha. */
const MAINTENANCE_MODE_STATIC = true;

/**
 * Lê o modo de manutenção diretamente do Realtime Database, em
 * configuracoes/maintenanceMode. Retorna false (não bloqueia nada) se o nó não
 * existir ou se houver qualquer erro de leitura — comportamento seguro
 * por padrão: uma falha de rede/permissão nunca deve travar o app
 * inteiro fora de um bloqueio deliberado.
 *
 * @param {import('https://www.gstatic.com/firebasejs/12.16.0/firebase-database.js').Database} rtdb
 * @returns {Promise<boolean>}
 */
/* CORREÇÃO CRÍTICA DE LOGIN TRAVADO ---------------------------------------
   dbGet() do Realtime Database NÃO rejeita quando a conexão está instável
   ou o cliente está offline — a Promise simplesmente nunca resolve. Como
   checkMaintenanceGate() faz `await` nesta leitura ANTES de chamar
   onNormalAccess(), uma leitura pendurada deixava o usuário preso na tela
   de carregamento indefinidamente, sem erro nenhum no console.
   withTimeout() garante que o app SEMPRE libera: se o banco não responder
   em MAINTENANCE_READ_TIMEOUT_MS, assume "sem bloqueio" (mesma decisão de
   segurança já adotada no catch abaixo — falha de leitura nunca deve
   impedir o acesso de quem tem credencial válida). */
const MAINTENANCE_READ_TIMEOUT_MS = 3500;

function withTimeout(promise, ms, fallbackValue, label) {
  return Promise.race([
    promise,
    new Promise((resolve) => setTimeout(() => {
      console.warn(`[maintenance] Leitura de ${label} excedeu ${ms}ms — liberando o app sem bloqueio (fallback de segurança).`);
      resolve(fallbackValue);
    }, ms)),
  ]);
}

async function readMaintenanceFlagFromDatabase(rtdb) {
  try {
    const snap = await withTimeout(
      dbGet(dbRef(rtdb, 'configuracoes/maintenanceMode')),
      MAINTENANCE_READ_TIMEOUT_MS,
      null,
      'configuracoes/maintenanceMode'
    );
    if (snap === null) return false; // timeout — assume modo normal
    return snap.exists() ? snap.val() === true : false;
  } catch (err) {
    console.warn('[maintenance] Não foi possível ler configuracoes/maintenanceMode — assumindo modo normal (sem bloqueio) por segurança:', err);
    return false;
  }
}

/* ---------- construção do DOM da tela de manutenção (uma única vez) ---------- */
function buildMaintenanceDom() {
  if (document.getElementById('maintenanceGate')) return; // já existe, não duplica

  const el = document.createElement('div');
  el.id = 'maintenanceGate';
  el.className = 'maintenance-gate';
  el.style.display = 'none';
  el.innerHTML = `
    <div class="maintenance-gate-card">
      <div class="maintenance-gate-icon">🚧</div>
      <h2>Sistema em manutenção</h2>
      <p>Estamos realizando melhorias no BDG. O acesso será restabelecido em breve. Pedimos desculpas pelo transtorno.</p>
      <button type="button" class="btn btn-ghost maintenance-admin-btn" id="btnMaintenanceAdminBypass">
        Entrar como Administrador
      </button>
      <div class="maintenance-gate-msg" id="maintenanceGateMsg"></div>
    </div>
  `;
  document.body.appendChild(el);

  document.getElementById('btnMaintenanceAdminBypass').addEventListener('click', onClickAdminBypass);
}

/* ---------- estado interno ---------- */
let currentRtdb = null;
let currentUserData = null;
let onBypassGranted = null; // callback chamado quando o bypass do gate GLOBAL é aceito

/**
 * Revalida o papel do usuário diretamente no Realtime Database (não confia
 * apenas no valor já carregado em memória, para o caso de o role ter sido
 * alterado por outro admin desde o login). Compartilhada entre o gate
 * GLOBAL (onClickAdminBypass) e o gate POR VIEW (Especificação de
 * melhorias, item 3) — mesma checagem, um único lugar, para as duas
 * nunca divergirem uma da outra.
 *
 * @param {string} uid
 * @returns {Promise<{ok: true, data: object} | {ok: false, reason: string}>}
 */
async function revalidateAdminRole(uid) {
  if (!uid) return { ok: false, reason: 'Nenhum usuário autenticado no momento.' };
  const snap = await dbGet(dbRef(currentRtdb, 'users/' + uid));
  const freshData = snap.exists() ? snap.val() : null;
  if (!freshData) return { ok: false, reason: 'Não foi possível confirmar seu cadastro no banco de dados.' };
  if (freshData.role !== 'admin') return { ok: false, reason: 'Apenas administradores podem acessar durante a manutenção.' };
  if (freshData.status !== 'approved') return { ok: false, reason: 'Seu cadastro de administrador ainda não está aprovado.' };
  return { ok: true, data: freshData };
}

/**
 * Trata o clique em "Entrar como Administrador" do gate GLOBAL de
 * manutenção — revalida via revalidateAdminRole() e libera o app inteiro
 * se aprovado.
 */
async function onClickAdminBypass() {
  const btn = document.getElementById('btnMaintenanceAdminBypass');
  const msgEl = document.getElementById('maintenanceGateMsg');
  msgEl.textContent = '';
  btn.disabled = true;
  const originalLabel = btn.textContent;
  btn.textContent = 'Verificando...';

  try {
    const result = await revalidateAdminRole(currentUserData && currentUserData.uid);
    if (!result.ok) {
      msgEl.textContent = result.reason;
      msgEl.classList.add('is-error');
      return;
    }

    // Bypass concedido: esconde a tela de manutenção e libera o app.
    document.getElementById('maintenanceGate').style.display = 'none';
    if (typeof onBypassGranted === 'function') onBypassGranted(result.data);
    applyBetaFeatureLocks(); // funcionalidades ainda em desenvolvimento continuam desabilitadas mesmo para o admin
  } catch (err) {
    console.error('[maintenance] Erro ao verificar bypass de administrador:', err);
    msgEl.textContent = 'Erro ao verificar permissão: ' + err.message;
    msgEl.classList.add('is-error');
  } finally {
    btn.disabled = false;
    btn.textContent = originalLabel;
  }
}

/**
 * Desabilita, via atributo disabled (nunca removendo do DOM), os elementos
 * marcados como "ainda em desenvolvimento" — convenção: qualquer elemento
 * com o atributo data-beta-feature="true" no HTML. Isso cumpre o requisito
 * de "não remover funcionalidades inacabadas, apenas impedir sua utilização"
 * também para o Administrador durante o bypass de manutenção, já que estar
 * em manutenção não significa que toda funcionalidade beta já está pronta
 * para teste.
 */
function applyBetaFeatureLocks() {
  document.querySelectorAll('[data-beta-feature="true"]').forEach((el) => {
    el.disabled = true;
    el.setAttribute('aria-disabled', 'true');
    if (!el.title) el.title = 'Funcionalidade em desenvolvimento — ainda não disponível.';
  });
}

/**
 * Ponto de entrada principal, chamado por auth.js no momento em que o
 * usuário está autenticado e aprovado (ver nota de integração no final
 * deste arquivo). Decide se mostra a Tela Global de Manutenção ou libera
 * o app normalmente.
 *
 * @param {object} userData  Dados do usuário já carregados por auth.js
 *   (deve incluir ao menos uid, role, status — mesmo shape gravado em
 *   users/{uid} no Realtime Database).
 * @param {() => void} onNormalAccess  Chamado se NÃO há bloqueio de
 *   manutenção (fluxo normal do app deve continuar como já funcionava).
 * @param {() => void} onBypass  Chamado se o usuário administrador clicou
 *   em "Entrar como Administrador" e foi validado com sucesso.
 */
export async function checkMaintenanceGate(userData, onNormalAccess, onBypass) {
  buildMaintenanceDom();

  const apps = getApps();
  if (!apps.length) {
    console.error('[maintenance] Nenhum app Firebase inicializado ainda — auth.js deveria ter inicializado antes de chamar checkMaintenanceGate.');
    onNormalAccess(); // segurança: não bloqueia o app por um problema de ordem de inicialização
    return;
  }
  currentRtdb = getDatabase(getApp());
  currentUserData = userData;
  onBypassGranted = onBypass;

  // Lê o flag do banco primeiro (fonte de verdade em produção); cai para
  // a constante estática apenas se a leitura falhar por completo — ver
  // readMaintenanceFlagFromDatabase() para o tratamento de erro.
  let maintenanceActive;
  try {
    maintenanceActive = await readMaintenanceFlagFromDatabase(currentRtdb);
  } catch {
    maintenanceActive = MAINTENANCE_MODE_STATIC;
  }

  if (!maintenanceActive) {
    document.getElementById('maintenanceGate').style.display = 'none';
    onNormalAccess();
    return;
  }

  // Modo de manutenção ativo: usuários Admin já aprovados passam direto
  // (sem precisar clicar no bypass) — a exigência de clique é reservada
  // para deixar claro ao usuário comum que ele está sendo bloqueado, não
  // para dificultar o acesso de quem já é administrador legítimo.
  if (userData && userData.role === 'admin' && userData.status === 'approved') {
    document.getElementById('maintenanceGate').style.display = 'none';
    onNormalAccess();
    applyBetaFeatureLocks();
    return;
  }

  // Usuário comum durante manutenção: mostra a tela de bloqueio.
  document.getElementById('appRoot').style.display = 'none';
  document.getElementById('maintenanceGate').style.display = 'flex';
}

/* ============================================================
   GATE DE MANUTENÇÃO POR VIEW (Especificação de melhorias, item 3)
   ============================================================
   Diferente do gate GLOBAL acima (bloqueia o app inteiro), este bloqueia
   só uma view específica — usado hoje pela view "Adicionar" (view-add),
   que tinha um overlay estático sem bypass nenhum. Decisão registrada
   com o usuário: reaproveitar o MESMO mecanismo de autorização do gate
   global (role 'admin' via revalidateAdminRole(), acima) em vez de criar
   uma senha nova — evita duas fontes de verdade sobre "quem é admin".

   Fonte do flag: configuracoes/viewMaintenance/{viewName} no Realtime Database
   (mesmo padrão de "ligar/desligar sem redeploy" que configuracoes/maintenanceMode
   já usa para o gate global) — permite bloquear só "add" sem afetar as
   demais views. Ausência do nó ou erro de leitura = não bloqueia (mesmo
   princípio de segurança do gate global: falha de rede/permissão nunca
   deve travar uma view que deveria estar liberada). */
async function readViewMaintenanceFlag(rtdb, viewName) {
  try {
    // Mesmo fallback de timeout do gate global (ver withTimeout acima):
    // uma leitura pendurada nunca deve deixar a view inacessível.
    const snap = await withTimeout(
      dbGet(dbRef(rtdb, 'configuracoes/viewMaintenance/' + viewName)),
      MAINTENANCE_READ_TIMEOUT_MS,
      null,
      'configuracoes/viewMaintenance/' + viewName
    );
    if (snap === null) return false; // timeout — assume view liberada
    return snap.exists() ? snap.val() === true : false;
  } catch (err) {
    console.warn('[maintenance] Não foi possível ler configuracoes/viewMaintenance/' + viewName + ' — assumindo view liberada por segurança:', err);
    return false;
  }
}

const viewGateState = {}; // viewName -> { locked: bool, userData, onUnlock, msgEl, btn }

function buildViewGateDom(viewName, containerEl) {
  const existing = containerEl.querySelector('.view-gate[data-view-gate="' + viewName + '"]');
  if (existing) return existing;

  const el = document.createElement('div');
  el.className = 'view-gate';
  el.dataset.viewGate = viewName;
  el.innerHTML = `
    <div class="view-gate-card">
      <div class="view-gate-icon">🚧</div>
      <h3>Funcionalidade em manutenção</h3>
      <p>Esta funcionalidade está temporariamente em manutenção. Em uma próxima versão ela será disponibilizada novamente com melhorias.</p>
      <button type="button" class="btn btn-ghost view-gate-admin-btn" data-view-gate-bypass="${viewName}">
        Entrar como Administrador
      </button>
      <div class="view-gate-msg" data-view-gate-msg="${viewName}"></div>
    </div>
  `;
  containerEl.prepend(el);

  el.querySelector('[data-view-gate-bypass]').addEventListener('click', () => onClickViewGateBypass(viewName));
  return el;
}

async function onClickViewGateBypass(viewName) {
  const state = viewGateState[viewName];
  if (!state) return;
  const btn = document.querySelector('[data-view-gate-bypass="' + viewName + '"]');
  const msgEl = document.querySelector('[data-view-gate-msg="' + viewName + '"]');
  if (!btn || !msgEl) return;
  msgEl.textContent = '';
  msgEl.classList.remove('is-error');
  btn.disabled = true;
  const originalLabel = btn.textContent;
  btn.textContent = 'Verificando...';

  try {
    const result = await revalidateAdminRole(state.userData && state.userData.uid);
    if (!result.ok) {
      msgEl.textContent = result.reason;
      msgEl.classList.add('is-error');
      return;
    }
    // Bypass concedido: some com o bloqueio SÓ desta view — nenhuma
    // outra view é afetada (diferente do bypass global, que libera o
    // app inteiro).
    const gateEl = document.querySelector('.view-gate[data-view-gate="' + viewName + '"]');
    if (gateEl) gateEl.remove();
    state.locked = false;
    if (typeof state.onUnlock === 'function') state.onUnlock(result.data);
  } catch (err) {
    console.error('[maintenance] Erro ao verificar bypass de view (' + viewName + '):', err);
    msgEl.textContent = 'Erro ao verificar permissão: ' + err.message;
    msgEl.classList.add('is-error');
  } finally {
    btn.disabled = false;
    btn.textContent = originalLabel;
  }
}

/**
 * Verifica se uma view específica deve ficar bloqueada por manutenção e,
 * se sim, injeta o banner de bloqueio (com bypass de admin) dentro do
 * elemento da view — ao contrário do gate global, o restante do app
 * continua acessível.
 *
 * @param {string} viewName - nome usado em configuracoes/viewMaintenance/{viewName}
 * @param {string} containerId - id do elemento da view (ex.: 'view-add')
 * @param {object} userData - mesmo shape usado em checkMaintenanceGate
 * @returns {Promise<boolean>} true se a view está (ou ficou) bloqueada
 */
export async function checkViewMaintenanceGate(viewName, containerId, userData) {
  const containerEl = document.getElementById(containerId);
  if (!containerEl) return false;

  const apps = getApps();
  if (!apps.length) {
    console.error('[maintenance] Nenhum app Firebase inicializado ainda — checkViewMaintenanceGate chamada cedo demais.');
    return false; // segurança: não bloqueia a view por um problema de ordem de inicialização
  }
  if (!currentRtdb) currentRtdb = getDatabase(getApp());

  // Admin já aprovado nunca vê o banner desta view — mesmo princípio do
  // gate global: a exigência de bypass existe para deixar claro ao
  // usuário comum que ele está bloqueado, não para dificultar o acesso
  // de quem já é administrador legítimo.
  if (userData && userData.role === 'admin' && userData.status === 'approved') {
    return false;
  }

  const locked = await readViewMaintenanceFlag(currentRtdb, viewName);
  if (!locked) return false;

  viewGateState[viewName] = {
    locked: true,
    userData,
    onUnlock: () => { /* nenhuma ação adicional necessária: remover o banner já libera a view */ },
  };
  buildViewGateDom(viewName, containerEl);
  return true;
}

/* ============================================================
   NOTA DE INTEGRAÇÃO COM auth.js (ação necessária, fora deste arquivo)
   ============================================================
   Em auth.js, dentro de onAuthStateChanged, no trecho onde o usuário é
   considerado aprovado (comentário "// aprovado" no código atual, logo
   antes de "$('appRoot').style.display = 'block';"), envolver a liberação
   do app com uma chamada a checkMaintenanceGate:

     import { checkMaintenanceGate } from './maintenance.js?v=20260802a';
     // ...
     // aprovado
     checkMaintenanceGate(
       { uid: user.uid, ...data },
       () => { // onNormalAccess — o mesmo bloco que já existia
         $('authGate').style.display = 'none';
         $('pendingGate').style.display = 'none';
         $('appRoot').style.display = 'block';
         if(window.showAppLoading) window.showAppLoading(data.username || data.email || '');
         startPresence(user.uid);
         logActivity(user.uid, data.username, 'Login realizado');
       },
       () => { // onBypass — chamado só quando um admin clica "Entrar como Administrador"
         $('authGate').style.display = 'none';
         $('pendingGate').style.display = 'none';
         $('appRoot').style.display = 'block';
       }
     );

   Esta é a ÚNICA mudança necessária em auth.js — nenhuma outra linha do
   arquivo precisa ser tocada. O comportamento quando maintenanceMode for
   false (ou o nó não existir no banco) é idêntico ao que já existe hoje.
   ========================================================================== */
