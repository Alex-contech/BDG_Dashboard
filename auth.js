/* ==========================================================================
   auth.js — BDG · Controle de Cotações (Samsung SDS Logistics)
   ==========================================================================
   MÓDULO DE SESSÃO. Responde a uma única pergunta, o tempo todo:
   "qual tela o usuário deve estar vendo agora?"

   Reescrito do zero. O arquivo anterior tinha 737 linhas e acumulava quatro
   responsabilidades distintas (sessão, perfil, gestão de usuários e
   infraestrutura do Firebase). Agora:

     firebase-core.js  → infraestrutura (Firebase, trinco de rede, prazos)
     auth.js           → sessão: entrar, cadastrar, sair, decidir a tela
     account.js        → telas de Perfil e Usuários, presença

   ---------------------------------------------------------------------------
   O QUE MUDOU DE FATO, E POR QUE O LOGIN TRAVAVA
   ---------------------------------------------------------------------------
   1. MÁQUINA DE ESTADOS ÚNICA.
      Antes, onAuthStateChanged tinha cinco saídas diferentes, e cada uma
      repetia à mão as mesmas cinco linhas de manipulação de tela (esconder
      isto, mostrar aquilo, liberar a trava, esconder o carregamento). Bastava
      uma saída esquecer uma linha para o aplicativo ficar num estado
      impossível — nenhuma tela visível, carregamento eterno. É o que
      acontecia no caminho do portão de manutenção, que nunca liberava a
      trava nem escondia a tela de carregamento.

      Agora existe UM lugar que decide o que está na tela (applyState) e UM
      caminho para mudar de estado (transitionTo). Todo estado é total: ele
      diz o que aparece, o que some e o que é liberado. Não há saída parcial.

   2. CÃO DE GUARDA DO BOOT.
      Nenhuma promessa individual precisa mais ser confiável. Ao entrar em
      "carregando", um cronômetro é armado; se em BOOT_WATCHDOG_MS o
      aplicativo não tiver chegado a um estado final, ele é levado à tela de
      erro com uma mensagem que diz o que falhou e o botão Sair disponível.
      Ficar preso na tela de carregamento deixa de ser possível por
      construção, e não por cada chamada ter lembrado do seu próprio prazo.

   3. FIM DA TRAVA COMPARTILHADA.
      O envio do formulário devolve o seu trinco sempre, no finally. O boot
      pega o seu próprio trinco. Não existe mais a passagem de trava de uma
      função para outra — que era a corrida descrita em firebase-core.js.
   ========================================================================== */

import {
  createUserWithEmailAndPassword, signInWithEmailAndPassword, onAuthStateChanged,
  sendPasswordResetEmail, updateProfile, signOut,
} from 'https://www.gstatic.com/firebasejs/12.16.0/firebase-auth.js';
import {
  ref as dbRef, set as dbSet, get as dbGet,
} from 'https://www.gstatic.com/firebasejs/12.16.0/firebase-database.js';

import {
  auth, rtdb, acquireNetworkLock, withDeadline, readUserProfile, logActivity, TimeoutError,
  appCheckDisabled,
} from './firebase-core.js?v=20260802a';
import { checkPasswordRules } from './core.js?v=20260802a';
import { checkMaintenanceGate } from './maintenance.js?v=20260802a';
import { mountAccountViews, unmountAccountViews, startPresence, stopPresence } from './account.js?v=20260802a';

const LOG_PREFIX = '[auth]';

/* Prazos. Todos generosos o bastante para rede móvel ruim e curtos o
   bastante para o usuário não achar que o aplicativo morreu. */
const SIGN_IN_TIMEOUT_MS = 15000;
const SIGN_UP_TIMEOUT_MS = 20000;
const PASSWORD_RESET_TIMEOUT_MS = 15000;
const MAINTENANCE_GATE_TIMEOUT_MS = 8000;
/* Precisa ser maior que a soma dos prazos das etapas do boot
   (leitura do perfil 10s + 5s, portão de manutenção 8s = 23s), senão o cão de
   guarda dispararia por cima de uma etapa que ainda ia responder com uma
   mensagem melhor. Ele é a última linha de defesa, não o prazo normal. */
const BOOT_WATCHDOG_MS = 28000;

/* --------------------------------------------------------------------------
   SESSÃO
   --------------------------------------------------------------------------
   Objeto único, passado por referência para account.js. Assim aquele módulo
   pode atualizar a foto ou o nome sem precisar chamar de volta este arquivo
   (o que criaria dependência circular entre os dois). */
const session = { uid: null, data: null };

/* Ponte com cotacoes.js, que lê o usuário logado para registrar quem fez
   cada alteração. Definida já na carga do módulo para nunca ser chamada
   antes de existir. */
window.getCurrentUser = () => ({ uid: session.uid, data: session.data });
window.logActivity = logActivity;

/* --------------------------------------------------------------------------
   ESTADOS
   -------------------------------------------------------------------------- */
const SessionState = Object.freeze({
  BOOTING: 'BOOTING',
  SIGNED_OUT: 'SIGNED_OUT',
  LOADING: 'LOADING',
  BLOCKED: 'BLOCKED',
  MAINTENANCE: 'MAINTENANCE',
  READY: 'READY',
});

const BlockedReason = Object.freeze({
  PENDING: {
    icon: '⏳',
    title: 'Cadastro pendente de aprovação',
    text: 'Sua conta foi criada. Um administrador precisa aprovar seu acesso antes de você entrar no painel.',
  },
  REJECTED: {
    icon: '✕',
    title: 'Cadastro não aprovado',
    text: 'Seu acesso não foi aprovado pelo administrador. Fale com ele para saber o motivo.',
  },
});

function blockedByError(detail) {
  return {
    icon: '⚠️',
    title: 'Não foi possível carregar seu cadastro',
    text: `${detail} Toque em Sair, confira sua conexão e entre de novo. Se continuar, avise o administrador com esta mensagem.`,
  };
}

let currentState = SessionState.BOOTING;
let bootWatchdogTimer = null;

/* --------------------------------------------------------------------------
   RENDERIZAÇÃO — o único lugar que mexe nas telas
   -------------------------------------------------------------------------- */
const byId = (id) => document.getElementById(id);

function setDisplay(id, value) {
  const el = byId(id);
  if (el) el.style.display = value;
}

function setText(id, value) {
  const el = byId(id);
  if (el) el.textContent = value;
}

/* As três telas de nível superior são mutuamente exclusivas. Passar por aqui
   garante isso — nenhuma combinação inválida sobrevive a uma transição. */
function showTopLevelScreen(screen) {
  setDisplay('authGate', screen === 'authGate' ? 'flex' : 'none');
  setDisplay('pendingGate', screen === 'pendingGate' ? 'flex' : 'none');
  setDisplay('appRoot', screen === 'appRoot' ? 'block' : 'none');
}

function hideLoadingOverlays() {
  if (typeof window.hideAppLoadingImmediately === 'function') {
    window.hideAppLoadingImmediately();
  }
}

function applyState(state, payload) {
  switch (state) {
    case SessionState.SIGNED_OUT:
      showTopLevelScreen('authGate');
      document.body.classList.remove('app-ready');
      hideLoadingOverlays();
      resetAuthForms();
      break;

    case SessionState.LOADING:
      /* Nenhuma tela muda aqui: o splash ou a tela de carregamento pós-login
         já estão visíveis. O estado existe para o cão de guarda ter o que
         vigiar. */
      break;

    case SessionState.BLOCKED:
      showTopLevelScreen('pendingGate');
      document.body.classList.remove('app-ready');
      hideLoadingOverlays();
      setText('pendingIcon', payload.icon);
      setText('pendingTitle', payload.title);
      setText('pendingText', payload.text);
      break;

    case SessionState.MAINTENANCE:
      /* maintenance.js desenha e controla a própria sobreposição. Aqui só
         garantimos que nada fique por cima dela nem rodando por baixo —
         era justamente o que faltava antes. */
      showTopLevelScreen('none');
      document.body.classList.remove('app-ready');
      hideLoadingOverlays();
      break;

    case SessionState.READY:
      showTopLevelScreen('appRoot');
      document.body.classList.add('app-ready');
      break;

    default:
      console.warn(`${LOG_PREFIX} Estado desconhecido:`, state);
      break;
  }
}

function transitionTo(state, payload = {}) {
  currentState = state;
  clearBootWatchdog();
  applyState(state, payload);
  if (state === SessionState.LOADING) armBootWatchdog();
}

function isTerminalState() {
  return currentState !== SessionState.LOADING && currentState !== SessionState.BOOTING;
}

/* --------------------------------------------------------------------------
   CÃO DE GUARDA DO BOOT
   -------------------------------------------------------------------------- */
function clearBootWatchdog() {
  if (bootWatchdogTimer) {
    clearTimeout(bootWatchdogTimer);
    bootWatchdogTimer = null;
  }
}

function armBootWatchdog() {
  bootWatchdogTimer = setTimeout(() => {
    bootWatchdogTimer = null;
    if (isTerminalState()) return;
    console.error(`${LOG_PREFIX} O boot passou de ${BOOT_WATCHDOG_MS} ms sem concluir. Levando o aplicativo à tela de erro.`);
    transitionTo(
      SessionState.BLOCKED,
      blockedByError('O servidor não concluiu a entrada a tempo.'),
    );
  }, BOOT_WATCHDOG_MS);
}

/* --------------------------------------------------------------------------
   MENSAGENS DO FORMULÁRIO
   -------------------------------------------------------------------------- */
function showAuthError(message) {
  const errorEl = byId('authError');
  const okEl = byId('authOk');
  if (errorEl) {
    errorEl.textContent = message;
    errorEl.classList.add('show');
  }
  if (okEl) okEl.classList.remove('show');
}

function showAuthOk(message) {
  const errorEl = byId('authError');
  const okEl = byId('authOk');
  if (okEl) {
    okEl.textContent = message;
    okEl.classList.add('show');
  }
  if (errorEl) errorEl.classList.remove('show');
}

function hideAuthMessages() {
  byId('authError')?.classList.remove('show');
  byId('authOk')?.classList.remove('show');
}

function resetAuthForms() {
  byId('formLogin')?.reset();
  byId('formSignup')?.reset();
  hideAuthMessages();
}

/* No celular não há console para conferir se o modo de diagnóstico pegou.
   Este aviso aparece na própria tela de acesso quando o App Check está
   desligado, e some assim que o usuário tenta entrar. */
if (appCheckDisabled) {
  showAuthOk('Modo de diagnóstico: App Check desligado neste aparelho. Para religar, abra a página com ?appcheck=on');
}

/* Map em vez de objeto literal: a chave vem de fora (err.code, definido pelo
   SDK do Firebase), e Map não expõe o protótipo a uma chave inesperada. */
const AUTH_ERROR_MESSAGES = new Map([
  ['auth/email-already-in-use', 'Este e-mail já tem cadastro. Entre pela aba Entrar ou use "Esqueci minha senha".'],
  ['auth/invalid-email', 'Esse e-mail não é válido. Confira se não faltou uma letra ou o @.'],
  ['auth/weak-password', 'O Firebase considerou essa senha fraca. Escolha outra combinação de letras e números.'],
  ['auth/user-not-found', 'Não existe conta com esse e-mail. Confira o endereço ou cadastre-se.'],
  ['auth/wrong-password', 'Senha incorreta. Tente de novo ou use "Esqueci minha senha".'],
  ['auth/invalid-credential', 'E-mail ou senha incorretos. Tente de novo ou use "Esqueci minha senha".'],
  ['auth/too-many-requests', 'Muitas tentativas seguidas. Espere alguns minutos e tente de novo.'],
  ['auth/network-request-failed', 'A conexão falhou ao contatar o servidor. Confira sua internet e tente de novo.'],
]);

function describeAuthError(err) {
  if (err instanceof TimeoutError) {
    return `${err.message} Confira sua conexão e tente de novo.`;
  }
  return AUTH_ERROR_MESSAGES.get(err?.code)
    || `Não foi possível concluir: ${err?.message || 'erro desconhecido'}.`;
}

/* Encapsula o par "desabilita o botão / devolve o botão", que antes era
   repetido e às vezes esquecido em um dos caminhos de erro. */
function useSubmitButton(buttonId, busyLabel) {
  const btn = byId(buttonId);
  const originalLabel = btn ? btn.textContent : '';
  if (btn) {
    btn.disabled = true;
    btn.textContent = busyLabel;
  }
  return () => {
    if (btn) {
      btn.disabled = false;
      btn.textContent = originalLabel;
    }
  };
}

/* --------------------------------------------------------------------------
   REGRAS DE SENHA
   -------------------------------------------------------------------------- */
function paintPasswordRules(password, confirmation) {
  const rules = checkPasswordRules(password, confirmation);
  Object.entries(rules).forEach(([rule, satisfied]) => {
    const el = document.querySelector(`.pwd-rules .r[data-r="${rule}"]`);
    if (el) el.classList.toggle('ok', satisfied);
  });
  return Object.values(rules).every(Boolean);
}

['suPass', 'suPass2'].forEach((id) => {
  byId(id)?.addEventListener('input', () => {
    paintPasswordRules(byId('suPass')?.value || '', byId('suPass2')?.value || '');
  });
});

/* --------------------------------------------------------------------------
   ENTRAR
   -------------------------------------------------------------------------- */
byId('formLogin')?.addEventListener('submit', (event) => {
  event.preventDefault();
  void handleSignIn();
});

async function handleSignIn() {
  hideAuthMessages();
  const email = byId('liEmail')?.value.trim() || '';
  const password = byId('liPass')?.value || '';
  if (!email || !password) {
    showAuthError('Informe e-mail e senha para entrar.');
    return;
  }

  const restoreButton = useSubmitButton('btnLogin', 'Entrando...');
  const releaseLock = acquireNetworkLock('entrar');
  try {
    await withDeadline(
      signInWithEmailAndPassword(auth, email, password),
      SIGN_IN_TIMEOUT_MS,
      'O servidor de autenticação',
    );
    /* Sucesso: onAuthStateChanged assume daqui, com o próprio trinco.
       O botão volta ao normal de qualquer forma no finally — se o boot
       falhar, o usuário volta para um formulário utilizável, não para um
       botão travado em "Entrando...". */
  } catch (err) {
    console.warn(`${LOG_PREFIX} Falha ao entrar:`, err);
    showAuthError(describeAuthError(err));
  } finally {
    releaseLock();
    restoreButton();
  }
}

/* --------------------------------------------------------------------------
   CADASTRAR
   -------------------------------------------------------------------------- */
byId('formSignup')?.addEventListener('submit', (event) => {
  event.preventDefault();
  void handleSignUp();
});

/**
 * Cria a conta e o registro em users/{uid}.
 * Regra de negócio preservada: o primeiro usuário do sistema vira
 * administrador já aprovado; os demais entram como pendentes.
 */
async function createAccount(username, email, password) {
  const credential = await createUserWithEmailAndPassword(auth, email, password);
  await updateProfile(credential.user, { displayName: username });

  const adminFlag = await dbGet(dbRef(rtdb, 'meta/adminExists'));
  const isFirstUser = !adminFlag.exists() || adminFlag.val() !== true;

  await dbSet(dbRef(rtdb, `users/${credential.user.uid}`), {
    username,
    email,
    status: isFirstUser ? 'approved' : 'pending',
    role: isFirstUser ? 'admin' : 'user',
    photoURL: null,
    createdAt: Date.now(),
    lastActive: Date.now(),
  });

  if (isFirstUser) {
    await dbSet(dbRef(rtdb, 'meta/adminExists'), true);
  }

  await logActivity(
    credential.user.uid,
    username,
    isFirstUser ? 'Conta criada (administrador inicial)' : 'Conta criada — aguardando aprovação',
  );
}

async function handleSignUp() {
  hideAuthMessages();
  const username = byId('suUser')?.value.trim() || '';
  const email = byId('suEmail')?.value.trim() || '';
  const password = byId('suPass')?.value || '';
  const confirmation = byId('suPass2')?.value || '';

  if (!username || !email) {
    showAuthError('Preencha o nome de usuário e o e-mail.');
    return;
  }
  if (!paintPasswordRules(password, confirmation)) {
    showAuthError('A senha ainda não atende a todas as regras listadas acima.');
    return;
  }

  const restoreButton = useSubmitButton('btnSignup', 'Cadastrando...');
  const releaseLock = acquireNetworkLock('cadastrar');
  try {
    await withDeadline(
      createAccount(username, email, password),
      SIGN_UP_TIMEOUT_MS,
      'O servidor de cadastro',
    );
  } catch (err) {
    console.warn(`${LOG_PREFIX} Falha ao cadastrar:`, err);
    showAuthError(describeAuthError(err));
  } finally {
    releaseLock();
    restoreButton();
  }
}

/* --------------------------------------------------------------------------
   ESQUECI MINHA SENHA
   -------------------------------------------------------------------------- */
byId('formForgot')?.addEventListener('submit', (event) => {
  event.preventDefault();
  void handlePasswordReset();
});

async function handlePasswordReset() {
  hideAuthMessages();
  const email = byId('fgEmail')?.value.trim() || '';
  if (!email) {
    showAuthError('Informe o e-mail da sua conta.');
    return;
  }

  const restoreButton = useSubmitButton('btnForgot', 'Enviando...');
  const releaseLock = acquireNetworkLock('redefinir senha');
  try {
    await withDeadline(
      sendPasswordResetEmail(auth, email),
      PASSWORD_RESET_TIMEOUT_MS,
      'O servidor de e-mail',
    );
    showAuthOk('Link de redefinição enviado. Confira sua caixa de entrada e a pasta de spam.');
  } catch (err) {
    console.warn(`${LOG_PREFIX} Falha ao enviar redefinição de senha:`, err);
    showAuthError(describeAuthError(err));
  } finally {
    releaseLock();
    restoreButton();
  }
}

/* --------------------------------------------------------------------------
   SAIR
   -------------------------------------------------------------------------- */
export function requestSignOut() {
  signOut(auth).catch((err) => {
    console.error(`${LOG_PREFIX} Falha ao sair:`, err);
    showAuthError('Não foi possível sair agora. Tente de novo.');
  });
}

byId('btnPendingLogout')?.addEventListener('click', requestSignOut);

/* --------------------------------------------------------------------------
   ROTA DA URL
   --------------------------------------------------------------------------
   window.applyPendingRouteFromHash é definida em cotacoes.js, que carrega
   depois deste arquivo. Se o Firebase responder muito rápido, ela ainda não
   existe — por isso a espera curta, que desiste em silêncio para nunca
   segurar o boot por causa de uma rota. */
const ROUTE_RETRY_INTERVAL_MS = 50;
const ROUTE_MAX_ATTEMPTS = 40;

function applyRouteFromUrlWhenReady(attempt = 0) {
  if (typeof window.applyPendingRouteFromHash === 'function') {
    window.applyPendingRouteFromHash();
    return;
  }
  if (attempt >= ROUTE_MAX_ATTEMPTS) return;
  setTimeout(() => applyRouteFromUrlWhenReady(attempt + 1), ROUTE_RETRY_INTERVAL_MS);
}

/* --------------------------------------------------------------------------
   ENTRADA LIBERADA
   -------------------------------------------------------------------------- */
function enterApplication(user, data, mode) {
  transitionTo(SessionState.READY);

  if (typeof window.showAppLoading === 'function') {
    window.showAppLoading(data.username || data.email || '');
  }
  window.dashboardMapReveal?.();
  applyRouteFromUrlWhenReady();

  mountAccountViews(user, session, { onSignOut: requestSignOut });
  startPresence(user.uid);

  void logActivity(
    user.uid,
    data.username,
    mode === 'bypass' ? 'Login realizado (bypass de manutenção)' : 'Login realizado',
  );
}

/**
 * Consulta o portão de manutenção e decide entre liberar o aplicativo ou
 * exibir a tela de bloqueio.
 *
 * Política de falha: liberar. Uma falha de leitura ou de rede nunca deve
 * impedir o acesso de quem tem credencial válida — é a mesma decisão já
 * adotada dentro de maintenance.js, agora aplicada também quando a própria
 * chamada ao portão trava.
 */
async function resolveMaintenanceGate(user, data) {
  let accessGranted = false;
  const grantAccess = (mode) => {
    accessGranted = true;
    enterApplication(user, data, mode);
  };

  try {
    await withDeadline(
      checkMaintenanceGate(
        { uid: user.uid, ...data },
        () => grantAccess('normal'),
        () => grantAccess('bypass'),
      ),
      MAINTENANCE_GATE_TIMEOUT_MS,
      'A verificação de manutenção',
    );
  } catch (err) {
    console.warn(`${LOG_PREFIX} O portão de manutenção falhou; liberando o acesso por segurança:`, err);
    if (!accessGranted) grantAccess('normal');
    return;
  }

  /* Nenhum dos dois retornos disparou: manutenção ativa e este usuário não é
     administrador. maintenance.js já colocou a própria tela no ar; aqui só
     fechamos o boot corretamente. Se o administrador clicar em "Entrar como
     Administrador" depois, grantAccess('bypass') assume a partir dali. */
  if (!accessGranted) {
    transitionTo(SessionState.MAINTENANCE);
  }
}

/* --------------------------------------------------------------------------
   ESTADO DE AUTENTICAÇÃO — ponto de entrada de tudo
   -------------------------------------------------------------------------- */
onAuthStateChanged(auth, (user) => {
  void handleAuthStateChange(user);
});

async function handleAuthStateChange(user) {
  /* Libera o splash assim que a autenticação se pronuncia, independente do
     que venha a seguir. */
  window.__splashAuthReady = true;
  if (typeof window.hideSplashIfReady === 'function') window.hideSplashIfReady();

  stopPresence();

  if (!user) {
    session.uid = null;
    session.data = null;
    unmountAccountViews();
    transitionTo(SessionState.SIGNED_OUT);
    return;
  }

  session.uid = user.uid;
  transitionTo(SessionState.LOADING);

  const releaseLock = acquireNetworkLock('boot da sessão');
  try {
    const data = await readUserProfile(user.uid);
    session.data = data;

    if (!data || data.status === 'pending') {
      transitionTo(SessionState.BLOCKED, BlockedReason.PENDING);
      return;
    }
    if (data.status === 'rejected') {
      transitionTo(SessionState.BLOCKED, BlockedReason.REJECTED);
      return;
    }

    await resolveMaintenanceGate(user, data);
  } catch (err) {
    console.error(`${LOG_PREFIX} Não foi possível concluir o boot da sessão:`, err);
    const detail = err instanceof TimeoutError
      ? `${err.message}`
      : 'O banco de dados recusou a leitura do seu cadastro.';
    transitionTo(SessionState.BLOCKED, blockedByError(detail));
  } finally {
    releaseLock();
  }
}
