/* ==========================================================================
   firebase-core.js — BDG · Controle de Cotações (Samsung SDS Logistics)
   ==========================================================================
   CAMADA DE INFRAESTRUTURA. Não toca no DOM, não conhece telas e não decide
   nada de negócio. Entrega quatro garantias, e só isso:

     1. Uma única instância de Firebase (app, auth, database), inicializada
        de forma determinística antes de qualquer outro módulo rodar.
     2. App Check isolado: se ele falhar, o aplicativo continua funcionando.
     3. Um TRINCO DE REDE POR CONTAGEM, que substitui o antigo
        window.__authBusy (ver "POR QUE O TRINCO MUDOU", abaixo).
     4. withDeadline(): nenhuma promessa de rede pode esperar para sempre.

   ---------------------------------------------------------------------------
   POR QUE O TRINCO MUDOU — a causa estrutural do login travado
   ---------------------------------------------------------------------------
   O código anterior coordenava três sistemas assíncronos independentes
   (autenticação, o laço de reconexão do Realtime Database em cotacoes.js e
   o portão de manutenção) através de um único booleano global compartilhado:
   window.__authBusy. Ele era escrito em 13 pontos diferentes, por dois
   arquivos distintos, sem dono definido.

   Um booleano não consegue representar "quantas operações estão em curso".
   Quando duas se sobrepunham — e o login sempre sobrepõe, porque o envio do
   formulário e o onAuthStateChanged rodam ao mesmo tempo — a primeira a
   terminar zerava a trava enquanto a segunda ainda estava no ar. Isso
   reabria exatamente a corrida que a trava existia para impedir: a
   reconexão automática derrubava a conexão do banco no meio da leitura do
   perfil, e a leitura ficava pendurada para sempre (dbGet não rejeita
   quando a conexão cai — ele simplesmente nunca resolve).

   Aqui o trinco passa a ser um CONTADOR com liberação idempotente:
   cada operação pega o seu próprio trinco e devolve o seu próprio trinco;
   devolver duas vezes não faz nada; o caminho só reabre quando o contador
   chega a zero. Não existe mais "passar a trava adiante" entre funções,
   que era de onde a corrida nascia.

   window.__authBusy continua sendo publicado como PROJEÇÃO do contador,
   porque cotacoes.js lê essa variável diretamente em dois pontos. Assim
   nenhuma linha de cotacoes.js precisa ser alterada.
   ========================================================================== */

import { initializeApp, getApps, getApp } from 'https://www.gstatic.com/firebasejs/12.16.0/firebase-app.js';
import {
  getAuth, setPersistence, browserLocalPersistence,
} from 'https://www.gstatic.com/firebasejs/12.16.0/firebase-auth.js';
import {
  getDatabase, ref as dbRef, set as dbSet, get as dbGet,
} from 'https://www.gstatic.com/firebasejs/12.16.0/firebase-database.js';
import {
  initializeAppCheck, ReCaptchaV3Provider,
} from 'https://www.gstatic.com/firebasejs/12.16.0/firebase-app-check.js';

/* --------------------------------------------------------------------------
   CONFIGURAÇÃO PÚBLICA DO FIREBASE
   --------------------------------------------------------------------------
   Estes valores NÃO são segredos e são publicados de propósito: qualquer
   aplicativo web do Firebase os entrega ao navegador em texto puro, por
   projeto. Eles apenas identificam o projeto. Quem protege os dados são as
   regras do Realtime Database e a autenticação por e-mail/senha. Analisadores
   estáticos costumam sinalizar "apiKey" como credencial embutida; neste
   contexto é falso positivo conhecido, e fica registrado aqui em vez de ser
   silenciado com uma anotação solta no meio do código. */
const PUBLIC_FIREBASE_CONFIG = Object.freeze({
  apiKey: 'AIzaSyBTRzLDGEsZWXDGvEksycxemUmmlBh3cLw',
  authDomain: 'bdg-sds.firebaseapp.com',
  databaseURL: 'https://bdg-sds-default-rtdb.firebaseio.com',
  projectId: 'bdg-sds',
  storageBucket: 'bdg-sds.firebasestorage.app',
  messagingSenderId: '30015192050',
  appId: '1:30015192050:web:7a9787ae83ab987fac496b',
  measurementId: 'G-E2WRE8BRYQ',
});

/* Chave do reCAPTCHA usada em bdg-cotacoes-staging (ambiente de testes).
   O projeto de produção (bdg-sds) nunca teve App Check registrado no
   console do Firebase — a versão anterior deste app não usava App Check.
   Deixar vazio aqui é proposital: startAppCheck() abaixo pula a
   inicialização quando não há chave, mantendo o comportamento atual
   (sem App Check) até que alguém registre um site key de reCAPTCHA v3
   para bdg-sds no Firebase Console e preencha esta constante. */
const RECAPTCHA_SITE_KEY = '';
const APP_CHECK_OFF_KEY = 'bdg:disableAppCheck';

const LOG_PREFIX = '[bdg]';

/* Reaproveita a instância existente em vez de criar outra. Criar duas vezes
   lança "Firebase App named '[DEFAULT]' already exists", que abortava o
   módulo inteiro e deixava o login sem nenhum ouvinte registrado. */
export const fbApp = getApps().length ? getApp() : initializeApp(PUBLIC_FIREBASE_CONFIG);

export const auth = getAuth(fbApp);
export const rtdb = getDatabase(fbApp, PUBLIC_FIREBASE_CONFIG.databaseURL);

/* Mantém a sessão salva no aparelho: atualizar a página ou fechar e reabrir
   o aplicativo não desloga. É o padrão do Firebase, declarado explicitamente
   para não depender de comportamento implícito. */
setPersistence(auth, browserLocalPersistence).catch((err) => {
  console.warn(`${LOG_PREFIX} Persistência local indisponível; a sessão pode não sobreviver ao fechamento do aplicativo:`, err);
});

/* --------------------------------------------------------------------------
   APP CHECK — camada adicional, nunca fatal, e agora desligável sem deploy
   --------------------------------------------------------------------------
   Quando o provedor reCAPTCHA falha (chave não registrada para o domínio em
   uso, ou o bloqueio de até 24 h que o próprio SDK ativa após um erro 403), o
   Firebase não lança exceção: ele passa a esperar por um token que nunca vem.
   Como o SDK de autenticação anexa o token do App Check às suas requisições,
   isso trava signInWithEmailAndPassword — a chamada simplesmente não resolve.

   INTERRUPTOR PELA BARRA DE ENDEREÇO (não precisa de console nem de deploy):
     ...?appcheck=off   desliga e memoriza no aparelho
     ...?appcheck=on    religa
   A preferência fica salva, então basta abrir uma vez com o parâmetro; as
   próximas aberturas já respeitam a escolha.

   URLSearchParams não está na lista de globais autorizadas pelo ESLint do
   projeto, então a leitura é por expressão regular sobre location.search. */
function readAppCheckSwitchFromUrl() {
  const match = /[?&]appcheck=(on|off)(?=&|$)/.exec(location.search);
  return match ? match[1] : null;
}

function readStoredAppCheckPreference() {
  try {
    return localStorage.getItem(APP_CHECK_OFF_KEY) === '1';
  } catch (err) {
    console.warn(`${LOG_PREFIX} localStorage indisponível ao ler o interruptor do App Check; seguindo com App Check ativo:`, err);
    return false;
  }
}

function storeAppCheckPreference(disabled) {
  try {
    if (disabled) localStorage.setItem(APP_CHECK_OFF_KEY, '1');
    else localStorage.removeItem(APP_CHECK_OFF_KEY);
  } catch (err) {
    console.warn(`${LOG_PREFIX} Não foi possível memorizar a preferência do App Check neste aparelho:`, err);
  }
}

function resolveAppCheckPreference() {
  const fromUrl = readAppCheckSwitchFromUrl();
  if (fromUrl === 'off') {
    storeAppCheckPreference(true);
    return true;
  }
  if (fromUrl === 'on') {
    storeAppCheckPreference(false);
    return false;
  }
  return readStoredAppCheckPreference();
}

/* Exportado para a tela de acesso avisar, de forma visível, que o modo de
   diagnóstico está ligado — no celular não há console para conferir. */
export const appCheckDisabled = resolveAppCheckPreference();

function startAppCheck() {
  if (!RECAPTCHA_SITE_KEY) {
    console.warn(`${LOG_PREFIX} App Check sem site key configurada para bdg-sds; seguindo sem ele (login e dados continuam protegidos pela autenticação e pelas regras do banco).`);
    return;
  }
  if (appCheckDisabled) {
    console.warn(`${LOG_PREFIX} App Check DESLIGADO neste aparelho. Login e dados seguem protegidos pela autenticação e pelas regras do banco. Para religar, abra a página com ?appcheck=on`);
    return;
  }
  try {
    initializeAppCheck(fbApp, {
      provider: new ReCaptchaV3Provider(RECAPTCHA_SITE_KEY),
      isTokenAutoRefreshEnabled: true,
    });
  } catch (err) {
    console.error(`${LOG_PREFIX} App Check não inicializou; o aplicativo segue sem ele:`, err);
  }
}
startAppCheck();

/* --------------------------------------------------------------------------
   TRINCO DE REDE POR CONTAGEM
   -------------------------------------------------------------------------- */

/* Teto de segurança por trinco. Nenhuma operação legítima do aplicativo passa
   disso; se passar, houve travamento e o trinco se devolve sozinho para não
   bloquear a reconexão automática permanentemente. */
const LOCK_MAX_HOLD_MS = 25000;

let operationsInFlight = 0;

function publishLegacyBusyFlag() {
  /* Compatibilidade: cotacoes.js lê window.__authBusy diretamente para adiar
     goOffline/goOnline durante operações de autenticação. Publicar o contador
     como booleano mantém aquele arquivo funcionando sem uma única alteração. */
  window.__authBusy = operationsInFlight > 0;
}
publishLegacyBusyFlag();

/**
 * Registra uma operação de rede em andamento e devolve a função que a encerra.
 * A função devolvida é idempotente: chamá-la mais de uma vez não tem efeito,
 * então nenhuma operação consegue liberar o trinco de outra.
 *
 * @param {string} label Nome da operação, usado só em log de diagnóstico.
 * @returns {() => void} Encerrador idempotente.
 */
export function acquireNetworkLock(label) {
  operationsInFlight += 1;
  publishLegacyBusyFlag();

  let alreadyReleased = false;

  const release = () => {
    if (alreadyReleased) return;
    alreadyReleased = true;
    clearTimeout(safetyTimer);
    operationsInFlight = Math.max(0, operationsInFlight - 1);
    publishLegacyBusyFlag();
  };

  const safetyTimer = setTimeout(() => {
    console.warn(`${LOG_PREFIX} A operação "${label}" passou de ${LOCK_MAX_HOLD_MS} ms sem encerrar. Liberando o trinco para não bloquear a reconexão.`);
    release();
  }, LOCK_MAX_HOLD_MS);

  return release;
}

/* --------------------------------------------------------------------------
   PRAZO EM TODA PROMESSA DE REDE
   -------------------------------------------------------------------------- */

/**
 * Erro lançado quando uma operação estoura o prazo. Ter um tipo próprio
 * permite distinguir "a rede demorou" de "o Firebase recusou", que exigem
 * mensagens diferentes na tela. O código anterior usava a string sentinela
 * '__TIMEOUT__' como valor de retorno, que colidiria com um dado real e
 * obrigava cada chamador a lembrar de comparar a string.
 */
export class TimeoutError extends Error {
  constructor(label, ms) {
    super(`${label} não respondeu em ${Math.round(ms / 1000)}s.`);
    this.name = 'TimeoutError';
    this.label = label;
    this.timeoutMs = ms;
  }
}

/**
 * Impõe um prazo a qualquer promessa.
 *
 * Necessário porque dbGet() do Realtime Database NÃO rejeita quando a conexão
 * cai — a promessa apenas nunca resolve. Sem prazo, um soluço de rede prende
 * a tela de carregamento indefinidamente, sem erro nenhum no console. É
 * exatamente o sintoma relatado ("demora e não entra").
 *
 * @template T
 * @param {Promise<T>} promise
 * @param {number} ms
 * @param {string} label
 * @returns {Promise<T>}
 */
export function withDeadline(promise, ms, label) {
  let timer = null;
  const deadline = new Promise((_resolve, reject) => {
    timer = setTimeout(() => reject(new TimeoutError(label, ms)), ms);
  });
  return Promise.race([promise, deadline]).finally(() => clearTimeout(timer));
}

/* --------------------------------------------------------------------------
   ACESSO A DADOS
   -------------------------------------------------------------------------- */

/* Prazos POR TENTATIVA, não um valor único.

   A primeira leitura carrega junto o handshake da conexão com o Realtime
   Database, que em rede móvel fraca sozinho já consome vários segundos — um
   prazo de 8s uniforme derrubava logins legítimos nessa etapa. A repetição
   reaproveita a conexão já aberta, então pode ser bem mais curta: se ela
   também estourar, o problema não é lentidão, é a conexão não ter subido. */
const PROFILE_READ_TIMEOUTS_MS = [10000, 5000];

/**
 * Lê users/{uid} com prazo e uma nova tentativa. Em rede móvel a primeira
 * tentativa falhar é comum e a segunda costuma passar.
 *
 * @param {string} uid
 * @returns {Promise<object|null>} Dados do usuário, ou null se não existir.
 * @throws {TimeoutError} Se todas as tentativas estourarem o prazo.
 */
export async function readUserProfile(uid) {
  let lastError = null;
  let attempt = 0;
  for (const timeoutMs of PROFILE_READ_TIMEOUTS_MS) {
    attempt += 1;
    try {
      const snapshot = await withDeadline(
        dbGet(dbRef(rtdb, `users/${uid}`)),
        timeoutMs,
        'A leitura do seu cadastro',
      );
      return snapshot.exists() ? snapshot.val() : null;
    } catch (err) {
      lastError = err;
      console.warn(`${LOG_PREFIX} Leitura do perfil falhou (tentativa ${attempt} de ${PROFILE_READ_TIMEOUTS_MS.length}):`, err);
    }
  }
  throw lastError;
}

/* Sufixo aleatório para IDs de log. Usa a Web Crypto API em vez de
   Math.random(): mesmo sem valor de segurança direto aqui, IDs previsíveis
   podem ser enumerados. */
function randomIdSuffix(length = 5) {
  const bytes = new Uint8Array(length);
  crypto.getRandomValues(bytes);
  return Array.from(bytes, (b) => (b % 36).toString(36)).join('');
}

/**
 * Registra uma ação no histórico de atividades.
 *
 * Falha aqui nunca interrompe o fluxo de quem chamou: o log é auditoria, não
 * requisito de acesso. Por isso a promessa é resolvida mesmo em erro, com o
 * problema registrado no console.
 *
 * @param {string} uid
 * @param {string} username
 * @param {string} action
 * @returns {Promise<void>}
 */
export async function logActivity(uid, username, action) {
  const id = `a_${Date.now()}_${randomIdSuffix(5)}`;
  try {
    await dbSet(dbRef(rtdb, `activityLog/${id}`), {
      uid, username, action, ts: Date.now(),
    });
  } catch (err) {
    console.warn(`${LOG_PREFIX} Não foi possível registrar a atividade "${action}":`, err);
  }
}
