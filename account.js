/* ==========================================================================
   account.js — BDG · Controle de Cotações (Samsung SDS Logistics)
   ==========================================================================
   Telas de PERFIL e USUÁRIOS, e o batimento de presença (online/offline).

   Separado do auth.js porque não é sessão: nada aqui decide se alguém entra
   no aplicativo. Eram ~330 das 737 linhas do arquivo antigo, e a mistura era
   o que tornava o módulo de login difícil de revisar com segurança.

   DUAS FALHAS REAIS CORRIGIDAS NESTA EXTRAÇÃO
   ---------------------------------------------------------------------------
   1. XSS armazenado via foto de perfil.
      O código anterior injetava data.photoURL e u.photoURL direto no
      atributo src de uma <img>, sem escapar e sem validar. Um valor gravado
      com aspas no campo escapa do atributo e injeta HTML — e a lista de
      usuários renderiza a foto de TODOS os cadastros, então um usuário
      comum conseguiria atingir o administrador que abrisse aquela tela.
      Agora toda origem de imagem passa por safeImageSource(), que só aceita
      data:image/... e https://, e todo texto passa por escapeHtml().

   2. Vazamento de ouvinte na lista de usuários.
      onValue em 'users' era registrado de novo a cada mudança de estado de
      autenticação, sem nunca ser cancelado. Entrar e sair algumas vezes
      acumulava ouvintes sobre o mesmo nó, multiplicando tráfego e
      redesenhos. Agora existe um único registro, cancelado ao sair.
   ========================================================================== */

import { updateProfile } from 'https://www.gstatic.com/firebasejs/12.16.0/firebase-auth.js';
import {
  ref as dbRef, update as dbUpdate, remove as dbRemove, onValue as dbOnValue,
  query as dbQuery, orderByChild as dbOrderByChild, equalTo as dbEqualTo,
  limitToLast as dbLimitToLast,
} from 'https://www.gstatic.com/firebasejs/12.16.0/firebase-database.js';

import { rtdb, withDeadline, logActivity } from './firebase-core.js?v=20260802a';

const LOG_PREFIX = '[account]';

const PHOTO_SAVE_TIMEOUT_MS = 15000;
const PRESENCE_INTERVAL_MS = 45000;
const ONLINE_WINDOW_MS = 120000;
const ACTIVITY_FEED_LIMIT = 50;
const TOAST_MS = 2500;

const DEFAULT_AVATAR = `data:image/svg+xml;utf8,${encodeURIComponent(
  '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="#7C87C2">'
  + '<circle cx="12" cy="8" r="4"/><path d="M4 20c0-4.4 3.6-8 8-8s8 3.6 8 8"/></svg>',
)}`;

const byId = (id) => document.getElementById(id);

/* --------------------------------------------------------------------------
   SANEAMENTO DE SAÍDA
   -------------------------------------------------------------------------- */
const HTML_ESCAPES = new Map([
  ['&', '&amp;'], ['<', '&lt;'], ['>', '&gt;'], ['"', '&quot;'], ["'", '&#39;'],
]);

function escapeHtml(value) {
  return String(value ?? '').replace(/[&<>"']/g, (char) => HTML_ESCAPES.get(char));
}

/**
 * Só devolve a origem da imagem se ela for uma foto embutida (data:image/...)
 * ou um endereço https. Qualquer outra coisa vira o avatar padrão.
 * Fecha o vetor de XSS descrito no cabeçalho deste arquivo.
 */
function safeImageSource(url) {
  const value = String(url ?? '').trim();
  if (!value) return DEFAULT_AVATAR;
  const isEmbeddedImage = value.startsWith('data:image/');
  const isHttps = value.startsWith('https://');
  if (!isEmbeddedImage && !isHttps) return DEFAULT_AVATAR;
  return escapeHtml(value);
}

/* --------------------------------------------------------------------------
   PRESENÇA
   -------------------------------------------------------------------------- */
let presenceTimer = null;

export function startPresence(uid) {
  stopPresence();
  const beat = () => {
    dbUpdate(dbRef(rtdb, `users/${uid}`), { lastActive: Date.now() })
      .catch((err) => console.warn(`${LOG_PREFIX} Não foi possível atualizar a presença:`, err));
  };
  beat();
  presenceTimer = setInterval(beat, PRESENCE_INTERVAL_MS);
}

export function stopPresence() {
  if (presenceTimer) {
    clearInterval(presenceTimer);
    presenceTimer = null;
  }
}

/* --------------------------------------------------------------------------
   MONTAGEM / DESMONTAGEM
   -------------------------------------------------------------------------- */
let unsubscribeUsers = null;
let unsubscribeActivity = null;

/**
 * Monta as telas de Perfil e Usuários para a sessão recém-aberta.
 *
 * @param {object} user Usuário do Firebase Auth (necessário para updateProfile).
 * @param {{uid: string, data: object}} session Sessão viva, compartilhada por
 *   referência com auth.js — atualizar session.data aqui mantém
 *   window.getCurrentUser() coerente sem dependência circular entre módulos.
 * @param {{onSignOut: () => void}} actions
 */
export function mountAccountViews(user, session, actions) {
  renderProfileView(user, session, actions);
  renderUsersView(session);
}

export function unmountAccountViews() {
  if (unsubscribeUsers) {
    unsubscribeUsers();
    unsubscribeUsers = null;
  }
  closeActivityFeed();
  const badge = byId('usersPendingBadge');
  if (badge) badge.classList.remove('show');
}

/* --------------------------------------------------------------------------
   PERFIL
   -------------------------------------------------------------------------- */
function renderProfileView(user, session, actions) {
  const root = byId('view-profile');
  if (!root) return;

  const data = session.data || {};
  const roleLabel = data.role === 'admin' ? 'Administrador' : 'Usuário';
  const removeButton = data.photoURL
    ? '<button type="button" class="btn btn-ghost" id="btnRemoveProfilePhoto" style="padding:5px 12px;font-size:11px;">Remover foto</button>'
    : '';

  root.innerHTML = `
    <div class="section-title"><h2>Meu perfil</h2><span>${roleLabel}</span></div>
    <div class="profile-card">
      <div class="profile-head">
        <div class="profile-upload">
          <img class="avatar lg" id="profileAvatar" src="${safeImageSource(data.photoURL)}" alt="Foto de perfil" style="transition:opacity .2s ease;">
          <div class="edit-badge">✎</div>
          <input type="file" id="profilePhotoInput" accept="image/*">
        </div>
        <div>
          <div class="field"><label for="profileName">Nome</label><input type="text" id="profileName" value="${escapeHtml(data.username)}"></div>
          <div class="hint" style="margin-top:6px;">${escapeHtml(data.email)}</div>
          <div style="margin-top:8px; display:flex; align-items:center; gap:10px;">
            <span id="profilePhotoStatus" style="font-size:11.5px; color:var(--ink-500);"></span>
            ${removeButton}
          </div>
        </div>
      </div>
      <button class="btn btn-primary" id="btnSaveProfile">Salvar alterações</button>
      <span class="toast" id="profileToast" style="margin-left:10px;">✓ Perfil atualizado</span>
      <div style="margin-top:22px;">
        <button class="btn btn-ghost" id="btnLogoutProfile">Sair da conta</button>
      </div>
    </div>
  `;

  bindProfilePhoto(user, session, actions);
  bindProfileName(user, session);
  byId('btnLogoutProfile')?.addEventListener('click', actions.onSignOut);
}

function setPhotoStatus(message) {
  const el = byId('profilePhotoStatus');
  if (el) el.textContent = message;
}

function bindProfilePhoto(user, session, actions) {
  byId('profilePhotoInput')?.addEventListener('change', (event) => {
    const [file] = event.target.files;
    if (!file) return;
    /* O recorte já devolve a imagem comprimida e no tamanho certo. */
    window.openPhotoCropper(file, {
      targetSize: 420,
      quality: 0.8,
      round: true,
      onConfirm: (dataUrl) => { void saveProfilePhoto(user, session, actions, dataUrl); },
    });
    event.target.value = '';
  });

  byId('btnRemoveProfilePhoto')?.addEventListener('click', () => {
    void removeProfilePhoto(user, session, actions);
  });
}

async function saveProfilePhoto(user, session, actions, dataUrl) {
  const avatar = byId('profileAvatar');
  if (avatar) avatar.style.opacity = '0.45';
  setPhotoStatus('Salvando foto...');
  try {
    await withDeadline(
      dbUpdate(dbRef(rtdb, `users/${user.uid}`), { photoURL: dataUrl }),
      PHOTO_SAVE_TIMEOUT_MS,
      'A gravação da foto',
    );
    session.data = { ...session.data, photoURL: dataUrl };
    setPhotoStatus('✓ Foto atualizada');
    await logActivity(user.uid, session.data.username, 'Foto de perfil atualizada');
    renderProfileView(user, session, actions);
  } catch (err) {
    console.error(`${LOG_PREFIX} Falha ao salvar a foto de perfil:`, err);
    setPhotoStatus('');
    if (avatar) avatar.style.opacity = '1';
    alert(`Não foi possível salvar essa foto: ${err.message}`);
  }
}

async function removeProfilePhoto(user, session, actions) {
  if (!confirm('Remover sua foto de perfil?')) return;
  try {
    await dbUpdate(dbRef(rtdb, `users/${user.uid}`), { photoURL: null });
    session.data = { ...session.data, photoURL: null };
    await logActivity(user.uid, session.data.username, 'Foto de perfil removida');
    renderProfileView(user, session, actions);
  } catch (err) {
    console.error(`${LOG_PREFIX} Falha ao remover a foto de perfil:`, err);
    alert(`Não foi possível remover a foto: ${err.message}`);
  }
}

function bindProfileName(user, session) {
  byId('btnSaveProfile')?.addEventListener('click', () => { void saveProfileName(user, session); });
}

async function saveProfileName(user, session) {
  const newName = byId('profileName')?.value.trim() || '';
  if (!newName) return;
  try {
    await updateProfile(user, { displayName: newName });
    await dbUpdate(dbRef(rtdb, `users/${user.uid}`), { username: newName });
    session.data = { ...session.data, username: newName };
    await logActivity(user.uid, newName, 'Nome de perfil alterado');
    const toast = byId('profileToast');
    if (toast) {
      toast.classList.add('show');
      setTimeout(() => toast.classList.remove('show'), TOAST_MS);
    }
  } catch (err) {
    console.error(`${LOG_PREFIX} Falha ao salvar o nome:`, err);
    alert(`Não foi possível salvar o nome: ${err.message}`);
  }
}

/* --------------------------------------------------------------------------
   USUÁRIOS
   -------------------------------------------------------------------------- */
function renderUsersView(session) {
  const root = byId('view-users');
  if (!root) return;

  root.innerHTML = `
    <div class="section-title"><h2>Usuários</h2><span id="usersCount">carregando...</span></div>
    <div id="pendingSection"></div>
    <div class="user-list" id="usersList"></div>
    <div class="section-title" id="activityTitle" style="display:none;"><h2>Histórico de atividades</h2><span>últimas ações registradas</span><button type="button" class="btn btn-ghost" id="btnCloseActivity" style="margin-left:auto;padding:6px 14px;font-size:11.5px;">Fechar</button></div>
    <div class="user-list" id="activityFeed" style="display:none;"></div>
  `;
  byId('btnCloseActivity')?.addEventListener('click', closeActivityFeed);

  subscribeToUsers(session);
}

function subscribeToUsers(session) {
  if (unsubscribeUsers) unsubscribeUsers();
  unsubscribeUsers = dbOnValue(
    dbRef(rtdb, 'users'),
    (snapshot) => paintUsers(snapshot.val() || {}, session),
    (err) => {
      console.error(`${LOG_PREFIX} Não foi possível carregar a lista de usuários:`, err);
      const list = byId('usersList');
      if (list) list.innerHTML = '<div class="user-row">Não foi possível carregar a lista agora.</div>';
    },
  );
}

function paintUsers(usersObject, session) {
  /* Map em vez do objeto cru: as chaves são uids vindos do banco, e daqui
     em diante todo acesso é por get(), sem tocar no protótipo. */
  const usersByUid = new Map(Object.entries(usersObject));
  const everyone = [...usersByUid].map(([uid, value]) => ({ uid, ...value }));
  const viewer = session.data || {};
  const isAdmin = viewer.role === 'admin';

  const count = byId('usersCount');
  if (count) count.textContent = `${everyone.length} usuário(s)`;

  const pending = everyone.filter((u) => u.status === 'pending');
  paintPendingBadge(pending.length, isAdmin);
  paintPendingSection(pending, usersByUid, session, isAdmin);
  paintApprovedList(everyone, usersByUid, session, isAdmin);
}

function paintPendingBadge(pendingCount, isAdmin) {
  const badge = byId('usersPendingBadge');
  if (!badge) return;
  if (isAdmin && pendingCount > 0) {
    badge.textContent = pendingCount > 99 ? '99+' : String(pendingCount);
    badge.classList.add('show');
  } else {
    badge.classList.remove('show');
  }
}

function pendingRowHtml(user) {
  return `
    <div class="user-row">
      <div class="user-pill">
        <img class="avatar" src="${safeImageSource(user.photoURL)}" alt="">
        <div class="who"><b>${escapeHtml(user.username)}</b><span>${escapeHtml(user.email)}</span></div>
      </div>
      <div class="approve-actions">
        <button class="btn-mini ok" data-approve="${escapeHtml(user.uid)}">Aprovar</button>
        <button class="btn-mini no" data-reject="${escapeHtml(user.uid)}">Rejeitar</button>
        <button class="btn-mini no" data-delete="${escapeHtml(user.uid)}" title="Remove o cadastro completamente">Excluir</button>
      </div>
    </div>`;
}

function paintPendingSection(pending, usersByUid, session, isAdmin) {
  const container = byId('pendingSection');
  if (!container) return;

  if (!isAdmin || pending.length === 0) {
    container.innerHTML = '';
    return;
  }

  container.innerHTML = `<div class="section-title"><h2>Aprovações pendentes</h2><span>${pending.length} aguardando</span></div>
    <div class="user-list">${pending.map(pendingRowHtml).join('')}</div>`;

  bindAction(container, 'approve', (uid) => changeUserStatus(uid, 'approved', usersByUid, session));
  bindAction(container, 'reject', (uid) => changeUserStatus(uid, 'rejected', usersByUid, session));
  bindAction(container, 'delete', (uid) => deleteUser(uid, usersByUid, session));
}

function bindAction(container, attribute, handler) {
  container.querySelectorAll(`[data-${attribute}]`).forEach((button) => {
    button.addEventListener('click', (event) => {
      event.stopPropagation();
      void handler(button.getAttribute(`data-${attribute}`));
    });
  });
}

async function changeUserStatus(uid, status, usersByUid, session) {
  const label = status === 'approved' ? 'Aprovou' : 'Rejeitou';
  try {
    await dbUpdate(dbRef(rtdb, `users/${uid}`), { status });
    await logActivity(
      session.uid,
      session.data?.username,
      `${label} o cadastro de ${usersByUid.get(uid)?.username || uid}`,
    );
  } catch (err) {
    console.error(`${LOG_PREFIX} Falha ao alterar o status do usuário:`, err);
    alert(`Não foi possível atualizar esse cadastro: ${err.message}`);
  }
}

async function deleteUser(uid, usersByUid, session) {
  const name = usersByUid.get(uid)?.username || uid;
  const confirmed = confirm(
    `Excluir o cadastro de "${name}"? Essa pessoa some da lista e perde o acesso. `
    + 'A conta de login no Firebase Authentication continua existindo e precisa ser apagada lá separadamente.',
  );
  if (!confirmed) return;
  try {
    await logActivity(session.uid, session.data?.username, `Excluiu o cadastro de ${name}`);
    await dbRemove(dbRef(rtdb, `users/${uid}`));
  } catch (err) {
    console.error(`${LOG_PREFIX} Falha ao excluir o cadastro:`, err);
    alert(`Não foi possível excluir esse cadastro: ${err.message}`);
  }
}

function approvedRowHtml(user, viewerUid, isAdmin, now) {
  const online = (now - (user.lastActive || 0)) < ONLINE_WINDOW_MS;
  const lastSeen = user.lastActive ? new Date(user.lastActive).toLocaleString('pt-BR') : '—';
  const adminSuffix = user.role === 'admin' ? ' · admin' : '';
  const deleteButton = isAdmin && user.uid !== viewerUid
    ? `<button class="btn-mini no" data-delete-approved="${escapeHtml(user.uid)}" title="Remove o cadastro completamente">Excluir</button>`
    : '';

  return `<div class="user-row" data-open="${escapeHtml(user.uid)}">
    <div class="user-pill">
      <img class="avatar" src="${safeImageSource(user.photoURL)}" alt="">
      <div class="who"><b>${escapeHtml(user.username)}${adminSuffix}</b><span>${escapeHtml(user.email)}</span></div>
    </div>
    <div class="meta">
      <div style="display:flex;align-items:center;gap:6px;justify-content:flex-end;">
        <span class="status-dot ${online ? 'on' : 'off'}"></span>${online ? 'Online' : 'Offline'}
      </div>
      <div class="last">último acesso: ${escapeHtml(lastSeen)}</div>
    </div>
    ${deleteButton}
  </div>`;
}

function paintApprovedList(everyone, usersByUid, session, isAdmin) {
  const list = byId('usersList');
  if (!list) return;

  const now = Date.now();
  const approved = everyone
    .filter((u) => u.status === 'approved')
    .sort((a, b) => (b.lastActive || 0) - (a.lastActive || 0));

  list.innerHTML = approved
    .map((u) => approvedRowHtml(u, session.uid, isAdmin, now))
    .join('') || '<div class="user-row">Nenhum usuário aprovado ainda.</div>';

  list.querySelectorAll('[data-open]').forEach((row) => {
    row.addEventListener('click', (event) => {
      if (event.target.closest('[data-delete-approved]')) return;
      const uid = row.getAttribute('data-open');
      openActivityFeed(uid, usersByUid.get(uid));
    });
  });

  bindAction(list, 'delete-approved', (uid) => deleteUser(uid, usersByUid, session));
}

/* --------------------------------------------------------------------------
   HISTÓRICO DE ATIVIDADES
   --------------------------------------------------------------------------
   Consulta indexada no servidor, limitada às últimas entradas do usuário.
   Requer a regra "activityLog": { ".indexOn": "uid" } no Realtime Database —
   sem o índice o Firebase ainda responde, mas baixa o nó inteiro e filtra no
   navegador. */
function openActivityFeed(uid, userInfo) {
  const title = byId('activityTitle');
  const feed = byId('activityFeed');
  if (!title || !feed) return;

  title.style.display = '';
  feed.style.display = '';
  const heading = title.querySelector('h2');
  if (heading) heading.textContent = `Histórico — ${userInfo?.username || uid}`;

  if (unsubscribeActivity) unsubscribeActivity();
  const activityQuery = dbQuery(
    dbRef(rtdb, 'activityLog'),
    dbOrderByChild('uid'),
    dbEqualTo(uid),
    dbLimitToLast(ACTIVITY_FEED_LIMIT),
  );

  unsubscribeActivity = dbOnValue(
    activityQuery,
    (snapshot) => {
      const entries = Object.values(snapshot.val() || {}).sort((a, b) => b.ts - a.ts);
      const rows = entries.map((entry) => `<div class="activity-item"><span>${escapeHtml(entry.action)}</span><span class="t">${escapeHtml(new Date(entry.ts).toLocaleString('pt-BR'))}</span></div>`).join('');
      feed.innerHTML = `<div class="activity-list" style="padding:14px;">${rows || '<div class="activity-item">Nenhuma atividade registrada ainda.</div>'}</div>`;
    },
    (err) => {
      console.error(`${LOG_PREFIX} Falha ao consultar o histórico (confira o índice .indexOn "uid"):`, err);
      feed.innerHTML = '<div class="activity-item">Não foi possível carregar o histórico agora.</div>';
    },
    { onlyOnce: true },
  );
}

function closeActivityFeed() {
  if (unsubscribeActivity) {
    unsubscribeActivity();
    unsubscribeActivity = null;
  }
  const title = byId('activityTitle');
  const feed = byId('activityFeed');
  if (title) title.style.display = 'none';
  if (feed) feed.style.display = 'none';
}
