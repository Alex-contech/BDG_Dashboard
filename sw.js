/* Auditoria Geral 01/08 (Bloco 4) — v4 → v5. A troca é obrigatória sempre que
   os sufixos ?v= dos módulos mudam: o activate abaixo apaga todo cache cuja
   chave não seja CACHE_NAME, garantindo que ninguém fique com a mistura de
   um index.html novo apontando para módulos antigos já em cache. */
const CACHE_NAME = 'bdg-sds-v5';
const APP_SHELL = ['./', './index.html', './manifest.json', './icon-192.png', './icon-512.png'];

/* CORREÇÃO CRÍTICA (Auditoria Geral 01/08, Bloco 4) — o Service Worker nunca
   chegava a instalar.

   cache.addAll() é tudo-ou-nada: basta UMA das URLs responder 404 para a
   Promise inteira rejeitar. Como a rejeição ia direto para event.waitUntil(),
   o evento install falhava, o worker jamais ativava, e o app não tinha
   funcionamento offline nenhum — silenciosamente, porque um install que
   falha não gera erro visível em tela.

   O repositório não contém nenhum arquivo .png, mas APP_SHELL lista
   icon-192.png e icon-512.png. Ou seja: a condição de falha estava presente
   desde sempre.

   Agora cada item é buscado individualmente e uma falha isolada é registrada
   sem derrubar os demais. O app passa a funcionar offline com o que existe,
   em vez de não funcionar por causa do que falta. */
self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => Promise.all(
      APP_SHELL.map((url) => cache.add(url).catch((err) => {
        console.warn('[sw] não foi possível pré-cachear', url, '— seguindo sem ele:', err);
      }))
    ))
  );
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k)))
    )
  );
  self.clients.claim();
});

/* CORREÇÃO CRÍTICA — login/carregamento travando em rede móvel instável
   ==========================================================================
   DIAGNÓSTICO: o fetch() de dentro do respondWith() abaixo não tinha
   NENHUM timeout, e interceptava TODA requisição GET — inclusive de
   terceiros (SDK do Firebase em www.gstatic.com, Chart.js/xlsx/jsPDF/
   Tesseract em cdnjs.cloudflare.com, Bootstrap em jsdelivr, fontes do
   Google). Em 5G com sinal fraco, se qualquer uma dessas respostas
   demorasse muito, o respondWith() ficava pendurado esperando a rede
   indefinidamente ANTES de cair para o cache — e como um import ES
   trava a execução do módulo inteiro até resolver, isso podia travar o
   próprio carregamento do SDK do Firebase dentro do auth.js, mesmo com
   `defer` e cache-buster corretos. É o mesmo sintoma de "demora e não
   entra" de antes, só que numa camada que nenhuma correção anterior
   tocou (este arquivo não tinha sido revisado ainda).

   CORREÇÃO (duas partes):
   1) Requisições de terceiros (outra origem) não são mais interceptadas
      — o navegador cuida delas do jeito normal dele, com o próprio cache
      HTTP (essas URLs já têm versão fixa no caminho, então isso é seguro
      e não perde nada).
   2) Para os arquivos do próprio app (mesma origem: index.html, auth.js,
      cotacoes.js, etc.), a rede continua vindo primeiro — sempre busca a
      versão mais nova — mas agora com um limite de NETWORK_TIMEOUT_MS:
      se a rede não responder a tempo, cai para o cache na hora em vez de
      ficar pendurado. Mesmo padrão de timeout já usado em auth.js e
      maintenance.js (Promise.race). */
/* Dois prazos distintos, e a diferença importa muito.

   O documento HTML (navegação) é o que decide QUAL versão de todo o resto do
   aplicativo vai carregar. Se ele cair para o cache cedo demais, o aparelho
   volta a rodar uma versão antiga inteira — foi exatamente o sintoma de
   "corrijo, faço deploy e continua igual". Por isso a navegação tem um prazo
   bem mais folgado: só recorre ao cache quando a rede realmente não vem.

   Os demais arquivos podem cair para o cache rápido sem esse risco, porque
   suas URLs carregam a versão (?v=...) — uma cópia em cache ali é, por
   definição, a cópia certa daquela versão. */
const NAVIGATION_TIMEOUT_MS = 12000;
const ASSET_TIMEOUT_MS = 4000;

function isDocumentRequest(request) {
  return request.mode === 'navigate'
    || (request.headers.get('accept') || '').includes('text/html');
}

function fetchWithTimeout(request, ms) {
  return new Promise((resolve, reject) => {
    const timer = setTimeout(() => reject(new Error('SW_NETWORK_TIMEOUT')), ms);
    fetch(request).then(
      (res) => { clearTimeout(timer); resolve(res); },
      (err) => { clearTimeout(timer); reject(err); }
    );
  });
}

self.addEventListener('fetch', (event) => {
  if (event.request.method !== 'GET') return;

  // Terceiros: não intercepta. O navegador resolve normalmente.
  if (new URL(event.request.url).origin !== self.location.origin) return;

  event.respondWith(
    fetchWithTimeout(event.request, isDocumentRequest(event.request) ? NAVIGATION_TIMEOUT_MS : ASSET_TIMEOUT_MS)
      .then((response) => {
        const copy = response.clone();
        caches.open(CACHE_NAME).then((cache) => cache.put(event.request, copy)).catch(()=>{});
        return response;
      })
      .catch(() => caches.match(event.request).then((cached) => cached || caches.match('./index.html')))
  );
});
