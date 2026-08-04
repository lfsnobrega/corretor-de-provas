// Service worker mínimo — existe só pra habilitar "Adicionar à tela inicial" no
// Android/Chrome (é um requisito técnico do navegador pra considerar o site
// instalável). NÃO faz cache nem funciona offline de propósito — o sistema
// depende do servidor pra tudo (banco de dados, processamento de OMR), então
// fingir que funciona offline ia confundir mais do que ajudar. Se um dia quiser
// suporte offline de verdade, esse arquivo é o lugar pra isso — mas é uma
// mudança de risco maior, então por enquanto ele só repassa as requisições.

self.addEventListener("install", (event) => {
  self.skipWaiting();
});

self.addEventListener("activate", (event) => {
  self.clients.claim();
});

self.addEventListener("fetch", (event) => {
  // Repassa a requisição direto pra rede, sem cache — mantém o comportamento
  // exatamente igual a não ter service worker nenhum, só que agora o navegador
  // reconhece o site como instalável.
  event.respondWith(fetch(event.request));
});