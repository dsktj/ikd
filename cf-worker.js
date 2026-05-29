// ═══════════════════════════════════════════
// Cloudflare Worker — Proxy GAS Web App
// ═══════════════════════════════════════════
// Hapus bar "This application was created by a Google Apps Script user"
// ═══════════════════════════════════════════

const GAS_URL = 'https://script.google.com/macros/s/AKfycbxonfj26eCaYt17zX0JqW-JC2o7bH8TxE3gTZKcdfTskstFY9sxRj0QJa03WayyOmiP/exec';

async function handleRequest(request) {
  const url = new URL(request.url);
  const target = GAS_URL + url.search;

  const response = await fetch(target, {
    method: request.method,
    headers: request.headers,
    body: ['GET','HEAD'].includes(request.method) ? null : request.body,
    redirect: 'follow',
  });

  const ct = response.headers.get('content-type') || '';
  if (!ct.includes('text/html')) {
    return new Response(response.body, { status: response.status, headers: response.headers });
  }

  // Stream + hapus elemen Google yang tidak diinginkan
  return new HTMLRewriter()
    .on('div[class*="google"]', { element: e => e.remove() })
    .on('div[id*="google"]', { element: e => e.remove() })
    .on('div[class*="drive"]', { element: e => e.remove() })
    .on('div[id*="drive"]', { element: e => e.remove() })
    .on('div[style*="position"][style*="fixed"]', { element: e => { if (e.getAttribute('style')?.includes('top')) e.remove(); }})
    .on('a[href*="support.google.com"]', { element: e => { const p = e.parentElement; if (p) p.remove(); }})
    .transform(response);
}

addEventListener('fetch', event => {
  event.respondWith(handleRequest(event.request));
});
