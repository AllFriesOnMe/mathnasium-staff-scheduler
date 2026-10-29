'use strict';
const AiFallback = (() => {
  async function resolve(cases, accessCode, signal, send = fetch) {
    if (!cases?.length) return [];
    if (!accessCode.trim())
      throw new Error(
        'Enter the center’s AI access code to check unresolved schedule assignments.',
      );
    const body = JSON.stringify({ cases });
    if (new TextEncoder().encode(body).length > 128 * 1024)
      throw new Error('Too much unresolved context for one AI check. Local results are unchanged.');
    const response = await send('/api/handoffs', {
      method: 'POST',
      credentials: 'same-origin',
      redirect: 'error',
      signal: AbortSignal.any([signal, AbortSignal.timeout(30000)]),
      headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${accessCode.trim()}` },
      body,
    });
    if (!response.headers.get('Content-Type')?.includes('application/json'))
      throw new Error('AI fallback is unavailable. Deploy the Cloudflare worker first.');
    const result = await response.json();
    if (!response.ok)
      throw new Error(result.error || 'AI check failed. Local results are unchanged.');
    if (!Array.isArray(result.resolutions) || result.resolutions.length > cases.length)
      throw new Error('Invalid AI response. Local results are unchanged.');
    return result.resolutions;
  }
  return { resolve };
})();
