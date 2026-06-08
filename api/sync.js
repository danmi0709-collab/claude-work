// Vercel 서버리스 함수 — Apps Script CORS 우회 프록시
module.exports = async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST') return res.status(405).json({ error: 'POST only' });

  const { appsScriptUrl, action, data } = req.body || {};
  if (!appsScriptUrl) return res.status(400).json({ error: 'appsScriptUrl 없음' });

  try {
    if (action === 'push') {
      const r = await fetch(appsScriptUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ action: 'set', data }),
        redirect: 'follow'
      });
      const text = await r.text();
      return res.status(200).json({ ok: true, raw: text });
    } else {
      const r = await fetch(appsScriptUrl + '?action=get', { redirect: 'follow' });
      const json = await r.json();
      return res.status(200).json(json);
    }
  } catch (e) {
    return res.status(500).json({ error: e.message });
  }
};
