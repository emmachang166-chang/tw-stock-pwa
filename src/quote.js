// api/quote.js — Vercel Serverless Function
// Proxies Yahoo Finance v8 quote API to avoid CORS issues in the browser.
// Usage: GET /api/quote?symbols=2330.TW,2454.TW,AAPL

export default async function handler(req, res) {
  // CORS headers — allow any origin (this is your own proxy)
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
  if (req.method === "OPTIONS") { res.status(200).end(); return; }

  const { symbols } = req.query;
  if (!symbols) {
    return res.status(400).json({ error: "symbols parameter required" });
  }

  // Yahoo Finance v8 endpoint — free, no API key needed
  const url = `https://query1.finance.yahoo.com/v8/finance/spark?symbols=${encodeURIComponent(symbols)}&range=1d&interval=5m&includePrePost=false&corsDomain=finance.yahoo.com`;
  // Fallback: v7 quoteSummary with price module
  const url2 = `https://query2.finance.yahoo.com/v7/finance/quote?symbols=${encodeURIComponent(symbols)}&fields=regularMarketPrice,regularMarketPreviousClose,regularMarketChangePercent,shortName,longName,currency,marketState`;

  try {
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), 8000);

    let data = null;
    let source = "v7";

    // Try v7 first (more reliable for current price)
    try {
      const r = await fetch(url2, {
        signal: controller.signal,
        headers: {
          "User-Agent": "Mozilla/5.0 (compatible; portfolio-tracker/1.0)",
          "Accept": "application/json",
        }
      });
      clearTimeout(timeout);
      if (r.ok) {
        const json = await r.json();
        const quotes = json?.quoteResponse?.result;
        if (quotes && quotes.length > 0) {
          // Normalise to { SYMBOL: { price, prevClose, changePct, name, currency, marketState } }
          const result = {};
          quotes.forEach(q => {
            result[q.symbol] = {
              price:       q.regularMarketPrice,
              prevClose:   q.regularMarketPreviousClose,
              changePct:   q.regularMarketChangePercent,
              name:        q.shortName || q.longName || q.symbol,
              currency:    q.currency || "TWD",
              marketState: q.marketState || "CLOSED",
            };
          });
          data = result;
          source = "v7";
        }
      }
    } catch (_) { /* fall through */ }

    if (!data) {
      return res.status(502).json({ error: "Failed to fetch quotes from Yahoo Finance" });
    }

    // Cache for 5 minutes on CDN edge
    res.setHeader("Cache-Control", "public, s-maxage=300, stale-while-revalidate=60");
    return res.status(200).json({ source, quotes: data, fetchedAt: new Date().toISOString() });

  } catch (err) {
    console.error("Quote proxy error:", err);
    return res.status(500).json({ error: err.message });
  }
}
