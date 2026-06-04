export default async function handler(req, res) {
  const { symbols } = req.query;
  if (!symbols) return res.status(400).json({ error: "no symbols" });
  try {
    const url = `https://query1.finance.yahoo.com/v7/finance/quote?symbols=${encodeURIComponent(symbols)}&fields=regularMarketPrice,regularMarketPreviousClose,regularMarketChangePercent,shortName,currency,marketState`;
    const r = await fetch(url, {
      headers: {
        "User-Agent": "Mozilla/5.0",
        "Accept": "application/json",
      },
      cache: "no-store",
    });
    const data = await r.json();
    const quotes = {};
    (data?.quoteResponse?.result || []).forEach((q) => {
      quotes[q.symbol] = {
        price: q.regularMarketPrice,
        prevClose: q.regularMarketPreviousClose,
        changePct: q.regularMarketChangePercent,
        name: q.shortName,
        currency: q.currency,
        marketState: q.marketState,
      };
    });
    res.setHeader("Cache-Control", "no-store, no-cache, must-revalidate");
    res.setHeader("Pragma", "no-cache");
    res.json({ quotes });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
}
