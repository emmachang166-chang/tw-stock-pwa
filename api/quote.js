export default async function handler(req, res) {
  const { symbols } = req.query;
  if (!symbols) return res.status(400).json({ error: "no symbols" });
  try {
    const symList = symbols.split(",");
    const quotes = {};
    await Promise.all(symList.map(async (sym) => {
      try {
        const url = `https://query2.finance.yahoo.com/v8/finance/chart/${encodeURIComponent(sym)}?interval=1d&range=1d`;
        const r = await fetch(url, {
          headers: {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
            "Accept": "application/json",
            "Referer": "https://finance.yahoo.com",
          },
        });
        const data = await r.json();
        const meta = data?.chart?.result?.[0]?.meta;
        if (meta) {
          quotes[sym] = {
            price: meta.regularMarketPrice,
            prevClose: meta.previousClose || meta.chartPreviousClose,
            changePct: meta.previousClose
              ? ((meta.regularMarketPrice - meta.previousClose) / meta.previousClose) * 100
              : null,
            name: meta.shortName || sym,
            currency: meta.currency,
            marketState: meta.marketState,
          };
        }
      } catch (_) {}
    }));
    res.setHeader("Cache-Control", "no-store");
    res.json({ quotes });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
}
