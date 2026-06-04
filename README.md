# 台股追蹤 Portfolio Tracker PWA

台灣股票投資組合追蹤 App，支援台股＋美股即時報價，可安裝為 PWA。

## 功能

- 📊 持股明細：即時市值、未實現損益
- 🟢 **自動報價**：透過 Yahoo Finance API，每 5 分鐘自動更新（支援台股 .TW、美股）
- 🏆 歷史績效：歷年已實現損益、各股報酬率
- 📋 交易記錄：新增/編輯/刪除，支援 CSV / Excel 批量匯入
- 💾 PWA：可安裝到手機桌面，離線使用快取版本
- 👤 多帳號：各帳號獨立資料，券商手續費自訂

## 本地開發

```bash
npm install
npm run dev
```

## 部署到 Vercel

### 方法一：Vercel CLI

```bash
npm install -g vercel
vercel
```

### 方法二：GitHub + Vercel Dashboard

1. 將此專案 push 到 GitHub repo
2. 到 [vercel.com](https://vercel.com) → New Project → Import 你的 repo
3. Framework: **Vite**（Vercel 會自動偵測）
4. Build Command: `npm run build`
5. Output Directory: `dist`
6. 點 Deploy

部署完成後，Vercel 會給你一個 `https://xxx.vercel.app` 的網址。

## 報價說明

- 報價來源：Yahoo Finance（免費，無需 API Key）
- 台股代號：直接輸入 `2330`（系統自動加 `.TW` 向 Yahoo 查詢）
- 美股代號：直接輸入 `AAPL`、`NVDA`、`TSM` 等
- 更新頻率：進入 App 時立即抓取，之後每 5 分鐘自動更新
- 盤後時間可能顯示最後收盤價

## 檔案結構

```
tw-stock-pwa/
├── api/
│   └── quote.js          # Vercel Serverless Function（Yahoo Finance Proxy）
├── public/
│   └── favicon.svg
├── src/
│   ├── App.jsx            # 主程式
│   └── main.jsx
├── index.html             # PWA 設定
├── vite.config.js         # Vite + PWA Plugin
├── vercel.json
└── package.json
```

## PWA 安裝

部署後，用手機 Chrome / Safari 開啟網址：
- **Android**：點右上角「⋮」→「加到主畫面」
- **iOS**：點底部「分享」→「加入主畫面」
- **桌面 Chrome**：網址列右側有安裝圖示

## 注意事項

- 資料儲存在瀏覽器 `localStorage`，清除瀏覽器資料會遺失
- 若要跨裝置同步，請先用「匯出紀錄」備份 CSV，再於新裝置「匯入交易」
- Yahoo Finance 免費 API 可能有請求限制，若無法取得報價請稍後重試
