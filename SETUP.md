# Setup

## 1. Install

```bash
cd excel-voice-agent
npm install
```

## 2. Create R2 bucket

```bash
wrangler r2 bucket create sheet-voice-agent
```

## 3. Set OpenAI API key

```bash
wrangler secret put OPENAI_API_KEY
```

## 4. Upload your spreadsheet (optional)

Export your Google Sheet as .xlsx (File → Download → Microsoft Excel), then upload to R2:

```bash
wrangler r2 object put sheet-voice-agent/spreadsheet.xlsx --file ./spreadsheet.xlsx
```

If you skip this, an empty spreadsheet is created automatically on first use.

## 5. Dev (local)

```bash
npm run dev
```

## 6. Deploy

```bash
npm run deploy
```

That's it. Cloudflare gives you a `*.workers.dev` URL.

## How it works

- Each browser connection gets its own Durable Object instance
- The DO proxies audio between the browser and OpenAI Realtime API
- Tool calls (read/write cell, etc.) are intercepted by the DO and executed against the xlsx file in R2
- Static files (`public/`) are served via Workers Assets
