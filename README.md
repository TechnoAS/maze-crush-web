# Maze Crush Production Report Parser

An intelligent production report parser that extracts key performance indicators (KPIs) from WhatsApp daily reports and exports them to Microsoft Excel with professional formatting.

## Features

- 📊 **Smart Text Parsing** – Extract 80+ structured KPIs from unstructured WhatsApp reports
- 📈 **Multiple Export Formats** – CSV, JSON, and professionally formatted Excel (.xlsx)
- 📁 **Excel Integration** – Append data to existing workbooks, sorted by date
- 🎨 **Modern Design** – Microsoft Office 365-style formatting with blue headers
- ⚡ **Production-Ready** – Built with Next.js 14, React, and TypeScript

## Getting Started

### Prerequisites
- Node.js 18+
- npm or yarn

### Installation

```bash
git clone <repository-url>
cd maze-crush-web
npm install
```

### Development

```bash
npm run dev
```

Open [http://localhost:3000](http://localhost:3000) to view the application.

### Build for Production

```bash
npm run build
npm start
```

## Usage

1. **Parse Report** – Paste your daily production report text
2. **Review Data** – Browse extracted KPIs by category
3. **Export** – Download as Excel, CSV, or JSON
4. **Append to Excel** – Add data to existing workbooks

## Deployment

### Vercel

```bash
npm install -g vercel
vercel
```

### GitHub

```bash
git add .
git commit -m "Initial commit"
git push origin main
```

## Tech Stack

- **Framework** – Next.js 14 with TypeScript
- **Styling** – CSS with modern design tokens
- **Excel** – XLSX/SheetJS for professional generation
- **File I/O** – File System Access API

## Project Structure

```
app/
  page.tsx              # Main application
  api/parse/route.ts    # Report parsing API
  components/
lib/
  parser.ts             # Parsing logic
  exportExcel.ts        # Export functionality
  appendExcel.ts        # Append functionality
```

## API

### POST `/api/parse`

Extracts KPIs from production report text.

**Request:** `{ "text": "Daily Production Report..." }`

**Response:** `{ "data": { ...kpis }, "csv": "..." }`

## License

MIT
