# ARG Converter Web

Convert CSV logs into formatted Excel and Word reports. A web-based version of the ARG Converter desktop application.

**Live Demo:** https://arg-converter-web.vercel.app

## Features

- **Drag & Drop CSV Upload** - Easy file selection with drag and drop support
- **Auto-Detect Log Type** - Automatically detects Behavior Logs, GUI Logs, or NPC Chat
- **JSON Data Flattening** - Flattens messy JSON string columns into proper rows/columns
- **Excel Export** - Generates formatted .xlsx files with auto-fitted columns
- **Word Export** - Creates professional .doc documents with grouped player data
- **Data Preview** - View first 10 rows before exporting

## Tech Stack

- **Framework:** Next.js 14
- **Language:** TypeScript
- **Styling:** Custom CSS with Sky Blue theme
- **CSV Parsing:** PapaParse
- **Excel Generation:** SheetJS (xlsx)
- **File Downloads:** FileSaver.js
- **Deployment:** Vercel

## Getting Started

### Local Development

```bash
# Clone the repository
git clone https://github.com/AfdalSusilo/arg_converter_web.git

# Navigate to project directory
cd arg_converter_web

# Install dependencies
npm install

# Run development server
npm run dev

# Open http://localhost:3000
```

### Deploy to Vercel

```bash
# Install Vercel CLI
npm install -g vercel

# Deploy
vercel
```

Or click the button below:

[![Deploy with Vercel](https://vercel.com/button)](https://vercel.com/new/clone?repository-url=https://github.com/AfdalSusilo/arg_converter_web)

## Usage

1. **Select CSV File** - Drag and drop or click to browse for your CSV file
2. **Choose Log Type** - Select auto-detect or manually choose the log type
3. **Preview Data** - View the parsed data in the preview table
4. **Export** - Click either "Export to Excel" or "Export to Word" button

## Log Types

- **Behavior Logs** - Mouse events, keystrokes, and user interactions
- **GUI Logs** - Input data and widget interactions
- **NPC Chat** - Chat history and NPC conversations

## Related Projects

- [ARG Converter Desktop App](https://github.com/AfdalSusilo/arg_converter_app) - Windows desktop application version

## License

MIT License
