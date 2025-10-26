# Node.js Document Viewer

A lightweight Node.js server that converts and previews various document formats (PDF, DOC, DOCX, XLS, XLSX, PPT, PPTX, CSV) in the browser. Built with Express.js, LibreOffice for document conversion, and Luckysheet for interactive Excel viewing.

## Features

- 📄 **Multiple Format Support**: PDF, DOC, DOCX, XLS, XLSX, PPT, PPTX, CSV
- 🚀 **Fast Caching**: Converted documents are cached for improved performance
- 📊 **Interactive Excel Viewer**: Full-featured spreadsheet interface powered by Luckysheet
- ⏳ **Loading Indicators**: Beautiful spinner for PDF and presentation loading
- 📥 **Unsupported Files**: Elegant download page for files that can't be previewed
- 🛡️ **Security Features**: Domain whitelist and file format restrictions
- ⚙️ **Configurable**: Control max file size, cache duration, allowed domains, and formats
- 🔒 **Safe**: Only accepts HTTP(S) URLs with configurable file size limits

## Prerequisites

- Node.js (v14 or higher)
- LibreOffice installed on your system

### Installing LibreOffice

**macOS:**
```bash
brew install --cask libreoffice
```

**Ubuntu/Debian:**
```bash
sudo apt-get install libreoffice
```

**Windows:**
Download and install from [LibreOffice official website](https://www.libreoffice.org/download/download/)

## Installation

1. Clone the repository:
```bash
git clone https://github.com/AzizSobirov/nodejs-docs-viewer.git
cd nodejs-docs-viewer
```

2. Install dependencies:
```bash
npm install
```

3. Start the server:
```bash
npm start
```

## Usage

### Basic Usage

Start the server and navigate to:
```
http://localhost:3000/preview?src=<URL_TO_DOCUMENT>
```

### Example

```
http://localhost:3000/preview?src=https://example.com/document.pdf
http://localhost:3000/preview?src=https://example.com/spreadsheet.xlsx
http://localhost:3000/preview?src=https://example.com/presentation.pptx
```

## Configuration

Create a `.env` file in the root directory to customize settings:

```env
# Server port (default: 3000)
PORT=3000

# Maximum file size in MB (default: 200)
MAX_FILE_SIZE=200

# Cache duration in seconds (default: 86400 = 24 hours)
CACHE_TIME=86400

# Allowed domains (comma-separated, leave empty to allow all)
ALLOWED_DOMAINS=example.com,docs.google.com,drive.google.com

# Accepted file formats (comma-separated)
ACCEPTED_FORMATS=.pdf,.doc,.docx,.xls,.xlsx,.ppt,.pptx,.csv,.odt,.ods,.odp
```

### Environment Variables

| Variable | Description | Default |
|----------|-------------|---------|
| `PORT` | Server port | `3000` |
| `MAX_FILE_SIZE` | Maximum file size in MB | `200` |
| `CACHE_TIME` | Cache duration in seconds | `86400` (24 hours) |
| `ALLOWED_DOMAINS` | Comma-separated list of allowed domains (empty = allow all) | `` (all allowed) |
| `ACCEPTED_FORMATS` | Comma-separated list of accepted file extensions | `.pdf,.doc,.docx,.xls,.xlsx,.ppt,.pptx,.csv,.odt,.ods,.odp` |

### Security Configuration Examples

**Allow specific domains:**
```env
ALLOWED_DOMAINS=mycompany.com,docs.google.com
# This allows mycompany.com, *.mycompany.com, docs.google.com, *.docs.google.com
```

**Restrict to only PDFs and Excel:**
```env
ACCEPTED_FORMATS=.pdf,.xlsx,.xls
```

**Allow all domains (default):**
```env
ALLOWED_DOMAINS=
```

## Supported Formats

| Format | Extensions | Preview Method | Features |
|--------|-----------|----------------|----------|
| PDF | `.pdf` | Direct display with loader | Spinner animation during load |
| Word | `.doc`, `.docx` | Convert to PDF with loader | LibreOffice conversion + PDF viewer |
| Excel | `.xls`, `.xlsx` | Interactive Luckysheet viewer | Full Excel-like interface with styles |
| PowerPoint | `.ppt`, `.pptx` | Convert to PDF with loader | LibreOffice conversion + PDF viewer |
| CSV | `.csv` | HTML table | Clean table display |
| Unsupported | Others | Download page | Elegant UI with download button |

## API

### GET /preview

Preview a document from a URL.

**Query Parameters:**
- `src` (required): HTTP(S) URL of the document to preview

**Response:**
- **PDFs**: HTML page with loading spinner, then displays PDF in iframe
- **Excel files**: Interactive Luckysheet spreadsheet viewer with full formatting
- **Word/PowerPoint**: Converted to PDF with loading animation
- **CSV**: Clean HTML table view
- **Unsupported files**: Beautiful download page with file info and download button

**Example:**
```bash
curl "http://localhost:3000/preview?src=https://example.com/document.pdf"
```

### GET /pdf/:key

Serves cached PDF files (used internally by the PDF loader).

### GET /download/:key

Downloads the original file for unsupported formats.

## Caching

The server caches converted documents to improve performance:
- Cached files are stored in the `cache/` directory
- Cache validity is determined by the `CACHE_TIME` environment variable
- Expired cache entries are automatically regenerated

## Development

Run the server in development mode:
```bash
npm run dev
```

## Project Structure

```
.
├── index.js           # Main application file
├── server.js          # Alternative server file
├── cache/             # Cached converted documents
├── package.json       # Dependencies and scripts
├── .env               # Environment configuration (create this)
├── .env.example       # Example environment variables
├── LICENSE            # MIT License
└── README.md          # This file
```

## User Experience Features

### PDF Loading Animation
- Elegant spinner with "Loading document..." text
- Smooth fade-out when document is ready
- Works for PDF, DOC, DOCX, PPT, PPTX files

### Excel Interactive Viewer (Luckysheet)
- Full Excel-like spreadsheet interface
- Preserves all formatting, colors, borders, and styles
- Support for multiple sheets with tabs
- Formula display and calculated values
- Merged cells, frozen rows/columns
- Read-only mode for safe viewing
- Responsive design

### Unsupported File Handler
- Modern gradient-styled page
- File icon and information display
- Download button with hover effects
- Clean, professional design
- Works for any file type that can't be previewed

## Error Handling

The server handles various error scenarios:
- Invalid or missing `src` parameter
- Non-HTTP(S) URLs
- Files exceeding `MAX_FILE_SIZE`
- Conversion failures (shows unsupported file page with download)
- Network errors
- Cache validation and expiration

## How It Works

1. **Request**: User provides document URL via `?src=` parameter
2. **Cache Check**: Server checks for valid cached version
3. **Download**: If not cached, downloads file from URL
4. **Detection**: Automatically detects file type
5. **Processing**:
   - **PDF**: Wraps in loader HTML → displays in iframe
   - **Excel**: Parses with LuckyExcel → renders with Luckysheet
   - **Word/PowerPoint**: Converts to PDF with LibreOffice → displays with loader
   - **CSV**: Converts to HTML table
   - **Others**: Tries conversion, falls back to download page
6. **Caching**: Stores result for future requests
7. **Delivery**: Returns processed document or download page

## Security Considerations

### Domain Whitelist
- Control which domains can be accessed via `ALLOWED_DOMAINS`
- Empty list allows all domains (development mode)
- Supports exact domain and subdomain matching
- Returns 403 error for non-whitelisted domains

### File Format Restrictions
- Only specified file formats in `ACCEPTED_FORMATS` are processed
- Non-accepted formats show download page instead of preview
- Prevents processing of potentially malicious file types

### Built-in Protections
- Only HTTP(S) URLs are accepted
- File size is limited by `MAX_FILE_SIZE`
- Downloaded files are validated before processing
- Cached files are hashed to prevent conflicts
- File type detection prevents extension spoofing

## Performance Tips

1. Adjust `CACHE_TIME` based on your use case
2. Increase `MAX_FILE_SIZE` only if necessary
3. Regularly clean up the `cache/` directory
4. Use a reverse proxy (nginx) for production deployments

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## Author

Aziz Sobirov

## Acknowledgments

- [LibreOffice](https://www.libreoffice.org/) for document conversion
- [Express.js](https://expressjs.com/) for the web framework
- [Luckysheet](https://github.com/mengshukeji/Luckysheet) for interactive Excel viewing
- [LuckyExcel](https://github.com/mengshukeji/Luckyexcel) for Excel file parsing
- [ExcelJS](https://github.com/exceljs/exceljs) for Excel file handling
