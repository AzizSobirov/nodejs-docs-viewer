# Node.js Document Viewer

A lightweight Node.js server that converts and previews various document formats (PDF, DOC, DOCX, XLS, XLSX, PPT, PPTX, CSV) in the browser. Built with Express.js and LibreOffice for document conversion.

## Features

- 📄 **Multiple Format Support**: PDF, DOC, DOCX, XLS, XLSX, PPT, PPTX, CSV
- 🚀 **Fast Caching**: Converted documents are cached for improved performance
- 🎨 **Styled Excel Preview**: Preserves formatting, colors, and styles from Excel files
- ⚙️ **Configurable**: Control max file size and cache duration via environment variables
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
```

### Environment Variables

| Variable | Description | Default |
|----------|-------------|---------|
| `PORT` | Server port | `3000` |
| `MAX_FILE_SIZE` | Maximum file size in MB | `200` |
| `CACHE_TIME` | Cache duration in seconds | `86400` (24 hours) |

## Supported Formats

| Format | Extensions | Preview Method |
|--------|-----------|----------------|
| PDF | `.pdf` | Direct display |
| Word | `.doc`, `.docx` | Convert to PDF |
| Excel | `.xls`, `.xlsx` | HTML with styles |
| PowerPoint | `.ppt`, `.pptx` | Convert to PDF |
| CSV | `.csv` | HTML table |

## API

### GET /preview

Preview a document from a URL.

**Query Parameters:**
- `src` (required): HTTP(S) URL of the document to preview

**Response:**
- For PDFs: Returns the PDF file directly
- For Excel: Returns styled HTML preview
- For other formats: Returns converted PDF or HTML

**Example:**
```bash
curl "http://localhost:3000/preview?src=https://example.com/document.pdf"
```

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
├── index.js          # Main application file
├── cache/             # Cached converted documents
├── package.json       # Dependencies and scripts
├── .env               # Environment configuration (create this)
├── .env.example       # Example environment variables
└── LICENSE            # MIT License
```

## Error Handling

The server handles various error scenarios:
- Invalid or missing `src` parameter
- Non-HTTP(S) URLs
- Files exceeding `MAX_FILE_SIZE`
- Conversion failures (falls back to raw file download)
- Network errors

## Security Considerations

- Only HTTP(S) URLs are accepted
- File size is limited by `MAX_FILE_SIZE`
- Downloaded files are validated before processing
- Cached files are hashed to prevent conflicts

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
- [ExcelJS](https://github.com/exceljs/exceljs) for Excel file handling
