# Markdown2Word

A simple web application that converts Markdown text to Word (.docx) documents.

## Features

- **Real-time Markdown Preview**: See your formatted Markdown as you type
- **Word Document Export**: Convert your Markdown to a downloadable .docx file
- **Client-side Processing**: All conversion happens in your browser for privacy and speed
- **Support for Common Markdown Elements**:
  - Headings (H1-H4)
  - Bold and italic text
  - Lists (bulleted and numbered)
  - Inline code and code blocks

## Usage

1. Enter or paste your Markdown text in the left panel
2. See the rendered preview in the right panel
3. Click "Convert to Word" to download your document

## Technologies

- **Frontend**: Blazor WebAssembly .NET 9.0
- **Markdown Parser**: [Markdig](https://www.nuget.org/packages/Markdig/)
- **Word Generator**: [DocumentFormat.OpenXml](https://www.nuget.org/packages/DocumentFormat.OpenXml/)

## Development

### Prerequisites

- .NET 9.0 SDK

### Build

```bash
cd Markdown2Word
dotnet build
```

### Run

```bash
cd Markdown2Word
dotnet run
```

The application will be available at http://localhost:5000

## Deployment

This application is automatically deployed to GitHub Pages via GitHub Actions when changes are pushed to the main branch.

## License

This project was created based on the specifications in PRD.md.
