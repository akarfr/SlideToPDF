# SlideToPDF 📄✨

A minimalist, highly scannable, and modern cross-platform desktop application designed to batch-convert Microsoft PowerPoint (`.pptx`) presentations into high-quality PDF documents using native system engines.

---

## Key Features

- **Drag & Drop Workspace**: Drop folders or multiple PowerPoint files directly into the drop zone for conversion.
- **Multi-File Browser**: Open a native dialog to select and queue multiple `.pptx` files.
- **Custom Output Destinations**: Choose exactly where your converted PDFs should go, or default to the input files' directories.
- **Real-Time Progress Tracking**: Watch live progress badges (e.g., `Converting (2/5): intro.pptx`) update dynamically.
- **Light & Dark Mode**: A sleek theme toggle that remembers your preferences using local storage.
- **Zero Cloud Dependencies**: Files are converted locally on your machine, ensuring total data privacy.

---

## Native Conversion Engines

SlideToPDF leverages the native engines installed on your operating system for optimal layout fidelity:

- **Windows (`win32`)**: Dynamically binds to the local **Microsoft PowerPoint COM Interface** via silent background PowerShell scripting.
- **macOS (`darwin`)**: Converts presentations silently utilizing the local headless **LibreOffice CLI** (`soffice`).

---

## Installation & Getting Started

### Prerequisites

1. Make sure you have [Node.js](https://nodejs.org) installed.
2. Ensure you have the appropriate native engine installed:
   - **Windows**: Microsoft PowerPoint must be installed on the machine.
   - **macOS**: [LibreOffice](https://www.libreoffice.org/) must be installed and accessible in the system path.

### Development Setup

1. **Clone the repository:**
   ```bash
   git clone https://github.com/akarfr/SlideToPDF.git
   cd SlideToPDF
   ```

2. **Install dependencies:**
   ```bash
   npm install
   ```

3. **Start the application:**
   ```bash
   npm start
   ```

---

## Development Stack

- **Core**: Electron, Node.js
- **Frontend**: Vanilla HTML5, JavaScript
- **Styling**: Vanilla CSS (CSS Custom Properties for Light/Dark themes)

---

## License

This project is licensed under the MIT License - see the LICENSE file for details.
