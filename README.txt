BANK2EXCEL — BANK STATEMENT CONVERTER

A browser-based PDF-to-Excel bank statement converter.

FEATURES
- Text-based and scanned/image-only PDF statements
- Automatic OCR with Tesseract
- Dates exported with month/day/year
- Works when transactions already include a year
- Works when transactions omit the year: detects a statement year when possible, otherwise uses the current year automatically
- Optional year override for historical statements
- Transaction preview before export
- Open file preview OR Download Excel
- 25 MB upload limit
- Docker deployment includes Tesseract, so end users install nothing

LOCAL WINDOWS
1. Run INSTALL_OCR.bat once if Tesseract is not installed.
2. Run START_HERE.bat.
3. Open http://127.0.0.1:8000.

DEPLOYMENT
The included Dockerfile installs Tesseract inside the server. Deploy the repository to a Docker-capable host such as Render. Users then only open your website and upload a PDF; they do not install Python, Tesseract, or any other software.

Render quick setup:
1. Push this folder to GitHub.
2. Create a new Web Service in Render and connect the GitHub repository.
3. Choose Docker as the runtime (the Dockerfile is included).
4. Deploy.

IMPORTANT PRIVACY NOTE
Bank statements contain sensitive financial information. Before public production use, configure your hosting provider and application with appropriate privacy, logging, retention, and security controls. This app processes the uploaded PDF in memory and does not intentionally save it as an upload file.

IMPORTANT EXTRACTION NOTE
Bank statements vary by bank, country, language, and layout. This is a generic parser and cannot guarantee perfect extraction from every possible statement. Always review transactions before relying on the Excel output.
