# Google Document Converter

A Python application that allows you to convert Google Sheets and Google Docs documents to PDF and DOCX formats.

## Features

- **Google Sheets and Google Docs Support**: Convert data from both Google Sheets and Google Docs to PDF and DOCX formats.
- **Credential Management**: Select and validate Google API credentials through the interface.
- **Preview Functionality**: Preview the document data before generating the output file.
- **Custom Document Formatting**: Add titles and format the output documents.
- **Template Support**: Use custom DOCX templates for document generation.
- **User-Friendly Interface**: Simple and intuitive Tkinter-based UI.

## Requirements

- Python 3.6+
- Required Python packages (install using `pip install -r requirements.txt`):
  - gspread
  - pandas
  - fpdf
  - google-api-python-client
  - google-auth
  - google-auth-oauthlib
  - docxtpl

## Installation

1. Clone this repository or download the source code.
2. Install the required packages:
   ```
   pip install -r requirements.txt
   ```
3. Obtain Google API credentials:
   - Go to the [Google Cloud Console](https://console.cloud.google.com/)
   - Create a new project
   - Enable the Google Sheets API and Google Docs API
   - Create a service account and download the JSON credentials file

## Usage

1. Run the application:
   ```
   python main.py
   ```
2. Select your Google API credentials file using the "Select Credentials File" button.
3. Validate your credentials using the "Validate Credentials" button.
4. Choose the document type (Google Sheets or Google Docs).
5. For Google Sheets, select the access method (by name, key, or URL).
6. Enter the document identifier (URL, key, or name for Sheets; document ID for Docs).
7. Optionally, select a document template for DOCX generation.
8. Optionally, enter a title for the output document.
9. Click "Preview and Generate" to process the document.
10. Review the data in the preview window.
11. Click "Generate Document" to save the file to your computer in your chosen format (PDF or DOCX).

## Architecture

The application follows an Object-Oriented Programming (OOP) approach with four main classes:

1. **GoogleConnection**: Handles authentication and credential validation.
2. **GoogleSheetsReader**: Reads data from Google Sheets.
3. **DocumentGenerator**: Generates and formats documents (PDF and DOCX) from the data.
4. **GoogleToDocApp**: Controls the Tkinter UI/UX.

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Acknowledgments

- [gspread](https://github.com/burnash/gspread) for Google Sheets API access
- [FPDF](https://pyfpdf.readthedocs.io/en/latest/) for PDF generation
- [pandas](https://pandas.pydata.org/) for data manipulation
- [docxtpl](https://docxtpl.readthedocs.io/en/latest/) for DOCX template processing
