# Automated Invoice Processor

A local Streamlit application that processes PDF invoices using OCR, automatically renames and sorts them, and generates an Excel report with extracted invoice data and net values.

## Demo

Demo video: **[Add your demo video link here]**

You can upload the video to YouTube as unlisted, Google Drive, LinkedIn, or include a short MP4/GIF inside a `demo` folder.

## Features

- Upload multiple PDF invoices directly from the interface
- Extract invoice information using OCR:
  - Seller name
  - Buyer name
  - Invoice number
  - Invoice date
  - Net invoice value
- Automatically classify invoices into:
  - `FacturiClienti`
  - `FacturiFurnizori`
- Rename invoices using the following format:

```text
Seller_InvoiceDate_InvoiceNumber.pdf
```

- Generate an Excel workbook with a separate sheet for each seller
- Calculate the total net invoice value for every sheet
- Download:
  - The generated Excel report
  - A ZIP archive containing the processed and sorted invoices
- Display processing results, OCR output, logs, and errors in the interface

## Technologies Used

- Python
- Streamlit
- Pandas
- Tesseract OCR
- Poppler
- pdf2image
- pytesseract
- openpyxl
- Pillow

## Project Structure

```text
FacturiApp/
│
├── app.py
├── requirements.txt
├── README.md
└── .gitignore
```

## Requirements

The application requires the following software to be installed locally:

### Python

Python 3.10 or Python 3.11 is recommended.

### Tesseract OCR

Install Tesseract OCR and select the path to `tesseract.exe` inside the application.

Example:

```text
C:\Program Files\Tesseract-OCR\tesseract.exe
```

Romanian and English OCR language data are recommended because the application uses:

```text
ron+eng
```

### Poppler

Install Poppler for Windows and select the path to its `bin` folder inside the application.

Example:

```text
C:\poppler\Library\bin
```

## Installation

Clone the repository:

```bash
git clone https://github.com/YOUR-USERNAME/YOUR-REPOSITORY.git
cd YOUR-REPOSITORY
```

Create and activate a Conda environment:

```bash
conda create -n facturiapp python=3.11 -y
conda activate facturiapp
```

Install the required Python packages:

```bash
pip install -r requirements.txt
```

## Running the Application

Start the application with:

```bash
streamlit run app.py
```

Streamlit will open the application automatically in your browser.

If it does not open automatically, use the local address displayed in the terminal, usually:

```text
http://localhost:8501
```

## How to Use

1. Open the application.
2. Enter the name of your company.
3. Select the local paths for Tesseract OCR and Poppler.
4. Upload one or more PDF invoices.
5. Click **Process invoices**.
6. Review the extracted information.
7. Download:
   - `ValoriNet.xlsx`
   - `FacturiPrelucrate.zip`

## Invoice Classification Logic

The application checks whether the configured company is the seller.

- If the configured company is the seller, the invoice is placed in `FacturiClienti`.
- Otherwise, the invoice is placed in `FacturiFurnizori`.

The invoice filename is always generated using the seller name:

```text
Seller_InvoiceDate_InvoiceNumber.pdf
```

## Excel Output

The generated Excel file contains a separate sheet for each seller.

Each sheet includes:

- Invoice date
- Invoice number
- Total net value
- A formula that calculates the total value of all invoices in that sheet

## Important Notes

- The application currently runs locally because it depends on external installations of Tesseract OCR and Poppler.
- OCR accuracy depends on the invoice layout, scan quality, and image resolution.
- The crop coordinates used for OCR are optimized for the invoice format used during development. Other invoice layouts may require adjustments.
- Do not upload confidential invoices, generated Excel files, ZIP archives, or personal company documents to a public repository.
- Local Tesseract and Poppler paths may be different on another computer.

## Possible Future Improvements

- Support for multiple invoice layouts
- Automatic detection of invoice fields without fixed crop coordinates
- Manual correction of OCR results before export
- Duplicate invoice detection
- VAT and gross value extraction
- Export to CSV
- Database integration
- Docker support for easier deployment
- Public cloud deployment

## Privacy

Invoice processing is performed locally on the user's computer. Uploaded files are temporarily processed by the application and are not sent to an external OCR service.

## Author

Developed as a Python automation and document-processing project for portfolio and CV purposes.
