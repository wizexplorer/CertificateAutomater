# CertificateAutomater

CertificateAutomater is a Windows desktop utility that generates certificate Excel files and their PDFs from a template Excel file and a directory of certificate Excel files.

## Features

- Choose a template Excel file
- Choose a folder containing certificate Excel files
- Generate completed certificate files from the template
- Export generated certificates to PDF
- Supports `.xls` and `.xlsx` files

## Requirements

- Windows 10/11
- Microsoft Excel installed

Microsoft Excel is required because PDF export uses Excel automation.

## Installation

1. Go to the latest release.
2. Download `CertificateAutomaterSetup.exe`.
3. Run the installer.
4. Open `CertificateAutomater.exe` from the Start Menu.

## Usage

1. Click **Browse** beside **Choose Template file**.
2. Select the template Excel file.
3. Click **Browse** beside **Choose certificate directory**.
4. Select the folder containing certificate files.
5. Click **Generate Certificates**.

Generated Excel files are saved in the template file's directory.

PDF files are saved in a `PDF` folder in the template file's directory.

## Notes

- Only the first worksheet is processed.
- Files containing `template` or `000` in the filename are skipped.
- Only `.xls` and `.xlsx` files are processed.