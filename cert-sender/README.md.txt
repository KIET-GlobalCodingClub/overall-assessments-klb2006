# Certificate Sender

A Python tool to **automatically generate and send personalized certificates** to participants using an Excel sheet and a certificate template.

## Features
- Read participant details from Excel (Name, Roll Number, Department)
- Generate certificates with custom text and logos
- Optionally send certificates via email
- Supports multiple logos and positions

## Requirements
- Python 3.x
- Libraries: `pandas`, `Pillow`
- Excel file with participant details
- Certificate template image (JPEG/PNG)
- Optional logos

## Usage
1. Place your Excel file, certificate template, and logos in the project folder.
2. Update file paths in `send_certificates.py`.
3. Run the script:
   ```bash
   python send_certificates.py
