# rechnungMARSters

A Python tool for generating professional invoices (Rechnungen) in PDF format from Excel data.

## Overview

This tool allows you to create customer offers and convert them into beautifully formatted PDF invoices. The system reads data from an Excel file and generates a professional LaTeX-based invoice document.

## How to Use

1. **Prepare your data**: Open the Excel file located in `resources/excelData.xlsx`
2. **Fill in customer information**: Add customer details, services, and costs in the appropriate sections
3. **Generate invoice**: Run the application using:
   ```bash
   python main.py
   ```
4. **Find your PDF**: The generated invoice will be saved in the `output/` folder

## Features

- **Automatic invoice numbering**: Each invoice gets a unique number
- **Customer data management**: Automatically extracts customer information from Excel
- **Service tables**: Supports both fixed goods and employee cost services
- **Professional formatting**: Clean, corporate-style PDF output
- **Tax calculations**: Automatic VAT calculations (19%)

## Excel File Structure

The tool expects specific data in `resources/excelData.xlsx`:
- Customer information (company, contact person, address)
- Goods/services with prices and quantities
- Employee costs with hours and rates
- Tax calculations

## Important Notes

⚠️ **Version 1.0 Limitations**: This is the first version of the tool. The system expects data in specific Excel rows and columns. **Adding or removing rows may cause errors** as the tool uses fixed row indices.

A future version will include:
- Flexible row detection
- Better error handling
- More customizable templates
- Dynamic Excel parsing

## Output

The generated PDF includes:
- Company header and branding
- Customer address
- Itemized services and costs
- Tax calculations
- Professional footer with company details

## Requirements

- Python 3.x
- pandas
- XeLaTeX (for PDF generation)

## File Structure

```
resources/
├── excelData.xlsx      # Your data input file
├── executeAPP.py       # Main processing logic
├── texComps/          # Generated LaTeX components
└── rechnung.tex       # Main LaTeX template

output/                # Generated PDF files
```

---

**Note**: Use this tool carefully and double-check the generated invoices before sending them to customers. Make sure your Excel data is properly formatted in the expected structure.
