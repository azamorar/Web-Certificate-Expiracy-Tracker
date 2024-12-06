# SSL Tracker

A Python-based application to track and monitor SSL certificates for specified domains. The program verifies SSL certificate details, including expiration dates, and exports the results to Excel. It also features a graphical user interface (GUI) for ease of use.

## Features

- **SSL Certificate Monitoring**:
  - Retrieves SSL certificate details such as expiration date and common name (CN).
  - Verifies the status of certificates for multiple domains.

- **Data Export**:
  - Exports certificate details to an Excel file for reporting or record-keeping.

- **Graphical Interface**:
  - User-friendly interface built with `tkinter`.
  - Provides options to input domain lists and manage exports.

- **Multi-threading**:
  - Simultaneously processes multiple domains for improved efficiency.

## Prerequisites

Ensure the following software and libraries are installed:

- **Python**: Version 3.8 or higher.
- Python Libraries:
  - `openpyxl`
  - `pandas`
  - `tkinter`

Install the required Python libraries using pip:

```bash
pip install openpyxl pandas
