A Python based automation system that collects sales and pricing data from CSV files, calculates pending payments and pending dispatch totals, generates an Excel report with detailed tables, and sends a styled HTML email with inlined pie charts showing payment and dispatch statuses. The system includes logging for monitoring, uses in-memory charts to avoid temporary files, and provides enterprise-level email reports for management.


Environment Setup

    python -m ensurepip --upgrade
    python -m pip install --upgrade pip

    python -m pip install pandas
    python -m pip install reportlab
    python -m pip install matplotlib openpyxl xlsxwriter



