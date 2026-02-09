import logging
import pandas as pd
from datetime import date
import os



#Logging Setup
os.makedirs("Logs", exist_ok=True)  #create log folder if not exist

logging.basicConfig(
    filename="Logs/automation.log",
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)
logging.info("Automation script started")







#Load Data
try:
    sale = pd.read_csv("Data/Sale.csv")
    price = pd.read_csv("Data/Price.csv")
    logging.info("Data loaded successfully")
except Exception as e:
    logging.error(f"Data loading failed: {e}")
    raise





#Data Calculation
try:
    df = sale.merge(price, on="product_id", how="left")
    df["total_price"] = df["quantity"] * df["price"]

    # 1. Order wise detail
    order_wise = (
        df.groupby("oder_id")
        .agg(
            total_price=("total_price", "sum"),
            payment_state=("payment_state", "first"),
            departure_state=("departure_state", "first"),
        )
        .reset_index()
    )

    # 2. Pending payments
    pending_payments = df[df["payment_state"] == "Unpaid"][
        ["oder_id", "product_name", "quantity", "total_price", "payment_state"]
    ]

    # 3. Pending departure
    pending_departure = df[df["departure_state"] == "Not Dispatch"][
        ["oder_id", "product_name", "quantity", "total_price", "payment_state", "departure_state"]
    ]

    logging.info("Data calculation successful")
except Exception as e:
    logging.error(f"Data calculation failed: {e}")
    raise





#Generate Report
try:
    report_path = f"Reports/daily_report_{date.today()}.xlsx"
    os.makedirs("Reports", exist_ok=True)

    if os.path.exists(report_path):
        os.remove(report_path)

    with pd.ExcelWriter(report_path, engine="xlsxwriter") as writer:
        workbook  = writer.book #entire Excel file
        worksheet = writer.sheets["Report"] #single sheet inside the workbook

        #Formats
        title_format = workbook.add_format({    #Formats are defined at workbook level then can apply it to any sheet or any cell
            "bold": True,
            "align": "center",
            "valign": "vcenter",
            "font_size": 12
        })


        #TABLE 1
        order_wise.to_excel(writer, sheet_name="Report", startrow=1, index=False)
        cols1 = [{"header": col} for col in order_wise.columns] #Take all column names from order_wise DataFrame
        rows1 = len(order_wise) #Take user count

        worksheet.merge_range(0, 0, 0, len(cols1)-1, "OrderWiseDetails", title_format)  #marge cell and add title

        worksheet.add_table(    #Add table below the title
            1, 0,
            rows1, len(cols1)-1,
            {"columns": cols1, "style": "Table Style Medium 9"}
        )

        start_row = rows1 + 4


        #TABLE 2
        pending_payments.to_excel(writer, sheet_name="Report", startrow=start_row+1, index=False)
        cols2 = [{"header": col} for col in pending_payments.columns]
        rows2 = len(pending_payments)

        worksheet.merge_range(start_row, 0, start_row, len(cols2)-1,"PendingPayments", title_format)

        worksheet.add_table(
            start_row+1, 0,
            start_row+rows2, len(cols2)-1,
            {"columns": cols2, "style": "Table Style Medium 9"}
        )

        start_row = start_row + rows2 + 4


        #TABLE 3
        pending_departure.to_excel(writer, sheet_name="Report", startrow=start_row+1, index=False)
        cols3 = [{"header": col} for col in pending_departure.columns]
        rows3 = len(pending_departure)

        worksheet.merge_range(
            start_row, 0, start_row, len(cols3)-1,
            "PendingDeparture", title_format
        )

        worksheet.add_table(
            start_row+1, 0,
            start_row+rows3, len(cols3)-1,
            {"columns": cols3, "style": "Table Style Medium 9"}
        )

    logging.info("Report generated successfully")

except Exception as e:
    logging.error(f"Report generation failed: {e}")
    raise
