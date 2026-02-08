import logging
import pandas as pd
import json
from datetime import date
import smtplib
from email.message import EmailMessage

 




logging.basicConfig(
    filename="Logs/automation.log",
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)

logging.info("Automation script started")






try:
    sale = pd.read_csv("Data/Sale.csv")
    price = pd.read_csv("Data/Price.csv")

    logging.info("Data loaded successfully")

except Exception as e:
    logging.error(f"Data loading failed: {e}")
    raise






            
try:
    # merge sales with price
    df = sale.merge(price, on="product_id", how="left")
                
                
    # calculate total price per row
    df["total_price"] = df["quantity"] * df["price"]
                
    #Order wise detail
    order_wise = (
        df.groupby("oder_id")
        .agg(
            total_price=("total_price", "sum"),
            payment_state=("payment_state", "first"),
            depature_stale=("departure_state", "first"),
        )
        .reset_index()
    )
                
                
    #Pending payments
    pending_payments = df[df["payment_state"] == "Unpaid"][
        ["oder_id", "product_name", "quantity", "total_price", "payment_state"]
    ]
                
                
    # 3. Pending departure orders
    pending_departure = df[df["departure_state"] == "Not Dispatch"][
        [
            "oder_id",
            "product_name",
            "quantity",
            "total_price",
            "payment_state",
            "departure_state",
        ]
    ]
    logging.info("Data Calculation successfully")

except Exception as e:
    logging.error(f"Data Calculation failed: {e}")
    raise
            
            



try:
    report_path = f"Reports/daily_report_{date.today()}.xlsx"
    
    with pd.ExcelWriter(report_path, engine="openpyxl") as writer:
        start_row = 0

        order_wise.to_excel(
            writer, sheet_name="Report", index=False, startrow=start_row
        )
        start_row += len(order_wise) + 3

        pending_payments.to_excel(
            writer, sheet_name="Report", index=False, startrow=start_row
        )
        start_row += len(pending_payments) + 3

        pending_departure.to_excel(
            writer, sheet_name="Report", index=False, startrow=start_row
        )
                
                
                
except Exception as e:
    logging.error(f"Report Generation failed: {e}")
    raise
    
    
    
