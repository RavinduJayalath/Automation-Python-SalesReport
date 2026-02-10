import logging
import pandas as pd
from datetime import date
import os
from io import BytesIO
import matplotlib.pyplot as plt
import smtplib
from email.message import EmailMessage
from pathlib import Path


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





#Data Calculation for excel
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









#Data Calculation for email
try:


    pending_payment_total = df[df["payment_state"] == "Unpaid"]["total_price"].sum()
    paid_payment_total = df[df["payment_state"] == "Paid"]["total_price"].sum()

    pending_dispatch_total = df[df["departure_state"] == "Not Dispatch"]["total_price"].sum()
    dispatched_total = df[df["departure_state"] == "Dispatch"]["total_price"].sum()

    #create pie chart
    payment_buf = BytesIO()
    plt.figure()
    plt.pie(
        [paid_payment_total, pending_payment_total],
        labels=["Paid", "Pending Payment"],
        autopct="%1.1f%%"
    )
    plt.title("Payment Status")
    plt.savefig(payment_buf, format="png")
    plt.close()
    payment_buf.seek(0)


    dispatch_buf = BytesIO()
    plt.figure()
    plt.pie(
        [dispatched_total, pending_dispatch_total],
        labels=["Dispatched", "Pending Dispatch"],
        autopct="%1.1f%%"
    )
    plt.title("Dispatch Status")
    plt.savefig(dispatch_buf, format="png")
    plt.close()
    dispatch_buf.seek(0)



    logging.info("Data calculation for email successful")
except Exception as e:
    logging.error(f"Data calculation for email failed: {e}")
    raise







#Generate Report
try:
    report_path = f"Reports/daily_report_{date.today()}.xlsx"
    os.makedirs("Reports", exist_ok=True)

    if os.path.exists(report_path):
        os.remove(report_path)

    with pd.ExcelWriter(report_path, engine="xlsxwriter") as writer:
        workbook  = writer.book #entire Excel file
        #worksheet = writer.sheets["Report"] #single sheet inside the workbook

        #Formats
        title_format = workbook.add_format({    #Formats are defined at workbook level then can apply it to any sheet or any cell
            "bold": True,
            "align": "center",
            "valign": "vcenter",
            "font_size": 12
        })


        #TABLE 1
        order_wise.to_excel(writer, sheet_name="Report", startrow=1, index=False)
        worksheet = writer.sheets["Report"]
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



#HTML Creation
try:
    html = f"""
    <html>
        <head>
            <style>
                body {{
                    font-family: Arial, sans-serif;
                    background:#f4f6f8;
                }}
                .card {{
                    background:#ffffff;
                    padding:20px;
                    border-radius:10px;
                    margin-bottom:20px;
                }}
                h1 {{
                    color:#1f2937;
                }}
                .stat {{
                    font-size:18px;
                    margin:10px 0;
                }}
                .footer {{
                    color:#6b7280;
                    font-size:12px;
                }}
            </style>
        </head>
        <body>
            <div class="card">
                <h1>Sales Status Report</h1>
                <p class="stat">Pending Payment Total: <b>LKR {pending_payment_total:,.2f}</b></p>
                <p class="stat">Pending Dispatch Total: <b>LKR {pending_dispatch_total:,.2f}</b></p>
            </div>

            <div class="card">
                <h2>Payment Overview</h2>
                <img src="cid:payment">
            </div>

            <div class="card">
                <h2>Dispatch Overview</h2>
                <img src="cid:dispatch">
            </div>

            <p class="footer">
                This is an automated enterprise sales report.<br>
                Generated by Python Automation System.
            </p>
        </body>
    </html>
    """



    logging.info("HTML Creation successfully")
except Exception as e:
    logging.error(f"HTML Creation failed: {e}")
    raise






#Send email
try:
    #email cinfig
    SENDER_EMAIL = "ravindu.01999@gmail.com"
    APP_PASSWORD = "app password of your email acc"
    RECEIVER_EMAIL = "ravindu0032@gmail.com"

    msg = EmailMessage()
    msg["Subject"] = "Enterprise Sales Payment & Dispatch Report"
    msg["From"] = SENDER_EMAIL
    msg["To"] = RECEIVER_EMAIL
    msg.set_content("HTML email required")
    msg.add_alternative(html, subtype="html")

    #Attach directly to email
    msg.get_payload()[1].add_related(
        payment_buf.read(),
        maintype="image",
        subtype="png",
        cid="payment"
    )

    msg.get_payload()[1].add_related(
        dispatch_buf.read(),
        maintype="image",
        subtype="png",
        cid="dispatch"
    )


    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(SENDER_EMAIL, APP_PASSWORD)
        server.send_message(msg)    

    logging.info("Send email successfully\n")
except Exception as e:
    logging.error(f"Send email failed: {e}\n")
    raise