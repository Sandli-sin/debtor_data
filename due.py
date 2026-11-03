import pandas as pd
from openpyxl import load_workbook

###fixing Dr Cr problem
wb = load_workbook("dataset_vin_31.xlsx", data_only=True)

ws = wb.active

for row in range(1, 30):
    cell = ws.cell(row=row, column=5)

    if cell.value is not None:
        print(
            "Cell:", cell.coordinate,
            "| Value:", cell.value,
            "| Format:", cell.number_format
        )

# load excel file
df = pd.read_excel("dataset_vin_31.xlsx")

print(df.columns.to_list())

opening_types = []
for row in range(2, ws.max_row + 1):

    cell = ws.cell(row=row, column=5)

    if cell.value is None:
        opening_types.append(None)

    elif "Dr" in cell.number_format:
        opening_types.append("Dr")

    elif "Cr" in cell.number_format:
        opening_types.append("Cr")

    else:
        opening_types.append(None)

df["Opening_Type"] = opening_types

df["Prior_Debt"] = 0.0
df["Advance_Payment"] = 0.0

df.loc[
    df["Opening_Type"] == "Dr",
    "Prior_Debt"
] = df.loc[
    df["Opening_Type"] == "Dr",
    "Opening"
]

df.loc[
    df["Opening_Type"] == "Cr",
    "Advance_Payment"
] = df.loc[
    df["Opening_Type"] == "Cr",
    "Opening"
]

#checking if that worked
#print(df[df["Opening"].notna()][["DEBTOR", "Opening", "Opening_Type","Prior_Debt", "Advance_Payment"]].to_string(index=False))
#print(df["Opening_Type"].value_counts(dropna=False))


### Remove extra spaces from column names
df.columns = df.columns.str.strip()

# Convert DATE to datetime
df["DATE"] = pd.to_datetime(df["DATE"], errors="coerce")

# Convert monetary columns to numeric
money_columns = ["Opening", "Debit", "Credit"]

for col in money_columns:
    df[col] = pd.to_numeric(df[col], errors="coerce")

# Transaction columns should have 0 instead of NaN
df["Debit"] = df["Debit"].fillna(0)
df["Credit"] = df["Credit"].fillna(0)

print("Dataset Info:")
print(df.info())



###fifo algo
pending_bills = []

for debtor, group in df.groupby("DEBTOR"):

    outstanding = []
    advance_payment = 0

    # -----------------------------------------
    # STEP 1: PROCESS OPENING BALANCE
    # -----------------------------------------

    for _, row in group.iterrows():

        if row["Opening_Type"] == "Dr":

            outstanding.append({
                "DEBTOR": debtor,
                "Bill_Date": pd.Timestamp("2026-04-01"),
                "VCH_NO": "Opening Balance",
                "Original_Amount": row["Opening"],
                "Pending_Amount": row["Opening"],
                "Source": "Opening Dr"
            })

        elif row["Opening_Type"] == "Cr":

            advance_payment += row["Opening"]


    # -----------------------------------------
    # STEP 2: PROCESS TRANSACTIONS BY DATE
    # -----------------------------------------

    for _, row in group.iterrows():

        # -----------------------------
        # NEW DEBIT / BILL
        # -----------------------------

        if row["Debit"] > 0:

            bill_amount = row["Debit"]

            # First use any advance payment
            if advance_payment > 0:

                if advance_payment >= bill_amount:

                    advance_payment -= bill_amount
                    bill_amount = 0

                else:

                    bill_amount -= advance_payment
                    advance_payment = 0


            # If something is still unpaid,
            # create a pending bill
            if bill_amount > 0:

                outstanding.append({
                    "DEBTOR": debtor,
                    "Bill_Date": row["DATE"],
                    "VCH_NO": row["VCH NO."],
                    "Original_Amount": row["Debit"],
                    "Pending_Amount": bill_amount,
                    "Source": "Debit"
                })


        # -----------------------------
        # CREDIT / PAYMENT
        # -----------------------------

        if row["Credit"] > 0:

            payment = row["Credit"]

            # Apply payment to oldest bills first
            while payment > 0 and outstanding:

                bill = outstanding[0]

                if payment >= bill["Pending_Amount"]:

                    payment -= bill["Pending_Amount"]

                    bill["Pending_Amount"] = 0

                    outstanding.pop(0)

                else:

                    bill["Pending_Amount"] -= payment

                    payment = 0


            # If payment remains after clearing
            # all existing bills, it becomes advance
            if payment > 0:

                advance_payment += payment


    # -----------------------------------------
    # STEP 3: SAVE REMAINING PENDING BILLS
    # -----------------------------------------

    for bill in outstanding:

        if bill["Pending_Amount"] > 0:

            pending_bills.append(bill)


pending_df = pd.DataFrame(pending_bills)
# Due date periods
pending_df["Due_Date_30"] = (
    pending_df["Bill_Date"]
    + pd.Timedelta(days=30)
)

pending_df["Due_Date_45"] = (
    pending_df["Bill_Date"]
    + pd.Timedelta(days=45)
)

pending_df["Due_Date_90"] = (
    pending_df["Bill_Date"]
    + pd.Timedelta(days=90)
)

today = pd.Timestamp.today().normalize()

pending_df["Status_30"] = pending_df.apply(
    lambda row:
        "Opening Balance - Date Unknown"
        if pd.isna(row["Bill_Date"])
        else (
            "Overdue"
            if row["Due_Date_30"] < today
            else "Pending"
        ),
    axis=1
)

pending_df["Status_45"] = pending_df.apply(
    lambda row:
        "Opening Balance - Date Unknown"
        if pd.isna(row["Bill_Date"])
        else (
            "Overdue"
            if row["Due_Date_45"] < today
            else "Pending"
        ),
    axis=1
)

pending_df["Status_90"] = pending_df.apply(
    lambda row:
        "Opening Balance - Date Unknown"
        if pd.isna(row["Bill_Date"])
        else (
            "Overdue"
            if row["Due_Date_90"] < today
            else "Pending"
        ),
    axis=1
)

date_columns = [
    "Bill_Date",
    "Due_Date_30",
    "Due_Date_45",
    "Due_Date_90",
]

for col in date_columns:
    pending_df[col] = pending_df[col].dt.strftime("%d-%m-%Y")

date_columns = [
    "Bill_Date",
    "Due_Date_30",
    "Due_Date_45",
    "Due_Date_90"
]

for col in date_columns:
    pending_df[col] = pd.to_datetime(
        pending_df[col],
        format="%d-%m-%Y",
        errors="coerce"
    )

#total outstanding
pending_df["Total_Outstanding"] = (
    pending_df.groupby("DEBTOR")["Pending_Amount"]
    .transform("sum")
)

pending_df = pending_df[
    [
        "Bill_Date",
        "DEBTOR",
        "VCH_NO",
        "Original_Amount",
        "Pending_Amount",
        "Total_Outstanding",
        "Source",
        "Due_Date_30",
        "Status_30",
        "Due_Date_45",
        "Status_45",
        "Due_Date_90",
        "Status_90"
    ]
]


#print(pending_df.to_string(index=False))

#pushing it in excel file
output_file = "Pending_Debtor_Report.xlsx"

pending_df.to_excel(
    output_file,
    index=False,
    sheet_name="Pending Bills"
)

wb2 = load_workbook(output_file)

ws2 = wb2["Pending Bills"]


headers = {
    cell.value: cell.column
    for cell in ws2[1]
}

due_30_col = headers["Due_Date_30"]
due_45_col = headers["Due_Date_45"]
due_90_col = headers["Due_Date_90"]

status_30_col = headers["Status_30"]
status_45_col = headers["Status_45"]
status_90_col = headers["Status_90"]

for row in range(2, ws2.max_row + 1):

    due_30_letter = ws2.cell(row=1, column=due_30_col).column_letter

    due_45_letter = ws2.cell(row=1, column=due_45_col).column_letter

    due_90_letter = ws2.cell(row=1, column=due_90_col).column_letter

    ws2.cell( row=row,column=status_30_col).value = f'=IF({due_30_letter}{row}<TODAY(),"Overdue","Pending")'

    ws2.cell(row=row,column=status_45_col).value = f'=IF({due_45_letter}{row}<TODAY(),"Overdue","Pending")'

    ws2.cell(row=row,column=status_90_col).value = f'=IF({due_90_letter}{row}<TODAY(),"Overdue","Pending")'

wb2.save(output_file)




