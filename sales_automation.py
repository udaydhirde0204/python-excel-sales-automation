import os
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.chart import BarChart, Reference

INPUT_FILE = "sales.xlsx"
OUTPUT_FILE = "sales_report.xlsx"


def create_sample_data():
    sample_data = {
        "Date": [
            "2026-01-02", "2026-01-05", "2026-01-10", "2026-01-15",
            "2026-02-02", "2026-02-08", "2026-02-12", "2026-02-20",
            "2026-03-03", "2026-03-09", "2026-03-15", "2026-03-21",
            "2026-03-25", ""
        ],
        "Product": [
            "Laptop", "Mouse", "Keyboard", "Laptop",
            "Mouse", "Laptop", "Keyboard", "Monitor",
            "Laptop", "Monitor", "Mouse", "Keyboard",
            "  mouse ", "Broken Item"
        ],
        "Category": [
            "Electronics", "Accessories", "Accessories", "Electronics",
            "Accessories", "Electronics", "Accessories", "Electronics",
            "Electronics", "Electronics", "Accessories", "Accessories",
            "accessories", "Accessories"
        ],
        "Quantity": [2, 10, 5, 1, 8, 3, 6, 2, 2, 3, 12, 4, 2, -1],
        "Unit_Price": [50000, 500, 1200, 50000, 500, 50000, 1200, 15000, 50000, 15000, 500, 1200, 500, 500]
    }
    pd.DataFrame(sample_data).to_excel(INPUT_FILE, index=False)
    print(f"Sample '{INPUT_FILE}' created.")


def generate_report():
    if not os.path.exists(INPUT_FILE):
        create_sample_data()

    df = pd.read_excel(INPUT_FILE)

    required_cols = {"Date", "Product", "Category", "Quantity", "Unit_Price"}
    if not required_cols.issubset(df.columns):
        raise ValueError(f"Missing columns: {required_cols - set(df.columns)}")

    # Clean data
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    df = df.dropna(subset=["Date"])
    df["Quantity"] = pd.to_numeric(df["Quantity"], errors="coerce")
    df["Unit_Price"] = pd.to_numeric(df["Unit_Price"], errors="coerce")
    df = df[(df["Quantity"] > 0) & (df["Unit_Price"] >= 0)].copy()

    df["Product"] = df["Product"].astype(str).str.strip().str.title()
    df["Category"] = df["Category"].astype(str).str.strip().str.title()
    df = df[(df["Product"] != "") & (df["Category"] != "")]

    # Calculate revenue & clean date format
    df["Revenue"] = df["Quantity"] * df["Unit_Price"]
    df["Month"] = df["Date"].dt.to_period("M").astype(str)
    df["Date"] = df["Date"].dt.strftime("%Y-%m-%d")

    monthly_summary = (
        df.groupby("Month", as_index=False)
        .agg(Total_Orders=("Product", "count"), Total_Revenue=("Revenue", "sum"))
    )

    product_summary = (
        df.groupby("Product", as_index=False)
        .agg(Quantity_Sold=("Quantity", "sum"), Revenue=("Revenue", "sum"))
        .sort_values("Revenue", ascending=False)
    )

    total_revenue = df["Revenue"].sum()
    total_orders = len(df)
    total_units = int(df["Quantity"].sum())
    top_product = product_summary.iloc[0]["Product"]

    dashboard = pd.DataFrame({
        "Business Metric": [
            "Total Revenue (INR)",
            "Total Valid Transactions",
            "Total Units Sold",
            "Best-Selling Product"
        ],
        "Value": [total_revenue, total_orders, total_units, top_product]
    })

    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        dashboard.to_excel(writer, sheet_name="Dashboard", index=False)
        monthly_summary.to_excel(writer, sheet_name="Monthly Summary", index=False)
        product_summary.to_excel(writer, sheet_name="Product Summary", index=False)
        df.to_excel(writer, sheet_name="Cleaned Data", index=False)

    format_excel_workbook(OUTPUT_FILE)

    print("\nREPORT GENERATED SUCCESSFULLY!")
    print(f"Total Revenue: Rs. {total_revenue:,.2f}")
    print(f"Total Transactions: {total_orders}")
    print(f"Total Units Sold: {total_units}")
    print(f"Top Product: {top_product}")
    print(f"Output File: {OUTPUT_FILE}")


def format_excel_workbook(file_path):
    wb = load_workbook(file_path)
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True, size=11)

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")

        # Extra wide padding so Rs. numbers never show #########
        for col in ws.columns:
            max_len = max(len(str(cell.value or "")) for cell in col)
            col_letter = col[0].column_letter
            ws.column_dimensions[col_letter].width = max(max_len + 8, 18)

    # Format currency cells
    ws_dash = wb["Dashboard"]
    ws_dash["B2"].number_format = "₹#,##0.00"

    for ws_name, col_letter in [("Monthly Summary", "C"), ("Product Summary", "C"), ("Cleaned Data", "E"), ("Cleaned Data", "F")]:
        ws = wb[ws_name]
        for row in range(2, ws.max_row + 1):
            ws[f"{col_letter}{row}"].number_format = "₹#,##0.00"

    # Chart 1: Monthly Revenue
    chart1 = BarChart()
    chart1.type = "col"
    chart1.style = 10
    chart1.title = "Monthly Revenue (INR)"
    chart1.y_axis.title = "Revenue (₹)"
    chart1.x_axis.title = "Month"
    ws_monthly = wb["Monthly Summary"]
    chart1.add_data(Reference(ws_monthly, min_col=3, min_row=1, max_row=ws_monthly.max_row), titles_from_data=True)
    chart1.set_categories(Reference(ws_monthly, min_col=1, min_row=2, max_row=ws_monthly.max_row))
    chart1.legend = None
    chart1.width = 16
    chart1.height = 9
    ws_dash.add_chart(chart1, "D2")

    # Chart 2: Product Revenue
    chart2 = BarChart()
    chart2.type = "col"
    chart2.style = 11
    chart2.title = "Revenue by Product (INR)"
    chart2.y_axis.title = "Revenue (₹)"
    chart2.x_axis.title = "Product"
    ws_prod = wb["Product Summary"]
    chart2.add_data(Reference(ws_prod, min_col=3, min_row=1, max_row=ws_prod.max_row), titles_from_data=True)
    chart2.set_categories(Reference(ws_prod, min_col=1, min_row=2, max_row=ws_prod.max_row))
    chart2.legend = None
    chart2.width = 16
    chart2.height = 9
    ws_dash.add_chart(chart2, "D20")

    wb.save(file_path)


if __name__ == "__main__":
    generate_report()