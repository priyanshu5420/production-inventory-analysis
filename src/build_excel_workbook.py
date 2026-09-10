from pathlib import Path
import csv
from artifact_tool import Workbook, SpreadsheetFile

BASE = Path(__file__).resolve().parents[1]
DATA = BASE / "data"
OUT = BASE / "excel" / "production_inventory_analysis.xlsx"

def read_csv(filename):
    with open(DATA / filename, newline="", encoding="utf-8") as f:
        rows = list(csv.reader(f))
    return rows

def main():
    wb = Workbook.create()

    dashboard = wb.worksheets.add("Dashboard")
    products = wb.worksheets.add("Products")
    sales = wb.worksheets.add("Sales")
    production = wb.worksheets.add("Production")
    inventory = wb.worksheets.add("Inventory")
    capacity = wb.worksheets.add("Capacity")
    kpi = wb.worksheets.add("KPI_Calculations")

    data_map = {
        products: "products.csv",
        sales: "sales.csv",
        production: "production.csv",
        inventory: "inventory.csv",
        capacity: "capacity.csv",
    }

    for sheet, filename in data_map.items():
        values = read_csv(filename)
        sheet.get_range_by_indexes(0, 0, len(values), len(values[0])).values = values
        sheet.get_range("A1:Z1").format = {
            "fill": "#1F4E78",
            "font": {"bold": True, "color": "#FFFFFF"},
            "horizontal_alignment": "center",
            "vertical_alignment": "center",
        }
        sheet.freeze_panes.freeze_rows(1)
        sheet.get_range("A:Z").format.wrap_text = False

    # KPI sheet
    kpi.get_range("A1:E1").merge()
    kpi.get_range("A1").values = [["Operational & Financial KPI Summary"]]
    kpi.get_range("A1:E1").format = {
        "fill": "#1F4E78",
        "font": {"bold": True, "color": "#FFFFFF", "size": 16},
        "horizontal_alignment": "center",
    }

    kpi.get_range("A3:E3").values = [[
        "KPI", "Formula", "Value", "Unit", "Business Meaning"
    ]]
    kpi.get_range("A3:E3").format = {
        "fill": "#5B9BD5",
        "font": {"bold": True, "color": "#FFFFFF"},
        "horizontal_alignment": "center",
    }

    rows = [
        ["Total Revenue", "=SUM(Sales!D2:D73)", None, "$", "Total sales revenue"],
        ["Production Cost", "=SUM(Production!F2:F73)", None, "$", "Direct manufacturing cost"],
        ["Gross Profit", "=C4-C5", None, "$", "Revenue less direct production cost"],
        ["Gross Margin %", "=C6/C4", None, "%", "Profit as a share of revenue"],
        ["Planned Production", "=SUM(Production!C2:C73)", None, "units", "Scheduled output"],
        ["Actual Production", "=SUM(Production!D2:D73)", None, "units", "Completed output"],
        ["Production Achievement %", "=C9/C8", None, "%", "Actual output vs plan"],
        ["Rejected Units", "=SUM(Production!E2:E73)", None, "units", "Units rejected by QC"],
        ["Rejection Rate %", "=C11/C9", None, "%", "Rejected share of output"],
        ["Available Capacity", "=SUM(Capacity!C2:C37)", None, "units", "Total available line capacity"],
        ["Capacity Utilization %", "=C9/C13", None, "%", "Actual output vs available capacity"],
        ["Average Inventory", "=AVERAGE(Inventory!F2:F73)", None, "units", "Average monthly closing stock"],
        ["Units Sold", "=SUM(Sales!C2:C73)", None, "units", "Total demand fulfilled"],
        ["Inventory Turnover", "=C16/C15", None, "x", "Approximate annual stock cycling"],
        ["Stock-out Instances", "=COUNTIF(Inventory!H2:H73,1)", None, "instances", "Monthly SKU stock-out records"],
        ["Stock-out Rate %", "=C18/COUNT(Inventory!H2:H73)", None, "%", "Stock-out instances as share of records"],
    ]
    for i, row in enumerate(rows, start=4):
        kpi.get_range(f"A{i}:E{i}").values = [[row[0], row[1], None, row[3], row[4]]]
        kpi.get_range(f"C{i}").formulas = [[row[1]]]
    kpi.get_range("A4:E19").format.wrap_text = True

    # Dashboard
    dashboard.get_range("A1:J1").merge()
    dashboard.get_range("A1").values = [["APEX TECH GEAR — PRODUCTION & INVENTORY DASHBOARD"]]
    dashboard.get_range("A1:J1").format = {
        "fill": "#1F4E78",
        "font": {"bold": True, "color": "#FFFFFF", "size": 16},
        "horizontal_alignment": "center",
        "vertical_alignment": "center",
    }

    cards = [
        ("A3:B3","A4:B4","TOTAL REVENUE","=KPI_Calculations!C4","$#,##0"),
        ("D3:E3","D4:E4","GROSS PROFIT","=KPI_Calculations!C6","$#,##0"),
        ("G3:H3","G4:H4","GROSS MARGIN","=KPI_Calculations!C7","0.0%"),
        ("A6:B6","A7:B7","CAPACITY UTILIZATION","=KPI_Calculations!C14","0.0%"),
        ("D6:E6","D7:E7","INVENTORY TURNOVER","=KPI_Calculations!C17","0.00x"),
        ("G6:H6","G7:H7","STOCK-OUT RATE","=KPI_Calculations!C19","0.0%"),
    ]
    for tr, vr, title, formula, fmt in cards:
        dashboard.get_range(tr).merge()
        dashboard.get_range(vr).merge()
        dashboard.get_range(tr.split(":")[0]).values = [[title]]
        dashboard.get_range(vr.split(":")[0]).formulas = [[formula]]
        dashboard.get_range(tr).format = {
            "fill": "#D9EAF7",
            "font": {"bold": True},
            "horizontal_alignment": "center",
        }
        dashboard.get_range(vr).format = {
            "fill": "#F3F6F9",
            "font": {"bold": True, "size": 16},
            "horizontal_alignment": "center",
        }
        dashboard.get_range(vr).format.number_format = fmt

    # Monthly helper table for charts
    months = ["2025-"+f"{m:02d}" for m in range(1,13)]
    dashboard.get_range("A10:C10").values = [["Month","Revenue","Actual Production"]]
    dashboard.get_range("A10:C10").format = {
        "fill": "#5B9BD5",
        "font": {"bold": True, "color": "#FFFFFF"},
    }
    for r, month in enumerate(months, start=11):
        dashboard.get_range(f"A{r}").values = [[month]]
        dashboard.get_range(f"B{r}").formulas = [[f'=SUMIF(Sales!A$2:A$73,A{r},Sales!D$2:D$73)']]
        dashboard.get_range(f"C{r}").formulas = [[f'=SUMIF(Production!A$2:A$73,A{r},Production!D$2:D$73)']]
    dashboard.get_range("B11:B22").format.number_format = "$#,##0"
    dashboard.get_range("C11:C22").format.number_format = "#,##0"

    try:
        chart = dashboard.charts.add("line", dashboard.get_range("A10:C22"))
        chart.title_text = "Monthly Revenue & Production"
        chart.has_legend = True
        chart.legend.position = "bottom"
        chart.set_position("J3", "P18")
    except Exception:
        pass

    # Formatting widths
    for sheet in [dashboard, products, sales, production, inventory, capacity, kpi]:
        sheet.get_range("A1:Z100").format.wrap_text = False

    SpreadsheetFile.export_xlsx(wb).save(OUT)

if __name__ == "__main__":
    main()
