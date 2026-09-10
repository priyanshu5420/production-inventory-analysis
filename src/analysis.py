import csv
from pathlib import Path
from collections import defaultdict

BASE = Path(__file__).resolve().parents[1]
DATA = BASE / "data"

def read_csv(filename):
    with open(DATA / filename, newline="", encoding="utf-8") as f:
        return list(csv.DictReader(f))

def main():
    products = read_csv("products.csv")
    sales = read_csv("sales.csv")
    production = read_csv("production.csv")
    inventory = read_csv("inventory.csv")
    capacity = read_csv("capacity.csv")

    product_names = {r["Product_ID"]: r["Product_Name"] for r in products}

    total_revenue = sum(float(r["Revenue"]) for r in sales)
    total_cost = sum(float(r["Production_Cost"]) for r in production)
    gross_profit = total_revenue - total_cost
    gross_margin = gross_profit / total_revenue if total_revenue else 0

    planned = sum(int(r["Planned_Production"]) for r in production)
    actual = sum(int(r["Actual_Production"]) for r in production)
    rejected = sum(int(r["Rejected_Units"]) for r in production)
    capacity_available = sum(int(r["Available_Capacity"]) for r in capacity)
    avg_inventory = (
        sum(int(r["Closing_Inventory"]) for r in inventory) / len(inventory)
        if inventory else 0
    )
    units_sold = sum(int(r["Units_Sold"]) for r in sales)
    demand = sum(int(r["Demand_Units"]) for r in sales)
    stockouts = sum(int(r["Stockout_Flag"]) for r in inventory)

    print("=== PRODUCTION & INVENTORY ANALYSIS ===")
    print(f"Products: {len(products)}")
    print(f"Total demand: {demand:,}")
    print(f"Units sold: {units_sold:,}")
    print(f"Total revenue: {total_revenue:,.2f}")
    print(f"Production cost: {total_cost:,.2f}")
    print(f"Gross profit: {gross_profit:,.2f}")
    print(f"Gross margin: {gross_margin:.2%}")
    print(f"Production achievement: {actual / planned:.2%}")
    print(f"Capacity utilization: {actual / capacity_available:.2%}")
    print(f"Rejection rate: {rejected / actual:.2%}")
    print(f"Inventory turnover: {units_sold / avg_inventory:.2f}x")
    print(f"Stock-out rate: {stockouts / len(inventory):.2%}")

    print("\nProduct performance")
    revenue_by_product = defaultdict(float)
    sold_by_product = defaultdict(int)
    demand_by_product = defaultdict(int)
    for r in sales:
        revenue_by_product[r["Product_ID"]] += float(r["Revenue"])
        sold_by_product[r["Product_ID"]] += int(r["Units_Sold"])
        demand_by_product[r["Product_ID"]] += int(r["Demand_Units"])

    for pid in product_names:
        rev = revenue_by_product[pid]
        fill_rate = sold_by_product[pid] / demand_by_product[pid] if demand_by_product[pid] else 0
        print(f"{pid} - {product_names[pid]}: revenue={rev:,.0f}, fill_rate={fill_rate:.2%}")

if __name__ == "__main__":
    main()
