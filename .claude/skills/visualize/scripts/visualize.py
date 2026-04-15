"""
Visualization script for MegaMart sales & returns data.

Reads CSV files from '.claude/skills/fetchAPI/data/<datetime>/',
computes KPIs, and saves PNG charts to
'.claude/skills/visualize/visualizations/<datetime>/'.
"""

import os
import pathlib

import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import pandas as pd
import seaborn as sns

# ── Configuration ─────────────────────────────────────────────────────────────
BASE_DATA_DIR = pathlib.Path(".claude/skills/fetchAPI/data")
VIZ_DIR = pathlib.Path(".claude/skills/visualize/visualizations")

sns.set_theme(style="whitegrid", palette="muted")
plt.rcParams.update({"figure.dpi": 120, "figure.facecolor": "white"})


# ── Helpers ───────────────────────────────────────────────────────────────────
def _latest_folder(base: pathlib.Path) -> pathlib.Path:
    """Return the most-recently-named datetime folder."""
    folders = sorted(
        [f for f in base.iterdir() if f.is_dir()],
        key=lambda p: p.name,
    )
    if not folders:
        raise FileNotFoundError(f"No data folders found under {base}")
    return folders[-1]


def load_data(folder: pathlib.Path) -> dict[str, pd.DataFrame]:
    """Load all CSV files in *folder* and return a name-keyed dict."""
    return {
        csv.stem: pd.read_csv(csv)
        for csv in sorted(folder.glob("*.csv"))
    }


# ── KPI Calculation ───────────────────────────────────────────────────────────
def compute_kpis(
    tables: dict[str, pd.DataFrame],
) -> dict[str, float | pd.DataFrame]:
    """Compute business KPIs from the loaded tables."""
    sales = tables["fact_sales"]
    returns = tables["fact_returns"]
    stores = tables["dim_store"]
    products = tables["dim_product"]
    customers = tables["dim_customer"]

    total_sales: float = sales["net_amount"].sum()
    total_returns: float = returns["refund_amount"].sum()
    net_sales: float = total_sales - total_returns

    # Average Sales per Store
    sales_by_store = (
        sales.groupby("store_sk")["net_amount"]
        .sum()
        .reset_index(name="store_net_sales")
        .merge(stores[["store_sk", "store_name"]], on="store_sk", how="left")
    )
    avg_sales_per_store: float = sales_by_store["store_net_sales"].mean()

    # Average Returns per Product
    returns_by_product = (
        returns.groupby("product_sk")["refund_amount"]
        .sum()
        .reset_index(name="product_refund")
        .merge(products[["product_sk", "product_name"]], on="product_sk", how="left")
    )
    avg_returns_per_product: float = returns_by_product["product_refund"].mean()

    # Average Sales per Customer
    sales_by_customer = (
        sales.groupby("customer_sk")["net_amount"]
        .sum()
        .reset_index(name="customer_net_sales")
        .merge(
            customers[["customer_sk", "first_name", "last_name"]],
            on="customer_sk",
            how="left",
        )
    )
    avg_sales_per_customer: float = sales_by_customer["customer_net_sales"].mean()

    return {
        "total_sales": total_sales,
        "total_returns": total_returns,
        "net_sales": net_sales,
        "avg_sales_per_store": avg_sales_per_store,
        "avg_returns_per_product": avg_returns_per_product,
        "avg_sales_per_customer": avg_sales_per_customer,
        # enriched frames for charts
        "sales_by_store": sales_by_store,
        "returns_by_product": returns_by_product,
        "sales_by_customer": sales_by_customer,
        "sales": sales,
        "returns": returns,
        "stores": stores,
        "products": products,
        "customers": customers,
    }


# ── Plotting ──────────────────────────────────────────────────────────────────
def _save(fig: plt.Figure, out_dir: pathlib.Path, name: str) -> None:
    path = out_dir / f"{name}.png"
    fig.savefig(path, bbox_inches="tight")
    plt.close(fig)
    print(f"  Saved → {path}")


def plot_kpi_summary(kpis: dict, out_dir: pathlib.Path) -> None:
    """Bar chart of the six scalar KPIs."""
    labels = [
        "Total Sales",
        "Total Returns",
        "Net Sales",
        "Avg Sales/Store",
        "Avg Returns/Product",
        "Avg Sales/Customer",
    ]
    values = [
        kpis["total_sales"],
        kpis["total_returns"],
        kpis["net_sales"],
        kpis["avg_sales_per_store"],
        kpis["avg_returns_per_product"],
        kpis["avg_sales_per_customer"],
    ]

    fig, ax = plt.subplots(figsize=(10, 5))
    bars = ax.bar(labels, values, color=sns.color_palette("muted", len(labels)))
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f"${x:,.0f}"))
    ax.set_title("KPI Summary", fontsize=14, fontweight="bold")
    ax.set_xlabel("KPI")
    ax.set_ylabel("USD ($)")
    for bar in bars:
        ax.text(
            bar.get_x() + bar.get_width() / 2,
            bar.get_height() * 1.01,
            f"${bar.get_height():,.0f}",
            ha="center",
            va="bottom",
            fontsize=8,
        )
    plt.xticks(rotation=20, ha="right")
    _save(fig, out_dir, "01_kpi_summary")


def plot_sales_by_store(kpis: dict, out_dir: pathlib.Path) -> None:
    """Horizontal bar chart – net sales per store."""
    df = kpis["sales_by_store"].sort_values("store_net_sales", ascending=True)

    fig, ax = plt.subplots(figsize=(9, 4))
    ax.barh(df["store_name"], df["store_net_sales"], color=sns.color_palette("Blues_d", len(df)))
    ax.xaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f"${x:,.0f}"))
    ax.set_title("Net Sales by Store", fontsize=14, fontweight="bold")
    ax.set_xlabel("Net Sales (USD)")
    _save(fig, out_dir, "02_sales_by_store")


def plot_returns_by_product(kpis: dict, out_dir: pathlib.Path) -> None:
    """Bar chart – total refund amount per product."""
    df = kpis["returns_by_product"].sort_values("product_refund", ascending=False)

    fig, ax = plt.subplots(figsize=(10, 5))
    ax.bar(df["product_name"], df["product_refund"], color=sns.color_palette("Oranges_d", len(df)))
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f"${x:,.0f}"))
    ax.set_title("Total Returns (Refund Amount) by Product", fontsize=14, fontweight="bold")
    ax.set_xlabel("Product")
    ax.set_ylabel("Refund Amount (USD)")
    plt.xticks(rotation=30, ha="right")
    _save(fig, out_dir, "03_returns_by_product")


def plot_sales_vs_returns_pie(kpis: dict, out_dir: pathlib.Path) -> None:
    """Pie chart – total sales vs total returns share."""
    sizes = [kpis["total_sales"], kpis["total_returns"]]
    labels = [f"Gross Sales\n${kpis['total_sales']:,.0f}", f"Returns\n${kpis['total_returns']:,.0f}"]

    fig, ax = plt.subplots(figsize=(6, 6))
    ax.pie(sizes, labels=labels, autopct="%1.1f%%", startangle=90,
           colors=["#4CAF50", "#F44336"])
    ax.set_title("Gross Sales vs Returns", fontsize=14, fontweight="bold")
    _save(fig, out_dir, "04_sales_vs_returns_pie")


def plot_payment_method_distribution(kpis: dict, out_dir: pathlib.Path) -> None:
    """Pie chart – sales split by payment method."""
    sales = kpis["sales"]
    pm = sales.groupby("payment_method")["net_amount"].sum().reset_index()

    fig, ax = plt.subplots(figsize=(6, 6))
    ax.pie(pm["net_amount"], labels=pm["payment_method"], autopct="%1.1f%%", startangle=140)
    ax.set_title("Sales by Payment Method", fontsize=14, fontweight="bold")
    _save(fig, out_dir, "05_payment_method_pie")


def plot_daily_sales_trend(kpis: dict, out_dir: pathlib.Path) -> None:
    """Line chart – daily net sales trend (using date_sk as a proxy)."""
    sales = kpis["sales"]
    daily = sales.groupby("date_sk")["net_amount"].sum().reset_index()

    fig, ax = plt.subplots(figsize=(12, 5))
    ax.plot(daily["date_sk"], daily["net_amount"], color="#1976D2", linewidth=1.5)
    ax.fill_between(daily["date_sk"], daily["net_amount"], alpha=0.15, color="#1976D2")
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f"${x:,.0f}"))
    ax.set_title("Daily Net Sales Trend", fontsize=14, fontweight="bold")
    ax.set_xlabel("Date Key (chronological)")
    ax.set_ylabel("Net Sales (USD)")
    _save(fig, out_dir, "06_daily_sales_trend")


def plot_sales_heatmap_store_product(kpis: dict, out_dir: pathlib.Path) -> None:
    """Heatmap – net sales by store × product."""
    sales = kpis["sales"]
    stores = kpis["stores"]
    products = kpis["products"]

    df = (
        sales.merge(stores[["store_sk", "store_name"]], on="store_sk", how="left")
        .merge(products[["product_sk", "product_name"]], on="product_sk", how="left")
        .groupby(["store_name", "product_name"])["net_amount"]
        .sum()
        .unstack(fill_value=0)
    )

    fig, ax = plt.subplots(figsize=(14, 5))
    sns.heatmap(
        df, annot=True, fmt=".0f", cmap="YlGnBu",
        linewidths=0.5, ax=ax,
        cbar_kws={"label": "Net Sales (USD)"},
    )
    ax.set_title("Net Sales Heatmap – Store × Product", fontsize=14, fontweight="bold")
    ax.set_xlabel("Product")
    ax.set_ylabel("Store")
    plt.xticks(rotation=30, ha="right")
    _save(fig, out_dir, "07_heatmap_store_product")


def plot_customer_sales_boxplot(kpis: dict, out_dir: pathlib.Path) -> None:
    """Box plot – distribution of per-customer net sales by loyalty tier."""
    sales = kpis["sales"]
    customers = kpis["customers"]

    df = (
        sales.groupby("customer_sk")["net_amount"]
        .sum()
        .reset_index(name="customer_total")
        .merge(customers[["customer_sk", "loyalty_tier"]], on="customer_sk", how="left")
    )
    tier_order = ["Platinum", "Gold", "Silver", "None"]

    fig, ax = plt.subplots(figsize=(8, 5))
    sns.boxplot(data=df, x="loyalty_tier", y="customer_total", order=tier_order,
                palette="Set2", ax=ax)
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f"${x:,.0f}"))
    ax.set_title("Customer Total Spend Distribution by Loyalty Tier", fontsize=14, fontweight="bold")
    ax.set_xlabel("Loyalty Tier")
    ax.set_ylabel("Total Net Spend (USD)")
    _save(fig, out_dir, "08_customer_spend_boxplot")


def plot_return_reason_distribution(kpis: dict, out_dir: pathlib.Path) -> None:
    """Bar chart – count of returns by reason."""
    returns = kpis["returns"]
    reason_counts = returns["return_reason"].value_counts().reset_index()
    reason_counts.columns = ["reason", "count"]

    fig, ax = plt.subplots(figsize=(8, 4))
    ax.bar(reason_counts["reason"], reason_counts["count"],
           color=sns.color_palette("Reds_d", len(reason_counts)))
    ax.set_title("Return Reason Distribution", fontsize=14, fontweight="bold")
    ax.set_xlabel("Return Reason")
    ax.set_ylabel("Number of Returns")
    plt.xticks(rotation=20, ha="right")
    _save(fig, out_dir, "09_return_reason_distribution")


# ── Entrypoint ────────────────────────────────────────────────────────────────
def main() -> None:
    """Main execution: load data, compute KPIs, generate and save charts."""
    data_folder = _latest_folder(BASE_DATA_DIR)
    folder_name = data_folder.name
    print(f"Using data folder: {data_folder}")

    tables = load_data(data_folder)
    print(f"Loaded tables: {list(tables.keys())}")

    kpis = compute_kpis(tables)

    # Print scalar KPIs
    print("\n── KPIs ──────────────────────────────────────────")
    print(f"  Total Sales          : ${kpis['total_sales']:>12,.2f}")
    print(f"  Total Returns        : ${kpis['total_returns']:>12,.2f}")
    print(f"  Net Sales            : ${kpis['net_sales']:>12,.2f}")
    print(f"  Avg Sales / Store    : ${kpis['avg_sales_per_store']:>12,.2f}")
    print(f"  Avg Returns / Product: ${kpis['avg_returns_per_product']:>12,.2f}")
    print(f"  Avg Sales / Customer : ${kpis['avg_sales_per_customer']:>12,.2f}")
    print("──────────────────────────────────────────────────\n")

    out_dir = VIZ_DIR / folder_name
    out_dir.mkdir(parents=True, exist_ok=True)
    print(f"Saving visualizations to: {out_dir}\n")

    plot_kpi_summary(kpis, out_dir)
    plot_sales_by_store(kpis, out_dir)
    plot_returns_by_product(kpis, out_dir)
    plot_sales_vs_returns_pie(kpis, out_dir)
    plot_payment_method_distribution(kpis, out_dir)
    plot_daily_sales_trend(kpis, out_dir)
    plot_sales_heatmap_store_product(kpis, out_dir)
    plot_customer_sales_boxplot(kpis, out_dir)
    plot_return_reason_distribution(kpis, out_dir)

    print("\nAll visualizations saved successfully.")


if __name__ == "__main__":
    main()
