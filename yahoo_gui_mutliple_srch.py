# Based on: user search downloads multiple company data 
import os
import io
import time
import pandas as pd
import yfinance as yf
import requests
import tkinter as tk
from tkinter import messagebox
from datetime import datetime

# ========== CONFIG ==========
OUTPUT_DIR = "exports"
os.makedirs(OUTPUT_DIR, exist_ok=True)

# ========== SCRAPER ==========
def safe_df(df):
    return df if isinstance(df, pd.DataFrame) else pd.DataFrame()

def scrape_yahoo_tables(url, retries=3, delay=5):
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/111.0.0.0 Safari/537.36",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Language": "en-US,en;q=0.5",
        "DNT": "1",
        "Connection": "close"
    }
    for attempt in range(retries):
        try:
            response = requests.get(url, headers=headers, timeout=10)
            response.raise_for_status()
            html_io = io.StringIO(response.text)
            tables = pd.read_html(html_io)
            if tables:
                return pd.concat(tables, ignore_index=True)
        except requests.exceptions.HTTPError as e:
            if "429" in str(e):
                print(f"⚠️ HTTP 429 — sleeping for {delay} seconds...")
                time.sleep(delay)
                delay += 5
                continue
            elif "404" in str(e):
                print(f"⚠️ HTTP 404 — URL not found: {url}")
                break
            else:
                print(f"HTTP Error: {e}")
                break
        except Exception as e:
            print(f"Error: {e}")
            break
    return pd.DataFrame()

# ========== FETCH & EXPORT ==========
def fetch_and_export(ticker):
    print(f"\n🔍 Fetching data for {ticker}...")
    stock = yf.Ticker(ticker)

    history_df = safe_df(stock.history(period="1Y"))
    if not history_df.empty and hasattr(history_df.index, "tz_localize"):
        history_df.index = history_df.index.tz_localize(None)

    balance_sheet_df = safe_df(stock.balance_sheet)
    financials_df = safe_df(stock.financials)
    cashflow_df = safe_df(stock.cashflow)

    info_dict = stock.info if isinstance(stock.info, dict) else {}
    calendar_data = stock.calendar if isinstance(stock.calendar, dict) else {}
    info_df = pd.DataFrame(list(info_dict.items()), columns=["Attribute", "Value"]) if info_dict else pd.DataFrame(columns=["Attribute", "Value"])

    stats_url = f"https://finance.yahoo.com/quote/{ticker}/key-statistics?p={ticker}"
    analysis_url = f"https://finance.yahoo.com/quote/{ticker}/analysis/"
    yahoo_stats_df = scrape_yahoo_tables(stats_url)
    yahoo_analysis_df = scrape_yahoo_tables(analysis_url)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"{ticker.replace('.', '_')}_YahooFinanceData_{timestamp}.xlsx"
    filepath = os.path.join(OUTPUT_DIR, filename)

    with pd.ExcelWriter(filepath, engine="openpyxl") as writer:
        if not info_df.empty:
            info_df.to_excel(writer, sheet_name="Company_Info", index=False)
        if not history_df.empty:
            history_df = history_df.round(2)
            history_df.to_excel(writer, sheet_name="Price_History")
        if not balance_sheet_df.empty:
            balance_sheet_df = balance_sheet_df.apply(pd.to_numeric, errors='coerce') / 1000
            balance_sheet_df.to_excel(writer, sheet_name="Balance_Sheet")
            # Access the workbook and worksheet
        workbook = writer.book
        worksheet = writer.sheets["Balance_Sheet"]

        # Apply number format to all data cells (excluding header)
        for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row,
                                       min_col=1, max_col=worksheet.max_column):
            for cell in row:
                if isinstance(cell.value, (int, float)):
                    cell.number_format = '0.0'

        if not financials_df.empty:
            financials_df = financials_df.apply(pd.to_numeric, errors='coerce') / 1000
            financials_df.to_excel(writer, sheet_name="Income_Stmnt")
            # Access the workbook and worksheet
        workbook = writer.book
        worksheet = writer.sheets["Income_Stmnt"]

        # Apply number format to all data cells (excluding header)
        for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row,
                                       min_col=1, max_col=worksheet.max_column):
            for cell in row:
                if isinstance(cell.value, (int, float)):
                    cell.number_format = '0.0'

        if not cashflow_df.empty:
            cashflow_df = cashflow_df.apply(pd.to_numeric, errors='coerce') / 1000
            cashflow_df.to_excel(writer, sheet_name="Cashflow")
        # Access the workbook and worksheet
        workbook = writer.book
        worksheet = writer.sheets["Cashflow"]

        # Apply number format to all data cells (excluding header)
        for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row,
                                       min_col=1, max_col=worksheet.max_column):
            for cell in row:
                if isinstance(cell.value, (int, float)):
                    cell.number_format = '0.0'

        if calendar_data:
            pd.DataFrame(list(calendar_data.items()), columns=["Event", "Date"]).to_excel(writer, sheet_name="Calendar", index=False)
        if not yahoo_stats_df.empty:
            
            yahoo_stats_df.to_excel(writer, sheet_name="Key_Statistics", index=False)
        if not yahoo_analysis_df.empty:
            if yahoo_analysis_df.columns[0] == 0:  # If columns are numeric (default from HTML table)
                yahoo_analysis_df.columns = ["Metric", "Value"]  # Assign meaningful headers
            yahoo_analysis_df.to_excel(writer, sheet_name="Analysis", index=False)
    # print(yahoo_stats_df.head())
    # print(yahoo_analysis_df.head())

    print(f"✅ Exported data for {ticker}")
    print(f"📂 Saved at: {filepath}")

# ========== GUI ==========
def launch_gui():
    def on_fetch_batch():
        raw_input = text_box.get("1.0", tk.END).strip()
        if not raw_input:
            messagebox.showerror("Error", "Please enter one or more company tickers.")
            return

        tickers = [t.strip().upper() for t in raw_input.replace(",", "\n").splitlines() if t.strip()]
        if not tickers:
            messagebox.showerror("Error", "No valid tickers found.")
            return

        for ticker in tickers:
            try:
                fetch_and_export(ticker)
            except Exception as e:
                print(f"❌ Failed to fetch {ticker}: {e}")

        messagebox.showinfo("Done", f"Batch export completed for {len(tickers)} companies.")

    root = tk.Tk()
    root.title("Yahoo Finance Batch Scraper")

    tk.Label(root, text="Enter Company Tickers (comma or newline separated):").pack(pady=10)
    text_box = tk.Text(root, height=8, width=40)
    text_box.pack(pady=5)

    tk.Button(root, text="Fetch & Export All", command=on_fetch_batch).pack(pady=10)

    root.mainloop()

# ========== RUN ==========
if __name__ == "__main__":
    launch_gui()