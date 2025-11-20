import sys
import os
import subprocess

def run_script(script, args=None):
    cmd = [sys.executable, script]
    if args:
        cmd.extend(args)
    print(f"\nRunning: {' '.join(cmd)}")
    result = subprocess.run(cmd)
    if result.returncode != 0:
        print(f"Error: {script} failed with exit code {result.returncode}")
        sys.exit(result.returncode)

def main():
    if len(sys.argv) < 2:
        print("Usage: python master_auction_sales_loader.py <QuickSearch.csv>")
        sys.exit(1)
    csv_path = sys.argv[1]
    # Step 1: Load CSV to Auctions
    run_script(os.path.join("backend", "process_quicksearch_to_auctions.py"), [csv_path])
    # Step 2: Update Sales table from Auctions
    # (Assume process_quicksearch_to_sales.py exists and takes county and date as args)
    import re
    base = os.path.basename(csv_path)
    m = re.match(r"([A-Za-z]+)[-_](\d{8})[_-]?QuickSearch\.csv", base, re.IGNORECASE)
    if not m:
        print(f"Could not parse county/date from filename: {base}")
        sys.exit(1)
    county, date = m.group(1), m.group(2)
    run_script(os.path.join("backend", "process_quicksearch_to_sales.py"), [county, date])
    print("\nAll steps complete.")

if __name__ == "__main__":
    main()
