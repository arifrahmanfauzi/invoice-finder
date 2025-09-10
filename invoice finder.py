import pandas as pd
import requests
import time
from tqdm import tqdm
import sys
from colorama import Fore, Style, init

# Initialize colorama for cross-platform colored output
init(autoreset=True)

def print_banner():
    """Display a nice banner"""
    print(f"\n{Fore.CYAN}{'='*60}")
    print(f"{Fore.CYAN}📊 Excel HTTP Transaction Checker")
    print(f"{Fore.CYAN}{'='*60}\n")

def get_excel_sheets(file_path):
    """Get list of available sheets in Excel file"""
    try:
        xl_file = pd.ExcelFile(file_path)
        return xl_file.sheet_names
    except Exception as e:
        print(f"{Fore.RED}❌ Error reading Excel file: {e}")
        return None

def select_sheet(sheets):
    """Let user select which sheet to process"""
    print(f"{Fore.YELLOW}📋 Available sheets:")
    for i, sheet in enumerate(sheets, 1):
        print(f"  {i}. {sheet}")
    
    while True:
        try:
            choice = int(input(f"\n{Fore.GREEN}Select sheet number: ")) - 1
            if 0 <= choice < len(sheets):
                return sheets[choice]
            else:
                print(f"{Fore.RED}Invalid choice. Please select 1-{len(sheets)}")
        except ValueError:
            print(f"{Fore.RED}Please enter a valid number")

def load_transaction_numbers(file_path, sheet_name):
    """Load transaction numbers from Excel file"""
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        
        # Look for Transaction Number column (case insensitive)
        trans_col = None
        for col in df.columns:
            #nomor transaksi(is the lowercase of the column to be searched)
            if 'nomor' in col.lower() and 'transaksi' in col.lower():
                trans_col = col
                break
        
        if trans_col is None:
            print(f"{Fore.RED}❌ 'Transaction Number' column not found!")
            print(f"{Fore.YELLOW}Available columns: {list(df.columns)}")
            return None
        
        # Remove NaN values and convert to string
        transactions = df[trans_col].dropna().astype(str).tolist()
        return transactions
    
    except Exception as e:
        print(f"{Fore.RED}❌ Error loading Excel file: {e}")
        return None

def check_transaction(transaction_id, session):
    """Check single transaction via HTTP GET"""
    url = f"https://example.com/core/api/v1/sales_invoices/{transaction_id}"
    headers = {"apikey": "dv14weccsdqwe12ewq"}
    
    try:
        response = session.get(url, headers=headers, timeout=10)
        return response.status_code, response
    except requests.exceptions.RequestException as e:
        return None, str(e)

def main():
    print_banner()
    
    # Get Excel file path
    excel_file = input(f"{Fore.GREEN}📁 Enter Excel file path: ").strip('"')
    
    # Get available sheets
    sheets = get_excel_sheets(excel_file)
    if not sheets:
        return
    
    # Select sheet
    selected_sheet = select_sheet(sheets)
    print(f"\n{Fore.GREEN}✅ Selected sheet: {selected_sheet}")
    
    # Load transaction numbers
    print(f"\n{Fore.YELLOW}📖 Loading transaction numbers...")
    transactions = load_transaction_numbers(excel_file, selected_sheet)
    
    if not transactions:
        return
    
    print(f"{Fore.GREEN}✅ Loaded {len(transactions)} transactions")
    
    # Setup session for connection pooling
    session = requests.Session()
    
    # Results tracking
    not_found = []
    success_count = 0
    error_count = 0
    
    # Process with progress bar
    print(f"\n{Fore.CYAN}🔄 Processing transactions...")
    
    with tqdm(transactions, desc="Checking", ncols=80, 
              bar_format="{l_bar}{bar}| {n_fmt}/{total_fmt} [{elapsed}<{remaining}]") as pbar:
        
        for transaction_id in pbar:
            # Update progress description
            pbar.set_description(f"Checking {transaction_id[:10]}...")
            
            # Make request
            status_code, response = check_transaction(transaction_id, session)
            
            if status_code == 422:
                not_found.append(transaction_id)
                pbar.write(f"{Fore.RED}❌ Not found: {transaction_id}")
            elif status_code == 200:
                success_count += 1
                pbar.write(f"{Fore.GREEN}✅ Found: {transaction_id}")
            elif status_code is None:
                error_count += 1
                pbar.write(f"{Fore.YELLOW}⚠️  Error: {transaction_id} - {response}")
            else:
                pbar.write(f"{Fore.YELLOW}⚠️  Status {status_code}: {transaction_id}")
            
            # Rate limiting - 1 request per second
            time.sleep(1)
    
    # Write not found transactions to file
    if not_found:
        with open('not_found_transactions.txt', 'w') as f:
            f.write("Not Found Transaction Numbers:\n")
            f.write("=" * 30 + "\n")
            for trans_id in not_found:
                f.write(f"{trans_id}\n")
        
        print(f"\n{Fore.YELLOW}📝 Written {len(not_found)} not found transactions to 'not_found_transactions.txt'")
    
    # Final summary
    print(f"\n{Fore.CYAN}{'='*60}")
    print(f"{Fore.GREEN}✅ Success: {success_count}")
    print(f"{Fore.RED}❌ Not Found (422): {len(not_found)}")
    print(f"{Fore.YELLOW}⚠️  Errors: {error_count}")
    print(f"{Fore.CYAN}{'='*60}\n")

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print(f"\n{Fore.RED}🛑 Process interrupted by user")
        sys.exit(1)
