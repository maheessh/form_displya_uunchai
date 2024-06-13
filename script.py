import pandas as pd
import colorama
from colorama import Fore, Style
import os

def main():
    colorama.init(autoreset=True)
    file_name = 'data.xlsx'
    
    # Check if the file exists
    if not os.path.isfile(file_name):
        print(Fore.RED + f"Error: The file '{file_name}' does not exist.")
        return

    while True:
        worksheet_name = input("Enter the worksheet name (or type 'exit' to quit): ")
        
        if worksheet_name.lower() == 'exit':
            print("Exiting the program.")
            break

        try:
            # Get the list of sheets to check if the provided sheet name exists
            xls = pd.ExcelFile(file_name)
            if worksheet_name not in xls.sheet_names:
                print(Fore.RED + f"Error: The worksheet '{worksheet_name}' does not exist in '{file_name}'.")
                continue
            
            # Read the specified worksheet
            df = pd.read_excel(file_name, sheet_name=worksheet_name)
            print(Fore.GREEN + "Data loaded successfully!")
            print(df.head())  # Print the first few rows of the dataframe
            break
        except Exception as e:
            print(Fore.RED + "Error loading data: " + str(e))

    # Wait for user input before exiting
    input("Press Enter to exit...")

if __name__ == "__main__":
    main()
