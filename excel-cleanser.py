# 1st Party Libraries
import os
import sys
import time

# 3rd Party Libraries
import pandas as pd
from pathlib import Path
from datetime import date

# Name introduction by the user
your_name = input("\nHello there! Before we introduce ourselves, may I have your name please?\n\nYou can enter it here: ")

# Greeting
print(f"\nHello {your_name}, nice to meet you. I can certainly assist you with your request! :) \n\nRemember prior to convert from .xlsx to .csv, make sure to delete the rows after the last one with useful data!\n\nThis has the purpose to do not affect the consolidation process, thank you :)\n")

# Input to determine if the user requires to convert different Excel files from a folder to CSV
yes_no = input("Before we proceed, would you like to convert Excel Files to CSV? (Yes/No): ")

# User is required to enter Yes or yes to proceed
if yes_no == 'Yes' or yes_no == 'yes':
    # Input where the user needs to indicate the path where the files are allocated
    location_xlsx = input('\nPlease, enter the folder path to convert .xlsx to .csv: ')
    
    # Unless the file exists, the input will continuously request to enter an existing path
    while not os.path.exists(location_xlsx):
        print("\nSorry to tell you this, but the path doesn't exist.")
        location_xlsx = input('\nPlease, enter the folder path to convert .xlsx to .csv: ')
    
    # Path is adjusted to a specific format avoiding any possible issues
    folder_xlsx = os.path.abspath(location_xlsx)
    
    # Listing all files from path
    files_xlsx = os.listdir(folder_xlsx)

    # Iterates through all over the files from the folder whose extension is ending in .xlsx to be read and changed to CSV
    for file_xlsx in files_xlsx:
        # At this point files must have the .xlsx extension
        if file_xlsx.endswith('.xlsx'):
            # Reading each Excel file
            convert_csv = pd.read_excel(f'{folder_xlsx}/{file_xlsx}')
            # File extension renaming
            new_name_csv = file_xlsx.replace('.xlsx', '')
            # Changes all Excel files to CSV
            convert_csv.to_csv(f'{folder_xlsx}/{new_name_csv}.csv', encoding = 'UTF-8', index = False, header = False)
    
    # Printing success message of CSV modifications applied
    print(f"\nThe process has been completed {your_name}! Hurrah :D")
    
    # Option available if the user requires to consolidate files
    Yes_No = input("\nWould you like to consolidate your files? (Yes/No): ")
    
    # If the option is Yes the user is required to enter the path once again for CSV file consolidation
    if Yes_No =='Yes' or Yes_No == 'yes':
        location_1 = input('\nPlease, re-enter the folder path to consolidate CSV files: ')
        
        # Unless the file exists, the input will continuously request to enter an existing path
        while not os.path.exists(location_1):
            print("\nSorry to tell you this, but the path doesn't exist.")
            location_1 = input('\nPlease, re-enter the folder path to consolidate CSV files: ')
        
        # Path is adjusted to a specific format avoiding any possible issues
        folder_1 = os.path.abspath(location_1)

        # Listing all files from path
        files_1 = os.listdir(folder_1)

        # Initializing empty DataFrame prepared to consolidating files in a single one
        df_total_1 = pd.DataFrame()
        
        # Reviewing if all files have the .csv extension
        for file_1 in files_1:
            # At this point files must have the .xlsx extension, this for loop operates individually and merges all DataFrames after completing
            if file_1.endswith('.csv'):
                # Reading each CSV file
                df = pd.read_csv(f'{folder_1}/{file_1}', sep = ',', header = None)
                # Identifying empty rows
                num = df[df.columns[-1]].first_valid_index()
                # Dropping empty rows
                df = df.drop(range(0, num))
                # Setting header
                df = df.rename(columns = df.loc[num])
                # Deleting empty columns
                df = df.dropna(how = 'all', axis = 1)
                # Deleting empty rows
                df = df.dropna(how = 'all', axis = 0)
                # Removing repeated headers
                df_total_1 = df_total_1[df_total_1[df_total_1.columns[0]] != df_total_1.columns[0]]
                # Appending DataFrames into a single one
                df_total_1 = df_total_1.append(df)
        
        # User needs to indicate once again the path where the consolidated file will be stored
        download_1 = input(f"\nWoo-hoo! You rock {your_name}! We have unified all files within this folder!\n\nNow please but not last, enter the folder path to save the consolidated file: ")
        
        # Unless the file exists, the input will continuously request to enter an existing path
        while not os.path.exists(download_1):
            print("\nSorry to tell you this, but the path doesn't exist.")
            download_1 = input('\nPlease, enter the folder path to save the consolidated file: ')
        
        # Path is adjusted to a specific format avoiding any possible issues
        new_folder_1 = os.path.abspath(download_1)     
        
        # Obtaining current date
        today_1 = date.today()
        
        # Changing current date to string
        to_string_1 = str(today_1)
        
        # Starting Excel file generation time
        t0_1 = time.time()
        
        # Consolidated Excel file generation
        df_total_1.to_excel(f'{new_folder_1}/Consolidate_File {to_string_1}.xlsx', index = False)
        print(f"\nConsolidate process completed in: {to_string_1}")
        
        # Finishing Excel file generation time
        t1_1 = time.time()

        # Calculating time duration
        total_1 = t1_1 - t0_1

        # Seconds calculation
        total_2_1 = total_1 % 60

        # Minutes calculation
        total_3_1 = total_1 // 60

        # If the process took less than five minutes
        if total_3_1 < 5:
            print(f"\nTime elapsed of process: {total_3_1:.0f} minutes {total_2_1:.2f} seconds.\n\nThat was fast! It didn't took so long to complete.\n\nHope to see you soon {your_name}! Have a wonderful day :)\n")
            time.sleep(3)

        # More than five minutes
        else:
            print(f"\nTime elapsed of process: {total_3_1:.0f} minutes {total_2_1:.2f} seconds.\n\nSorry that it took you so long, those files were kind of heavy :( \n\nAt least all files have been consolidated!\n\nHope to see you soon {your_name}! Have a wonderful day :)\n")
            time.sleep(3)
    
    # In case that the user does not require to change from  
    elif Yes_No == 'No' or Yes_No == 'no': 
        print(f"\nYour files have been converted {your_name}. Have a nice day :D\n")
        time.sleep(3)
        sys.exit()
    else:
        print("\nInvalid Option! Have a nice day :D\n")
        time.sleep(3)
        sys.exit()

# Option for file straightforward file consolidation
elif yes_no == 'No' or yes_no=='no':    
    Yes_No = input("\nWould you like to consolidate your files? (Yes/No): ")

    # If the answer is Yes or yes, the user needs to indicate the path for file consolidation
    if Yes_No =='Yes' or Yes_No == 'yes':
        location = input("\nPlease, enter the folder path to begin the consolidation: ")

        # Checking the existence of the path
        folder = os.path.abspath(location)

        # Unless the file exists, the input will continuously request to enter an existing path
        while not os.path.exists(location):
            print("\nSorry to tell you this, but the path doesn't exist.")
            location = input('\nPlease, enter the folder path to begin the consolidation: ')
        
        # Initializing empty DataFrame prepared to consolidating files in a single one
        df_total = pd.DataFrame()

        # Path is adjusted to a specific format avoiding any possible issues
        folder = os.path.abspath(location)

        # Listing all files from path
        files = os.listdir(folder)
        
        # Placing the delimiter
        delimiter = input('\nPlease indicate if the files are separated with (,) or (;): ')
        
        # Unless is a valid delimeter, the input will continuously request to enter a correct one
        while delimiter not in (',',';'):
            print('\nInvalid option!')
            delimiter = input('\nPlease indicate if the files are separated with (,) or (;): ')
        
        # Reviewing if all files have the .csv extension
        for file in files:
            # At this point files must have the .xlsx extension, this for loop operates individually and merges all DataFrames after completing
            if file.endswith('.csv'):
                # Reading each CSV file
                df = pd.read_csv(f'{folder}/{file}', sep = str(delimiter), error_bad_lines = False, index_col = False, dtype = 'unicode', header = None, low_memory = False)
                # Dropping first row
                df = df.drop(df.index[0])
                # Identifying empty rows
                num = df[df.columns[-1]].first_valid_index()
                # Dropping empty rows
                df = df.drop(range(0,num))
                # Setting header
                df_total = df_total.rename(columns = df_total.loc[num])
                # Defining headers
                df_total = df_total.set_index(df_total.columns[0])
                # Deleting empty columns
                df_total = df_total.dropna(how = 'all', axis = 1)
                # Deleting empty rows
                df_total = df_total.dropna(how = 'all', axis = 0)
                # Appending DataFrames into a single one
                df_total = df_total.append(df)
                # Dropping rows equivalent to headers
                df_total = df_total[df_total[df_total.columns[0]] != df_total.columns[0]]
        
        # User needs to indicate once again the path where the consolidated file will be stored
        download = input(f"\nWoo-hoo! You rock {your_name}! We have unified all files within this folder!\n\nNow please but not last, enter the path to save the consolidated file: ") 
        
        # Unless the file exists, the input will continuously request to enter an existing path
        while not os.path.exists(download):
            print("\nSorry to tell you this, but the path doesn't exist.")
            download = input('\nPlease, enter the folder path to save the consolidated file: ')
        
        # Path is adjusted to a specific format avoiding any possible issues
        new_folder = os.path.abspath(download)
        
        # Obtaining current date
        today = date.today()

        # Changing current date to string
        to_string = str(today)

        # Starting Excel file generation time
        t0 = time.time()

        # Consolidated Excel file generation
        df_total.to_excel(f'{new_folder}/Consolidate_File {to_string}.xlsx',index = False)
        print(f"\nConsolidate process completed in: {to_string}")

        # Finishing Excel file generation time
        t1 = time.time()

        # Calculating time duration
        total = t1 - t0

        # Seconds calculation
        total_2 = total % 60

        # Minutes calculation
        total_3 = total // 60

        # If the process took less than five minutes
        if total_3 < 5:
            print(f"\nTime elapsed of process: {total_3:.0f} minutes {total_2:.2f} seconds.\n\nThat was fast! It didn't took so long to complete.\n\nHope to see you soon {your_name}! Have a wonderful day :)\n")
            time.sleep(3)
        
        # More than five minutes
        else:
            print(f"\nTime elapsed of process: {total_3:.0f} minutes {total_2:.2f} seconds.\n\nSorry that it took you so long, those files were kind of heavy :( \n\nAt least all files have been consolidated!\n\nHope to see you soon {your_name}! Have a wonderful day :)\n")
            time.sleep(3)
    
    # If no file consolidation is required
    else:
        print("\nHave a nice day :D\n")
        time.sleep(3)
        sys.exit()  

# For incorrect options entered program exits
else: 
    print("\nInvalid Option! Have a nice day :D\n")
    time.sleep(3)
    sys.exit()