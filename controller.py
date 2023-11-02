from tkinter import messagebox
import webbrowser
import os
import re
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
def create_model():
     messagebox.showinfo("")

def download_data():
    download_url = "https://www.example.com/download"
    webbrowser.open_new(download_url)


def update_data():
    # Define the regular expression pattern to match the timestamp
    pattern = r"_\d{8}_\d{2}_\d{2}_\d{2}_\w{2}\.xlsx"

    # Define the source folder path
    source_folder = "C:/Users/yli6/Downloads/test/"  # Adjust this path to your source folder

    # Loop through each file in the folder
    for filename in os.listdir(source_folder):
        if not re.search(pattern, filename):  # Check if the file name does not match the pattern
            # Delete the file
            os.remove(os.path.join(source_folder, filename))
        else:
            # Generate the new file name by removing the timestamp
            new_filename = re.sub(pattern, ".xlsx", filename)
            
            # Rename the file
            os.rename(os.path.join(source_folder, filename), os.path.join(source_folder, new_filename))

def manipulation_powerTech():
    #create_model()
    #learning_records_not_finished=data_learning_records_not_finished()
    #print(learning_records_not_finished[learning_records_not_finished['Training - Training Start Date'].notnull()]['Training - Training Start Date'])
    learning_records_completed=data_learning_records_completed()

def data_learning_records_not_finished():
    file_path = 'C:\\Users\\yli6\\Downloads\\source\\[Learning]_2023_Learning_Management_(In_Progress__Not_Started__Others)_20231031_02_02_28_PM.csv'

    # Specify the columns you want to read
    columns_to_read = [
                    'User - User Last Name', 'User - User First Name', 'User - User ID','User - Business Group', 'User - Product Group', 
                    'User - Location Parent', 'User - Location', 'Training - Training Title', 'Training - Training Type', 
                    'Transcript - Transcript Assigned Date', 'Transcript - Transcript Status', 'Transcript - Transcript Status Group',
                    'Training - Training Hours', 'User - Manager - User ID', 'Training - Training Start Date', 'Training - Training End Date', 
                    'Training - Training Provider', 'User - Position ID'
                    ]

    # Load the CSV file
    try:
        csv_data = pd.read_csv(file_path, header=6, usecols=columns_to_read)
        # Apply the transformation to the 'User - Location' column
        csv_data['User - Location'] = csv_data['User - Location'].apply(lambda x: x if "- " not in x else x.split("- ", 1)[1])
        # Filter based on "PTS Powertrain Systems" in the 'User - Product Group' column
        filtered_data_1 = csv_data[csv_data['User - Business Group'] == 'PTS Powertrain Systems']
        # Filter based on specific values in the 'Transcript - Transcript Status Group' column
        filter_values = ["Central R&D", "Group R&D", "PowerTECH Knowledge", "CDA Academy", "THS Academy", "VisiTech"]
        filtered_data_2 = filtered_data_1[filtered_data_1['Training - Training Provider'].isin(filter_values)]
        # Drop the 'User - Business Group' column
        filtered_data_2.drop('User - Business Group', axis=1, inplace=True)
        csv_data=filtered_data_2
        # Deleting unnecessary variables
        del filtered_data_1
        del filtered_data_2
        # Change the date format
        # This line changes the date format to 'mm/dd/yyyy hh:mm AM/PM'
        csv_data['Training - Training Start Date'] = pd.to_datetime(csv_data['Training - Training Start Date']).dt.strftime('%m/%d/%Y %I:%M %p')
        # Change the date format
        # This line changes the date format to 'mm/dd/yyyy hh:mm AM/PM'
        csv_data['Training - Training End Date'] = pd.to_datetime(csv_data['Training - Training End Date']).dt.strftime('%m/%d/%Y %I:%M %p')
        
        #print(csv_data)
        #print(csv_data[csv_data['Training - Training Start Date'].notnull()]['Training - Training Start Date'])
        return csv_data

    except FileNotFoundError:
        print("File not found. Please provide the correct file path.")
    except Exception as e:
        print("An error occurred:", e)

def data_learning_records_completed():
    file_path = 'C:\\Users\\yli6\\Downloads\\source\\[Learning]_2023_Learning_Completions_20231031_04_08_28_PM.csv'

    # Specify the columns you want to read
    columns_to_read = [
                    'User - User Last Name', 'User - User First Name', 'User - User ID','BG', 'User - Product Group', 
                    'Country', 'Site', 'Training - Training Title', 'Training - Training Type', 
                    'Transcript - Transcript Assigned Date', 'Transcript - Transcript Status', 'Transcript - Transcript Completed Date',
                    'Training - Training Hours', 'User - Manager - User ID', 
                    'Training - Training Provider', 'User - Position ID'
                    ]

    # Load the CSV file
    try:
        csv_data = pd.read_csv(file_path, header=6, usecols=columns_to_read)
        print(csv_data)

        # Apply the transformation to the 'User - Location' column
        csv_data['User - Location'] = csv_data['User - Location'].apply(lambda x: x if "- " not in x else x.split("- ", 1)[1])
        # Filter based on "PTS Powertrain Systems" in the 'User - Product Group' column
        filtered_data_1 = csv_data[csv_data['BG'] == 'PTS Powertrain Systems']
        # Filter based on specific values in the 'Transcript - Transcript Status Group' column
        filter_values = ["Central R&D", "Group R&D", "PowerTECH Knowledge", "CDA Academy", "THS Academy", "VisiTech"]
        filtered_data_2 = filtered_data_1[filtered_data_1['Training - Training Provider'].isin(filter_values)]
        # Drop the 'User - Business Group' column
        filtered_data_2.drop('BG', axis=1, inplace=True)
        csv_data=filtered_data_2
        # Deleting unnecessary variables
        del filtered_data_1
        del filtered_data_2
        # Change the date format
        # This line changes the date format to 'mm/dd/yyyy hh:mm AM/PM'
        csv_data['Training - Training Start Date'] = pd.to_datetime(csv_data['Training - Training Start Date']).dt.strftime('%m/%d/%Y %I:%M %p')
        # Change the date format
        # This line changes the date format to 'mm/dd/yyyy hh:mm AM/PM'
        csv_data['Training - Training End Date'] = pd.to_datetime(csv_data['Training - Training End Date']).dt.strftime('%m/%d/%Y %I:%M %p')
        
        #print(csv_data)
        #print(csv_data[csv_data['Training - Training Start Date'].notnull()]['Training - Training Start Date'])
        return csv_data

    except FileNotFoundError:
        print("File not found. Please provide the correct file path.")
    except Exception as e:
        print("An error occurred:", e)

manipulation_powerTech()
