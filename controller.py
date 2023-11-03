from tkinter import messagebox
import webbrowser
import os
import re
import numpy as np
import pandas as pd
from model import file_path_2023_Learning_Management ,file_path_2023_Learning_Completions,file_path_learning_record_21_22
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
            #os.remove(os.path.join(source_folder, filename))
            print("keep old data filename " + filename)
        else:
            # Generate the new file name by removing the timestamp
            new_filename = re.sub(pattern, ".xlsx", filename)
            
            # Rename the file
            os.rename(os.path.join(source_folder, filename), os.path.join(source_folder, new_filename))

def manipulation_powerTech():
    #create_model()
    learning_records_not_finished=data_learning_records_not_finished()
    #print(learning_records_not_finished)
    learning_records_completed=data_learning_records_completed()
    #print(learning_records_completed)
    learning_records_21_22=data_learning_records_21_22()
    #print(learning_records_21_22)
    result = pd.concat([learning_records_not_finished, learning_records_completed])
    result = pd.concat([result, learning_records_21_22])
    result = result.reset_index(drop=True)
    del learning_records_not_finished
    del learning_records_completed
    del learning_records_21_22
     # Setting 'Data Studio Training Type' based on 'Training Type'
    conditions = [
                    (result['Training Type'] == 'Event'),
                    (result['Training Type'] == 'Session')
                 ]
    choices = ['Event & Session', 'Event & Session']
    # Set values in 'Data Studio Training Type' based on 'Training Type' conditions
    result['Data Studio Training Type'] = np.select(conditions, choices, default=result['Training Type'])

    # Setting 'Data Studio Training status' based on 'Transcript - Transcript Status'
    conditions = [
                    (result['Transcript - Transcript Status Group'] == 'Completed'),
                    (result['Transcript - Transcript Status'] == 'Registered'),
                    (result['Transcript - Transcript Status'] == 'Registered / Past Due'),
                    (result['Transcript - Transcript Status'] == 'In Progress'),
                    (result['Transcript - Transcript Status'] == 'In Progress / Past Due'),
                    (result['Transcript - Transcript Status'] == 'Exception Requested'),
                    (result['Transcript - Transcript Status'] == 'Exception Requested / Past Due'),
                    (result['Transcript - Transcript Status'] == 'Pending Approval'),
                    (result['Transcript - Transcript Status'] == 'Pending Pre-Work'),
                    (result['Transcript - Transcript Status'] == 'Pending Prerequisite'),
                    (result['Transcript - Transcript Status'] == 'Pending Pre-Work/Past Due'),
                    (result['Transcript - Transcript Status'] == 'Pending Approval / Past Due'),
                    (result['Transcript - Transcript Status'] == 'Incomplete'),
                    (result['Transcript - Transcript Status'] == 'Incomplete / Past Due'),
                    (result['Transcript - Transcript Status'] == 'Denied'),
                    (result['Transcript - Transcript Status'] == 'Denied / Past Due')
                 ]
    choices = ['5.Completed','4.In Progress & Registered','4.In Progress & Registered','4.In Progress & Registered','4.In Progress & Registered',
               '3.Pending Process','3.Pending Process','3.Pending Process','3.Pending Process','3.Pending Process','3.Pending Process','3.Pending Process',
               '2.Incomplete & Event denied','2.Incomplete & Event denied','2.Incomplete & Event denied','2.Incomplete & Event denied']
    # Set values in 'Data Studio Training Type' based on 'Transcript - Transcript Status' conditions
    result['Data Studio Training status'] = np.select(conditions, choices, default="1.Not Started")

    result['Transcript Assigned Date'] = pd.to_datetime(result['Transcript Assigned Date'], format='mixed')
    result['Transcript Completed Date'] = pd.to_datetime(result['Transcript Completed Date'], format='mixed')

    result.sort_values(by=['ID', 'Training Title', 'Transcript Assigned Date'], ascending=[True, True, False], inplace=True)

    result = result.reset_index(drop=True)
    
    # Filter and keep the desired rows
    result = result.sort_values('Transcript Assigned Date').groupby(['ID', 'Training Title']).apply(
        lambda x: x[
        (x['Data Studio Training status'] == '5.Completed') & (x['Transcript Assigned Date'] == x['Transcript Assigned Date'].max())
            ].head(1) if '5.Completed' in x['Data Studio Training status'].values else (
         x[
            (x['Training Type'] == 'Session') & (x['Data Studio Training status'] != '2.Incomplete & Event denied') & 
             (x['Transcript Assigned Date'] == x['Transcript Assigned Date'].max())
         ].head(1) if ((x['Training Type'] == 'Session') & (x['Data Studio Training status'] != '2.Incomplete & Event denied')).any() else (
             x[x['Transcript Assigned Date'] == x['Transcript Assigned Date'].max()].head(1)) )
             
    ).reset_index(drop=True)

    result['BG'] = 'PTS'
    result.to_excel('result.xlsx', index=False)


def data_learning_records_not_finished():
    # Specify the columns you want to read
    columns_to_read = [
                    'User - User Last Name', 'User - User First Name', 'User - User ID','User - Business Group', 'User - Product Group', 
                    'User - Location Parent', 'User - Location', 'Training - Training Title', 'Training - Training Type', 
                    'Transcript - Transcript Assigned Date', 'Transcript - Transcript Status', 'Transcript - Transcript Status Group',
                    'Training - Training Provider', 'User - Position ID'
                    ]

    # Load the CSV file
    try:
        csv_data = pd.read_csv(file_path_2023_Learning_Management, header=6, usecols=columns_to_read)
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
        csv_data['Data Studio Training Type'] = None
        csv_data['Data Studio Training status'] = None
        csv_data['Transcript Completed Date'] = None
        csv_data['BG'] = None
        csv_data.rename(columns={'User - User Last Name': 'Last Name'}, inplace=True)
        csv_data.rename(columns={'User - User First Name': 'First Name'}, inplace=True)
        csv_data.rename(columns={'User - User ID': 'ID'}, inplace=True)
        csv_data.rename(columns={'User - Product Group': 'PG'}, inplace=True)
        csv_data.rename(columns={'User - Location Parent': 'Country'}, inplace=True)
        csv_data.rename(columns={'User - Location': 'Site'}, inplace=True)
        csv_data.rename(columns={'Training - Training Title': 'Training Title'}, inplace=True)
        csv_data.rename(columns={'Training - Training Type': 'Training Type'}, inplace=True)
        csv_data.rename(columns={'Transcript - Transcript Assigned Date': 'Transcript Assigned Date'}, inplace=True)
        csv_data.rename(columns={'Transcript - Transcript Completed Date': 'Transcript Completed Date'}, inplace=True)
        csv_data.rename(columns={'Training - Training Provider': 'Training Provider'}, inplace=True)
        csv_data.rename(columns={'User - Position ID': 'Position ID'}, inplace=True)
        columns_ordered = ['Last Name', 'First Name', 'ID', 'PG','Country','Site','Training Title','Training Type',
                           'Transcript Assigned Date','Transcript Completed Date','Transcript - Transcript Status','Transcript - Transcript Status Group',
                           'Training Provider','Data Studio Training Type','Data Studio Training status','Position ID','BG'
                          ]
        csv_data = csv_data[columns_ordered]
        # Reset the index to start from 0 and drop the current index
        csv_data = csv_data.reset_index(drop=True)
        return csv_data

    except FileNotFoundError:
        print("File not found. Please provide the correct file path.")
    except Exception as e:
        print("An error occurred:", e)

def data_learning_records_completed():
    # Specify the columns you want to read
    columns_to_read = [
                    'User - User Last Name', 'User - User First Name', 'User - User ID','BG', 'User - Product Group', 
                    'Country', 'Site', 'Training - Training Title', 'Training - Training Type', 
                    'Transcript - Transcript Assigned Date', 'Transcript - Transcript Status', 'Transcript - Transcript Completed Date',
                     'User - Manager - User ID', 'Training - Training Provider', 'User - Position ID'
                    ]

    # Load the CSV file
    try:
        csv_data = pd.read_csv(file_path_2023_Learning_Completions, header=6, usecols=columns_to_read)

        # Apply the transformation to the 'User - Location' column
        csv_data['Site'] = csv_data['Site'].apply(lambda x: x if "- " not in x else x.split("- ", 1)[1])
        # Filter based on "PTS Powertrain Systems" in the 'User - Product Group' column
        filtered_data_1 = csv_data[csv_data['BG'] == 'PTS Powertrain Systems']
        # Filter based on specific values in the 'Transcript - Transcript Status Group' column
        filter_values = ["Central R&D", "Group R&D", "PowerTECH Knowledge", "CDA Academy", "THS Academy", "VisiTech"]
        filtered_data_2 = filtered_data_1[filtered_data_1['Training - Training Provider'].isin(filter_values)]
        filter_values = ["Completed (Equivalent)", "Completed"]
        filtered_data_2 =  filtered_data_2[ filtered_data_2['Transcript - Transcript Status'].isin(filter_values)]
        # Drop the 'User - Business Group' column
        filtered_data_2.drop('BG', axis=1, inplace=True)
        csv_data=filtered_data_2
        # Deleting unnecessary variables
        del filtered_data_1
        del filtered_data_2
        csv_data['Data Studio Training Type'] = None
        csv_data['Data Studio Training status'] = None
        csv_data['Transcript - Transcript Status Group'] = 'Completed'
        csv_data['BG'] = None
        csv_data.rename(columns={'User - User Last Name': 'Last Name'}, inplace=True)
        csv_data.rename(columns={'User - User First Name': 'First Name'}, inplace=True)
        csv_data.rename(columns={'User - User ID': 'ID'}, inplace=True)
        csv_data.rename(columns={'User - Product Group': 'PG'}, inplace=True)
        csv_data.rename(columns={'Training - Training Title': 'Training Title'}, inplace=True)
        csv_data.rename(columns={'Training - Training Type': 'Training Type'}, inplace=True)
        csv_data.rename(columns={'Transcript - Transcript Assigned Date': 'Transcript Assigned Date'}, inplace=True)
        csv_data.rename(columns={'Transcript - Transcript Completed Date': 'Transcript Completed Date'}, inplace=True)
        csv_data.rename(columns={'Training - Training Provider': 'Training Provider'}, inplace=True)
        csv_data.rename(columns={'User - Position ID': 'Position ID'}, inplace=True)

        columns_ordered = ['Last Name', 'First Name', 'ID', 'PG','Country','Site','Training Title','Training Type',
                           'Transcript Assigned Date','Transcript Completed Date','Transcript - Transcript Status','Transcript - Transcript Status Group',
                           'Training Provider','Data Studio Training Type','Data Studio Training status','Position ID','BG'
                          ]
        csv_data = csv_data[columns_ordered]
        # Reset the index to start from 0 and drop the current index
        csv_data = csv_data.reset_index(drop=True)

        return csv_data

    except FileNotFoundError:
        print("File not found. Please provide the correct file path.")
    except Exception as e:
        print("An error occurred:", e)

def data_learning_records_21_22():
    # Specify the columns you want to read
    columns_to_read = [
                        'Last Name', 'First Name', 'ID', 'PG','Country','Site','Training Title','Training Type',
                        'Transcript Assigned Date','Transcript Completed Date','Transcript - Transcript Status','Transcript - Transcript Status Group',
                        'Training Provider','Data Studio Training Type','Data Studio Training status','Position ID','BG'
                      ]

    # Load the CSV file
    try:
        csv_data = pd.read_csv(file_path_learning_record_21_22,header=0,usecols=columns_to_read,delimiter=';')

        columns_ordered = ['Last Name', 'First Name', 'ID', 'PG','Country','Site','Training Title','Training Type',
                           'Transcript Assigned Date','Transcript Completed Date','Transcript - Transcript Status','Transcript - Transcript Status Group',
                           'Training Provider','Data Studio Training Type','Data Studio Training status','Position ID','BG'
                          ]
        csv_data = csv_data[columns_ordered]
        # Reset the index to start from 0 and drop the current index
        csv_data = csv_data.reset_index(drop=True)

        return csv_data

    except FileNotFoundError:
        print("File not found. Please provide the correct file path.")
    except Exception as e:
        print("An error occurred:", e)

manipulation_powerTech()
