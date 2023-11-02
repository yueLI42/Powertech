import tkinter as tk
from tkinter import messagebox

from controller import download_data, update_data

# Function to download data
def download_data_interface():
    # Add your code here for downloading data from MyLearning platform
    # For example:
    # data = your_download_function()
    messagebox.showinfo("Download", "Data downloaded successfully!")

# Function to update the data
def update_data_interface():
    # Add your code here for updating the data
    # For example:
    # updated_data = your_update_function()
    messagebox.showinfo("Update", "Data updated successfully!")

# Create the main window
def create_window():
    window = tk.Tk()
    window.title("Data Manipulation PowerTech")

    # Setting a larger window size
    window.geometry("400x300")

    # Function to handle button 1 (Download)
    def handle_download():
        download_data_interface()

    # Function to handle button 2 (Update)
    def handle_update():
        update_data_interface()

    # Creating buttons
    download_button = tk.Button(window, text="Download Data", command=handle_download)
    download_button.pack(pady=20)

    update_button = tk.Button(window, text="Update Data", command=handle_update)
    update_button.pack(pady=20)

    # Function to exit the application
    def close_window():
        window.destroy()

    exit_button = tk.Button(window, text="Exit", command=close_window)
    exit_button.pack(pady=20)

    window.mainloop()

if __name__ == "__main__":
    #download_data()
    update_data()
    #create_window()
