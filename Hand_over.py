import os
import shutil
from glob import glob
from tkinter import font
import pandas as pd
import time
from plyer import notification
import schedule
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.firefox.service import Service
from webdriver_manager.firefox import GeckoDriverManager
import win32com.client as win32
import getpass
from datetime import datetime
from datetime import date
import sys
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import win32com.client
import tkinter.simpledialog

status_label = None
root = None
def get_outlook_email():
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    return namespace.Accounts[0].DisplayName
sender_email = get_outlook_email()

current_date = date.today().strftime("%m-%d-%Y")

# Global variable to track if export has been done
export_done = False

def click_button():
    """
    Performs the button click action and downloads the file to Desktop/handover
    """
    options = Options()
    
    username = getpass.getuser()
    desktop_path = rf'#Give path where you want to save your exported sheet'
    download_dir = os.path.join(desktop_path)
    
    options.set_preference("browser.download.folderList", 2)
    options.set_preference("browser.download.manager.showWhenStarting", False)
    options.set_preference("browser.download.dir", download_dir)
    options.set_preference("browser.download.useDownloadDir", True)
    options.set_preference("browser.helperApps.neverAsk.saveToDisk", "text/csv")

    firefox_profile_directory = rf"C:\Users\{username}\AppData\Roaming\Mozilla\Firefox\Profiles"
    profile_folders = [folder for folder in os.listdir(firefox_profile_directory) if folder.endswith(".default-esr")]

    if not profile_folders:
        raise Exception("Firefox profile not found!")

    firefox_default_profile_directory = os.path.join(firefox_profile_directory, profile_folders[0])
    options.profile = firefox_default_profile_directory

    try:
        print("Initializing Firefox WebDriver...")
        driver = webdriver.Firefox(service=Service(GeckoDriverManager().install()), options=options)
        driver.get('#Link where you want to download the export')
        
        WebDriverWait(driver, 45).until(ec.visibility_of_element_located((By.XPATH, "//span[text()='Export all']")))
        button_one = driver.find_element(By.XPATH, "//span[text()='Export all']")
        button_one.click()
       
        WebDriverWait(driver, 45).until(ec.visibility_of_element_located((By.XPATH, "//span[text()='Download CSV']")))
        button_two = driver.find_element(By.XPATH, "//span[text()='Download CSV']")
        button_two.click()
        
        notification.notify(title="Export Done", message="Successfully Exported the queue!")

    except Exception as e:
        notification.notify(title="Error", message=f"Error occurred: {str(e)}")
        print(f"Error occurred: {str(e)}")
    
    finally:
        driver.quit()
def get_latest_csv(folder_path):
    """
    Get the latest CSV file from the specified folder based on the modification time.
    """
    csv_files = [f for f in os.listdir(folder_path) if f.endswith('.csv')]
    
    if not csv_files:
        raise FileNotFoundError("No CSV files found in the specified folder.")
    
    full_paths = [os.path.join(folder_path, f) for f in csv_files]
    latest_file = max(full_paths, key=os.path.getmtime)
    
    return latest_file

def create_additional_columns(df):
    """
    Adds 5 additional columns to the DataFrame.
    """
    df['Age Of the Ticket'] = " "
    df['Root Cause Summary'] = " "
    df['Action Summary'] = " "
    df['Recent Update from Team'] = " "
    df['Next Steps from Team'] = " "
    return df

def create_outlook_email(recipient, subject, table_headers, table_rows, sender_email={sender_email}):
    """
    Creates and displays an Outlook email with a customized table.
    """
    try:
        outlook = win32.Dispatch('outlook.application')
    except Exception as e:
        print(f"Error initializing Outlook: {e}")
        sys.exit(1)

    mail = outlook.CreateItem(0)
    mail.To = recipient
    mail.Subject = subject
    mail.BodyFormat = 2

    html_body = """
    <html>
    <body>
        <p>Hello Team,</p>
        <p><b><u>List of open High Severity tickets in our queue:</u></b></p>
        <table style="border-collapse: collapse; width: 100%;">
    """

    if table_headers:
        html_body += "<tr>"
        for header in table_headers:
            html_body += f"""
                <th style="
                    border: 1px solid black;
                    padding: 8px;
                    background-color: black;
                    color: white;
                    text-align: left;
                ">{header}</th>
            """
        html_body += "</tr>"

    for row in table_rows:
        html_body += "<tr>"
        for cell in row:
            html_body += f"""
                <td style="
                    border: 1px solid black;
                    padding: 8px;
                ">{cell}</td>
            """
        html_body += "</tr>"

    html_body += f"""
        </table>
        <br>
        <p>Best regards,<br>{sender_email}</p>
    </body>
    </html>
    """

    mail.HTMLBody = html_body
    mail.Display()
    formatted_html = html_body.format(sender_email=sender_email)
    print(formatted_html)

def create_html_table_with_data(df, mapping, row_index):
    try:
        def get_mapped_value(key, with_link=False):
            try:
                if key in mapping:
                    column_name = mapping[key]
                    if column_name in df.columns:
                        value = str(df[column_name].iloc[row_index])
                        if with_link and f'{key}_link' in mapping:
                            link_column = mapping[f'{key}_link']
                            if link_column in df.columns:
                                url = str(df[link_column].iloc[row_index])
                                return f'<a href="{url}" target="_blank">{value}</a>'
                        return value
                return ''
            except Exception as e:
                print(f"Error getting value for {key}: {e}")
                return ''

        html_table = f"""
        <table class="custom-table">
           <tr>
                <td class="first-column"><b><center>Ticket ID</center></b></td>
                <td colspan="8"><center><u>{get_mapped_value('a', with_link=True)}</u></center></td>
            </tr>
            <tr>
                <td class="first-column"><b><center>Problem Statement</center></b></td>
                <td colspan="8"><center>{get_mapped_value('b')}</center></td>
            </tr>
            <tr>
                <td class="first-column"><b><center>High Severity Justification</center></b></td>
                <td colspan="8"><center>{get_mapped_value('c')}</center></td>
            </tr>
            <tr>
                <td class="first-column" rowspan="2"><b><center>Tickets Details</center></b></td>
                <td class="header-row"><center>Created Date</center></td>
                <td class="header-row"><center>Resolver Group</center></td>
                <td class="header-row"><center>Overall Age</center></td>
                <td class="header-row"><center>Days Open as High Severity</center></td>
                <td class="header-row"><center>Days Open as Low Severity</center></td>
                <td class="header-row"><center>Assigned date to Ops (As High Severity)</center></td>
                <td class="header-row"><center>Last Assigned Date</center></td>
                <td class="header-row yellow-cell"><center>Status of the ticket</center></td>
            </tr>
            <tr>
                <td><center>{get_mapped_value('j')}</center></td>
                <td><center>{get_mapped_value('k')}</center></td>
                <td><center>{get_mapped_value('l')}</center></td>
                <td><center>{get_mapped_value('m')}</center></td>
                <td><center>{get_mapped_value('n')}</center></td>
                <td><center>{get_mapped_value('o')}</center></td>
                <td><center>{get_mapped_value('p')}</center></td>
                <td class="yellow-cell"><center>{get_mapped_value('q')}</center></td>
            </tr>
            <tr>
                <td class="first-column" rowspan="2"><b><center>Title Details</center></b></td>
                <td class="header-row"><center>Title Name</center></td>
                <td class="header-row"><center>Partner Name</center></td>
                <td class="header-row"><center>POM Name</center></td>
                <td class="header-row"><center>POM Action required (Yes/No)</center></td>
                <td class="header-row"><center>Title type (High-Valued Content, Non HVC)</center></td>
                <td class="header-row"><center>Series/Movie</center></td>
                <td class="header-row"><center>Season and Episode Details</center></td>
                <td class="header-row"><center>Region</center></td>
            </tr>
            <tr>
                <td><center>{get_mapped_value('r')}</center></td>
                <td><center>{get_mapped_value('s')}</center></td>
                <td><center>{get_mapped_value('t')}</center></td>
                <td><center>{get_mapped_value('u')}</center></td>
                <td><center>{get_mapped_value('v')}</center></td>
                <td><center>{get_mapped_value('w')}</center></td>
                <td><center>{get_mapped_value('x')}</center></td>
                <td><center>{get_mapped_value('y')}</center></td>
            </tr>
            <tr>
                <td class="first-column"><b><center>Root Cause Summary</center></b></td>
                <td colspan="8"><center>{get_mapped_value('f')}</center></td>
            </tr>
            <tr>
                <td class="first-column"><b><center>Action Summary</center></b></td>
                <td colspan="8"><center>{get_mapped_value('g')}</center></td>
            </tr>
            <tr>
                <td class="first-column"><b><center>Help Required: \nYes/No</center></b></td>
                <td colspan="8"><center>{get_mapped_value('h')}</center></td>
            </tr>
            <tr>
                <td class="first-column"><b><center>Next Steps to be performed by DF MAD</center></b></td>
                <td colspan="8"><center>{get_mapped_value('i')}</center></td>
            </tr>
        </table>
        """
        return html_table

    except Exception as e:
        print(f"Error creating table: {e}")
        return None
def normal_flow():
    folder_path = os.path.join(os.path.expanduser('~'), 'Documents', 'handover')

    try:
        latest_csv_file = get_latest_csv(folder_path)
        print(f"Latest CSV file: {latest_csv_file}")

        # First check if file is empty
        if os.path.getsize(latest_csv_file) == 0:
            try:
                outlook = win32.Dispatch('outlook.application')
                mail = outlook.CreateItem(0)
                
                current_date = datetime.now().strftime("%d-%b-%Y")
                mail.Subject = f"High Severity Tickets || Handover - {current_date}"
                
                html_body = f"""
                <html>
                <body style="font-family: Calibri;">
                    <p>Hello Team,</p>
                    <p>Currently we do not have any High Severity Tickets in our queue.</p>
                    <br>
                    <p>Best regards,<br>
                    {sender_email}</p>
                </body>
                </html>
                """
                
                mail.HTMLBody = html_body
                mail.Display()
                return
            except Exception as e:
                print(f"Error creating empty data email: {str(e)}")
                messagebox.showerror("Error", str(e))
                return

        # Try to read the CSV file
        try:
            df = pd.read_csv(latest_csv_file)
        except pd.errors.EmptyDataError:
            outlook = win32.Dispatch('outlook.application')
            mail = outlook.CreateItem(0)
            
            current_date = datetime.now().strftime("%d-%b-%Y")
            mail.Subject = f"High Severity Tickets || Handover - {current_date}"
            
            html_body = f"""
            <html>
            <body style="font-family: Calibri;">
                <p>Hello Team,</p>
                <p>Currently we do not have any High Severity Tickets in our queue.</p>
                <br>
                <p>Best regards,<br>
                {sender_email}</p>
            </body>
            </html>
            """
            
            mail.HTMLBody = html_body
            mail.Display()
            return

        current_date = datetime.now().strftime("%d-%b-%Y")

        # Check if DataFrame is empty
        if df.empty or len(df.index) == 0:
            outlook = win32.Dispatch('outlook.application')
            mail = outlook.CreateItem(0)
            
            mail.Subject = f"High Severity Tickets || Handover - {current_date}"
            
            html_body = f"""
            <html>
            <body style="font-family: Calibri;">
                <p>Hello Team,</p>
                <p>Currently we do not have any High Severity Tickets in our queue.</p>
                <br>
                <p>Best regards,<br>
                {sender_email}</p>
            </body>
            </html>
            """
            
            mail.HTMLBody = html_body
            mail.Display()
            return

        print("\nAll data:")
        print(df)  # This will print all data

        total_rows = len(df)
        
        # Add row selection dialog
        choice = messagebox.askquestion("Row Selection", "Do you want to include all rows?")
        if choice == 'yes':
            rows_to_include = list(range(total_rows))
        else:
            row_input = tk.simpledialog.askstring("Input", f"Enter row numbers to include (1-{total_rows}, separated by commas):")
            try:
                rows_to_include = [int(x.strip())-1 for x in row_input.split(',')]
                if not all(0 <= x < total_rows for x in rows_to_include):
                    messagebox.showerror("Error", "Invalid row numbers")
                    return
            except:
                messagebox.showerror("Error", "Invalid input")
                return

        csv_headers_to_extract = ['TicketLink', 'CreateDate', 'Partner', 'Status','Age']
        custom_headers_for_email = ['Ticket Link', 'Created Date', 'Partner', 'Ticket Status','Ticket Age']

        missing_columns = [col for col in csv_headers_to_extract if col not in df.columns]
        if missing_columns:
            print(f"The following required columns are missing in the CSV file: {missing_columns}")
            return

        # Filter DataFrame to include only selected rows
        df = df.iloc[rows_to_include]

        df_filtered = df[csv_headers_to_extract]
        df_with_additional_columns = create_additional_columns(df_filtered)
        desired_column_order = ['TicketLink', 'CreateDate','Ticket Age', 'Age', 'Partner', 'Root Cause Summary', 
                              'Action Summary', 'Recent Update from Team', 'Next Steps from Team', 'Status']
        df_rearranged = df_with_additional_columns[desired_column_order]
        table_headers = ['Ticket Link', 'Created Date', 'Ticket Age', 'Ticket Age', 'Partner/POM', 'Root Cause Summary', 
                        'Action Summary', 'Recent Update by Team', 'Next Steps from Team', 'Ticket Status']
        table_rows = df_rearranged.values.tolist()

        create_outlook_email(
            recipient="",
            subject=f"High Severity Tickets || Handover - {current_date}",
            table_headers=table_headers,
            table_rows=table_rows
        )
    except Exception as e:
        print(f"An error occurred: {str(e)}")
        messagebox.showerror("Error", str(e))


def sla_24hrs_flow():
    try:
        file_path = get_latest_csv(os.path.join(os.path.expanduser('~'), 'Documents', 'handover'))
        
        # First check if file is empty
        if os.path.getsize(file_path) == 0:
            try:
                outlook = win32.Dispatch('outlook.application')
                mail = outlook.CreateItem(0)
                
                current_date = datetime.now().strftime("%d-%b-%Y")
                mail.Subject = f"High Severity Tickets || Handover - {current_date}"
                
                html_body = f"""
                <html>
                <body style="font-family: Calibri;">
                    <p>Hello Team,</p>
                    <p>Currently we do not have any High Severity Tickets in our queue.</p>
                    <br>
                    <p>Best regards,<br>
                    {sender_email}</p>
                </body>
                </html>
                """
                
                mail.HTMLBody = html_body
                mail.Display()
                return
            except Exception as e:
                print(f"Error creating empty data email: {str(e)}")
                messagebox.showerror("Error", str(e))
                return

        # Try to read the file
        try:
            file_ext = os.path.splitext(file_path)[1].lower()
            if file_ext == '.csv':
                df = pd.read_csv(file_path)
            else:
                df = pd.read_excel(file_path)
        except pd.errors.EmptyDataError:
            outlook = win32.Dispatch('outlook.application')
            mail = outlook.CreateItem(0)
            
            current_date = datetime.now().strftime("%d-%b-%Y")
            mail.Subject = f"High Severity Tickets || Handover - {current_date}"
            
            html_body = f"""
            <html>
            <body style="font-family: Calibri;">
                <p>Hello Team,</p>
                <p>Currently we do not have any High Severity tickets in our queue.</p>
                <br>
                <p>Best regards,<br>
                {sender_email}</p>
            </body>
            </html>
            """
            
            mail.HTMLBody = html_body
            mail.Display()
            return

        print("\nAll data:")
        print(df)  # This will print all data instead of just the first few rows
        
        total_rows = len(df)
        
        custom_mapping = {
            'a': 'ShortId',
            'a_link': 'TicketLink',
            'b': 'Example 2',
            'c': 'Example 3',
            'j': 'CreateDate',
            'k': 'AssignedGroup',
            'l': 'Age',
            'm': 'Example 13',
            'n': 'Example 14',
            'o': 'Example 15',
            'p': 'LastAssignedDate',
            'q': 'Status',
            'r': 'Example 18',
            's': 'VendorId',
            't': 'ShipOrigin',
            'u': 'Example 21',
            'v': 'Example 22',
            'w': 'Example 23',
            'x': 'Example 24',
            'y': 'PhysicalLocation',
            'f': 'Example 6',
            'g': 'Example 7',
            'h': 'Example 8',
            'i': 'Example 9'
        }

        choice = messagebox.askquestion("Row Selection", "Do you want to include all rows?")
        if choice == 'yes':
            rows_to_include = list(range(total_rows))
        else:
            row_input = tk.simpledialog.askstring("Input", f"Enter row numbers to include (1-{total_rows}, separated by commas):")
            try:
                rows_to_include = [int(x.strip())-1 for x in row_input.split(',')]
                if not all(0 <= x < total_rows for x in rows_to_include):
                    messagebox.showerror("Error", "Invalid row numbers")
                    return
            except:
                messagebox.showerror("Error", "Invalid input")
                return

        css_style = """
        <style>
        .custom-table {
            border-collapse: collapse;
            width: 150px;
            font-family: Calibri;
            table-layout: Auto;
            margin-bottom: 20px;
            text-align: center;
        }
        .custom-table td {
            border: 1px solid black;
            padding: 5px;
            text-align: left;
            overflow: hidden;
            text-overflow: ellipsis;
            white-space: nowrap;
        }
        .first-column {
            width: 15%;
            background-color: #F2F2F2;
            text-align: center;
        }
        .yellow-cell {
            background-color: yellow;
        }
        .header-row {
            background-color: #F2F2F2;
            font-weight: bold;
            text-align: center;
        }
        a {
            color: #0066cc;
            text-decoration: none;
        }
        a:hover {
            text-decoration: underline;
        }
        </style>
        """

        current_date = datetime.now().strftime("%d-%b-%Y")
        
        print("\nCreating email...")
        all_tables = []
        for row_index in rows_to_include:
            html_table = create_html_table_with_data(df, custom_mapping, row_index)
            if html_table:
                all_tables.append(f'{html_table}<div style="height: 50px;"></div>')
            else:
                raise Exception(f"Failed to create table for row {row_index + 1}")

        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.Subject = f"High Severity Tickets || Handover - {current_date}"
        
        mail.HTMLBody = f"""
        <html>
        <head>
        {css_style}
        </head>
        <body>
        <p style="font-family: Calibri;">Hello Team,</p>
        <p style="font-family: Calibri;">Please find below the list of open High Severity tickets in our queue:</p>
        {''.join(all_tables)}
        <p style="font-family: Calibri;">Best regards,<br>{sender_email}</p>
        </body>
        </html>
        """
        
        mail.Display()
        
    except Exception as e:
        print(f"An error occurred: {str(e)}")
        messagebox.showerror("Error", str(e))


def check_and_export():
    global export_done
    if not export_done:
        click_button()
        time.sleep(5)  # Wait for download to complete
        export_done = True

def run_normal_flow():
    check_and_export()
    normal_flow()

def run_sla_flow():
    check_and_export()
    sla_24hrs_flow()

def main():
    global status_label, root
    root = tk.Tk()
    root.title("High Severity tickets Handover")
    root.geometry("400x300")
    root.configure(bg="#f0f0f0")  # Light gray background

    # Custom fonts
    title_font = ("Helvetica", 16, "bold")
    button_font = ("Helvetica", 12)

    # Title with custom styling
    title_label = tk.Label(root, 
                          text="High Severity tickets Handover", 
                          font=title_font, 
                          bg="#f0f0f0", 
                          fg="#333333")
    title_label.pack(pady=20)

    # Frame for buttons with custom background
    button_frame = tk.Frame(root, bg="#f0f0f0")
    button_frame.pack(expand=True)

    # Custom button styling
    normal_btn = tk.Button(button_frame, 
                          text="Normal", 
                          command=run_normal_flow,
                          font=button_font,
                          bg="#4CAF50",
                          fg="white",
                          width=20,
                          relief=tk.RAISED,
                          cursor="hand2")
    normal_btn.pack(pady=10)

    # Bind hover effects
    normal_btn.bind("<Enter>", lambda e: e.widget.configure(bg="#45a049"))
    normal_btn.bind("<Leave>", lambda e: e.widget.configure(bg="#4CAF50"))

    sla_btn = tk.Button(button_frame, 
                        text="24hrs SLA", 
                        command=run_sla_flow,
                        font=button_font,
                        bg="#4CAF50",
                        fg="white",
                        width=20,
                        relief=tk.RAISED,
                        cursor="hand2")
    sla_btn.pack(pady=10)

    # Bind hover effects
    sla_btn.bind("<Enter>", lambda e: e.widget.configure(bg="#45a049"))
    sla_btn.bind("<Leave>", lambda e: e.widget.configure(bg="#4CAF50"))

    # Status label
    status_label = tk.Label(root, 
                           text="Ready", 
                           bg="#f0f0f0", 
                           fg="#666666",
                           font=("Helvetica", 10))
    status_label.pack(pady=20)

    # Center the window on the screen
    root.update_idletasks()
    width = root.winfo_width()
    height = root.winfo_height()
    x = (root.winfo_screenwidth() // 2) - (width // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    root.geometry(f'{width}x{height}+{x}+{y}')

    root.mainloop()

if __name__ == "__main__":
    main()
