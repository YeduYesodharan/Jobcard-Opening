import requests
from datetime import datetime
import pandas as pd
import pygetwindow as gw
from RPA.Windows import Windows
import json
import time
import pyautogui
import win32gui
import win32con




# def send_post_request(url, data_dict):
#     """
#     Sends a POST request to the given URL with the provided dictionary as the JSON body.

#     :param url: The endpoint URL where the request should be sent.
#     :param data_dict: The dictionary to be sent as the JSON body of the POST request.
#     :return: Response object containing the result of the request.
#     """
#     # Ensure the dictionary is converted to a JSON string
#     payload = json.dumps(data_dict)
    
#     # Set the headers for the POST request
#     headers = {
#         'Content-Type': 'application/json',
#     }
    
#     # Send the POST request with the JSON payload
#     response = requests.post(url, data=payload, headers=headers)
    
#     # Return the response for further handling
#     return response

def send_post_request(url, data_dict):
    """
    Sends a POST request to the given URL with the provided dictionary as the JSON body.

    :param url: The endpoint URL where the request should be sent.
    :param data_dict: The dictionary to be sent as the JSON body of the POST request.
    :return: Response object containing the result of the request.
    """
    # Convert dictionary to JSON string
    payload = json.dumps(data_dict)

    # Set headers for JSON content
    headers = {
        'Content-Type': 'application/json',
    }

    try:
        # Send POST request
        response = requests.post(url, data=payload, headers=headers)
        return response
    except requests.exceptions.RequestException as e:
        print(f"Request error: {e}")
        raise


# def bot_run_status_save_to_db(url, consolidated_report, region_mapping_sheet, jc_no, date_timestamp):
    
#     try:

#         # Load Excel files
#         df = pd.read_excel(consolidated_report, dtype=str)
#         region_df = pd.read_excel(region_mapping_sheet, dtype=str)

#         # Strip spaces from column names
#         df.columns = df.columns.str.strip()
#         region_df.columns = region_df.columns.str.strip()

#         # Create a mapping dictionary from Location Mapping DMS ERP
#         region_dict = dict(zip(region_df["DMS Location description"], region_df["ERP location code"]))

#         row = None
#         # Iterate through each row, and if a match row found, then copy to row variable.
#         for _, each_row in df.iterrows():
#             if each_row['Job Card No'] == jc_no:
#                 row = each_row
#                 break

#         # # Iterate through each row
#         # for _, row in df.iterrows():
#         job_card_no = row["Job Card No"]
#         branch = row["Branch"]
#         service_type_code = row["Service Type Code"]
#         service_type_desc = row["Service Type Description"]
#         recall_status = row["Recall Status"]
#         exception_reason= row["Exception Reason"]
        

#         # Fetch Region value (ERP location code) if Branch exists in mapping
#         region = region_dict.get(branch, "Unknown")

#         # Check Execution Status
#         execution_status = "Success" if row["DMS Execution Status"] == "Success" and row["ERP Execution Status"] == "Success" else "Fail"

#         # Create dictionary entry
#         data_dict = {
#             "Job Card No": job_card_no,
#             "Region": region,
#             "Branch": branch,
#             "Service Type Code": service_type_code,
#             "Service Type Description": service_type_desc,
#             "Recall Status": recall_status,
#             "Date and Time": date_timestamp,
#             "Execution Status": execution_status,
#             "Exception Reason": exception_reason
#         }

#         response = send_post_request(url, data_dict)

#         #Check the response
#         if response.status_code == 200:
#             print("Request was successful.")
#             print("Response:", response.json())  # Display response data
#         else:
#             print(f"Request failed with status code {response.status_code}")
#             print("Response:", response.text)

#     except Exception as e:
#         print(f"Error processing files: {e}")
#         return []

def bot_run_status_save_to_db(url, consolidated_report, region_mapping_sheet, jc_no, date_timestamp):
    try:
        # Load Excel files
        df = pd.read_excel(consolidated_report, dtype=str)
        region_df = pd.read_excel(region_mapping_sheet, dtype=str)

        # Strip whitespace from column names
        df.columns = df.columns.str.strip()
        region_df.columns = region_df.columns.str.strip()

        # Create mapping dictionary for location
        region_dict = dict(zip(region_df["DMS Location description"], region_df["ERP location code"]))

        # Find the matching row based on Job Card No
        match = df.loc[df['Job Card No'] == jc_no]
        if match.empty:
            print(f"No matching job card found for: {jc_no}")
            return []

        row = match.iloc[0]

        # Extract required fields
        job_card_no = row["Job Card No"]
        branch = row["Branch"]
        service_type_code = row["Service Type Code"]
        service_type_desc = row["Service Type Description"]
        recall_status = row["Recall Status"]
        exception_reason = row["Exception Reason"]

        # Determine region from mapping
        region = region_dict.get(branch, "Unknown")

        # Determine execution status
        execution_status = "Success" if row["DMS Execution Status"] == "Success" and row["ERP Execution Status"] == "Success" else "Fail"

        # Create data dictionary
        data_dict = {
            "Job Card No": job_card_no,
            "Region": region,
            "Branch": branch,
            "Service Type Code": service_type_code,
            "Service Type Description": service_type_desc,
            "Recall Status": recall_status,
            "Date and Time": date_timestamp,
            "Execution Status": execution_status,
            "Exception Reason": exception_reason
        }

        # Send data to API
        response = send_post_request(url, data_dict)

        # Handle response
        if response.status_code == 200:
            print("Request was successful.")
            print("Response:", response.json())
        else:
            print(f"Request failed with status code {response.status_code}")
            print("Response:", response.text)

        return  'success'

    except Exception as e:
        print(f"Error processing files: {e}")
        return []
    

    

# date_timestamp = datetime.now().strftime("%Y-%m-%dT%H:%M:%S")
# url = 'http://rpa.popularmaruti.com/store-jobcards-creation-data'
# consolidated_report = r"D:\JobcardOpeningIntegrated_Api\ProcessRelatedFolders\2025-04-27\ProcessedERP\consolidated_jobcard_report20250427233553.xlsx"
# region_mapping_sheet = r"D:\JobcardOpeningIntegrated_Api\Mapping\Location Mapping DMS ERP.xlsx"
# jc_no = r"JC25001458"
# bot_run_status_save_to_db(url, consolidated_report, region_mapping_sheet, jc_no, date_timestamp)

# def get_window_title():
#     windows = gw.getAllWindows()

#     # Print titles of all visible windows
#     for window in windows:
#         if window.title:  # Only print non-empty titles
#             print(window.title)

# get_window_title()


# def bring_window_to_front(partial_title):
#     windows = Windows()
#     pyautogui.press('alt')  # Helps allow the foreground change

#     success = False

#     # Find all windows matching the partial title
#     matching_windows = gw.getWindowsWithTitle(partial_title)

#     for window in matching_windows:
#         full_title = window.title

#         try:
#             print(f"Attempting to bring window to front: {full_title}")

#             # Restore the window if minimized
#             if window.isMinimized:
#                 window.restore()

#             # Activate using pygetwindow
#             window.activate()
#             time.sleep(0.2)

#             # Try activating with RPA.Windows (no 'action' or 'title' argument)
#             try:
#                 windows.control_window(locator=full_title)
#             except Exception as rpa_error:
#                 print(f"RPA.Windows activation failed: {rpa_error}")

#             # Last resort: use win32gui
#             hwnd = window._hWnd
            
#             # Maximize the window
#             win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)  # Maximize the window
#             print(f"Successfully brought window to front and maximized: {full_title}")
#             success = True
#             break

#         except Exception as e:
#             print(f"Could not activate window '{full_title}': {e}")
#             continue

#     if not success:
#         raise Exception(f"Window with partial title '{partial_title}' not found or could not be activated.")


def bring_window_to_front(partial_title):
    windows = Windows()
    pyautogui.press('alt')  # Helps allow foreground change

    start_time = time.time()
    success = False

    matching_windows = gw.getWindowsWithTitle(partial_title)

    for window in matching_windows:
        full_title = window.title
        try:
            print(f"Attempting to bring window to front: {full_title}")

            if window.isMinimized:
                window.restore()

            # Try with pygetwindow first (usually fastest)
            window.activate()
            time.sleep(0.1)

            # Use win32gui to maximize and bring to front
            hwnd = window._hWnd
            win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
            win32gui.SetForegroundWindow(hwnd)
            print(f"Successfully brought window to front and maximized: {full_title}")
            success = True
            break

        except Exception as e:
            print(f"Primary methods failed for '{full_title}', trying RPA.Windows: {e}")

            try:
                windows.control_window(locator=full_title)
                print(f"Activated with RPA.Windows: {full_title}")
                success = True
                break
            except Exception as rpa_error:
                print(f"RPA.Windows also failed: {rpa_error}")
                continue

    if not success:
        raise Exception(f"Window with partial title '{partial_title}' not found or could not be activated.")

    print(f"Time taken to bring window to front: {round(time.time() - start_time, 2)} seconds")
