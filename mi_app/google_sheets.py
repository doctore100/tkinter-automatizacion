import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import gspread
from google.oauth2 import service_account
import pandas as pd
import json
import os
from gspread.utils import GridRangeType
from mi_app.utils import validate_json_file, get_credentials_path


class GoogleConnection:
    """Class for handling Google API connections and credential validation"""

    def __init__(self, credentials_path=None):
        self.credentials_path = credentials_path if credentials_path else get_credentials_path()
        self.client = None
        self.scope = [
            "https://spreadsheets.google.com/feeds",
            "https://www.googleapis.com/auth/drive"
        ]

    def select_credentials(self):
        """Allow user to select credentials file"""
        file_path = filedialog.askopenfilename(
            title="Select Google API Credentials",
            filetypes=[("JSON files", "*.json")]
        )
        if file_path:
            self.credentials_path = file_path
            return True
        return False

    def validate_credentials(self):
        """Validate the selected credentials file"""
        if not self.credentials_path:
            return False, "No credentials file selected"

        try:
            # Check if it has the required fields for a service account
            required_fields = ['client_email', 'private_key', 'project_id']
            is_valid, message = validate_json_file(self.credentials_path, required_fields)
            if not is_valid:
                return False, message

            # Try to connect to Google API using the newer google-auth library
            creds = service_account.Credentials.from_service_account_file(
                self.credentials_path, scopes=self.scope
            )
            self.client = gspread.authorize(creds)

            # Test connection by trying to list all spreadsheets in the drive
            self.client.list_spreadsheet_files()

            return True, "Credentials validated successfully"
        except Exception as e:
            return False, f"Validation failed: {str(e)}"

    def show_validation_window(self, parent):
        """Display a validation window to verify credentials"""
        if not self.credentials_path:
            messagebox.showerror("Error", "Please select credentials file first")
            return False

        validation_window = tk.Toplevel(parent)
        validation_window.title("Credentials Validation")
        validation_window.geometry("400x200")
        validation_window.transient(parent)
        validation_window.grab_set()

        # Center the window
        validation_window.update_idletasks()
        width = validation_window.winfo_width()
        height = validation_window.winfo_height()
        x = (validation_window.winfo_screenwidth() // 2) - (width // 2)
        y = (validation_window.winfo_screenheight() // 2) - (height // 2)
        validation_window.geometry('{}x{}+{}+{}'.format(width, height, x, y))

        # Add a label
        ttk.Label(
            validation_window,
            text="Validating credentials...",
            font=("Arial", 12)
        ).pack(pady=20)

        # Add a progress bar
        progress = ttk.Progressbar(
            validation_window,
            orient="horizontal",
            length=300,
            mode="indeterminate"
        )
        progress.pack(pady=10)
        progress.start()

        result_var = tk.StringVar()
        result_label = ttk.Label(
            validation_window,
            textvariable=result_var,
            font=("Arial", 10)
        )
        result_label.pack(pady=10)

        # Validate in a separate function to avoid freezing the UI
        def validate():
            is_valid, message = self.validate_credentials()
            progress.stop()
            result_var.set(message)

            if is_valid:
                result_label.config(foreground="green")
                validation_window.after(2000, validation_window.destroy)
                return True
            else:
                result_label.config(foreground="red")
                return False

        # Schedule the validation to run after the window is shown
        validation_window.after(100, validate)

        # Wait for the window to be destroyed
        parent.wait_window(validation_window)

        return self.client is not None

    def connect(self):
        """Connect to Google API using validated credentials"""
        if not self.client:
            try:
                creds = service_account.Credentials.from_service_account_file(
                    self.credentials_path, scopes=self.scope
                )
                self.client = gspread.authorize(creds)
            except Exception as e:
                raise ConnectionError(f"Failed to connect: {str(e)}")
        return self.client


class GoogleSheetsReader:
    """
    Class for reading data from Google Sheets.

    This class provides functionality to connect to Google Sheets using the provided
    credentials and read data from spreadsheets. It supports accessing spreadsheets
    by name, key, or URL.
    """

    def __init__(self, google_connection):
        """
        Initializes a class instance by setting up the Google connection and preparing
        the client attribute for later assignment.

        :param google_connection: A connection instance to interact with Google services.
        :type google_connection: GoogleConnection
        """
        self.connection = google_connection
        self.client = None

    def connect(self):
        """
        Connect to the external service and initialize the client if it has not
        already been connected.

        This method ensures that the `client` attribute is initialized with an
        active connection. If the connection is already established, it will not
        attempt to reconnect or reinitialize the client.

        :return: Returns the initialized client object.
        :rtype: Any
        """
        if not self.client:
            self.client = self.connection.connect()
        return self.client

    def read_sheets(self, access_type, identifier):
        """
        Reads data from a Google Sheets spreadsheet and returns it as a pandas DataFrame.
        This function connects to a Google Sheets client and retrieves the data from
        a specified spreadsheet based on the provided access type (name, key, or URL).
        It converts the sheet data into a DataFrame for further processing or analysis.

        :param access_type: A string indicating the type of identifier used to access
            the spreadsheet. Options are "name", "key", or "url".
        :param identifier: A string representing the value associated with the
            specified access type. For example, a spreadsheet name for "name",
            a unique key for "key", or a complete URL for "url".
        :return: A pandas DataFrame containing the data from the spreadsheet.
        """
        client = self.connect()

        if access_type == "name":
            spreadsheet = client.open(identifier)
        elif access_type == "key":
            spreadsheet = client.open_by_key(identifier)
        elif access_type == "url":
            spreadsheet = client.open_by_url(identifier)
        else:
            raise ValueError("Invalid access type")
        spreadsheet_data = spreadsheet.sheet1.get(return_type=GridRangeType.ListOfLists)
        df = pd.DataFrame(spreadsheet_data)
        return [('Sheet1', df)]