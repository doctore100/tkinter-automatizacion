import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import gspread
from google.oauth2 import service_account
import pandas as pd
from fpdf import FPDF
import os
import json
from gspread.utils import GridRangeType


# Try to import Google Docs API modules, but make them optional
GOOGLE_DOCS_AVAILABLE = True
try:
    from googleapiclient.discovery import build
except ImportError:
    GOOGLE_DOCS_AVAILABLE = False


class GoogleConnection:
    """Class for handling Google API connections and credential validation"""

    def __init__(self, credentials_path=None, ):
        self.credentials_path = credentials_path
        self.client = None
        self.scope = [
            "https://spreadsheets.google.com/feeds",
            "https://www.googleapis.com/auth/drive",
            "https://www.googleapis.com/auth/documents"
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
            # Try to load the credentials file
            with open(self.credentials_path, 'r') as f:
                json_data = json.load(f)

            # Check if it has the required fields for a service account
            required_fields = ['client_email', 'private_key', 'project_id']
            for field in required_fields:
                if field not in json_data:
                    return False, f"Invalid credentials: missing {field}"

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


class GoogleDocumentReader:
    """Class for reading data from Google Sheets and Google Docs"""

    def __init__(self, google_connection):
        """
        Initializes a class instance by setting up the Google connection and preparing
        the client attribute for later assignment.

        :param google_connection: A connection instance to interact with Google services.
        :type google_connection: Any
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
        return df

    def read_document(self, doc_id):
        """
        Read content from a Google Doc

        Args:
            doc_id: The document ID

        Returns:
            Document content as text
        """
        # Check if Google Docs API is available
        if not GOOGLE_DOCS_AVAILABLE:
            raise ValueError(
                "Google Docs functionality is not available. Please install the 'google-api-python-client' package."
            )

        # For Google Docs, we need to use the Google Docs API
        # This requires additional setup with the googleapiclient
        try:
            # Create credentials object for Docs API
            creds = service_account.Credentials.from_service_account_file(
                self.connection.credentials_path,
                scopes=['https://www.googleapis.com/auth/documents.readonly']
            )

            # Build the Docs API client
            docs_service = build('docs', 'v1', credentials=creds)

            # Retrieve the document
            document = docs_service.documents().get(documentId=doc_id).execute()

            # Extract text from the document
            doc_content = document.get('body').get('content')
            text_content = self._extract_text_from_doc_content(doc_content)

            # Return as a single "worksheet" with the document content
            df = pd.DataFrame({'Content': [text_content]})
            return [('Document', df)]

        except Exception as e:
            raise ValueError(f"Failed to read Google Doc: {str(e)}")

    def _extract_text_from_doc_content(self, content):
        """Helper method to extract text from Google Doc content"""
        text = []
        for element in content:
            if 'paragraph' in element:
                for paragraph_element in element['paragraph']['elements']:
                    if 'textRun' in paragraph_element:
                        text.append(paragraph_element['textRun']['content'])
            elif 'table' in element:
                # Handle tables
                for table_row in element['table']['tableRows']:
                    row_text = []
                    for cell in table_row['tableCells']:
                        if 'content' in cell:
                            cell_text = self._extract_text_from_doc_content(cell['content'])
                            row_text.append(cell_text)
                    text.append(' | '.join(row_text))
            elif 'tableOfContents' in element:
                text.append('[Table of Contents]')
        return '\n'.join(text)


class PDFGenerator:
    """Class for generating and formatting PDFs from Google documents"""

    def __init__(self):
        self.pdf = None
#TODO: Reformat the function for generate a document templete
    def generate_from_dataframes(self, dataframes, output_path, title=None):
        """
        Generate a PDF file from a list of dataframes

        Args:
            dataframes: List of tuples containing (worksheet_name, dataframe)
            output_path: Path to save the PDF file
            title: Optional title for the PDF
        """
        self.pdf = FPDF()
        self.pdf.set_auto_page_break(auto=True, margin=15)

        # Add title page if title is provided
        if title:
            self._add_title_page(title)

        for sheet_name, df in dataframes:
            self._add_dataframe_page(sheet_name, df)

        self.pdf.output(output_path)
        return output_path

    def _add_title_page(self, title):
        """Add a title page to the PDF"""
        self.pdf.add_page()
        self.pdf.set_font("Arial", "B", 24)

        # Calculate position for centered text
        title_w = self.pdf.get_string_width(title)
        self.pdf.set_xy((self.pdf.w - title_w) / 2, self.pdf.h / 3)

        # Add title
        self.pdf.cell(title_w, 10, title, ln=True, align="C")

        # Add date
        self.pdf.set_font("Arial", "I", 12)
        import datetime
        date_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        date_w = self.pdf.get_string_width(date_str)
        self.pdf.set_xy((self.pdf.w - date_w) / 2, self.pdf.h / 2)
        self.pdf.cell(date_w, 10, date_str, ln=True, align="C")

    def _add_dataframe_page(self, sheet_name, df):
        """Add a page with dataframe content"""
        self.pdf.add_page()
        self.pdf.set_font("Arial", "B", 12)
        self.pdf.cell(200, 10, f"Sheet: {sheet_name}", ln=True, align="C")
        self.pdf.ln(10)

        # Check if this is a document content dataframe (special case)
        if list(df.columns) == ['Content'] and len(df) == 1:
            self._add_document_content(df.iloc[0]['Content'])
            return

        # Regular dataframe handling
        self.pdf.set_font("Arial", "", 10)

        # Calculate column width based on number of columns
        col_width = self.pdf.w / (len(df.columns) + 1)

        # Write headers
        for col in df.columns:
            self.pdf.cell(col_width, 10, str(col), border=1)
        self.pdf.ln()

        # Write rows
        for _, row in df.iterrows():
            for item in row:
                # Truncate long text to fit in cell
                text = str(item)
                if len(text) > 30:  # Arbitrary limit to prevent overflow
                    text = text[:27] + "..."
                self.pdf.cell(col_width, 10, text, border=1)
            self.pdf.ln()

    def _add_document_content(self, content):
        """Add document content to the PDF"""
        self.pdf.set_font("Arial", "", 10)

        # Split content into lines and add them to the PDF
        lines = content.split('\n')
        for line in lines:
            # Check if line is too long and needs to be wrapped
            if self.pdf.get_string_width(line) > self.pdf.w - 20:
                words = line.split()
                current_line = ""
                for word in words:
                    test_line = current_line + " " + word if current_line else word
                    if self.pdf.get_string_width(test_line) < self.pdf.w - 20:
                        current_line = test_line
                    else:
                        self.pdf.multi_cell(0, 10, current_line)
                        current_line = word
                if current_line:
                    self.pdf.multi_cell(0, 10, current_line)
            else:
                self.pdf.multi_cell(0, 10, line)


class GoogleToPDFApp:
    """Class to control the Tkinter UI/UX for the Google to PDF converter application"""

    def __init__(self, root):
        self.root = root
        self.root.title("Google to PDF Converter")
        self.root.geometry("700x500")
        self.root.configure(padx=20, pady=20)

        # Initialize components
        self.connection = GoogleConnection()
        self.document_reader = None
        self.pdf_generator = PDFGenerator()
        self.current_data = []

        # UI variables
        self.access_var = tk.StringVar(value="url")
        self.doc_type_var = tk.StringVar(value="sheets")
        self.identifier_var = tk.StringVar()
        self.status_var = tk.StringVar(value="Ready")
        self.title_var = tk.StringVar()

        # Setup UI
        self._setup_styles()
        self._create_widgets()

    def _setup_styles(self):
        """Setup ttk styles"""
        style = ttk.Style()
        style.configure("TFrame", padding=10)
        style.configure("TButton", padding=6)
        style.configure("TRadiobutton", padding=5)
        style.configure("Header.TLabel", font=("Arial", 16, "bold"))
        style.configure("Subheader.TLabel", font=("Arial", 10))
        style.configure("Status.TLabel", padding=2, relief="sunken")
        # Add a special style for the Generate PDF button
        style.configure("Generate.TButton", padding=10, font=("Arial", 12, "bold"))

    def _create_widgets(self):
        """Create all UI widgets"""
        # Main container
        main_container = ttk.Frame(self.root)
        main_container.pack(fill="both", expand=True)

        # Header
        header_frame = ttk.Frame(main_container)
        header_frame.pack(fill="x", pady=10)

        ttk.Label(
            header_frame,
            text="Google to PDF Converter",
            style="Header.TLabel"
        ).pack()

        ttk.Label(
            header_frame,
            text="Convert Google Sheets or Google Docs to PDF format",
            style="Subheader.TLabel"
        ).pack(pady=5)

        # Credentials frame
        creds_frame = ttk.LabelFrame(main_container, text="Google API Credentials")
        creds_frame.pack(fill="x", pady=10, padx=5)

        ttk.Button(
            creds_frame,
            text="Select Credentials File",
            command=self._select_credentials
        ).pack(side="left", padx=5, pady=10)

        self.creds_label = ttk.Label(creds_frame, text="No credentials selected")
        self.creds_label.pack(side="left", padx=5, fill="x", expand=True)

        ttk.Button(
            creds_frame,
            text="Validate Credentials",
            command=self._validate_credentials
        ).pack(side="right", padx=5, pady=10)

        # Document type selection
        doc_type_frame = ttk.LabelFrame(main_container, text="Document Type")
        doc_type_frame.pack(fill="x", pady=10, padx=5)

        ttk.Radiobutton(
            doc_type_frame,
            text="Google Sheets",
            variable=self.doc_type_var,
            value="sheets",
            command=self._update_input_label
        ).pack(side="left", padx=20, pady=10)

        # Create Google Docs radio button
        docs_radio = ttk.Radiobutton(
            doc_type_frame,
            text="Google Docs",
            variable=self.doc_type_var,
            value="docs",
            command=self._update_input_label
        )
        docs_radio.pack(side="left", padx=20, pady=10)

        # Disable Google Docs option if the API is not available
        if not GOOGLE_DOCS_AVAILABLE:
            docs_radio.configure(state="disabled")
            docs_warning = ttk.Label(
                doc_type_frame,
                text="(Install google-api-python-client to enable)",
                foreground="red",
                font=("Arial", 8)
            )
            docs_warning.pack(side="left", padx=5, pady=10)

        # Access method (for Sheets only)
        self.access_frame = ttk.LabelFrame(main_container, text="Access Method (Sheets Only)")
        self.access_frame.pack(fill="x", pady=10, padx=5)

        ttk.Radiobutton(
            self.access_frame,
            text="By Name",
            variable=self.access_var,
            value="name"
        ).pack(side="left", padx=10, pady=10)

        ttk.Radiobutton(
            self.access_frame,
            text="By Key",
            variable=self.access_var,
            value="key"
        ).pack(side="left", padx=10, pady=10)

        ttk.Radiobutton(
            self.access_frame,
            text="By URL",
            variable=self.access_var,
            value="url"
        ).pack(side="left", padx=10, pady=10)

        # Input fields
        input_frame = ttk.Frame(main_container)
        input_frame.pack(fill="x", pady=10, padx=5)

        self.identifier_label = ttk.Label(input_frame, text="Spreadsheet URL:")
        self.identifier_label.grid(row=0, column=0, sticky="w", pady=5)

        ttk.Entry(
            input_frame,
            textvariable=self.identifier_var,
            width=60
        ).grid(row=0, column=1, sticky="ew", pady=5, padx=5)

        ttk.Label(input_frame, text="PDF Title (optional):").grid(row=1, column=0, sticky="w", pady=5)
        ttk.Entry(
            input_frame,
            textvariable=self.title_var,
            width=60
        ).grid(row=1, column=1, sticky="ew", pady=5, padx=5)

        input_frame.columnconfigure(1, weight=1)

        # Buttons
        button_frame = ttk.Frame(main_container)
        button_frame.pack(fill="x", pady=20, padx=5)

        buttons_container = ttk.Frame(button_frame)
        buttons_container.pack(pady=10)

        ttk.Button(
            buttons_container,
            text="Preview and Generate PDF",
            command=self._process_document
        ).pack(side=tk.LEFT, padx=5)

        # Add a more prominent Generate PDF button
        generate_button = ttk.Button(
            button_frame,
            text="Generate PDF",
            command=self._generate_pdf_directly,
            style="Generate.TButton"
        )
        generate_button.pack(pady=10)

        # Status bar
        status_bar = ttk.Label(
            self.root,
            textvariable=self.status_var,
            style="Status.TLabel",
            anchor=tk.W
        )
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)

    def _update_input_label(self):
        """Update the input label based on document type"""
        doc_type = self.doc_type_var.get()
        if doc_type == "sheets":
            self.identifier_label.config(text="Spreadsheet URL:")
            self.access_frame.pack(fill="x", pady=10, padx=5)
        else:
            self.identifier_label.config(text="Document ID:")
            self.access_frame.pack_forget()

    def _select_credentials(self):
        """Handle credentials selection"""
        if self.connection.select_credentials():
            filename = os.path.basename(self.connection.credentials_path)
            self.creds_label.config(text=f"Selected: {filename}")
            self.status_var.set(f"Credentials file selected: {filename}")
        else:
            self.status_var.set("Credentials selection cancelled")

    def _validate_credentials(self):
        """Validate the selected credentials"""
        if not self.connection.credentials_path:
            messagebox.showerror("Error", "Please select credentials file first")
            return

        if self.connection.show_validation_window(self.root):
            self.document_reader = GoogleDocumentReader(self.connection)
            self.status_var.set("Credentials validated successfully")
        else:
            self.status_var.set("Credentials validation failed")

    def _process_document(self):
        """Process the document and show preview"""
        if not self.document_reader:
            messagebox.showwarning(
                "Warning",
                "Please select and validate credentials first"
            )
            return

        identifier = self.identifier_var.get()
        if not identifier:
            messagebox.showwarning(
                "Warning",
                "Please enter the document identifier"
            )
            return

        try:
            self.status_var.set("Processing document...")

            # Get data based on document type
            doc_type = self.doc_type_var.get()
            if doc_type == "sheets":
                access_type = self.access_var.get()
                self.current_data = self.document_reader.read_sheets(access_type, identifier)
            else:
                # Check if Google Docs API is available before proceeding
                if not GOOGLE_DOCS_AVAILABLE:
                    messagebox.showerror(
                        "Error",
                        "Google Docs functionality is not available. Please install the 'google-api-python-client' package."
                    )
                    self.status_var.set("Error: Google Docs API not available")
                    return
                self.current_data = self.document_reader.read_document(identifier)

            if not self.current_data:
                messagebox.showwarning("Warning", "No data found in the document")
                return

            self.status_var.set("Document processed successfully")
            self._show_preview()
        except Exception as e:
            self.status_var.set(f"Error: {str(e)}")
            messagebox.showerror("Error", f"An error occurred: {e}")

    def _show_preview(self):
        """Show a preview of the data before generating the PDF"""
        preview_window = tk.Toplevel(self.root)
        preview_window.title("Preview")
        preview_window.geometry("800x600")
        preview_window.transient(self.root)

        # Create a notebook (tabbed interface)
        notebook = ttk.Notebook(preview_window)
        notebook.pack(fill="both", expand=True, padx=10, pady=10)

        # Create a tab for each worksheet
        for sheet_name, df in self.current_data:
            frame = ttk.Frame(notebook)
            notebook.add(frame, text=sheet_name)

            # Create a treeview to display the data
            tree = ttk.Treeview(frame)
            tree.pack(fill="both", expand=True)

            # Define columns
            tree["columns"] = list(df.columns)
            tree["show"] = "headings"

            # Set column headings
            for col in df.columns:
                tree.heading(col, text=col)
                tree.column(col, width=100)

            # Add data rows
            for i, row in df.iterrows():
                tree.insert("", "end", values=list(row))

            # Add scrollbars
            vsb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
            vsb.pack(side="right", fill="y")
            tree.configure(yscrollcommand=vsb.set)

            hsb = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
            hsb.pack(side="bottom", fill="x")
            tree.configure(xscrollcommand=hsb.set)

        # Add a button to generate the PDF
        def on_generate():
            preview_window.destroy()
            self._save_pdf()

        ttk.Button(
            preview_window,
            text="Generate PDF",
            command=on_generate
        ).pack(pady=10)

    def _save_pdf(self):
        """Save the PDF file"""
        file_path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")]
        )
        if file_path:
            try:
                self.status_var.set("Generating PDF...")
                title = self.title_var.get() if self.title_var.get() else None
                self.pdf_generator.generate_from_dataframes(self.current_data, file_path, title)
                self.status_var.set("PDF generated successfully")
                messagebox.showinfo("Success", "PDF generated successfully")
            except Exception as e:
                self.status_var.set(f"Error generating PDF: {str(e)}")
                messagebox.showerror("Error", f"Failed to generate PDF: {e}")

    def _generate_pdf_directly(self):
        """Process document and generate PDF directly without preview"""
        if not self.document_reader:
            messagebox.showwarning(
                "Warning",
                "Please select and validate credentials first"
            )
            return

        identifier = self.identifier_var.get()
        if not identifier:
            messagebox.showwarning(
                "Warning",
                "Please enter the document identifier"
            )
            return

        try:
            self.status_var.set("Processing document...")

            # Get data based on document type
            doc_type = self.doc_type_var.get()
            if doc_type == "sheets":
                access_type = self.access_var.get()
                self.current_data = self.document_reader.read_sheets(access_type, identifier)
            else:
                # Check if Google Docs API is available before proceeding
                if not GOOGLE_DOCS_AVAILABLE:
                    messagebox.showerror(
                        "Error",
                        "Google Docs functionality is not available. Please install the 'google-api-python-client' package."
                    )
                    self.status_var.set("Error: Google Docs API not available")
                    return
                self.current_data = self.document_reader.read_document(identifier)

            if not self.current_data:
                messagebox.showwarning("Warning", "No data found in the document")
                return

            self.status_var.set("Document processed successfully")
            # Skip preview and directly save PDF
            self._save_pdf()
        except Exception as e:
            self.status_var.set(f"Error: {str(e)}")
            messagebox.showerror("Error", f"An error occurred: {e}")


# Main entry point
def main():
    root = tk.Tk()
    app = GoogleToPDFApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
