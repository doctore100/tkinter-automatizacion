import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import pandas as pd

from mi_app.google_sheets import GoogleConnection, GoogleSheetsReader
from mi_app.docx_generator import DocumentGenerator


class GoogleToDocApp:
    """
    Class to control the Tkinter UI/UX for the Google Sheets to Document converter application.

    This class provides a graphical user interface for converting Google Sheets data to
    formatted Word documents using templates. It allows users to:

    1. Select and validate Google API credentials
    2. Choose a Google Sheets document by name, key, or URL
    3. Select a custom template or use the default template
    4. Preview the data before generating the document
    5. Generate a Word document with the data formatted according to the template

    The application uses the GoogleSheetsReader class to read data from Google Sheets
    and the DocumentGenerator class to generate documents from templates.
    """

    def __init__(self, root):
        self.root = root
        self.root.title("Conversor de Google Sheets para Documento Word")
        self.root.geometry("700x700")
        self.root.configure(padx=20, pady=20, bg="lightblue")

        # Initialize components
        self.connection = GoogleConnection()
        self.sheets_reader = None
        self.doc_generator = DocumentGenerator()
        self.current_data = []

        # UI variables
        self.access_var = tk.StringVar(value="url")
        self.identifier_var = tk.StringVar()
        self.status_var = tk.StringVar(value="Ready")
        self.title_var = tk.StringVar()
        self.job_title_var = tk.StringVar()
        self.level_hierarchy_var = tk.StringVar()
        self.template_path_var = tk.StringVar(value=self.doc_generator.default_template_path)

        # Lists for dropdown values
        self.job_titles = []
        self.level_hierarchies = []

        # Setup UI
        self._setup_styles()
        self._configure_ui_layout()

    def _setup_styles(self):
        """Setup custom styles for the UI"""
        style = ttk.Style()
        style.configure("TFrame", padding=10)
        style.configure("TButton", padding=6)
        style.configure("TRadiobutton", padding=5)
        style.configure("", font=("Arial", 16, "bold"))
        style.configure("Sub.TLabel", font=("Arial", 10))
        style.configure("Status.TLabel", padding=2, relief="sunken")
        # Add a special style for the Generate PDF button
        style.configure("Generate.TButton", padding=10, font=("Arial", 12, "bold"))

    def _configure_ui_layout(self):
        """Create all UI widgets"""
        # Main container
        main_container = ttk.Frame(self.root)
        main_container.pack(fill="both", expand=True)

        # Header
        header_frame = ttk.Frame(main_container)
        header_frame.pack(fill="x", pady=10)

        ttk.Label(
            header_frame,
            text="Convertidor de Google Sheets para Documento Word",
            style="Sub.TLabel"
        ).pack(pady=5)

        ttk.Label(
            header_frame,
            text="Conversor de Google Sheets a Documento Word",
            style="Sub.TLabel"
        ).pack(pady=5)

        # Credentials frame for validation
        creds_frame = ttk.LabelFrame(main_container, text="Credenciales de la api de Google")
        creds_frame.pack(fill="x", pady=10, padx=5)
        # Validation de las credenciales de google sheet
        ttk.Button(
            creds_frame,
            text="Select Credentials File",
            command=self._select_credentials # retorna true o false
        ).pack(side="left", padx=5, pady=10)

        self.creds_label = ttk.Label(creds_frame, text="No credentials selected")
        self.creds_label.pack(side="left", padx=5, fill="x", expand=True)

        ttk.Button(
            creds_frame,
            text="Validate Credentials",
            command=self._validate_credentials
        ).pack(side="right", padx=5, pady=10)

        # Template selection frame
        template_frame = ttk.LabelFrame(main_container, text="Templete del documento Word:")
        template_frame.pack(fill="x", pady=10, padx=5)

        ttk.Label(
            template_frame,
            text="Template:"
        ).pack(side="left", padx=5, pady=10)

        template_entry = ttk.Label(
            template_frame,
            textvariable=self.template_path_var,
            width=40
        )
        template_entry.pack(side="left", padx=5, pady=10, fill="x", expand=True)

        # ttk.Button(
        #     template_frame,
        #     text="Browse",
        #     command=self._select_template
        # ).pack(side="left", padx=5, pady=10)
        #
        # ttk.Button(
        #     template_frame,
        #     text="Reset to Default",
        #     command=self._reset_template
        # ).pack(side="left", padx=5, pady=10)

        # Access method for Google Sheets and data manipulations
        self.access_frame = ttk.LabelFrame(main_container, text="Access Method")
        self.access_frame.pack(fill="x", pady=10, padx=5)

        ttk.Radiobutton(
            self.access_frame,
            text="By Name",
            variable=self.access_var,
            value="name",
            command=self._update_label
        ).pack(side="left", padx=10, pady=10)

        ttk.Radiobutton(
            self.access_frame,
            text="By Key",
            variable=self.access_var,
            value="key",
            command=self._update_label
        ).pack(side="left", padx=10, pady=10)

        ttk.Radiobutton(
            self.access_frame,
            text="By URL",
            variable=self.access_var,
            value="url",
            command=self._update_label
        ).pack(side="left", padx=10, pady=10)

        # Input fields
        input_frame = ttk.Frame(main_container)
        input_frame.pack(fill="x", pady=10, padx=5)

        self.identifier_label = ttk.Label(input_frame, text="Spreadsheet by: " )
        self.identifier_label.grid(row=0, column=0, sticky="w", pady=5)
        self._update_label()

        ttk.Entry(
            input_frame,
            textvariable=self.identifier_var,# el contenido entero del entry
            width=60
        ).grid(row=0, column=1, sticky="ew", pady=5, padx=5)

        ttk.Label(input_frame, text="PDF Title (optional):").grid(row=1, column=0, sticky="w", pady=5)
        ttk.Entry(
            input_frame,
            textvariable=self.title_var,
            width=60
        ).grid(row=1, column=1, sticky="ew", pady=5, padx=5)

        # Add job title dropdown
        ttk.Label(input_frame, text="Job Title:").grid(row=2, column=0, sticky="w", pady=5)
        self.job_title_combobox = ttk.Combobox(
            input_frame,
            textvariable=self.job_title_var,
            width=58,
            values=self.job_titles
        )
        self.job_title_combobox.grid(row=2, column=1, sticky="ew", pady=5, padx=5)

        # Add level hierarchy dropdown
        ttk.Label(input_frame, text="Level Hierarchy:").grid(row=3, column=0, sticky="w", pady=5)
        self.level_hierarchy_combobox = ttk.Combobox(
            input_frame,
            textvariable=self.level_hierarchy_var,
            width=58,
            values=self.level_hierarchies
        )
        self.level_hierarchy_combobox.grid(row=3, column=1, sticky="ew", pady=5, padx=5)

        input_frame.columnconfigure(1, weight=1)

        # Buttons: definen el contenedor de los botones
        button_frame = ttk.Frame(main_container)
        button_frame.pack(fill="x", pady=20, padx=5)

        buttons_container = ttk.Frame(button_frame)
        buttons_container.pack(pady=10)

        # Add a more prominent Load Data button
        generate_button = ttk.Button(
            button_frame,
            text="Load Data",
            command=self._generate_document_directly,
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
            self.sheets_reader = GoogleSheetsReader(self.connection)
            self.status_var.set("Credentials validated successfully")
        else:
            self.status_var.set("Credentials validation failed")

    def _select_template(self):
        """Allow user to select a template file"""
        file_path = filedialog.askopenfilename(
            title="Select Template File",
            filetypes=[("Word Templates", "*.docx")]
        )
        if file_path:
            if self.doc_generator.set_template(file_path):
                self.template_path_var.set(file_path)
                self.status_var.set(f"Template selected: {os.path.basename(file_path)}")
            else:
                messagebox.showerror("Error", "Invalid template file")

    def _reset_template(self):
        """Reset to the default template"""
        self.doc_generator.reset_to_default_template()
        self.template_path_var.set(self.doc_generator.default_template_path)
        self.status_var.set("Reset to default template")


    def _save_document(self):
        """Save the document file"""
        file_path = filedialog.asksaveasfilename(
            defaultextension=".docx",
            filetypes=[("Word Document", "*.docx")]
        )
        if file_path:
            try:
                self.status_var.set("Generating document...")

                # Get values from input fields
                title = self.title_var.get() if self.title_var.get() else None
                job_title = self.job_title_var.get() if self.job_title_var.get() else None
                level_hierarchy = self.level_hierarchy_var.get() if self.level_hierarchy_var.get() else None

                # Job title and level hierarchy should already be validated before calling this method

                # Generate document with all available parameters
                self.doc_generator.generate_from_dataframes_title_page(
                    self.current_data, 
                    file_path, 
                    title,
                    job_title,
                    level_hierarchy
                )

                self.status_var.set("Document generated successfully")
                messagebox.showinfo("Success", "Document generated successfully")
            except Exception as e:
                self.status_var.set(f"Error generating document: {str(e)}")
                messagebox.showerror("Error", f"Failed to generate document: {e}")

    def _generate_document_directly(self):
        """Process spreadsheet and open selection window"""
        if not self.sheets_reader:
            messagebox.showwarning(
                "Warning",
                "Please select and validate credentials first"
            )
            return

        identifier = self.identifier_var.get()
        if not identifier:
            messagebox.showwarning(
                "Warning",
                "Please enter the spreadsheet identifier"
            )
            return

        try:
            self.status_var.set("Processing spreadsheet...")

            # Get data from Google Sheets
            access_type = self.access_var.get()
            self.current_data = self.sheets_reader.read_sheets(access_type, identifier)

            if self.current_data is None or self.current_data.empty:
                messagebox.showwarning("Warning", "No data found in the spreadsheet")
                return

            # Extract job titles and level hierarchies from the data
            self._extract_job_data_from_dataframe()

            self.status_var.set("Spreadsheet processed successfully")

            # Open selection window
            self._show_selection_window()
        except Exception as e:
            self.status_var.set(f"Error: {str(e)}")
            messagebox.showerror("Error", f"An error occurred: {e}")

    def _show_selection_window(self):
        """Show a window for selecting job title and level hierarchy"""
        selection_window = tk.Toplevel(self.root)
        selection_window.title("Select Categories")
        selection_window.geometry("600x300")
        selection_window.transient(self.root)
        selection_window.grab_set()  # Make the window modal

        # Create a frame for the content
        content_frame = ttk.Frame(selection_window, padding=20)
        content_frame.pack(fill="both", expand=True)

        # Add instructions
        ttk.Label(
            content_frame,
            text="Select Job Title and Level Hierarchy to generate the document:",
            font=("Arial", 12)
        ).pack(pady=(0, 20))

        # Add job title dropdown
        job_frame = ttk.Frame(content_frame)
        job_frame.pack(fill="x", pady=5)

        ttk.Label(job_frame, text="Job Title:", width=15).pack(side="left")
        job_combobox = ttk.Combobox(
            job_frame,
            width=40,
            values=self.job_titles
        )
        job_combobox.pack(side="left", padx=5, fill="x", expand=True)

        # Add level hierarchy dropdown
        level_frame = ttk.Frame(content_frame)
        level_frame.pack(fill="x", pady=5)

        ttk.Label(level_frame, text="Level Hierarchy:", width=15).pack(side="left")
        level_combobox = ttk.Combobox(
            level_frame,
            width=40,
            values=self.level_hierarchies
        )
        level_combobox.pack(side="left", padx=5, fill="x", expand=True)

        # Add a button frame
        button_frame = ttk.Frame(content_frame)
        button_frame.pack(pady=20)

        def on_generate():
            job_title = job_combobox.get()
            level_hierarchy = level_combobox.get()

            if not job_title or not level_hierarchy:
                messagebox.showwarning(
                    "Warning", 
                    "Please select both Job Title and Level Hierarchy before generating the document.",
                    parent=selection_window
                )
                return

            # Set the values in the main window variables
            self.job_title_var.set(job_title)
            self.level_hierarchy_var.set(level_hierarchy)

            # Close the selection window
            selection_window.destroy()

            # Generate the document
            self._save_document()

        ttk.Button(
            button_frame,
            text="Generate Document",
            command=on_generate,
            style="Generate.TButton"
        ).pack(pady=10)

    def _update_label(self):
        btn_selected = self.access_var.get()
        self.identifier_label.config(text=f"Spreadsheet by {btn_selected}: " )

    def _extract_job_data_from_dataframe(self):
        """Extract job titles and level hierarchies from the dataframe"""
        if self.current_data is None or self.current_data.empty:
            return

        try:
            # Process the dataframe to extract job titles and level hierarchies
            # Based on the structure in the dataset, job titles are in column 1 and level hierarchies in column 0
            # Starting from row 11 (index 10) to skip headers
            df_filtered = self.current_data.iloc[10:, :2].copy()

            # Clean the data
            df_filtered = df_filtered.replace('', pd.NA).replace(' ', pd.NA).dropna()

            # Extract unique values
            if df_filtered.shape[1] > 1:
                # Get job titles from column 1
                job_titles = df_filtered.iloc[:, 1].dropna().unique().tolist()
                self.job_titles = [str(title).strip() for title in job_titles if str(title).strip()]

                # Get level hierarchies from column 0
                level_hierarchies = df_filtered.iloc[:, 0].dropna().unique().tolist()
                self.level_hierarchies = [str(level).strip() for level in level_hierarchies if str(level).strip()]

                # Update the comboboxes
                if hasattr(self, 'job_title_combobox') and self.job_titles:
                    self.job_title_combobox['values'] = self.job_titles

                if hasattr(self, 'level_hierarchy_combobox') and self.level_hierarchies:
                    self.level_hierarchy_combobox['values'] = self.level_hierarchies
        except Exception as e:
            print(f"Error extracting job data: {e}")
            # Don't raise the exception, just log it
