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
        self.template_path_var = tk.StringVar(value=self.doc_generator.default_template_path)

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

        input_frame.columnconfigure(1, weight=1)

        # Buttons: definen el contenedor de los botones
        button_frame = ttk.Frame(main_container)
        button_frame.pack(fill="x", pady=20, padx=5)

        buttons_container = ttk.Frame(button_frame)
        buttons_container.pack(pady=10)

        ttk.Button(
            buttons_container,
            text="Preview and Generate Document",
            command=self._process_document
        ).pack(side=tk.LEFT, padx=5)

        # Add a more prominent Generate Document button
        generate_button = ttk.Button(
            button_frame,
            text="Generate Document",
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

    def _process_document(self):
        """Process the spreadsheet and show preview"""
        if not self.sheets_reader:
            messagebox.showwarning(
                "Warning",
                "Please select and validate credentials first"
            )
            return

        identifier = self.identifier_var.get() # Este es el valor de identificador que controla el dato entero del entry
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
            #TODO-> hace la primera lectura y devuelve un df, esto es lo primero que hay que arreglar para que use el templete
            self.current_data = self.sheets_reader.read_sheets(access_type, identifier)

            if not self.current_data:
                messagebox.showwarning("Warning", "No data found in the spreadsheet")
                return

            self.status_var.set("Spreadsheet processed successfully")
            self._show_preview()
        except Exception as e:
            self.status_var.set(f"Error: {str(e)}")
            messagebox.showerror("Error", f"An error occurred: {e}")

    def _show_preview(self):
        """Show a preview of the spreadsheet data before generating the document"""
        preview_window = tk.Toplevel(self.root)
        preview_window.title("Preview")
        preview_window.geometry("800x600")
        preview_window.transient(self.root)

        # Create a notebook (tabbed interface)
        notebook = ttk.Notebook(preview_window)
        notebook.pack(fill="both", expand=True, padx=10, pady=10)

        # Create a tab for each worksheet
        # TODO -> aqui recibe el db y lo analiza por eso no usa el templete en el preview

        frame = ttk.Frame(notebook)
        notebook.add(frame, text="Data Preview")

        # Create a treeview to display the data
        file_path = filedialog.asksaveasfilename(
            defaultextension=".docx",
            filetypes=[("Word Document", "*.docx")]
        )
        if file_path:
            try:
                self.status_var.set("Generating document...")
                title = self.title_var.get() if self.title_var.get() else None
                self.doc_generator.generate_from_dataframes_title_page(self.current_data, file_path, title)
                self.status_var.set("Document generated successfully")
                messagebox.showinfo("Success", "Document generated successfully")
            except Exception as e:
                self.status_var.set(f"Error generating document: {str(e)}")
                messagebox.showerror("Error", f"Failed to generate document: {e}")

        # Add scrollbars
        vsb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        vsb.pack(side="right", fill="y")
        tree.configure(yscrollcommand=vsb.set)

        hsb = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
        hsb.pack(side="bottom", fill="x")
        tree.configure(xscrollcommand=hsb.set)

        # Add a button to generate the document
        def on_generate():
            preview_window.destroy()
            self._save_document()

        ttk.Button(
            preview_window,
            text="Generate Document",
            command=on_generate
        ).pack(pady=10)

    def _save_document(self):
        """Save the document file"""
        # TODO-> Esta es la otra parte que hay que arreglar para generar el documento completo a partir del tmeplete y luego guardarlo
        file_path = filedialog.asksaveasfilename(
            defaultextension=".docx",
            filetypes=[("Word Document", "*.docx")]
        )
        if file_path:
            try:
                self.status_var.set("Generating document...")
                title = self.title_var.get() if self.title_var.get() else None
                self.doc_generator.generate_from_dataframes_title_page(self.current_data, file_path, title)
                self.status_var.set("Document generated successfully")
                messagebox.showinfo("Success", "Document generated successfully")
            except Exception as e:
                self.status_var.set(f"Error generating document: {str(e)}")
                messagebox.showerror("Error", f"Failed to generate document: {e}")

    def _generate_document_directly(self):
        """Process spreadsheet and generate document directly without preview"""
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

            if not self.current_data:
                messagebox.showwarning("Warning", "No data found in the spreadsheet")
                return

            self.status_var.set("Spreadsheet processed successfully")
            # Skip preview and directly save document
            self._save_document()
        except Exception as e:
            self.status_var.set(f"Error: {str(e)}")
            messagebox.showerror("Error", f"An error occurred: {e}")

    def _update_label(self):
        btn_selected = self.access_var.get()
        self.identifier_label.config(text=f"Spreadsheet by {btn_selected}: " )