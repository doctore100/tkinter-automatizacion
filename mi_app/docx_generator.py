import os

import pandas as pd
from docxtpl import DocxTemplate

from mi_app.utils import get_default_template_path


class DocumentGenerator:
    """
    Class for generating documents from templates using data from Google Sheets.

    This class provides functionality to generate Word documents from templates using
    data from Google Sheets. It supports custom templates with a default fallback and
    uses the DocxTemplate library for document generation.

    The generated documents will contain the data from the Google Sheets in a structured
    format based on the template. The template can include variables for title, date,
    and sheet data (headers and rows).
    """

    def __init__(self, template_path=None):
        """
        Initialize the DocumentGenerator with an optional template path.

        Args:
            template_path: Optional path to a custom template file. If not provided,
                          the default template will be used.
        """
        self.default_template_path = get_default_template_path()
        self.template_path = template_path if template_path else self.default_template_path

    def set_template(self, template_path):
        """
        Set a custom template to use for document generation.

        Args:
            template_path: Path to the template file

        Returns:
            True if template was set successfully, False otherwise
        """
        if not os.path.exists(template_path):
            return False
        self.template_path = template_path
        return True

    def reset_to_default_template(self):
        """Reset to the default template"""
        self.template_path = self.default_template_path

    def generate_from_dataframes(self, dataframes, output_path, title=None):
        """
        Generate a document from dataframes using a template.

        Args:
            dataframes: List of tuples containing (worksheet_name, dataframe)
            output_path: Path to save the output document
            title: Optional title for the document

        Returns:
            Path to the generated document
        """
        # Load the template
        try:
            doc = DocxTemplate(self.template_path)
        except Exception as e:
            raise ValueError(f"Failed to load template: {str(e)}")

        # Definir un mapeo de campos a posiciones (fila, columna)
        executive_summary = {
            'author': (5, 42),
            'review': (6, 42),
            'release': (7, 42),
            'version': (3, 42),
            'date': (9, 42),
            'state': (8, 42)
        }
        page_header = {
            'code': (2, 42),
            # 'revision':"",
            'f_emission': (4, 42)
        }
        # The base filter is the first two columns of the data frame that contain the Nivel Jerárquico and the Puesto
        # by which the filter will then be made
        base_filter = dataframes.iloc[10:, 2:3]
        another_data = dataframes.iloc[11:, 4:]

        # Crear el diccionario usando comprensión de diccionarios
        context = {
            field: '' if pd.isna(dataframes.iloc[row, col]) else dataframes.iloc[row, col]
            for field, (row, col) in executive_summary.items()
        }
        doc.render(context)

        # Save the document
        doc.save(output_path)
        return doc.save('index.docx')

        # return output_path
