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
        Generates a Word document based on a given template, using data from specified dataframes
        to populate the document's context. This method processes the data, maps relevant fields
        to their locations in the document template, dynamically generates the content, and saves
        the resulting Word document to the specified output path.

        Data from the input dataframes is used to create an executive summary and populate specific
        page headers in the template. The function assumes that certain rows and columns have specific
        meanings and uses them to extract, filter, and process required data. After processing the
        data, a context dictionary is created and rendered into the template, followed by saving the
        generated document.

        The method ensures consistency between the base filter and additional data, aligning and
        concatenating them as necessary for accurate processing. It requires a pre-defined template
        available at the specified path to render the final document.

        :param dataframes: A pandas DataFrame containing input data; assumes specific
            columns and rows contain required fields such as "Nivel Jerárquico" and "Puesto".
        :type dataframes: pandas.DataFrame
        :param output_path: File path where the generated Word document will be saved.
        :type output_path: str
        :param title: Optional parameter to specify the title of the document.
        :type title: str, optional
        :return: A successful save operation for the document at the specified location.
        :rtype: None
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
        base_filter = (
            dataframes.iloc[10:, 2:3]
            .replace('', pd.NA)
            .replace(' ', pd.NA)
            .dropna()
        )
        # The other data contains the rest of the fields in the data frame.
        another_data = dataframes.iloc[11:, 4:]
        if base_filter.shape[0] == another_data.shape[0]:
            df_final = pd.DataFrame({
                **base_filter.reset_index(drop=True).to_dict('list'),
                **another_data.reset_index(drop=True).to_dict('list')
            })
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
