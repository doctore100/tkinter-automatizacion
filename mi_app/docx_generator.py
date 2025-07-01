import os
from typing import Optional, Dict, Tuple

import pandas as pd
from docxtpl import DocxTemplate

from mi_app.utils import get_default_template_path, clean_data


class DocumentGenerator:
    """Generates documents from templates using data from Google Sheets.

    This class provides functionality to generate Word documents from templates using
    data from Google Sheets. It supports custom templates with a default fallback and
    uses the DocxTemplate library for document generation.
    """

    def __init__(self, template_path: Optional[str] = None) -> None:
        """Initialize the DocumentGenerator with an optional template path.

        Args:
            template_path: Optional path to a custom template file. If not provided,
                          the default template will be used.
        """
        self.default_template_path = get_default_template_path()
        self.template_path = template_path or self.default_template_path

    def set_template(self, template_path: str) -> bool:
        """Set a custom template for document generation.

        Args:
            template_path: Path to the template file

        Returns:
            bool: True if template was set successfully, False otherwise
        """
        if not os.path.exists(template_path):
            return False
        self.template_path = template_path
        return True

    def reset_to_default_template(self) -> None:
        """Reset to using the default template."""
        self.template_path = self.default_template_path

    def generate_from_dataframes(
            self,
            dataframes: pd.DataFrame,
            output_path: str,
            title: Optional[str] = None
    ) -> None:
        """Generate a Word document from template using dataframe data.

        Processes the data, maps fields to template locations, and saves the document.

        Args:
            dataframes: DataFrame containing input data with specific columns/rows
            output_path: File path to save the generated document
            title: Optional document title

        Raises:
            ValueError: If template loading fails
        """
        try:
            doc = DocxTemplate(self.template_path)
        except Exception as e:
            raise ValueError(f"Failed to load template: {str(e)}")

        # Define field position mappings
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
            'f_emission': (4, 42)
        }

        field_position_mapping = {**executive_summary, **page_header}
        doc.render(clean_data(field_position_mapping))

        # Process data and generate document
        self._process_data(dataframes)
        doc.save(output_path)

    def _process_data(self, dataframes: pd.DataFrame) -> pd.DataFrame:
        """Process and combine dataframe sections.

        Args:
            dataframes: Input dataframe containing raw data

        Returns:
            pd.DataFrame: Processed dataframe
        """
        base_filter = (
            dataframes.iloc[10:, 2:3]
            .replace('', pd.NA)
            .replace(' ', pd.NA)
            .dropna()
        )

        another_data = dataframes.iloc[11:, 4:]

        if base_filter.shape[0] == another_data.shape[0]:
            return pd.DataFrame({
                **base_filter.reset_index(drop=True).to_dict('list'),
                **another_data.reset_index(drop=True).to_dict('list')
            })
        raise ValueError("DataFrames have mismatched lengths after processing")