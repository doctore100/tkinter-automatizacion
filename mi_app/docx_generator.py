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

    def __init__(self,) -> None:
        """Initialize the DocumentGenerator with an optional template path.

        Args:
            template_path: Optional path to a custom template file. If not provided,
                          the default template will be used.
        """
        self.default_template_path = get_default_template_path()
        #Todo desactivar la funcionalidad de tomar un templete path
        self.template_path = self.default_template_path

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

    def generate_from_dataframes_title_page(
            self,
            dataframes: pd.DataFrame,
            output_path: str,
            job_title: str,
            level_hierarchy: str,
            title: Optional[str] = None,

    ) -> None:
        """Generate a Word document from template using dataframe data.

        Processes the data, maps fields to template locations, and saves the document.

        Args:
            dataframes: DataFrame containing input data with specific columns/rows
            output_path: File path to save the generated document
            title: Optional document title
            job_title: Optional job title to filter data
            level_hierarchy: Optional level hierarchy to filter data

        Raises:
            ValueError: If template loading fails
        """
        try:
            doc = DocxTemplate(self.template_path)
            print(f'este es el path del templete: {self.template_path}')
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
        print(f"Field position mapping: {field_position_mapping}")
        ##### clean_data devuelve lo correcto
        doc.render(clean_data(field_position_mapping, dataframes))

        # Process data and generate document
        df_data_general = self._process_data(dataframes)
        print(f"Data to generate PDF para procesar el resto del dato: {df_data_general}")

        # Only process job-specific data if both job_title and level_hierarchy are provided
        # Check if job_title and level_hierarchy are strings and not empty
        if isinstance(job_title, str) and job_title.strip() and isinstance(level_hierarchy, str) and level_hierarchy.strip():
            data_to_generate_pdf = self._process_general_data(df_data_general, job_title, level_hierarchy)
            # Add the job-specific data to the document context
            doc.render(data_to_generate_pdf)

        doc.save(output_path)

    def _process_data(self, dataframes: pd.DataFrame) -> pd.DataFrame:
        """Process and combine dataframe sections.

        Args:
            dataframes: Input dataframe containing raw data

        Returns:
            pd.DataFrame: Processed dataframe
        """
        base_filter = (
            dataframes.iloc[10:, 2:4]
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

    def _process_general_data(self, dataframes: pd.DataFrame, job_title: str, level_hierarchy: str) -> dict:
        # print(f'job_title: {job_title} y level_hierarchy {level_hierarchy} \n el dataframe {dataframes} dentro de _process_general_data')
        """Process and combine dataframe sections.

        Searches for a row where the first column contains level_hierarchy and
        the second column contains job_title, then returns the complete row
        with keys mapped to template placeholders.
        This version handles whitespace, case variations, and missing values robustly.

        Args:
            dataframes: Input dataframe containing the data
            job_title: Value to search for in the second column
            level_hierarchy: Value to search for in the first column

        Returns:
            dict: Dictionary with keys matching template placeholders

        Raises:
            ValueError: If no matching row is found or multiple rows match
        """
        # Get the first two columns
        first_col = dataframes.iloc[:, 0]
        second_col = dataframes.iloc[:, 1]
        print(f'first_col: {dataframes.iloc[:, 0].tolist()} \n y second_col: {dataframes.iloc[:, 1].tolist()} dentro de _process_general_data')

        # Convert to string, handle NaN values, strip whitespace, and convert to lowercase for comparison
        first_col_clean = first_col.astype(str).str.strip().str.lower()
        second_col_clean = second_col.astype(str).str.strip().str.lower()

        # Clean the search parameters
        job_title_clean = str(job_title).strip().lower()
        level_hierarchy_clean = str(level_hierarchy).strip().lower()

        # Filter out 'nan' values that come from converting NaN to string


        # Create boolean mask searching for the specific values
        mask = (first_col_clean == level_hierarchy_clean) & (second_col_clean == job_title_clean)
        print(f'paso 1 de la mascara :{first_col_clean == level_hierarchy_clean}')
        print(f'paso 2 de la mascara :{second_col_clean == job_title_clean}')
        print(f"Mask: {mask}")

        # Filter the dataframe
        matching_rows = dataframes[mask]

        # Check if exactly one row matches
        if matching_rows.empty:
            raise ValueError(
                f"No row found with '{level_hierarchy}' in first column and '{job_title}' in second column. "
                f"Available values in first column: {first_col.dropna().unique().tolist()[:10]}... "
                f"Available values in second column: {second_col.dropna().unique().tolist()[:10]}..."
            )
        elif len(matching_rows) > 1:
            raise ValueError(
                f"Multiple rows found with '{level_hierarchy}' in first column and '{job_title}' in second column. "
                f"Found {len(matching_rows)} matching rows at indices: {matching_rows.index.tolist()}"
            )

        # Get the matching row
        row = matching_rows.iloc[0]

        # Map the row values to template placeholders
        # Based on the template structure and the datasheet columns
        template_data = {
            'puesto': row.iloc[1] if len(row) > 1 else '',  # Job title
            'n_jerarquico': row.iloc[0] if len(row) > 0 else '',  # Hierarchical level
            'a_trabajo': row.iloc[2] if len(row) > 2 else '',  # Work area
            'p_participa': row.iloc[3] if len(row) > 3 else '',  # Processes
            'is_supervisado': row.iloc[4] if len(row) > 4 else '',  # Supervised by
            'supervisa_to': row.iloc[5] if len(row) > 5 else '',  # Supervises
            'replace_to': row.iloc[6] if len(row) > 6 else '',  # Replaces
            'is_replace': row.iloc[7] if len(row) > 7 else '',  # Is replaced by
            'objective_position': row.iloc[8] if len(row) > 8 else '',  # Job objective
            'responsibilities': row.iloc[9] if len(row) > 9 else '',  # Responsibilities
            'specific_responsibilities': row.iloc[10] if len(row) > 10 else '',  # Specific functions
            'sgi_specific': row.iloc[11] if len(row) > 11 else '',  # SGI functions
            'specific_functions': row.iloc[12] if len(row) > 12 else '',  # RASCI matrix
            'educations': row.iloc[13] if len(row) > 13 else '',  # Education
            'work_experience': row.iloc[14] if len(row) > 14 else '',  # Work experience
            'proactivity': row.iloc[15] if len(row) > 15 else '',  # Proactivity
            'oral_expression': row.iloc[16] if len(row) > 16 else '',  # Oral expression
            'teamwork': row.iloc[17] if len(row) > 17 else '',  # Teamwork
            'digital_tools': row.iloc[18] if len(row) > 18 else '',  # Digital tools
            't_quality_control': row.iloc[19] if len(row) > 19 else '',  # Quality control
            'num_geom_skills': row.iloc[20] if len(row) > 20 else '',  # Numerical skills
            'project_management': row.iloc[21] if len(row) > 21 else '',  # Project management
            'troubleshooting': row.iloc[22] if len(row) > 22 else '',  # Troubleshooting
            'change_management': row.iloc[23] if len(row) > 23 else '',  # Change management
            'innovation_creativity': row.iloc[24] if len(row) > 24 else '',  # Innovation
            'business_skills': row.iloc[25] if len(row) > 25 else '',  # Business skills
            'textile_techniques': row.iloc[26] if len(row) > 26 else ''  # Textile techniques
        }

        return template_data
