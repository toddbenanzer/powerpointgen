from pptx import Presentation
from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.util import Inches
import pandas as pd

class PySlide:
    def __init__(self, pptx_slide):
        """Initializes the PySlide.

        Args:
            pptx_slide (pptx.slide.Slide): The python-pptx Slide object.
        """
        self.pptx_slide = pptx_slide

    def set_title(self, text):
        """Sets the title of this slide.

        Args:
            text (str): The text to set as the title.

        Raises:
            AttributeError: If the slide does not have a title placeholder.
        """
        if self.pptx_slide.shapes.title:
            self.pptx_slide.shapes.title.text = text
        else:
            raise AttributeError("This slide does not have a title placeholder.")

    def set_subtitle(self, text):
        """Sets the subtitle of this slide.

        Args:
            text (str): The text to set as the subtitle.

        Raises:
            AttributeError: If the slide does not have a suitable subtitle placeholder.
        """
        subtitle_shape = None
        for shape in self.pptx_slide.placeholders:
            if shape.placeholder_format.type == PP_PLACEHOLDER.SUBTITLE:
                subtitle_shape = shape
                break

        if subtitle_shape:
            subtitle_shape.text = text
        else:
            raise AttributeError("This slide does not have a clear subtitle placeholder.")

    def set_footer_text(self, text):
        """Sets the footer text on this slide.

        Args:
            text (str): The text to set in the footer.

        Raises:
            AttributeError: If the slide does not have a footer placeholder.
        """
        footer_shape = None
        for shape in self.pptx_slide.placeholders:
            if shape.placeholder_format.type == PP_PLACEHOLDER.FOOTER:
                footer_shape = shape
                break

        if footer_shape:
            footer_shape.text = text
        else:
            raise AttributeError("This slide does not have a footer placeholder.")

    def add_text_box(self, text, left, top, width, height):
        """Adds a text box with plain text to this slide.
        Positions and dimensions are in Inches.

        Args:
            text (str): The plain text to add to the text box.
            left (float): The left position of the text box (Inches).
            top (float): The top position of the text box (Inches).
            width (float): The width of the text box (Inches).
            height (float): The height of the text box (Inches).
        """
        shape = self.pptx_slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
        shape.text_frame.text = text
        return shape

    def add_bullet_point_box(self, items, left, top, width, height):
        """Adds a text box with a list of items as bullet points to this slide.
        Positions and dimensions are in Inches.

        Args:
            items (list[str]): A list of strings, each will be a bullet point.
            left (float): The left position (in Inches).
            top (float): The top position (in Inches).
            width (float): The width (in Inches).
            height (float): The height (in Inches).
        """
        left_emu = Inches(left)
        top_emu = Inches(top)
        width_emu = Inches(width)
        height_emu = Inches(height)

        shape = self.pptx_slide.shapes.add_textbox(left_emu, top_emu, width_emu, height_emu)
        tf = shape.text_frame
        tf.clear()

        for item_text in items:
            p = tf.add_paragraph()
            p.text = item_text
            p.level = 0

        return shape

    def add_table_from_dataframe(self, dataframe, left, top, width, height,
                                 column_labels=None, number_formats=None,
                                 include_index=False, index_label=None):
        """Adds a table to the slide populated from a Pandas DataFrame.

        Args:
            dataframe (pd.DataFrame): The Pandas DataFrame to display.
            left (float): The left position of the table (in Inches).
            top (float): The top position of the table (in Inches).
            width (float): The width of the table (in Inches).
            height (float): The height of the table (in Inches).
            column_labels (dict, optional): Maps original DataFrame column names
                                          to custom display names for the table header.
                                          e.g., {'col_df_name': 'Display Name'}
            number_formats (dict, optional): Maps original DataFrame column names
                                           to Python format strings for cell values.
                                           e.g., {'price_col': '$,.2f', 'qty_col': ',d'}
            include_index (bool): If True, include the DataFrame's index as the
                                  first column. Defaults to False.
            index_label (str, optional): Header for the index column if
                                       include_index is True. Defaults to DataFrame's
                                       index name or "Index".
        Returns:
            pptx.shapes.graphfrm.GraphicFrame: The table shape object.
        """
        if not isinstance(dataframe, pd.DataFrame):
            raise ValueError("Input 'dataframe' must be a pandas DataFrame.")

        rows = len(dataframe) + 1  # +1 for header row
        cols = len(dataframe.columns)
        if include_index:
            cols += 1

        # Create the table shape
        table_shape = self.pptx_slide.shapes.add_table(
            rows, cols, Inches(left), Inches(top), Inches(width), Inches(height)
        )
        table = table_shape.table

        # --- Populate Header Row ---
        current_col_idx = 0
        if include_index:
            header_text = index_label if index_label is not None else (dataframe.index.name if dataframe.index.name is not None else "Index")
            table.cell(0, current_col_idx).text = str(header_text)
            current_col_idx += 1

        for df_col_name in dataframe.columns:
            display_name = str(df_col_name) # Default to original column name
            if column_labels and df_col_name in column_labels:
                display_name = str(column_labels[df_col_name])
            table.cell(0, current_col_idx).text = display_name
            current_col_idx += 1

        # --- Populate Data Rows ---
        for i, df_row_tuple in enumerate(dataframe.itertuples(index=include_index, name=None)):
            # df_row_tuple contains actual data values. If include_index=True, first element is index.

            current_cell_idx_in_row = 0
            data_tuple_offset = 0 # if include_index is False, df_row_tuple starts with first data col

            if include_index:
                index_val = df_row_tuple[0]
                # Apply formatting to index if 'index_label' (or a convention) is in number_formats
                # For now, index is added as string without specific formatting via number_formats
                table.cell(i + 1, current_cell_idx_in_row).text = str(index_val)
                current_cell_idx_in_row += 1
                data_tuple_offset = 1 # Data starts from 2nd element of df_row_tuple

            for col_idx, df_col_name in enumerate(dataframe.columns):
                cell_value = df_row_tuple[col_idx + data_tuple_offset]

                formatted_value = str(cell_value) # Default to string representation
                if pd.isna(cell_value): # Handle NaN/None before formatting
                    formatted_value = "" # Or some other placeholder like "N/A"
                elif number_formats and df_col_name in number_formats:
                    try:
                        fmt_spec = number_formats[df_col_name]
                        # Ensure value is appropriate for format spec (e.g. numeric for 'f', 'd')
                        if isinstance(cell_value, (int, float)):
                             formatted_value = f"{cell_value:{fmt_spec}}"
                        # else: formatted_value remains str(cell_value) if format spec is for numbers but type isn't
                    except ValueError:
                        # print(f"Warning: Could not apply format '{fmt_spec}' to value '{cell_value}' in col '{df_col_name}'. Using default string.")
                        pass # Keep default string if formatting fails

                table.cell(i + 1, current_cell_idx_in_row).text = formatted_value
                current_cell_idx_in_row += 1

        return table_shape

class PyPPT:
    def __init__(self, pptx_path=None):
        if pptx_path:
            self.presentation = Presentation(pptx_path)
        else:
            self.presentation = Presentation()
        # Store the path if provided, though it's less relevant if creating new
        self.pptx_path = pptx_path

    def add_slide(self, layout_index=5):
        """Adds a new slide to the presentation and returns its wrapper.

        Args:
            layout_index (int): The index of the slide layout to use.
                                Defaults to 5 (often 'Title and Content').
        Returns:
            PySlide: A wrapper for the newly added slide.
        """
        slide_layout = self.presentation.slide_layouts[layout_index]
        new_pptx_slide = self.presentation.slides.add_slide(slide_layout)
        return PySlide(new_pptx_slide)

    def get_slide(self, slide_index):
        """Gets the slide at the specified index as a PySlide instance.

        Args:
            slide_index (int): The index of the slide to retrieve.

        Returns:
            PySlide: A wrapper for the specified slide.

        Raises:
            IndexError: If slide_index is out of range.
        """
        if slide_index < 0 or slide_index >= len(self.presentation.slides):
            raise IndexError(f"Slide index {slide_index} is out of range.")
        pptx_slide = self.presentation.slides[slide_index]
        return PySlide(pptx_slide)

    @property
    def slides(self):
        """Returns a list of PySlide instances for all slides in the presentation."""
        return [PySlide(s) for s in self.presentation.slides]

    def save(self, filename):
        """Saves the presentation to the given filename."""
        self.presentation.save(filename)

    def set_slide_numbers_visibility(self, visible=True):
        """Attempts to set visibility of slide numbers on each slide by interacting with placeholders.

        For hiding (visible=False): Clears text from slide number placeholders.
        For showing (visible=True): This method doesn't explicitly force show if master/layout hides it,
                                   but ensures text isn't cleared if a placeholder exists.
                                   Proper enabling is best done via slide master/layout design.
        Args:
            visible (bool): True to attempt to show/ensure not hidden, False to attempt to hide.
        """
        print("INFO: Slide number visibility is best configured in the slide master/layout.")
        for i, slide_obj in enumerate(self.presentation.slides): # Renamed slide to slide_obj to avoid conflict
            slide_number_shape = None
            # Check shapes that are placeholders first
            for shape in slide_obj.placeholders: # Use slide_obj here
                if shape.placeholder_format.type == PP_PLACEHOLDER.SLIDE_NUMBER:
                    slide_number_shape = shape
                    break
            # If not found in placeholders, check all shapes (less common for true slide numbers)
            if not slide_number_shape:
                for shape in slide_obj.shapes: # Use slide_obj here
                    if hasattr(shape, "is_placeholder") and shape.is_placeholder and \
                       hasattr(shape, "placeholder_format") and \
                       shape.placeholder_format.type == PP_PLACEHOLDER.SLIDE_NUMBER:
                        slide_number_shape = shape
                        break

            if slide_number_shape:
                if not visible:
                    slide_number_shape.text_frame.clear() # Clears text, effectively hiding it.
                    print(f"INFO: Cleared text from slide number placeholder on slide {i} to hide it.")
                else:
                    # If it was cleared, this won't automatically restore it.
                    # python-pptx should fill it if the layout/master expects a slide number.
                    # We are not adding text here as it should be automatic.
                    print(f"INFO: For slide {i}, ensure layout/master enables slide numbers for placeholder to fill.")
            elif visible:
                print(f"WARNING: Slide {i} does not seem to have a slide number placeholder to make visible.")
