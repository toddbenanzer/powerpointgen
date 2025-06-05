from pptx import Presentation
from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.util import Inches

class PresentationWrapper:
    def __init__(self, pptx_path=None):
        if pptx_path:
            self.presentation = Presentation(pptx_path)
        else:
            self.presentation = Presentation()
        # Store the path if provided, though it's less relevant if creating new
        self.pptx_path = pptx_path

    def add_slide(self, layout_index=5): # Assuming 5 is 'Title and Content' or a common layout
        """Adds a slide to the presentation.
        layout_index: The index of the slide layout to use.
                      Defaults to 5, often 'Title and Content'.
        """
        slide_layout = self.presentation.slide_layouts[layout_index]
        self.presentation.slides.add_slide(slide_layout)

    def save(self, filename):
        """Saves the presentation to the given filename."""
        self.presentation.save(filename)

    def set_title(self, slide_index, text):
        """Sets the title of the slide at the given index.

        Args:
            slide_index (int): The index of the slide to modify.
            text (str): The text to set as the title.

        Raises:
            IndexError: If slide_index is out of range.
            AttributeError: If the slide does not have a title placeholder.
        """
        if slide_index < 0 or slide_index >= len(self.presentation.slides):
            raise IndexError(f"Slide index {slide_index} is out of range.")

        slide = self.presentation.slides[slide_index]

        if slide.shapes.title:
            slide.shapes.title.text = text
        else:
            # Or log a warning, or try to add a title shape if appropriate
            # For now, raising an error if no title placeholder is clear.
            raise AttributeError(f"Slide {slide_index} does not have a title placeholder.")

    def set_subtitle(self, slide_index, text):
        """Sets the subtitle of the slide at the given index.

        Args:
            slide_index (int): The index of the slide to modify.
            text (str): The text to set as the subtitle.

        Raises:
            IndexError: If slide_index is out of range.
            AttributeError: If the slide does not have a suitable subtitle placeholder.
        """
        if slide_index < 0 or slide_index >= len(self.presentation.slides):
            raise IndexError(f"Slide index {slide_index} is out of range.")

        slide = self.presentation.slides[slide_index]

        # Attempt to find the subtitle placeholder
        subtitle_shape = None
        for shape in slide.placeholders:
            if shape.placeholder_format.type == PP_PLACEHOLDER.SUBTITLE:
                subtitle_shape = shape
                break

        if subtitle_shape:
            subtitle_shape.text = text
        else:
            # Check if there's a shape named "Subtitle" or similar as a fallback,
            # or if there's a placeholder at index 1 (common for title slides)
            # For now, we'll stick to the typed placeholder.
            raise AttributeError(f"Slide {slide_index} does not have a clear subtitle placeholder.")

    def set_footer_text(self, slide_index, text):
        """Sets the footer text on a specific slide.

        Args:
            slide_index (int): The index of the slide to modify.
            text (str): The text to set in the footer.

        Raises:
            IndexError: If slide_index is out of range.
            AttributeError: If the slide does not have a footer placeholder.
        """
        if slide_index < 0 or slide_index >= len(self.presentation.slides):
            raise IndexError(f"Slide index {slide_index} is out of range.")

        slide = self.presentation.slides[slide_index]

        footer_shape = None
        for shape in slide.placeholders:
            if shape.placeholder_format.type == PP_PLACEHOLDER.FOOTER:
                footer_shape = shape
                break

        if footer_shape:
            footer_shape.text = text
        else:
            raise AttributeError(f"Slide {slide_index} does not have a footer placeholder.")

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
        for i, slide in enumerate(self.presentation.slides):
            slide_number_shape = None
            # Check shapes that are placeholders first
            for shape in slide.placeholders:
                if shape.placeholder_format.type == PP_PLACEHOLDER.SLIDE_NUMBER:
                    slide_number_shape = shape
                    break
            # If not found in placeholders, check all shapes (less common for true slide numbers)
            if not slide_number_shape:
                for shape in slide.shapes:
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

    def add_text_box(self, slide_index, text, left, top, width, height):
        """Adds a text box with plain text to the specified slide.
        Positions and dimensions are in Inches.

        Args:
            slide_index (int): The index of the slide.
            text (str): The plain text to add to the text box.
            left (float): The left position of the text box (Inches).
            top (float): The top position of the text box (Inches).
            width (float): The width of the text box (Inches).
            height (float): The height of the text box (Inches).
        Raises:
            IndexError: If slide_index is out of range.
        """
        if slide_index < 0 or slide_index >= len(self.presentation.slides):
            raise IndexError(f"Slide index {slide_index} is out of range.")

        slide = self.presentation.slides[slide_index]

        shape = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
        shape.text_frame.text = text
        return shape

    def add_bullet_point_box(self, slide_index, items, left, top, width, height):
        """Adds a text box with a list of items as bullet points.
        Positions and dimensions are in Inches.

        Args:
            slide_index (int): The index of the slide.
            items (list[str]): A list of strings, each will be a bullet point.
            left (float): The left position (in Inches).
            top (float): The top position (in Inches).
            width (float): The width (in Inches).
            height (float): The height (in Inches).
        Raises:
            IndexError: If slide_index is out of range.
        """
        if slide_index < 0 or slide_index >= len(self.presentation.slides):
            raise IndexError(f"Slide index {slide_index} is out of range.")

        slide = self.presentation.slides[slide_index]

        left_emu = Inches(left)
        top_emu = Inches(top)
        width_emu = Inches(width)
        height_emu = Inches(height)

        shape = slide.shapes.add_textbox(left_emu, top_emu, width_emu, height_emu)
        tf = shape.text_frame
        tf.clear() # Clear existing paragraph (if any)

        for item_text in items:
            p = tf.add_paragraph()
            p.text = item_text
            p.level = 0 # Top-level bullet

        return shape
