from pptx import Presentation
from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.util import Inches

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
