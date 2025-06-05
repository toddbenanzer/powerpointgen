from pptx import Presentation

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
