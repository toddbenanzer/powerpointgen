# powerpointgen

## PPTX Wrapper Module

This module provides a simple wrapper around the `python-pptx` library to simplify creating and manipulating PowerPoint presentations.

### Installation

To install the necessary dependencies, run:
```bash
pip install -r requirements.txt
```

### Basic Usage

Here's how to create a new presentation, add a slide, and save it:

```python
from pyppt import PresentationWrapper

# Create a new presentation
preso_wrapper = PresentationWrapper()

# Add a slide (uses default 'Title and Content' layout)
preso_wrapper.add_slide()

# Add another slide using a specific layout index (e.g., 0 for 'Title Slide')
# Note: Available layout indices can vary depending on the PowerPoint version
# and default template. Index 5 is typically 'Title and Content'.
# You might need to inspect your available layouts if using a custom template.
preso_wrapper.add_slide(layout_index=0)

# Save the presentation
preso_wrapper.save("my_presentation.pptx")

print("Presentation created and saved as my_presentation.pptx")
```

To open an existing presentation, you can pass the file path to the constructor:

```python
from pyppt import PresentationWrapper

# Open an existing presentation
# (Assuming you have a presentation named 'existing_presentation.pptx')
# preso_wrapper = PresentationWrapper("existing_presentation.pptx")

# Add a new slide
# preso_wrapper.add_slide()

# Save changes (either to a new file or overwrite)
# preso_wrapper.save("existing_presentation_modified.pptx")
```