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

### Text Manipulation

The `PresentationWrapper` class provides methods to add and modify various text elements on your slides.

**Setting Titles and Subtitles**

You can set the title and subtitle for a slide if its layout includes placeholders for them.

```python
# Assuming 'preso_wrapper' is an instance of PresentationWrapper
# and has at least one slide (e.g., added via preso_wrapper.add_slide())
# Set title for the first slide (index 0)
try:
    preso_wrapper.set_title(0, "My Presentation Title")
    preso_wrapper.set_subtitle(0, "A Subtitle for the First Slide")
except IndexError:
    print("Error: Slide index out of range. Add a slide first.")
except AttributeError as e:
    print(f"Error setting title/subtitle: {e}. Ensure slide layout has these placeholders.")
```

**Managing Footers and Slide Numbers**

Set footer text on a specific slide:
```python
# Set footer text for the first slide
try:
    preso_wrapper.set_footer_text(0, "Confidential - Company Use Only")
except IndexError:
    print("Error: Slide index out of range.")
except AttributeError as e:
    print(f"Error setting footer: {e}. Ensure slide layout has a footer placeholder.")
```

Control slide number visibility (attempts to hide/show based on placeholder text):
```python
# Attempt to make slide numbers visible
preso_wrapper.set_slide_numbers_visibility(True)
# Attempt to hide slide numbers
# preso_wrapper.set_slide_numbers_visibility(False)
```
*Note: Slide number visibility is heavily dependent on the slide master and layout configurations.*

**Adding Text Boxes**

Add a simple text box to a slide at a specified position (in inches):
```python
# Add a text box to the first slide
try:
    preso_wrapper.add_text_box(0, "This is some important text.", left=1, top=2, width=4, height=1)
except IndexError:
    print("Error: Slide index out of range.")
```

**Adding Bullet Points**

Add a text box with bullet points:
```python
# Add bullet points to the first slide
try:
    items = ["Bullet point 1", "Another bullet point", "Final bullet"]
    preso_wrapper.add_bullet_point_box(0, items, left=1, top=3.5, width=5, height=2)
except IndexError:
    print("Error: Slide index out of range.")
```