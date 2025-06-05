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
from pyppt import PyPPT

# Create a new presentation
preso_wrapper = PyPPT()

# Add a slide (uses default 'Title and Content' layout)
# This now returns a PySlide instance
new_slide = preso_wrapper.add_slide()
new_slide.set_title("Title for New Slide") # Example of using the returned slide

# Add another slide using a specific layout index (e.g., 0 for 'Title Slide')
title_slide_obj = preso_wrapper.add_slide(layout_index=0)
title_slide_obj.set_title("Main Title Slide")
try:
    # Assuming layout 0 (Title Slide) typically has a subtitle placeholder
    title_slide_obj.set_subtitle("My Presentation Subtitle")
except AttributeError as e:
    print(f"Could not set subtitle on title_slide_obj: {e}")


# You can also get existing slides using an index (returns a PySlide)
if len(preso_wrapper.slides) > 0:
    first_slide_wrapper = preso_wrapper.get_slide(0)
    # first_slide_wrapper.set_title("Updated title for first slide") # Already set by 'new_slide.set_title' if it was the first

# Or iterate through all slides
# for idx, slide_wrapper in enumerate(preso_wrapper.slides):
#     slide_wrapper.set_footer_text(f"Slide {idx+1}")

# Save the presentation
preso_wrapper.save("my_presentation.pptx")

print("Presentation created and saved as my_presentation.pptx")
```

To open an existing presentation, you can pass the file path to the constructor:

```python
from pyppt import PyPPT

# Open an existing presentation
# (Assuming you have a presentation named 'existing_presentation.pptx')
# preso_wrapper_existing = PyPPT("existing_presentation.pptx")

# Get the first slide
# if len(preso_wrapper_existing.slides) > 0:
#     slide_to_modify = preso_wrapper_existing.slides[0] # or .get_slide(0)
#     slide_to_modify.set_title("Modified Title for Existing Slide")
#     # Add more modifications as needed
# else:
#     print("Existing presentation has no slides.")

# Add a new slide to the existing presentation
# new_slide_in_existing = preso_wrapper_existing.add_slide()
# new_slide_in_existing.set_title("Newly Added Slide in Existing Presentation")

# Save changes (either to a new file or overwrite)
# preso_wrapper_existing.save("existing_presentation_modified.pptx")
```

### Text Manipulation

The `PyPPT` allows access to `PySlide` objects, which provide methods to add and modify various text elements on your slides.

**Setting Titles and Subtitles**

You can set the title and subtitle for a slide if its layout includes placeholders for them using methods on the `PySlide` object.

```python
# Assuming 'preso_wrapper' is an instance of PyPPT
# Add a slide first if the presentation is new
# new_slide = preso_wrapper.add_slide()
# new_slide.set_title("A Title for the New Slide")

# Or, if slides already exist:
if len(preso_wrapper.slides) > 0:
    # Get a specific slide by index
    slide_to_edit = preso_wrapper.get_slide(0) # Gets the first slide as a PySlide
    try:
        slide_to_edit.set_title("My Presentation Title")
        slide_to_edit.set_subtitle("A Subtitle for the First Slide")
    except AttributeError as e:
        print(f"Error setting title/subtitle on slide 0: {e}. Ensure slide layout has these placeholders.")
else:
    print("No slides to set title/subtitle on.")
```

**Managing Footers and Slide Numbers**

Set footer text on a specific slide using its `PySlide` object:
```python
if len(preso_wrapper.slides) > 0:
    slide_for_footer = preso_wrapper.slides[0] # Get the first slide wrapper
    try:
        slide_for_footer.set_footer_text("Confidential - Company Use Only")
    except AttributeError as e:
        print(f"Error setting footer on slide 0: {e}. Ensure slide layout has a footer placeholder.")
else:
    print("No slides to set footer on.")
```

Control slide number visibility (this method remains on `PyPPT` as it affects multiple slides or presentation-level settings):
```python
# Attempt to make slide numbers visible
preso_wrapper.set_slide_numbers_visibility(True)
# Attempt to hide slide numbers
# preso_wrapper.set_slide_numbers_visibility(False)
```
*Note: Slide number visibility is heavily dependent on the slide master and layout configurations.*

**Adding Text Boxes**

Add a simple text box to a slide at a specified position (in inches) using its `PySlide` object:
```python
if len(preso_wrapper.slides) > 0:
    slide_for_textbox = preso_wrapper.slides[0] # Get the first slide wrapper
    try:
        slide_for_textbox.add_text_box("This is some important text.", left=1, top=2, width=4, height=1)
    except Exception as e: # General exception for add_text_box
        print(f"Error adding text box to slide 0: {e}")
else:
    print("No slides to add a text box to.")
```

**Adding Bullet Points**

Add a text box with bullet points using its `PySlide` object:
```python
if len(preso_wrapper.slides) > 0:
    slide_for_bullets = preso_wrapper.slides[0] # Get the first slide wrapper
    try:
        items = ["Bullet point 1", "Another bullet point", "Final bullet"]
        slide_for_bullets.add_bullet_point_box(items, left=1, top=3.5, width=5, height=2)
    except Exception as e: # General exception for add_bullet_point_box
        print(f"Error adding bullet points to slide 0: {e}")
else:
    print("No slides to add bullet points to.")
```

### Adding Pandas DataFrames as Tables

The `PySlide` class provides a method to easily insert a Pandas DataFrame as a table onto a slide, with options for custom column names and number formatting.

First, ensure you have `pandas` installed (it's listed in `requirements.txt`).

```python
import pandas as pd
# Assuming 'slide' is a PySlide object obtained from PyPPT
# (e.g., new_slide = preso.add_slide() or existing_slide = preso.slides[0])

# Create a sample DataFrame
data = {
    'product_name': ['Laptop', 'Mouse', 'Keyboard', 'Monitor'],
    'sales_id': [101, 102, 103, 104],
    'quantity_sold': [150, 300, 200, 100],
    'unit_price': [1200.00, 25.50, 75.00, 300.99],
    'revenue': [180000.00, 7650.00, 15000.00, 30099.00]
}
df = pd.DataFrame(data)
df.set_index('sales_id', inplace=True) # Example with an index

# Define custom column labels for the table
custom_labels = {
    'product_name': 'Product',
    'quantity_sold': 'Units Sold',
    'unit_price': 'Price per Unit ($)',
    'revenue': 'Total Revenue ($)'
}

# Define number formats for specific columns
# Uses Python's format string syntax
number_formatting = {
    'unit_price': '.2f',    # Format as float with 2 decimal places
    'revenue': ',.2f',     # Format as float with comma separator and 2 decimal places
    'quantity_sold': ',d'  # Format as integer with comma separator
}

# Add the DataFrame as a table to the slide
try:
    table_shape = slide.add_table_from_dataframe(
        dataframe=df,
        left=0.5, top=2.0, width=9.0, height=1.5, # Position and size in Inches
        column_labels=custom_labels,
        number_formats=number_formatting,
        include_index=True,
        index_label="Sales ID"
    )
    print("DataFrame table added to the slide.")
except Exception as e:
    print(f"Error adding DataFrame table: {e}")
```

**Parameters for `add_table_from_dataframe`:**

*   `dataframe` (pd.DataFrame): The Pandas DataFrame to insert.
*   `left`, `top`, `width`, `height` (float): Position and size of the table in Inches.
*   `column_labels` (dict, optional): A dictionary mapping original DataFrame column names to custom display names for the table header (e.g., `{'df_col_name': 'Display Name'}`).
*   `number_formats` (dict, optional): A dictionary mapping original DataFrame column names to Python format strings (e.g., `{'price_col': '$,.2f'}`). Applied to cell values.
*   `include_index` (bool, optional): If `True`, includes the DataFrame's index as the first column. Defaults to `False`.
*   `index_label` (str, optional): If `include_index` is `True`, this string is used as the header for the index column. Defaults to the DataFrame's index name or "Index".

The method returns the `GraphicFrame` object representing the table.