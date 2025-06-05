# powerpointgen

## PPTX Wrapper Module

This module provides a simple wrapper around the `python-pptx` library to simplify creating and manipulating PowerPoint presentations.

### Installation

To install the necessary dependencies, run:
```bash
pip install -r requirements.txt
```

**Note on Importing Enums:**
Common `python-pptx` enums needed for certain operations, like `MSO_SHAPE` (for adding shapes), `XL_CHART_TYPE` (for adding charts), and `PP_PLACEHOLDER` (relevant for understanding slide structure if customizing layouts or dealing with placeholders directly), are re-exported for convenience. You can import them directly from `pyppt`:
```python
from pyppt import PyPPT, PySlide, MSO_SHAPE, XL_CHART_TYPE, PP_PLACEHOLDER
```

### Basic Usage

Here's how to create a new presentation, add a slide, and save it:

```python
from pyppt import PyPPT

# Create a new presentation
preso_wrapper = PyPPT()

# Add a slide (uses default 'Title and Content' layout)
# This now returns a PySlide instance
# The add_slide method can take a layout index (int) or layout name (str).
new_slide = preso_wrapper.add_slide() # Defaults to layout_ref=5
new_slide.set_title("Title for New Slide") # Example of using the returned slide

# Add another slide using a specific layout index (e.g., 0 for 'Title Slide')
# Or by layout name, e.g., preso_wrapper.add_slide("Title Slide")
# (Actual layout names depend on the presentation's design template)
title_slide_obj = preso_wrapper.add_slide(layout_ref=0)
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

# Define styling options
font_options = {
    'font_name': 'Arial',
    'font_size': 10
}
header_style_options = {
    'header_bold': True,
    'header_font_color_rgb': (255, 255, 255), # White text
    'header_fill_color_rgb': (79, 129, 189)  # A common blue shade
}
# Example column widths (first column for index, others for df.columns)
# Assumes df has 4 columns, and index is included, total 5 columns in table.
# The method converts these floats to Inches internally.
col_widths_example = [1.5, 2.0, 1.0, 1.5, 2.0]


# Add the DataFrame as a table to the slide with styling
try:
    table_shape = slide.add_table_from_dataframe(
        dataframe=df,
        left=0.5, top=2.0, width=9.0, height=1.5, # Position and size
        column_labels=custom_labels,
        number_formats=number_formatting,
        include_index=True,
        index_label="Sales ID",
        font_name=font_options['font_name'],
        font_size=font_options['font_size'],
        header_bold=header_style_options['header_bold'],
        header_font_color_rgb=header_style_options['header_font_color_rgb'],
        header_fill_color_rgb=header_style_options['header_fill_color_rgb'],
        column_widths=col_widths_example
    )
    print("DataFrame table with custom styling added to the slide.")
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
*   `font_name` (str, optional): Font name for all text in the table (e.g., "Arial").
*   `font_size` (int, optional): Font size in points for all text in the table (e.g., 10).
*   `column_widths` (list or dict, optional): List of widths (Inches) for columns by index, or a dict mapping column index to width.
*   `row_heights` (list or dict, optional): List of heights (Inches) for rows by index, or a dict mapping row index to height.
*   `header_bold` (bool, optional): If `True` (default), makes header text bold.
*   `header_font_color_rgb` (tuple, optional): An RGB tuple (e.g., `(255, 255, 255)` for white) for header text color.
*   `header_fill_color_rgb` (tuple, optional): An RGB tuple (e.g., `(0, 0, 0)` for black) for header row background fill.

The method returns the `GraphicFrame` object representing the table.

### Adding and Styling Basic Shapes

You can add various predefined shapes to your slides and apply basic styling.

**Adding a Shape**

Use the `PySlide.add_shape()` method. You'll need to import `MSO_SHAPE` from `pptx.enum.shapes`.

```python
from pptx.enum.shapes import MSO_SHAPE
# Assuming 'slide' is a PySlide object (e.g., slide = preso.slides[0])

# Add a rectangle
rect = slide.add_shape(
    MSO_SHAPE.RECTANGLE,
    left=1, top=1, width=2, height=1, # Inches
    shape_name="MyRectangle"
)
print(f"Added rectangle named: {rect.name}")

# Add an oval
oval = slide.add_shape(
    MSO_SHAPE.OVAL,
    left=4, top=1, width=2, height=1.5 # Inches
)
print("Added an oval.")
```

**Styling a Shape**

Once a shape is added (or retrieved), you can style its fill and line.
The styling methods take a reference to the shape, which can be the shape object itself, its name (if set), or its index on the slide.

```python
# Assuming 'slide' and 'rect' (the rectangle shape object from above) exist

# Set fill color for the rectangle (e.g., a light blue)
slide.set_shape_fill_color(rect, r=173, g=216, b=230)
# Or using its name:
# slide.set_shape_fill_color("MyRectangle", r=173, g=216, b=230)

# Set line color and weight for the oval
slide.set_shape_line_color(oval, r=255, g=0, b=0) # Red line
slide.set_shape_line_weight(oval, weight_pt=2.5) # 2.5 points thick

print("Applied styling to shapes.")
```

**Helper for Shape Reference (`shape_ref`)**

The `shape_ref` argument in styling methods can be:
*   The shape object itself (returned by `add_shape` or `slide.shapes[idx]`).
*   The name of the shape (string), if you assigned one using `shape_name` in `add_shape`.
*   The index of the shape (integer) in the slide's shapes collection.

**Available Styling Methods on `PySlide`**:
*   `set_shape_fill_color(shape_ref, r, g, b)`
*   `set_shape_line_color(shape_ref, r, g, b)`
*   `set_shape_line_weight(shape_ref, weight_pt)`

### Adding Charts

You can add common chart types like line and bar charts to your slides using data from Python dictionaries.

**Method: `PySlide.add_chart()`**

The `add_chart` method on a `PySlide` object allows you to insert charts.

**Parameters:**

*   `chart_type` (`XL_CHART_TYPE`): The type of chart to create (e.g., `XL_CHART_TYPE.LINE`, `XL_CHART_TYPE.COLUMN_CLUSTERED`). You'll need to import this enum: `from pptx.enum.chart import XL_CHART_TYPE`.
*   `chart_data_dict` (dict): A dictionary containing the data for the chart. It must have the following structure:
    ```python
    {
        'categories': ['Category 1', 'Category 2', 'Category 3'], # List of category labels (X-axis)
        'series': [ # List of series, each series is a dictionary
            {'name': 'Series A Name', 'values': [10.5, 20.2, 15.7]},
            {'name': 'Series B Name', 'values': [12.0, 18.5, 19.2]}
            # Add more series as needed
        ]
    }
    ```
    The length of `values` in each series must match the length of the `categories` list.
*   `left`, `top`, `width`, `height` (float): Position and dimensions of the chart on the slide, in Inches.
*   `chart_title` (str, optional): An optional title for the chart.

The method returns the `GraphicFrame` object containing the chart.

**Example: Adding a Line Chart**

```python
from pptx.enum.chart import XL_CHART_TYPE
# Assuming 'slide' is a PySlide object

line_chart_data = {
    'categories': ['Jan', 'Feb', 'Mar', 'Apr', 'May'],
    'series': [
        {'name': 'Product A Sales', 'values': [100, 120, 90, 110, 130]},
        {'name': 'Product B Sales', 'values': [80, 85, 100, 95, 105]}
    ]
}

try:
    line_chart_frame = slide.add_chart(
        XL_CHART_TYPE.LINE,
        line_chart_data,
        left=1, top=2, width=8, height=4, # Inches
        chart_title="Monthly Product Sales (Line Chart)"
    )
    print("Line chart added to the slide.")
except ValueError as e:
    print(f"Error adding line chart: {e}")
```

**Example: Adding a Clustered Column Chart**

```python
# from pptx.enum.chart import XL_CHART_TYPE # Already imported
# Assuming 'slide' is a PySlide object

column_chart_data = {
    'categories': ['Q1', 'Q2', 'Q3', 'Q4'],
    'series': [
        {'name': 'Region North', 'values': [250, 260, 280, 270]},
        {'name': 'Region South', 'values': [220, 230, 210, 240]}
    ]
}

try:
    column_chart_frame = slide.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        column_chart_data,
        left=1, top=2, width=8, height=4, # Inches
        chart_title="Quarterly Regional Performance (Column Chart)"
    )
    print("Column chart added to the slide.")
except ValueError as e:
    print(f"Error adding column chart: {e}")
```

### Presentation and Slide Management

The `PyPPT` object provides methods to manage slides within the presentation.

**Deleting a Slide**

You can delete a slide by its index:

```python
# Assuming 'preso' is a PyPPT object
# Example: Delete the slide currently at index 2
try:
    preso.delete_slide(2)
    print("Slide at index 2 deleted.")
except IndexError as e:
    print(e)
```
*Note: This operation directly modifies the presentation structure and should be used with awareness of `python-pptx` internal behaviors.*

**Moving a Slide**

Reorder slides within the presentation:

```python
# Assuming 'preso' is a PyPPT object with at least 3 slides
# Example: Move the slide currently at index 0 to be at index 2
try:
    preso.move_slide(current_index=0, new_index=2)
    print("Moved slide from index 0 to index 2.")
except IndexError as e:
    print(e) # If current_index is invalid
```
*The `new_index` is handled robustly; if out of bounds, it's adjusted to be within valid insertion points (e.g., negative becomes 0, too large appends).*

**Duplicating a Slide (Basic Implementation)**

Create a copy of an existing slide. The new slide is added at the end of the presentation.

```python
# Assuming 'preso' is a PyPPT object with at least one slide
try:
    duplicated_slide_wrapper = preso.duplicate_slide(0) # Duplicate the first slide
    print(f"Slide at index 0 was duplicated. New slide index: {len(preso.slides)-1}")
    # You can now work with duplicated_slide_wrapper (a PySlide object)
    # duplicated_slide_wrapper.set_title("Copy of First Slide")
except IndexError as e:
    print(e)
```

**Important Limitations of `duplicate_slide`:**
This is a basic implementation and does *not* perform a perfect, deep copy.
*   It copies the layout and attempts to replicate text from common placeholders (title, content - content often becomes new text boxes).
*   It tries to copy basic auto-shapes (like rectangles, ovals, lines) and their text, position, and size.
*   **It does NOT copy**: Complex shape formatting, tables, charts, images, grouped shapes, animations, transitions, or slide master details.
*   For a true clone, more advanced manipulation of the underlying presentation XML would be needed, which is beyond the scope of this basic function.