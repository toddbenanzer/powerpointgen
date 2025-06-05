from pyppt import PyPPT
import pandas as pd

# --- Create a new presentation ---
print("Creating a new presentation...")
new_preso = PyPPT()

# Add slides and get their PySlide instances
print("Adding slide 1 (default layout)...")
slide_one = new_preso.add_slide() # Returns PySlide
try:
    slide_one.set_title("Slide One Title (Default Layout)")
    print("Set title for slide one.")
    # Example: Add a quick bullet list to the first slide
    slide_one.add_bullet_point_box(["Item A", "Item B"], left=1, top=2, width=5, height=2)
    print("Added bullets to slide one.")
except AttributeError as e:
    print(f"Could not set title or add content for slide one: {e} (Layout might not support title or check placeholder names)")
except Exception as e:
    print(f"An error occurred with slide one: {e}")


print("\nAdding slide 2 (layout index 0 - typically Title Slide)...")
slide_two = new_preso.add_slide(layout_index=0) # Returns PySlide
try:
    slide_two.set_title("Slide Two (Title Slide Layout)")
    slide_two.set_subtitle("This is a subtitle for the Title Slide")
    print("Set title and subtitle for slide two.")
except AttributeError as e:
    print(f"Could not set title/subtitle for slide two: {e} (Layout might not have these placeholders).")
except Exception as e:
    print(f"An error occurred with slide two: {e}")

print("\nAdding slide 3 (layout index 1 - typically Title and Content)...")
slide_three = new_preso.add_slide(layout_index=1) # Returns PySlide
try:
    slide_three.set_title("Slide Three (Title and Content Layout)")
    # Add a text box to this slide
    slide_three.add_text_box("Content text box on slide three.", left=1, top=2, width=5, height=2)
    print("Set title and added text box for slide three.")
except AttributeError as e:
    print(f"Could not set title or add content for slide three: {e} (Layout might not have these placeholders).")
except Exception as e:
    print(f"An error occurred with slide three: {e}")

# Save the new presentation (first version)
output_filename = "sample_presentation_initial.pptx"
print(f"\nSaving initial presentation as {output_filename}...")
new_preso.save(output_filename)
print(f"Initial presentation saved: {output_filename}")


# --- Demonstrating Further Text Manipulation on Existing Slides ---
print("\n--- Demonstrating Further Text Manipulation ---")

# Check if there are any slides to manipulate
if len(new_preso.slides) > 0:
    # Get the first slide (slide_one) using the .slides property or get_slide()
    # target_slide_ops = new_preso.slides[0]
    target_slide_ops = new_preso.get_slide(0) # Using get_slide for variety

    slide_title_for_print = "N/A"
    try:
        # Attempt to get the title text if the title shape and text exist
        if target_slide_ops.pptx_slide.shapes.title and target_slide_ops.pptx_slide.shapes.title.has_text_frame and target_slide_ops.pptx_slide.shapes.title.text_frame.text:
            slide_title_for_print = target_slide_ops.pptx_slide.shapes.title.text_frame.text
    except AttributeError: # Handles cases where title shape might not exist as expected
        pass # slide_title_for_print remains "N/A"

    print(f"\nModifying first slide (Index 0 - Current Title: '{slide_title_for_print}')...")

    print("Updating title and subtitle for first slide...")
    try:
        target_slide_ops.set_title("Updated Title for First Slide")
        target_slide_ops.set_subtitle("This is an updated subtitle for the first slide.")
        print("Title and subtitle updated for first slide.")
    except AttributeError as e:
        print(f"Could not set/update title/subtitle for first slide: {e} (Layout might not have these placeholders).")

    print("\nSetting footer text for first slide...")
    try:
        target_slide_ops.set_footer_text("Updated Footer Text for First Slide")
        print("Footer text set for first slide.")
    except AttributeError as e:
        print(f"Could not set footer for first slide: {e} (Layout might not have a footer placeholder).")

    print("\nAdding another text box to the first slide...")
    try:
        target_slide_ops.add_text_box("Another text box, added during text manipulation phase.", left=1, top=4, width=5, height=1)
        print("Another text box added to the first slide.")
    except Exception as e:
        print(f"Error adding another text box to first slide: {e}")

    print("\nAdding more bullet points to the first slide...")
    try:
        more_bullet_items = ["Additional Point 1", "Additional Point 2"]
        # This will add a new text box with bullets. If you want to add to existing, that's a different logic.
        target_slide_ops.add_bullet_point_box(more_bullet_items, left=1, top=5, width=5, height=1.5)
        print("Additional bullet point box added to the first slide.")
    except Exception as e:
        print(f"Error adding more bullet points to first slide: {e}")
else:
    print("\nSkipping text manipulation examples as no slides were added.")

# --- Demonstrating Adding DataFrame as Table ---
print("\n--- Demonstrating Adding DataFrame as Table ---")

if len(new_preso.slides) == 0:
    print("Adding a new slide for the table example...")
    table_slide = new_preso.add_slide()
    try:
        table_slide.set_title("DataFrame Table Example")
    except AttributeError:
        print("INFO: New slide for table has no title placeholder.")
else:
    print("Using the first slide for the table example.")
    table_slide = new_preso.slides[0] # Get the first slide
    try:
        if not table_slide.pptx_slide.shapes.title:
             table_slide.set_title("DataFrame Table Example (on existing slide)")
        elif not table_slide.pptx_slide.shapes.title.text:
             table_slide.set_title("DataFrame Table Example (on existing slide)")
        # If title exists and has text, we might not want to overwrite it here, or choose to.
        # For this example, we'll assume if it has text, we leave it.
    except AttributeError:
        print(f"INFO: Slide {new_preso.slides.index(table_slide)} has no title placeholder, adding table without setting/checking title here.")


# Create a sample DataFrame
data = {
    'Product Name': ['Apples', 'Bananas', 'Cherries', 'Dates'],
    'Category': ['Fruit', 'Fruit', 'Fruit', 'Fruit'],
    'Quantity': [120, 250, 75, 100],
    'Unit Price': [0.99, 0.59, 2.49, 3.00],
    'Total Value': [118.8, 147.5, 186.75, 300.00]
}
df_sample = pd.DataFrame(data)
df_sample.set_index('Product Name', inplace=True)

print("Sample DataFrame created:")
print(df_sample)

custom_headers = {
    'Category': 'Type',
    'Quantity': 'Amount (Units)',
    'Unit Price': 'Price/Unit',
    'Total Value': 'Subtotal'
}

number_formats_for_table = {
    'Unit Price': '$.2f',
    'Total Value': '$,.2f',
    'Quantity': ',d'
}

try:
    print("\nAdding DataFrame to slide...")
    # Determine slide index for print message
    slide_idx_for_msg = "N/A"
    for idx, s_wrapper in enumerate(new_preso.slides):
        if s_wrapper.pptx_slide == table_slide.pptx_slide:
            slide_idx_for_msg = idx
            break

    table_shape = table_slide.add_table_from_dataframe(
        dataframe=df_sample,
        left=1, top=3, width=8, height=1.5, # Adjusted top position based on typical slide content
        column_labels=custom_headers,
        number_formats=number_formats_for_table,
        include_index=True,
        index_label='Product'
    )
    print(f"DataFrame table added to slide index {slide_idx_for_msg}.")
except Exception as e:
    print(f"Error adding DataFrame table to slide: {e}")


print("\nAttempting to set slide numbers visibility (True)...")
new_preso.set_slide_numbers_visibility(True) # This method is on PyPPT

# Re-save the presentation to include these text changes
updated_output_filename = "sample_presentation_with_text.pptx"
print(f"\nSaving presentation with all manipulations as {updated_output_filename}...")
new_preso.save(updated_output_filename)
print(f"Presentation saved: {updated_output_filename}")

print("\n--- Example Finished ---")
print(f"Please open '{output_filename}' and '{updated_output_filename}' to view the results.")

# --- Optional: Example of opening an existing presentation ---
# print("\n--- Opening an existing presentation (example) ---")
# try:
#     # To run this, first ensure 'existing_example.pptx' exists.
#     # You can create one by renaming one of the outputs from this script.
    # dummy_preso_for_opening = PyPPT()
#     # slide = dummy_preso_for_opening.add_slide()
#     # slide.set_title("Existing Presentation Example")
#     # dummy_preso_for_opening.save("existing_example.pptx")
#
#     print("Opening 'existing_example.pptx'...")
#     existing_preso = PyPPT("existing_example.pptx")
#
#     if len(existing_preso.slides) > 0:
#         print("Modifying the first slide of the existing presentation...")
#         slide_to_modify = existing_preso.slides[0] # Get PySlide for the first slide
#         slide_to_modify.set_title("Title Modified in Existing Presentation")
#         slide_to_modify.add_text_box("New text box in existing.", 1, 2, 3, 1)
#
#         print("Adding a new slide to the existing presentation...")
#         new_slide_in_existing = existing_preso.add_slide()
#         new_slide_in_existing.set_title("Newly Added Slide in Existing Presentation")
#
#         existing_output_filename = "existing_example_modified.pptx"
#         print(f"Saving modified existing presentation as {existing_output_filename}...")
#         existing_preso.save(existing_output_filename)
#         print(f"Modified existing presentation saved: {existing_output_filename}")
#     else:
#         print("'existing_example.pptx' has no slides to modify.")
#
# except FileNotFoundError:
#     print("Error: 'existing_example.pptx' not found. Please create it to run this part of the example.")
# except Exception as e:
#     print(f"Error in opening/modifying existing presentation example: {e}")
