from pyppt import PresentationWrapper

# --- Create a new presentation ---
print("Creating a new presentation...")
new_preso = PresentationWrapper()

# Add a slide using the default layout (typically 'Title and Content')
print("Adding slide 1 (default layout)...")
new_preso.add_slide()

# Add a title-only slide (layout_index=5 is often 'Title and Content', 0 is often 'Title Slide', 1 is often 'Title and Content')
# We'll try to add a few common ones.
# Specific indices can vary, so this is for demonstration.
# Users should check their specific default template's slide layouts.
print("Adding slide 2 (layout index 0 - typically Title Slide)...")
new_preso.add_slide(layout_index=0)

print("Adding slide 3 (layout index 1 - typically Title and Content)...")
new_preso.add_slide(layout_index=1)

# Save the new presentation
output_filename = "sample_presentation.pptx"
print(f"Saving new presentation as {output_filename}...")
new_preso.save(output_filename)
print(f"Presentation saved: {output_filename}")


# (Assuming new_preso is an existing PresentationWrapper instance from earlier in the script)
print("\n--- Demonstrating Text Manipulation ---")

if len(new_preso.presentation.slides) > 0:
    print("\nSetting title and subtitle for slide 0...")
    try:
        new_preso.set_title(0, "Example Title")
        # For the first slide, let's use a layout that typically has a subtitle.
        # If new_preso.add_slide() added a "Title and Content" (layout 1 or 5), it should have a title.
        # If new_preso.add_slide(layout_index=0) was the first slide, it's a "Title Slide", which has a subtitle.
        # We'll assume the first slide added (index 0) is suitable.
        # The example script adds slides with layout_index default, then 0, then 1.
        # So, slide 0 (first one added) is default layout. Slide 1 is layout 0. Slide 2 is layout 1.
        # Let's try to set title on slide 1 (which used layout_index=0, often "Title Slide")
        # and subtitle on that same slide.
        # Or, to be safe, let's target slide 1 (which was created with layout_index=0) for both.
        # The example script adds 3 slides. Slide 0, 1, 2.
        # Slide 1 (index 1) was created with layout_index=0 (Title Slide)
        new_preso.set_title(1, "Title for Slide with Layout 0")
        new_preso.set_subtitle(1, "This is a subtitle on a Title Slide Layout.")
        print("Title and subtitle set for slide 1.")
    except AttributeError as e:
        print(f"Could not set title/subtitle for slide 1: {e} (Layout might not have these placeholders).")
    except IndexError:
        print("Cannot set title/subtitle, slide 1 not found.")

    print("\nSetting footer text for slide 0...")
    try:
        new_preso.set_footer_text(0, "Sample Footer Text - Slide 1")
        print("Footer text set for slide 0.")
    except AttributeError as e:
        print(f"Could not set footer for slide 0: {e} (Layout might not have a footer placeholder).")
    except IndexError:
        print("Cannot set footer, slide 0 not found.")

    print("\nAdding a text box to slide 0...")
    new_preso.add_text_box(0, "Hello from a text box on Slide 1!", left=1, top=2.5, width=3, height=0.5)
    print("Text box added to slide 0.")

    print("\nAdding bullet points to slide 0...")
    bullet_items = ["Introduction", "Main Points", "Conclusion"]
    new_preso.add_bullet_point_box(0, bullet_items, left=1, top=3.5, width=4, height=1.5)
    print("Bullet point box added to slide 0.")
else:
    print("\nSkipping text manipulation examples as no slides were added.")

print("\nAttempting to set slide numbers visibility (True)...")
new_preso.set_slide_numbers_visibility(True)
# (Add a note that user should check the output file)

# Re-save the presentation to include these text changes
updated_output_filename = "sample_presentation_with_text.pptx"
print(f"\nSaving presentation with text manipulations as {updated_output_filename}...")
new_preso.save(updated_output_filename)
print(f"Presentation saved: {updated_output_filename}")

print("\n--- Example Finished ---")
print(f"Please open {output_filename} and {updated_output_filename} to view the results.")

# --- Optional: Example of opening an existing presentation ---
# print("\n--- Opening an existing presentation (example) ---")
# try:
#     # Create a dummy presentation first to ensure the open example can run
#     # In a real scenario, this file would already exist.
#     dummy_preso_for_opening = PresentationWrapper()
#     dummy_preso_for_opening.add_slide(layout_index=5) # Add a blank slide
#     dummy_preso_for_opening.save("existing_example.pptx")
#
#     print("Opening 'existing_example.pptx'...")
#     existing_preso = PresentationWrapper("existing_example.pptx")
#
#     print("Adding a new slide to the existing presentation...")
#     existing_preso.add_slide(layout_index=5) # Add another blank slide
#
#     existing_output_filename = "existing_example_modified.pptx"
#     print(f"Saving modified presentation as {existing_output_filename}...")
#     existing_preso.save(existing_output_filename)
#     print(f"Modified presentation saved: {existing_output_filename}")
#
# except Exception as e:
#     print(f"Error in opening/modifying existing presentation example: {e}")
#     print("This part of the example requires a file named 'existing_example.pptx' to be present.")
