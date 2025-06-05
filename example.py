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

print("\n--- Example Finished ---")
print(f"Please open {output_filename} to view the result.")

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
