# Background-reprocessing-xlsx-file
Background reprocessing xlsx file
# The function first defines a nested function bg_processing_action(), which performs the actual background processing task.
# The bg_processing_action() function first prompts the user to select an input XLSX file.
# If a file is selected, the function reads the file and renames the sheets.
# The function then creates a Tkinter window for column selection.
# The window has a label for column selection, a combobox for column selection, and a button to confirm column selection.
# When the user clicks the confirm button, the function gets the chosen column and creates a new sheet for processing.
# The function then prompts the user to select an output directory.
# The function then gets the input filename without extension, constructs the output file path, and saves the processed sheet to a new workbook.
# Finally, the function opens the folder where the output files are saved.
