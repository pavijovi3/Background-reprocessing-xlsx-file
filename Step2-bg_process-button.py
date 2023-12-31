import os
import tkinter as tk
from tkinter import Tk, ttk, filedialog, messagebox
import pandas as pd


# The function first defines a nested function bg_processing_action(), which performs the actual background processing task.
# The bg_processing_action() function first prompts the user to select an input XLSX file.
# If a file is selected, the function reads the file and renames the sheets.
# The function then creates a Tkinter window for column selection.
# The window has a label for column selection, a combobox for column selection, and a button to confirm column selection.
# When the user clicks the confirm button, the function gets the chosen column and creates a new sheet for processing.
# Calculate the differences between the selected column and other columns and store the processed data in a DataFrame.
# The function then prompts the user to select an output directory.
# The function then gets the input filename without extension, constructs the output file path, and saves the processed sheet to a new workbook.
# Finally, the function opens the folder where the output files are saved.


def bg_processing():
    def bg_processing_action():
        # Select the CSV file
        xlsx_file_path = filedialog.askopenfilename(title="Select Input XLSX File", filetypes=[("XLSX Files", "*.xlsx")])

        if xlsx_file_path:
            try:
                # Read the input XLSX file and rename the sheets
                xlsx = pd.ExcelFile(xlsx_file_path)
                sheets = xlsx.sheet_names
                df = None  # Initialize df variable
                wavenumber_column = None  # Initialize wavenumber_column variable
                for i, sheet in enumerate(sheets):
                    df_sheet = pd.read_excel(xlsx_file_path, sheet_name=sheet)
                    df_sheet.to_excel(xlsx_file_path, sheet_name=f"Sheet{i + 1}", index=False)
                    if i == 0:
                        df = df_sheet
                        wavenumber_column = df_sheet["Wavenumber"]

                # Create a Tkinter window for column selection
                column_window = tk.Tk()
                column_window.title("Select Column")
                column_window.geometry("300x100")

                # Create a label for column selection
                column_label = ttk.Label(column_window, text="Choose a column:")
                column_label.pack()

                # Create a combobox for column selection
                column_combobox = ttk.Combobox(column_window, values=df.columns.tolist())
                column_combobox.pack()

                # Create a button to confirm column selection
                confirm_button = ttk.Button(column_window, text="Confirm", command=column_window.quit)
                confirm_button.pack()

                # Run the column selection window
                column_window.mainloop()

                # Get the chosen column
                chosen_column = column_combobox.get()

                # Create a new sheet for processing
                processed_sheet = pd.DataFrame()
                processed_sheet["Wavenumber"] = wavenumber_column
                for column in df.columns[1:]:
                    if column == chosen_column:
                        processed_sheet[column] = 0
                    else:
                        processed_sheet[column] = df[column] - df[chosen_column]

                # Prompt for the output directory
                xlsx_output_dir = filedialog.askdirectory(title="Select Output Directory")

                # Get the input filename without extension
                input_xlsx_file_name = os.path.splitext(os.path.basename(xlsx_file_path))[0]

                # Construct the output file path
                output_xlsx_file_name = f"{input_xlsx_file_name}_{chosen_column}.xlsx"
                output_xlsx_file_path = os.path.join(xlsx_output_dir, output_xlsx_file_name)

                # Save the processed sheet to a new workbook
                with pd.ExcelWriter(output_xlsx_file_path, engine="openpyxl") as writer:
                    processed_sheet.to_excel(writer, sheet_name="Sheet1", index=False)

                # Open the folder where the output files are saved
                os.startfile(xlsx_output_dir)

                # Show completion message
                message = f"Processing completed! Output saved as:\n{output_xlsx_file_path}"
                messagebox.showinfo("Processing Complete", message)
            except Exception as e:
                messagebox.showerror("Error", "An error occurred: " + str(e))

    root = Tk()
    root.withdraw()
    bg_processing_action()


# Create the main GUI window
window = tk.Tk()
window.title("FTIR Data Processing")
window.geometry("300x230")

# Create a frame to contain the buttons and align it to the left
button_frame = tk.Frame(window, padx=20, pady=20)
button_frame.pack(anchor="w")  # Use "w" to anchor (justify) the frame to the left

# Create buttons with some styling
button_style = {
    "font": ("Helvetica", 14),
    "width": 30,
    "height": 2,
    "bd": 2,  # Border thickness
    "relief": "solid",  # Border style
}
# Create buttons for each functionality
process_background_data_button = tk.Button(window, text="Step 2: Reprocess Background", command=bg_processing,
                                           bg="light blue")

# Pack the buttons with some spacing and justify to the left
process_background_data_button.pack(pady=10, anchor="w")

# Start the tkinter event loop
window.mainloop()
