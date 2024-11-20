import logging
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import warnings
import json

# Initialize the logger
logger = logging.getLogger(__name__)

def open_file_dialog(title, filetypes):
    """
    Opens a file dialog and returns the selected file path.

    Args:
        title (str): The title of the file dialog window.
        filetypes (list of tuples): A list of tuples containing file types and extensions (e.g., [("Excel files", "*.xlsx")]).

    Returns:
        str: The path to the selected file, or an empty string if no file is selected.
    """
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    file_path = filedialog.askopenfilename(title=title, filetypes=filetypes)
    if not file_path:
        messagebox.showinfo("Info", "No file selected")
        logger.info("User cancelled file selection.")
    else:
        logger.info(f"File selected: {file_path}")
    return file_path

def save_file_dialog(defaultextension, filetypes, title):
    """
    Opens a save dialog and returns the selected file path.

    Args:
        defaultextension (str): The default extension to use for the saved file.
        filetypes (list of tuples): A list of tuples containing file types and extensions (e.g., [("Excel files", "*.xlsx")])).
        title (str): The title of the save dialog window.

    Returns:
        str: The path to the selected save location, or an empty string if the save operation is cancelled.
    """
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    file_path = filedialog.asksaveasfilename(
        defaultextension=defaultextension,
        filetypes=filetypes,
        title=title
    )
    if not file_path:
        messagebox.showinfo("Info", "Save operation cancelled")
        logger.info("User cancelled save operation.")
    else:
        logger.info(f"File saved to: {file_path}")
    return file_path

def append_comments_to_tracking_sheet():
    """
    Appends comments from the 'Comments' sheet of an Excel file to the corresponding entries in the 
    'Supplier Enablement Tracking Sh' sheet. Comments are maintained in their original order.
    """
    try:
        # Open the input Excel file dialog
        logger.info("Initiating file selection for input Excel.")
        input_file = open_file_dialog(
            title="Select Excel file",
            filetypes=[("Excel files", "*.xlsx")]
        )
        input_file = 'Supplier Enablement Tracking Sheet (12).xlsx'
        # If no file is selected, exit the function
        if not input_file:
            return

        # Load the Excel file
        logger.info(f"Loading Excel file from: {input_file}")
        xls = pd.ExcelFile(input_file)
        supplier_tracking_df = pd.read_excel(xls, sheet_name='Supplier Enablement Tracking Sh')
        comments_df = pd.read_excel(xls, sheet_name='Comments')

        logger.info("Processing comments.")

        # Rename columns for clarity
        comments_df.columns = ['row_number', 'comment_text', 'commentor', 'timestamp']

        # Group comments by row number and create structured objects
        grouped_comments = comments_df.groupby('row_number').apply(
            lambda x: x[['comment_text', 'commentor', 'timestamp']].to_dict('records')
        ).to_dict()

        # Convert row numbers from 'Row X' format to integers
        cleaned_comments = {}
        for key, value in grouped_comments.items():
            if isinstance(key, str) and key.startswith('Row '):
                row_num = int(key.split('Row ')[1])
                cleaned_comments[row_num] = value

        # Add comments to supplier tracking sheet
        for row_num, comments in cleaned_comments.items():
            if row_num <= len(supplier_tracking_df):
                formatted_comments = [comment['comment_text'] for comment in comments]
                # Join all comments with line breaks
                supplier_tracking_df.at[row_num-1, 'Comments'] = '\n'.join(formatted_comments)

        # Open the save dialog to specify output file path
        logger.info("Initiating file selection for saving updated Excel.")
        output_file = save_file_dialog(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            title="Save Updated Excel File"
        )

        # If no file is selected, exit the function
        if not output_file:
            return

        # Suppress openpyxl warnings and save the updated DataFrame
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            logger.info(f"Saving updated data to: {output_file}")
            supplier_tracking_df.to_excel(output_file, index=False)

        messagebox.showinfo("Success", f"File saved successfully to:\n{output_file}")
        logger.info("File save operation completed successfully.")

    except Exception as e:
        logger.error(f"An unexpected error occurred: {str(e)}")
        messagebox.showerror("Error", f"An unexpected error occurred:\n{str(e)}")

if __name__ == "__main__":
    # Run the main function to append comments
    append_comments_to_tracking_sheet()