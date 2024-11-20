import logging
import pandas as pd
import tkinter as tk
from tkinter import filedialog
from io import BytesIO
import warnings

def append_comments_to_tracking_sheet():
    try:
        # Create and hide the tkinter root window
        root = tk.Tk()
        root.withdraw()

        # Open file selection dialog
        input_file = filedialog.askopenfilename(
            title="Select Excel file",
            filetypes=[("Excel files", "*.xlsx")]
        )
        
        if not input_file:
            tk.messagebox.showinfo("Info", "No file selected")
            return

        try:
            # Load the Excel file
            xls = pd.ExcelFile(input_file)
            supplier_tracking_df = pd.read_excel(xls, sheet_name='Supplier Enablement Tracking Sh')
            comments_df = pd.read_excel(xls, sheet_name='Comments')
        except Exception as e:
            tk.messagebox.showerror("Error", f"Error reading Excel file:\n{str(e)}")
            return

        # Extract and reorder columns from the Comments sheet (D, C, B)
        reordered_comments = comments_df.iloc[:, [3, 2, 1]]  # Reorder D, C, B as columns

        # Loop through each row and append the reorganized comment to the corresponding row
        for index, row in reordered_comments.iterrows():
            supplier_index = index + 1  # Offset by 1 as per instruction
            comment_text = ' '.join(row.dropna().astype(str))  # Join non-null values as string
            
            if supplier_index < len(supplier_tracking_df):
                # Append the comment to the existing comments in column U
                existing_comment = supplier_tracking_df.at[supplier_index, 'Comments'] if pd.notna(supplier_tracking_df.at[supplier_index, 'Comments']) else ''
                supplier_tracking_df.at[supplier_index, 'Comments'] = existing_comment + ' ' + comment_text

        # Ask user where to save the output file
        if output_file := filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            title="Save Updated Excel File"
        ):
            # Suppress openpyxl warnings
            with warnings.catch_warnings():
                warnings.simplefilter("ignore")
                supplier_tracking_df.to_excel(output_file, index=False)
            tk.messagebox.showinfo("Success", f"File saved successfully to:\n{output_file}")

    except Exception as e:
        tk.messagebox.showerror("Error", f"An unexpected error occurred:\n{str(e)}")

if __name__ == "__main__":
    append_comments_to_tracking_sheet()
