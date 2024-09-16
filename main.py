import logging
import pandas as pd
import azure.functions as func
import os
from io import BytesIO

def append_comments_to_tracking_sheet(excel_file):
    # Load the Excel file into a BytesIO object
    xls = pd.ExcelFile(BytesIO(excel_file))

    # Load the relevant sheets into DataFrames
    supplier_tracking_df = pd.read_excel(xls, sheet_name='Supplier Enablement Tracking Sh')
    comments_df = pd.read_excel(xls, sheet_name='Comments')

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

    # Save the updated Supplier Enablement Tracking Sheet to a BytesIO object
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    supplier_tracking_df.to_excel(writer, index=False)
    writer.save()
    output.seek(0)
    
    return output

def main(req: func.HttpRequest) -> func.HttpResponse:
    logging.info('Processing Excel file.')

    # Ensure the request is a POST with a file
    try:
        # Get the uploaded file
        file_bytes = req.get_body()

        # Process the Excel file
        updated_excel = append_comments_to_tracking_sheet(file_bytes)

        # Create a response with the updated file
        return func.HttpResponse(
            updated_excel.getvalue(),
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={
                'Content-Disposition': 'attachment; filename=Updated_Supplier_Enablement_Tracking_Sheet.xlsx'
            }
        )
    except Exception as e:
        logging.error(f"Error processing file: {str(e)}")
        return func.HttpResponse(f"An error occurred: {str(e)}", status_code=500)
import logging
import pandas as pd
import azure.functions as func
import os
from io import BytesIO

def append_comments_to_tracking_sheet(excel_file):
    # Load the Excel file into a BytesIO object
    xls = pd.ExcelFile(BytesIO(excel_file))

    # Load the relevant sheets into DataFrames
    supplier_tracking_df = pd.read_excel(xls, sheet_name='Supplier Enablement Tracking Sh')
    comments_df = pd.read_excel(xls, sheet_name='Comments')

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

    # Save the updated Supplier Enablement Tracking Sheet to a BytesIO object
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    supplier_tracking_df.to_excel(writer, index=False)
    writer.save()
    output.seek(0)
    
    return output

def main(req: func.HttpRequest) -> func.HttpResponse:
    logging.info('Processing Excel file.')

    # Ensure the request is a POST with a file
    try:
        # Get the uploaded file
        file_bytes = req.get_body()

        # Process the Excel file
        updated_excel = append_comments_to_tracking_sheet(file_bytes)

        # Create a response with the updated file
        return func.HttpResponse(
            updated_excel.getvalue(),
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={
                'Content-Disposition': 'attachment; filename=Updated_Supplier_Enablement_Tracking_Sheet.xlsx'
            }
        )
    except Exception as e:
        logging.error(f"Error processing file: {str(e)}")
        return func.HttpResponse(f"An error occurred: {str(e)}", status_code=500)
