# -*- coding: utf-8 -*-
"""
Created on Tue Nov  9 12:07:11 2021

@author: AllegraNoto

"""

# Import packages
import xlsxwriter
import pandas as pd
import math 
import datetime as dt
import numpy as np
from io import BytesIO

###############################################################################
# Read in the data
###############################################################################

def make_excel_sheet(title, excel_data, highlight, column_widths, excel_sheet_title, pdf):

    # If user wants an excel sheet, initialize the Binary object to write to
    if(not pdf):
        output = BytesIO()
    else:
        # Else, create a local excel file to write to, so it can be converted later 
        file_title = title + ".xlsx"

        # Creates the Excel File
        output = xlsxwriter.Workbook(file_title)
        pdf_sheet = output.add_worksheet()
        pdf_sheet.write(0, 0, 'Hello Excel')
        output.close()

        
    # Read json data into a dataframe
    df = pd.DataFrame(excel_data)

    # Drop CO Date, which is not to be printed
    if(title == "Flight Schedule"):
        df = df.drop(columns = 'CO Date')

    ###############################################################################
    # Create a binary list to decide when to highlight a row
    ###############################################################################

    if(highlight):

        # Extract the CO Time and Flight Time from the larget dataset
        df_highlight = df[['CO Time', 'Flight Time']].copy()

        # Create an indicator column for whether to add an additional day to the Flight Time
        df_highlight['add day'] = (df_highlight['Flight Time'].str.slice(6) == '(+1)')

        # Convert both times to datetime format 
        df_highlight[['CO Time', 'Flight Time']] = (df_highlight[['CO Time', 'Flight Time']]
                                                    .apply(lambda x: pd.to_datetime(x.str.slice(0,5), format = '%H:%M'), axis=1))

        # Add an extra day where necessary
        df_highlight['Flight Time'] = np.where(df_highlight['add day'],
                                            df_highlight['Flight Time'] + dt.timedelta(days = 1),
                                            df_highlight['Flight Time'])

        # Find the time between CO Time and Flight Time
        df_highlight['delta time'] = (df_highlight['Flight Time'] - df_highlight['CO Time']).dt.seconds / 3600

        # Create a binary indicator of whether the change in time is less than 7 hours
        #   or more than 12 hours
        df_highlight['highlight'] = np.where((df_highlight['delta time'] > 12) | (df_highlight['delta time'] < 7),
                                            True,
                                            False)

        # Only retain the indicator column of whether to highlight
        df_highlight = df_highlight[['highlight']].copy()
    else:
        df_highlight = df.copy()


    ###############################################################################
    # Prepare export
    ###############################################################################

    # Instantiate the excel writer
    writer = pd.ExcelWriter(output if not pdf else file_title, engine='xlsxwriter', datetime_format='mm/dd/yyyy')

    

    # Create the workbook
    workbook = writer.book

    # Set the global font size to 8
    workbook.formats[0].set_font_size(8)  

    # Set the global font name to Arial 
    workbook.formats[0].set_font_name('Arial') 

    # Choose the name of the worksheet
    wkst_name = title

    # Add a worksheet with chosen name
    worksheet = workbook.add_worksheet(wkst_name)

    # Freeze the top four rows
    worksheet.freeze_panes(4,0)

    # Remove gridlines (option 2 means to hide printed and screen grid lines)
    worksheet.hide_gridlines(2)

    # Set zoom to 120%
    worksheet.set_zoom(120)

    #Set Column Widths, first Column is a blank column
    for index, width in enumerate(column_widths):
        worksheet.set_column(index, index, width)

    # Set title row height
    worksheet.set_row(1, 20.25) 

    # Create the header format
    title_format = workbook.add_format({
        'bold': True,
        'text_wrap': True,
        'font_name': 'Arial',
        'font_size': 16,
        'valign': 'bottom',
        'align':'left'})

    # Write the title
    worksheet.merge_range(
        
        # Merge rows 1 to 1
        1, 1, 
        
        # Merge columns 1 to 11 (B:L)
        1, 11,
        
        # Write the title text
        excel_sheet_title,
        
        # Use the title format
        title_format)

    # Set the column header height
    worksheet.set_row(3, 27.75) 

    # Create the column header format
    header_format = workbook.add_format({
        'bold': True,
        'text_wrap': True,
        'font_name': 'Arial',
        'font_size': 9,
        'valign': 'bottom',
        'align': 'center',
        'fg_color': '#E2EFDA', # Fill color hex code
        'right':1,
        'left': 1,
        'top':1,
        'border_color': '#000000'})

    # Traverse the columns and associated indices
    for idx, col_name in enumerate(df.columns):
        worksheet.write(
            # All column headers are in row 3
            3,
            
            # Add one to the index to account for the blank column
            1 + idx,
            
            # Write the column name 
            col_name, 
            
            # Use the header format
            header_format)
        
    # If we need two columns, set up the second set of column headers
    if len(df) > 54:

        #Set Column Widths
        for index, width in enumerate(column_widths[1:]):
            worksheet.set_column(index + 9, index + 9, width)
        
        # Again, traverse the columns and associated indices
        for idx, col_name in enumerate(df.columns):
            worksheet.write(
                # All column headers are in row 3
                3,
                
                # Add one to the index to account for the first set of columns 
                9 + idx, 
                
                # Write the column name
                col_name, 
                
                # Use the header format
                header_format)

    # Find the length of each column
    # If the length of the dataset is more than 65, we need two columns
    if len(df) > 54:
        # If the length of the dataset is more than one page of data, we want an 
        #   approximately even amount of data in each column
        if len(df) > 108:
            col1 = round(math.ceil(len(df))/2)
            col2 = len(df) - col1
            
        # If the data will fit on one page, fill the first column before moving to 
        #   the second 
        else:
            col1 = 54
            col2 = len(df) - col1
    # If the data will fit in one column, the length is the length of the dataframe
    else:
        col1 = len(df)
        col2 = 0
        
    # Create format for text with highlight
    text_format1 = workbook.add_format({
        'font_name': 'Arial',
        'font_size': 8,
        'align': 'center',
        'fg_color':'#FCE4D6',
        'text_wrap': True,
        'right':1,
        'left': 1,
        'top':1,
        'bottom':1,
        'border_color': '#000000'}) 

    # Create format for text without highlight
    text_format2 = workbook.add_format({
        'font_name': 'Arial',
        'font_size': 8,
        'align': 'center',
        'text_wrap': True,
        'right':1,
        'left': 1,
        'top':1,
        'bottom':1,
        'border_color': '#000000'}) 

    # Write in the actual data
    # Traverse each row in the first column
    for i in range(0, col1):
        # Set each row height 
        # worksheet.set_row(i + 4, 13.5) 
        # Initialized value, will be multiplied by 11.25 depending on how wide cell content is
        row_height = 1
        
        # Traverse each column of data
        for j in range(0, len(df.columns)):

            # Will check to see if contents of a cell are longer than it's width
            # Row height will later be adjusted depending on how long the text is
            if(type(df.iloc[i,j]) == str):
                num_of_chars = len(df.iloc[i,j]) / 14
                row_height = num_of_chars if num_of_chars > row_height else row_height     
            
            # If the corresponding highlight indicator is true
            if df_highlight.iloc[i, 0] == True:
                worksheet.write(
                    # Write to row of the index plus 4 (the size of the header rows)
                    i + 4,
                    
                    # Write to the column of the index plus 1 (for the blank column)
                    j + 1,
                    
                    # Write the data in row i and column j
                    df.iloc[i, j],
                    
                    # Use the format corresponding to highlighted data
                    text_format1)
                
                
            # If the corresponding highlight indicator is false
            else: 
                worksheet.write(
                    # Write to row of the index plus 4 (the size of the header rows)
                    i + 4,
                    
                    # Write to the column of the index plus 1 (for the blank column)
                    j + 1,
                    
                    # Write the data in row i and column j
                    df.iloc[i, j],
                    
                    # Use the format corresponding to the non-highlighted data
                    text_format2)
        # Will adjust row height if contents of row are longer than cell
        worksheet.set_row(i + 4, 11.25 * row_height if row_height % 1 == 0 else 11.25 * math.ceil(row_height))
            
            
                
                
                
    # If a second column exists
    if col2 != 0:
        
        # Traverse data that will be in column 2
        for i in range(0, col2):
            row_height = 1
            
            # Again, set the row height
            # worksheet.set_row(i + 4, 13.5) 
            
            # Traverse each column of data
            for j in range(0, len(df.columns)):
                if(type(df.iloc[i,j]) == str):
                    num_of_chars = len(df.iloc[i,j]) / 14
                    row_height = num_of_chars if num_of_chars > row_height else row_height  
                
                # If the corresponding highlight indicator is true
                if df_highlight.iloc[col1 + i, 0] == True:
                    worksheet.write(
                        # Write to row of the index plus 4 (the size of the header rows)
                        i + 4,
                        
                        # Write to the column of the index plus 9 (first set of columns)
                        j + 9,
                        
                        # Write the data in row i past the first column of data and 
                        #   column j
                        df.iloc[col1 + i, j],
                        
                        # Use the format corresponding to the non-highlighting data
                        text_format1)
                
                # If the corresponding highlight indicator is false
                else:
                    worksheet.write(
                        # Write to row of the index plus 4 (the size of the header rows)
                        i + 4,
                        
                        # Write to the column of the index plus 9 (first set of columns)
                        j + 9,
                        
                        # Write the data in row i past the first column of data and 
                        #   column j
                        df.iloc[col1 + i, j],
                        
                        # Use the format corresponding to the non-highlighting data
                        text_format2)
            # Will adjust row height if contents of row are longer than cell
            worksheet.set_row(i + 4, 11.25 * row_height if row_height % 1 == 0 else 11.25 * math.ceil(row_height))

    # Set print scale to 80%
    worksheet.set_print_scale(80)

    # Set the margins for print options
    worksheet.set_margins(
        left = 0.25,
        right = 0.25,
        top = 0.4,
        bottom = 0.4)
        
    # Save the workbook
    writer.close()

    # If no PDF needed, return the binary Stream data
    if(not pdf):
        output.seek(0)
        return output
    else:
    #Else return 0
        return 0


