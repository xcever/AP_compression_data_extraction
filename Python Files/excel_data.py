# This module is used for excel related manipulation

import pandas as pd
import open_file
import openpyxl
import data_manipulation

def open_excel(file_directory):

    # uses the open_file module to transform the directory into just the name
    file_name = open_file.file_name_only(file_directory)
    # Opens the rheology data file using the file_directory and file_name
    # Makes a copy of the entire data set for further processing
    print("copying dataset using pandas.")
    # Excel worksheets can have a maximum of 31 characters. Since the sheet name is based on the data_file name, we limit it to 31 characters.
    full_data_set = pd.read_excel(file_directory, sheet_name=file_name[0:31], header=None, index_col=None)

    return full_data_set

def convert_to_numpy(full_data_set):
    #Converts pandas dataset to numpy dataset for further manipulation.
    np_full_data_set = full_data_set.to_numpy()
    return np_full_data_set

def create_new_sheet(sheet_names, destination_path):
    # uses a list of names to create a new sheet for each name.
    # Load the Excel workbook
    workbook = openpyxl.load_workbook(filename=destination_path)

    # Get the source worksheet
    source_sheet = workbook['Mech. Test auto']

    for name in sheet_names:
        # Copy the source worksheet
        new_sheet = workbook.copy_worksheet(source_sheet)

        # Rename the new worksheet
        new_sheet.title = name

    # Save the changes to the workbook
    workbook.save(destination_path)
    return

def copy_data_to_sheet(test_starting_rows, full_data_set, sheet_names, destination_path):
    # Extracts and copy's data to the worksheets in the new file
    print('copying data to new sheets:')
    # collects and pastes the data into the separate sheets
    for test_row in range(len(test_starting_rows)):
        print('test row: ' + str(test_starting_rows[test_row]))
        # index over the numbers in the list until the last one that should go to the end
        if test_row == len(test_starting_rows) - 1:
            sample_data = data_manipulation.extract_sample_data(full_data_set, test_starting_rows[test_row], None)
            write_data_to_sheet(sample_data, sheet_names[test_row], destination_path)
        else:
            sample_data = data_manipulation.extract_sample_data(full_data_set, test_starting_rows[test_row], test_starting_rows[test_row + 1])
            print(sample_data)
            write_data_to_sheet(sample_data, sheet_names[test_row], destination_path)
    return

def write_data_to_sheet(sample_data, sample_name, destination_path):

    with pd.ExcelWriter(destination_path, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
        sample_data.to_excel(writer, sheet_name=sample_name, index=False, header=False, startrow=2, startcol=0)
    return
