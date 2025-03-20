# Anton Parr Compression Test data extraction script

'''
Description: This script is designed to seperate your exported experimental data from a CSV file into a single excel document with 
a sheet for each experiment. See the example on Github for details.

'''
##

# Modules used in this script
import data_manipulation
import excel_data
import open_file

# Main module
if __name__ == '__main__':
    print("Please select your Rheology file. Make sure this file is in .xlsx (Excel workbook) format."
          "\n Please also make sure that the name of your workbook matches the same of the worksheet. \nlong names tend to be shortend in the worksheet.")
    # User selects the data file from their directory
    # returns the file directory and file name
    file_directory = open_file.locate_file()
    file_name = open_file.file_name_only(file_directory)

    # Extracts datasets from excel file inb pandas format
    full_data_set = excel_data.open_excel(file_directory)
    # Converts pandas dataset to numpy
    np_data_set = excel_data.convert_to_numpy(full_data_set)

    # collects all the starting rows of the experiments by looking for "Test:" in the first column.
    # no longer necessary since it's done by the "extract_sample_name" function
    test_starting_rows = data_manipulation.find_starting_row(np_data_set)

    # collects all the experiment names (labels) by using the rows identified by test_starting_rows
    sample_names = data_manipulation.extract_sample_name(full_data_set, test_starting_rows)


    # User selects the automation file from the directory and
    # provides the destination path (folder)
    # Creates a new file at the destination path
    # and returns the destination path
    destination_path = open_file.copy_file(file_name)

    # Cleans up sample names to become useful sheet names
    sheet_names = data_manipulation.clean_sample_names(sample_names)

    # Creates new sheets based on the list sheet_names
    excel_data.create_new_sheet(sheet_names, destination_path)

    excel_data.copy_data_to_sheet(test_starting_rows, full_data_set, sheet_names, destination_path)

    #data_manipulation.remove_second_header(sheet_names, destination_path)

    print(
        "processing has finished. Don't forget to change the area number for your sample. \ncurrently it's also advised to remove any additional headers that might have been copied over.")
    print("to make the most of this system, please do the following: under Formulas > define Name")
