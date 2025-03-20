# This module is for compression data specific actions
# IMPORT
import numpy as np
import pandas as pd
import excel_data
import re

def find_starting_row(np_data_set):
    # Finds the rows that separate the different experiments by searching for "Test:". Returns them as a list
    print('collecting sample starting rows:')
    find_test = 'Test:'
    test_starting_rows = np.where(np_data_set == find_test)[0]
    print(test_starting_rows)
    return test_starting_rows


def extract_sample_name(full_data_set, test_starting_rows):
    # Takes a numpy dataset
    # Finds the rows where the tests starts
    # test_starting_rows = find_starting_row(np_data_set)
    # sample_names list is used to store all the different test names
    sample_names = []
    # Looks in column 1 at all the row numbers provided by test_starting_rows and returns the name
    for name in range(0, len(test_starting_rows)):
        sample_names.append(full_data_set.iloc[test_starting_rows[name], 1])

    print(sample_names)
    return sample_names


def clean_sample_names(sample_names):
    # Creates new sheets for each experiment using the sample_names as their worksheet titles
    print('collecting sheet names:')
    sheet_names = []
    for name in sample_names:

        # split for newbs
        #new_name = name.split('pm-')
        # typical split
        new_name = name.split(' ')
        clean_name = new_name[-1]
        sheet_names.append(clean_name)

    print(sheet_names)
    return sheet_names


def extract_sample_data(full_data_set, test_start, test_finish):
    # If there is no next row in the test_starting_rows, test_finish will be None
    # and the data will be collected until the last full cell.
    if test_finish is None:
        sample_data = full_data_set.iloc[test_start:, :]
    else:
        sample_data = full_data_set.iloc[test_start:test_finish, :]

    print("these are the row numbers.")
    print(sample_data)
    clean_sample_data = find_second_header(sample_data)
    return clean_sample_data

def is_number(number):
    # simple test used by find_second_header to test if something is a number.
    return isinstance(number, (int, float))

def find_second_header(sample_data):
    # grabs the 4th column and compares the cells -> are they numbers or nan/text.
    # if they are text or nan, they are added to a list of rows that are removed at the end.
    # the data is then returned as clean data.
    remove_rows = []

    for rows in range(len(sample_data)):
        #print(sample_data.iloc[rows,3])
        if is_number(sample_data.iloc[rows, 3]):
            if str(sample_data.iloc[rows, 3]) == 'nan':
                remove_rows.append(rows)
            continue
        else:
            remove_rows.append(rows)
    print(remove_rows)

    # drops rows that are empty or not numerical data
    clean_data = sample_data.drop(index=sample_data.index[remove_rows[8:]])

    return clean_data


def remove_second_header(sample_data, second_row):
    # Convert sample_data to DataFrame
    panda_sample_data = pd.DataFrame(sample_data)
    # Index number is required because index isn't changed during the copying process.
    sample_data_index = panda_sample_data.index[0]
    print(sample_data_index)
    # Calculate the range of rows to be removed
    rows_to_remove = list(range(sample_data_index + second_row - 1, sample_data_index + second_row + 3))
    print(rows_to_remove)
    # Drop the rows and create a new DataFrame
    clean_data = panda_sample_data.drop(index=rows_to_remove, axis=0, inplace=False)

    return clean_data


def remove_empty_rows(clean_data, empty_row):
    # Convert sample_data to DataFrame
    panda_sample_data = pd.DataFrame(clean_data)
    # Index number is required because index isn't changed during the copying process.
    sample_data_index = panda_sample_data.index[0]
    print(sample_data_index)
    # Calculate the range of rows to be removed
    rows_to_remove = sample_data_index + empty_row
    print("remove row: " + rows_to_remove)
    # Drop the rows and create a new DataFrame
    clean_data = panda_sample_data.drop(index=rows_to_remove, axis=0, inplace=False)

    return clean_data
