
# Dependencies
from tkinter import filedialog
import os
import shutil

def locate_file():

    # Using the file dialog box to find a file and return the directory and filename
    file_directory = filedialog.askopenfilename()
    print(file_directory)
    return file_directory

def locate_folder():

    # Using the file dialog box to find the directory and return the folder path
    folder_directory = filedialog.askdirectory()
    print(folder_directory)
    return folder_directory

def file_name_with_extension(file_directory):
    file_name_w_extension = os.path.basename(file_directory).split('/')[-1]
    print(file_name_w_extension)
    return file_name_w_extension

def file_name_only(file_directory):
    # the short name removes the xlsx part and keeps only the name
    file_name = os.path.basename(file_directory).split('/')[-1].split('.')[0]
    print(file_name)
    return file_name


def copy_file(file_name):
    # locate the excel file and return the full filepath and name
    print("located the Compression_test_auto file.")
    source_path = locate_file()
    # set the destination folder for processed data to be placed.
    print("set the destination folder for processed data to be placed.")
    destination_folder = locate_folder()

    extension = os.path.splitext(source_path)
    new_file_name = file_name + "_automated" + extension[1]
    # Ensure the source file exists before attempting to copy
    if source_path and os.path.exists(source_path):
        # Construct the destination path by joining the destination folder and the source file name
        destination_path = os.path.join(destination_folder, new_file_name)

        # Perform the copy operation
        try:
            shutil.copy(source_path, destination_path)
            print(f"File copied from {source_path} to {destination_path}")
            return destination_path
        except IOError as e:
            print(f"Unable to copy file. Error: {e}")
    else:
        print("Source file not found.")
        return None, None

    # destination_path is the full path + filename of the new file into which we can do manipulations.
    print(destination_path)
    return destination_path
