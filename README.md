# AP_compression_data_extraction
Script for Anton Parr Compression data CSV files to extract individual experiments into separate excel sheets within one excel file.

## Description:
This package consists of a python script, an excel template sheet, and some test data you can use to try out the script.

The main purpose of these components is to enable the processing of large datasets that you collect from your Anton Parr compression experiments.
Instead of manually processing each individual experiment, the script extracts each experiment and places it into a separate sheet with the experiment name as its label.

### The excel template sheet has 3 sheets:
1. **ImageJ**
  - used to calculate pixel to mm conversion
  - used to input the area data for each experimental sample.
  - has the code for setting SheetNames (needed to run macros that automate the excel sheet)
2. **Summary sheet**
  - gives a summary of the storage moduli in Pa, kPa with it's R² value.
  - we do triplicate repeats of our samples, so averages for that are also included.
3. **Mech. Test auto**
  - This is the template sheet that the script will look for and will add the data into.
    - it will make a copy and paste in the data.
  - If you are running this script for the first time, make sure that
    - your data is correctly stored
    - check the Parameters to match your requirements
    - manually check if the upper and lower strains match the rows they provide in the grey box
  - _Strain start:_ indicates the "starting row" where we assume the measurement is actually beginning
  - _Strain lower end:_ returns the row number and Stress value for that particular strain
  - _Strain upper end:_ returns the row number and Stree value for that particular strain
  - The slope is caluclated from these two values and returned as the youngs modulus.

## Pre-requesits:
In order to run this script you need to have the following:
- python 3
- pip installer for libaries
    - numpy
    - pandas
    - openpyxl
- python compatible compiler or software that can run the script. (Pycharm, Thonny, Sublime Text, VSC...)
    - If you're stuck on this step, I suggest following a youtube tutorial on how to set up python on your computer.


## Running the Program:
Once you've downloaded the file package, open it in your preferred coding environment to run it.
Remember this golden rule: crap in = crap out. 
So let's do a quick preflight check:
1. _Naming your experiments:_
   - Anton Parr typically gives you an experiment name that looks like this: "12/03/2024 2:50 pm experiment_name"
   - The script breaks this string (that entire name) down into pieces based on where the spaces are in the text and grabs the last piece of text.
     - Make sure that your experiment name therefore is connected in some way.
     - **Good example:** "12/03/2024 2:50 pm s_1_cond_1_rep_1" → sheetname output: "s_1_cond_1_rep_1"
     - **Bad example:** "12/03/2024 2:50 pm s 1 cond 1 rep 1" → sheetname output: "1"
   - Note: there is a character limit of 31 charcters for a sheet name in excel.
