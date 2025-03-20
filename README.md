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
### Pre-flight Check
1. **Naming your experiments:**
   - Anton Parr typically gives you an experiment name that looks like this: "12/03/2024 2:50 pm experiment_name"
   - The script breaks this string (that entire name) down into pieces based on where the spaces are in the text and grabs the last piece of text.
     - Make sure that your experiment name therefore is connected in some way.
     - **Good example:** "12/03/2024 2:50 pm s_1_cond_1_rep_1" → sheetname output: "s_1_cond_1_rep_1"
     - **Bad example:** "12/03/2024 2:50 pm s 1 cond 1 rep 1" → sheetname output: "1"
   - Note: there is a character limit of 31 charcters for a sheet name in excel.
    
2. **Convert your .csv to an .xlsx file:**
   - Simply open the .csv file you got from the Anton Parr device and open it in excel (if it prompts you about converting something say no).
   - Save as: whatever_your_name.xlsx
   - Note: I expect you know how to save all your experiments into one csv file. That's where the strength of this script lies.
   
3. **Measuring the (area) size of your samples:**
   - Depending on your application, you may need to determine the cross-section of your sample before you begin your experiments.
   - In our lab we use casted hydrogels that have individual swelling behaviour, therefore we need to quantify this before we begin compressing them.
   - For this we use ImageJ, however if you have a better way of going about it, please go ahead.

4. **Check data columns:**
   - Because our compression setup in the Anton Parr software might have slightly different settings than yours, ensure that you have matching columns so that all the important information is copied over.
   - Incase this isn't the case, you can either change the data collection in the script, or in the excel sheet.
   - If you run into problems with this, feel free to contact me about it. This was one of the first big scripts I wrote, so it might not be the cleanest thing to work through.
  
5. **Check the parameters in the template:**
   - Depending on how soft or hard your material is, you may want to adjust the starting strain, as well as your lower and upper limits.
   - Talk with your local Rheology expert :D 


### Running the script
1. Run the script (I will assume you have the libraries installed and no errors popped up as yet.
2. The program will open a window which prompts you to open your data file. (find it and hit open)
   **Important:** make sure the data file you get from Anton Parr (.csv) is converted to **xlsx**.
3. The program will open a window to prompt you to find the template file. ("20240520 Mech testing automatic v5.xlsx" can be found in the "Excelf files" folder)
4. The program will open final window to prompt you where you would like the processed data to be stored. (pick a directory (folder) and press open).
5. The script will now begin doing its thing. Depending on how much data and how powerful your computer is, this might take a while. Grab a coffee and sit back.
6. When the script is done, it should let you know in the terminal (little box in the software executing the script).
7. The new file should be found in the directory you pointed towards, with the original name of the file with "_automated" stuck on the end.

### Final steps to processing the data:
1. Open up the newly processed file. It Should look something like this:
2. Where **ImageJ** and **Summary sheet** have a bunch of weird looking data. Don't be scared, we're about to fix that.
3. Navigate to the **ImageJ** sheet
    1. In row P1 it says: **SheetNames** → note the capital N
    2. In row P2 it says: **=REPLACE(GET.WORKBOOK(1),1,FIND("]",GET.WORKBOOK(1)),"")**
    3. In the toolbar (top) Go to: Formula > Define Name
       - For _Name:_ type in SheetNames
       - For _Refers to:_ type in =REPLACE(GET.WORKBOOK(1),1,FIND("]",GET.WORKBOOK(1)),"")
       - Press OK
    4. Most of the #NAME? should now be replaced with names of your experiments
5. Finally, we need to modify the **Area /[px/]** column (H) and fill in the correct area for each corresponding experiment (As found in sample number)
6. Now navigate to the summary sheet and collect your final data.
7. Happy researching! :D
    
