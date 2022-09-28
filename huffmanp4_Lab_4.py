################################################################################
#
# File        huffmanp4_Lab4.py
# Author      Paul Huffman
# Email:      huffmanp4@nku.edu
# Course:     DSC200
# Section:    001
# Assignment: 4
# Date:       9/28/2022
# Brief:
#  This file contains the implementation of the prompts for Lab 4
#  
#  Hours spent on this assignment: 3
#
################################################################################
import openpyxl as oxl
from csv import writer

# This loads the file, and assigns it to the variable myWorkbook
myWorkbook = oxl.load_workbook("Lab4Data.xlsx", read_only=True, keep_vba=False, data_only=True)

# This is making a dictionary of possible catagory titles
header = ['CountryName', 'CategoryName', 'CategoryTotal']
catNames = {
    3: "child_labor_total",
    5: "child_labor_male",
    7: "child_labor_female",
    9: "child_marrage_by_15",
    11: "child_marrage_by_18",
    13: "birth_registration_total",
    15: "female_genital_mutilation_prevalence_woman",
    17: "female_genital_mutilation_prevalence_girls",
    19: "female_genital_mutilation_support_for_practice",
    21: "justification_of_wife_beating_male",
    23: "justification_of_wife_beating_female",
    25: "violent_discipline_total",
    27: "violent_discipline_male",
    29: "violent_discipline_female"
}
data = []


wsht = myWorkbook.active
ctrTotal = 1

# This loops through the active worksheet to pull the necessary data
for row in wsht.iter_rows(min_col = 2, max_col = 32, min_row = 15, max_row = 211):
    colCntr = 0
    for cell in row:
        if colCntr == 0:
            countryName = cell.value
            colCntr += 1
        elif colCntr < 2 or colCntr > 30 :
            colCntr += 1
            continue
        elif colCntr % 2 == 0 :
            colCntr += 1
            continue
        elif cell.value == '\u2013' :
            colCntr += 1
            continue
        elif cell.value == "\u2013 ":
            colCntr += 1
            continue
        else:
            catName = catNames[colCntr]
            value = cell.value
            data.append([countryName,catName,value])
            ctrTotal += 1
            colCntr += 1
        
       

# This writes the data to a csv file, then prints the total number of lines printed
with open("huffmanp4_sample.csv", "w", newline='') as wfileObj:
        mwriter = writer(wfileObj)

        mwriter.writerow(header)
        mwriter.writerows(data)

print("Total lines: %d"%(ctrTotal))