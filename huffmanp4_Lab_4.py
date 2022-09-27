import openpyxl as oxl
from csv import writer

myWorkbook = oxl.load_workbook("Lab4Data.xlsx", read_only=True, keep_vba=False, data_only=True)

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

for row in wsht.iter_rows(min_col = 2, max_col = 32, min_row = 15, max_row = 211):
    colCntr = 0
    for cell in row:
        if colCntr == 0:
            countryName = cell.value
        elif colCntr > 2 and colCntr < 30 and colCntr % 2 == 1 and cell.value != '\u2013' and cell.value != "\u2013 ":
            catName = catNames[colCntr]
            value = cell.value
            data.append([countryName,catName,value])
            ctrTotal += 1
        
        colCntr += 1


with open("huffmanp4_sample.csv", "w", newline='') as wfileObj:
        mwriter = writer(wfileObj)

        mwriter.writerow(header)
        mwriter.writerows(data)

print("Total lines: %d"%(ctrTotal))