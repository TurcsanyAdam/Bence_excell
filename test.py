from openpyxl import Workbook
import string
import os

file_name = []
folder = "C:/Users/Turi/PycharmProjects/test"
for file in os.listdir(folder):
    if file.endswith(".txt"):
        file_name.append(file)


workbook = Workbook()
for i in file_name:
    workbook.create_sheet(i)

abc = string.ascii_uppercase

for txt_file in file_name:
    row = 1
    column = 0
    header = True
    sheet = workbook[txt_file]
    with open(txt_file, "r") as file:
        lines = file.readlines()
        for line in lines:
            line = line.split('|')
            if len(line) > 20 and line[1].startswith("P") and header:
                for header_data in line:
                    cell = abc[column] + str(row)
                    sheet[cell] = header_data
                    column += 1
                    header = False
                row += 1
            if len(line) > 20 and line[1].startswith("3"):
                column = 0
                for data in line:
                    cell = abc[column] + str(row)
                    sheet[cell] = data
                    column += 1
                column = 0
                row += 1

workbook.save(filename="bence.xlsx")

