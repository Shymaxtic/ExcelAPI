import sys

sys.path.append("../ExcelApi")
import openpyxl


from ExcelFile import ExcelFile


outputFile = ExcelFile("Template_1.xlsx")

outputFile.Open("Sheet1", "A1:G2")
