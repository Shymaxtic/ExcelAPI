import sys

sys.path.append("../ExcelApi")
import openpyxl


from ExcelFile import ExcelFile


outputFile = ExcelFile("Template_2.xlsx")

# Open excel file at Sheet1, with header range is "A1:G2"
print("------------------Sheet1--------------------------")
outputFile.Open("Sheet1", "A1:F2")
# test read: read data at 1st data row
print("---------------------------Read data at 1st data row")
ret = outputFile.Read(1)
for key in ret:
    print(key, ret[key])

# test read: read data at field "Name"
print("---------------------------Read data at field \"Name\"")
ret = outputFile.ReadByField("Name")
print(ret)

# test read: read data at field "Start Date:Coding" and "Start Date:Testing" where field "Name" has value "Person1" and field "Age" has value "1"
print("---------------------------Read data at condition")
ret = outputFile.ReadByCondition(["Start Date:Coding", "Start Date:Testing"], {"Name":"Person1"})
for key in ret:
    print(key, ret[key])

print("---------------------------Read row at condition")    
ret = outputFile.ReadRowByCondition({"Name":"Person1"})
for key in ret:
    print(key, ret[key])

outputFile.Save()