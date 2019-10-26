import sys

sys.path.append("../ExcelApi")
import openpyxl


from ExcelFile import ExcelFile


outputFile = ExcelFile("Template_1.xlsx")

# Open excel file at Sheet1, with header range is "A1:G2"
print("------------------Sheet1--------------------------")
outputFile.Open("Sheet1", "A1:G2")

# test write
# write to 9 first rows at each field correspondly
for i in range(1, 10):
    outputFile.Write("Name:", i,"Person" + str(i))
    outputFile.Write("Age", i, str(i))
    outputFile.Write("Address", i,  "Zone" +  str(i))
    outputFile.Write("Start Date:Coding", i, "10/1" + str(i))
    outputFile.Write("End Date:Coding", i, "10/2" + str(i))
    outputFile.Write("Start Date:Testing", i, "11/1" + str(i))
    outputFile.Write("End Date:Testing", i, "11/2" + str(i))

# test read: read data at 1st data row
ret = outputFile.Read(1)
for key in ret:
    print(key, ret[key])

# test read: read data at field "Name"
ret = outputFile.ReadByField("Name")
print(ret)

# test read: read data at field "Start Date:Coding" and "Start Date:Testing" where field "Name" has value "Person1" and field "Age" has value "1"
ret = outputFile.ReadByCondition(["Start Date:Coding", "Start Date:Testing"], {"Name":"Person1", "Age": "1"})
for key in ret:
    print(key, ret[key])

print("------------------Sheet2--------------------------")
# load sheet
outputFile.LoadSheet("Sheet2", "C5:I6")
# test write
# write to 9 first rows at each field correspondly
for i in range(1, 10):
    outputFile.Write("Name:", i,"Person" + str(i))
    outputFile.Write("Age", i, str(i))
    outputFile.Write("Address", i,  "Zone" +  str(i))
    outputFile.Write("Start Date:Coding", i, "10/1" + str(i))
    outputFile.Write("End Date:Coding", i, "10/2" + str(i))
    outputFile.Write("Start Date:Testing", i, "11/1" + str(i))
    outputFile.Write("End Date:Testing", i, "11/2" + str(i))

# test read: read data at 1st data row
ret = outputFile.Read(1)
for key in ret:
    print(key, ret[key])

# test read: read data at field "Name"
ret = outputFile.ReadByField("Name")
print(ret)

# test read: read data at field "Start Date:Coding" and "Start Date:Testing" where field "Name" has value "Person1" and field "Age" has value "1"
ret = outputFile.ReadByCondition(["Start Date:Coding", "Start Date:Testing"], {"Name":"Person1", "Age": "1"})
for key in ret:
    print(key, ret[key])

outputFile.Save()