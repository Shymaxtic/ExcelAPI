import sys

sys.path.append("../ExcelApi")
import openpyxl


from ExcelFile import ExcelFile


outputFile = ExcelFile("Template_1.xlsx")

outputFile.Open("Sheet1", "A1:G2")

# test read 
ret = outputFile.Read(1)

for key in ret:
    print(key, ret[key])

# test write
for i in range(1, 10):
    outputFile.Write("Name:", i,"Person" + str(i))
    outputFile.Write("Age", i, str(i))
    outputFile.Write("Address", i,  "Zone" +  str(i))
    outputFile.Write("Start Date:Coding", i, "10/1" + str(i))
    outputFile.Write("End Date:Coding", i, "10/2" + str(i))
    outputFile.Write("Start Date:Testing", i, "11/1" + str(i))
    outputFile.Write("End Date:Testing", i, "11/2" + str(i))


outputFile.Close()