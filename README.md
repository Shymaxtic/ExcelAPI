API for readinng/writing Excel file having header

![output](https://user-images.githubusercontent.com/23006460/67298217-26779400-f515-11e9-970f-9de2672d5789.png)




1. Open Excel file at Sheet1 and header range "A1:G2":

    testFile = ExcelFile("Template_1.xlsx")
    testFile.Open("Sheet1", "A1:G2")

2. Read data at 1st row:

    testFile.Read(1)

2. Read data at filed "Name":

    testFile.ReadByField("Name")

3. Write to filed "Name", at 1st row:

    testFile.Write("Name", 1, "Person1")

4. Read data at field "Start Date:Coding" and "Start Date:Testing" where field "Name" has value "Person1" and field "Age" has value "1":

    testFile.ReadByCondition(["Start Date:Coding", "Start Date:Testing"], {"Name":"Person1", "Age": "1"})
