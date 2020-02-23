<a href="https://scan.coverity.com/projects/shymaxtic-excelapi">
  <img alt="Coverity Scan Build Status"
       src="https://scan.coverity.com/projects/19458/badge.svg"/>
</a>

# ExcelAPI

API for reading/writing Excel file having header

![output](https://user-images.githubusercontent.com/23006460/67298217-26779400-f515-11e9-970f-9de2672d5789.png)




<b>1. Open Excel file at Sheet1 and header range "A1:G2":</b>

    testFile = ExcelFile("Template_1.xlsx")
    testFile.Open("Sheet1", "A1:G2")

<b>2. Read data at 1st row:</b>

    testFile.Read(1)

<b>3. Read data at filed "Name":</b>

    testFile.ReadByField("Name")

<b>4. Write to filed "Name", at 1st row:</b>

    testFile.Write("Name", 1, "Person1")

<b>5. Read data at field "Start Date:Coding" and "Start Date:Testing" where field "Name" has value "Person1" and field "Age" has value "1":</b>

    testFile.ReadByCondition(["Start Date:Coding", "Start Date:Testing"], {"Name":"Person1", "Age": "1"})
