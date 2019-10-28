import unittest
import sys

sys.path.append("../ExcelApi")
import openpyxl


from ExcelFile import ExcelFile

class TestExcelApi(unittest.TestCase):
    def setUp(self):
        self.testFile = ExcelFile("Template_2.xlsx")
        self.testFile.Open("Sheet1", "A1:F2")

    def test_ReadAtRow(self):
        expectedValue = {"Name:":"Person1",
        "Task:":"Task1", 
        "Start Date:Coding:":"10/11",
        "End Date:Coding:":"10/21",
        "Start Date:Testing:":"11/11",
        "End Date:Testing:":"11/21"}
        ret = self.testFile.Read(1)
        self.assertEqual(ret, expectedValue)

    def test_ReadAtField(self):
        expectedValue = ["10/21","10/22","10/23","10/24","10/25","10/26"]
        ret = self.testFile.ReadByField("End Date:Coding")
        self.assertEqual(ret, expectedValue)
        ret = self.testFile.ReadByField("End Date:Coding:")
        self.assertEqual(ret, expectedValue)

    def test_ReadByCondition(self):
        expectedValue = {"Task":["Task1", "Task2"], 
        "Start Date:Coding":["10/11", "10/12"],
        "Start Date:Testing":["11/11", "11/12"]}
        ret = self.testFile.ReadByCondition(["Task", "Start Date:Coding", "Start Date:Testing"],
        {"Name":"Person1"})
        self.assertEqual(ret, expectedValue)

    def test_ReadRowByCondition(self):
        expectedValue = {"Name:":["Person1", "Person1"],
        "Task:":["Task1", "Task2"], 
        "Start Date:Coding:":["10/11", "10/12"],
        "End Date:Coding:":["10/21", "10/22"],
        "Start Date:Testing:":["11/11", "11/12"],
        "End Date:Testing:":["11/21", "11/22"]}
        ret = self.testFile.ReadRowByCondition({"Name":"Person1"})
        self.assertEqual(ret, expectedValue)

if __name__ == '__main__':
    unittest.main()        