# Copyright (C) 2019 QuynhPP
# 
# This file is part of ExcelAPI.
# 
# ExcelAPI is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
# 
# ExcelAPI is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
# 
# You should have received a copy of the GNU General Public License
# along with ExcelAPI.  If not, see <http://www.gnu.org/licenses/>.

import openpyxl
from HeaderInfo import HeaderInfo 
from CellInfo import CellInfo
from HeaderInfo import D_SEPERATOR
import Utils

class ExcelFile:

    mPath = ""
    mHeaderRange = ""           # Ex: "A1:N1"
    mWorkBook = None
    mSheet = None
    mHeaderList = {}                # {fullheadername:header info}   
    mDataColumnSize = 0                # number of data column
    mDataRowSize  = 0                 # number of data row
    mHeaderCellInfo = {}           # list of info of main cell of header {coordinate:cell info}
    mMergedDataCellInfo = {}        # list of cell in merged cell area {coordinate:cell info}
    mPivotRow = 0                   # row at lowest header.
    mHeaderInfoColumCache = {}      # {column:fullheadername}
    mDictData = {}       


    def __init__(self, path: str):
        self.mPath = path     

    def __PostProcessMergedCell(self):
        self.mMergedDataCellInfo = {}
        self.mHeaderCellInfo = {}
        for cellRange in self.mSheet.merged_cells.ranges:
            for rowOfCell in self.mSheet[cellRange.__str__()]:
                for cell in rowOfCell:
                    # if this cell is in header range and is top-left cell
                    if Utils.IsCellInCellRange(cell.coordinate, self.mHeaderRange) and \
                        cell == self.mSheet[cellRange.__str__()][0][0]:
                        colNum = len(self.mSheet[cellRange.__str__()][0]) # get column size
                        rowNum = len(self.mSheet[cellRange.__str__()]) # get row size
                        topLeftCell = self.mSheet[cellRange.__str__()][0][0] # get top-left cell
                        self.mHeaderCellInfo[cell.coordinate] = CellInfo(cell, topLeftCell, rowNum, colNum)
                        continue
                    else:
                        colNum = len(self.mSheet[cellRange.__str__()][0]) # get column size
                        rowNum = len(self.mSheet[cellRange.__str__()]) # get row size
                        topLeftCell = self.mSheet[cellRange.__str__()][0][0] # get top-left cell
                        self.mMergedDataCellInfo[cell.coordinate] = CellInfo(cell, topLeftCell, rowNum, colNum)

        # print(self.mHeaderCellInfo)                

    def __GetHeaderCellInfo(self):
        """
        Get info of cells containing header name, including merged cells or singel cell.
        """
        self.__PostProcessMergedCell()
        # For each row in header sheet    
        for rowOfCell in self.mSheet[self.mHeaderRange]:
            for cell in rowOfCell:
                if cell.value != None and cell.coordinate not in self.mHeaderCellInfo:  # this single cell has value                    
                    self.mHeaderCellInfo[cell.coordinate] = CellInfo(cell, cell, 1, 1)

    def __GetHeaderInfo(self):
        """
        Update relationship of header cells. Create Header Info    
        """
        self.__GetHeaderCellInfo()
        sortedKeys = sorted(self.mHeaderCellInfo.keys())
        self.mHeaderList = {}
        self.mHeaderInfoColumCache = {}
        for key in sortedKeys:
            # prepare parent header
            icell = self.mHeaderCellInfo[key].mCell
            if (icell.row > 1):
                # if upper cell (merged cells/single cell) has value. It is parent of this cell
                upperCell = self.mSheet.cell(row=icell.row - 1, column=icell.column)
                # Check if upper cell is in any created header info
                for headerInfo in self.mHeaderList.values():
                    if headerInfo.mCellInfo.Has(upperCell):
                        newHeaderInfo = HeaderInfo(self.mHeaderCellInfo[key], headerInfo)
                        self.mHeaderList[newHeaderInfo.mFullName] = HeaderInfo(self.mHeaderCellInfo[key], headerInfo)
                        self.mHeaderInfoColumCache[self.mHeaderCellInfo[key].mCell.column] = newHeaderInfo.mFullName
                        break
            else:
                newHeaderInfo = HeaderInfo(self.mHeaderCellInfo[key], None) 
                self.mHeaderList[newHeaderInfo.mFullName] = (newHeaderInfo)
                self.mHeaderInfoColumCache[self.mHeaderCellInfo[key].mCell.column] = newHeaderInfo.mFullName
        # for key in self.mHeaderList:
        #     print(key, ":::", self.mHeaderList[key])
        # for key in self.mHeaderInfoColumCache:
        #     print(key, ":::", self.mHeaderInfoColumCache[key])                            

    def __CheckMatchHeader(self, checkingName: str, headerFullName: str):
        """
        Check if a name is a header full name. Ex: "Function" is name of header full name "Function:FileName:PIC"
        Note: "Function:" is not name of header full name "Function:FileName:PIC"
        """
        if (checkingName == headerFullName):
            return True
        inputStruct = checkingName.split(D_SEPERATOR, 1)
        checkStruct = headerFullName.split(D_SEPERATOR, 1)
        if (inputStruct[0] == checkStruct[0]):
            if len(inputStruct) > 1 and len(checkStruct) > 1:
                if inputStruct[1] and inputStruct[1] in checkStruct[1]:
                    return True
                else:
                    return False                    
            else:
                return True
        return False  
        
    
    # Get list of matched header with a input name
    # headerStruct: "End Date:Test term"
    # return list of CellInfo (s)
    def __GetMatchHeader(self, headerStruct: str):
        matchHeader = []
        for key in self.mHeaderList:
            if (self.__CheckMatchHeader(headerStruct, key)):
                matchHeader.append(self.mHeaderList[key].mCellInfo)
        return matchHeader     

    # Get a cell at header name and row offset. Row offset start at 1
    # Ex: GetCell("Function:PIC:Company", 1):
    # return Cell                 
    def GetCell(self, headerStruct: str, rowOffset: int):
        matchHeader = self.__GetMatchHeader(headerStruct)
        if (len(matchHeader) == 0):
                raise Exception("Cannot find header: " +  headerStruct)
        if (len(matchHeader) > 1):
            raise Exception("More than one header: " + headerStruct)
        # print(matchHeader)
        col = matchHeader[0].mCell.column
        row = (matchHeader[0].mCell.row + matchHeader[0].mRowSize - 1) + rowOffset
        # print(row, col)
        # return self.mSheet.cell(row=row, column=col)
        cell = self.mSheet.cell(row=row, column=col)
        if cell.coordinate in self.mMergedDataCellInfo:
            return self.mMergedDataCellInfo[cell.coordinate].mTopLeftCell
        return cell            

    # Write to cell at header name and row offset.
    def Write(self, headerStruct: str, rowOffset: int, value: str):
        cell = self.GetCell(headerStruct, rowOffset)
        cell.value = value          

    def GetValue(self, cell):
        if cell.coordinate in self.mMergedDataCellInfo:
            return self.mMergedDataCellInfo[cell.coordinate].mTopLeftCell.value
        return cell.value            

    def Open(self, sheet: str, headerRange:str, readOnly=False):
        self.mHeaderRange = headerRange
        self.mWorkBook = openpyxl.load_workbook(self.mPath, readOnly)
        self.mSheet = self.mWorkBook[sheet]
        self.mDataColumnSize = Utils.GetDimension(headerRange)[1]
        self.__GetHeaderInfo()
        self.mPivotRow = openpyxl.utils.cell.coordinate_from_string(self.mHeaderRange.split(":")[1])[1]

    def Close(self):
        self.mWorkBook.save(self.mPath)

    def Read(self, rowOffset: int):
        returnValue = {}
        for i in range(self.mDataColumnSize):
            headerName = self.mHeaderInfoColumCache[i+1]
            # print(headerName)
            cell = self.mSheet.cell(row=self.mPivotRow+rowOffset,column=i+1)
            returnValue[headerName] = self.GetValue(cell)
        return returnValue
