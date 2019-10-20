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
import Utils

class ExcelFile:

    mPath = ""
    mHeaderRange = ""           # Ex: "A1:N1"
    mWorkBook = None
    mSheet = None
    mHeaderList = {}       
    mDataColumnSize = 0                # number of data column
    mDataRowSize  = 0                 # number of data row
    mDictData = {}       
    mHeaderCellInfo = {}           # list of info of main cell of header
    mMergedDataCellInfo = {}
    mPivotRow = 0
    mHeaderInfoColumCache = {}


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

    def Read(self, rowOffset: int):
        returnValue = {}
        for i in range(self.mDataColumnSize):
            headerName = self.mHeaderInfoColumCache[i+1]
            # print(headerName)
            cell = self.mSheet.cell(row=self.mPivotRow+rowOffset,column=i+1)
            returnValue[headerName] = self.GetValue(cell)
        return returnValue
