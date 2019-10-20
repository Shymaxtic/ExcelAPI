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
    mDataCellInfo = {}


    def __init__(self, path: str):
        self.mPath = path     

    def __PostProcessMergedCell(self):
        self.mDataCellInfo = {}
        self.mHeaderCellInfo = {}
        for cellRange in self.mSheet.merged_cells.ranges:
            # get column size
            colNum = len(self.mSheet[cellRange.__str__()][0])
            # get row size
            rowNum = len(self.mSheet[cellRange.__str__()])
            # get top-left cell
            topLeftCell = self.mSheet[cellRange.__str__()][0][0]
            # if cell range is in header area
            if Utils.IsCellRangeInCellRange(cellRange.__str__(), self.mHeaderRange):
                # add to dictionary
                self.mHeaderCellInfo[topLeftCell.coordinate] = CellInfo(topLeftCell, topLeftCell, rowNum, colNum)
            # if cell range is in data area
            else:
                self.mDataCellInfo[topLeftCell.coordinate] = CellInfo(topLeftCell, topLeftCell, rowNum, colNum)
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
        for key in sortedKeys:
            # prepare parent header
            parentHeader = None
            icell = self.mHeaderCellInfo[key].mCell
            if (icell.row > 1):
                # if upper cell (merged cells/single cell) has value. It is parent of this cell
                upperCell = self.mSheet.cell(row=icell.row - 1, column=icell.column)
                # Check if upper cell is in any created header info
                for headerInfo in self.mHeaderList.values():
                    if headerInfo.mCellInfo.Has(upperCell):
                        parentHeader = headerInfo
                        newHeaderInfo = HeaderInfo(self.mHeaderCellInfo[key], parentHeader)
                        self.mHeaderList[newHeaderInfo.mFullName] = newHeaderInfo
                        break
            else:
                newHeaderInfo = HeaderInfo(self.mHeaderCellInfo[key], None) 
                self.mHeaderList[newHeaderInfo.mFullName] = (newHeaderInfo)
        for key in self.mHeaderList:
            print(key, ":::", self.mHeaderList[key])
                        
    def Open(self, sheet: str, headerRange:str, readOnly=False):
        self.mHeaderRange = headerRange
        self.mWorkBook = openpyxl.load_workbook(self.mPath, readOnly)
        self.mSheet = self.mWorkBook[sheet]
        self.mDataColumnSize = Utils.GetDimension(headerRange)[1]
        self.__GetHeaderInfo()

    # def Read(self, rowOffset: int):
