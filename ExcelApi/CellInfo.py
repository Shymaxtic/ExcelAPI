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
import Utils

class CellInfo:
    """
    Cell info for cell in merged cell area or single cell.
    """

    mCell = None
    mTopLeftCell = None
    mRowSize = 0
    mColumnSize = 0

    def __init__(self, cell : openpyxl.cell.Cell, topleftCell: openpyxl.cell.Cell, rowSize: int, colSize: int):
        self.mCell = cell                   # this cell
        self.mTopLeftCell = topleftCell     # top-left cell of merged area
        self.mRowSize, self.mColumnSize = rowSize, colSize


    def __eq__(self, other): 
        if not isinstance(other, CellInfo):
            return NotImplemented
        return self.mTopLeftCell == other.mTopLeftCell and \
                    self.mCell == other.mCell and \
                    self.mRowSize == other.mRowSize and \
                    self.mColumnSize == other.mColumnSize

    def __str__(self):        
        return self.mTopLeftCell.__str__() + "::" + self.mCell.__str__() +  "::" + str(self.mRowSize) + "::" + str(self.mColumnSize)

    def Value(self):
        """
        Get value of top-left cell if in merged cells
        """    
        return self.mTopLeftCell.value

    def RawValue(self):
        """
        Get value of this cell
        """
        return self.mCell.value        

    def GetRelativeCells(self):
        """
        Get list of merged cells if this cell is in a merged cell area, or single cell
        """
        sheet = self.mCell.parent
        for cellRange in sheet.merged_cells.ranges:
            if self.mTopLeftCell in cellRange:
                return sheet[cellRange]        
        return self.mCell

    def Has(self, cell: openpyxl.cell.Cell):
        pivotRow, pivotCol = self.mCell.row, self.mCell.column
        for row in range(self.mRowSize):
            for col in range(self.mColumnSize):
                if (cell == self.mCell.parent.cell(pivotRow + row, pivotCol + col)):
                    return True
        return False

    