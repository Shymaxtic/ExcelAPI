# Copyright (C) 2019 Shymaxtic
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

from CellInfo import CellInfo

D_SEPERATOR = ":"

class HeaderInfo:
    def __init__(self, cellInfo, parentHeaderInfo):
        self.mCellInfo = cellInfo
        self.mParenHeaderInfo = parentHeaderInfo
        self.mFullName = self.__GetFullHeaderName()

    def __eq__(self, other): 
        if not isinstance(other, HeaderInfo):
            return NotImplemented
        return self.mCellInfo == other.mCellInfo and \
                self.mParenHeaderInfo == other.mParenHeaderInfo and \
                self.mFullName == other.mFullName

    def __str__(self):        
        return "HeaderInfo " + self.mCellInfo.__str__()  + "::"  + self.mFullName 

    def __GetFullHeaderName(self):
        """
        Full name of a header name. Ex: "HeaderName:ParentHeaderName:TopLevelHeaderName:"
        """
        headerName = self.mCellInfo.mTopLeftCell.value
        parentName = ""
        parentInfo = self.mParenHeaderInfo
        while (parentInfo != None):
            parentName = parentName + D_SEPERATOR + parentInfo.mCellInfo.mTopLeftCell.value
            parentInfo = parentInfo.mParenHeaderInfo
        return headerName + parentName + D_SEPERATOR
