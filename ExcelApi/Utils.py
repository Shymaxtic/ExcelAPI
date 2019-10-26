import openpyxl

def GetDimension(rangeString: str):
    """Get dimension of a range
    
    Arguments:
        rangeString {str} -- range in excel style. Ex: "A1:B2"
    
    Returns:
        int, int -- row size, column size
    """
    min_col, min_row, max_col, max_row = openpyxl.utils.range_boundaries(rangeString)
    return (max_row - min_row) + 1, (max_col - min_col) + 1

def IsCellRangeInCellRange(checkRange: str, targetRange: str):
    min_col, min_row, max_col, max_row = openpyxl.utils.range_boundaries(checkRange)
    tmin_col, tmin_row, tmax_col, tmax_row = openpyxl.utils.range_boundaries(targetRange)
    return tmin_col <= min_col and \
            tmax_col >= max_col and \
            tmin_row <= min_row and \
            tmax_row >= max_row

def IsCellInCellRange(cellCoor: str, targetRange: str):
    row, column = openpyxl.utils.coordinate_to_tuple(cellCoor)
    tmin_col, tmin_row, tmax_col, tmax_row = openpyxl.utils.range_boundaries(targetRange)
    return tmin_row <= row and \
            tmax_row >= row and \
            tmin_col <= column and \
            tmax_col >= column
