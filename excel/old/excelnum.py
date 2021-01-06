def excelNum(string):
    num=0
    while len(string)>0:
        num=num+(ord(string.pop(0).upper())-64)*(26**(len(string)))
    return num
def strNum(string):
    num=0
    while len(string)>0:
        num=num+int(string.pop(0))*10**(len(string))
    return num
def cellRange(startCell,endCell):
    startRow=excelNum(list(filter(str.isalpha,startCell)))
    startCol=strNum(list(filter(str.isdigit,startCell)))
    endRow=excelNum(list(filter(str.isalpha,endCell)))
    endCol=strNum(list(filter(str.isdigit,endCell)))
    if startRow>endRow or startCol>endCol:
        print("error range")
        return "error range"
    else:
        return ([startRow,startCol,endRow,endCol])