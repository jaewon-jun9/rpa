def excelNum(string):
    num=0
    for i,j in enumerate(string):
        if ord(j)>57:
            num=num+(ord(j.upper())-64)*(26**(len(string)-i-1))
        else:
            num=num+int(j)*10**(len(string)-i-1)
    return num
def excelString(string):
    row=excelNum(list(filter(str.isalpha,string)))
    col=excelNum(list(filter(str.isdigit,string)))
    return ([row,col])
def cellRange(startCell,endCell):
    start=excelString(startCell)
    end=excelString(endCell)
    if start[0]>end[0] or start[1]>end[1]:
        print("error range")
        return "error range"
    else:
        return (start+end)
def numToExNum(i,j):
    col=[]
    if i==26:
        col.insert(0,"Z")
    else:        
        while i>0:
            col.insert(0,chr(i%26+64))
            i=i//26
    j=str(j)
    col.append(j)
    excol=''.join(col)
    return (excol)
def appExcelList(num1,ws):
    for row in readCell(num1):
        ws.append(row)