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
def appExcelList(num1,ws,load_ws):
    for row in readCell(num1,load_ws):
        ws.append(row)
def readCell(addList,load_ws):
    all_values=[]
    for row in range(addList[1],addList[3]+1): #row 단위로 읽기
        row_value=[] #row 마다 셀의 값을 저장
        for cell in range(addList[0],addList[2]+1): #row 마다 cell 단위로 데이터 읽기
            row_value.append(load_ws.cell(row,cell).value) #row_value에 cell의 내용을 넣어주어 row를 완성
        all_values.append(row_value) #row를 전체 리스트에 추가    
    return all_values