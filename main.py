import xlrd
import xlwt

def getTable(file, sheet_name):
    workbook = xlrd.open_workbook(file)
    worksheet = workbook.sheet_by_name(sheet_name)
    table = []
    index = 0
    for row in worksheet.get_rows():
        table.append([])
        for cell in row:
            table[index].append(cell.value)
        index += 1
    return table

def doIttMethod(table, startGameNumber, E):
    iterationsTable = []
    currentBrow = []
    currentArow = []
    sumBrow = []
    sumArow = []
    ittIndex = 0;
    ittNumber = 1
    iCol = startGameNumber
    jCol = 0
    curA = 0
    curB = 0;
    v_down = 0
    v_avg = 1
    v_up = 0
    cur_v = 0
    changeFlag = True
    for j in range(len(table)):
        sumBrow.append(0);

    for j in range(len(table[0])):
        sumArow.append(0);

    while (abs(v_avg - cur_v)) > E:
        print(ittNumber)
        if(changeFlag):
            v_avg = 0
            changeFlag = False
        cur_v = v_avg
        iterationsTable.append([])
        for j in range(len(table)):
            currentBrow.append(table[iCol][j]);

        for j in range(len(sumBrow)):
            sumBrow[j] += currentBrow[j]

        lst_num_B = list(enumerate(sumBrow, 0))
        min_num_B = min(lst_num_B, key=lambda i: i[1])
        jCol = min_num_B[0]
        for j in range(len(table[0])):
            currentArow.append(table[j][jCol]);

        for j in range(len(sumArow)):
            sumArow[j] += currentArow[j]

        iterationsTable[ittIndex].append(ittNumber)
        iterationsTable[ittIndex].append(iCol+1)

        lst_num_A = list(enumerate(sumArow, 0))
        max_num_A = max(lst_num_A, key=lambda i: i[1])
        iCol = max_num_A[0]

        curA = max_num_A[1] / ittNumber
        curB = min_num_B[1] / ittNumber

        v_down = curB
        v_avg = (curA + curB) / 2
        v_up = curA

        for j in range(len(sumBrow)):
            iterationsTable[ittIndex].append(sumBrow[j])
        iterationsTable[ittIndex].append(jCol+1)
        for j in range(len(sumArow)):
            iterationsTable[ittIndex].append(sumArow[j])
        iterationsTable[ittIndex].append(v_down)
        iterationsTable[ittIndex].append(v_avg)
        iterationsTable[ittIndex].append((v_up))

        ittIndex += 1
        ittNumber += 1

        currentArow.clear()
        currentBrow.clear()

    return iterationsTable


def writeToXls(book, table, nameIndex):
    sheet1 = book.add_sheet("Sheet"+str(nameIndex))
    for i in range(len(table)):
        for j in range(len(table[0])):
            sheet1.write(i, j, table[i][j])

if __name__ == '__main__':
    book = xlwt.Workbook(encoding="utf-8")
    ittTable = getTable("itt_table.xls", "Sheet1")
    wrTable = doIttMethod(ittTable, 0, 0.1)
    for row in wrTable:
        print(row)
    writeToXls(book, wrTable, 2)
    book.save("result.xls")
    pass


def thisNewFunction:
    pass