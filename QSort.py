__author__ = 'zxz'

def PreSort(SList, Left, Right):
    if Left >= Right:
        return
    p, i, j = Left, Left + 1, Left + 1
    while (i <= Right):
        if SList[i] < SList[p]:
            SList[i], SList[j] = SList[j], SList[i]
            j = j + 1
        print(SList)
        i = i + 1
    SList[p], SList[j-1] = SList[j-1], SList[p]

if __name__ == '__main__':
    TestList = [5,3,2,6,3,9,0,6,7]
    PreSort(TestList, 0, len(TestList) - 1)
    print(TestList)
