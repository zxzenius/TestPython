__author__ = 'zenius'

def qsort(sl, left, right):
    if left >= right:
        return
    p = sl[left]
    i, j = left, left + 1
    while(j <= right):
        if sl[j] < p:
            i = i + 1
            sl[i], sl[j] = sl[j], sl[i]
        j = j + 1
    sl[left], sl[i] = sl[i], sl[left]

    qsort(sl, left, i - 1)
    qsort(sl, i + 1, right)

if __name__ == '__main__':
    tlist = [3, 2, 5, 7, 3, 1, 4]
    qsort(tlist, 0, len(tlist) - 1)
    print(tlist)

