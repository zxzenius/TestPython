import _thread, threading, time


tester = 0

def subt(tid):
    global tester
    for i in range(1):
        tester+= 1
        time.sleep(5)
        tester+= 1
        #print('[%s]--->%s:%s'%(tid, i, tester))

def subt2():
    time.sleep(1)
        
if __name__ == '__main__':
    ths = []
    for i in range(100):
        #_thread.start_new_thread(subt, (i,))
        thread = threading.Thread(target=subt, args=(i,))
        thread.daemon = True
        thread.start()
    """for i in range(2):
        thread = threading.Thread(target=subt2)
        ths.append(thread)
        thread.start()
        #ths.append(thread)
    #for thread in ths: thread.join()
    for th in ths:
        th.join()"""

    print(tester)
                