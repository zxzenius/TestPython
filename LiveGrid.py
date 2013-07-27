from tkinter import *
#import pprint
import random
import time

class livegrid:
    def __init__(self, canvas, col = 8, row = 8):
        #self.grid = []
        self.col = col
        self.row = row
        self.canvas = canvas
        self.reset()
        
    def reset(self):
        self.grid = []
        self.nextgd = []
        for x in range(self.col):
            ylist = []
            for y in range(self.row):
                flag = random.randrange(2)
                ylist.append([self.drawgrid(x, y, flag),flag])
                #self.drawgrid(x, y, flag)
            self.grid.append(ylist)
                #self.grid.append([random.randrange(2) for y in range(self.row)])
        self.nextgd = self.grid[:]
        #print(self.grid)
                
    def drawgrid(self, x, y, flag):
        i=x*8
        j=y*8
        if flag:
            color='green'
        else:
            color='gray'
        return(self.canvas.create_rectangle(i+8,j+8,i+16,j+16, width=0, fill=color))
    
    def fillgrid(self, x, y, flag):
        if flag:
            color='green'
        else:
            color='gray'
        self.canvas.itemconfig(self.grid[x][y][0], fill=color)
        #print(self.grid[x][y][0], color)
    
    def setstat(self, x, y, flag):
        if self.nextgd[x][y][1] == flag:
            pass
        else:
            self.nextgd[x][y][1] = flag
            self.fillgrid(x, y, flag)
    
    def judge(self, x, y):        
        fillcounter = self.adjacency(x, y)              
        if fillcounter == 2:             
            pass
        elif fillcounter == 3:
            self.setstat(x, y, 1)
        else:
            self.setstat(x, y, 0)
        
            
    def steprun(self):
        for x in range(self.col):
            for y in range(self.row):
                self.judge(x, y)
        self.grid = self.nextgd[:]    
                
    def run(self):
        while True:
            self.steprun()
            time.sleep(2)
                     
            
             
    def adjacency(self, x, y):
        alive = 0
        for i in (-1, 0, 1):
            for j in (-1, 0, 1):
                if i or j:
                    try: 
                        alive += self.grid[x+i][y+j][1]                 
                        #yield(self.grid[x+i][y+j])                        
                    except:
                        continue                   
        return(alive)
        
        

if __name__ == '__main__':
    #for i in range(10):
    #    start(4,3)
    #print(random.randrange(2))
    #drawgrid()
    canvas = Canvas(bg='gray')
    canvas.pack(side=TOP, expand=YES, fill=BOTH)    
    grid = livegrid(canvas, 64, 64)
    frame = Frame()
    frame.pack(side=BOTTOM)
    button = Button(frame, text='start',command=grid.steprun)
    button.pack(side=LEFT)
    button2 = Button(frame, text='show', command=(lambda:print(grid.grid)))
    button2.pack(side=RIGHT)
    mainloop()