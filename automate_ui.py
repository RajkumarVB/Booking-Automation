from tkinter import *

ws = Tk()
ws.title("Automation")
ws.geometry("500x300")


pawin = PanedWindow(orient ='vertical')
top = Button(pawin, text ="Panel 1")
top.pack(side = TOP)

pawin.add(top)

pawin.pack(fill = BOTH, expand = True)

pawin.configure(relief = RAISED)

ws.mainloop()