from tkinter import *
from tkinter import PhotoImage as pimg


class App():
    def __init__(self):
        self.win = Tk()
        win = self.win
        img = pimg(file = 'img\\cart.png')

        Label(win, image = img).pack()
        win.mainloop()
App()