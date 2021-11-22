import tkinter
from convertedor import convertedor_dexion
from tkinter import *
from tkinter import filedialog

app = Tk()
def open():
  app.filename = filedialog.askopenfilename()
  convertedor_dexion(app.filename)
  text['text'] = 'Arquivo convertido com sucesso'


app.title('Convertedor planilhas dexion')
app.geometry("300x100")


title = Label(app, text='Escolha um arquivo dexion para converter')
title.place(relx=0.5, rely=0.15,anchor=CENTER)

button = Button(app, text="Escolher Arquivo", command=open)
button.place(relx=0.5, rely=0.50, anchor=CENTER)

text = Label(app, text='')
text.place(relx=0.5, rely=0.80, anchor=CENTER)

app.mainloop()
