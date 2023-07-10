from datetime import date
import datetime
import tkinter
from tkinter import *
import conexao

from models.Estacao1221 import Estacao1221
from models.Estacao1222 import Estacao1222
from models.Estacao1226 import Estacao1226
from models.Estacao1227 import Estacao1227
from models.Estacao1228 import Estacao1228
from models.Estacao1229 import Estacao1229
from models.Estacao1230 import Estacao1230
from models.Estacao1243 import Estacao1243
from models.Estacao1244 import Estacao1244
from models.Estacao1245 import Estacao1245
from models.Estacao1246 import Estacao1246
from models.Estacao1247 import Estacao1247

try:
    class Principal():

        # Seleção da estação de consulta
        def executa(self):
            data1 = self.entry2.get()
            data2 = self.entry3.get()
            Nome_Arquivo = self.entry4.get()
            Nome_Salvar = self.entry5.get()
            estacao = self.rb_value.get()
            resultado = str("Arquivo excel criado com sucesso!!!")
            self.v.set("")
            Window.update()

            # Validação de dados
            if ((data1 == "") or (data2 == "") or (Nome_Arquivo == "") or (Nome_Salvar == "") or (estacao == 0)):
                validacao = str("Preencha os campos corretamente!!!")
                self.v.set(validacao)
                Window.update()
            
            elif estacao == 1:  # Estação 1221
                objeto = Estacao1221()
                instancia = objeto._init_(
                    data1, data2, Nome_Arquivo, Nome_Salvar)
                self.v.set(resultado)

            elif estacao == 2:  # Estação 1222
                objeto = Estacao1222()
                instancia = objeto._init_(
                    data1, data2, Nome_Arquivo, Nome_Salvar)
                self.v.set(resultado)

            elif estacao == 3:  # Estação 1226
                objeto = Estacao1226()
                instancia = objeto._init_(
                    data1, data2, Nome_Arquivo, Nome_Salvar)
                self.v.set(resultado)

            elif estacao == 4:  # Estação 1227
                objeto = Estacao1227()
                instancia = objeto._init_(
                    data1, data2, Nome_Arquivo, Nome_Salvar)
                self.v.set(resultado)

            elif estacao == 5:  # Estação 1228
                objeto = Estacao1228()
                instancia = objeto._init_(
                    data1, data2, Nome_Arquivo, Nome_Salvar)
                self.v.set(resultado)

            elif estacao == 6:  # Estação 1229
                objeto = Estacao1229()
                instancia = objeto._init_(
                    data1, data2, Nome_Arquivo, Nome_Salvar)
                self.v.set(resultado)

            elif estacao == 7:  # Estação 1230
                objeto = Estacao1230()
                instancia = objeto._init_(
                    data1, data2, Nome_Arquivo, Nome_Salvar)
                self.v.set(resultado)

            elif estacao == 8:  # Estação 1243
                objeto = Estacao1243()
                instancia = objeto._init_(
                    data1, data2, Nome_Arquivo, Nome_Salvar)
                self.v.set(resultado)

            elif estacao == 9:  # Estação 1244
                objeto = Estacao1244()
                instancia = objeto._init_(
                    data1, data2, Nome_Arquivo, Nome_Salvar)
                self.v.set(resultado)

            elif estacao == 10:  # Estação 1245
                objeto = Estacao1245()
                instancia = objeto._init_(
                    data1, data2, Nome_Arquivo, Nome_Salvar)
                self.v.set(resultado)

            elif estacao == 11:  # Estação 1246
                objeto = Estacao1246()
                instancia = objeto._init_(
                    data1, data2, Nome_Arquivo, Nome_Salvar)
                self.v.set(resultado)

            elif estacao == 12:  # Estação 1247
                objeto = Estacao1247()
                instancia = objeto._init_(
                    data1, data2, Nome_Arquivo, Nome_Salvar)
                self.v.set(resultado)

        def limpar(self):
            self.entry2.delete(0, END)
            self.entry3.delete(0, END)
            self.entry4.delete(0, END)
            self.entry5.delete(0, END)
            self.rb_value.set(0)
            self.v.set("")
            self.entry2.focus()
            Window.update()

# Parâmetros da tela
        def __init__(self, instancia_de_tk):
            frame1 = Frame(instancia_de_tk)
            frame1.configure(border=5)
            frame1.pack()
            frame2 = Frame(instancia_de_tk)
            frame2.configure(border=5)
            frame2.pack()
            frame3 = Frame(instancia_de_tk)
            frame3.configure(border=5)
            frame3.pack()
            frame4 = Frame(instancia_de_tk)
            frame4.configure(border=5)
            frame4.pack()
            frame5 = Frame(instancia_de_tk)
            frame5.configure(border=5)
            frame5.pack()
            framecontainer = Frame(instancia_de_tk)
            framecontainer.configure()
            framecontainer.pack()
            frame6 = Frame(master=framecontainer)
            frame6.configure(border=5)
            frame6.pack(side="left")
            frame7 = Frame(master=framecontainer)
            frame7.configure(border=5)
            frame7.pack(side="left")
            frame8 = Frame(master=framecontainer)
            frame8.configure(border=5)
            frame8.pack(side="left")
            frame9 = Frame(instancia_de_tk)
            frame9.configure(border=5)
            frame9.pack()
            frame10 = Frame(instancia_de_tk)
            frame10.configure(border=5)
            frame10.pack()

# Parâmetros dos dados apresentados na tela
            label1 = Label(frame1, text="Formato de data: yyyy-mm-dd")
            label1.pack()

            label2 = Label(frame2, text="Insira a data inicio para consulta: ")
            label2.pack()
            self.entry2 = Entry(frame2)
            self.entry2.pack()

            label3 = Label(frame3, text="Insira a data final para consulta: ")
            label3.pack()
            self.entry3 = Entry(frame3)
            self.entry3.pack()

            label4 = Label(frame4, text="Insira o nome do arquivo existente: ")
            label4.pack()
            self.entry4 = Entry(frame4)
            self.entry4.pack()

            label5 = Label(
                frame5, text="Insira um nome para salvar o arquivo: ")
            label5.pack()
            self.entry5 = Entry(frame5)
            self.entry5.pack()

            self.rb_value = IntVar()
            self.rb1 = Radiobutton(
                frame6, text="Estação FLU 17", value=1, variable=self.rb_value).pack()
            self.rb2 = Radiobutton(
                frame6, text="Estação FLU 16", value=2, variable=self.rb_value).pack()
            self.rb3 = Radiobutton(
                frame6, text="Estação PN1-01", value=3, variable=self.rb_value).pack()
            self.rb4 = Radiobutton(
                frame6, text="Estação PSL-01", value=4, variable=self.rb_value).pack()
            self.rb5 = Radiobutton(
                frame7, text="Estação PSL-02", value=5, variable=self.rb_value).pack()
            self.rb6 = Radiobutton(
                frame7, text="Estação PSL-03", value=6, variable=self.rb_value).pack()
            self.rb7 = Radiobutton(
                frame7, text="Estação PSL-04", value=7, variable=self.rb_value).pack()
            self.rb8 = Radiobutton(
                frame7, text="Estação FLU 11", value=8, variable=self.rb_value).pack()
            self.rb9 = Radiobutton(
                frame8, text="Estação FLU 13", value=9, variable=self.rb_value).pack()
            self.rb10 = Radiobutton(
                frame8, text="Estação FLU 12", value=10, variable=self.rb_value).pack()
            self.rb11 = Radiobutton(
                frame8, text="Estação FLU 15", value=11, variable=self.rb_value).pack()
            self.rb12 = Radiobutton(
                frame8, text="Estação FLU 14", value=12, variable=self.rb_value).pack()

            label6 = Label(frame9, text="Resultado: ")
            label6.pack()
            self.v = StringVar()
            label6 = Label(frame9, text="", textvariable=self.v, background="white", font="14")
            label6.pack()

# Parâmetros de execução
            button1 = Button(frame10, text="Buscar", borderwidth=5, relief="raised", command=lambda: self.executa())
            button1.pack(side="right")

            button2 = Button(frame10, text="Limpar", borderwidth=5, relief="raised", command=lambda: self.limpar())
            button2.pack(side="left")

# Parâmetros da tela
    Window = tkinter.Tk()
    Window.title("Dados Estações Hidro")
    Window.minsize(width=350, height=450)
    Window.maxsize(width=450, height=600)
    Principal(Window)
    Window.mainloop()

# Encerrar conexao com o banco de dados
    if (conexao.con.is_connected()):
        conexao.cursor.close()
        conexao.con.close()
        print("Conexão ao MySQL encerrada.\n")

except OSError as e:
    print("Erro: ", e)
