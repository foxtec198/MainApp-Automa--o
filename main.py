from sqlite3 import connect
from tkinter import *
from tkinter import ttk, messagebox, PhotoImage, filedialog
from time import strftime as st

from metas import Metas
from vj import VJ
from pa import PA
from hxh import mainHxh
from estore import es

import os
from functools import partial


class App():
    def __init__(self):
        # Vars
        self.win = Tk()
        self.conn = connect("db\\dados.db")
        self.c = self.conn.cursor()
        self.bg = "#333"
        self.nome = 'Undefined'
        self.dia = int(st("%d"))
        self.fls = ["L025","L051","L054","L056","L057","L059","L061","L062","L194","L201","L287","L306","L313","L316","L326","L391","L393"]
        self.metasFl = Metas()
        self.metaLoja = self.metasFl.soma
        self.fla = open('txt\\fl.txt').read()
        
        # callback's
        self.config_win()
        self.widget()
        self.inst()
        self.win.mainloop()

    def config_win(self):
        self.win.geometry("1300x680+0+0")
        self.win['bg'] = 'Black'
        self.win.iconbitmap("img/icon.ico")
        self.win.title("Automação Acompanhamento")
        self.win.resizable(width=False, height=False)

    def widget(self):
        win = self.win

        # Frames
        self.loginFrame = Frame(win, bg= self.bg, width= 210, height = 660)
        
        self.hxhFrame = Frame(win, bg = self.bg, width= 260, height = 80)
        self.paFrame = Frame(win, bg = self.bg, width=260, height= 80)
        self.vjFrame = Frame(win, bg = self.bg, width = 260, height = 80)
        self.esFrame = Frame(win, bg = self.bg, width = 250, height = 80)
        
        self.txtFrame = Frame(win, bg = self.bg, width= 550, height= 540)
        self.setFrame = Frame(win, bg = self.bg, width = 500, height = 540)

        # IMAGENS
        pimg = PhotoImage
        self.imglogin = pimg(file =r"img\\man.png")
        self.paImg = pimg(file = r"img\\pa.png")
        self.hxhImg = pimg(file = r"img\\hxh.png")
        self.vjImg = pimg(file = r"img\\vj.png")
        self.userImg = pimg(file = r"img\\user.png")
        self.userImg = pimg(file = r"img\\user.png")
        self.cartImg = pimg(file = r"img\\cart.png")
        self.upUser = pimg(file=r'img\\update.png')
        self.delUser = pimg(file=r'img\\del.png')
        self.darkImg = pimg(file=r'img\\dark.png')


        # DIREITOS AUTORAIS ----------------------------------------------------------------------------------------
        self.ms = 'Deselvolvido por TecnoBreve Enterprise, Direitos autorais reservados a mesma © Londrina 2022'
        Label(
            self.win,
            text = self.ms,
            font = 'Arial 8 italic',
            bg = 'black', fg = 'grey'
            ).place(x=690, y= 660, anchor = 'center')
        
        # Login Vermelho: d90429
        self.lblImgLogin = Label(
            self.loginFrame,
            image = self.imglogin,
            bg =self.bg, width= 150,
            height=150
            )

        self.lblUser = Label(
            self.loginFrame,
            text='👤 Undefined',
            bg = self.bg,
            fg = 'white',
            font = 'sansserif 12 bold'
            )

        self.lblHora = Label(
            self.loginFrame,
            font = "sansserif 12",
            bg = self.bg,
            fg = "white"
            )

        self.lblMeta = Label(
            self.loginFrame,
            text = f"Meta Loja: R$ {self.metaLoja:.2f}",
            font="Arial 12 bold",
            bg = self.bg,
            fg = 'white'
            )

        self.btnLogin = Button(
            self.loginFrame,
            text = "Login",
            width = 25,
            font = "monospace 16 bold",
            bg = "Black",
            fg = "ghostwhite",
            command = self.login
            )

        self.btnLogout = Button(
            self.loginFrame,
            text = "Logout",
            width = 25,
            font = "monospace 16",
            bg = "#d90429",
            fg = "white",
            command = self.logout
            )
        
        self.ttCombo = Button(
            self.loginFrame,
            text='Escolha sua Filial:',
            bg = '#d90429',
            activebackground = '#d90429',
            activeforeground='white',
            fg = 'white',
            borderwidth = 0,
            width = 200,
            font='Arial 12 bold',
            command = self.upfl
            )

        self.combo = ttk.Combobox(
            self.loginFrame,
            values=self.fls,
            width=32,
            ) 
        
        # TEXT FRAME
        self.txt = Text(
            self.txtFrame,
            bg = self.bg,
            fg = 'white',
            borderwidth = 3,
            height = 30,
            width=65
            )

        Label(
            self.txtFrame,
            text = '🖊 Cole o texto aqui:',
            font = 'Arial 14 bold',
            bg = self.bg,
            fg = "GhostWhite"
            ).place(x = 280, y = 15 , anchor='center' )

        # HXH
        self.btnHxh = Button(
            self.hxhFrame,
            text = 'Comercial',
            image = self.hxhImg,
            width= '64',
            bg = self.bg,
            borderwidth = 0,
            activebackground= self.bg,
            command = partial(self.hxh, 1)
            )
        
        self.lblHxh = Button(
            self.hxhFrame,
            text = 'Comercial',
            font = 'monospace 14 bold',
            borderwidth=0,
            activebackground = self.bg,
            bg = self.bg,
            fg = 'white',
            command = partial(self.hxh, 1)
            )
        
        # PA
        self.btnPA = Button(
            self.paFrame,
            text = 'Participalçao',
            image = self.paImg,
            width= '64', 
            bg = self.bg, 
            borderwidth = 0, 
            activebackground= self.bg, 
            command = self.pa
            )
            
        self.lblPA = Button(
            self.paFrame, 
            text = 'Participação', 
            font = 'monospace 14 bold',
            borderwidth=0,
            activebackground = self.bg,
            bg = self.bg, 
            fg = 'white',
            command = self.pa
            )
        
        # VJ
        self.btnVJ= Button(
            self.vjFrame, 
            text = 'VendaComJuros', 
            image = self.vjImg, 
            width= '64', 
            bg = self.bg, 
            borderwidth = 0, 
            activebackground= self.bg, 
            command = self.vj
            )

        self.lblVJ = Button(
            self.vjFrame, 
            text = 'Vendas c/ Juros', 
            font = 'monospace 14 bold',
            borderwidth=0,
            activebackground = self.bg, 
            bg = self.bg, 
            fg = 'white',
            command = self.vj 
            )
        
        # ESTORE
        self.btnCart = Button(
            self.esFrame, 
            text = 'E-Store', 
            image = self.cartImg, 
            width= '64', 
            bg = self.bg, 
            borderwidth = 0, 
            activebackground= self.bg, 
            command = partial(self.estore, 1)
            )

        self.lblEs = Button(
            self.esFrame, 
            text = 'E-store', 
            font = 'monospace 14 bold',
            borderwidth=0,
            activebackground = self.bg, 
            bg = self.bg, 
            fg = 'white',
            command = partial(self.estore, 1)
            )
        
        # SET FRAME
        self.ttAdd = Label(
            self.setFrame,
            text = 'Adicionar Usuário:',
            fg = 'White',
            bg = self.bg,
            font = 'Arial 12 bold'
        )
        self.btnAdd = Button(
            self.setFrame,
            image = self.userImg,
            bg = self.bg,
            activebackground = self.bg,
            borderwidth = 0,
            command = self.Add
        )

        self.ttUp = Label(
            self.setFrame,
            text = 'Atualizar Usuário:',
            fg = 'White',
            bg = self.bg,
            font = 'Arial 12 bold'
        )
        self.btnUp = Button(
            self.setFrame,
            image = self.upUser,
            bg = self.bg,
            activebackground = self.bg,
            borderwidth = 0,
            command = self.Update
        )

        self.ttDel = Label(
            self.setFrame,
            text = 'Excluir Usuário:',
            fg = 'White',
            bg = self.bg,
            font = 'Arial 12 bold'
        )
        self.btnDel = Button(
            self.setFrame,
            image = self.delUser,
            bg = self.bg,
            activebackground = self.bg,
            borderwidth = 0,
            command = self.Del
        )

        self.ttDark = Label(
            self.setFrame,
            text = 'Modo Claro:',
            fg = 'White',
            bg = self.bg,
            font = 'Arial 12 bold'
        )
        self.btnDark = Button(
            self.setFrame,
            image = self.darkImg,
            bg = self.bg,
            activebackground = self.bg,
            borderwidth = 0,
        )
        
        self.ttArq = Label(
            self.setFrame,
            text = '''
                Para Atualizar as Metas, Exporte no PowerBI com a 
                função "Dados Resumidos", salve como "metas.xlsx",
                e cole no diretório, caso necessário, substitua!!
            ''',
            bg = self.bg,
            fg = 'white',
            font = 'Arial 11 bold'
        )
        self.btnArq = Button(
            self.setFrame, 
            text = 'Choose File',
            command = self.dirMt
        )
        
        # UpDate and Bind
        self.win.bind('<F5>', self.hxh)
        self.win.bind('<F12>', self.estore)
        self.atualizarM()
        self.atualizar()

    def inst(self):
        # Frame
        self.loginFrame.place(x= 10, y = 10)
        
        self.hxhFrame.place(x= 230, y= 10)
        self.esFrame.place(x= 500, y = 10)
        self.vjFrame.place(x = 760, y = 10)
        self.paFrame.place(x=1030, y = 10)
        self.txtFrame.place(x = 230, y = 101 )
        
        self.setFrame.place(x = 790, y = 101)
        
        # Login
        self.lblImgLogin.place(x = 25, y = 10)  
        self.lblUser.place(x = 100, y = 160, anchor="center")
        self.lblHora.place(x = 100, y = 180, anchor="center")
        self.lblMeta.place(x= 100, y = 220, anchor="center")
        self.btnLogin.place(x = 100, y = 600, anchor="center")
        self.btnLogout.place(x = 100, y = 640, anchor="center")

        self.combo.set(self.fla)
        self.combo['state'] = 'readonly'
        self.ttCombo.place(x = 100, y = 255, anchor = 'center')
        self.combo.place(x = 105, y = 280, anchor = 'center')
        
        # Text
        self.txt.place(x=10, y=30)

        # DP
        self.btnHxh.place(x = 10, y = 10)
        self.lblHxh.place(x = 170, y = 40, anchor = 'center')
        self.btnPA.place(x = 10, y = 10)
        self.lblPA.place(x = 170, y = 40, anchor = 'center')
        self.btnVJ.place(x = 10, y = 10)
        self.lblVJ.place(x = 170, y = 40, anchor = 'center')
        self.btnCart.place(x = 10, y = 10)
        self.lblEs.place(x = 170, y = 40, anchor = 'center')
        
        # SET FRAME
        self.btnAdd.place(x = 125, y = 100, anchor = 'center')
        self.ttAdd.place(x = 125, y = 150, anchor = 'center')

        self.btnUp.place(x = 350, y = 100, anchor = 'center')
        self.ttUp.place(x = 350, y = 150, anchor = 'center')

        self.btnDel.place(x= 125, y= 300, anchor = 'center')
        self.ttDel.place(x= 125, y= 350, anchor = 'center')
        
        self.btnDark.place(x= 350, y= 300, anchor = 'center')
        self.ttDark.place(x= 350, y= 350, anchor = 'center')
        
        self.ttArq.place(x = 170, y = 500, anchor = 'center')
        self.btnArq.place(x = 450, y = 500, anchor = 'center')

    # Functions
    def msg(self, tp, msg):
        if tp == 1:
            messagebox.showinfo('Auto Acompanhamento', msg)
        elif tp == 2:
            messagebox.showerror('Auto Acompanhamento', msg)
        elif tp == 3:
            messagebox.showwarning('Auto Acompanhamento', msg)

    def atualizarM(self):
        self.metasFl = Metas
        self.metasFl.fl = self.combo.get()
        self.metasFl = Metas()
        self.metaLoja = self.metasFl.soma
        self.lblHora.after(3000, self.atualizarM)

    def atualizar(self):
        self.hora = st("%H:%M:%S - %d/%m/%Y")
        self.lblHora["text"] = self.hora
        self.lblMeta['text'] = f"Meta Loja: R$ {self.metaLoja:.2f}"
        self.lblHora.after(1000, self.atualizar)

    def login(self):
        win2 = Toplevel()       
        win2.geometry("250x250")
        win2['bg'] = '#333'
        win2.resizable(width=False, height=False)
        win2.title("LOGIN")
        
        def cons(event):
            self.userg = self.user.get()
            self.senhag = self.senha.get()
            user = self.c.execute(f"SELECT nome FROM user WHERE matricula = '{self.userg}' ").fetchone()
            senha = self.c.execute(f"SELECT senha FROM user WHERE matricula = '{self.userg}' ").fetchone()
            if senha != None:
                senha = senha[0]

            if self.senhag == senha:
                user = user[0]
                user = user.split()
                user = user[0]
                self.lblUser["text"] = f'👤 {user}'
                self.msg(1, "Login realizado com sucesso!")
                win2.destroy()
            else:
                self.msg(2, 'Senha ou Matricula Incorreta!')
            
            self.nome = user

        
        Label(win2, text = "🧑 Matricula:", bg = self.bg, fg="white", font = 'arial 12 bold').pack(pady = 5)
        self.user = Entry(win2, font = 'arial 12 bold', justify='center')
        self.user.pack(anchor="center")
        Label(win2, text= "🔑 Senha:", bg = self.bg, fg="white", font = 'arial 12 bold').pack(pady = 5)
        self.senha = Entry(win2, font = 'arial 12 bold', show = '*', justify='center')
        self.senha.pack(anchor="center")
        Button(win2, text = "Login", font = "arial 10", command= cons).pack(pady = 10)

        win2.bind('<Return>', cons)
        
    def logout(self):
        if self.nome != 'Undefined':
            self.lblUser['text'] = 'Undefined'
            self.nome = 'Undefined'
            self.msg(1, 'Lougout realizado com sucesso!!')
        else:
            self.msg(1, 'Realize Login !!!')

    def Add(self):
        def enviar(event):
            nomedb = nome.get()
            matdb = mat.get()
            senhadb = senha.get()
            print(nomedb, senhadb, matdb)
            if nomedb != '' and matdb != '' and senhadb != '':
                self.c.execute(f'insert into user(nome, matricula, senha) values ("{nomedb}","{matdb}","{senhadb}")')
                self.conn.commit()
                nomep = nomedb.split()
                self.msg(1, f'{nomep[0]} adicionado com sucesso!')
                add.destroy()
            else:
                self.msg(2, 'Dados Inválidos')
        
        add = Toplevel()
        if self.nome == 'Undefined':
            add.destroy()
            self.msg(3, 'Realize Login!')
        else:
            pass
        add.geometry('400x400')
        add.resizable(width=False, height=False)
        add['bg'] = self.bg

        tti = Label(
            add,
            image = self.userImg,
            borderwidth = 0,
            bg = self.bg,
            activebackground = self.bg
        )
        tt = Label(
            add,
            text='Adicionar novo Usuário!',
            bg = self.bg,
            fg = 'GhostWhite',
            font = 'Fantasy 15 bold'
        )
        nomelbl = Label(
            add,
            text = 'Nome:',
            bg = self.bg,
            fg = 'White',
            font = 'Arial 12 bold'
            )
        nome = Entry(
            add,
            bg = 'Black',
            fg = 'white',
            font = 'Arial 12 bold'
            )
        matlbl = Label(
            add,
            text = 'Matricula:',
            bg = self.bg,
            fg = 'White',
            font = 'Arial 12 bold'
            )
        mat = Entry(
            add,
            bg = 'Black',
            fg = 'white',
            font = 'Arial 12 bold'
            )
        senhalbl = Label(
            add,
            text = 'Senha:',
            bg = self.bg,
            fg = 'White',
            font = 'Arial 12 bold'
            )
        senha = Entry(
            add,
            bg = 'Black',
            fg = 'white',
            font = 'Arial 12 bold',
            show = '*'
            )
        btnEnviar = Button(
            add,
            text = 'Adicionar!',
            bg = '#d90429',
            fg = 'White',
            font ='Arial 14 bold',
            command = partial(enviar, 1)
        )
        
        tti.place(x = 60, y = 40, anchor = 'center')
        tt.place(x = 220, y = 40, anchor = 'center')
        nomelbl.place(x = 200, y = 110, anchor = 'center')
        nome.place(x = 200, y = 140, anchor = 'center')
        matlbl.place(x = 200, y = 170, anchor = 'center')
        mat.place(x = 200, y = 200, anchor = 'center')
        senhalbl.place(x = 200, y = 230, anchor = 'center')
        senha.place(x = 200, y = 260, anchor = 'center')
        btnEnviar.pack(fill='x', side = 'bottom')
        add.bind('<Return>', enviar)

    def dirMt(self):
        os.system('explorer metas')

    def upfl(self): 
        if self.nome != 'Undefined': 
            a = open('txt\\fl.txt', 'w')
            a.write(self.combo.get())
            a.close()
        else:
            self.msg(2, 'Realize Login!')

    def Update(self):
        def up(event):
            matdb = mat.get()
            senhadb = senha.get()
            nomedb = nome.get()
            if matdb != '':
                self.c.execute(f'UPDATE user SET nome = "{nomedb}", senha = "{senhadb}" WHERE matricula = "{matdb}" ')
                self.msg(1, 'Atualizado com sucesso!!')
        update = Toplevel()
        if self.nome == 'Undefined':
            update.destroy()
            self.msg(3, 'Realize Login!')
        else:
            pass
        update.geometry('400x400')
        update.resizable(width=False, height=False)
        update['bg'] = self.bg

        tti = Label(
            update,
            image = self.upUser,
            borderwidth = 0,
            bg = self.bg,
            activebackground = self.bg
        )
        tt = Label(
            update,
            text='Atualizar Usuário!',
            bg = self.bg,
            fg = 'GhostWhite',
            font = 'Fantasy 15 bold'
        )
        nomelbl = Label(
            update,
            text = 'Nome:',
            bg = self.bg,
            fg = 'White',
            font = 'Arial 12 bold'
            )
        nome = Entry(
            update,
            bg = 'Black',
            fg = 'white',
            font = 'Arial 12 bold'
            )
        matlbl = Label(
            update,
            text = 'Matricula(não pode ser alterada!):',
            bg = self.bg,
            fg = 'White',
            font = 'Arial 12 bold'
            )
        mat = Entry(
            update,
            bg = 'Black',
            fg = 'white',
            font = 'Arial 12 bold'
            )
        senhalbl = Label(
            update,
            text = 'Senha:',
            bg = self.bg,
            fg = 'White',
            font = 'Arial 12 bold'
            )
        senha = Entry(
            update,
            bg = 'Black',
            fg = 'white',
            font = 'Arial 12 bold',
            show = '*'
            )
        btnEnviar = Button(
            update,
            text = 'Atualizar!!',
            bg = '#d90429',
            fg = 'White',
            font ='Arial 14 bold',
            command = partial(up, 1)
        )

        tti.place(x = 60, y = 40, anchor = 'center')
        tt.place(x = 220, y = 40, anchor = 'center')
        nomelbl.place(x = 200, y = 110, anchor = 'center')
        nome.place(x = 200, y = 140, anchor = 'center')
        matlbl.place(x = 200, y = 170, anchor = 'center')
        mat.place(x = 200, y = 200, anchor = 'center')
        senhalbl.place(x = 200, y = 230, anchor = 'center')
        senha.place(x = 200, y = 260, anchor = 'center')
        btnEnviar.pack(fill='x', side = 'bottom')
        update.bind('<Return>', up)

    def Del(self):
        def deletar(event):
            matdb = mat.get()
            if matdb != '':
                self.c.execute(f'DELETE FROM user WHERE matricula = "{matdb}"')
                self.conn.commit()
                self.msg(1, 'Usuário removido com sucesso!! ')
                dell.destroy()
            else:
                self.msg(2, 'Dadso Inválidos!')

        dell = Toplevel()
        if self.nome == 'Undefined':
            dell.destroy()
            self.msg(3, 'Realize Login!')
        else:
            pass
        dell.geometry('400x400')
        dell.resizable(width=False, height=False)
        dell['bg'] = self.bg

        tti = Label(
            dell,
            image = self.delUser,
            borderwidth = 0,
            bg = self.bg,
            activebackground = self.bg
        )
        tt = Label(
            dell,
            text='Deletar Usuário!',
            bg = self.bg,
            fg = 'GhostWhite',
            font = 'Fantasy 15 bold'
        )

        matlbl = Label(
            dell,
            text = 'Matricula:',
            bg = self.bg,
            fg = 'White',
            font = 'Arial 12 bold'
            )
        mat = Entry(
            dell,
            bg = 'Black',
            fg = 'white',
            font = 'Arial 12 bold'
            )
        btnEnviar = Button(
            dell,
            text = 'Deletar!!',
            bg = '#d90429',
            fg = 'White',
            font ='Arial 14 bold',
            command = partial(deletar, 1)
        )

        dell.bind('<Return>', deletar)
        tti.place(x = 60, y = 40, anchor = 'center')
        tt.place(x = 220, y = 40, anchor = 'center')
        matlbl.place(x = 200, y = 170, anchor = 'center')
        mat.place(x = 200, y = 200, anchor = 'center')
        btnEnviar.pack(fill='x', side = 'bottom')
#========================CALL=BACK=============================

    def hxh(self, event):
        txt = self.txt.get('1.0', END)
        arquivo = open('texto.txt', 'w')
        arquivo.write(txt)
        arquivo.close()
        
        self.msg(1, 'Realizando parcial')
        mainHxh()
        os.system('cd pln && horaxhora.xlsx')

    def pa(self):
        txt = self.txt.get('1.0', END)
        arquivo = open('clb.txt', 'w')
        arquivo.write(txt)
        arquivo.close()
        
        PA().cons()
        self.msg(1, 'Realizando Parcial')
        os.system('cd pln &&parcial.xlsx')

    def vj(self):
        txt = self.txt.get('1.0', END)
        arquivo = open('vj.txt', 'w')
        arquivo.write(txt)
        arquivo.close()
        VJ().code()
        self.msg(1, 'Realizando Parcial')
        os.system('cd pln && vj.xlsx')
        
    def estore(self, event):
        if self.nome != 'Undefined':
            self.msg(1, 'Abrindo E-Store')
            os.system('cd pln && estore.xlsx')
            es().iniciar()
        else:
            self.msg(1,'Realize Login!')

App()
