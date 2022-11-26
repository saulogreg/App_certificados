import fitz
from tkinter import filedialog
from tkinter import *
from tkinter import ttk
import sqlite3


root = Tk()


class app():
    def __init__(self):
        self.root = root
        self.conecta_db()
        self.criar_bd()
        self.TelaPrincipal()
        self.frames()
        self.botoes()
        self.Entrada_texto()
        self.func_limpa_tela()
        self.Frame_visualizacao()
        self.Visualizar_dados()
        self.Labels()
        root.mainloop()

    def TelaPrincipal(self):

        self.root.configure(background='#c0c0c0')
        # self.root.config(bg='#add123')
        # self.root.wm_attributes('-transparentcolor', '#add123')
        # self.root.attributes('-alpha', 0.5)
        self.root.title("App Anexar Certificado")
        self.root.geometry('1280x720')
        # self.image = Image.open(
        #     "estudo/appcertificado/Blueshift BackGroundpng.png")
        # self.imagepdf = Image.open(
        #     "estudo/appcertificado/pdfexemplo.png")
        # self.pdfimage = ImageTk.PhotoImage(self.imagepdf)
        # self.BackgroundBlue = ImageTk.PhotoImage(self.image)
        # ttk.Label(root, image=self.BackgroundBlue).pack(fill=BOTH, expand=True)

        self.root.resizable('true', 'true')  # ou true
        self.root.wm_maxsize(width=1920, height=1080)
        self.root.wm_minsize(width=640, height=480)

    def frames(self):

        # self.frame1 = Frame(self.root, bd=4, bg='#76877D',
        #                     highlightbackground='#82A6B1', highlightthickness=5)
        # ttk.Label(self.frame1, image=self.pdfimage).pack(fill=BOTH, expand=1)
        # ttk.Label(self.frame1, image=self.BackgroundBlue).pack()

        # self.frame1.place(relx=0.02, rely=0.02,
        #                   relwidth=0.46, relheight=0.46)
        self.frame2 = Frame(self.root, bd=4, bg='#76877D',
                            highlightbackground='#82A6B1', highlightthickness=5)
        # ttk.Label(self.frame2, image=self.BackgroundBlue).pack()
        self.frame2.place(relx=0.35, rely=0.07,
                          relwidth=0.32, relheight=0.12)
        self.textoframe1 = Text(self.frame2, width=640, height=480)
        self.textoframe1.place(relx=0.02, rely=0.02,
                               relwidth=0.96, relheight=0.96)
        self.frame3 = Frame(self.root, bd=4, bg='#76877D',
                            highlightbackground='#82A6B1', highlightthickness=5)
        # ttk.Label(self.frame3, image=self.BackgroundBlue).pack()
        self.frame3.place(relx=0.02, rely=0.20,
                          relwidth=0.96, relheight=0.75)

    def botoes(self):

        self.AnexarPdf = Button(
            self.root, text='Anexar PDF', bd=3, bg='#82A6B1', fg='black', font=("verdana", 8, 'bold'), command=lambda: [self.func_limpa_tela(), self.open_pdf()])
        self.AnexarPdf.place(relx=0.35, rely=0.01,
                             relwidth=0.1, relheight=0.04)

        self.AnexarPdf = Button(
            self.root, text='Enviar Dados', bd=3, bg='#008000', fg='black', font=("verdana", 8, 'bold'), command=lambda: [self.func_limpa_tela(), self.popular_tabela(), self.Visualizar_dados()])
        self.AnexarPdf.place(relx=0.57, rely=0.01,
                             relwidth=0.1, relheight=0.04)

        self.AnexarPdf = Button(
            self.root, text='Atualizar', bd=3, bg='#82A6B1', fg='black', font=("verdana", 8, 'bold'), command=lambda: [self.Visualizar_dados()])
        self.AnexarPdf.place(relx=0.46, rely=0.01,
                             relwidth=0.1, relheight=0.04)

        self.bt_buscar = Button(self.root, text='Buscar por Nome', bd=3,
                                bg='#82A6B1', fg='black', font=("verdana", 8, 'bold'), command=lambda: [self.Buscar_dados(), self.func_limpa_tela()])
        self.bt_buscar.place(relx=0.18, rely=0.15,
                             relwidth=0.1, relheight=0.04)

        self.bt_Remover = Button(self.root, text='Deletar selecionado', bd=3,
                                 bg='#a80000', fg='white', font=("verdana", 8, 'bold'), command=lambda: [self.selectedRow(), self.Remover_item(), self.Visualizar_dados()])
        self.bt_Remover.place(relx=0.70, rely=0.15,
                             relwidth=0.11, relheight=0.04)

    def Labels(self):
        self.buscar_label = Label(self.root, text='Busca Rápida', bd=3,
                                  bg='#caeaf5', font=("verdana", 8, 'bold'))
        self.buscar_label.place(relx=0.02, rely=0.11)

    def Entrada_texto(self):
        self.buscar_entry = Entry(self.root, bg='#82A6B1')
        self.buscar_entry.place(relx=0.02, rely=0.15,
                                relwidth=0.15, relheight=0.04)

    def open_pdf(self):
        self.file = filedialog.askopenfilename(title="Selecione seu PDF", filetype=(
            ("PDF    Files", "*.pdf"), ("All Files", "*.*")))

        if self.file:

            pdf_file = fitz.open(self.file)

            for pageNumber, page in enumerate(pdf_file.pages(), start=1):
                text = page.get_text()
            self.text = text

            self.lines = text.split('\n')

            # guardando o conteúdo extraído em uma variável

            if "Válido at" in text:

                self.nome = self.lines[1]
                self.tipocertificado = self.lines[4].replace(
                    'Microsoft Certified: ', "")
                self.datacertificado = self.lines[5].replace(
                    'Data da conquista: ', "")
                self.numerocertificado = self.lines[7].replace(
                    "Número da certificação: ", "")
                self.validade = self.lines[6].replace("Válido até: ", "")

                # adicionando o conteúdo extraído na tela
                self.textoframe1.insert(
                    1.0, self.nome + '\n' + self.tipocertificado + '\n' + self.datacertificado + '\n' + self.numerocertificado + '\n' + self.validade)

            elif len(text) < 300:

                self.nome = self.lines[1]
                self.tipocertificado = self.lines[4].replace(
                    'Microsoft Certified: ', "")
                self.datacertificado = self.lines[5].replace(
                    'Data da conquista: ', "")
                self.numerocertificado = self.lines[6].replace(
                    "Número de certificação: ", "").replace(
                    "Número da certificação: ", "")

                # adicionando o conteúdo extraído na tela
                self.textoframe1.insert(
                    1.0, self.nome + '\n' + self.tipocertificado + '\n' + self.datacertificado + '\n' + self.numerocertificado)

            elif "conseguiu a" in text:

                self.nome = self.lines[-6]
                self.tipocertificado = self.lines[-4].replace(
                    'Microsoft Certified: ', "")
                self.datacertificado = self.lines[-3].replace(
                    'Data da conquista: ', "")
                self.numerocertificado = self.lines[-2].replace(
                    "Número de certificação: ", "").replace(
                    "Número da certificação: ", "")

                # adicionando o conteúdo extraído na tela
                self.textoframe1.insert(
                    1.0, self.nome + '\n' + self.tipocertificado + '\n' + self.datacertificado + '\n' + self.numerocertificado)

            elif "http" in text:

                self.nome = self.lines[-7]
                self.tipocertificado = self.lines[-4].replace(
                    'Microsoft Certified: ', "")
                self.datacertificado = self.lines[-3].replace(
                    'Data da conquista: ', "")
                self.numerocertificado = self.lines[-2].replace(
                    "Número de certificação: ", "").replace(
                    "Número da certificação: ", "")

                # adicionando o conteúdo extraído na tela
                self.textoframe1.insert(
                    1.0, self.nome + '\n' + self.tipocertificado + '\n' + self.datacertificado + '\n' + self.numerocertificado)

            else:
                print("erro, contatar o adm")

    def conecta_db(self):
        # conectando...
        self.conn = sqlite3.connect('certificacoes.db')
        # definindo um cursor
        self.cursor = self.conn.cursor()

    def desconecta_db(self):
        self.conn.close()

    def criar_bd(self):

        try:
            self.conecta_db()

            self.cursor.execute("""
            CREATE TABLE trainee_certificados (
                    id INTEGER PRIMARY KEY,
                    nome TEXT NOT NULL,
                    Tipo_de_Certificacao INTEGER,
                    Data_da_Conquista TEXT NOT NULL,
                    Numero_da_Certificacao INTEGER,
                    Validade TEXT
                    
                    
            );
            """)
            print('Tabela criada com sucesso.')
            # desconectando...
            # self.desconecta_db()
        except:
            pass
    
    def Frame_visualizacao(self):

        self.visualizartabela = ttk.Notebook(self.frame3)
        self.visualizartabela.place(relx=0, rely=0, relwidth=1, relheight=1)

        self.listadados = ttk.Treeview(
            self.visualizartabela, height=3, column=("col1", "col2", "col3", "col4", "col5", "col6"))
        self.listadados.place(
            relx=0.02, rely=0.02, relwidth=0.96, relheight=0.96)
        self.listadados.heading("#0", text="")
        self.listadados.heading("#1", text="ID")
        self.listadados.heading("#2", text="Nome")
        self.listadados.heading("#3", text="Tipo de Certificação")
        self.listadados.heading("#4", text="Data da Conquista")
        self.listadados.heading("#5", text="Número da Certificação")
        self.listadados.heading("#6", text="Validade")

        self.listadados.column("#0", width=-200)
        self.listadados.column("#1", width=10)
        self.listadados.column("#2", width=100)
        self.listadados.column("#3", width=100)
        self.listadados.column("#4", width=100)
        self.listadados.column("#5", width=100)
        self.listadados.column("#6", width=100)
        

        self.scroll_painel = Scrollbar(self.frame3, orient=VERTICAL)
        self.listadados.configure(yscroll=self.scroll_painel.set)
        self.scroll_painel.place(
            relx=0.98, rely=0.02, relwidth=0.02, relheight=0.98)

    def selectedRow(self):
       
        selected_item = self.listadados.focus()
        item_details = self.listadados.item(selected_item)
        self.Id_button = item_details.get("values")[0]
        print(self.Id_button)
        print(item_details.get("values"))

    def Remover_item(self):
        self.conecta_db()
        selected_items = self.listadados.selection()
        for selected_item in selected_items:
            self.listadados.delete(selected_item)
            print(selected_item)
            
                 
        self.cursor.execute(f"""
            DELETE FROM trainee_certificados where id={self.Id_button}""")
        self.conn.commit()

    def Visualizar_dados(self):
        self.listadados.delete(*self.listadados.get_children())
        self.conecta_db()
        lista = self.cursor.execute(
            """SELECT * FROM trainee_certificados""")
        
        for i in lista:

            self.listadados.insert(
                "", END, values=(i[0], i[1], i[2], i[3], i[4], i[5]))

    def Buscar_dados(self):

        self.conecta_db()
        self.listadados.delete(*self.listadados.get_children())
        self.buscar_entry.insert(END, "%")
        nome = self.buscar_entry.get()

        self.cursor.execute(
            """ SELECT id, nome, Tipo_de_Certificacao, Data_da_Conquista, Numero_da_Certificacao, Validade from trainee_certificados where nome like '%s' order by nome asc """ % nome)

        buscar_nome = self.cursor.fetchall()

        for i in buscar_nome:
            self.listadados.insert("", END, values=(
                i[0], i[1], i[2], i[3], i[4], i[5]))

    def func_limpa_tela(self):

        self.textoframe1.delete("1.0", "end")
        self.buscar_entry.delete(0, END)

    def popular_tabela(self):

        self.conecta_db()

        # self.cursor.execute(F"""
        # INSERT INTO trainee_certificados VALUES ('{self.nome}','{self.tipocertificado}','{self.datacertificado}','{self.numerocertificado}')""")

        if "Válido at" in self.text:
            self.cursor.execute(F""" insert into trainee_certificados (nome, Tipo_de_Certificacao, Data_da_Conquista, Numero_da_Certificacao, Validade)
                                VALUES ('{self.nome}','{self.tipocertificado}','{self.datacertificado}','{self.numerocertificado}', '{self.validade}')""")
        else:
            self.cursor.execute(F""" insert into trainee_certificados (nome, Tipo_de_Certificacao, Data_da_Conquista, Numero_da_Certificacao, Validade)
                                VALUES ('{self.nome}','{self.tipocertificado}','{self.datacertificado}','{self.numerocertificado}','Não possui')""")

        self.conn.commit()

        print('Dados inseridos com sucesso.')


app()
