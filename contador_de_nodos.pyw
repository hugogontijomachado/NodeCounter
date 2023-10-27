# -*- coding: UTF-8 -*-
from tkinter import *
import tkinter as tk
from tkinter import ttk


from tkinter import messagebox
from tkinter import filedialog

import threading
import json
import os


################ FrontEnd ##################
def check_and_install_libraries():
    try:        
        #importlib.import_module("pandas")
        #importlib.import_module("openpyxl")
        import pandas as pd
        import openpyxl
        root.build_main_frame()
    except ImportError:
        root.check_and_install()
       
        import subprocess
        #install_process = subprocess.Popen(["pip", "install", "pandas openpyxl"])
        subprocess.check_call(["pip", "install", "pandas", "openpyxl"])

        root.check_install_label.destroy()
        root.build_main_frame()
    

class CustomMenu(Menu):

    def __init__(self, parent, *args, **kwargs):
        Menu.__init__(self, parent, *args, **kwargs)
        self.parent = parent

        # Defina a cor de fundo do menu
        self.configure(bg=self.parent.bg)

        # Adicione os itens de menu
        self.add_command(label="Sobre", command=self.about, accelerator="")
        self.add_command(label="Como usar", command=self.how_to_use, accelerator="")
        self.add_separator()
        self.add_command(label="Sair", command=self.parent.quit)


    def about(self):
        t = Toplevel(self)
        #t = create_scrollable_toplevel()
        t.geometry("600x180")
        #t.resizable(False, False)
        t.configure(background=self.parent.bg)

        Label(t,
               text="""\nContador de nodos (2023)\nSoftware desenvolvido por Hugo G. Machado
                    \n\nApós processar seus espectros no GNPS (Global Natural Products Social Molecular Networking)\ne converter os dados no software Cytoscape, obtenha a contagem de classes, subclasses,\nclasses e parents de maneira rápida e prática.""",
                font=("Arial", "10"),
                bg=self.parent.bg
        ).pack()


    def how_to_use(self):
        basepath = os.path.dirname(os.path.abspath(__file__))        
        os.startfile(os.path.join(basepath, "manual.pdf"))
        

class Main(Tk):

    def __init__(self):
        super().__init__()
        self.title('Contador de Nodos')
        self.geometry('400x250+50+50')
        try:
            icon_path = os.path.dirname(os.path.abspath(__file__))
            self.wm_iconbitmap(os.path.join(icon_path, r'contador_de_nodos.ico'))
        except:
            pass


        #self.resizable(False, False)
        self.bg = "#dde"
        self.fg = "#002060"#"#0070C0"#"#2C48E5"
        self.font = ("Arial", "12")
        self.font_title = ("Arial", "18", "bold")
        self.configure(background=self.bg)
        
        ##### Menu
        # Crie um rótulo que abrirá o menu personalizado
        self.label = Label(self, text="Ajuda", bg=self.bg, cursor="hand2")
        self.label.pack(anchor='nw', padx=5)#, pady=10)  # Alinhe o rótulo no canto superior esquerdo

        # Configure um evento de clique para o rótulo
        self.label.bind("<Button-1>", self.open_menu)

        # Crie um Frame que pareça uma linha horizontal
        separator = Frame(self, height=1, bd=1, relief="ridge", bg="#cdd")
        separator.pack(fill="x")#, padx=10, pady=10)


        ##### Título
        Label(self, text='Contador de Nodos', font=self.font_title, bg=self.bg, fg=self.fg).pack(side=TOP, fill=X, padx=10, pady=15)


    def check_and_install(self):
        self.check_install_label = Label(self, text="Bibliotecas Pandas e OpenPyXL não encontradas.\n\nA instalação pode levar alguns minutos, aguarde...", font=self.font, bg=self.bg)
        self.check_install_label.pack(side=TOP, fill=X)#, padx=10, pady=30 )


    def build_main_frame(self):

        self.bt_open = Button(self, text="Abrir arquivo '.cyjs'", command=self.open_file, font=self.font)
        self.bt_open.pack(side=TOP, padx=10, pady=15)

        self.bt_save = Button(self, text="Salvar resultados da contagem", command=self.save_file, font=self.font, state=DISABLED)
        self.bt_save.pack(side=TOP, padx=10)##, pady=5)


    def open_menu(self, event):
        # Crie uma instância do menu personalizado e exiba-o
        menu = CustomMenu(self)
        
        # Remova a primeira opção do menu (a opção especial)
        menu.delete(0)
        
        # menu.post(self.button.winfo_rootx(), self.button.winfo_rooty() + self.button.winfo_height())
        menu.post(self.label.winfo_rootx(), self.label.winfo_rooty() + self.label.winfo_height())


    def open_file(self):
        filename = filedialog.askopenfilename()
        if not filename:
            return
        try:
            with open(filename) as f:
                data = json.load(f)
            nodes = data['elements']['nodes']
        except:
            messagebox.showerror('Erro', "Arquivo inválido. Para mais informações acesse o menu 'Ajuda'")
            self.bt_save.config(state=DISABLED)
            return
        
        self.classes_dict, self.subclasses_dict,  self.superclasses_dict, self.parents_dict = contar_nodos(data)
        messagebox.showinfo('Sucesso', "Contagem realizada com sucesso!")
        self.bt_save.config(state=NORMAL)


    def save_file(self):
        import pandas as pd
        
        classes_df = pd.DataFrame({
            "Classes": list(self.classes_dict.keys()),
            "Count": list(self.classes_dict.values())
        })

        sub_classes_df = pd.DataFrame({
            "Sub Classes": list(self.subclasses_dict.keys()),
            "Count": list(self.subclasses_dict.values())
        })

        super_classes_df = pd.DataFrame({
            "Super Classes": list(self.superclasses_dict.keys()),
            "Count": list(self.superclasses_dict.values())
        })

        parents_df = pd.DataFrame({
            "Super Classes": list(self.parents_dict.keys()),
            "Count": list(self.parents_dict.values())
        })

        # Save this two dataframes in one excel file in diferent sheets
        filename = filedialog.asksaveasfilename(defaultextension=".xlsx")
        if not filename:
            return

        with pd.ExcelWriter(filename) as writer:
            classes_df.to_excel(writer, sheet_name='Classes', index=False)
            sub_classes_df.to_excel(writer, sheet_name='Sub Classes', index=False)            
            super_classes_df.to_excel(writer, sheet_name='Super Classes', index=False)
            parents_df.to_excel(writer, sheet_name='Parents', index=False)

        messagebox.showinfo('Sucesso', "Arquivo salvo com sucesso!")


################ BackEnd ##################


def format_and_count(classes, c_dict):
    if classes[-1][:5] == ' and ':
        classes[-1] = classes[-1][5:]
    classes = [c.strip() for c in classes]
    for c in classes:
        c = c.capitalize()
        if c not in c_dict:
            c_dict[c] = 0
        c_dict[c] += 1


def contar_nodos(data):
    classes_dict = {}
    subclasses_dict = {}
    superclasses_dict = {}
    parent_dict = {}
    
    for node in data['elements']['nodes']:
        # if node['data']['CF_subclass'] == 'no matches':
        #     continue
        classes = node['data']['CF_class'].lower().split(',')
        subclasses = node['data']['CF_subclass'].lower().split(',')
        superclasses = node['data']['CF_superclass'].lower().split(',')
        parent = node['data']['CF_Dparent'].lower().split(',')
        
        format_and_count(subclasses, subclasses_dict)
        format_and_count(classes, classes_dict)
        format_and_count(superclasses, superclasses_dict)
        format_and_count(parent, parent_dict)

    return classes_dict, subclasses_dict, superclasses_dict, parent_dict



if __name__ == '__main__':
    root = Main()

    install_thread = threading.Thread(target=check_and_install_libraries)
    install_thread.start()

    root.mainloop()
    