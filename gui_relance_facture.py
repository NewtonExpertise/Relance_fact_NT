from tkinter import *
from tkinter.ttk import *
from tkinter import messagebox
from os import system
import webbrowser as wb

class Gui_relance_client():
    def __init__(self):
        super().__init__()
        # paramètre : 
        couleur = "#E4AB5B"

        # fenêtre
        self.window=Tk()
        menubar = Menu(self.window)
        menubar.add_cascade(label="Notice d'utilisation", command=self.help)
        self.window.title("Relance client")
        self.window.wm_attributes("-topmost", 1)
        self.window.geometry("1000x661")
        self.window.resizable(0,0)
        self.window.iconbitmap("img/iconRC.ico")
        self.window.config(background=couleur, menu=menubar)

        # background
        self.Background_window = PhotoImage(file = 'img/bg.png')
        self.ib = Label(self.window, background=couleur, image = self.Background_window)
        self.ib.place(x=-15,y=280)

        # gestion du style des widget
        style = Style()
        style.configure("BW.TLabel", foreground="white", background=couleur, font=("Arial",12, ) )
        style.configure("BC.TLabel", foreground="white", background=couleur, font=("Arial",9, ))
        style.configure("BF.TLabel", foreground="white", background=couleur, font=("Arial",10, 'bold'))
        style.configure("TButton", foreground="green", background="#ccc", padding=6)
        style.map("TButton", foreground=[('pressed',"red"),('active','blue')])
        style.configure("TSeparator",  background="orange", borderwidth=2)

        # barre de défilement horizontal : 
        xDefilB_relance_valide = Scrollbar(self.window, orient='horizontal')
        xDefilB_relance_valide.place(x=10, y=399, width = 300)

        xDefilB_relance_ignore = Scrollbar(self.window, orient='horizontal')
        xDefilB_relance_ignore.place(x=347, y=399, width =300)

        xDefilB_relance_invalide = Scrollbar(self.window, orient='horizontal')
        xDefilB_relance_invalide.place(x=720, y=318, width = 261)
    

        # Images
        self.Logo_envoi_mail = PhotoImage(file= 'img/send_mail.png')
        self.logo_excel = PhotoImage(file= 'img/excelbtn.png')

        # boutons
        self.envoie_mail = Button(self.window, text="Envoyer les relances", image = self.Logo_envoi_mail, compound = LEFT)
        self.envoie_mail.place(x=233 , y=426)

        self.xl_fact_relance =Button(self.window, text="Liste factures à relancer", image = self.logo_excel, compound = LEFT )
        self.xl_fact_relance.place(x=96 , y=554, width = 200)

        self.xl_all_dest =Button(self.window, text="Liste destinataires relances", image = self.logo_excel, compound = LEFT)
        self.xl_all_dest.place(x=392, y=554, width = 200)

        self.trie_relances_pretes =Button(self.window, text="Trier")
        self.trie_relances_pretes.place(x=224, y=46)

        self.trie_relances_annulees =Button(self.window, text="Trier")
        self.trie_relances_annulees.place(x=560, y=46)

        # Progress bar : 
        self.progressbar = Progressbar(self.window, orient="horizontal", length=636, mode = "determinate")
        self.progressbar.place_forget

        

        # Label : 
        lb_relance_ok = Label(self.window, text="Liste d'envoi :",style='BW.TLabel', justify= LEFT)
        lb_relance_ok.place(x=15 , y=51)

        lb_relance_annulee = Label(self.window, text="Liste d'envoi annulé :",style='BW.TLabel', justify= LEFT)
        lb_relance_annulee.place(x=355 , y=51)

        lb_relance_no_contact = Label(self.window, text='Destinataires relance à paramétrer \n dans Quadra GI :',style='BW.TLabel', justify= LEFT)
        lb_relance_no_contact.place(x=725 , y=341)

        lb_relance_adress = Label(self.window, text='Adresse mail à contrôler :',style='BW.TLabel', justify= LEFT)
        lb_relance_adress.place(x=725 , y=10)

        lb_relance_adress2 = Label(self.window, 
        text="Si l'adresse mail est valide cliquez sur la ligne\n pour la basculer dans la liste d'envoie.\n Sinon corrigez l'adresse mail dans Quadra GI.",
        style="BC.TLabel", justify= LEFT)
        lb_relance_adress2.place(x=725 , y=30, )
        

        lb_fleche_D = Label(self.window, text='>>>',style='BF.TLabel')
        lb_fleche_D.place(x=315 , y=203)

        lb_fleche_G = Label(self.window, text='<<<',style='BF.TLabel')
        lb_fleche_G.place(x=315 , y=238)


        # Liste : 
        self.liste_relance_valide = Listbox(self.window,
                                            selectbackground=couleur,
                                            selectforeground='white',
                                            highlightbackground=couleur,
                                            highlightcolor=couleur,
                                            xscrollcommand=xDefilB_relance_valide.set)
        self.liste_relance_valide.place(x = 10 , y=76, width=300 , height = 323)
        xDefilB_relance_valide['command'] = self.liste_relance_valide.xview
   
        
        self.liste_relance_ignore = Listbox(self.window,
                                            selectbackground=couleur,
                                            selectforeground='white',
                                            highlightbackground=couleur,
                                            highlightcolor=couleur,
                                            xscrollcommand=xDefilB_relance_ignore.set)
        self.liste_relance_ignore.place(x = 347 , y=76 ,width=300 , height = 323)
        xDefilB_relance_ignore['command'] = self.liste_relance_ignore.xview

        self.liste_relance_invalide = Listbox(self.window,
                                            selectbackground=couleur,
                                            selectforeground='white',
                                            highlightbackground=couleur,
                                            highlightcolor=couleur,
                                            xscrollcommand=xDefilB_relance_invalide.set)
        self.liste_relance_invalide.config(width = 43 , height = 15)
        self.liste_relance_invalide.place(x = 720 , y=76)
        xDefilB_relance_invalide['command'] = self.liste_relance_invalide.xview

        self.liste_relance_noDest = Listbox(self.window,
                                            selectbackground=couleur,
                                            selectforeground='white',
                                            highlightbackground=couleur,
                                            highlightcolor=couleur)
        self.liste_relance_noDest.config(width = 20 , height = 15)
        self.liste_relance_noDest.place(x = 780 , y=381)

        # Combobox
        self.combobox_base_de_donnee = Combobox(self.window,xscrollcommand=xDefilB_relance_ignore.set,state="readonly")
        self.combobox_base_de_donnee.place(x=10, y=10)
        # Separateur 
        separateur = Separator(self.window,  style='TSeparator', orient=VERTICAL)
        separateur.place(x=686, y=0, height=641 )

        separateur2 = Separator(self.window,  style='TSeparator', orient=HORIZONTAL)
        separateur2.place(x=0, y=526, width=686 )


        # Messagebox : 
    def message_box(self, nb_relance):
        res = messagebox.askokcancel("Compte-rendu de l'envoie : ", f"""Nous avons envoyé : {nb_relance} mails de relances.


Voulez-vous quitter l'application ?""")
        return res

    def message_box_noexcel(self, pathfile):
        messagebox.showerror(title='Erreur Excel non présent', message=f"Merci de prendre contact avec le service informatique\nLe fichier Excel : {pathfile} n'a pas été trouvé.")
        self.window.destroy()

    def show(self):
        self.window.mainloop()
        
    def help(self):
        wb.open_new(r'.\doc\relanceclient.pdf')


if __name__ == "__main__":
    view = Gui_relance_client()
    view.show()