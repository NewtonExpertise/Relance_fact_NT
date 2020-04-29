
from tkinter import END, DISABLED, NORMAL
from outil_excel import Outil_Excel
from Query_releve_compte_contact import get_datas_relance, get_ecritures_non_lettrees, Query_all_client_for_relance
import jinja2
from mail import mail
from datetime import datetime
from DAO_PostgreSQL import BDDPostgreSQL
from postgre_rqt import espion_postgre, UtilisationTotal
from os import getlogin , getcwd, path
import time 
import sys


class Controler_relance_client():

    def __init__(self, view):
        # Création de la fram
        self.fram = view
        # Set du chemin mennant aux factures à joindres.
        self.path_BDD_GI = r'\\srvquadra\Qappli\Quadra\DATABASE\gi\0000\Qgi.mdb'
        self.path_BDD_Audit = r'\\srvquadra\Qappli\Quadra\DATABASE\cpta\DC\000500\QCompta.mdb'
        self.path_PJ_Audit = r'\\srvquadra\Qappli\Quadra\DATABASE\cpta\DC\000500\Images'
        self.path_BDD_Expertise = r'\\srvquadra\Qappli\Quadra\DATABASE\cpta\DC\000101\QCompta.mdb'
        self.path_PJ_Expertise = r'\\srvquadra\Qappli\Quadra\DATABASE\cpta\DC\000101\Images'
        self.NEWTON_IBAN_AUDIT = 'FR76 3007 6024 4123 0865 0020 089'
        self.NEWTON_BIC_AUDIT = 'NORDFRPP'
        self.NEWTON_IBAN_EXPERTISE = 'FR76 3007 6024 4122 3451 0020 080'
        self.NEWTON_BIC_EXPERTISE = 'NORDFRPP'
        self.template_basic_expertise = 'template/template_brut_expertise.txt'
        self.template_basic_audit = 'template/template_brut_audit.txt'
        self.excel_expertise = "excel/Relance_clients_expertise.xlsx"
        self.excel_audit = "excel/Relance_clients_audit.xlsx"

        # Variable a set au moment du choix de dossier
        self.NEWTON_ENTITY = ''
        self.NEWTON_ACTIVE_IBAN = ''
        self.NEWTON_ACTIVE_BIC = ''
        self.template_basic = ''
        self.active_BDD_Newton = ''
        self.active_PJ_NewTon = ''
        self.xlwings_wb = ''
        self.outil_excel= 0
        self.path_wb = []
        self.active_path_wb = ''
        self.stop_process=''

        # Déclencheurs et appels aux fonctions
        self.fram.liste_relance_valide.bind('<ButtonRelease-1>', self.on_select_relance_valide)
        self.fram.liste_relance_ignore.bind('<ButtonRelease-1>', self.on_select_relance_ignore)
        self.fram.liste_relance_invalide.bind('<ButtonRelease-1>', self.on_select_relance_invalide)
        self.fram.envoie_mail.bind('<ButtonRelease-1>', self.on_select_envoie_mail)
        self.fram.trie_relances_annulees.bind('<ButtonRelease-1>', self.tri_liste_annulees)
        self.fram.trie_relances_pretes.bind('<ButtonRelease-1>', self.tri_liste_pretes)
        self.fram.combobox_base_de_donnee.bind("<<ComboboxSelected>>", self.Select_entity)
        self.fram.window.protocol('WM_DELETE_WINDOW', self.quitte_application)


        # Set combobox
        self.fram.combobox_base_de_donnee["values"]= ["Newton Expertise","Newton Audit"]
        # Rendre le bouton excel liste des factures dû grisé
        self.fram.xl_fact_relance.config(state=DISABLED)
        self.fram.xl_all_dest.config(state=DISABLED)


    def Select_entity(self, evt):
        """
        Définit la base de donnée à intéroger selon si l'utilisateur sélectionne Audit ou Expertise.
        """

        Newton_entity = self.fram.combobox_base_de_donnee.get()
        if Newton_entity == "Newton Expertise":
            self.active_BDD_Newton = self.path_BDD_Expertise 
            self.active_PJ_NewTon = self.path_PJ_Expertise 
            self.template_basic = self.template_basic_expertise 
            self.NEWTON_ENTITY = 'Expertise'
            self.NEWTON_ACTIVE_IBAN = self.NEWTON_IBAN_EXPERTISE
            self.NEWTON_ACTIVE_BIC = self.NEWTON_BIC_EXPERTISE
            self.xlwings_wb =  self.excel_expertise
            self.active_path_wb = self.excel_expertise

        elif Newton_entity == "Newton Audit":
            self.active_BDD_Newton = self.path_BDD_Audit
            self.active_PJ_NewTon = self.path_PJ_Audit
            self.template_basic = self.template_basic_audit
            self.NEWTON_ENTITY = 'AUDIT SARL'
            self.NEWTON_ACTIVE_IBAN = self.NEWTON_IBAN_AUDIT
            self.NEWTON_ACTIVE_BIC = self.NEWTON_BIC_AUDIT
            self.xlwings_wb =  self.excel_audit
            self.active_path_wb =  self.excel_expertise

        else:
            x='Merci de prendre contact avec le service informatique'
            self.fram.liste_relance_valide.insert(END, x)

        self.fram.xl_fact_relance.config(state=NORMAL)
        self.fram.xl_all_dest.config(state=NORMAL)
        self.set_list_fram()
        self.fram.progressbar.place_forget()
        self.fram.envoie_mail.place(x=233 , y=426)
        self.fram.xl_fact_relance.bind('<ButtonRelease-1>', self.on_select_xl_fact_relance)
        self.fram.xl_all_dest.bind('<ButtonRelease-1>', self.on_select_xl_all_dest)

    def set_list_fram(self):
        """
        Alimente l'IHM
        """
        self.fram.liste_relance_valide.delete(0,END)
        self.fram.liste_relance_ignore.delete(0,END)
        self.fram.liste_relance_invalide.delete(0,END)
        self.fram.liste_relance_noDest.delete(0,END)
        all_client = Query_all_client_for_relance(self.path_BDD_GI)
        ecriture = get_ecritures_non_lettrees(self.active_BDD_Newton)

        # Récupération des relevés de compte clients à relancer.
        self.clients_relance = get_datas_relance(all_client,ecriture)
        # # --
        # Alimentation des listes de l'IHM.
        for compte_client, list_datas_relance in self.clients_relance.items():
            for data_relance in list_datas_relance:
            # - Liste des relances valides:
                if data_relance['contact']['Statut_mail']:
                    ligne_contact_valide = f"{compte_client} - {data_relance['contact']['RaisonSocial']} - {data_relance['contact']['Email']}"
                    self.fram.liste_relance_valide.insert(END, ligne_contact_valide)
            # - Liste des relances, aux mails non conforme: 
                if data_relance['contact']['Statut_mail']==0:
                    ligne_contact_invalide = f"{compte_client} - {data_relance['contact']['RaisonSocial']} - {data_relance['contact']['Email']}"
                    self.fram.liste_relance_invalide.insert(END, ligne_contact_invalide)
            # - Liste dossier n'ayant pas de contact désigné en tant que destinataire: 
                if data_relance['contact']['Statut_mail']==None:
                    self.fram.liste_relance_noDest.insert(END, compte_client)

        try:
            self.outil_excel = Outil_Excel(self.xlwings_wb)
        except:
            self.no_excel()

        if Outil_Excel(self.xlwings_wb).wb.fullname not in self.path_wb:
            self.path_wb.append(Outil_Excel(self.xlwings_wb).wb.fullname)
        self.active_path_wb = Outil_Excel(self.xlwings_wb).wb.fullname
        # Set niveaux de relance des factures envoyées des mois précédents.
        self.niveau_relance_factures = self.outil_excel.historique_relance()

        # Appels aux fonctions de trie des listes par numéo de compte client
        self.tri_nodest()
        self.tri_mail_ok()

    def on_select_relance_valide(self,evt):
        """
        Bascule l'item cliqué de la liste relance valide à la  liste relance annulée
        """
        index = self.fram.liste_relance_valide.curselection()
        value = self.fram.liste_relance_valide.get(index)
        
        self.fram.liste_relance_valide.delete(index)
        self.fram.liste_relance_ignore.insert(END,value)
      
    def on_select_relance_ignore(self,evt):
        """
        Bascule l'item cliqué de la liste relance annulées à la  liste relance valide
        """
        index = self.fram.liste_relance_ignore.curselection()
        value = self.fram.liste_relance_ignore.get(index)

        self.fram.liste_relance_ignore.delete(index)
        self.fram.liste_relance_valide.insert(END, value)

    def on_select_relance_invalide(self, evt):
        """
        Bascule l'item cliqué de la liste email invalide à la  liste relance valide
        """
        index = self.fram.liste_relance_invalide.curselection()
        value = self.fram.liste_relance_invalide.get(index)

        self.fram.liste_relance_invalide.delete(index)
        self.fram.liste_relance_valide.insert(END, value)


    def on_select_xl_fact_relance(self, evt):
        """
        Méthode permettant d'écrire dans l'onglet Facture_dû du tableau excel
        """
        Outil_Excel(self.xlwings_wb).fact_du(get_ecritures_non_lettrees(self.active_BDD_Newton))
        
    def on_select_xl_all_dest(self, evt):
        """
        Méthode permettant d'écrire dans l'onglet No_destinataire
        """
        Outil_Excel(self.xlwings_wb).No_mail_dest(Query_all_client_for_relance(self.path_BDD_GI))

    def on_select_envoie_mail(self, evt):
        """
        Répond au clique d'envoie de mail, set le boutton et la progress bar
        enregistre les donnnées sur espion de posgreSQL
        """

        self.stop_process=0
        # bouton d'envoie masqué
        self.fram.envoie_mail.place_forget()

        # mise en place de la progressbar
        self.fram.progressbar.place(x=10 , y=461)
        self.fram.progressbar['value'] = 0
        self.fram.progressbar['maximum'] = self.fram.liste_relance_valide.size()

        #envoie des mails
        relance, factures = self.envoie_des_relance()
        # affichage message_box de fin
        answer = self.fram.message_box(relance)
        self.quitte_application(boolean=answer)

        try:
            with BDDPostgreSQL() as PGSQL:
                
                sql = espion_postgre(getlogin(), datetime.now().strftime('%Y-%m-%d %H:%M:%S'), path.basename(path.dirname(self.active_BDD_Newton)),relance,factures )
                PGSQL.execute(sql)
        except:
            pass

    def envoie_des_relance(self):
        """
        Set des variables de la classe mail pour envoie des relances sélectionné dans la liste "Liste relance prêtes"
        de l'IHM
        """
        # compteur permettant un retour des envoies effectués à l'utilisateur
        nb_relance = 0
        nb_factures = 0
        # Récupération des comptes clients de la liste "Liste relance prêtes"
        # et alimentation de liste_relance_du_mois.
        liste_relance_du_mois =[]
        [liste_relance_du_mois.append(string.split(' - ')[0]) for string in self.fram.liste_relance_valide.get(0,END)]
        # Liste permettant d'alimenter l'onglet Relance du fichier excel après chaque envoie de mail.
        
        
        # Pour chaque relance client souhaité
        for compte_client in liste_relance_du_mois:
            # Création de l'objet mail
            set_mail = mail()
            compte_rendu_envoie_mail = [] # par mail écrit dans excel
            all_datas_template={} # dictionnaire qui contiendra les datas qui serviront à alimenter le template HTML.
            data_template_corps_tableau = [] # Liste destiné à alimenter le tableau dans le template HTML

            # Pour un client à relancé
            for data_relance in self.clients_relance[compte_client]:
                #Pour chaques factures
                niveau_relance=''
                for facture in data_relance['factures']:                      
                    if facture['num_fact']: # si nous avons un numéros de facture : 
                        try:
                            # Les niveaus de relances récupérés via Excel sur les relances envoyées peuvent être des int.
                            # Nous devons donc cast en int les variables de notre dictionnaires pour une bonne comparaison.
                            test_niveau_relance = int(facture['num_fact'])
                        except:
                            test_niveau_relance = facture['num_fact']
                        # Si la facture fait partie de la liste des niveaux de relances récupéré via Excel
                        # Alors on ajout 1 a ce niveau de relance
                        if test_niveau_relance in self.niveau_relance_factures:
                            niveau_relance = self.niveau_relance_factures[test_niveau_relance] +1
                        else:
                        # Sinon on le set à 1
                            niveau_relance = 1
                    # Ajuout des données destinnées à alimenter le tableau HTML
                    if facture['montantDebit']>0:
                        montant = ("%.*f" %(2,facture['montantDebit']))
                    elif facture['montantCredit']>0:
                        montant = '- '+("%.*f" %(2,facture['montantCredit']))
                    data_template_corps_tableau.append({"niveau_relance":niveau_relance,
                    "date_fact":facture['date_fact'].strftime("%d/%m/%Y"),
                    'num_fact':facture['num_fact'],
                    'echeance':facture['echeance'].strftime("%d/%m/%Y"),
                    'montantDebit':montant,
                    })

                    # Si la facture est associée à un fichier
                    if facture['pdf']:
                        # On ajoute en pièce jointe cette facture.
                        set_mail.piece_jointe(self.active_PJ_NewTon,facture['pdf'])
                        # Préparation compte rendu email : 
                        compte_rendu_envoie_mail.append([datetime.now(),
                                                        facture['date_fact'],
                                                        facture['echeance'],
                                                        compte_client, facture['num_fact'],
                                                        niveau_relance, facture['pdf'],
                                                        data_relance['contact']['RaisonSocial'],
                                                        data_relance['contact']['Email'],
                                                        facture['montantDebit'],
                                                        facture['montantCredit']
                                                        ])

                        nb_factures+=1 # Incrémentation de 1 par facture traitée
                
                # Mise en forme des données destinnées à alimenter le tableau HTML par un trie par date de facture.
                data_template_corps_tableau = sorted(data_template_corps_tableau, key = self.fctSortDict, reverse=False)

                # Regroupement des datas destinées à alimenter le mail HTML.
                all_datas_template['total_relance']= data_relance['solde_relance']
                all_datas_template['compte_client']=compte_client
                all_datas_template['data_tableau']=data_template_corps_tableau
                all_datas_template['entity']=self.NEWTON_ENTITY
                all_datas_template['IBAN']=self.NEWTON_ACTIVE_IBAN
                all_datas_template['BIC']=self.NEWTON_ACTIVE_BIC

                # Set du corps de mail par retour de template généré par Jinja2
                set_mail.render_template_html(all_datas_template, 'template/block_mail.html')
                # Set basic template : 
                set_mail.render_basic_template(self.template_basic)
                # Set de l'objet du mail
                subject = data_relance['contact']['Email']

                # Envoie Email, test si le processus est toujours accepté par l'utilisateur
                if self.stop_process:
                    return nb_relance, nb_factures
                set_mail.send_mail(subject,"mathieu.leroy@newtonexpertise.com","mathieu.leroy@newtonexpertise.com")
                nb_relance+= 1

                # Maj de la progress bar tout les 0.3 sec.
                self.fram.progressbar.after(300, self.progress(nb_relance))
                self.fram.progressbar.update()

                # Ectriture du rapport de l'email envoyé 
                try:
                    Outil_Excel(self.active_path_wb).relance(compte_rendu_envoie_mail)
                    Outil_Excel(self.active_path_wb).save_wb(self.path_wb)
                except:
                    # si le fichier excel est introuvable : message erreur et close app
                    self.no_excel()

                

        return nb_relance, nb_factures
    
    def fctSortDict(self, Value):
        """
        Méthode permettant un tri par date de la liste : data_template_corps_tableau créer dans la méthode : envoie_des_relance
        """
        return datetime.strptime(Value["date_fact"], "%d/%m/%Y") 

    def progress(self, currentValue):
        """
        set de la valeur de la progress bar
        """
        self.fram.progressbar['value']=currentValue

    def tri_delete_mail(self):
        """
        Tri les élément d'une liste par code client
        """
        tri=[]
        [tri.append((string[:8], string)) for string in self.fram.liste_relance_ignore.get(0,END)]
        tri = sorted(tri)
        self.fram.liste_relance_ignore.delete(0,END)
        for item in tri:
            self.fram.liste_relance_ignore.insert(END,item[1])

    def tri_mail_ok(self):
        """
        Tri les élément d'une liste par code client
        """
        tri=[]
        [tri.append((string[:8], string)) for string in self.fram.liste_relance_valide.get(0,END)]
        tri = sorted(tri)
        self.fram.liste_relance_valide.delete(0,END)
        for item in tri:
            self.fram.liste_relance_valide.insert(END,item[1])

    def tri_nodest(self):
        """
        Tri les élément d'une liste par code client
        """
        tri=[]
        [tri.append((string[:8], string)) for string in self.fram.liste_relance_noDest.get(0,END)]
        tri = sorted(tri)
        self.fram.liste_relance_noDest.delete(0,END)
        for item in tri:
            self.fram.liste_relance_noDest.insert(END,item[1])

    def tri_liste_annulees(self, evt):
        """
        Tri les élément d'une liste par code client
        """
        self.tri_delete_mail()

    def tri_liste_pretes(self,evt):
        """
        Tri les élément d'une liste par code client
        """
        self.tri_mail_ok()

    def quitte_application(self, boolean=True):
        """
        enregistre et ferme excel et ferme l'application
        """
        # test etat processus
        if self.stop_process==0:
            self.stop_process=1
        else:
            if self.path_wb:
                Outil_Excel(self.active_path_wb).save_and_close(self.path_wb)
        if boolean:
            self.fram.window.quit()
        
    def no_excel(self):
        """
        Lorsque le fichier excel n'est pas présent affiche un message informatif à l'utilisateur et ferme le programme
        """
        self.fram.message_box_noexcel(self.active_path_wb)
        sys.exit(0)
