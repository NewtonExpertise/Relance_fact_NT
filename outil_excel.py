import xlwings as xw
from datetime import datetime

class Outil_Excel():

    def __init__(self, path_excel):

        self.wb = xw.Book(path_excel)
        self.wb.activate(steal_focus=True)
 
        self.ws_fact_du = self.wb.sheets["Facture_du"]
        self.sht_all_client = self.wb.sheets["No_destinataire"]
        self.sht_relance = self.wb.sheets["Relance"]

    def fact_du(self, give_ecritures_non_lettrees):
        """
        Ecrit dans le fichier excel dans l'onglet fact_dû l'ensemble des factures en souffrance à un instant T
        """
        facture_impayee = []
        ecritures_non_lettrees  = give_ecritures_non_lettrees
        for compteClient , listeEcrtiture in ecritures_non_lettrees.items():
            for ecriture in listeEcrtiture:
                if ecriture['pdf'] and ecriture['CodeJournal']=='VE1' and ecriture['CodeJournal']=='VE1':
                    facture_impayee.append([ecriture['date_fact'] ,ecriture['echeance'] ,
                    compteClient ,ecriture['montantDebit'] ,ecriture['montantCredit'] ,
                    ecriture['pdf'] ,ecriture['num_fact'] ,ecriture['CodeJournal']])

        firstLine =[("Date ecriture", "Date écheance", "Num compte", "Debit", "Credit" , "Ref image", "Num fact", "Code journal")]
        tab_fact_du = firstLine+facture_impayee
        self.ws_fact_du.clear()
        self.ws_fact_du.range('A1').value = tab_fact_du

    def No_mail_dest(self, Query_all_client_for_relance):
        """
        Répertorie dans un fichier excel l'ensemble des individues faisant
        partie de l'ensemble des comptes clients n'ayant pas de destinataire paramétré.
        """
        all_client = Query_all_client_for_relance
        list_all_client=[]
        list_client_dest_valide =[]
        for compte, contact in all_client.items():
            if contact['DestRelance']:
                list_client_dest_valide.append(compte)

        for compte, contact in all_client.items():
            if compte in list_client_dest_valide:
                continue
            else:
                list_all_client.append([compte,contact['RaisonSocial'],
                                        contact['Email'],
                                        contact['Nom'],
                                        contact['Prenom'],
                                        contact['DestRelance']
                                        ])
        firstLine = [("Code","RaisonSocial","Email","Nom","Prenom","DestRelance")]
        tab_all_client = firstLine+list_all_client
        self.sht_all_client.clear()
        self.sht_all_client.range('A1').value = tab_all_client

    def relance(self, list_send_relance):
        """
        Ajoute le compte rendu des relances envoyées dans l'onglet Relance
        """

        nbligne = self.sht_relance.cells(self.sht_relance.api.rows.count, "A").end(-4162).row +1
        self.sht_relance.range('A'+str(nbligne)).value = list_send_relance
        
    def historique_relance(self):
        """
        Permet d'obtenir un dictionnaire du niveau de relance de chaques factures
        """

        nbligne = self.sht_relance.cells(self.sht_relance.api.rows.count, "A").end(-4162).row
        dict_fact_deja_relancee = self.sht_relance.range('E1:F'+str(nbligne)).options(dict, numbers=int).value
        
        return dict_fact_deja_relancee

    def save_and_close(self, liste_workook):
        """
        Permet de fermer excel après utilisation du module de relance.
        """
        for book in liste_workook:
            try:
                self.wb = xw.Book(book)
                self.wb.save()
                self.wb.close()
            except:
                print('rip')

    def save_wb(self, liste_workook):
        """
        Save wb
        """
        for book in liste_workook:
            try:
                self.wb = xw.Book(book)
                self.wb.save()
            except:
                pass


# if __name__ == "__main__":

#     wb =Outil_Excel(r'V:\Mathieu\relance_facture_client\excel\Relance_clients_audit.xlsx')
    
#     print(xw.books(wb.wb))
    
