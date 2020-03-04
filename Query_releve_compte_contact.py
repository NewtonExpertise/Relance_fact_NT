from DAO import BDDQuadra
from collections import defaultdict
from datetime import datetime
import re

def test_mail_valide(email):
    """
    Permet de controler la validité d'un mail.
    retourn True ou False sur le test.
    retourn None si le paramètre attendu n'est pas conforme
    """
    email_regex = re.compile(
    r"(^[-!#$%&'*+/=?^_`{}|~0-9A-Z]+(\.[-!#$%&'*+/=?^_`{}|~0-9A-Z]+)*"  # dot-atom
    # quoted-string
    r'|^"([\001-\010\013\014\016-\037!#-\[\]-\177]|\\[\001-011\013\014\016-\177])*"'
    r")@(?:[A-Z0-9](?:[A-Z0-9-]{0,61}[A-Z0-9])?\.)+[A-Z]{2,6}\.?$",
    re.IGNORECASE,
)  # domain
    try:
        if email_regex.search(email):
            return True
        else:
            return False
    except:
        return False

def get_facture_impayee(BDD):
    """
    Retourn l'écriture comptable pour une facture non lettrée avec une échéance suppérieur à 25 jours.
    On se réfère au journal 'VE1'
    """
    # BDD Qcompta.mdb
    sql = """
    SELECT
    DateSerial(Year(E.PeriodeEcriture), Month(E.PeriodeEcriture), E.JourEcriture) as DateEcr,
    DateSerial(Year(E.PeriodeEcriture), Month(E.PeriodeEcriture), E.JourEcriture+25) as DateEcheance,
    E.NumeroCompte, E.MontantTenuDEbit, E.RefImage,
    Mid(E.RefImage,(InStr(1,RefImage,'_')+1), InStr(1, Mid(E.RefImage,(InStr(1,RefImage,'_')+1),) ,'.')-1) as NumFact
    FROM Ecritures E
    WHERE E.NumeroCompte LIKE '9%'
    AND E.CodeJournal='VE1'
    AND E.CodeLettrage=''
    AND E.TypeLigne='E'
    AND E.MontantTenuCredit=0
    AND DateSerial(Year(E.PeriodeEcriture), Month(E.PeriodeEcriture), E.JourEcriture)<(DATE()-25)
    """
    with BDDQuadra(BDD) as requete_Quadra:
        rq = requete_Quadra.execute(sql)
        rt = rq.fetchall()

    dict_fact_impayee = defaultdict(list)
    for (dateFact, Echeance, NumeroCompte, montantDebit, montantCredit, pdf, numFact, CodeJournal, ) in rt:
        dict_fact_impayee[NumeroCompte].append(
            {
                "date_fact": dateFact,
                "echeance": Echeance,
                "num_fact": numFact,
                "montantDebit": montantDebit,
                "montantCredit": montantCredit,
                "pdf": pdf,
                "CodeJournal":CodeJournal
            }
        )

    return dict(dict_fact_impayee)

def get_ecritures_non_lettrees(BDD):
    """
    Retourn les écritures comptables, non lettrées, pour les clients, avec une échéance suppérieur à 25 jours.
    Tous journaux confondus
    """
    # BDD Qcompta.mdb
    sql = """
    SELECT
    DateSerial(Year(E.PeriodeEcriture), Month(E.PeriodeEcriture), E.JourEcriture) as DateEcr,
    DateSerial(Year(E.PeriodeEcriture), Month(E.PeriodeEcriture), E.JourEcriture+25) as DateEcheance,
    E.NumeroCompte, E.MontantTenuDEbit, E.MontantTenuCredit, E.RefImage,
    E.NumeroPiece,
    E.CodeJournal
    FROM Ecritures E
    WHERE E.NumeroCompte LIKE '90%'
    AND E.CodeLettrage=''
    AND E.TypeLigne='E'
    AND DateSerial(Year(E.PeriodeEcriture), Month(E.PeriodeEcriture), E.JourEcriture)<(DATE()-25)
    """
    with BDDQuadra(BDD) as requete_Quadra:
        rq = requete_Quadra.execute(sql)
        rt = rq.fetchall()

    dict_releve_relance = defaultdict(list)
    for (dateFact, Echeance, NumeroCompte, montantDebit, montantCredit, pdf, numFact, CodeJournal, ) in rt:
        dict_releve_relance[NumeroCompte].append(
            {
                "date_fact": dateFact,
                "echeance": Echeance,
                "num_fact": numFact,
                "montantDebit": montantDebit,
                "montantCredit": montantCredit,
                "pdf": pdf,
                "CodeJournal":CodeJournal
            }
        )

    return dict(dict_releve_relance)

def Query_all_client_for_relance(BDD):
    """
    Retourn tous les contacts des clients susceptible d'être destinataire des relances
    """
    sql = """
    SELECT
    C.NumeroCompte,
    I.Nom as RaisonSocial,
    A.Texte3 as Email,
    A.Texte1 as Nom,
    AN.Prenom,
    AN.DestRelance
    FROM ((Intervenants I
    LEFT JOIN Annexe A  ON I.Code=A.Code1 AND I.Code=A.Code2)
    LEFT JOIN AnnexeSuite AN ON A.Code1=AN.Code1 AND A.Code2=AN.Code2 AND A.Numero=AN.Numero AND A.Type=AN.Type)
    LEFT JOIN Clients C ON C.Code=I.Code
    WHERE A.Type='T'
    OR A.Type IS NULL
    AND C.NumeroCompte <> ''
    """
    with BDDQuadra(BDD) as requete_Quadra:
        rq = requete_Quadra.execute(sql)
        rt = rq.fetchall()

    contacts_statut = {}
    

    for (Code, RaisonSocial, Email, Nom, Prenom, DestRelance) in rt:
        # contact valide et mail conforme : 
        # Si test_mail_valide = True et que l'individu est désigné comme le destinataire des relance alors ajout au dict

        if DestRelance:

            if test_mail_valide(Email):
                contacts_statut[Code] = {'RaisonSocial':RaisonSocial, 'Email':Email, 'Nom':Nom,
                                            'Prenom':Prenom, 'DestRelance':DestRelance, 'Statut_mail':1}
            else:
                contacts_statut[Code] = {'RaisonSocial':RaisonSocial, 'Email':Email, 'Nom':Nom,
                                                'Prenom':Prenom, 'DestRelance':DestRelance, 'Statut_mail':0}
        else:

            if Code in contacts_statut:
                continue
            else:
                contacts_statut[Code] = {'RaisonSocial':RaisonSocial, 'Email':Email, 'Nom':Nom,
                                                'Prenom':Prenom, 'DestRelance':DestRelance, 'Statut_mail':None}

    return contacts_statut

def get_datas_relance(Query_all_client_for_relance, get_ecritures_non_lettrees):
    """
    Retourn un dictionaire faisant état d'un relevé de compte débiteur pour les clients composé comme suit :
    dict[compte client] = {'facture': liste de dictionnaire(s) contenant le(s) information(s) des écritures non lettrées ,
                            'solde_releve': solde débiteur du relevé de comptes,
                             'mail' : un dictionnaire contenant les informations des contacts ou None s'il n'y a pas de destinataire désigné dans Quadra
                          }
    Si la clé mail est différente de None alors elle contiendra un dictionnaire qui, en son sein, en plus des informations du contacts,  aura
    une clé Statut_mail faisant état de la validité de l'adresse mail selon le regex suivant :

    re.compile(
    "(^[-!#$%&'*+/=?^_`{}|~0-9A-Z]+(\.[-!#$%&'*+/=?^_`{}|~0-9A-Z]+)*"  # dot-atom
    # quoted-string
    '|^"([\001-\010\013\014\016-\037!#-\[\]-\177]|\\[\001-011\013\014\016-\177])*"'
    ")@(?:[A-Z0-9](?:[A-Z0-9-]{0,61}[A-Z0-9])?\.)+[A-Z]{2,6}\.?$",
    re.IGNORECASE,)

    si Statut_mail == 1 alors l'adresse mail est considéré valide si Statut_mail vaut 0 il est considéré comme non valide
    """
    data_relance= defaultdict(list) # Constitura l'ensemble des datas pour la relance
    ecritures_non_lettrees = get_ecritures_non_lettrees
    contact_client = Query_all_client_for_relance

    # On boucle sur le dictionnaire des écritures non lettrées
    for compte_client, ecritures in ecritures_non_lettrees.items():
        # On cherche a savoir si les éctritures non lettrées d'un client constitue un solde débiteur.
        # # Si c'est le cas alors on ajoute le total de la créance au dictionnaire à retourner
        # # # Sinon on passe à l'itération suivante.

        solde_ecriture = 0.
        for ecriture in ecritures:
            if ecriture["montantDebit"]:
                solde_ecriture += ecriture["montantDebit"]
            elif ecriture["montantCredit"]:
                solde_ecriture -= ecriture["montantCredit"]
        if solde_ecriture > 0.:
            # Si les écritures non lettrées constitue un solde débiteur.
            # Nous associons le contacts du client à relancer.
            # if compte_client in contact_client:
            data_relance[compte_client].append(
                    {
                        "factures": ecritures,
                        "solde_relance": round(solde_ecriture, 2),
                        "contact": contact_client[compte_client]
                    }
            )

    return dict(data_relance)





# if __name__ == "__main__":
#     # import pprint

    # pp = pprint.PrettyPrinter(indent=4)

    # bdd = r'\\srvquadra\Qappli\Quadra\DATABASE\gi\0000\Qgi.mdb'

    # pp.pprint(Query_all_client_for_relance())
    # pp = pprint.PrettyPrinter(indent=4)
    # x= Query_all_client_for_relance()
    # d=0
    # for c in x:
    #     d+=1
    # print(d)
    # info_envoie_mail()
    # pp.pprint(x["valide"])
    # pp.pprint(x["email_invalide"])
    # pp.pprint(x["no_dest_relance"])
    # print(get_facture_impaye())
    # # get_client_with_dest()
    # pp.pprint(get_ecritures_non_lettrees())
    # pp.pprint(get_facture_impaye())
    # pp.pprint(x)
    # print(len(Query_all_client_for_relance()))
    # pp.pprint(get_dest_incorrect_mail(Query_all_client_for_relance()))
    # pp.pprint(Query_all_client_for_relance())

    # print(test_mail_valide(None))

#####/////////////////////////////////////////////////////
    # ecritures = get_ecritures_non_lettrees()
    # clients = Query_all_client_for_relance(bdd)
    # pp.pprint(clients)
    # f = get_datas_relance()
    # pp.pprint(f)
    # print(len(f))
    # exit()
    # pp.pprint(ecritures)
    # clients_mail_valide = get_client_with_valide_mail(clients)
    # clients_mail_invalide = get_clients_invalide_mail(clients)
    # clients_no_mail = get_clients_no_mail(clients)
    # pp.pprint(datarelance_global(ecritures , clients))
    # # pp.pprint(clients_mail_valide)
    # # pp.pprint(clients_mail_invalide)
    # # pp.pprint(clients_no_mail)

    # RCV = get_relances_releve_compte_mail_valide(ecritures, clients_mail_valide)
    # RCI = get_relances_releve_compte_mail_invalide(ecritures, clients_mail_invalide)
    # RCND = get_relance_sans_destinataire(ecritures, clients_no_mail)

    # print(type(RCV))
    # print(type(RCI))
    # print(type(RCND))

