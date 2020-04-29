        
def InfoInsertPJDBPostgre(collab, horodata, dossier, compte, document, nomref):
    
    sql = """
    INSERT INTO instadoc (collab,horodat,dossier,compte,document, nomref)
    VALUES (%s,%s,%s,%s,%s,%s)
    """
    data = [collab, horodata, dossier, compte, document, nomref]
    return sql, data

def espion_postgre(collab, horodata, base, nb_mails, nb_factures):

    sql = f"""
    INSERT INTO relance_client(collab, horodat, base, nb_mails, nb_factures)
    VALUES ('{collab}','{horodata}', '{base}','{nb_mails}', '{nb_factures}')
    """
    return sql


def UtilisationTotal():
    
    sql = """
    SELECT *
    FROM espion
    """
    return sql