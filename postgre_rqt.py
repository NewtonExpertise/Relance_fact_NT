        
def InfoInsertPJDBPostgre(collab, horodata, dossier, compte, document, nomref):
    
    sql = """
    INSERT INTO instadoc (collab,horodat,dossier,compte,document, nomref)
    VALUES (%s,%s,%s,%s,%s,%s)
    """
    data = [collab, horodata, dossier, compte, document, nomref]
    return sql, data

def espion_postgre(collab, horodata, dossier, base, args):

    argument = ";".join(args)
    sql = f"""
    INSERT INTO espion (collab, horodat, dossier, base, args)
    VALUES ('{collab}','{horodata}','{dossier}','{base}','{argument}')
    """
    return sql


def UtilisationTotal():
    
    sql = """
    SELECT *
    FROM espion
    """
    return sql