import psycopg2
import logging
import datetime
import sys

class BDDPostgreSQL(object):
    def __init__(self):
        self.cursor = ""
        self.conx = ""

    def __enter__(self):

        db_name='outils'
        db_user='admin'
        db_host='10.0.0.17'
        db_password='Zabayo@@'
        db_port='5432'

        try:
            self.conx = psycopg2.connect(database=db_name, user=db_user,host=db_host, password=db_password, port=db_port)
            logging.info(f"Connexion Postgre ok")
            self.cursor = self.conx.cursor()
            return self.cursor

        except (Exception, psycopg2.InterfaceError) as error:
            logging.error(f"Echec connexion Postgre : {error}")
            return 0
        except (Exception, psycopg2.DatabaseError) as error:
            logging.error(f"Echec database Postgre : {error}")
            return 0
        return 1

    def __exit__(self, type, value, traceback):
        logging.info('Fermeture de la base Postgre')
        self.conx.commit()
        self.conx.close()