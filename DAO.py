import pyodbc
import logging
import sys

class BDDQuadra():
    def __init__(self, chem_base):
        self.cursor = None
        self.chem_base = chem_base.lower()
        self.conx = None

    def __enter__(self):

        constr = 'Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=' + \
            self.chem_base
        try:
            self.conx = pyodbc.connect(constr, autocommit=True)
            logging.info('Ouverture de {}'.format(self.chem_base))
            self.cursor = self.conx.cursor()
            return self.cursor

        except pyodbc.Error:
            logging.error("erreur requete base {} \n {}".format(
                self.chem_base, sys.exc_info()[1]))
        except:
            logging.error("erreur ouverture base {} \n {}".format(
                self.chem_base, sys.exc_info()[0]))
        
    def __exit__(self, type, value, traceback):
        logging.info(f'fermeture de la base {self.chem_base}')
        self.conx.commit()
        self.conx.close()
        print('close')
