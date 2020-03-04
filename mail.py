# coding: utf-8

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
import mimetypes
from email import encoders
import os
from jinja2 import Environment, FileSystemLoader, Template

class mail():
    def __init__(self):

        self.msg = MIMEMultipart()
        self.host = 'smtp.serinyatelecom.fr'
        self.port = 25

    def render_basic_template(self, template_file):
        """
        Set un message basic pour les destinataires ne pouvant supporter le format HTML
        """
        loader = FileSystemLoader('.')
        env = Environment(loader=loader)
        template = env.get_template(template_file)
        msg = template.render()
        self.msg.attach(MIMEText(msg, "plain"))

    def render_template_html(self, data_dict, template_file):
        ''' 
        Set le message html via jinja2
        '''
        loader = FileSystemLoader('.')
        env = Environment(loader=loader)
        template = env.get_template(template_file)
        msg = template.render(data_dict)
        self.msg.attach(MIMEText(msg,'html'))
   


    def piece_jointe(self, path_pj , nompiècejointe):
        """
        Ajoute une pièce jointe au mail.
        """
        nom_fichier = nompiècejointe    ## Spécification du nom de la pièce jointe
        with open(os.path.join(path_pj, nompiècejointe), "rb") as f:   ## Ouverture du fichier
            pieceJointe = MIMEBase('application', 'octet-stream')    ## Encodage de la pièce jointe en Base64
            pieceJointe.set_payload(f.read())
            encoders.encode_base64(pieceJointe)
            pieceJointe.add_header('Content-Disposition', "piece; filename= %s" % nom_fichier)
            self.msg.attach(pieceJointe)   ## Attache de la pièce jointe à l'objet "message" 

    def add_img_corp_mail(self, pathimg, id_img):
        with open(pathimg, 'rb') as i:
            # maintype, subtype = mimetypes.guess_type(i.name)[0].split('/')
            img = MIMEImage(i.read(), id_img )
        img.add_header('Content-ID', '<'+id_img+'>')
        img.add_header('Content-Disposition', 'attachment', filename=id_img)
        self.msg.attach(img)

    def send_mail(self, objet, expediteur, destinataire, cc=None, bcc=None):
        """
        envoie l'email.
        """
        # try:
        #     self.add_img_corp_mail(r'C:\Users\mathieu.leroy\Desktop\relance_client_test\logo.jpg','1')
        # except:
        #     pass
        # try:
        #     self.add_img_corp_mail(r'C:\Users\mathieu.leroy\Desktop\relance_client_test\fb.png','2')
        # except:
        #     pass
        # try:
        #     self.add_img_corp_mail(r'C:\Users\mathieu.leroy\Desktop\relance_client_test\twitter.png','3')
        # except:
        #     pass

        if type(destinataire) is not list:
            destinataire = destinataire.split()

        destinataires_list = destinataire + [cc] + [bcc]
        destinataires_list = list(filter(None, destinataires_list)) # remove null emails
        print(destinataires_list)
        print(expediteur)
        self.msg['From'] = expediteur
        self.msg['Subject'] = objet
        self.msg['To']= ', '.join(destinataires_list)
        self.msg['Cc']      = cc
        self.msg['Bcc']     = bcc
        mailserver = smtplib.SMTP(self.host , self.port)
        try:
            x=mailserver.sendmail(expediteur, destinataires_list, self.msg.as_string())
            
            print(x)
        except smtplib.SMTPException as e:
            print(e)
            return False
        mailserver.quit()
        

# nicolas.rollet@newtonexpertise.com
# mathieu.leroy@newtonexpertise.com
# jasmine.lefebvre@newtonexpertise.com

    # def set_mail_relance(self):

# if __name__ == '__main__':
#     set_mail=mail()
#     data_exemple = {'total_relance': 1977.0,
#                     'compte_client': '90000729',
#                     'data_tableau': [{'niveau_relance': 3, 
#                                     'date_fact': '14/02/2019',
#                                     'num_fact': '0190211545',
#                                     'echeance': '11/03/2019',
#                                     'montantDebit': 431.4,
#                                     'montantCredit': 0.0},
#                                     {'niveau_relance': 3,
#                                     'date_fact': '15/11/2016',
#                                     'num_fact': '0161105903',
#                                     'echeance': '10/12/2016',
#                                     'montantDebit': 1692.0,
#                                     'montantCredit': 0.0},
#                                     {'niveau_relance': 1,
#                                     'date_fact': '16/02/2017',
#                                     'num_fact': '010',
#                                     'echeance': '13/03/2017',
#                                     'montantDebit': 0.0,
#                                     'montantCredit': 2052.0},
#                                     {'niveau_relance': 3,
#                                     'date_fact': '16/04/2018',
#                                     'num_fact': '0180409348',
#                                     'echeance': '11/05/2018',
#                                     'montantDebit': 1155.6,
#                                     'montantCredit': 0.0},
#                                     {'niveau_relance': 3,
#                                     'date_fact': '18/11/2019',
#                                     'num_fact': '0191113367',
#                                     'echeance': '13/12/2019',
#                                     'montantDebit': 750.0,
#                                     'montantCredit': 0.0}
#                                     ]}

    
#     set_mail.render_template_html(data_exemple , 'block_mail.html')
#     set_mail.render_basic_template('template_brut_ppl.txt')
#     set_mail.piece_jointe(r'V:\Procédures','Instadoc.pdf')
#     set_mail.send_mail('objet', 'mathieu.leroy@newtonexpertise.com', 'mathieu.leroy@newtonexpertise.com')