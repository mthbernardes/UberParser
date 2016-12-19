# -*- coding: utf-8 -*-

import os
import imapy
import pdfkit
from lxml import html
from pprint import pprint
from dateutil import parser
from imapy.query_builder import Q

#Vars
email = 'youremail@gmail.com'
password = 'yourpassword'
imap_server = 'imap.gmail.com'
emails_since = '27-Oct-2016'
subject = 'Sua viagem de'

box = imapy.connect(host=imap_server,username=email,password=password,ssl=True)
q = Q()

print 'Searching e-mails...'
emails = box.folder('INBOX').emails(q.subject(subject).sender("uber.com").since(emails_since))
print 'Were found %d e-mails' % len(emails)
options = {'quiet': ''}
for email in emails:
    pdf_file = '%s.pdf' % email.uid
    pdf_path = os.path.join('saida',pdf_file)
    print pdf_file
    corpo = email['html'][0].decode('utf-8','ignore')
    try:
        pdfkit.from_string(corpo,pdf_path,options=options)
    except Exception as e:
        print e
    with open('saida.csv','a') as f:
        tree = html.fromstring(email['html'][0])
        valor = tree.xpath('//*[@class="totalPrice topPrice tal black"]/text()')
        partida = tree.xpath('//*[@id="normalABlock"]/tbody/tr/td/table/tbody/tr[1]/td[5]/text()')
        destino = tree.xpath('//*[@id="normalABlock"]/tbody/tr/td/table/tbody/tr[2]/td[5]/text()')
        dt = parser.parse(email['date'])
        valor = valor[0].strip().replace(',','.').replace('R$','')
        partida = partida[0].strip().replace(',','')
        destino = destino[0].strip().replace(',','')
        date = '{0}/{1}/{2}'.format(dt.day,dt.month,dt.year)
        result = '%s,%s,%s,%s,%s\n' %(email.uid,date,partida,destino,valor)
        f.write(result.encode('utf-8'))
print 'All Done Fella'
