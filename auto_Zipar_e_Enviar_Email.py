import zipfile as zipf #imports para gerar arquivo.zip
import os

import smtplib #imports para criar função de enviar e-mail (GMAIL)
import email.message

import ctypes #import para gerar poup up de confirmação

import mimetypes #imports e from para criar função de anexar arquivos no e-mail
from email import encoders
from email.mime.audio import MIMEAudio
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

from datetime import datetime #from para capturar data e hora

now = datetime.now()
mes = now.month
mes = mes - 1
ano = now.year
empresa_arq = 'FranciscaLopesME'
empresa = 'Francisca L S Lopes ME'

if (mes < 1):
    mes = 12
    ano = (ano - 1)

if (mes < 10):
    mes = f'0{mes}'

nomeArquivo = (f'({empresa_arq}_{mes}-{ano})')

def zipar(arqs):
    with zipf.ZipFile('Autorizadas ' + nomeArquivo + '.zip','w', zipf.ZIP_DEFLATED) as z:
        for arq in arqs:
            if(os.path.isfile(arq)): # se for ficheiro
                z.write(arq)
            else: # se for diretorio
                for root, dirs, files in os.walk(arq):
                    for f in files:
                        z.write(os.path.join(root, f))

zipar([f'../Autorizadas/{ano}-{mes}'])
########################################################################################################

def zipar(arqs):
    with zipf.ZipFile('Canceladas ' + nomeArquivo + '.zip','w', zipf.ZIP_DEFLATED) as z:
        for arq in arqs:
            if(os.path.isfile(arq)): # se for ficheiro
                z.write(arq)
            else: # se for diretorio
                for root, dirs, files in os.walk(arq):
                    for f in files:
                        z.write(os.path.join(root, f))

zipar([f'../Canceladas/{ano}-{mes}'])
########################################################################################################

def zipar(arqs):
    with zipf.ZipFile('Cce ' + nomeArquivo + '.zip','w', zipf.ZIP_DEFLATED) as z:
        for arq in arqs:
            if(os.path.isfile(arq)): # se for ficheiro
                z.write(arq)
            else: # se for diretorio
                for root, dirs, files in os.walk(arq):
                    for f in files:
                        z.write(os.path.join(root, f))

zipar([f'../Cce/{ano}-{mes}'])
########################################################################################################

def zipar(arqs):
    with zipf.ZipFile('Geradas ' + nomeArquivo + '.zip','w', zipf.ZIP_DEFLATED) as z:
        for arq in arqs:
            if(os.path.isfile(arq)): # se for ficheiro
                z.write(arq)
            else: # se for diretorio
                for root, dirs, files in os.walk(arq):
                    for f in files:
                        z.write(os.path.join(root, f))

zipar([f'../Geradas/{ano}-{mes}'])
########################################################################################################
### Criando e Executando Função de Zipar (Autorizadas, Canceladas, Cce e Geradas) ###########################################################


def adiciona_anexo(msg, filename):
    if not os.path.isfile(filename):
        return

    ctype, encoding = mimetypes.guess_type(filename)

    if ctype is None or encoding is not None:
        ctype = 'application/octet-stream'

    maintype, subtype = ctype.split('/', 1)

    if maintype == 'text':
        with open(filename) as f:
            mime = MIMEText(f.read(), _subtype=subtype)
    elif maintype == 'image':
        with open(filename, 'rb') as f:
            mime = MIMEImage(f.read(), _subtype=subtype)
    elif maintype == 'audio':
        with open(filename, 'rb') as f:
            mime = MIMEAudio(f.read(), _subtype=subtype)
    else:
        with open(filename, 'rb') as f:
            mime = MIMEBase(maintype, subtype)
            mime.set_payload(f.read())

        encoders.encode_base64(mime)

    mime.add_header('Content-Disposition', 'attachment', filename=filename)
    msg.attach(mime)
    ### Criando Função para Anexar Arquivos ###########################################################

from config import email, pswd
from smtplib import SMTP

de = email
para = ['j.costa13@hotmail.com']
acesso = pswd

msg = MIMEMultipart()
msg['From'] = de
msg['To'] = ', '.join(para)
msg['Subject'] = f'XML/NFe: {empresa} - {mes} de {ano}.'

# Corpo da mensagem
msg.attach(MIMEText(f"""
<p> Segue em anexo o Envio da XML, de <strong>{empresa}</strong></p> 
<p> referente à <strong>{mes} de {ano}</strong>.<p>
<br/>
<br/>
<br/>
<hr/>

<p>Att.</p>
<p><strong>Empresa</strong></p>



""", 'html', 'utf-8'))

# Arquivos anexos.
adiciona_anexo(msg, 'Autorizadas ' + nomeArquivo + '.zip')
adiciona_anexo(msg, 'Canceladas ' + nomeArquivo + '.zip')
adiciona_anexo(msg, 'Cce ' + nomeArquivo + '.zip')
adiciona_anexo(msg, 'Geradas ' + nomeArquivo + '.zip')

raw = msg.as_string()

smtp=SMTP('smtp.live.com',587)
smtp.starttls()
smtp.login(email, acesso)
smtp.sendmail(de, para, raw)
smtp.quit()
### Parametros para Enviar E-mail (GMAIL) ###########################################################

def Mbox(title, text, style):
    return ctypes.windll.user32.MessageBoxW(0, text, title, style)

Mbox('ENVIO DE XML/NFe', 'E-mail Enviado!', 1)
### Função para Poup up de Confirmação ###########################################################

###  ###########################################################
##### Jean da Costa ###########################################################
###  ###########################################################

