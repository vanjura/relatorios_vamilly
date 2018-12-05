from __future__ import print_function

import os
from pathlib import Path

from googleapiclient.discovery import build
from httplib2 import Http
from oauth2client import file, client, tools
from apiclient import errors
import base64
import email
import json

# modelo: analista - mes - ano

# If modifying these scopes, delete the file token.json.
SCOPES = 'https://www.googleapis.com/auth/gmail.readonly'
MES = input('Digite o mês:')
ANO = input('Digite o ano:')
OPERACAO = input('Digite o codigo da operação:')
ANALISTAS = [
    ['Matheus', 'matheus.fujihara@vamilly.com.br', 'Matheus Shiniti Fujihara'],
    ['Andre Silva', 'andre.silva@vamilly.com.br', ''],
    ['Edinei', 'edinei@vamilly.com.br', ''],
    ['Anderson', 'anderson.scheffer@vamilly.com.br', ''],
    ['Thales', 'thales.marenda@vamilly.com.br', ''],
    ['Iohan', 'iohan.camargo@vamilly.com.br', ''],
    ['Tiago Sc.', 'tscheneider@vamilly.com.br', 'Tiago Scheneider'],
    ['Thiago B.', 'thiago.bach@vamilly.com.br', 'Tiago Bach'],
    ['Gilvan', 'gilvan@vamilly.com.br', 'Gilvan Ramos Prioto'],
    ['Cesar', 'cesar.affonso@vamilly.com.br', 'Cesar Affonso Oliveira'],
    ['Marcos B.', 'marcos.borck@vamilly.com.br', 'Marcos Vinicius Borck'],
    ['Maycon', 'maycon@vamilly.com.br', 'Maycon Robson Branco'],
    ['Alessandro', 'alessandro.sperandio@vamilly.com.br', 'Alessandro Sperandio'],
    ['Edenilson', 'edenilson.santos@vamilly.com.br', 'Edenilson dos Santos']
]


def get_mes(num):
    meses = [
        'Janeiro',
        'Fevereiro',
        'Março',
        'Abril',
        'Maio',
        'Junho',
        'Julho',
        'Agosto',
        'Setembro',
        'Outubro',
        'Novembro',
        'Dezembro'
    ]
    num = int(num) - 1
    return meses[num]


def create_file_name(analista):
    file_name = analista + ' - ' + get_mes(MES) + ' ' + str(ANO) + '.xlsx'
    return file_name


def get_attachments(service, user_id, msg_id, analista=''):
    try:
        message = service.users().messages().get(userId=user_id, id=msg_id).execute()

        for part in message['payload']['parts']:
            if part['filename']:
                if 'data' in part['body']:
                    data = part['body']['data']
                else:
                    att_id = part['body']['attachmentId']
                    att = service.users().messages().attachments().get(userId=user_id, messageId=msg_id,
                                                                       id=att_id).execute()
                    data = att['data']
                file_data = base64.urlsafe_b64decode(data.encode('UTF-8'))
                path = create_file_name(analista)

                with open(path, 'wb') as f:
                    f.write(file_data)
        return True
    except errors.HttpError as error:
        print('An error occurred: %s' % error)
        return False


def get_message(service, user_id, msg_id):
    try:
        message = service.users().messages().get(userId=user_id, id=msg_id).execute()
        return message
    except errors.HttpError as error:
        print('An error occurred: %s' % error)


def list_messages_matching_query(service, user_id, query=''):
    try:
        response = service.users().messages().list(userId=user_id, q=query).execute()
        messages = []
        if 'messages' in response:
            messages.extend(response['messages'])

        while 'nextPageToken' in response:
            page_token = response['nextPageToken']
            response = service.users().messages().list(userId=user_id, q=query,
                                                       pageToken=page_token).execute()
            messages.extend(response['messages'])

        return messages
    except errors.HttpError as error:
        print('An error occurred: %s' % error)


def get_analista(msg):
    headers = msg['payload']['headers']
    for head in headers:
        if head['name'] == 'Subject':
            subject_ = head['value']
        if head['name'] == 'From':
            from_ = head['value']
    de = from_.split('<')
    for d in de:
        if '@vamilly.com.br' in d:
            from_email = d.strip('>')
    from_analista = 'ERR'
    for analista in ANALISTAS:
        if from_email == analista[1]:
            from_analista = analista[0]
            break
        else:
            sub = subject_.split('-')
            for s in sub:
                if s != from_analista:
                    for analist in ANALISTAS:
                        if s == analist[2]:
                            from_analista = s
                            break
                else:
                    break
    if from_analista == 'ERR':
        print('########## ERRO ##########')
        print('Analista do email "' + subject_ + '" não encontrado. From:', from_)
        exit()
    return from_analista


def init(service):
    print('Iniciando leitura de dados.')
    print('Gerando lista de emails não lidos')
    print('Abrindo emails de:')
    mensagens = list_messages_matching_query(service, 'me', 'is:unread')
    # print(mensagens)
    for mensagem in mensagens:
        msg = get_message(service, 'me', mensagem['id'])
        analista = get_analista(msg)
        if analista:
            print('\t', analista)
            get_attachments(service, 'me', mensagem['id'], analista)
        else:
            print('Analista não encontrado no email id', msg['id'])
    print('Excel dos analistas acima baixado com sucesso.')


def main():
    store = file.Storage('token.json')
    creds = store.get()
    if not creds or creds.invalid:
        flow = client.flow_from_clientsecrets('credentials.json', SCOPES)
        creds = tools.run_flow(flow, store)
    service = build('gmail', 'v1', http=creds.authorize(Http()))

    init(service)


if __name__ == '__main__':
    main()
