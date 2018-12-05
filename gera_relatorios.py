
ANALISTAS = []
ANO = 0
MES = ''
OPERACAO = 0

def init(analistas, operacao, mes, ano):
    ANO = ano
    MES = mes
    OPERACAO = operacao
    print('Gerando relatorios...')
    print(MES, 'de', ANO)
    print('Operação:', OPERACAO)
    if OPERACAO == 'makita1':
