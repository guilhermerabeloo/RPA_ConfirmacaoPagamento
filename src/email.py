import win32com.client as win32
from jinja2 import Environment, FileSystemLoader
import os
import json

def envioDoEmail(tipo, dados):
    with open("../config/config.json", "r", encoding="utf-8") as file:
        sensitive_data = json.load(file)
        emailsCopiaOculta = sensitive_data["enderecosEmailsCCo"]

    outlook = win32.Dispatch('outlook.application')
    email = outlook.CreateItem(0)
    if tipo == 'fatura':
        env = Environment(loader=FileSystemLoader(os.path.join(os.path.dirname(__file__), '../assets')))
        template = env.get_template('emailFatura.html')
        html_body = template.render(data=dados['parcelas'])
        html_body = html_body.format(dataEmail=dados['dataEmail'], empresa=dados['empresa'], fornecedor=dados['fornecedor'], totalBruto=dados['totalBruto'], totalJuros=dados['totalJuros'], deducoesImpostos=dados['deducoesImpostos'], descontos=dados['descontos'], totalLiquido=dados['totalLiquido'])

    else:
        print('oi')
        html_file_path_lancamento = os.path.join(os.path.dirname(__file__), '../assets', 'emailLancamento.html')
        with open(html_file_path_lancamento, 'r', encoding='utf-8') as file:
            html_body = file.read()
            html_body = html_body.format(referencia=dados['referencia'], dataPagamento=dados['dataPagamento'], valorTitulo=dados['valorTitulo'], dataEmail=dados['dataEmail'], empresa=dados['empresa'], fornecedor=dados['fornecedor'], totalBruto=dados['totalBruto'], totalJuros=dados['totalJuros'], deducoesImpostos=dados['deducoesImpostos'], descontos=dados['descontos'], totalLiquido=dados['totalLiquido'])

    email.To = dados['para']
    email.BCC = emailsCopiaOculta
    email.Subject = 'Comprovante de pagamento'
    email.HTMLBody = html_body

    email.Send()

# envioDoEmail('fatura', {
#     'para': 'guilherme.rabelo@grupofornecedora.com.br',
#     'dataEmail': '20/12/2023',
#     'empresa': 'FME MATRIZ 0001/69',
#     'fornecedor': 'FORNECEDORA MAQUINAS E EQUIPAMENTOS LTDA',
#     'totalBruto': '1.000,00',
#     'totalJuros': '54,00',
#     'deducoesImpostos': '1212,54',
#     'descontos': '121,35',
#     'totalLiquido': '875,28',
#     'parcelas': [
#         {
#             'referencia': '122',
#             'parcela': '01',
#             'dataPagamento': '15/11/2023',
#             'valorTitulo': '124,00'
#         },
#         {
#             'referencia': '122',
#             'parcela': '01',
#             'dataPagamento': '15/11/2023',
#             'valorTitulo': '124,00'
#         },
#     ]
# })

# envioDoEmail('lancamento', {
#     'para': 'guilherme.rabelo@grupofornecedora.com.br',
#     'dataEmail': '20/12/2023',
#     'empresa': 'FME MATRIZ 0001/69',
#     'fornecedor': 'FORNECEDORA MAQUINAS E EQUIPAMENTOS LTDA',
#     'totalBruto': '1.000,00',
#     'totalJuros': '54,00',
#     'deducoesImpostos': '1212,54',
#     'descontos': '121,35',
#     'totalLiquido': '875,28',
#     'referencia': '122',
#     'dataPagamento': '15/11/2023',
#     'valorTitulo': '124,00'
# })