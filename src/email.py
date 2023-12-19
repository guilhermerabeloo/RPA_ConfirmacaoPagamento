import win32com.client as win32
import os
import json

# html_file_path_fatura = os.path.join(os.path.dirname(__file__), 'assets', 'emailFatura.html')
# html_file_path_lancamento = os.path.join(os.path.dirname(__file__), 'assets', 'emailLancamento.html')

def envioDoEmail(tipo, dados):
    with open("../config/config.json", "r", encoding="utf-8") as file:
        sensitive_data = json.load(file)
        emailsCopiaOculta = sensitive_data["enderecosEmailsCCo"]

    outlook = win32.Dispatch('outlook.application')
    email = outlook.CreateItem(0)
    if tipo == 'fatura':
        # with open(html_file_path_fatura, 'r', encoding='utf-8') as file:
        #     html_body = file.read()
        #     html_body = 'Body fatura'
        html_body = 'Body fatura'
            # html_body = html_body.format(cliente=dados['cliente'], emissao=dados['emissao'], empresa=dados['empresa'], numeroNotaFiscal=dados['numeroNota'], vencimentoNotaFiscal=dados['vencimento'], valorBoleto=dados['boleto'], valorTitulo=dados['nota'])

    else:
        # with open(html_file_path_lancamento, 'r', encoding='utf-8') as file:
        #     html_body = file.read()
        #     html_body = 'Body lancamento'
        html_body = 'Body lancamento'
            # html_body = html_body.format(cliente=dados['cliente'], emissao=dados['emissao'], empresa=dados['empresa'], numeroNotaFiscal=dados['numeroNota'], vencimentoNotaFiscal=dados['vencimento'], valorBoleto=dados['boleto'], valorTitulo=dados['nota'])

    # email.To = dados['email']
    email.To = 'guilherme.rabelo@grupofornecedora.com.br'
    email.BCC = emailsCopiaOculta
    email.Subject = 'Comprovante de pagamento'
    email.HTMLBody = html_body

    email.Send()

envioDoEmail('fatura', '')
envioDoEmail('lancamento', '')