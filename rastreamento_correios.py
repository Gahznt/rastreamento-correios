import requests
from tqdm import tqdm
import pandas as pd

def main():
    plan = pd.read_excel("Correios.xlsx")
    sro_list = plan['SRO'].tolist()

    print('------------------')
    print('## Consulta Objeto-Correios ##')

    # codInput = input('Digite o Codigo: ')
    rastreamentos = []
    for sro in tqdm(sro_list):
        request = requests.get(' https://proxyapp.correios.com.br/v1/sro-rastro/{}'.format(sro))
        result = request.json()

        if 'mensagem' not in result['objetos'][0]:
            rastreamentos.append((result['objetos'][0]['codObjeto'], result['objetos'][0]['eventos'][0]['descricao'], result['objetos'][0]['eventos'][0]['dtHrCriado']))
        else:
            rastreamentos.append((result['objetos'][0]['codObjeto'], result['objetos'][0]['mensagem']))
    print("==> Rastreamento completo")

    completed = pd.DataFrame(rastreamentos)
    completed.to_excel('Result_Rastreamento.xlsx', index = False)  #Transforma o dataframe em excel

if __name__ == '__main__':
    main()