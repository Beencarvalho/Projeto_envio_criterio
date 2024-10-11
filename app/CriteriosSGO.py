import pandas as pd
import requests
import json
import os
import sys
import time
from util.api_token import url, headers

def show_startup_animation():
    # Desenho simples em ASCII
    logo = [
        "  #####   #####    ##### ",
        " #     # #     #  #     #",
        " #       #        #     #",
        "  #####  #  ####  #     #",
        "       # #     #  #     #",
        " #     # #     #  #     #",
        "  #####   #####    ##### "
    ]

    # Animação do desenho
    for line in logo:
        print(line)
        time.sleep(0.1)  # Pequeno delay para criar o efeito de "desenho"

    # Mensagem de inicialização
    print("\n\nIniciando conexão com API SGO")
    
    # Animação de carregamento
    loading_animation = ["[=     ]", "[==    ]", "[===   ]", "[====  ]", "[===== ]", "[======]"]
    for i in range(3):  # Repetir a animação algumas vezes
        for frame in loading_animation:
            sys.stdout.write("\r" + frame)
            sys.stdout.flush()
            time.sleep(0.2)  # Delay entre os frames
    print("\n\nConexão estabelecida com sucesso!")

# Chamar a função para exibir a animação
show_startup_animation()

# Caminho para a área de trabalho
desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')

# Caminho para a pasta que contém os arquivos Excel
pasta_arquivos = os.path.join(desktop_path, 'dados')
pasta_json = os.path.join(pasta_arquivos, 'json')  # 'json' dentro da pasta 'dados'

# Certifique-se de que a pasta 'dados' exista
if not os.path.exists(pasta_arquivos):
    os.makedirs(pasta_arquivos)

# Certifique-se de que a pasta 'json' dentro de 'dados' exista
if not os.path.exists(pasta_json):
    os.makedirs(pasta_json)

# Mensagem para o usuário
print('\n\n\n\nAgora você precisa adicionar seus arquivos de critérios na pasta "DADOS" em sua área de trabalho.')
print('\nCaso a pasta não exista, já criei uma para você, pode conferir.')
print('\nInclua os arquivos de critérios na pasta.')

# Loop para verificar se os arquivos foram incluídos
while True:
    resposta = input('\n\nArquivos incluídos? (sim/nao): ').strip().lower()
    if resposta == 'sim':
        # Listar os arquivos na pasta "dados"
        arquivos_excel = [os.path.join(pasta_arquivos, f)
                          for f in os.listdir(pasta_arquivos)
                          if f.endswith(('.xlsx', '.xls')) and not f.startswith('~$')]

        # Verificar se encontrou algum arquivo
        if len(arquivos_excel) == 0:
            print('\nNenhum arquivo de critérios foi encontrado na pasta "DADOS".')
            print('Por favor, adicione os arquivos e informe quando estiver pronto.')
            time.sleep(2)  # Espera 2 segundos antes de perguntar novamente
        else:
            print(f'\n{len(arquivos_excel)} arquivo(s) encontrado(s). Vamos prosseguir com o processamento.')
            break
    elif resposta == 'nao':
        print('\nPor favor, adicione os arquivos à pasta "DADOS" e informe quando estiver pronto.')
        time.sleep(2)  # Espera 2 segundos antes de perguntar novamente
    else:
        print('\nResposta inválida. Por favor, responda "sim" ou "nao".')

# Função para criar o JSON para cada linha do DataFrame
def criar_item_json(linha, id_item):
    # Garantir que os valores ausentes sejam tratados
    cod_setor = linha['cod_setor'] if pd.notnull(linha['cod_setor']) else 0
    centro_custo = linha['centro_custo'] if pd.notnull(linha['centro_custo']) else 'N/A'
    base = linha['base'] if pd.notnull(linha['base']) else 0
    distrib = linha['distribuicao'] if pd.notnull(linha['distribuicao']) else 0
    
    # Verificar se a empresa tem um formato válido
    if pd.notnull(linha['empresa']) and " - " in linha['empresa']:
        cod_empresa = linha['empresa'].split(" - ")[0]
        try:
            id_empresa = int(cod_empresa)
        except ValueError:
            id_empresa = 0  # Valor padrão se não for possível converter
        name_emp = linha['empresa']
        code_emp = cod_empresa
    else:
        id_empresa = 0
        name_emp = 'Empresa Desconhecida'
        code_emp = '0'
    
    return {
        "active": True,
        "base": base,
        "distribution": distrib,
        "apportionmentId": id_item,
        "sectorId": int(cod_setor),
        "sector": {
            "active": True,
            "id": int(cod_setor),
            "name": linha['setor'],
            "code": str(int(cod_setor)),
            "codeCostCenter": centro_custo,
            "company": {
                "active": True,
                "id": id_empresa,
                "name": name_emp,
                "code": code_emp
            },
            "companyId": id_empresa
        },
        "orderBy": None
    }

# Função para processar um único arquivo Excel
def processar_arquivo(pasta_arquivos):
    nome_arquivo = os.path.splitext(os.path.basename(pasta_arquivos))[0]
    df2 = pd.read_excel(pasta_arquivos)
    description = df2.iloc[3, 5]

    df = pd.read_excel(pasta_arquivos, skiprows=8)
    df = df.iloc[:-1, :6]
    df.columns = ['empresa', 'cod_setor', 'centro_custo', 'setor', 'base', 'distribuicao']
    df['distribuicao'] = pd.to_numeric(df['distribuicao'], errors='coerce').apply(
        lambda x: round(x * 100, 4) if pd.notnull(x) else 0.0000)

    id_item = 0
    itens = [criar_item_json(linha, id_item) for _, linha in df.iterrows()]

    json_final = {
        "active": True,
        "name": nome_arquivo,
        "description": description,
        "items": itens,
        "cycle": None,
        "cycleId": 5,
        "baseSum": None,
        "totalItems": None
    }

    return json_final

# Função para processar o arquivo e salvar o JSON na pasta 'json' dentro de 'dados'
def salvar_json_local(arquivo, pasta_json):
    json_resultado = processar_arquivo(arquivo)
    nome_arquivo_json = os.path.splitext(os.path.basename(arquivo))[0] + '.json'
    
    with open(os.path.join(pasta_json, nome_arquivo_json), 'w') as f:
        json.dump(json_resultado, f, indent=4)
    
    return nome_arquivo_json

# Função para enviar um JSON via POST
def enviar_json(json_dados, nome_arquivo_json):
    try:
        response = requests.post(url, headers=headers, data=json.dumps(json_dados))
        response.raise_for_status()
    except requests.exceptions.HTTPError as errh:
        try:
            error_message = response.json().get('message', 'Erro desconhecido')
            print(f"Falha ao enviar o JSON: ({nome_arquivo_json}). Status: {response.status_code}, Resposta: {error_message}")
        except ValueError:
            print(f"Falha ao enviar o JSON: ({nome_arquivo_json}). Status: {response.status_code}, Resposta: {response.text}")
    except requests.exceptions.ConnectionError as errc:
        print("Error de conexão:", errc)
    except requests.exceptions.Timeout as errt:
        print("Timeout Error:", errt)
    except requests.exceptions.RequestException as err:
        print("OOps: algo estranho aconteceu", err)
    else:
        print(f"JSON: ({nome_arquivo_json}) enviado com sucesso!")

# Função para enviar JSONs da pasta 'json' dentro de 'dados'
def enviar_jsons_da_pasta():
    json_files = [f for f in os.listdir(pasta_json) if f.endswith('.json')]
    for json_file in json_files:
        with open(os.path.join(pasta_json, json_file), 'r') as f:
            json_dados = json.load(f)
        enviar_json(json_dados, json_file)

# Processar e salvar JSONs para todos os arquivos Excel na pasta 'dados'
for arquivo in arquivos_excel:
    salvar_json_local(arquivo, pasta_json)

# Após salvar os JSONs localmente, enviar os JSONs
enviar_jsons_da_pasta()

print("\n\nTudo pronto, pode verificar no site do SGO.")
print("\nAté a próxima ;)")