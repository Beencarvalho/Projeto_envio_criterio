import pandas as pd
import requests
import json
import os
from util.api_token import url,headers

# Caminho para a pasta que contém os arquivos Excel
pasta_arquivos = 'data'
# Definindo um ID inicial
ultimo_id_criterio = int(input("Digite o numero do ID do ultimo criterio cadastrado: "))
# Listar todos os arquivos Excel na pasta
arquivos_excel = [os.path.join(pasta_arquivos, f) 
                  for f in os.listdir(pasta_arquivos) 
                  if f.endswith('.xlsx' or '.xls') 
                  and not f.startswith('~$')]

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
            "codeCostCenter": str(centro_custo),
            "company": {
                "active": True,
                "id": id_empresa,  # ID tratado
                "name": name_emp,  # Nome da empresa tratado
                "code": code_emp   # Código tratado
            },
            "companyId": id_empresa
        },
        "orderBy": None
    }

# Função para processar um único arquivo Excel
def processar_arquivo(file_path, ultimo_id_criterio):
    # Nome de cada arquivo para ser o "name"
    nome_arquivo = os.path.splitext(os.path.basename(file_path))[0]

    # Atualizar o ID sequencial somando 1
    id_atual = ultimo_id_criterio + 1

    # Ler o arquivo Excel, removendo as 8 primeiras linhas e mantendo as 6 primeiras colunas
    df = pd.read_excel(file_path, skiprows=8)
    df = df.iloc[:-1, :6] # Manter apenas as 6 primeiras colunas e todas as linhas exeto a ultima. 

    # Renomear as colunas para facilitar o acesso (com base nas colunas identificadas)
    df.columns = ['empresa', 'cod_setor', 'centro_custo', 'setor', 'base', 'distribuicao']

    # Criar a lista de itens para a requisição
    id_item = id_atual
    itens = [criar_item_json(linha, id_item) for _, linha in df.iterrows()]

    # Montando o JSON final para a requisição
    json_final = {
        "active": True,
        "id": id_atual,
        "name": nome_arquivo,
        "description": "Adicione uma descricao aqui",
        "items": itens,
        "cycle": None,
        "cycleId": 5,
        "baseSum": None,
        "totalItems": None
    }

    return json_final

# Função para processar o arquivo e salvar o JSON na pasta
def salvar_json_local(arquivo, ultimo_id_criterio):
    json_resultado = processar_arquivo(arquivo, ultimo_id_criterio)
    
    # Nome do arquivo JSON a ser salvo
    nome_arquivo_json = os.path.splitext(os.path.basename(arquivo))[0] + '.json'
    
    # Salvar o JSON em um arquivo local
    with open(f"jsons/{nome_arquivo_json}", 'w') as f:
        json.dump(json_resultado, f, indent=4)
    
    return nome_arquivo_json  # Retorna o nome do arquivo salvo

# Função para enviar um JSON via POST
def enviar_json(json_dados, nome_arquivo_json):
    try:
        response = requests.post(url, headers=headers, data=json.dumps(json_dados))
        response.raise_for_status()  # Isso levanta um erro se o status code for 4xx ou 5xx

    except requests.exceptions.HTTPError as errh:
        try:
            # Tentando capturar a mensagem de erro do JSON retornado pela API
            error_message = response.json().get('message', 'Erro desconhecido')
            print(f"Falha ao enviar o JSON:  ({nome_arquivo_json}).   Status: {response.status_code}, Resposta: {error_message}")
        except ValueError:
            # Caso o JSON não esteja disponível ou seja inválido
            print(f"Falha ao enviar o JSON:  ({nome_arquivo_json}).   Status: {response.status_code}, Resposta: {response.text}")
    except requests.exceptions.ConnectionError as errc:
        print("Error de conexão:", errc)
    except requests.exceptions.Timeout as errt:
        print("Timeout Error:", errt)
    except requests.exceptions.RequestException as err:
        print("OOps: algo estranho aconteceu", err)
    else:
        print(f"JSON:  ({nome_arquivo_json})  enviado com sucesso!")


# Função para enviar JSONs da pasta jsons/
def enviar_jsons_da_pasta():
    json_files = [f for f in os.listdir('jsons/') if f.endswith('.json')]
    
    for json_file in json_files:
        # Carregar o JSON do arquivo
        with open(f"jsons/{json_file}", 'r') as f:
            json_dados = json.load(f)
        
        # Enviar o JSON carregado
        enviar_json(json_dados, json_file)

# Certifique-se de que a pasta 'jsons/' exista
if not os.path.exists('jsons/'):
    os.makedirs('jsons/')

# Processar e salvar JSONs para todos os arquivos Excel na pasta
for arquivo in arquivos_excel:
    salvar_json_local(arquivo, ultimo_id_criterio)
    ultimo_id_criterio += 1  # Incrementar o ID para o próximo arquivo

# Após salvar os JSONs localmente, você pode enviá-los
enviar_jsons_da_pasta()
