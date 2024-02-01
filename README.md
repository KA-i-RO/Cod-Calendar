# Cod-Calendar
import requests
import time
import pandas as pd
import getpass
import os
from dotenv import load_dotenv

load_dotenv()

pipefy_token = os.getenv("PIPEFY_TOKEN")

pipeId = "303683960"
relatorioID = "300551552"
nome_arquivo_final = r'C:\Users\kairo_szczepkowski\OneDrive - Sicredi\Público\Database\19 Calendário Compartilhado\base_calendario_pipe.xlsx'.replace('kairo_szczepkpowski', getpass.getuser(), 1)

url = "https://api.pipefy.com/graphql"

# Mutation para exportar o pipeReport
mutation_query = f'''
mutation {{
    exportPipeReport(input: {{
        pipeId: "{pipeId}",
        pipeReportId: "{relatorioID}"
    }}) {{
        pipeReportExport {{
            id
        }}
    }}
}}
'''

headers = {
    "authorization": f"Bearer {pipefy_token}",
    "content-type": "application/json"
}

# Enviar a primeira consulta para exportar o pipeReport
response = requests.post(url, json={'query': mutation_query}, headers=headers)

if response.status_code == 200:
    result = response.json()
    print("está aqui o erro", result)
    pipe_report_export_id = result['data']['exportPipeReport']['pipeReportExport']['id']

    # Consulta para obter o pipeReportExport
    query = f'''
    {{
    pipeReportExport(id: "{pipe_report_export_id}") {{
        fileURL
        state
        startedAt
        requestedBy {{
        id
        }}
    }}
    }}
    '''

    done = False

    # Loop para verificar o estado do pipeReportExport até que seja concluído
    while not done:
        response = requests.post(url, json={'query': query}, headers=headers)
        
        if response.status_code == 200:
            result = response.json()
            pipe_report_export = result['data']['pipeReportExport']
            state = pipe_report_export['state']
            
            if state == 'done':
                file_url = pipe_report_export['fileURL']
                print(f"URL do arquivo: {file_url}")
                done = True

                exportPipe = pd.read_excel(file_url)
                exportPipe.to_excel(nome_arquivo_final, index=False)
                print('Arquivo gerado com sucesso!')

            elif state == 'processing':
                print("Ainda processando o pipeReportExport...")
                time.sleep(1)  # Esperar 1 segundo antes de fazer a próxima verificação
            else:
                print("Erro ao exportar o pipeReport.")
                break
        else:
            print(f"Falha ao obter os detalhes do pipeReportExport. Código de status: {response.status_code}")
            break
else:
    print(f"Falha ao exportar o pipeReport. Código de status: {response.status_code}")
