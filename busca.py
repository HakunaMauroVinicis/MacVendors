import pdfplumber
import pandas as pd
import re
import requests
from time import sleep

linha = 0
pagina = 0

# Caminho para o arquivo PDF
pdf_path = r"C:\Users\Colaborador\Downloads\Mauro\IPV6-clientes\ClientesAracatubaB.pdf"

# Inicializar uma lista para armazenar os dados
data = []

# Definir os nomes das colunas
columns = ["ID Cliente", "Login", "Senha", "IP", "Concentrador", "MAC", "Status", "Marca"]

# Expressão regular para identificar padrões de IP, MAC e Status
ip_pattern = re.compile(r'\b(?:[0-9]{1,3}\.){3}[0-9]{1,3}\b')
mac_pattern = re.compile(r'\b(?:[0-9A-Fa-f]{2}:){5}[0-9A-Fa-f]{2}\b')
status_pattern = re.compile(r'\b(?:On-line|Off-line)\b')

# Função para consultar a marca pelo endereço MAC
def get_mac_brand(mac_address):
    url = f"http://www.macvendorlookup.com/api/v2/{mac_address}"
    response = requests.get(url)
    if response.status_code == 200:
        data = response.json()
        if data and isinstance(data, list):
            return data[0].get('company', 'Unknown')
    return 'Unknown'

with pdfplumber.open(pdf_path) as pdf:
    # Iterar por cada página do PDF
    for page in pdf.pages:
        print('pagina: ', pagina)
        # Extrair o texto da página
        text = page.extract_text()
        if text:
            # Dividir o texto em linhas
            lines = text.split('\n')
            for line in lines:
                print('linha:', linha)
                # Encontrar IP e MAC na linha
                ip_match = ip_pattern.search(line)
                mac_match = mac_pattern.search(line)
                status_match = status_pattern.search(line)

                if ip_match and mac_match and status_match:
                    # Dividir a linha com base nos padrões
                    parts = line.split()
                    ip = ip_match.group()
                    mac = mac_match.group()
                    status = status_match.group()

                    # Remover IP, MAC e Status dos 'parts' para evitar duplicação
                    parts.remove(ip)
                    parts.remove(mac)
                    parts.remove(status)

                    # Assegurar que restam 4 partes (ID Cliente, Login, Senha, Concentrador)
                    if len(parts) >= 4:
                        id_cliente = parts[0]
                        login = parts[1]
                        senha = parts[2]
                        concentrador = ' '.join(parts[3:])

                        # Obter a marca do dispositivo pelo MAC address
                        marca = get_mac_brand(mac)
                        
                        # Adicionar um atraso entre as consultas
                        sleep(2)

                        # Adicionar os dados na lista
                        data.append([id_cliente, login, senha, ip, concentrador, mac, status, marca])
                linha = linha + 1
        pagina = pagina + 1

# Converter os dados em um DataFrame do pandas
df = pd.DataFrame(data, columns=columns)

# Salvar o DataFrame em um arquivo Excel
excel_path = r"C:\Users\Colaborador\Downloads\Mauro\IPV6-clientes\ClientesAracatubaB.xlsx"
df.to_excel(excel_path, index=False)

print(f"Dados salvos com sucesso em {excel_path}")
