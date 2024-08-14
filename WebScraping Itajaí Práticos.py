import time
import pyodbc
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By

# Definir o caminho para o ChromeDriver
selenium_service = Service('')
driver = webdriver.Chrome(service=selenium_service)

# Definir a URL e o intervalo de atualização em segundos
url = 'https://www.itajaipraticos.com.br/'
intervalo_atualizacao = 60  # 1 minuto

# Definir os detalhes de conexão ao SQL Server
server = 'CPTSC-SI104\SQLEXPRESS'
database = 'itajaiPraticos'
conn_str = 'DRIVER={SQL Server};SERVER=' + server + ';DATABASE=' + database + ';Trusted_Connection=yes'

# Criação da conexão com o banco de dados SQL Server
conn = pyodbc.connect(conn_str)
cursor = conn.cursor()

# Acessar o site e extrair os dados iniciais
driver.get(url)

table = driver.find_element(By.XPATH, '/html/body/div/div[2]/div/div/main/article/div/div/div/div/section[3]/div/div/div/div/div/div[3]/div/table')

# Extrair as linhas da tabela e criar uma lista de dicionários para os dados
data = []
rows = table.find_elements(By.TAG_NAME, 'tr')
for row in rows:
    cells = row.find_elements(By.TAG_NAME, 'td')
    if len(cells) > 0:
        row_data = {
            'Data': cells[0].text,
            'Horario': cells[1].text,
            'Manobra': cells[2].text,
            'Berco': cells[3].text,
            'Bordo': cells[4].text,
            'Navio': cells[5].text,
            'Loa': cells[6].text,
            'Boca': cells[7].text,
            'Calado': cells[8].text,
            'Situacao': cells[9].text
        }
        data.append(row_data)
        print(row_data)

# Verificar se a tabela já existe
if cursor.tables(table='tabela_itajai', tableType='TABLE').fetchone():
    print("A tabela já existe.")
else:
    # Criar a tabela se não existir
    cursor.execute('''CREATE TABLE tabela_itajai (
                        
                        Data VARCHAR(10),
                        Horario VARCHAR(10),
                        Manobra VARCHAR(50),
                        Berco VARCHAR(50),
                        Bordo VARCHAR(50),
                        Navio VARCHAR(50),
                        Loa VARCHAR(10),
                        Boca VARCHAR(10),
                        Calado VARCHAR(50),
                        Situacao VARCHAR(50)
                    )''')
    print("Tabela criada com sucesso")

# Excluir os registros existentes na tabela
cursor.execute('DELETE FROM tabela_itajai')
print("Registros excluídos da tabela.")

# Inserir os valores atualizados na tabela do banco de dados
for row in data:
    cursor.execute('INSERT INTO tabela_itajai (Data, Horario, Manobra, Berco, Bordo, Navio, Loa, Boca, Calado, Situacao) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)',
                   (row['Data'], row['Horario'], row['Manobra'], row['Berco'], row['Bordo'], row['Navio'], row['Loa'], row['Boca'], row['Calado'], row['Situacao']))
    print("Valores inseridos:", row)

# Confirmar as alterações
conn.commit()

while True:
    # Acessar o site e extrair os dados atualizados
    driver.get(url)

    table = driver.find_element(By.XPATH, '/html/body/div/div[2]/div/div/main/article/div/div/div/div/section[3]/div/div/div/div/div/div[3]/div/table')

    # Extrair as linhas da tabela e atualizar os valores correspondentes na tabela do banco de dados
    rows = table.find_elements(By.TAG_NAME, 'tr')
    new_data = []
    for row in rows:
        cells = row.find_elements(By.TAG_NAME, 'td')
        if len(cells) > 0:
            row_data = {
                'Data': cells[0].text,
                'Horario': cells[1].text,
                'Manobra': cells[2].text,
                'Berco': cells[3].text,
                'Bordo': cells[4].text,
                'Navio': cells[5].text,
                'Loa': cells[6].text,
                'Boca': cells[7].text,
                'Calado': cells[8].text,
                'Situacao': cells[9].text
            }
            new_data.append(row_data)
            print(row_data)

    # Excluir os registros existentes na tabela
    cursor.execute('DELETE FROM tabela_itajai')
    print("Registros excluídos da tabela.")

    # Inserir os valores atualizados na tabela do banco de dados
    for row in new_data:
        cursor.execute('INSERT INTO tabela_itajai (Data, Horario, Manobra, Berco, Bordo, Navio, Loa, Boca, Calado, Situacao) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)',
                       (row['Data'], row['Horario'], row['Manobra'], row['Berco'], row['Bordo'], row['Navio'], row['Loa'], row['Boca'], row['Calado'], row['Situacao']))
        print("Valores inseridos:", row)

    # Confirmar as alterações
    conn.commit()

    # Aguardar o intervalo de atualização
    time.sleep(intervalo_atualizacao)
