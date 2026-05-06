import pytz
from selenium import webdriver
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import json
import re

from datetime import datetime
import openpyxl

def extrair_info_calendario():
    print("Abrindo o navegador")
    firefox_options = Options()
    firefox_options.add_argument('--headless')
    firefox_options.binary_location = "C:\\Program Files (x86)\\Mozilla Firefox\\firefox.exe"  # Adicione o caminho para o executável do Firefox

    driver = webdriver.Firefox(options=firefox_options)

    # url = 'https://www.espn.com.br/nba/estatisticas'
    url = 'https://www.espn.com.br/nba/calendario'
    driver.get(url)
    print("Navegador aberto")

    try:
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//script[contains(text(), "__espnfitt__")]'))
        )

        # Use XPath para encontrar o script desejado
        script_element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//script[contains(text(), "__espnfitt__")]'))
        )
        print("XPath encontrado")

        # Extraia o conteúdo do script
        script_content = script_element.get_attribute('innerHTML')
        espnfitt_match = re.search(r'window\[\'__espnfitt__\'\]=(\{.*?\});', script_content)

        print("Conteúdo extraído")

        if espnfitt_match:
            espnfitt_json = espnfitt_match.group(1)
            print("Json Carregado com sucesso")
            espnfitt_data = json.loads(espnfitt_json)
            print("Content do Json carregado")

            # Access the list of events
            events_list = espnfitt_data['page']['content']['events']

            # Nome do arquivo txt para salvar as informações
            nome_arquivo_txt = 'TXT/informacoes_eventos.txt'
            # Nome do arquivo Excel para salvar as informações
            nome_arquivo_excel = 'Excel/informacoes_eventos.xlsx'

            # Obtém a data atual
            data_atual = datetime.now().strftime("%Y%m%d")

            # Cria um novo arquivo Excel
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Jogos de Hoje"
            
            # Cria as colunas no arquivo Excel
            columns = ["Competidores", "Data do Evento", "Horário do Evento", "Local do Evento"]
            ws.append(columns)

            # Itera sobre cada evento
            for date, events in events_list.items():
                # Converte a data do evento para o formato desejado ("%Y%m%d")
                data_do_evento = datetime.strptime(date, "%Y%m%d").strftime("%Y%m%d")

                # Verifica se a data do evento é igual à data atual
                if data_do_evento == data_atual:
                    for event in events:
                        id_do_evento = event['id']

                        # Modificação: Extrai os nomes das equipes
                        team_names = [competitor['name'] for competitor in event['competitors']]
                        # Formata as equipes como "Equipe A vs Equipe B"
                        teams_formatted = f"{team_names[0]} vs {team_names[1]}"

                        data_do_evento = datetime.strptime(date, "%Y%m%d").strftime("%d-%m-%Y")
                        local_do_evento = event['venue']['address']['city']
                        
                        # Extrai e converte o horário
                        horario_raw = event['status']['detail']
                        horario_match = re.search(r"at (\d+:\d+ [AP]M)", horario_raw)
                        if horario_match:
                            horario_bruto = horario_match.group(1)
                            horario_obj = datetime.strptime(horario_bruto, "%I:%M %p").replace(
                                                                year=datetime.now().year, 
                                                                month=datetime.now().month, 
                                                                day=datetime.now().day
                                                            )

                            est_tz = pytz.timezone("US/Eastern")
                            horario_com_fuso = est_tz.localize(horario_obj)
                            brasilia_tz = pytz.timezone("America/Sao_Paulo")
                            horario_brasilia = horario_com_fuso.astimezone(brasilia_tz)

                            horario_do_evento = horario_brasilia.strftime("%H:%M")
                        else:
                            horario_do_evento = "Horário não disponível"

                        # Escreve as informações no arquivo de texto
                        with open(nome_arquivo_txt, 'a', encoding='utf-8') as arquivo_txt:
                            arquivo_txt.write(f"Competidores: {teams_formatted}\n")  # Usando a nova formatação
                            arquivo_txt.write(f"Data do Evento: {data_do_evento}\n")
                            arquivo_txt.write(f"Horário do Evento: {horario_do_evento}\n")
                            arquivo_txt.write(f"Local do Evento: {local_do_evento}\n\n")

                            print(f"Competidores: {teams_formatted}")  # Usando a nova formatação
                            print(f"Data do Evento: {data_do_evento}")
                            print(f"Horário do Evento: {horario_do_evento}")
                            print(f"Local do Evento: {local_do_evento}\n")

                        # Escreve as informações no arquivo Excel
                        ws.append([teams_formatted, data_do_evento, horario_do_evento, local_do_evento])

            # Salva o arquivo Excel
            wb.save(nome_arquivo_excel)

            print(f"Arquivo Excel '{nome_arquivo_excel}' criado com sucesso.")

        else:
            print("JSON de '__espnfitt__' não encontrado no script")

    except Exception as e:
        print(f"Erro: {e}")

    finally:
        driver.quit()

