import pandas as pd
from selenium import webdriver
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import json
import re
from datetime import datetime

def extrair_info_classificacao():
    print("Abrindo o navegador")
    firefox_options = Options()
    firefox_options.add_argument('--headless')

    driver = webdriver.Firefox(options=firefox_options)

    url = 'https://www.espn.com.br/nba/classificacao'
    driver.get(url)
    print("Navegador aberto")

    try:
        # Aguarde o elemento desejado carregar
        script_element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//script[contains(text(), "__espnfitt__")]'))
        )
        print("Elemento encontrado")

        # Extraia o conteúdo do script
        script_content = script_element.get_attribute('innerHTML')
        espnfitt_match = re.search(r'window\[\'__espnfitt__\'\]=(\{.*?\});', script_content)

        if espnfitt_match:
            espnfitt_json = espnfitt_match.group(1)
            print("Json Carregado com sucesso")
            espnfitt_data = json.loads(espnfitt_json)
            print("Content do Json carregado")

            # Especificando o caminho do arquivo de texto
            caminho_arquivo = 'TXT/classificacao_nba.txt'

            # Convertendo a estrutura combinada para uma string formatada
            json_formatado = json.dumps(espnfitt_data['page']['content']['standings']['groups']['groups'], indent=2, ensure_ascii=False)

            # Escrevendo o conteúdo no arquivo
            with open(caminho_arquivo, 'w', encoding='utf-8') as arquivo:
                arquivo.write(json_formatado)

            print(f"\nConteúdo do '__espnfitt__' também salvo em '{caminho_arquivo}'")

            try:
                # Criar listas para armazenar os dados temporariamente
                data = []

                # Preencher listas com informações Conferência Leste
                for group in espnfitt_data['page']['content']['standings']['groups']['groups']:
                    if group['name'] == "Conferência Leste":
                        conference = group['name']
                        standings = group['standings']
                        
                        for position, leader in enumerate(standings, start=1):  # 'position' começa em 1
                            team = leader['team']['displayName']
                            stats = leader['stats']  # Lista de estatísticas

                            # Extração das métricas
                            victories = stats[15]
                            defeats = stats[16]
                            win_percentage = stats[13]
                            games_behind = stats[4]
                            home_record = stats[17]
                            away_record = stats[18]
                            division_record = stats[19]
                            conference_record = stats[20]
                            points_per_game = stats[1]
                            opponent_points_per_game = stats[0]
                            point_difference = stats[2]
                            streak = stats[12]
                            last_10_games = stats[20]

                            # Adicionando ao DataFrame
                            data.append({
                                'Posição': position,
                                'Conferencia': conference,
                                'Nome do Time': team,
                                'Vitórias': victories,
                                'Derrotas': defeats,
                                '% Vitórias': win_percentage,
                                'Jogos Atrás': games_behind,
                                'Casa': home_record,
                                'Visitante': away_record,
                                'Divisão': division_record,
                                'Conferência': conference_record,
                                'PTS': points_per_game,
                                'PTS Contra': opponent_points_per_game,
                                'DIF': point_difference,
                                'Sequência': streak,
                                'Últimos 10 Jogos': last_10_games
                            })

                # Preencher listas com informações Conferência Oeste
                for group in espnfitt_data['page']['content']['standings']['groups']['groups']:
                    if group['name'] == "Conferência Oeste":
                        conference = group['name']
                        standings = group['standings']
                        
                        for position, leader in enumerate(standings, start=1):  # 'position' começa em 1
                            team = leader['team']['displayName']
                            stats = leader['stats']  # Lista de estatísticas

                            # Extração das métricas
                            victories = stats[15]
                            defeats = stats[16]
                            win_percentage = stats[13]
                            games_behind = stats[4]
                            home_record = stats[17]
                            away_record = stats[18]
                            division_record = stats[19]
                            conference_record = stats[20]
                            points_per_game = stats[1]
                            opponent_points_per_game = stats[0]
                            point_difference = stats[2]
                            streak = stats[12]
                            last_10_games = stats[20]

                            # Adicionando ao DataFrame
                            data.append({
                                'Posição': position,
                                'Conferencia': conference,
                                'Nome do Time': team,
                                'Vitórias': victories,
                                'Derrotas': defeats,
                                '% Vitórias': win_percentage,
                                'Jogos Atrás': games_behind,
                                'Casa': home_record,
                                'Visitante': away_record,
                                'Divisão': division_record,
                                'Conferência': conference_record,
                                'PTS': points_per_game,
                                'PTS Contra': opponent_points_per_game,
                                'DIF': point_difference,
                                'Sequência': streak,
                                'Últimos 10 Jogos': last_10_games
                            })

                # Criar DataFrame do Pandas com os dados coletados
                df_geral = pd.DataFrame(data)

                # Salvar DataFrame em arquivo Excel
                caminho_excel = 'Excel/classificacao_nba.xlsx'

                # Cria um escritor Excel
                with pd.ExcelWriter(caminho_excel, engine='xlsxwriter') as writer:
                    # Filtra e salva cada tabela sem a coluna 'Conferencia'
                    for tipo, tabela in df_geral.groupby('Conferencia'):
                        tabela_sem_tipo = tabela.drop(columns=['Conferencia'])
                        tabela_sem_tipo.to_excel(writer, sheet_name=tipo, index=False)
                print(f"\nConferencias salvas em '{caminho_excel}'")
            except KeyError as e:
                print(f"Erro: A chave {e} não foi encontrada na estrutura do JSON. Verifique se a estrutura mudou.")
        else:
            print("JSON de '__espnfitt__' não encontrado no script")

    except Exception as e:
        print(f"Erro: {e}")

    finally:
        driver.quit()


