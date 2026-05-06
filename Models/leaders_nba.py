from selenium import webdriver
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import re
import pandas as pd
import json

def extrair_info_lideres():
    print("Abrindo o navegador")
    firefox_options = Options()
    firefox_options.add_argument('--headless')

    driver = webdriver.Firefox(options=firefox_options)

    url = 'https://www.espn.com.br/nba/estatisticas'
    driver.get(url)
    print("Navegador aberto")

    try:
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, 'espnfitt'))
        )
        print("Estrutura da página foi carregada com sucesso")
    except TimeoutException:
        print("Timeout: A estrutura da página não foi carregada dentro do tempo limite.")

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
            print("Conteúdo do Json carregado")

            # Especificando o caminho do arquivo de texto
            caminho_arquivo = 'TXT/conteudo_espnfitt.txt'

            # Convertendo a estrutura combinada para uma string formatada
            json_formatado = json.dumps(espnfitt_data['page']['content']['statistics']['leaders'], indent=2, ensure_ascii=False)

            # Escrevendo o conteúdo no arquivo
            with open(caminho_arquivo, 'w', encoding='utf-8') as arquivo:
                arquivo.write(json_formatado)

                print(f"\nConteúdo do '__espnfitt__' também salvo em '{caminho_arquivo}'")

            try:

                leaders_ofensivos = espnfitt_data['page']['content']['statistics']['leaders']['0']['groups']
                leaders_defensivos = espnfitt_data['page']['content']['statistics']['leaders']['1']['groups']

                # Cria um escritor Excel
                caminho_excel = 'Excel/leaders_nba.xlsx'
                writer = pd.ExcelWriter(caminho_excel, engine='xlsxwriter')

                # Função para adicionar tabelas em uma única aba
                def adicionar_tabelas_aba(writer, nome_aba, grupos):
                    workbook = writer.book
                    worksheet = workbook.add_worksheet(nome_aba)
                    writer.sheets[nome_aba] = worksheet

                    linha_inicial = 0  # Linha inicial para começar a escrever os dados

                    for group in grupos:
                        header = group['header']
                        leaders = group['leaders']

                        # Criar DataFrame para o grupo atual
                        tabela_grupo = pd.DataFrame([
                            {
                                'Rank': leader['rank'],
                                'Nome de Jogador': leader['name'],
                                'Time': leader['team'],
                                'Valor': leader['statValue']
                            }
                            for leader in leaders
                        ])

                        # Escreve o cabeçalho do grupo
                        worksheet.write(linha_inicial, 0, header)
                        linha_inicial += 1  # Pula uma linha

                        # Escreve os dados do grupo como tabela
                        for r, row in tabela_grupo.iterrows():
                            for c, value in enumerate(row):
                                worksheet.write(linha_inicial + r, c, value)

                        # Avança linhas para próxima tabela
                        linha_inicial += len(tabela_grupo) + 2  # Pula mais 2 linhas para separar tabelas

                # Adiciona os líderes ofensivos na aba correspondente
                adicionar_tabelas_aba(writer, 'Líderes Ofensivos', leaders_ofensivos)

                # Adiciona os líderes defensivos na aba correspondente
                adicionar_tabelas_aba(writer, 'Líderes Defensivos', leaders_defensivos)

                # Fecha o escritor Excel
                writer.close()

                print(f"\nLíderes salvos em '{caminho_excel}'")

            except KeyError as e:
                print(f"Erro: A chave {e} não foi encontrada na estrutura do JSON. Verifique se a estrutura mudou.")
        else:
            print("JSON de '__espnfitt__' não encontrado no script")

    except Exception as e:
        print(f"Erro: {e}")

    finally:
        driver.quit()

