from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import json
import re
import pandas as pd

def extrair_info_estatistica(url, estatisticas, quantidade=5):
    print("Abrindo o navegador")
    firefox_options = Options()
    firefox_options.add_argument('--headless')
    
    driver = webdriver.Firefox(options=firefox_options)

    driver.get(url)
    print("Navegador aberto para:", url)

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
            print("Content do Json carregado")

            # Especificando o caminho do arquivo de texto
            caminho_arquivo = f'TXT/{estatisticas.lower()}.txt'

            # Extrair informações dos 5 primeiros jogadores
            jogadores_info = espnfitt_data['page']['content']['statistics']['playerStats'][:quantidade]
            with open(caminho_arquivo, 'w', encoding='utf-8') as arquivo_txt:
                for i, jogador in enumerate(jogadores_info, start=1):
                    # Acessa as informações do 'athlete' para cada jogador individualmente
                    athlete_info = jogador['athlete']
                    arquivo_txt.write(f"Informações do Atleta {i}:\n")
                    arquivo_txt.write(f"Nome do Jogador: {athlete_info['name']}\n")
                    arquivo_txt.write(f"Time do Jogador: {athlete_info['team']}\n")
                    arquivo_txt.write(f"Rank: {athlete_info['rank']}\n")

                    # Acessa 'stats' para cada jogador individualmente
                    stats_jogador = jogador['stats']
                    arquivo_txt.write(f"Estatísticas do Jogador {i}:\n")

                    # Itera sobre a lista de estatísticas do jogador
                    for stat in stats_jogador:
                        # Verifica se a estatística é a desejada (doubleDouble ou tripleDouble)
                        if stat['name'] == estatisticas:
                            arquivo_txt.write(f"Estatística: {estatisticas}\n")
                            arquivo_txt.write(f"Quantidade: {stat['value']}\n")
    
            print(f"As informações dos 5 primeiros jogadores foram salvas em {caminho_arquivo}")

            # Retorna os dados extraídos
            return jogadores_info

    finally:
        # Feche o navegador
        driver.quit()

def criar_excel_com_abas(nome_arquivo_txt, estatisticas):
    jogadores_info = []
    with open(nome_arquivo_txt, 'r') as arquivo:
        linhas = arquivo.readlines()

    jogador_info = {}
    for linha in linhas:
        if linha.startswith("Nome do Jogador"):
            jogador_info = {}
            jogador_info["Nome do Jogador"] = linha.split(": ")[1].strip()
        elif linha.startswith("Time do Jogador"):
            jogador_info["Time do Jogador"] = linha.split(": ")[1].strip()
        elif linha.startswith("Rank"):
            jogador_info["Rank"] = int(linha.split(": ")[1].strip())
        elif linha.startswith("Quantidade"):
            jogador_info["Quantidade"] = int(linha.split(": ")[1].strip())
            jogadores_info.append(jogador_info)


    # Define os nomes das abas
    aba = estatisticas.capitalize()

    # Especifica o caminho do arquivo Excel
    caminho_excel = f'Excel/{estatisticas.lower()}.xlsx'

    # Define as colunas para cada estatística
    colunas = ["Rank", "Nome do Jogador", "Time do Jogador","Quantidade"]

    # Cria DataFrame a partir dos dados extraídos
    df = pd.DataFrame(jogadores_info, columns=colunas)

    # Salva o DataFrame no arquivo Excel
    df.to_excel(caminho_excel, index=False, sheet_name=aba)

    print(f"As informações foram consolidadas em {caminho_excel}")

def main():
    # Chamar a função extrair_info_estatistica para o URL de doubleDouble (limitado aos 5 primeiros jogadores)
    extrair_info_estatistica('https://www.espn.com.br/nba/estatisticas/jogador/_/table/general/ordenar/doubleDouble/dir/desce', 'doubleDouble', quantidade=5)

    # Chamar a função extrair_info_estatistica para o URL de tripleDouble (limitado aos 5 primeiros jogadores)
    extrair_info_estatistica('https://www.espn.com.br/nba/estatisticas/jogador/_/table/general/ordenar/tripleDouble/dir/desce', 'tripleDouble', quantidade=5)

    # Chamar a função criar_excel_com_abas com os dados dos arquivos TXT
    criar_excel_com_abas("TXT/doubledouble.txt", 'doubleDouble')
    criar_excel_com_abas("TXT/tripledouble.txt", 'tripleDouble')

if __name__ == "__main__":
    main()
