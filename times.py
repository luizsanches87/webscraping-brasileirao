import requests
import openpyxl
import json
from pathlib import Path

# Abrindo arquivo times.json que contem os IDs dos times do site Sofascore para realizar as chamadas no requests
with open('times.json') as f:
    jsondata = json.load(f)

# Criando a planilha
book = openpyxl.Workbook()
book.create_sheet("times")
time_page = book["times"]

# Inserindo cabeçalho na primeira linha da planilha
time_page.append(['Time', 'Partidas', 'Gols Marcados', 'Gols Sofridos', 'Media de Gols', 'Media de Finalizacoes',
                  'Media de Finalizacoes no Gol', 'Media de Escanteios', 'Media Gols Sofridos',
                  'Media de Impedimentos', 'Media de Faltas Cometidas', 'Cartoes Amarelos',
                  'Cartoes Vermelhos', 'Media de Cartoes Amarelos', 'Media de Cartoes Vermelhos',
                  'Media de Finalizacoes no Gol Adversario', 'Media de Finalizacoes Adversario',
                  'Media de Escanteios Adversario', 'Media de Cartoes Amarelos Adversario',
                  'Media de Cartoes Vermelhos Adversario', 'Media de Impedimentos Adversario'])

# Coletando dados dos times utilizando a biblioteca requests e inserindo nas linhas da planilha utilizando a biblioteca openpyxl
for time in jsondata:
    id = time['id']
    name = time['name']
    # Montando a chamada para ser realizada no requests
    url = f'https://www.sofascore.com/api/v1/team/{id}/unique-tournament/325/season/58766/statistics/overall'
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/113.0.0.0 Safari/537.36'}
    # Coletando os dados do time atraves do requests.get
    site = requests.get(url, headers=headers)
    # Criando um dicionario com os dados coletados no requests.get
    dados_time = json.loads(site.content)

    # Gravando em variaveis os dados coletados no dicionario de dados, realizando as medias necessarias
    partidas = dados_time['statistics']['matches']
    gols_marcados = dados_time['statistics']['goalsScored']
    gols_sofridos = dados_time['statistics']['goalsConceded']
    media_gols_partida = gols_marcados / partidas
    media_chutes_partida = dados_time['statistics']['shots'] / partidas
    media_chutes_gol_partida = dados_time['statistics']['shotsOnTarget'] / partidas
    media_escanteios = dados_time['statistics']['corners'] / partidas
    media_gols_sofridos = gols_sofridos / partidas
    media_impedimentos = dados_time['statistics']['offsides'] / partidas
    media_faltas_cometidas = dados_time['statistics']['fouls'] / partidas
    total_cartoes_amarelos = dados_time['statistics']['yellowCards']
    total_cartoes_vermelhos = dados_time['statistics']['redCards']
    media_cartoes_amarelos = total_cartoes_amarelos / partidas
    media_cartoes_vermelhos = total_cartoes_vermelhos / partidas
    media_chutes_gol_adversario = dados_time['statistics']['shotsOnTargetAgainst'] / partidas
    media_chutes_adversario = dados_time['statistics']['shotsAgainst'] / partidas
    media_escanteios_adversario = dados_time['statistics']['cornersAgainst'] / partidas
    media_cartoes_amarelos_adversario = dados_time['statistics']['yellowCardsAgainst'] / partidas
    media_cartoes_vermelhos_adversario = dados_time['statistics']['redCardsAgainst'] / partidas
    media_impedimentos_adversario = dados_time['statistics']['offsidesAgainst'] / partidas

    # Inserindo na linha da planilha do time os dados das variaveis
    time_page.append([name, partidas, gols_marcados, gols_sofridos, round(media_gols_partida, 1),
                      round(media_chutes_partida, 1), round(
                          media_chutes_gol_partida, 1), round(media_escanteios, 1),
                      round(media_gols_sofridos, 1), round(
                          media_impedimentos, 1), round(media_faltas_cometidas, 1),
                      total_cartoes_amarelos, total_cartoes_vermelhos, round(
                          media_cartoes_amarelos, 1),
                      round(media_cartoes_vermelhos, 1), round(
                          media_chutes_gol_adversario, 1),
                      round(media_chutes_adversario, 1), round(
                          media_escanteios_adversario, 1),
                      round(media_cartoes_amarelos_adversario, 1), round(
                          media_cartoes_vermelhos_adversario, 1),
                      round(media_impedimentos_adversario, 1)])

    # Salvando a planilha
    book.save(
        f"{Path.home()}/Desktop/webscraping_brasileirao/estatistica_times_br24.xlsx")
