import os
import glob
import pandas as pd
import numpy as np
from pathlib import Path
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta


path = os.getcwd()

conteudo_colunas = ["Data", "Espectadores únicos", "Comentários adicionados", "Compartilhamentos", 'Marcações "Gostei"', "RPM (USD)", "Taxa de cliques de impressões (%)", "Impressões", "Inscritos", "Porcentagem visualizada média (%)", "Visualizações", "Tempo de exibição (horas)", "Duração média da visualização", "Receita estimada (USD)"]
conteudo_total_colunas = ["Data", "Espectadores únicos", "Comentários adicionados", "Compartilhamentos", 'Marcações "Gostei"', "RPM (USD)", "Taxa de cliques de impressões (%)", "Impressões", "Vídeos publicados", "Inscritos", "Porcentagem visualizada média (%)", "Visualizações", "Tempo de exibição (horas)", "Duração média da visualização", "Receita estimada (USD)"]
total_colunas = ["Data", 'Média de "Gostei" da postagem (%)', '"Gostei" da postagem', "Impressões da postagem", "Espectadores únicos", "Comentários adicionados", "Compartilhamentos", 'Marcações "Gostei"', "RPM (USD)", "Taxa de cliques de impressões (%)", "Impressões", "Vídeos publicados", "Inscritos", "Porcentagem visualizada média (%)", "Visualizações", "Tempo de exibição (horas)", "Duração média da visualização", "Receita estimada (USD)","CPM (USD)", "CPM baseado em exibição (USD)","Respostas à postagem","Inscrições obtidas","Inscrições perdidas"]


def salvar_arquivo_csv_novas_colunas(colunas_a_incluir: dict, colunas_padrao: list[str], arquivo: str, df: pd.DataFrame) -> None:
    for coluna in colunas_a_incluir:
        df[coluna] = None
    df = df[colunas_padrao]
    df.to_csv(arquivo, index=False)


def buscar_lista_artistas():
    # Acesso exports.txt para buscar o nome dos artistas
    lines = ''
    with open(path+'/exports.txt') as f:
        lines = f.readlines()
    lines = [i.rstrip() for i in lines]
    return lines


def ajustar_padrao_colunas(artista):
    """
    Ajusta e reordena as colunas dos arquivos CSV de acordo com padrões específicos
    baseados no tipo de conteúdo identificado pelo nome do arquivo.

    Parâmetros:
        artista (str): nome do artista ou pasta alvo dentro de /dados_full/
        path (str): caminho base até o diretório de dados
        total_colunas (list): colunas padrão para arquivos 'total'
        conteudo_total_colunas (list): colunas padrão para vídeos, lives e shorts
        conteudo_colunas (list): colunas padrão para outros arquivos
    """
    arquivos_csv = glob.glob(f"{path}/dados_full/{artista}/*.csv")

    for arquivo in arquivos_csv:
        df = pd.read_csv(arquivo)
        nome_arquivo = os.path.basename(arquivo).split('.')[0]

        if "Receita estimada (BRL)" in df.columns:
            print(f"⚠️ A extração de {artista} foi feita em Real (BRL)")

        # Ignora arquivos que não devem ser ajustados
        if nome_arquivo in ['origem_lives', 'origem_vods', 'comunidade']:
            continue

        # Define o conjunto de colunas conforme o tipo de arquivo
        if nome_arquivo == 'total':
            colunas_padrao = total_colunas
        elif nome_arquivo in ['lives', 'videos', 'shorts']:
            colunas_padrao = conteudo_total_colunas
        else:
            colunas_padrao = conteudo_colunas

        # Adiciona colunas ausentes com valor None
        for col in colunas_padrao:
            if col not in df.columns:
                df[col] = None

        # Mantém a ordem definida + quaisquer colunas extras ao final
        colunas_presentes = [col for col in colunas_padrao if col in df.columns]
        colunas_extras = [col for col in df.columns if col not in colunas_padrao]
        df_reordenado = df[colunas_presentes + colunas_extras]

        # Salva o arquivo sobrescrevendo o original
        df_reordenado.to_csv(arquivo, index=False)


def read_and_process_file(artista):
    df = pd.read_csv(f"dados_full/{artista}/total.csv")
    df.drop(index=0, inplace=True)
    df['Data'] = pd.to_datetime(df['Data'], format='%Y-%m')
    df = df.sort_values(by='Data', ascending=False)
    most_recent_month = df.iloc[0]['Data']
    months_list = []
    for i in range(7):
        months_list.append(most_recent_month.strftime("%Y-%m"))
        most_recent_month -= relativedelta(months=1)
    return months_list


def process_post_data(artista):
    arquivo = f"dados_full/{artista}/comunidade.csv"

    # 🔹 Carregar o CSV garantindo que o cabeçalho seja mantido
    df = pd.read_csv(arquivo)
    # 🔹 Se o DataFrame estiver vazio (apenas cabeçalho, sem dados), adiciona uma linha padrão
    if len(df) == 1:
        # Criar a nova linha com os valores padrão, usando as colunas do próprio CSV
        nova_linha = pd.DataFrame([[
            "Ugxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx",  # ID
            "Texto da postagem",  # Texto
            "Jan 01, 2001",  # Data
            0, 0, 0, 0, 0  # Métricas
        ]], columns=df.columns)  # 🔹 Mantém as colunas do próprio CSV

        # Adicionar a nova linha e salvar no arquivo
        df = pd.concat([df, nova_linha], ignore_index=True)
        df.to_csv(arquivo, index=False)


def completar_data(artista):
  arquivos_csv = glob.glob(path+'/dados_full/'+artista+'/*.csv')
  for arquivo in arquivos_csv:
    arq = pd.read_csv(arquivo)

    if 'Data' in arq.columns and not 'origem_' in arquivo and not 'comunidade' in arquivo:
        meses = arq['Data'].to_list()
        meses_esperados = read_and_process_file(artista)
        meses_faltantes = [m for m in meses_esperados if m not in meses]

        if meses_faltantes:
            for mes in meses_faltantes:
                arq.loc[len(arq)] = [mes] + [''] * (len(arq.columns) - 1)

        arq.to_csv(arquivo, index=False)
    
    else:
        continue


def corrigir_csv_por_prefixo(
    artista: str,
    prefixo: str = "Ug"
) -> None:
    """
    Lê um arquivo CSV, une linhas que não começam com o prefixo à linha anterior que começa com ele,
    e salva o resultado em um novo arquivo.

    Args:
        caminho_entrada (str): Caminho para o arquivo CSV de entrada.
        caminho_saida (str): Caminho onde o arquivo corrigido será salvo.
        prefixo (str): Prefixo que identifica o início de uma nova entrada. Padrão: "Ug"
    """
    with open(f'dados_full/{artista}/comunidade.csv', "r", encoding="utf-8") as f:
        linhas = f.readlines()

        # Preservar as duas primeiras linhas
    linhas_corrigidas = linhas[:2]
    linha_corrente = ""

    for linha in linhas[2:]:
        if linha.startswith(prefixo):
            if linha_corrente:
                linhas_corrigidas.append(linha_corrente)
            linha_corrente = linha.strip()
        else:
            linha_corrente += " " + linha.strip()

    if linha_corrente:
        linhas_corrigidas.append(linha_corrente)

    with open(f'dados_full/{artista}/comunidade.csv', "w", encoding="utf-8") as f:
        for linha in linhas_corrigidas:
            f.write(linha + "\n")


def preencher_colunas_vazias(artista):
    arquivos_csv = glob.glob(path+'/dados_full/'+artista+'/*.csv')

    for arquivo in arquivos_csv:
        arq = pd.read_csv(arquivo)
        
        if 'Duração média da visualização' in arq.columns:
            arq['Duração média da visualização'] = arq['Duração média da visualização'].apply(lambda x: '0:00:00' if pd.isna(x) else '0:00:00' if x == '0.0' else x)
            arq.to_csv(arquivo, index=False)
    
    for arquivo in arquivos_csv:
        arq = pd.read_csv(arquivo)
        arq = arq.fillna(0)
        arq.to_csv(arquivo, index=False)


def update_traffic_source(artista):
    base_path = f'{path}/dados_full/{artista}'
    arquivos_csv = glob.glob(f'{base_path}/*.csv')
    origem_trafego = ['Recursos de navegação', 'Vídeos sugeridos', 'Feed dos Shorts', 
                      'Externa', 'Notificações', 'Pesquisa do YouTube', 'Playlists', 
                      'Publicidade no YouTube']
    
    # Passo 1: Encontrar e processar 'total.csv' para definir 'unique_months'
    total_file_path = f'{base_path}/total.csv'
    unique_months = []  # Inicializa como uma lista vazia para segurança
    try:
        arq_total = pd.read_csv(total_file_path)
        unique_months_all = arq_total['Data'].unique()
        unique_months_sorted = np.sort(unique_months_all)
        unique_months = unique_months_sorted[1:7]  # Pega do 2º ao 7º mês
    except FileNotFoundError:
        print(f"Aviso: Arquivo 'total.csv' não encontrado para {artista}. Não é possível atualizar as origens de tráfego.")
        return  # Sai da função se 'total.csv' não for encontrado
    except Exception as e:
        print(f"Erro ao processar 'total.csv' para {artista}: {e}")
        return

    # Passo 2: Iterar por todos os arquivos CSV e atualizar aqueles com "origem_"
    for arquivo in arquivos_csv:
        if "origem_" in arquivo:
            arq = pd.read_csv(arquivo)
            for origem in origem_trafego:
                if origem not in arq['Origem do tráfego'].values:
                    # Cria novas linhas apenas se 'unique_months' não estiver vazio
                    if len(unique_months) > 0:
                        new_rows = pd.DataFrame({
                            'Data': unique_months,
                            'Origem do tráfego': [origem] * len(unique_months),
                            'Visualizações': [0] * len(unique_months)
                        })
                        arq = pd.concat([arq, new_rows], ignore_index=True)
            
            arq.to_csv(arquivo, index=False)


def run():
    lista_de_artistas = buscar_lista_artistas()

    for artista in lista_de_artistas:
        print(f'Tratando: {artista}')
        ajustar_padrao_colunas(artista)
        completar_data(artista)
        process_post_data(artista)
        corrigir_csv_por_prefixo(str(artista))
        preencher_colunas_vazias(artista)
        update_traffic_source(artista)

if __name__ == '__main__':
    run()

