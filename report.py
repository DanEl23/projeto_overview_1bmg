import os
import sys
import numpy as np
import pandas as pd
from typing import List

# Obtém o caminho do diretório onde o script está sendo executado
path = os.getcwd()

def divisao_segura(numerador, denominador):
    """
    Realiza divisão segura elemento a elemento para escalares, arrays NumPy ou objetos pandas.
    É robusto para cenários de broadcasting (ex: 2D / 1D).
    Retorna 0 onde o denominador for 0.
    """
    if isinstance(denominador, (pd.Series, pd.DataFrame)):
        denominador_seguro = denominador.replace(0, np.nan)
        resultado = (numerador / denominador_seguro)
        return resultado.replace([np.inf, -np.inf], 0).fillna(0)

    numerador = np.asarray(numerador, dtype=float)
    denominador = np.asarray(denominador, dtype=float)

    if numerador.ndim == 2 and denominador.ndim == 1 and numerador.shape[0] == denominador.shape[0]:
        denominador = denominador.reshape(-1, 1)

    resultado = np.zeros_like(numerador, dtype=float)
    mascara_valida = denominador != 0
    
    np.divide(numerador, denominador, out=resultado, where=mascara_valida)
    
    return resultado


def carregar_e_processar_dados(caminho_arquivo: str, tipo_conteudo: str) -> pd.DataFrame:
    """
    Carrega e pré-processa um arquivo CSV de dados do YouTube.
    """
    df = pd.read_csv(caminho_arquivo)
    if df.empty:
        return pd.DataFrame()

    df = df.drop(df.index[0])
    # Linha corrigida com o formato explícito
    df['Data'] = pd.to_datetime(df['Data'], format='%Y-%m').dt.to_period('M')
    
    duracao_timedelta = pd.to_timedelta(df['Duração média da visualização'], errors='coerce')
    minutos = duracao_timedelta.dt.components['minutes'].astype(str).str.zfill(2)
    segundos = duracao_timedelta.dt.components['seconds'].astype(str).str.zfill(2)
    df['Duração média da visualização'] = minutos + ':' + segundos

    df['Taxa de cliques de impressões (%)'] = pd.to_numeric(df['Taxa de cliques de impressões (%)'], errors='coerce').fillna(0) / 100
    df['Porcentagem visualizada média (%)'] = pd.to_numeric(df['Porcentagem visualizada média (%)'], errors='coerce').fillna(0) / 100
    
    df.sort_values(by='Data', inplace=True)
    if not df.empty:
      df.drop(df.index[0], inplace=True, errors='ignore')
    df.set_index('Data', inplace=True)

    if tipo_conteudo in ['videos', 'shorts', 'lives']:
        numerador_engajamento = abs(df['Comentários adicionados'].astype(float)) + abs(df['Marcações \"Gostei\"'].astype(float)) + abs(df['Compartilhamentos'].astype(float))
        denominador_engajamento = df['Vídeos publicados'].astype(float)
        df['Engajamento'] = divisao_segura(numerador_engajamento, denominador_engajamento)
        
    return df.T


def carregar_dados_mensais(artista: str, tipo: str) -> pd.DataFrame:
    """
    Carrega e concatena dados mensais de um determinado tipo (videos, lives, shorts).
    O intervalo de 6 meses é determinado automaticamente a partir do último mês presente no arquivo total.csv.
    """
    # 1. Ler o arquivo total.csv para encontrar a última data
    caminho_total = f'dados_full/{artista}/total.csv'
    if not os.path.exists(caminho_total):
        print(f"Erro: Arquivo 'total.csv' não encontrado para o artista {artista}.")
        return pd.DataFrame()
        
    df_total = pd.read_csv(caminho_total)

    # Adicionando a linha para remover a primeira linha do DataFrame,
    # que provavelmente contém a string "Total".
    # Isso evita o erro de conversão de data.
    if not df_total.empty:
      df_total = df_total.drop(df_total.index[0])

    # Usando errors='coerce' para lidar com valores não-data (que agora devem ter sido removidos)
    # e format='%Y-%m' para evitar o UserWarning anterior e melhorar a performance.
    df_total['Data'] = pd.to_datetime(df_total['Data'], format='%Y-%m', errors='coerce')
    # Removendo quaisquer linhas que não puderam ser convertidas para data (serão NaT)
    df_total.dropna(subset=['Data'], inplace=True)
    
    ultima_data = df_total['Data'].max()
    
    # ... o restante da função é o mesmo ...

    # 2. Calcular os 6 meses anteriores à última data encontrada
    meses_para_analise = []
    data_atual = ultima_data
    for _ in range(6):
        meses_para_analise.insert(0, data_atual.strftime('%Y-%m'))
        if data_atual.month == 1:
            data_atual = data_atual.replace(year=data_atual.year - 1, month=12)
        else:
            data_atual = data_atual.replace(month=data_atual.month - 1)

    # 3. Processar os arquivos mensais com o novo intervalo
    dfs = []
    for i, mes_str_analise in enumerate(meses_para_analise):
        nome_arquivo_mes = str(i + 1).zfill(2)
        caminho_arquivo = f'dados_full/{artista}/{tipo}_{nome_arquivo_mes}.csv'

        if os.path.exists(caminho_arquivo):
            df_mes = pd.read_csv(caminho_arquivo)
            df_mes = df_mes[df_mes['Data'] == mes_str_analise]
            dfs.append(df_mes)

    # O restante da função para concatenar e processar os DataFrames permanece o mesmo
    if not dfs:
        return pd.DataFrame()
        
    df_concatenado = pd.concat(dfs, ignore_index=True)

    df_concatenado['Duração média da visualização'] = pd.to_timedelta(
        df_concatenado['Duração média da visualização'], errors='coerce'
    ).apply(lambda x: x.total_seconds()) / 60

    df_concatenado['Data'] = pd.to_datetime(df_concatenado['Data'], format='%Y-%m').dt.to_period('M')
    df_concatenado.set_index('Data', inplace=True)
    
    return df_concatenado.T


def calcular_metricas_por_publicacao(df_metricas, df_publicacoes, metrica_prefixo, metrica_col_idx, pub_col_idx):
    """
    Calcula uma métrica por número de publicações de forma segura.
    """
    if df_metricas.empty or df_publicacoes.empty:
        return pd.DataFrame()
        
    metricas = df_metricas.iloc[[metrica_col_idx]]
    pubs = df_publicacoes.iloc[[pub_col_idx]]
    
    metricas_valores = metricas.values.flatten().astype(float)
    pubs_valores = pubs.values.flatten().astype(float)

    resultados = [divisao_segura(m, p) for m, p in zip(metricas_valores, pubs_valores)]

    df_resultado = pd.DataFrame([resultados], columns=metricas.columns)
    nome_da_metrica = df_publicacoes.index[pub_col_idx].replace('Número de', '').replace('Números de','').strip()
    df_resultado.index = [f"{metrica_prefixo} por {nome_da_metrica} Novo"]
    return df_resultado

def processar_fontes_de_trafego(artista: str) -> pd.DataFrame:
    """
    Processa e combina os dados de origem de tráfego para VODs e Lives.
    É robusto contra arquivos de origem vazios ou ausentes.
    """
    origem_vods_path = f'dados_full/{artista}/origem_vods.csv'
    origem_lives_path = f'dados_full/{artista}/origem_lives.csv'

    colunas_desejadas = [
        'Recursos de navegação', 'Vídeos sugeridos', 'Feed dos Shorts', 'Externa',
        'Notificações', 'Pesquisa do YouTube', 'Playlists', 'Publicidade no YouTube'
    ]

    def pivot_origem(path):
        if not os.path.exists(path): return pd.DataFrame()
        try:
            df = pd.read_csv(path)
            if df.empty or not all(col in df.columns for col in ['Data', 'Origem do tráfego', 'Visualizações']):
                return pd.DataFrame()
            pivot = df.pivot_table(index='Data', columns='Origem do tráfego', values='Visualizações', aggfunc='first').fillna(0)
        except Exception:
            return pd.DataFrame()

        for col in colunas_desejadas:
            if col not in pivot.columns: pivot[col] = 0
        
        outros_cols = [col for col in pivot.columns if col not in colunas_desejadas]
        outros = pivot[outros_cols].sum(axis=1)
        pivot = pivot[colunas_desejadas].copy()
        pivot['Outros'] = outros
        return pivot

    origem_vods = pivot_origem(origem_vods_path)
    origem_lives = pivot_origem(origem_lives_path)
    
    origem_total = origem_vods.add(origem_lives, fill_value=0)
    
    if origem_total.empty:
        return pd.DataFrame()

    soma_total_por_linha = origem_total.sum(axis=1)
    divisor = pd.DataFrame([soma_total_por_linha.values] * len(origem_total.columns), index=origem_total.columns, columns=origem_total.index).T
    perc_origem = divisao_segura(origem_total.mul(100), divisor)

    perc_origem.reset_index(inplace=True)
    perc_origem['Data'] = pd.to_datetime(perc_origem['Data']).dt.to_period('M')
    perc_origem.rename(columns={'Publicidade no YouTube': 'Tráfego Pago', 'index': 'Data'}, inplace=True)
    
    return perc_origem.set_index('Data').T

def subtract_and_reindex(df1, df1_row_idx, df2, df2_row_idx, new_index_name):
    """
    Subtrai duas linhas de DataFrames ignorando seus índices e retorna um novo DataFrame
    de uma linha com o índice especificado.
    """
    cols = df1.columns
    df2_aligned = df2.reindex(columns=cols).fillna(0)

    row1_values = df1.iloc[[df1_row_idx]].values.astype(float)
    row2_values = df2_aligned.iloc[[df2_row_idx]].values.astype(float)

    result_values = row1_values - row2_values
    
    return pd.DataFrame(result_values, columns=cols, index=[new_index_name])

def gerar_relatorio_para_artista(artista: str):
    """
    Função principal que orquestra todo o processo de geração de relatório para um artista.
    """
    print(f"Iniciando o processamento para o artista: {artista}...")

    # --- 1. Carregar e Processar Dados Base ---
    base_path = f'dados_full/{artista}/'
    df = carregar_e_processar_dados(f'{base_path}total.csv', 'total')
    df_videos = carregar_e_processar_dados(f'{base_path}videos.csv', 'videos')
    df_shorts = carregar_e_processar_dados(f'{base_path}shorts.csv', 'shorts')
    df_lives = carregar_e_processar_dados(f'{base_path}lives.csv', 'lives')
    
    # As chamadas são atualizadas para não passarem o argumento 'meses_range'
    videos_novos = carregar_dados_mensais(artista, 'videos')
    lives_novos = carregar_dados_mensais(artista, 'lives')
    shorts_novos = carregar_dados_mensais(artista, 'shorts')
    
    cols = df.columns
    videos_novos = videos_novos.reindex(columns=cols).fillna(0)
    lives_novos = lives_novos.reindex(columns=cols).fillna(0)
    shorts_novos = shorts_novos.reindex(columns=cols).fillna(0)
    
    total_novo = videos_novos.add(lives_novos, fill_value=0).add(shorts_novos, fill_value=0)

    # --- 2. Cálculos e Definições de Métricas ---
    receita_shorts_calc = df_shorts.iloc[[13]]
    receita_sem_shorts = subtract_and_reindex(df, 16, receita_shorts_calc, 0, "Receita Sem Shorts")
    
    impress_shorts_calc = df_shorts.iloc[[6]]
    impress_sem_shorts = subtract_and_reindex(df, 9, impress_shorts_calc, 0, "Impressões sem shorts")

    views_shorts_calc = df_shorts.iloc[[10]]
    views_sem_shorts = subtract_and_reindex(df, 13, views_shorts_calc, 0, "Visualizações sem Shorts")
    
    receita_vod = df_videos.iloc[[13]].rename(index={"Receita estimada (USD)":"Receita VOD's"})
    receita_lives = df_lives.iloc[[13]].rename(index={"Receita estimada (USD)":"Receita Lives "})
    receita_shorts = df_shorts.iloc[[13]].rename(index={"Receita estimada (USD)":"Receita Shorts"})
    
    impress_vods = df_videos.iloc[[6]].rename(index={"Impressões":"Impressões VOD's"})
    impress_lives = df_lives.iloc[[6]].rename(index={"Impressões":"Impressões Lives"})
    impressoes_shorts = df_shorts.iloc[[6]].rename(index={"Impressões": "Impressões Shorts"})
    
    views_vod = df_videos.iloc[[10]].rename(index={"Visualizações":"Visualizações VOD's"})
    views_lives = df_lives.iloc[[10]].rename(index={"Visualizações":"Visualizações Lives"})
    views_shorts = df_shorts.iloc[[10]].rename(index={"Visualizações": "Visualizações Shorts"})

    # Definição das novas métricas de WatchTime
    watchtime_vod = df_videos.iloc[[11]].rename(index={"Tempo de exibição (horas)": "WatchTime VOD's"})
    watchtime_lives = df_lives.iloc[[11]].rename(index={"Tempo de exibição (horas)": "WatchTime Lives"})
    watchtime_shorts = df_shorts.iloc[[11]].rename(index={"Tempo de exibição (horas)": "WatchTime Shorts"})

    inscritos_ganhos = df.iloc[[20]].rename(index={"Inscritos":"Número de Inscritos Ganhos"})
    inscritos_perdidos = df.iloc[[21]].rename(index={"Inscritos":"Número de Inscritos Perdidos"})
    saldo_inscritos = df.iloc[[11]].rename(index={"Inscritos":"Saldo de Inscritos"})

    rpm_total_vod = df_videos.iloc[[4]].rename(index={"RPM (USD)": "RPM VOD's"})
    rpm_total_lives = df_lives.iloc[[4]].rename(index={"RPM (USD)": "RPM Lives"})
    rpm_total_shorts = df_shorts.iloc[[4]].rename(index={"RPM (USD)": "RPM Shorts"})
    rpm_total = df.iloc[[7]].rename(index={"RPM (USD)": "RPM Total"})


    vods_pubs = df_videos.iloc[[7]]; vods_pubs.index = ["Número de VOD's"]
    lives_pubs = df_lives.iloc[[7]]; lives_pubs.index = ["Números de Lives "]
    shorts_pubs = df_shorts.iloc[[7]]; shorts_pubs.index = ["Números de Shorts"]
    watchtime = df.iloc[[14]]; watchtime.index = ["Watch Time Total"]
    cpm = df.iloc[[17]]; cpm.index = ["CPM (USD)"]

    impress_vods_calc = df_videos.iloc[[6]].astype(float)
    impress_lives_calc = df_lives.iloc[[6]].astype(float)
    total_impress_sem_shorts = impress_vods_calc.values + impress_lives_calc.values
    
    peso_videos = divisao_segura(impress_vods_calc.values, total_impress_sem_shorts)
    peso_lives = divisao_segura(impress_lives_calc.values, total_impress_sem_shorts)
    
    rpm_vod_vals = df_videos.iloc[[4]].values.astype(float)
    rpm_lives_vals = df_lives.iloc[[4]].values.astype(float)
    
    rpm_sem_short_vals = (rpm_vod_vals * peso_videos) + (rpm_lives_vals * peso_lives)
    rpm_sem_shorts = pd.DataFrame(rpm_sem_short_vals, columns=df.columns, index=['RPM sem Short'])

    views_velho = subtract_and_reindex(df, 13, total_novo, 9, "temp_index")
    receita_velho = subtract_and_reindex(df, 16, total_novo, 12, "temp_index")
    rpm_velho_vals = divisao_segura(receita_velho.values.astype(float) * 1000, views_velho.values.astype(float))
    rpm_velho = pd.DataFrame(rpm_velho_vals, columns=df.columns, index=['RPM Velho'])

    views_novo_vod = videos_novos.iloc[[9]].rename(index={"Visualizações":"Visualizações VOD's Novo"})
    views_novo_live = lives_novos.iloc[[9]].rename(index={"Visualizações":"Visualizações Lives Novo"})
    views_novo_shorts = shorts_novos.iloc[[9]].rename(index={"Visualizações":"Visualizações Shorts Novo"})

    views_velho_vod = subtract_and_reindex(df_videos, 10, videos_novos, 9, "Visualizações VOD's Velho")
    views_velho_live = subtract_and_reindex(df_lives, 10, lives_novos, 9, "Visualizações Lives Velho")
    views_velho_shorts = subtract_and_reindex(df_shorts, 10, shorts_novos, 9, "Visualizações Shorts Velho")

    views_total = df.iloc[[13]].rename(index={"Visualizações":"Visualizações Total"})

    rpm_novo_vals = divisao_segura(total_novo.iloc[[12]].values.astype(float) * 1000, total_novo.iloc[[9]].values.astype(float))
    rpm_novo = pd.DataFrame(rpm_novo_vals, columns=total_novo.columns, index=['RPM Novo'])

    taxa_preenchimento_vals = divisao_segura(df.iloc[[18]].values.astype(float), cpm.values.astype(float))
    taxa_preenchimento = pd.DataFrame(taxa_preenchimento_vals, columns=df.columns, index=['Taxa de Preenchimento'])


    minutagem_videos_vals = divisao_segura(videos_novos.iloc[[11]].values.astype(float) * 100, videos_novos.iloc[[8]].values.astype(float)) 
    minutagem_videos = pd.DataFrame(minutagem_videos_vals, columns=videos_novos.columns, index=["Tamanhos de VOD's Novo"])
    
    minutagem_lives_vals = divisao_segura(lives_novos.iloc[[11]].values.astype(float) * 100, lives_novos.iloc[[8]].values.astype(float))
    minutagem_lives = pd.DataFrame(minutagem_lives_vals, columns=lives_novos.columns, index=["Tamanho de Lives Novo"])

    minutagem_shorts_vals = divisao_segura(shorts_novos.iloc[[11]].values.astype(float) * 100, shorts_novos.iloc[[8]].values.astype(float)) 
    minutagem_shorts = pd.DataFrame(minutagem_shorts_vals, columns=shorts_novos.columns, index=["Tamanhos de Shorts Novo"])


    def calcular_engajamento_novo(df_novo, df_pubs, content_type):
        if df_novo.empty or df_pubs.empty: return pd.DataFrame()
        engaj_novo_total = (abs(df_novo.T['Comentários adicionados'].astype(float)) + abs(df_novo.T['Marcações \"Gostei\"'].astype(float)) + abs(df_novo.T['Compartilhamentos'].astype(float)))
        engaj_por_pub = divisao_segura(engaj_novo_total.values, df_pubs.iloc[0].T.values.astype(float))
        df_out = pd.DataFrame(engaj_por_pub.reshape(1, -1), columns=df_pubs.columns)
        df_out.index = [f"Engajamento {content_type}"]
        return df_out
    engajamento_vod_novo = calcular_engajamento_novo(videos_novos, vods_pubs, 'VOD')
    engajamento_lives_novo = calcular_engajamento_novo(lives_novos, lives_pubs, 'Lives')
    engajamento_shorts_novo = calcular_engajamento_novo(shorts_novos, shorts_pubs, 'Shorts')
    
    # --- 3. Montagem da Tabela Final ---
    df_geral_1 = pd.concat([
        receita_sem_shorts, receita_vod, receita_lives, receita_shorts,
        impress_sem_shorts, impress_vods, impress_lives, 
        views_sem_shorts, views_shorts, views_vod, views_lives,
        vods_pubs, lives_pubs, shorts_pubs, watchtime
    ])
    
    with open(f"dados_full/{artista}/sub.txt", "r") as f: insc_ant = int(''.join((f.readline()).split('.')))
    inscritos_totais = pd.DataFrame((df.T['Inscritos'].astype(float).cumsum() + (insc_ant - df.T['Inscritos'].astype(float).sum()))).T.rename(index={"Inscritos":"Número de Inscritos Total"})
    inscritos_sem_shorts = subtract_and_reindex(df, 11, df_shorts, 8, "Número de Inscritos Sem Shorts")
    df_geral_2 = pd.concat([
        df_videos.iloc[[8]].rename(index={"Inscritos":"Número de Inscritos VOD's "}), df_lives.iloc[[8]].rename(index={"Inscritos":"Número de Inscritos LIves"}),
        df_shorts.iloc[[8]].rename(index={"Inscritos":"Número de Inscritos Shorts"}), inscritos_totais, inscritos_sem_shorts
    ])
    
    lista_dfs_finais = [
        df_geral_1,
        calcular_metricas_por_publicacao(videos_novos, vods_pubs, 'Impressões', 6, 0),
        calcular_metricas_por_publicacao(lives_novos, lives_pubs, 'Impressões', 6, 0),
        calcular_metricas_por_publicacao(shorts_novos, shorts_pubs, 'Impressões', 6, 0),
        rpm_velho, rpm_novo, rpm_sem_shorts, videos_novos.iloc[[4]].rename(index={"RPM (USD)":"RPM VOD's Novo"}),
        lives_novos.iloc[[4]].rename(index={"RPM (USD)":"RPM Lives Novo"}), shorts_novos.iloc[[4]].rename(index={"RPM (USD)":"RPM Shorts Novo"}),
        videos_novos.iloc[[5]].rename(index={"Taxa de cliques de impressões (%)":"CTR VOD's Novo"}),
        lives_novos.iloc[[5]].rename(index={"Taxa de cliques de impressões (%)":"CTR Lives Novo"}),
        videos_novos.iloc[[8]].rename(index={"Porcentagem visualizada média (%)":"Porcentagem Média Assistirda VOD's Novo"}),
        lives_novos.iloc[[8]].rename(index={"Porcentagem visualizada média (%)":"Porcentagem Média Assistirda Live Novo"}),
        minutagem_videos, minutagem_lives,
        videos_novos.iloc[[11]].rename(index={"Duração média da visualização":"Tempo Médio Assistido VOD's Novo"}),
        lives_novos.iloc[[11]].rename(index={"Duração média da visualização":"Tempo Médio Assistido Lives Novo"}),
        processar_fontes_de_trafego(artista),
        videos_novos.iloc[[7]].rename(index={"Inscritos":"Número de Inscritos VOD Novo"}),
        lives_novos.iloc[[7]].rename(index={"Inscritos":"Número de Inscritos Live Novo"}),
        shorts_novos.iloc[[7]].rename(index={"Inscritos":"Número de Inscritos Shorts Novo"}),
        subtract_and_reindex(df_videos, 8, videos_novos, 7, "Número de Inscritos vods Velho"),
        subtract_and_reindex(df_lives, 8, lives_novos, 7, "Número de Inscritos Lives Velho"),
        subtract_and_reindex(df_shorts, 8, shorts_novos, 7, "Número de Inscritos Shorts Velho"),
        df_geral_2,
        calcular_metricas_por_publicacao(videos_novos, vods_pubs, "Receita", 12, 0),
        calcular_metricas_por_publicacao(lives_novos, lives_pubs, "Receita", 12, 0),
        calcular_metricas_por_publicacao(shorts_novos, shorts_pubs, "Receita", 12, 0),
        videos_novos.iloc[[12]].rename(index={"Receita estimada (USD)":"Receita VOD Novo"}),
        lives_novos.iloc[[12]].rename(index={"Receita estimada (USD)":"Receita Live Novo"}),
        cpm,
        subtract_and_reindex(df_videos, 13, videos_novos, 12, "Receita VOD's Velho"),
        subtract_and_reindex(df_shorts, 13, shorts_novos, 12, "Receita Shorts Velho"),
        subtract_and_reindex(df_lives, 13, lives_novos, 12, "Receita Lives Velho"),
        shorts_novos.iloc[[12]].rename(index={"Receita estimada (USD)":"Receita Shorts Novo"}),
        taxa_preenchimento, engajamento_vod_novo, engajamento_lives_novo,
        impressoes_shorts,
        # Adicionando as novas métricas de Watch Time ao final da lista
        watchtime_vod,
        watchtime_lives,
        watchtime_shorts,
        rpm_total_vod,
        rpm_total_lives,
        rpm_total_shorts,
        rpm_total,
        inscritos_ganhos,
        inscritos_perdidos,
        saldo_inscritos,
        engajamento_shorts_novo,
        minutagem_shorts,
        shorts_novos.iloc[[8]].rename(index={"Porcentagem visualizada média (%)":"Porcentagem Média Assistirda Shorts Novo"}),
        shorts_novos.iloc[[11]].rename(index={"Duração média da visualização":"Tempo Médio Assistido Shorts Novo"}),
        views_novo_vod,
        views_novo_live,
        views_novo_shorts,
        views_velho_vod,
        views_velho_live,
        views_velho_shorts,
        views_total

    ]

    lista_dfs_validos = [df for df in lista_dfs_finais if df is not None and not df.empty]
    tabela = pd.concat(lista_dfs_validos)
    
    # --- Pós-processamento e Exportação ---
    colunas_de_meses = tabela.columns.tolist()
    
    tabela_calc = tabela.copy()
    for col in colunas_de_meses:
        tabela_calc[col] = pd.to_numeric(tabela_calc[col], errors='coerce')
    tabela_calc.fillna(0, inplace=True)

    media_series = tabela_calc[colunas_de_meses].mean(axis=1)
    tabela.insert(0, 'Média', media_series)
    tabela.index.name = "Data"

    tabela_numerica_sem_media = tabela_calc[colunas_de_meses]
    
    resultado_array = divisao_segura(tabela_numerica_sem_media.sub(media_series, axis=0), media_series.values)
    tabela_desvio_media = pd.DataFrame(resultado_array, 
                                       index=tabela_numerica_sem_media.index, 
                                       columns=tabela_numerica_sem_media.columns)
    tabela_desvio_media.index.name = "Data"
    
    df_shifted = tabela_numerica_sem_media.shift(1, axis=1)
    numerador_pct = tabela_numerica_sem_media.sub(df_shifted)
    tabela_desvio_anterior = divisao_segura(numerador_pct, df_shifted)
    tabela_desvio_anterior.index.name = "Data"
    
    # 1. Identifique as linhas que serão formatadas como strings de tempo mais tarde.
    linhas_de_tempo = [
        "Tamanhos de VOD's Novo", "Tempo Médio Assistido VOD's Novo",
        "Tamanho de Lives Novo", "Tempo Médio Assistido Lives Novo",
        "Tamanhos de Shorts Novo", "Tempo Médio Assistido Shorts Novo",
    ]

    # 2. Identifique TODAS as outras linhas que devem ser puramente numéricas.
    # Usamos 'index.difference' para pegar todos os índices EXCETO os de tempo.
    linhas_numericas = tabela.index.difference(linhas_de_tempo)

    # 3. Converta APENAS as linhas numéricas para float.
    # Isso resolve o problema original (inteiros no Excel) sem criar o novo erro.
    # Usamos .loc para garantir que estamos modificando o DataFrame original.
    for col in colunas_de_meses:
        tabela.loc[linhas_numericas, col] = pd.to_numeric(tabela.loc[linhas_numericas, col], errors='coerce')
    
    # Garante que, após a conversão, o tipo seja float.
    tabela.loc[linhas_numericas] = tabela.loc[linhas_numericas].astype(float)


    # 4. AGORA, com os tipos numéricos já corrigidos, aplique a formatação de tempo.
    # Este código agora opera de forma segura, sem ser sobrescrito depois.
    for metrica in linhas_de_tempo:
        if metrica in tabela.index:
            # Garante que estamos trabalhando com números antes de formatar
            valores_numericos = pd.to_numeric(tabela.loc[metrica, colunas_de_meses], errors='coerce').fillna(0)
            
            # Aplica a formatação para MM:SS ou HH:MM:SS
            tabela.loc[metrica, colunas_de_meses] = valores_numericos.apply(
                lambda minutos: 
                    # SE o valor de minutos for 60 ou mais, usa o formato HH:MM:SS
                    f"{int(minutos // 60):02d}:{int(minutos % 60):02d}:00" 
                    if minutos >= 60 
                    # SENÃO, usa o formato MM:SS para minutos decimais
                    else f"{int(minutos):02d}:{int((minutos % 1) * 60):02d}"
            )
    
    output_filename = f'exports_tabelas/tabela_4.1_{artista}.xlsx'
    os.makedirs('exports_tabelas', exist_ok=True)
    
    with pd.ExcelWriter(output_filename, engine='openpyxl') as writer:
        tabela.to_excel(writer, sheet_name='Resultado', index=True)
        tabela_desvio_media.to_excel(writer, sheet_name='Desvio', index=True)
        tabela_desvio_anterior.to_excel(writer, sheet_name='Mês Anterior', index=True)

    print(f"Relatório para {artista} salvo com sucesso")


def run(artista):
    print(f"Gerando report para: {artista}")
    
    try:
        gerar_relatorio_para_artista(artista)
    except FileNotFoundError as e:
        print(f"\nERRO: Não foi possível processar '{artista_selecionado}'. Arquivo não encontrado: {e.filename}\n")
    except Exception as e:
        print(f"\nERRO: Ocorreu um problema inesperado ao processar '{artista_selecionado}': {e}\n")
        import traceback
        traceback.print_exc()
    print(f"Report para {artista} finalizado.")


if __name__ == "__main__":
    # Pega o nome do artista do argumento passado pelo main.py
    if len(sys.argv) < 2:
        print("Erro: Nenhum artista fornecido. Este script deve ser chamado pelo main.py")
        sys.exit(1) # Sai com erro
    
    artista_argumento = sys.argv[1]
    
    # Executa a função run APENAS para esse artista
    run(artista_argumento)
