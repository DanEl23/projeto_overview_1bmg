import pandas as pd
import numpy as np
import os
from typing import List
from dateutil.relativedelta import relativedelta

# Obtém o caminho do diretório onde o script está sendo executado
path = os.getcwd()

def divisao_segura(numerador, denominador):
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

def carregar_e_processar_dados(caminho_arquivo: str) -> pd.DataFrame:
    if not os.path.exists(caminho_arquivo): return pd.DataFrame().T
    df = pd.read_csv(caminho_arquivo)
    if df.empty: return pd.DataFrame().T
    if 'Total' in df.iloc[0].values:
        df = df.drop(df.index[0])
    df['Data'] = pd.to_datetime(df['Data'], format='%Y-%m').dt.to_period('M')
    for col in ['Taxa de cliques de impressões (%)', 'Porcentagem visualizada média (%)']:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0) / 100
    df.sort_values(by='Data', inplace=True)
    df.set_index('Data', inplace=True)
    return df.T

def carregar_dados_mensais(artista: str, tipo: str) -> pd.DataFrame:
    caminho_total = f'dados_full/{artista}/total.csv'
    if not os.path.exists(caminho_total): return pd.DataFrame()
    df_total = pd.read_csv(caminho_total)
    if not df_total.empty and 'Total' in df_total.iloc[0].values:
        df_total = df_total.drop(df_total.index[0])
    df_total['Data'] = pd.to_datetime(df_total['Data'], format='%Y-%m', errors='coerce')
    df_total.dropna(subset=['Data'], inplace=True)
    if df_total.empty: return pd.DataFrame()
    ultima_data = df_total['Data'].max()
    meses_para_analise = []
    for i in range(5, -1, -1):
        mes_alvo = ultima_data - relativedelta(months=i)
        meses_para_analise.append(mes_alvo.strftime('%Y-%m'))
    coluna_data_publicacao = 'Horário de publicação do vídeo'
    dfs = []
    for i in range(1, 13):
        nome_arquivo_mes = str(i).zfill(2)
        caminho_arquivo = f'dados_full/{artista}/{tipo}_{nome_arquivo_mes}.csv'
        if os.path.exists(caminho_arquivo):
            df_mes = pd.read_csv(caminho_arquivo)
            if coluna_data_publicacao in df_mes.columns:
                datas_convertidas = pd.to_datetime(df_mes[coluna_data_publicacao], format='%b %d, %Y', errors='coerce')
                df_mes['mes_publicacao'] = datas_convertidas.dt.strftime('%Y-%m')
                df_filtrado = df_mes[df_mes['mes_publicacao'].isin(meses_para_analise)].copy()
                dfs.append(df_filtrado)
    if not dfs: return pd.DataFrame()
    df_concatenado = pd.concat(dfs, ignore_index=True)

    if 'Duração média da visualização' in df_concatenado.columns:
        df_concatenado['Duração média da visualização'] = pd.to_timedelta(df_concatenado['Duração média da visualização'], errors='coerce').dt.total_seconds()

    numeric_cols = ['Receita estimada (USD)', 'Porcentagem visualizada média (%)', 'Impressões', 'Taxa de cliques de impressões (%)', 'RPM (USD)', 'Marcações "Gostei"', 'Compartilhamentos', 'Comentários adicionados', 'Espectadores únicos', 'Visualizações', 'Tempo de exibição (horas)', 'Inscritos', 'Duração média da visualização', 'Duração']
    for col in numeric_cols:
        if col in df_concatenado.columns:
            df_concatenado[col] = pd.to_numeric(df_concatenado[col], errors='coerce')
            
    agg_dict = {'Receita estimada (USD)': 'sum', 'Porcentagem visualizada média (%)': 'mean', 'Impressões': 'sum', 'Taxa de cliques de impressões (%)': 'mean', 'RPM (USD)': 'mean', 'Marcações "Gostei"': 'sum', 'Compartilhamentos': 'sum', 'Comentários adicionados': 'sum', 'Espectadores únicos': 'sum', 'Visualizações': 'sum', 'Tempo de exibição (horas)': 'sum', 'Inscritos': 'sum', 'Duração média da visualização': 'mean', 'Duração': 'mean'}
    agg_dict_existente = {k: v for k, v in agg_dict.items() if k in df_concatenado.columns}
    if not agg_dict_existente or df_concatenado.empty: return pd.DataFrame()
    df_agrupado = df_concatenado.groupby('mes_publicacao').agg(agg_dict_existente)
    df_transposto = df_agrupado.T
    if not df_transposto.empty:
        df_transposto.columns = pd.to_datetime(df_transposto.columns, format='%Y-%m').to_period('M')
    return df_transposto

def subtract_by_label(df1, label1, df2, label2, new_label):
    cols = df1.columns
    df2_aligned = df2.reindex(columns=cols).fillna(0)
    if label1 not in df1.index or label2 not in df2.index: return pd.DataFrame(index=[new_label], columns=cols).fillna(0)
    row1_values = df1.loc[[label1]].values.astype(float)
    row2_values = df2_aligned.loc[[label2]].values.astype(float)
    return pd.DataFrame(row1_values - row2_values, columns=cols, index=[new_label])

def metric_per_publication_novo(df_metrics, df_publications, metric_label, pub_label, new_metric_prefix):
    clean_pub_label = pub_label.replace('Número de', '').replace('Números de','').strip()
    new_index_name = f"{new_metric_prefix} por {clean_pub_label} Novo"
    cols = df_metrics.columns
    if df_metrics.empty or df_publications.empty or metric_label not in df_metrics.index or pub_label not in df_publications.index:
        return pd.DataFrame(index=[new_index_name], columns=cols).fillna(0)
    metric_values = df_metrics.loc[[metric_label]].values.flatten().astype(float)
    pub_values = df_publications.loc[[pub_label]].values.flatten().astype(float)
    result = divisao_segura(metric_values, pub_values)
    return pd.DataFrame([result], columns=cols, index=[new_index_name])

def processar_fontes_de_trafego(artista: str, meses_desejados: pd.PeriodIndex) -> pd.DataFrame:
    origem_vods_path = f'dados_full/{artista}/origem_vods.csv'
    origem_lives_path = f'dados_full/{artista}/origem_lives.csv'
    colunas_desejadas = ['Recursos de navegação', 'Vídeos sugeridos', 'Feed dos Shorts', 'Externa', 'Notificações', 'Pesquisa do YouTube', 'Playlists', 'Publicidade no YouTube']
    def pivot_origem(path):
        if not os.path.exists(path): return pd.DataFrame()
        df = pd.read_csv(path)
        if df.empty or not all(col in df.columns for col in ['Data', 'Origem do tráfego', 'Visualizações']): return pd.DataFrame()
        pivot = df.pivot_table(index='Data', columns='Origem do tráfego', values='Visualizações', aggfunc='first').fillna(0)
        for col in colunas_desejadas:
            if col not in pivot.columns: pivot[col] = 0
        outros_cols = [col for col in pivot.columns if col not in colunas_desejadas]
        pivot['Outros'] = pivot[outros_cols].sum(axis=1)
        return pivot[colunas_desejadas + ['Outros']]
    origem_vods = pivot_origem(origem_vods_path)
    origem_lives = pivot_origem(origem_lives_path)
    origem_total = origem_vods.add(origem_lives, fill_value=0)
    if origem_total.empty: return pd.DataFrame()
    origem_total.index = pd.to_datetime(origem_total.index).to_period('M')
    origem_total = origem_total[origem_total.index.isin(meses_desejados)]
    if origem_total.empty: return pd.DataFrame()
    soma_total_por_linha = origem_total.sum(axis=1)
    resultado_divisao_np = divisao_segura(origem_total.mul(100), soma_total_por_linha.values)
    perc_origem = pd.DataFrame(resultado_divisao_np, index=origem_total.index, columns=origem_total.columns)
    perc_origem.reset_index(inplace=True)
    perc_origem.rename(columns={'Publicidade no YouTube': 'Tráfego Pago', 'Data': 'Data'}, inplace=True)
    return perc_origem.set_index('Data').T


def gerar_relatorio_para_artista(artista: str):
    print(f"Iniciando o processamento para o artista: {artista}...")

    base_path = f'dados_full/{artista}/'

    caminho_total_ref = f'{base_path}total.csv'
    if not os.path.exists(caminho_total_ref):
        print(f"Arquivo 'total.csv' não encontrado para determinar o período. Saindo.")
        return

    df_ref_data = pd.read_csv(caminho_total_ref)
    if not df_ref_data.empty and 'Total' in df_ref_data.iloc[0].values:
        df_ref_data = df_ref_data.drop(df_ref_data.index[0])
    df_ref_data['Data'] = pd.to_datetime(df_ref_data['Data'], format='%Y-%m', errors='coerce')
    df_ref_data.dropna(subset=['Data'], inplace=True)
    ultima_data = df_ref_data['Data'].max()
    meses_para_analise_dt = []
    for i in range(5, -1, -1):
        mes_alvo = ultima_data - relativedelta(months=i)
        meses_para_analise_dt.append(mes_alvo)
    colunas_de_meses_desejadas = pd.to_datetime(meses_para_analise_dt).to_period('M')

    df = carregar_e_processar_dados(f'{base_path}total.csv')
    df_videos = carregar_e_processar_dados(f'{base_path}videos.csv')
    df_shorts = carregar_e_processar_dados(f'{base_path}shorts.csv')
    df_lives = carregar_e_processar_dados(f'{base_path}lives.csv')
    
    df = df[df.columns.intersection(colunas_de_meses_desejadas)]
    df_videos = df_videos[df_videos.columns.intersection(colunas_de_meses_desejadas)]
    df_shorts = df_shorts[df_shorts.columns.intersection(colunas_de_meses_desejadas)]
    df_lives = df_lives[df_lives.columns.intersection(colunas_de_meses_desejadas)]
    
    videos_novos = carregar_dados_mensais(artista, 'videos')
    lives_novos = carregar_dados_mensais(artista, 'lives')
    shorts_novos = carregar_dados_mensais(artista, 'shorts')
    
    cols = df.columns
    all_new_metrics = pd.concat([videos_novos, lives_novos, shorts_novos]).index.unique()
    videos_novos = videos_novos.reindex(index=all_new_metrics, columns=cols).fillna(0)
    lives_novos = lives_novos.reindex(index=all_new_metrics, columns=cols).fillna(0)
    shorts_novos = shorts_novos.reindex(index=all_new_metrics, columns=cols).fillna(0)
    total_novo = videos_novos.add(lives_novos, fill_value=0).add(shorts_novos, fill_value=0)

    # --- 2. Recriação Fiel dos Cálculos Originais ---
    receita_sem_shorts = subtract_by_label(df, 'Receita estimada (USD)', df_shorts, 'Receita estimada (USD)', "Receita Sem Shorts")
    impress_sem_shorts = subtract_by_label(df, 'Impressões', df_shorts, 'Impressões', "Impressões sem shorts")
    views_sem_shorts = subtract_by_label(df, 'Visualizações', df_shorts, 'Visualizações', "Visualizações sem Shorts")
    receita_vod = df_videos.loc[['Receita estimada (USD)']].rename(index={"Receita estimada (USD)":"Receita VOD's"})
    receita_lives = df_lives.loc[['Receita estimada (USD)']].rename(index={"Receita estimada (USD)":"Receita Lives "})
    receita_shorts = df_shorts.loc[['Receita estimada (USD)']].rename(index={"Receita estimada (USD)":"Receita Shorts"})
    impress_vods = df_videos.loc[['Impressões']].rename(index={"Impressões":"Impressões VOD's"})
    impress_lives = df_lives.loc[['Impressões']].rename(index={"Impressões":"Impressões Lives"})
    impressoes_shorts = df_shorts.loc[['Impressões']].rename(index={"Impressões": "Impressões Shorts"})
    views_vod = df_videos.loc[['Visualizações']].rename(index={"Visualizações":"Visualizações VOD's"})
    views_lives = df_lives.loc[['Visualizações']].rename(index={"Visualizações":"Visualizações Lives"})
    views_shorts = df_shorts.loc[['Visualizações']].rename(index={"Visualizações": "Visualizações Shorts"})
    watchtime_vod = df_videos.loc[['Tempo de exibição (horas)']].rename(index={"Tempo de exibição (horas)": "WatchTime VOD's"})
    watchtime_lives = df_lives.loc[['Tempo de exibição (horas)']].rename(index={"Tempo de exibição (horas)": "WatchTime Lives"})
    watchtime_shorts = df_shorts.loc[['Tempo de exibição (horas)']].rename(index={"Tempo de exibição (horas)": "WatchTime Shorts"})
    inscritos_ganhos = df.loc[['Inscrições obtidas']]
    inscritos_perdidos = df.loc[['Inscrições perdidas']]
    saldo_inscritos = df.loc[['Inscritos']].rename(index={"Inscritos":"Saldo de Inscritos"})
    rpm_total_vod = df_videos.loc[['RPM (USD)']].rename(index={"RPM (USD)": "RPM VOD's"})
    rpm_total_lives = df_lives.loc[['RPM (USD)']].rename(index={"RPM (USD)": "RPM Lives"})
    rpm_total_shorts = df_shorts.loc[['RPM (USD)']].rename(index={"RPM (USD)": "RPM Shorts"})
    rpm_total = df.loc[['RPM (USD)']].rename(index={"RPM (USD)": "RPM Total"})
    vods_pubs = df_videos.loc[['Vídeos publicados']].rename(index={"Vídeos publicados":"Número de VOD's"})
    lives_pubs = df_lives.loc[['Vídeos publicados']].rename(index={"Vídeos publicados":"Números de Lives "})
    shorts_pubs = df_shorts.loc[['Vídeos publicados']].rename(index={"Vídeos publicados":"Números de Shorts"})
    watchtime = df.loc[['Tempo de exibição (horas)']].rename(index={'Tempo de exibição (horas)':"Watch Time Total"})
    cpm = df.loc[['CPM (USD)']].rename(index={"CPM (USD)":"CPM (USD)"}) if 'CPM (USD)' in df.index else pd.DataFrame()
    rpm_sem_shorts = pd.DataFrame(0.0, index=['RPM sem Short'], columns=df.columns)
    total_impress_sem_shorts_vals = df_videos.loc['Impressões'].astype(float) + df_lives.loc['Impressões'].astype(float)
    peso_videos = divisao_segura(df_videos.loc['Impressões'].astype(float), total_impress_sem_shorts_vals)
    peso_lives = divisao_segura(df_lives.loc['Impressões'].astype(float), total_impress_sem_shorts_vals)
    rpm_sem_short_vals = (df_videos.loc['RPM (USD)'].astype(float) * peso_videos) + (df_lives.loc['RPM (USD)'].astype(float) * peso_lives)
    rpm_sem_shorts.loc['RPM sem Short'] = rpm_sem_short_vals
    views_velho = subtract_by_label(df, 'Visualizações', total_novo, 'Visualizações', "temp_index")
    receita_velho_df = subtract_by_label(df, 'Receita estimada (USD)', total_novo, 'Receita estimada (USD)', "temp_index")
    rpm_velho_vals = divisao_segura(receita_velho_df.values * 1000, views_velho.values)
    rpm_velho = pd.DataFrame(rpm_velho_vals, columns=df.columns, index=['RPM Velho'])
    rpm_novo_vals = divisao_segura(total_novo.loc[['Receita estimada (USD)']].values * 1000, total_novo.loc[['Visualizações']].values)
    rpm_novo = pd.DataFrame(rpm_novo_vals, columns=total_novo.columns, index=['RPM Novo'])
    views_novo_vod = videos_novos.loc[['Visualizações']].rename(index={"Visualizações":"Visualizações VOD's Novo"})
    views_novo_live = lives_novos.loc[['Visualizações']].rename(index={"Visualizações":"Visualizações Lives Novo"})
    views_novo_shorts = shorts_novos.loc[['Visualizações']].rename(index={"Visualizações":"Visualizações Shorts Novo"})
    views_velho_vod = subtract_by_label(df_videos, 'Visualizações', videos_novos, 'Visualizações', "Visualizações VOD's Velho")
    views_velho_live = subtract_by_label(df_lives, 'Visualizações', lives_novos, 'Visualizações', "Visualizações Lives Velho")
    views_velho_shorts = subtract_by_label(df_shorts, 'Visualizações', shorts_novos, 'Visualizações', "Visualizações Shorts Velho")
    views_total = df.loc[['Visualizações']].rename(index={"Visualizações":"Visualizações Total"})
    taxa_preenchimento = pd.DataFrame(index=['Taxa de Preenchimento'], columns=cols).fillna(0)
    if 'CPM (USD)' in df.index and 'CPM baseado em exibição (USD)' in df.index:
        taxa_preenchimento_vals = divisao_segura(df.loc[['CPM baseado em exibição (USD)']].values, df.loc[['CPM (USD)']].values)
        taxa_preenchimento = pd.DataFrame(taxa_preenchimento_vals, columns=df.columns, index=['Taxa de Preenchimento'])
    minutagem_videos = videos_novos.loc[['Duração']].rename(index={"Duração": "Tamanhos de VOD's Novo"})
    minutagem_lives = lives_novos.loc[['Duração']].rename(index={"Duração": "Tamanho de Lives Novo"})
    minutagem_shorts = shorts_novos.loc[['Duração']].rename(index={"Duração": "Tamanhos de Shorts Novo"})
    tempo_medio_vod_novo = videos_novos.loc[['Duração média da visualização']].rename(index={"Duração média da visualização":"Tempo Médio Assistido VOD's Novo"})
    tempo_medio_lives_novo = lives_novos.loc[['Duração média da visualização']].rename(index={"Duração média da visualização":"Tempo Médio Assistido Lives Novo"})
    tempo_medio_shorts_novo = shorts_novos.loc[['Duração média da visualização']].rename(index={"Duração média da visualização":"Tempo Médio Assistido Shorts Novo"})
    def calcular_engajamento_novo(df_novo, df_pubs, pub_label, content_type):
        new_index_name = f"Engajamento {content_type}"
        cols_engajamento = ['Comentários adicionados', 'Marcações "Gostei"', 'Compartilhamentos']
        if df_novo.empty or df_pubs.empty or not all(c in df_novo.index for c in cols_engajamento):
            return pd.DataFrame(index=[new_index_name], columns=cols).fillna(0)
        engaj_total = df_novo.loc[cols_engajamento].sum()
        engaj_df = pd.DataFrame(engaj_total).T; engaj_df.index = ['engaj_temp']
        pub_values = df_pubs.loc[pub_label].values
        result = divisao_segura(engaj_df.values, pub_values)
        return pd.DataFrame(result, columns=cols, index=[new_index_name])
    engajamento_vod_novo = calcular_engajamento_novo(videos_novos, vods_pubs, "Número de VOD's", 'VOD')
    engajamento_lives_novo = calcular_engajamento_novo(lives_novos, lives_pubs, "Números de Lives ", 'Lives')
    engajamento_shorts_novo = calcular_engajamento_novo(shorts_novos, shorts_pubs, "Números de Shorts", 'Shorts')
    inscritos_vod_novo = pd.DataFrame(0.0, index=['Número de Inscritos VOD Novo'], columns=cols)
    inscritos_live_novo = pd.DataFrame(0.0, index=['Número de Inscritos Live Novo'], columns=cols)
    inscritos_shorts_novo = pd.DataFrame(0.0, index=['Número de Inscritos Shorts Novo'], columns=cols)
    inscritos_vod_velho = pd.DataFrame(0.0, index=['Número de Inscritos vods Velho'], columns=cols)
    inscritos_lives_velho = pd.DataFrame(0.0, index=['Número de Inscritos Lives Velho'], columns=cols)
    inscritos_shorts_velho = pd.DataFrame(0.0, index=['Número de Inscritos Shorts Velho'], columns=cols)
    inscritos_vod = pd.DataFrame(0.0, index=["Número de Inscritos VOD's"], columns=cols)
    inscritos_lives = pd.DataFrame(0.0, index=['Número de Inscritos LIves'], columns=cols)
    inscritos_shorts = pd.DataFrame(0.0, index=['Número de Inscritos Shorts'], columns=cols)
    inscritos_sem_shorts = subtract_by_label(df, 'Inscritos', df_shorts, 'Inscritos', "Número de Inscritos Sem Shorts")
    inscritos_totais = pd.DataFrame(index=["Número de Inscritos Total"], columns=cols).fillna(0)
    sub_txt_path = f'dados_full/{artista}/sub.txt'
    if os.path.exists(sub_txt_path):
        with open(sub_txt_path, "r") as f: insc_ant = int(''.join((f.readline()).split('.')))
        inscritos_totais = pd.DataFrame((df.T['Inscritos'].astype(float).cumsum() + (insc_ant - df.T['Inscritos'].astype(float).sum()))).T.rename(index={"Inscritos":"Número de Inscritos Total"})
    
    df_geral_1 = pd.concat([receita_sem_shorts, receita_vod, receita_lives, receita_shorts, impress_sem_shorts, impress_vods, impress_lives, views_sem_shorts, views_shorts, views_vod, views_lives, vods_pubs, lives_pubs, shorts_pubs, watchtime])
    df_geral_2 = pd.concat([inscritos_vod, inscritos_lives, inscritos_shorts, inscritos_totais, inscritos_sem_shorts])
    lista_dfs_finais = [
        df_geral_1,
        metric_per_publication_novo(videos_novos, vods_pubs, 'Impressões', "Número de VOD's", 'Impressões'),
        metric_per_publication_novo(lives_novos, lives_pubs, 'Impressões', "Números de Lives ", 'Impressões'),
        metric_per_publication_novo(shorts_novos, shorts_pubs, 'Impressões', "Números de Shorts", 'Impressões'),
        rpm_velho, rpm_novo, rpm_sem_shorts,
        videos_novos.loc[['RPM (USD)']].rename(index={"RPM (USD)":"RPM VOD's Novo"}),
        lives_novos.loc[['RPM (USD)']].rename(index={"RPM (USD)":"RPM Lives Novo"}),
        shorts_novos.loc[['RPM (USD)']].rename(index={"RPM (USD)":"RPM Shorts Novo"}),
        videos_novos.loc[['Taxa de cliques de impressões (%)']].rename(index={"Taxa de cliques de impressões (%)":"CTR VOD's Novo"}),
        lives_novos.loc[['Taxa de cliques de impressões (%)']].rename(index={"Taxa de cliques de impressões (%)":"CTR Lives Novo"}),
        videos_novos.loc[['Porcentagem visualizada média (%)']].rename(index={"Porcentagem visualizada média (%)":"Porcentagem Média Assistida VOD's Novo"}),
        lives_novos.loc[['Porcentagem visualizada média (%)']].rename(index={"Porcentagem visualizada média (%)":"Porcentagem Média Assistida Live Novo"}),
        minutagem_videos, minutagem_lives,
        tempo_medio_vod_novo, tempo_medio_lives_novo,
        processar_fontes_de_trafego(artista, colunas_de_meses_desejadas),
        inscritos_vod_novo, inscritos_live_novo, inscritos_shorts_novo,
        inscritos_vod_velho, inscritos_lives_velho, inscritos_shorts_velho,
        df_geral_2,
        metric_per_publication_novo(videos_novos, vods_pubs, "Receita estimada (USD)", "Número de VOD's", "Receita"),
        metric_per_publication_novo(lives_novos, lives_pubs, "Receita estimada (USD)", "Números de Lives ", "Receita"),
        metric_per_publication_novo(shorts_novos, shorts_pubs, "Receita estimada (USD)", "Números de Shorts", "Receita"),
        videos_novos.loc[['Receita estimada (USD)']].rename(index={"Receita estimada (USD)":"Receita VOD Novo"}),
        lives_novos.loc[['Receita estimada (USD)']].rename(index={"Receita estimada (USD)":"Receita Live Novo"}),
        cpm,
        subtract_by_label(df_videos, 'Receita estimada (USD)', videos_novos, 'Receita estimada (USD)', "Receita VOD's Velho"),
        subtract_by_label(df_shorts, 'Receita estimada (USD)', shorts_novos, 'Receita estimada (USD)', "Receita Shorts Velho"),
        subtract_by_label(df_lives, 'Receita estimada (USD)', lives_novos, 'Receita estimada (USD)', "Receita Lives Velho"),
        shorts_novos.loc[['Receita estimada (USD)']].rename(index={"Receita estimada (USD)":"Receita Shorts Novo"}),
        taxa_preenchimento,
        engajamento_vod_novo, engajamento_lives_novo,
        impressoes_shorts,
        watchtime_vod, watchtime_lives, watchtime_shorts,
        rpm_total_vod, rpm_total_lives, rpm_total_shorts, rpm_total,
        inscritos_ganhos, inscritos_perdidos, saldo_inscritos,
        engajamento_shorts_novo,
        minutagem_shorts,
        shorts_novos.loc[['Porcentagem visualizada média (%)']].rename(index={"Porcentagem visualizada média (%)":"Porcentagem Média Assistida Shorts Novo"}),
        tempo_medio_shorts_novo,
        views_novo_vod, views_novo_live, views_novo_shorts,
        views_velho_vod, views_velho_live, views_velho_shorts,
        views_total
    ]

    lista_dfs_validos = [df for df in lista_dfs_finais if df is not None and not df.empty]
    tabela = pd.concat(lista_dfs_validos)
    
    # --- 4. Pós-processamento e Exportação ---
    colunas_de_meses = tabela.columns.tolist()
    tabela_calc = tabela.copy()
    for col in colunas_de_meses:
        tabela_calc[col] = pd.to_numeric(tabela_calc[col], errors='coerce')
    tabela_calc.fillna(0, inplace=True)
    media_series = tabela_calc[colunas_de_meses].mean(axis=1)
    tabela.insert(0, 'Média', media_series)
    tabela.index.name = "Data"
    
    tabela = tabela.astype(object)

    linhas_de_tempo_formatar = [ "Tamanhos de VOD's Novo", "Tamanho de Lives Novo", "Tempo Médio Assistido VOD's Novo", "Tempo Médio Assistido Lives Novo", "Tamanhos de Shorts Novo", "Tempo Médio Assistido Shorts Novo" ]
    colunas_para_formatar = tabela.columns.tolist()
    for metrica in linhas_de_tempo_formatar:
        if metrica in tabela.index:
            valores_numericos = pd.to_numeric(tabela.loc[metrica, colunas_para_formatar], errors='coerce').fillna(0)
            tabela.loc[metrica, colunas_para_formatar] = valores_numericos.apply( lambda total_seconds: f"{int(total_seconds // 60):02d}:{int(total_seconds % 60):02d}" )

    tabela_numerica_sem_media = tabela_calc[colunas_de_meses]
    resultado_array = divisao_segura(tabela_numerica_sem_media.sub(media_series, axis=0), media_series.values)
    tabela_desvio_media = pd.DataFrame(resultado_array, index=tabela.index, columns=colunas_de_meses)
    df_shifted = tabela_numerica_sem_media.shift(1, axis=1)
    tabela_desvio_anterior = divisao_segura(tabela_numerica_sem_media.sub(df_shifted), df_shifted)
    
    output_filename = f'exports_tabelas/tabela_4.1_{artista}.xlsx'
    os.makedirs('exports_tabelas', exist_ok=True)
    
    with pd.ExcelWriter(output_filename, engine='openpyxl') as writer:
        tabela.to_excel(writer, sheet_name='Resultado', index=True)
        tabela_desvio_media.to_excel(writer, sheet_name='Desvio', index=True)
        tabela_desvio_anterior.to_excel(writer, sheet_name='Mês Anterior', index=True)

    print(f"Relatório para {artista} salvo com sucesso ✅")


def buscar_lista_artistas():
    caminho_exports = os.path.join(path, 'exports.txt')
    if not os.path.exists(caminho_exports): return []
    with open(caminho_exports) as f: lines = f.readlines()
    return [i.strip() for i in lines if i.strip()]


def run():
    lista_de_artistas = buscar_lista_artistas()
    for artista_selecionado in lista_de_artistas:
        try:
            gerar_relatorio_para_artista(artista_selecionado)
        except FileNotFoundError as e:
            print(f"\nERRO: Arquivo não encontrado ao processar '{artista_selecionado}': {e.filename}\n")
        except Exception as e:
            print(f"\nERRO: Ocorreu um problema inesperado ao processar '{artista_selecionado}': {e}\n")
            import traceback
            traceback.print_exc()


if __name__ == '__main__':
    run()