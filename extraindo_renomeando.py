import os
import sys
import glob
import zipfile
import re
from collections import Counter

# --- CONFIGURAÇÃO DE IDIOMAS E PALAVRAS-CHAVE ---
# O script usará estas palavras para adivinhar o idioma dos arquivos.
CONFIG = {
    'pt': {
        'keywords': ['Data', 'Origem', 'Postar'],
        'table_data_name': 'Dados da tabela.csv',
        'chart_data_name': 'Dados do gráfico.csv'
    },
    'en': {
        'keywords': ['Date', 'Traffic source', 'Post'],
        'table_data_name': 'Table data.csv',
        'chart_data_name': 'Chart data.csv'
    },
    'es': {
        'keywords': ['Fecha', 'Fuente', 'Publicaci'],
        'table_data_name': 'Datos de la tabla.csv',
        'chart_data_name': 'Datos del gráfico.csv'
    }
}

# --- MAPEAMENTO DE NOMES DE ARQUIVOS (sem alteração) ---
OLDEST_MONTH_MAP = {
    '': 'total.csv', '(1)': 'videos.csv', '(2)': 'lives.csv', '(3)': 'shorts.csv'
}
NEWEST_MONTH_MAP = {
    '': 'videos_01.csv', '(1)': 'lives_01.csv', '(2)': 'shorts_01.csv',
    '(3)': 'videos_02.csv', '(4)': 'lives_02.csv', '(5)': 'shorts_02.csv',
    '(6)': 'videos_03.csv', '(7)': 'lives_03.csv', '(8)': 'shorts_03.csv',
    '(9)': 'videos_04.csv', '(10)': 'lives_04.csv', '(11)': 'shorts_04.csv',
    '(12)': 'videos_05.csv', '(13)': 'lives_05.csv', '(14)': 'shorts_05.csv',
    '(15)': 'videos_06.csv', '(16)': 'lives_06.csv', '(17)': 'shorts_06.csv'
}

def remover_csv_antigos(artista):
    """Para cada artista, remove os .csv antigos, exceto 'postagem.csv'."""
    arquivos_csv_path = glob.glob(f"dados_full/{artista}/*.csv")
    for arquivo_csv in arquivos_csv_path:
        if os.path.basename(arquivo_csv) != 'postagem.csv':
            try:
                os.remove(arquivo_csv)
            except OSError as e:
                print(f"Erro ao remover o arquivo {arquivo_csv}: {e}")

def detectar_idioma(arquivos_zip_path):
    """
    Analisa os nomes dos arquivos .zip para detectar o idioma (pt, en, es).
    Retorna o código do idioma ('pt', 'en', 'es') ou 'pt' como padrão.
    """
    if not arquivos_zip_path:
        return 'pt' # Retorna padrão se não houver arquivos

    contador_idioma = Counter()
    nomes_arquivos = [os.path.basename(p) for p in arquivos_zip_path]

    for lang_code, config in CONFIG.items():
        for keyword in config['keywords']:
            for nome_arquivo in nomes_arquivos:
                if nome_arquivo.startswith(keyword):
                    contador_idioma[lang_code] += 1
    
    # Retorna o idioma com a maior contagem de palavras-chave.
    # Se nenhum for encontrado, retorna 'pt' como padrão.
    if not contador_idioma:
        return 'pt'
    
    idioma_detectado = contador_idioma.most_common(1)[0][0]
    return idioma_detectado


def identificar_arquivos_zip(artista):
    """
    Função dinâmica que detecta o idioma, identifica, extrai e renomeia os arquivos.
    """
    arquivos_zip_path = glob.glob(f"dados_full/{artista}/raw_data/*.zip")
    if not arquivos_zip_path:
        print(f"Nenhum arquivo .zip encontrado para o artista {artista}.")
        return

    # --- 1. DETECÇÃO AUTOMÁTICA DE IDIOMA ---
    lang = detectar_idioma(arquivos_zip_path)
    cfg = CONFIG[lang]
    date_keyword = cfg['keywords'][0] # Assume que a primeira keyword é sempre a de data
    print(f"Idioma detectado para {artista}: {lang.upper()}")

    # --- 2. Lógica para detecção dinâmica de meses ---
    meses_encontrados = set()
    date_pattern = re.compile(rf"{re.escape(date_keyword)}\s(\d{{4}}-\d{{2}})")
    
    for arq_path in arquivos_zip_path:
        match = date_pattern.search(os.path.basename(arq_path))
        if match:
            meses_encontrados.add(match.group(1))
    
    sorted_meses = sorted(list(meses_encontrados))
    
    mes_antigo, mes_recente = None, None
    if len(sorted_meses) >= 2:
        mes_antigo = sorted_meses[0]
        mes_recente = sorted_meses[-1]
    elif len(sorted_meses) == 1:
        mes_antigo = mes_recente = sorted_meses[0]
        print(f"AVISO: Apenas um mês de dados ('{mes_antigo}') encontrado para {artista}.")

    if mes_antigo:
        print(f"Processando para {artista}: Mês antigo: {mes_antigo}, Mês recente: {mes_recente}")

    # --- 3. Processamento e Renomeação ---
    arquivos_processados = 0
    for arq_zip_path in arquivos_zip_path:
        nome_arq_zip = os.path.basename(arq_zip_path)
        path_extracao = os.path.dirname(arq_zip_path)
        sucesso = False

        with zipfile.ZipFile(arq_zip_path, "r") as zip_ref:
            sufixo_match = re.search(r'(\(\d+\))\.zip$', nome_arq_zip)
            sufixo = sufixo_match.group(1) if sufixo_match else ''
            
            # Mês antigo
            if mes_antigo and nome_arq_zip.startswith(f"{date_keyword} {mes_antigo}"):
                if sufixo in OLDEST_MONTH_MAP:
                    novo_nome = OLDEST_MONTH_MAP[sufixo]
                    zip_ref.extractall(path_extracao)
                    os.rename(os.path.join(path_extracao, cfg['table_data_name']), f"dados_full/{artista}/{novo_nome}")
                    sucesso = True
            
            # Mês recente
            elif mes_recente and nome_arq_zip.startswith(f"{date_keyword} {mes_recente}"):
                if sufixo in NEWEST_MONTH_MAP:
                    novo_nome = NEWEST_MONTH_MAP[sufixo]
                    zip_ref.extractall(path_extracao)
                    os.rename(os.path.join(path_extracao, cfg['table_data_name']), f"dados_full/{artista}/{novo_nome}")
                    sucesso = True

            # Outros arquivos (usando as keywords do idioma detectado)
            elif nome_arq_zip.startswith(cfg['keywords'][1]): # Traffic/Origem/Fuente
                novo_nome = 'origem_lives.csv' if nome_arq_zip.endswith('(1).zip') else 'origem_vods.csv'
                zip_ref.extractall(path_extracao)
                os.rename(os.path.join(path_extracao, cfg['chart_data_name']), f"dados_full/{artista}/{novo_nome}")
                sucesso = True

            elif nome_arq_zip.startswith(cfg['keywords'][2]): # Post/Postar/Publicaci
                zip_ref.extractall(path_extracao)
                os.rename(os.path.join(path_extracao, cfg['table_data_name']), f"dados_full/{artista}/comunidade.csv")
                sucesso = True

        if sucesso:
            arquivos_processados += 1

    if arquivos_processados == len(arquivos_zip_path):
        print(f'Sucesso! {arquivos_processados} arquivos foram processados para {artista}.')
    else:
        nao_processados = len(arquivos_zip_path) - arquivos_processados
        print(f'AVISO: {nao_processados} arquivos não foram processados para {artista}.')


def run(artista):
    
    print(f"Extraindo dados para: {artista}")
    print(f"Finalizado: {artista}")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Erro: Nenhum artista fornecido. Este script deve ser chamado pelo main.py")
        sys.exit(1)
    
    artista_argumento = sys.argv[1]
    run(artista_argumento)
