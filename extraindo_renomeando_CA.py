import os
import sys
import glob
import zipfile
import re


def remover_csv_antigos(artista):
    # (código idêntico ao anterior)
    arquivos_csv_path = glob.glob(f"dados_full/{artista}/*.csv")
    for arquivo_csv in arquivos_csv_path:
        if os.path.basename(arquivo_csv) != 'postagem.csv':
            try:
                os.remove(arquivo_csv)
            except OSError as e:
                print(f"Erro ao remover o arquivo {arquivo_csv}: {e}")

def identificar_arquivos_zip(artista):
    arquivos_zip_path = glob.glob(f"dados_full/{artista}/raw_data/*.zip")
    arquivos_processados = 0

    # --- Lógica Dinâmica de Mapeamento de Datas a partir dos Nomes dos Arquivos ---
    
    # 1. Encontra o mês para os arquivos "Data"
    prefixo_data = None
    data_pattern = re.compile(r"Data (\d{4}-\d{2})")
    for arq_path in arquivos_zip_path:
        match = data_pattern.search(os.path.basename(arq_path))
        if match:
            prefixo_data = f"Data {match.group(1)}"
            break # Encontrou o primeiro, assume que é o único
    
    # 2. Cria o mapa de meses para os arquivos "Conteúdo"
    conteudo_meses = set()
    conteudo_pattern = re.compile(r"Conteúdo (\d{4}-\d{2})")
    for arq_path in arquivos_zip_path:
        match = conteudo_pattern.search(os.path.basename(arq_path))
        if match:
            conteudo_meses.add(match.group(1))

    # Ordena os meses do mais antigo para o mais novo para criar os sufixos _01, _02, etc.
    mapa_meses = {mes: f"_{i+1:02d}" for i, mes in enumerate(sorted(list(conteudo_meses)))}
    
    print(f"Para {artista} -> Prefixo 'Data' detectado: {prefixo_data}")
    print(f"Para {artista} -> Mapa de 'Conteúdo' gerado: {mapa_meses}")


    for arq_zip_path in arquivos_zip_path:
        nome_arq_zip = os.path.basename(arq_zip_path)
        path_extracao = os.path.dirname(arq_zip_path)
        sucesso = False

        with zipfile.ZipFile(arq_zip_path, "r") as zip_ref:

            # --- Bloco 1: Arquivos 'Conteúdo' (lógica dinâmica) ---
            if nome_arq_zip.startswith('Conteúdo'):
                try:
                    chave_mes_arquivo = nome_arq_zip.split(' ')[1][:7]
                except IndexError:
                    continue # Ignora arquivos mal formatados

                if chave_mes_arquivo in mapa_meses:
                    sufixo = mapa_meses[chave_mes_arquivo]
                    zip_ref.extractall(path_extracao)
                    
                    if nome_arq_zip.endswith('(1).zip'):
                        os.rename(os.path.join(path_extracao, "Dados da tabela.csv"), f"dados_full/{artista}/lives{sufixo}.csv")
                    elif nome_arq_zip.endswith('(2).zip'):
                        os.rename(os.path.join(path_extracao, "Dados da tabela.csv"), f"dados_full/{artista}/shorts{sufixo}.csv")
                    else:
                        os.rename(os.path.join(path_extracao, "Dados da tabela.csv"), f"dados_full/{artista}/videos{sufixo}.csv")
                    sucesso = True
            
            # --- Bloco 2: Arquivos 'Data' (lógica dinâmica) ---
            elif prefixo_data and nome_arq_zip.startswith(prefixo_data):
                zip_ref.extractall(path_extracao)
                if nome_arq_zip.endswith('(1).zip'):
                    os.rename(os.path.join(path_extracao, "Dados da tabela.csv"), f"dados_full/{artista}/videos.csv")
                elif nome_arq_zip.endswith('(2).zip'):
                    os.rename(os.path.join(path_extracao, "Dados da tabela.csv"), f"dados_full/{artista}/lives.csv")
                elif nome_arq_zip.endswith('(3).zip'):
                    os.rename(os.path.join(path_extracao, "Dados da tabela.csv"), f"dados_full/{artista}/shorts.csv")
                else:
                    os.rename(os.path.join(path_extracao, "Dados da tabela.csv"), f"dados_full/{artista}/total.csv")
                sucesso = True

            # --- Bloco 3: Outros arquivos com nomes fixos ---
            elif nome_arq_zip.startswith('Origem'):
                zip_ref.extractall(path_extracao)
                nome_final = 'origem_lives.csv' if nome_arq_zip.endswith('(1).zip') else 'origem_vods.csv'
                os.rename(os.path.join(path_extracao, "Dados do gráfico.csv"), f"dados_full/{artista}/{nome_final}")
                sucesso = True

            elif nome_arq_zip.startswith('Postar'):
                zip_ref.extractall(path_extracao)
                os.rename(os.path.join(path_extracao, "Dados da tabela.csv"), f"dados_full/{artista}/comunidade.csv")
                sucesso = True
        
        if sucesso:
            arquivos_processados += 1

    if arquivos_processados == len(arquivos_zip_path):
        print(f'Sucesso! {arquivos_processados} arquivos foram processados para {artista}.')
    else:
        nao_processados = len(arquivos_zip_path) - arquivos_processados
        print(f'AVISO: {nao_processados} arquivos não foram processados para {artista}.')

def run(artista):
    print(f"\n--- Iniciando processamento CA para: {artista} ---")
    remover_csv_antigos(artista)
    identificar_arquivos_zip(artista)
    print(f"--- Finalizado: {artista} ---")


if __name__ == "__main__":
    # Pega o nome do artista do argumento passado pelo main.py
    if len(sys.argv) < 2:
        print("Erro: Nenhum artista fornecido. Este script deve ser chamado pelo main.py")
        sys.exit(1) # Sai com erro
    
    artista_argumento = sys.argv[1]
    
    # Executa a função run APENAS para esse artista
    run(artista_argumento)
