import os
import glob
import zipfile
from datetime import date
from dateutil.relativedelta import relativedelta # Requer: pip install python-dateutil


def buscar_lista_artistas():
    # Acesso exports.txt para buscar o nome dos artistas
    lines = ''
    with open('exports.txt') as f:
        lines = f.readlines()
    lines = [i.rstrip() for i in lines]
    return lines


def remover_csv_antigos(artista):
    # Para cada artista, remove os .csv antigos caso existam
    arquivos_csv_path = glob.glob("dados_full/"+artista+'/*.csv')

    for arquivo_csv in arquivos_csv_path:
        # Verifica se o nome do arquivo é 'postagem.csv'
        if os.path.basename(arquivo_csv) == 'postagem.csv':
            continue  # Pula para a próxima iteração do loop sem excluir o arquivo

        try:
            os.remove(arquivo_csv)
        except:
            pass


# ==============================================================================
# VERSÃO FINAL DA FUNÇÃO identificar_arquivos_zip
# ==============================================================================
def identificar_arquivos_zip(artista):
    # --- Lógica Dinâmica de Mapeamento de Datas ---
    hoje = date.today()

    # 1. Define o prefixo dinâmico para os arquivos 'Data' (mês atual - 7 meses)
    mes_data = hoje - relativedelta(months=7)
    prefixo_data = f"Data {mes_data.strftime('%Y-%m')}" # Ex: 'Data 2025-02'

    # --- ALTERAÇÃO AQUI: Lógica do mapeamento de meses invertida ---
    # 2. Cria o mapa de meses para os arquivos 'Conteúdo' (últimos 6 meses)
    #    Agora, _01 = mês mais antigo | _06 = mês mais recente
    mapa_meses = {}
    # delta_meses irá de 6, 5..1 (do mais antigo para o mais novo)
    # num_sufixo irá de 1, 2..6 (do menor para o maior)
    for delta_meses, num_sufixo in zip(range(6, 0, -1), range(1, 7)):
        mes_alvo = hoje - relativedelta(months=delta_meses)
        chave_mes = mes_alvo.strftime('%Y-%m')
        sufixo = f"_{num_sufixo:02d}"
        mapa_meses[chave_mes] = sufixo
    # ----------------------------------------------------------------------

    arquivos_zip_path = glob.glob(f"dados_full/{artista}/raw_data/*.zip")
    arquivos_processados = 0

    for arq_zip_path in arquivos_zip_path:
        nome_arq_zip = os.path.basename(arq_zip_path)
        sucesso = False
        path_extracao = os.path.dirname(arq_zip_path)

        with zipfile.ZipFile(arq_zip_path, "r") as zip_ref:

            # --- Bloco 1: Arquivos 'Conteúdo' (lógica dinâmica) ---
            if nome_arq_zip.startswith('Conteúdo'):
                try:
                    chave_mes_arquivo = nome_arq_zip.split(' ')[1][:7]
                except IndexError:
                    print(f"AVISO: Formato de nome inesperado, não foi possível extrair data: {nome_arq_zip}")
                    continue

                if chave_mes_arquivo in mapa_meses:
                    sufixo = mapa_meses[chave_mes_arquivo]
                    zip_ref.extractall(path_extracao)
                    
                    if nome_arq_zip.endswith('(1).zip'):
                        os.rename(os.path.join(path_extracao, "Dados da tabela.csv"), f"dados_full/{artista}/lives{sufixo}.csv")
                        sucesso = True
                    elif nome_arq_zip.endswith('(2).zip'):
                        os.rename(os.path.join(path_extracao, "Dados da tabela.csv"), f"dados_full/{artista}/shorts{sufixo}.csv")
                        sucesso = True
                    else: # Arquivo "sem nada" no final
                        os.rename(os.path.join(path_extracao, "Dados da tabela.csv"), f"dados_full/{artista}/videos{sufixo}.csv")
                        sucesso = True
            
            # --- Bloco 2: Arquivos 'Data' (lógica dinâmica) ---
            elif nome_arq_zip.startswith(prefixo_data):
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
                if nome_arq_zip.endswith('(1).zip'):
                    os.rename(os.path.join(path_extracao, "Dados do gráfico.csv"), f"dados_full/{artista}/origem_lives.csv")
                else:
                    os.rename(os.path.join(path_extracao, "Dados do gráfico.csv"), f"dados_full/{artista}/origem_vods.csv")
                sucesso = True

            elif nome_arq_zip.startswith('Postar'):
                zip_ref.extractall(path_extracao)
                os.rename(os.path.join(path_extracao, "Dados da tabela.csv"), f"dados_full/{artista}/comunidade.csv")
                sucesso = True
        
        if sucesso:
            arquivos_processados += 1

    # Relatório final
    if arquivos_processados == len(arquivos_zip_path):
        print('Todos os arquivos foram processados e renomeados com sucesso!')
    else:
        nao_processados = len(arquivos_zip_path) - arquivos_processados
        print(f'Arquivos não processados: {nao_processados}')


def run():
    lista_de_artistas = buscar_lista_artistas()
    
    for artista in lista_de_artistas:
        remover_csv_antigos(artista)
        identificar_arquivos_zip(artista)

        print('Done: '+ artista)


if __name__ == "__main__":
    run()