import os
import glob
import zipfile


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



def identificar_arquivos_zip(artista):
    arquivos_zip_path = glob.glob("dados_full/"+artista+'/raw_data/*.zip')
    arquivos_renomeados = 0
    
    for arq_zip_path in arquivos_zip_path:
        with zipfile.ZipFile(arq_zip_path, "r") as zip_ref:
            path = os.path.join(*arq_zip_path.split("/")[:-1])
            arq_zip = arq_zip_path.split("/")[3]
        
            if arq_zip.startswith("Fecha 2024-12") and not arq_zip.endswith('(1).zip') and not arq_zip.endswith('(2).zip') and not arq_zip.endswith('(3).zip'):
                zip_ref.extractall(path)
                os.rename(path + "/Datos de la tabla.csv", "dados_full/"+artista+"/total.csv")
                arquivos_renomeados += 1

            elif arq_zip.startswith('Fecha 2024-12') and arq_zip.endswith('(1).zip'):
                zip_ref.extractall(path)
                os.rename(path + "/Datos de la tabla.csv", "dados_full/"+artista+"/videos.csv")
                arquivos_renomeados += 1

            elif arq_zip.startswith('Fecha 2024-12') and arq_zip.endswith('(2).zip'):
                zip_ref.extractall(path)
                os.rename(path + "/Datos de la tabla.csv", "dados_full/"+artista+"/lives.csv")
                arquivos_renomeados += 1

            elif arq_zip.startswith('Fecha 2024-12') and arq_zip.endswith('(3).zip'):
                zip_ref.extractall(path)
                os.rename(path + "/Datos de la tabla.csv", "dados_full/"+artista+"/shorts.csv")
                arquivos_renomeados += 1

            elif arq_zip.startswith('Fecha 2025-01') and not arq_zip.endswith('(1).zip') and not arq_zip.endswith('(2).zip') and not arq_zip.endswith('(3).zip') and not arq_zip.endswith('(4).zip') and not arq_zip.endswith('(5).zip') and not arq_zip.endswith('(6).zip') and not arq_zip.endswith('(7).zip') and not arq_zip.endswith('(8).zip') and not arq_zip.endswith('(9).zip') and not arq_zip.endswith('(10).zip') and not arq_zip.endswith('(11).zip') and not arq_zip.endswith('(12).zip') and not arq_zip.endswith('(13).zip') and not arq_zip.endswith('(14).zip') and not arq_zip.endswith('(15).zip') and not arq_zip.endswith('(16).zip') and not arq_zip.endswith('(17).zip'):
                zip_ref.extractall(path)
                os.rename(path + "/Datos de la tabla.csv", "dados_full/"+artista+"/videos_01.csv")
                arquivos_renomeados += 1

            elif arq_zip.startswith('Fecha 2025-01') and arq_zip.endswith('(1).zip'):
                zip_ref.extractall(path)
                os.rename(path + "/Datos de la tabla.csv", "dados_full/"+artista+"/lives_01.csv")
                arquivos_renomeados += 1

            elif arq_zip.startswith('Fecha 2025-01') and arq_zip.endswith('(2).zip'):
                zip_ref.extractall(path)
                os.rename(path + "/Datos de la tabla.csv", "dados_full/"+artista+"/shorts_01.csv")
                arquivos_renomeados += 1

            elif arq_zip.startswith('Fecha 2025-01') and arq_zip.endswith('(3).zip'):
                zip_ref.extractall(path)
                os.rename(path + "/Datos de la tabla.csv", "dados_full/"+artista+"/videos_02.csv")
                arquivos_renomeados += 1

            elif arq_zip.startswith('Fecha 2025-01') and arq_zip.endswith('(4).zip'):
                zip_ref.extractall(path)
                os.rename(path + "/Datos de la tabla.csv", "dados_full/"+artista+"/lives_02.csv")
                arquivos_renomeados += 1

            elif arq_zip.startswith('Fecha 2025-01') and arq_zip.endswith('(5).zip'):
                zip_ref.extractall(path)
                os.rename(path + "/Datos de la tabla.csv", "dados_full/"+artista+"/shorts_02.csv")
                arquivos_renomeados += 1

            elif arq_zip.startswith('Fecha 2025-01') and arq_zip.endswith('(6).zip'):
                zip_ref.extractall(path)
                os.rename(path + "/Datos de la tabla.csv", "dados_full/"+artista+"/videos_03.csv")
                arquivos_renomeados += 1

            elif arq_zip.startswith('Fecha 2025-01') and arq_zip.endswith('(7).zip'):
                zip_ref.extractall(path)
                os.rename(path + "/Datos de la tabla.csv", "dados_full/"+artista+"/lives_03.csv")
                arquivos_renomeados += 1

            elif arq_zip.startswith('Fecha 2025-01') and arq_zip.endswith('(8).zip'):
                zip_ref.extractall(path)
                os.rename(path + "/Datos de la tabla.csv", "dados_full/"+artista+"/shorts_03.csv")
                arquivos_renomeados += 1

            elif arq_zip.startswith('Fecha 2025-01') and arq_zip.endswith('(9).zip'):
                zip_ref.extractall(path)
                os.rename(path + "/Datos de la tabla.csv", "dados_full/"+artista+"/videos_04.csv")
                arquivos_renomeados += 1

            elif arq_zip.startswith('Fecha 2025-01') and arq_zip.endswith('(10).zip'):
                zip_ref.extractall(path)
                os.rename(path + "/Datos de la tabla.csv", "dados_full/"+artista+"/lives_04.csv")
                arquivos_renomeados += 1

            elif arq_zip.startswith('Fecha 2025-01') and arq_zip.endswith('(11).zip'):
                zip_ref.extractall(path)
                os.rename(path + "/Datos de la tabla.csv", "dados_full/"+artista+"/shorts_04.csv")
                arquivos_renomeados += 1

            elif arq_zip.startswith('Fecha 2025-01') and arq_zip.endswith('(12).zip'):
                zip_ref.extractall(path)
                os.rename(path + "/Datos de la tabla.csv", "dados_full/"+artista+"/videos_05.csv")
                arquivos_renomeados += 1

            elif arq_zip.startswith('Fecha 2025-01') and arq_zip.endswith('(13).zip'):
                zip_ref.extractall(path)
                os.rename(path + "/Datos de la tabla.csv", "dados_full/"+artista+"/lives_05.csv")
                arquivos_renomeados += 1

            elif arq_zip.startswith('Fecha 2025-01') and arq_zip.endswith('(14).zip'):
                zip_ref.extractall(path)
                os.rename(path + "/Datos de la tabla.csv", "dados_full/"+artista+"/shorts_05.csv")
                arquivos_renomeados += 1

            elif arq_zip.startswith('Fecha 2025-01') and arq_zip.endswith('(15).zip'):
                zip_ref.extractall(path)
                os.rename(path + "/Datos de la tabla.csv", "dados_full/"+artista+"/videos_06.csv")
                arquivos_renomeados += 1

            elif arq_zip.startswith('Fecha 2025-01') and arq_zip.endswith('(16).zip'):
                zip_ref.extractall(path)
                os.rename(path + "/Datos de la tabla.csv", "dados_full/"+artista+"/lives_06.csv")
                arquivos_renomeados += 1

            elif arq_zip.startswith('Fecha 2025-01') and arq_zip.endswith('(17).zip'):
                zip_ref.extractall(path)
                os.rename(path + "/Datos de la tabla.csv", "dados_full/"+artista+"/shorts_06.csv")
                arquivos_renomeados += 1

            elif arq_zip.startswith('Fuente') and not arq_zip.endswith('(1).zip'):
                zip_ref.extractall(path)
                os.rename(path + "/Datos del gráfico.csv", "dados_full/"+artista+"/origem_vods.csv")
                arquivos_renomeados += 1

            elif arq_zip.startswith('Fuente') and arq_zip.endswith('(1).zip'):
                zip_ref.extractall(path)
                os.rename(path + "/Datos del gráfico.csv", "dados_full/"+artista+"/origem_lives.csv")
                arquivos_renomeados += 1

            elif arq_zip.startswith('Publicaci') and arq_zip.endswith('.zip'):
                zip_ref.extractall(path)
                os.rename(path + "/Datos de la tabla.csv", "dados_full/"+artista+"/comunidade.csv")
                arquivos_renomeados += 1


    if arquivos_renomeados == len(arquivos_zip_path):
        print('Arquivos renomeados com sucesso!')
    
    else:
        arquivos_nao_renomeados = len(arquivos_zip_path) - arquivos_renomeados
        print('Arquivos não renomeados: ', arquivos_nao_renomeados)
        

def run():
    lista_de_artistas = buscar_lista_artistas()
    
    for artista in lista_de_artistas:
        remover_csv_antigos(artista)
        identificar_arquivos_zip(artista)

        print('Done: '+ artista)


if __name__ == "__main__":
    run()
