import os
import json
import subprocess
import sys

def buscar_lista_artistas():
    try:
        with open('exports.txt', 'r', encoding='utf-8') as f:
            return [line.strip() for line in f if line.strip()]
    except FileNotFoundError:
        print("ERRO CRÍTICO: Arquivo 'exports.txt' não encontrado.")
        return []

def carregar_configuracao(filepath):
    """Função genérica para carregar um arquivo JSON."""
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            return json.load(f)
    except FileNotFoundError:
        print(f"[AVISO] Arquivo '{filepath}' não encontrado. Retornando configuração vazia.")
        return {}
    except json.JSONDecodeError:
        print(f"ERRO: O arquivo '{filepath}' contém um erro de sintaxe. Retornando configuração vazia.")
        return {}

def executar_script(script_name, artista):
    if not os.path.exists(script_name):
        print(f"    -> AVISO: O script '{script_name}' não foi encontrado. Pulando.")
        return False
    print(f"    -> Executando '{script_name}'...")
    try:
        subprocess.run(
            [sys.executable, script_name, artista], 
            check=True, text=True, capture_output=True, encoding='utf-8'
        )
        print(f"    -> SUCESSO: '{script_name}' finalizado.")
        return True
    except subprocess.CalledProcessError as e:
        print(f"    -> ERRO na execução de '{script_name}':\n{e.stderr}")
        return False
    except Exception as e:
        print(f"    -> Um erro inesperado ocorreu ao executar '{script_name}': {e}")
        return False

def main():
    """Função principal que orquestra todo o fluxo de execução."""
    config_principal = carregar_configuracao('config.json').get('artistas', {})
    mapa_de_apresentacao = carregar_configuracao('presentation_config.json')
    lista_de_artistas = buscar_lista_artistas()

    if not lista_de_artistas or not config_principal:
        print("Execução encerrada devido a erros de configuração nos arquivos principais.")
        return

    print("=======================================================")
    print("           INICIANDO FLUXO DE PROCESSAMENTO            ")
    print(f"Artistas a serem processados: {', '.join(lista_de_artistas)}")
    print("=======================================================\n")

    for artista in lista_de_artistas:
        print(f"-------------------------------------------------------")
        print(f"PROCESSANDO ARTISTA: {artista}")
        print(f"-------------------------------------------------------")

        if artista not in config_principal:
            print(f"AVISO: O artista '{artista}' não foi encontrado no config.json. Pulando.")
            continue

        config_artista = config_principal[artista]
        tipo_proc = config_artista.get("tipo_processamento", "padrao").upper()

        # Define e executa os scripts para os grupos 1 a 4
        if tipo_proc == "CA":
            scripts_base = {"grupo_1_extracao": "extraindo_renomeando_CA.py", "grupo_2_tratamento": "tratamento_CA.py", "grupo_3_report": "report_CA.py"}
        else:
            scripts_base = {"grupo_1_extracao": "extraindo_renomeando.py", "grupo_2_tratamento": "tratamento.py", "grupo_3_report": "report.py"}
        
        scripts_a_executar = {**scripts_base, **config_artista}
        
        for i in range(1, 5):
            nome_grupo = f"grupo_{i}_{['extracao', 'tratamento', 'report', 'graficos'][i-1]}"
            print(f"\n  [GRUPO: {nome_grupo.upper()}]")
            script = scripts_a_executar.get(nome_grupo)
            if script:
                executar_script(script, artista)
            else:
                print(f"    -> Nenhum script definido para este grupo. Pulando.")

        # --- Lógica Final para o Grupo 5 (Apresentação) ---
        print("\n  [GRUPO: GRUPO_5_APRESENTACAO]")
        
        # 1. Pega a lista de apresentações a serem geradas do 'presentation_config.json'
        #    Se o artista não estiver lá, o padrão é ['report']
        tipos_desejados = mapa_de_apresentacao.get(artista, ["report"])
        print(f"    -> Apresentações definidas em 'presentation_config.json': {tipos_desejados}")

        # 2. Pega o "cardápio" de scripts de apresentação possíveis do 'config.json'
        apresentacoes_possiveis = config_artista.get("grupo_5_apresentacao", [])
        if not isinstance(apresentacoes_possiveis, list):
            apresentacoes_possiveis = [apresentacoes_possiveis]

        # 3. Filtra o "cardápio" com base no "pedido"
        scripts_para_rodar_hoje = []
        for tipo in tipos_desejados:
            for script_possivel in apresentacoes_possiveis:
                if tipo in script_possivel:
                    scripts_para_rodar_hoje.append(script_possivel)

        if not scripts_para_rodar_hoje:
            print("    -> Nenhum script de apresentação corresponde à seleção para hoje.")
        else:
            print(f"    -> Scripts a serem executados: {list(set(scripts_para_rodar_hoje))}")
            for script in set(scripts_para_rodar_hoje): # `set` evita rodar o mesmo script duas vezes
                executar_script(script, artista)

    print("\n=======================================================")
    print("            FLUXO DE PROCESSAMENTO FINALIZADO            ")
    print("=======================================================")

if __name__ == "__main__":
    main()