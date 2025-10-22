# Projeto Orquestrador de Processamento e Relatórios de Overview 1BMG

## 1. Visão Geral

Este projeto automatiza o fluxo de trabalho de extração, tratamento, análise e geração de relatórios para dados de múltiplos artistas. Ele é projetado para ser flexível e configurável, permitindo que diferentes artistas sigam fluxos de processamento distintos através de um sistema de configuração centralizado.

O processo é controlado por um script principal, `main.py`, que lê as configurações e executa os scripts necessários na ordem correta, desde a extração de dados brutos de arquivos `.zip` até a criação de apresentações `.pptx` com gráficos e análises.

## 2. Extração dos Arquivos .zip (Etapa Manual com Macros)

Antes de executar o fluxo de automação em Python, os arquivos de dados brutos (`.zip`) precisam ser baixados do YouTube Studio. Este processo é semi-automatizado através de macros de automação de navegador reproduzidos através da ferramenta *Ui.Vision* baseada em Selenium IDE. Para um melhor funcionamento da extensão recomenda-se utiliza-la no navegador Firefox.

Instale a extensão para executa-la: https://addons.mozilla.org/pt-BR/firefox/addon/rpa/


### Ferramenta Utilizada
- **Tipo de Ferramenta:** Ui.Vision RPA - Extensão de automação de navegador.
- **Arquivos de Macro:** `3 - Extração Overview.json`, `3 - Extração Overview Inglês.json`, `3 - Extração Overview C.A.json`.

### Descrição das Macros

Existem três macros distintas, cada uma projetada para um tipo específico de canal. É crucial usar a macro correta para cada artista para garantir que os arquivos sejam baixados com os nomes e formatos esperados pelos scripts de automação.

---

#### Macro 1: Canais Padrão (Português)
- **Arquivo da Macro:** `3 - Extração Overview.json`
- **Para quais canais:** A maioria dos canais, com idioma configurado em português.
- **O que ela faz:**
    1.  Acessa o YouTube Studio do canal.
    2.  Navega para a área de **Estatísticas -> Modo Avançado**.
    3.  Executa uma série de cliques e digitação para baixar **22 relatórios** no total:
        - **Dados do Mês Antigo:** Baixa 4 relatórios de "Data" para o primeiro mês de análise.
        - **Dados do Mês Recente:** Baixa 18 relatórios de "Data" para o segundo mês de análise.
        - **Dados Complementares:** Baixa os relatórios de "Origem de tráfego" (VODs e Lives) e "Postagens na Comunidade".
- **É necessário ajustar as datas presentes na macro para o período que quer analisar**


---

#### Macro 2: Canais em Inglês
- **Arquivo da Macro:** `3 - Extração Overview Inglês.json`
- **Para quais canais:** Canais com idioma configurado em inglês (ex: 3LittleWords, deiveLeonardoEnglish).
- **O que ela faz:**
    - O fluxo é idêntico à macro padrão, mas interage com a interface do YouTube Studio em inglês. Ela baixa os mesmos 22 relatórios, cujos nomes de arquivo estarão em inglês (ex: `Date`, `Traffic source`, `Post`).
- **É necessário ajustar as datas presentes na macro para o período que quer analisar**

---

#### Macro 3: Canais C.A.
- **Arquivo da Macro:** `3 - Extração Overview C.A.json`
- **Para quais canais:** Canais que seguem o fluxo de processamento "CA", conforme definido no `config.json`.
- **O que ela faz:**
    - Esta macro possui um fluxo de download diferente, focado em extrair dados dos últimos 6 meses, além de outros relatórios específicos.
    1.  Acessa o **YouTube Studio -> Estatísticas -> Modo Avançado**.
    2.  Baixa múltiplos relatórios de "Conteúdo" para cada um dos últimos 6 meses.
    3.  Baixa relatórios de "Data" para um período específico.
    4.  Baixa os dados complementares de "Origem de tráfego" e "Postagens na Comunidade".
- **É necessário ajustar as datas presentes na macro para o período que quer analisar**

**Instruções Importantes:**
- Após executar a macro apropriada, mova todos os arquivos `.zip` gerados para a pasta `raw_data` do artista correspondente antes de executar o script `main.py`.

## 3. Estrutura de Diretórios

Para que o projeto funcione corretamente, a seguinte estrutura de pastas deve ser mantida no diretório raiz:

```
/
|-- dados_full/
|   |-- 3Palavrinhas/
|   |   |-- raw_data/
|   |   |   |-- Data 2024-01-01.zip
|   |   |   |-- ...
|   |   |-- (arquivos .csv gerados aqui)
|   |-- 3LittleWords/
|   |   |-- raw_data/
|   |   |-- ...
|   |-- ... (outros artistas)
|
|-- resources/
|   |-- (templates, fontes, imagens usadas nas apresentações)
|
|-- main.py                         # Script principal para executar o fluxo
|-- extracao_unificada.py           # Script para extração padrão (PT, EN, ES)
|-- extraindo_renomeando_CA.py      # Script para extração do tipo CA
|-- tratamento.py                   # Script para tratamento de dados padrão
|-- tratamento_CA.py                # Script para tratamento de dados CA
|-- report.py                       # Script para geração de report padrão
|-- report_CA.py                    # Script para geração de report CA
|-- gerar_graficos.py               # Script para gerar gráficos em português
|-- gerar_graficos_ingles.py        # Script para gerar gráficos em inglês
|-- apresentacao_report.py          # Script para gerar a apresentação no modelo de Report
|-- apresentacao_cluster.py         # Script para gerar a apresentação no modelo de Overview com Cluster
|-- apresentacao_midias.py          # Script para gerar a apresentação no modelo de Overview para Grupos de mídia
|-- apresentacao_report_ingles.py   # Script para gerar a apresentação no modelo de Report em inglês
|-- apresentacao_cluster_ingles.py  # Script para gerar a apresentação no modelo de Overview com Cluster em inglês
|
|-- config.json                 # Configuração base dos artistas
|-- comparacoes.json            # Configuração dos nomes dos artistas na apresentação
|-- presentation_config.json    # Configuração da apresentação a ser gerada
|-- exports.txt                 # Lista de artistas a serem processados
|-- requirements.txt            # Dependências do projeto
```

## 4. Arquivos de Configuração

### `exports.txt`
Lista os artistas a serem processados, um por linha.

### `config.json`
Configuração principal e de longo prazo para cada artista.
- **`tipo_processamento`**: (Opcional) Define se o artista usa o fluxo "CA". Se omitido, usa o fluxo padrão.
- **`grupo_4_graficos`**: Define qual script de geração de gráficos usar.
- **`grupo_5_apresentacao`**: Lista TODOS os scripts de apresentação que o artista pode gerar.

### `presentation_config.json`
Define qual(is) apresentação(ões) serão geradas na execução atual.
- A chave é o nome do artista.
- O valor é uma **lista** de tipos (`["report"]`, `["cluster"]`, `["report", "midias"]`).
- Se um artista não estiver neste arquivo, o padrão `["report"]` será usado.

## 5. Como Usar

### Passo 1: Instalação das Dependências
Execute no terminal, no diretório do projeto:
```bash
pip install -r requirements.txt
```

### Passo 2: Preparação dos Dados (Extração com Macros)
1.  Execute a macro de extração correta para o artista (Padrão, Inglês ou C.A.), conforme descrito na Seção 2.
2.  Mova todos os arquivos `.zip` baixados para a pasta `dados_full/<nome_do_artista>/raw_data/`.
3.  Copie o numero total de inscritos mostrado ao final da execução da macro para o arquivo sub.txt presente na pasta `dados_full/<nome_do_artista>/`.

### Passo 3: Configuração do Processo
1.  **`exports.txt`**: Adicione os nomes dos artistas que você deseja processar.
2.  **`config.json`**: Verifique se os artistas têm as configurações base corretas.
3.  **`comparacoes.json`**: Verifique se os artistas têm as correspondências de nome corretas.
4.  **`presentation_config.json`**: Defina qual(is) apresentação(ões) você quer gerar.

### Passo 4: Execução do Fluxo Automatizado
Para iniciar o processo completo, execute o script `main.py` pelo terminal:

```bash
python main.py
```

O script cuidará de todo o resto, desde a extração dos zips até a geração dos relatórios finais.