import os
import sys
import glob
import openpyxl
import matplotlib as mpl
import matplotlib.font_manager as fm
import warnings
warnings.filterwarnings('ignore')
import seaborn as sns
sns.set_style('whitegrid')
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import matplotlib.patheffects as path_effects
from matplotlib.ticker import MaxNLocator, FuncFormatter
from datetime import datetime,date
from openpyxl import load_workbook



try:
    mpl.rcParams["font.family"] = "DM Sans"
    font_path = "resources/Fonts/DMSans-Bold.ttf"
    dm_sans_bold = fm.FontProperties(fname=font_path)
except Exception as e:
    print(f"⚠️ Aviso: Fonte 'DM Sans' não encontrada. Usando fonte padrão. Erro: {e}")
    mpl.rcParams["font.family"] = "sans-serif"
    dm_sans_bold = fm.FontProperties(weight='bold')


STYLE_CONFIG = {
    'colors': {
        'primary_blue': '#4f46e5', 'secondary_blue': '#c7d2fe', 'primary_purple': '#7c3aed',
        'primary_yellow':'#E2FC51', 'secondary_yellow': '#3157F7', 'accent_blue': '#60a5fa',
        'accent_purple': '#a78bfa', 'text_dark': '#1f2937', 'text_light': '#6b7280',
        'background': '#f9fafb', 'positive': '#10b981', 'negative': '#ef4444',
        'neutral': '#d1d5db', 'vod': '#4f46e5', 'live': '#3157F7', 'shorts': '#E2FC51',
        'primary_red': '#B22222',
    },
    'font_bold': dm_sans_bold,
    'font_props_title': {'fontproperties': dm_sans_bold, 'size': 20},
    'font_props_subtitle': {'fontproperties': dm_sans_bold, 'size': 16},
    'font_props_label': {'size': 12},
    'label_font_props': { # Fonte em negrito
        'fontproperties': dm_sans_bold, 
        'size': 10
        # A chave 'color' foi removida para evitar conflitos
    },
    'label_bbox_props': { # Fundo arredondado
        'boxstyle': 'round,pad=0.3',
        'facecolor': 'white',
        'edgecolor': 'none',
        'alpha': 0.8
    },
    'table_font_props': {
        'fontproperties': dm_sans_bold,
        'size': 14
    },
    'dpi': 300, 'transparent': True,
    'figsize_wide': (12, 7), 'figsize_standard': (10, 6),
    'figsize_table_metricas': (14, 7), 'figsize_table_small': (12, 4)
}


METRIC_ICON_MAP = {
    "Vídeos publicados": "▶",
    "Impressões": "👁",
    "Taxa de cliques de impressões (%)": "🖱",
    "Visualizações": "📊",
    "Duração média da visualização": "⏰",
    "Porcentagem visualizada média (%)": "📈",
    "Inscritos": "👥",
    "RPM (USD)": "$",
    "Receita estimada (USD)": "💰"
}


def custom_format(num): 
    return f"{int(num):,}".replace(',', '.')


def dec_format(num):
    if isinstance(num, (int, float)): 
        return f'{num:,.2f}'.replace(',', 'X').replace('.', ',').replace('X', '.')
    return str(num)


def safe_float_conversion(value):
    try: return float(str(value).replace('.', '').replace(',', '.'))
    except (ValueError, TypeError): return 0.0


def converter_excel_time_para_segundos(time_val):
    """
    Converte um valor de tempo do Excel para segundos.
    Reutiliza a função de minutos e multiplica o resultado por 60.
    """
    minutos_decimais = converter_excel_time_para_minutos(time_val)
    if minutos_decimais is not None:
        return minutos_decimais * 60
    return 0.0


def value_with_arrow(value):
    numeric_value = float(value)
    if numeric_value < 0: return f"{numeric_value:.2f}% ↓", STYLE_CONFIG['colors']['negative']
    return f"{numeric_value:.2f}% ↑", STYLE_CONFIG['colors']['positive']


def format_currency(value): 
    return f'${float(value):,.2f}'.replace(',', 'X').replace('.', ',').replace('X', '.')


def converter_tempo_flexivel_para_minutos(tempo_str):
    if not isinstance(tempo_str, str) or ':' not in tempo_str: return 0.0
    parts = tempo_str.split(':')
    try:
        if len(parts) == 3: return (int(parts[0]) * 60) + int(parts[1]) + (int(parts[2]) / 60.0)
        if len(parts) == 2: return int(parts[0]) + (int(parts[1]) / 60.0)
        return 0.0
    except (ValueError, TypeError): return 0.0


def formatar_eixo_numeros(tick_val, pos):
    if abs(tick_val) >= 1_000_000: return f'{tick_val/1_000_000:.1f}M'
    if abs(tick_val) >= 1_000: return f'{tick_val/1_000:.1f}K'
    return str(int(tick_val))


def formatar_milhoes_mil(num):
    num = float(num)
    if abs(num) >= 1_000_000: return f'{num/1_000_000:.2f} Mi'
    if abs(num) >= 1_000: return f'{num/1_000:.0f} mil'
    return str(int(num))


def converter_excel_time_para_minutos(time_val):
    """
    Converte um valor de tempo do Excel para minutos decimais.
    Funciona com floats (formato de tempo do Excel), strings ('HH:MM:SS')
    e objetos datetime.time.
    """
    # Se for um número (formato comum do Excel para tempo)
    if isinstance(time_val, (int, float)):
        # Excel armazena tempo como fração de 1 dia (24 horas * 60 minutos)
        return float(time_val) * 24 * 60
    
    # Se for um texto como 'HH:MM:SS' ou 'MM:SS'
    if isinstance(time_val, str) and ':' in time_val:
        parts = time_val.split(':')
        try:
            if len(parts) == 3: return (int(parts[0]) * 60) + int(parts[1]) + (int(parts[2]) / 60.0)
            if len(parts) == 2: return int(parts[0]) + (int(parts[1]) / 60.0)
        except (ValueError, TypeError):
            return 0.0 # Retorna 0 se o texto não for um formato válido
            
    # Se for um objeto de tempo do Python
    if hasattr(time_val, 'hour'):
        return time_val.hour * 60 + time_val.minute + time_val.second / 60.0

    # Se não for nenhum dos formatos esperados, retorna 0
    return 0.0


def extract_numeric_value(value):
    """Extrai um valor numérico de uma string formatada (K, M, %, :, $)."""
    if isinstance(value, (int, float)):
        return float(value)
    if not isinstance(value, str):
        return 0.0
    
    cleaned_str = value.replace('$', '').replace('R$', '').replace('%', '').replace('+', '').replace(',', '.').strip()
    
    if ':' in cleaned_str:
        try:
            parts = cleaned_str.split(':');
            return int(parts[0]) * 60 + int(parts[1]) if len(parts) >= 2 else 0
        except: return 0.0
    
    if 'M' in cleaned_str.upper():
        return float(cleaned_str.upper().replace('M', '')) * 1000000
    if 'K' in cleaned_str.upper():
        return float(cleaned_str.upper().replace('K', '')) * 1000
        
    try: return float(cleaned_str)
    except ValueError: return 0.0


def get_trend(current_val, prev_val):
    """Determina a tendência (up, down, neutral) comparando dois valores numéricos."""
    if not all(isinstance(v, (int, float)) for v in [current_val, prev_val]):
        return "neutral"
    if current_val > prev_val:
        return "up"
    if current_val < prev_val:
        return "down"
    return "neutral"


def get_performance_color(value, row_data):
    """
    Calcula a cor de fundo da célula com base na performance relativa da métrica em sua própria linha.
    Esta é a versão Python da sua função getPerformanceColor.
    """
    # Filtra apenas valores numéricos para não distorcer a escala
    numeric_values = [v for v in row_data if isinstance(v, (int, float)) and v > 0]
    if not numeric_values:
        return {'facecolor': '#f9fafb', 'textcolor': '#1f2937'} # Cor neutra

    max_val, min_val = max(numeric_values), min(numeric_values)
    range_val = max_val - min_val

    if range_val == 0 or not isinstance(value, (int, float)):
        return {'facecolor': '#f9fafb', 'textcolor': '#1f2937'} # Cor neutra

    # Normaliza o valor entre 0 e 1
    normalized = (value - min_val) / range_val
    
    # Retorna cores de fundo e de texto com base na performance
    if normalized >= 0.75:
        return {'facecolor': '#d1fae5', 'textcolor': '#065f46'} # Verde
    if normalized >= 0.35:
        return {'facecolor': '#dbeafe', 'textcolor': '#1e40af'} # Azul
    return {'facecolor': '#fee2e2', 'textcolor': '#991b1b'}     # Vermelho


def format_for_table(value):
    """Formata os números para exibição na tabela (K para mil, M para milhões)."""
    if not isinstance(value, (int, float)):
        return str(value)
    if abs(value) >= 1_000_000:
        return f'{value/1_000_000:.1f}M'
    if abs(value) >= 1_000:
        return f'{value/1_000:.1f}K'
    if isinstance(value, float) and value < 100: # Para porcentagens e valores pequenos
        return f'{value:.2f}'
    return str(int(value))


def gerar_tabela_metricas_avancada(artista, tipo_conteudo, nome_arquivo, plot_index):
    """
    Versão final e corrigida que gera a tabela de métricas com controle de fonte funcional.
    """
    try:
        # --- 1. CARREGAMENTO E PREPARAÇÃO DOS DADOS (sem alterações) ---
        data = pd.read_csv(f'dados_full/{artista}/{nome_arquivo}'); data.drop(0, inplace=True)
        def time_to_seconds(time_str):
            if isinstance(time_str, str) and ':' in time_str:
                parts = time_str.split(':')
                try:
                    if len(parts) == 3: return int(parts[0]) * 3600 + int(parts[1]) * 60 + int(parts[2])
                    elif len(parts) == 2: return int(parts[0]) * 60 + int(parts[1])
                except (ValueError, IndexError): return 0
            return 0
        data['Duração média da visualização'] = data['Duração média da visualização'].apply(time_to_seconds)
        data["Data"] = pd.to_datetime(data['Data']); data.sort_values('Data', ascending=True, inplace=True, ignore_index=True)
        if len(data) >= 7: data = data.tail(7)
        data.reset_index(drop=True, inplace=True); data["Mes"] = data.Data.apply(lambda i: i.strftime("%b"))
        ordem_colunas = ['Mes', 'Vídeos publicados', 'Impressões', 'Taxa de cliques de impressões (%)', 'Visualizações', 'Duração média da visualização', 'Porcentagem visualizada média (%)', 'Inscritos', 'RPM (USD)', 'Receita estimada (USD)']; df_full = data[ordem_colunas].copy().set_index('Mes').T; df_display = df_full.iloc[:, 1:]

        # --- 2. CRIAÇÃO DA TABELA ---
        fig, ax = plt.subplots(figsize=(16, 7)); ax.axis('off')
        
        tabela = ax.table(
            cellText=[[''] * len(df_display.columns)] * len(df_display.index),
            rowLabels=df_display.index, colLabels=df_display.columns,
            loc='center', cellLoc='right'
        )
        tabela.auto_set_font_size(False) # Impede o ajuste automático
        tabela.scale(1, 3) # Ajusta a altura das células

        # --- 3. ESTILIZAÇÃO E PREENCHIMENTO (LÓGICA CORRIGIDA) ---
        for (row, col_display), cell in tabela.get_celld().items():
            col_full = col_display + 1; cell.set_edgecolor('none')
            
            # Pega as propriedades de fonte do STYLE_CONFIG
            font_properties = STYLE_CONFIG['table_font_props']
            
            if row == 0: # Cabeçalho
                cell.set_text_props(**font_properties, color='white', ha='center')
                cell.set_facecolor('#3157F7'); cell.set_height(0.12)
            elif col_display == -1: # Coluna de Métricas
                cell.set_text_props(**font_properties, ha='left', va='center', color='#142544')
                cell.get_text().set_text(f"  {df_display.index[row-1]}")
                cell.set_facecolor('#F0F4FF'); cell.set_edgecolor('#E5E7EB'); cell.set_linewidth(1); cell.set_width(0.35)
            else: # Células de Dados
                cell.set_facecolor('#FFFFFF' if row % 2 != 0 else '#F0F4FF')
                metric_name = df_full.index[row-1]; current_value = df_full.iloc[row-1, col_full]; prev_value = df_full.iloc[row-1, col_full - 1]
                trend = get_trend(current_value, prev_value)
                
                display_val = ""
                if pd.notna(current_value):
                    if metric_name in ["Impressões", "Visualizações", "Inscritos", "Vídeos publicados"]: display_val = custom_format(current_value)
                    elif metric_name == "Duração média da visualização": minutes, seconds = divmod(int(current_value), 60); display_val = f"{minutes}:{seconds:02d}"
                    elif metric_name == "Receita estimada (USD)": display_val = format_currency(current_value)
                    else: display_val = dec_format(current_value)
                    if metric_name in ["Taxa de cliques de impressões (%)", "Porcentagem visualizada média (%)"]: display_val += "%"
                
                icon = ''; color = STYLE_CONFIG['colors']['text_dark']
                if trend == "up": icon = '↑ '; color = STYLE_CONFIG['colors']['positive']
                elif trend == "down": icon = '↓ '; color = STYLE_CONFIG['colors']['negative']
                
                cell.get_text().set_text(f"{icon}{display_val}")
                # APLICA AS PROPRIEDADES CORRETAMENTE
                cell.get_text().set_font_properties(font_properties['fontproperties'])
                cell.get_text().set_size(font_properties['size'])
                cell.get_text().set_color(color)

        fig.suptitle("", y=0.5, **STYLE_CONFIG['font_props_title'])
        plt.tight_layout(rect=[0, 0, 1, 0.95])
        
        plt.savefig(
            f'dados_full/{artista}/plots/{plot_index} - Métricas_{tipo_conteudo}_Avancada.png', 
            dpi=STYLE_CONFIG['dpi'], bbox_inches='tight', pad_inches=0.1, transparent=STYLE_CONFIG['transparent']
        )
        plt.close(fig)
    except Exception as e:
        print(f"❌ Erro em gerar_tabela_metricas_avancada para '{artista}' ({tipo_conteudo}): {e}")
                

def formatar_numero_card(value):
    """Formata números grandes para o padrão do novo card (mil, mi)."""
    if not isinstance(value, (int, float)):
        return str(value)
    if abs(value) >= 1_000_000:
        return f'{value / 1_000_000:.1f} mi'
    if abs(value) >= 1_000:
        return f'{value / 1_000:.0f} mil'
    return str(int(value))


def criar_big_number_card(config, output_filename):
    """
    Cria uma imagem de um card de métrica, agora com posicionamento de valor personalizável.
    """
    BRAND_COLORS = {
        'azulEscuro': "#142544", 'azulPrimario': "#3157F7", 'roxo': "#8140FA",
        'amareloMarcaTexto': "#E2FC51", 'branco': "#FFFFFF", 'cinzaClaro': "#E8E8E9",
        'positive_text': '#16a34a', 'positive_bg': '#f0fdf4',
        'negative_text': '#dc2626', 'negative_bg': '#fef2f2',
    }
    variant = config.get('variant', 'default')
    if variant == 'primary': accent_color, title_color = BRAND_COLORS['azulPrimario'], BRAND_COLORS['azulPrimario']
    elif variant == 'secondary': accent_color, title_color = BRAND_COLORS['roxo'], BRAND_COLORS['roxo']
    elif variant == 'accent': accent_color, title_color = BRAND_COLORS['amareloMarcaTexto'], BRAND_COLORS['azulEscuro']
    else: accent_color, title_color = BRAND_COLORS['azulEscuro'], BRAND_COLORS['azulEscuro']

    fig, ax = plt.subplots(figsize=(5, 5)); fig.set_facecolor('#f0f2f5'); ax.axis('off')
    ax.add_patch(mpatches.FancyBboxPatch((0.05, 0.05), 0.9, 0.9, facecolor='#d1d5db', edgecolor='none', boxstyle="round,pad=0,rounding_size=0.04", transform=ax.transAxes, alpha=0.5))
    ax.add_patch(mpatches.FancyBboxPatch((0.05, 0.05), 0.9, 0.9, facecolor=BRAND_COLORS['branco'], edgecolor='#e5e7eb', boxstyle="round,pad=0,rounding_size=0.04", transform=ax.transAxes))
    ax.add_patch(mpatches.Rectangle((0.05, 0.91), 0.9, 0.04, facecolor=accent_color, transform=ax.transAxes, clip_on=False, zorder=2))
    ax.text(0.5, 0.85, config['title'], ha='center', va='center', fontsize=18, fontproperties=STYLE_CONFIG['font_bold'], color=title_color, transform=ax.transAxes)
    ax.text(0.5, 0.70, config['mainValue'], ha='center', va='center', fontsize=40, fontproperties=STYLE_CONFIG['font_bold'], color=BRAND_COLORS['azulEscuro'], transform=ax.transAxes)
    
    y_pos = 0.48
    for item in config['breakdown']:
        ax.text(0.15, y_pos, f"{item['label']}:", ha='left', va='center', fontsize=14, color=BRAND_COLORS['azulEscuro'], transform=ax.transAxes)
        
        # --- ALTERAÇÃO APLICADA AQUI ---
        # Pega a posição personalizada se ela existir, senão usa o padrão 0.35
        valor_x_pos = item.get('value_x_pos', 0.35)
        ax.text(valor_x_pos, y_pos, item['value'], ha='left', va='center', fontsize=14, fontproperties=STYLE_CONFIG['font_bold'], color=BRAND_COLORS['azulPrimario'], transform=ax.transAxes)
        
        if 'change' in item:
            is_positive = item['changeType'] == 'positive'
            badge_color_text = BRAND_COLORS['positive_text'] if is_positive else BRAND_COLORS['negative_text']
            badge_color_bg = BRAND_COLORS['positive_bg'] if is_positive else BRAND_COLORS['negative_bg']
            icon = '↑' if is_positive else '↓'
            badge_text = f"{icon} {item['change']}"
            text_box = ax.text(0.85, y_pos, badge_text, ha='center', va='center', fontsize=10, color=badge_color_text, fontproperties=STYLE_CONFIG['font_bold'], transform=ax.transAxes)
            text_box.set_bbox(dict(facecolor=badge_color_bg, edgecolor=badge_color_text, boxstyle='round,pad=0.4', alpha=0.7, linewidth=0.5))
        
        y_pos -= 0.12

    plt.savefig(output_filename, dpi=300, bbox_inches='tight', transparent = True); plt.close(fig)


def gerar_cards_detalhados(artista, file_path):
    """
    Função principal que lê os dados e adiciona uma configuração de posição
    personalizada para o card de Inscritos.
    """
    try:
        data = pd.read_excel(file_path, sheet_name='Resultado')
        data2 = pd.read_excel(file_path, sheet_name='Desvio')
        
        metricas_config = [
            {'title': 'VISUALIZAÇÕES', 'variant': 'primary', 'total': (data.iloc[8, 7] + data.iloc[9, 7] + data.iloc[10, 7]), 'formatter': formatar_numero_card, 'plot_index': '4a', 
             'breakdown': [
                 {'label': 'VOD', 'value': data.iloc[9, 7], 'deviation': data2.iloc[9, 6], 'value_x_pos': 0.28}, 
                 {'label': 'LIVE', 'value': data.iloc[10, 7], 'deviation': data2.iloc[10, 6], 'value_x_pos': 0.28}, 
                 {'label': 'SHORT', 'value': data.iloc[8, 7], 'deviation': data2.iloc[8, 6], 'value_x_pos': 0.34}
             ]},

            {'title': 'RECEITA (USD)', 'variant': 'secondary', 'total': (data.iloc[1, 7] + data.iloc[2, 7] + data.iloc[3, 7]), 'formatter': format_currency, 'plot_index': '4b', 
             'breakdown': [
                 {'label': 'VOD', 'value': data.iloc[1, 7], 'deviation': data2.iloc[1, 6], 'value_x_pos': 0.29}, 
                 {'label': 'LIVE', 'value': data.iloc[2, 7], 'deviation': data2.iloc[2, 6], 'value_x_pos': 0.29}, 
                 {'label': 'SHORT', 'value': data.iloc[3, 7], 'deviation': data2.iloc[3, 6], 'value_x_pos': 0.34}
             ]},
            
            # --- ALTERAÇÃO APLICADA AQUI ---
            {'title': 'INSCRITOS', 'variant': 'accent', 'total': data.iloc[50, 7], 'formatter': custom_format, 'plot_index': '4c', 
             'breakdown': [
                 {'label': 'Ganhos', 'value': data.iloc[73, 7], 'deviation': data2.iloc[73, 6]},
                 # Adicionada a chave 'value_x_pos' para ajustar apenas este item
                 {'label': 'Perdidos', 'value': data.iloc[74, 7], 'deviation': data2.iloc[74, 6], 'value_x_pos': 0.38}, 
             ]},
            
            {'title': 'RPM (USD)', 'variant': 'primary', 'total': data.iloc[72, 7], 'formatter': lambda x: f"${dec_format(x)}", 'plot_index': '4d', 
             'breakdown': [
                 {'label': 'VOD', 'value': data.iloc[69, 7], 'deviation': data2.iloc[69, 6], 'value_x_pos': 0.29}, 
                 {'label': 'LIVE', 'value': data.iloc[70, 7], 'deviation': data2.iloc[70, 6], 'value_x_pos': 0.29}, 
                 {'label': 'SHORT', 'value': data.iloc[71, 7], 'deviation': data2.iloc[71, 6], 'value_x_pos': 0.34}
             ]},

            {'title': 'IMPRESSÕES', 'variant': 'secondary', 'total': (data.iloc[5, 7] + data.iloc[6, 7] + data.iloc[65, 7]), 'formatter': formatar_numero_card, 'plot_index': '4e', 
             'breakdown': [
                 {'label': 'VOD', 'value': data.iloc[5, 7], 'deviation': data2.iloc[5, 6], 'value_x_pos': 0.28}, 
                 {'label': 'LIVE', 'value': data.iloc[6, 7], 'deviation': data2.iloc[6, 6], 'value_x_pos': 0.28}, 
                 {'label': 'SHORT', 'value': data.iloc[65, 7], 'deviation': data2.iloc[65, 6], 'value_x_pos': 0.34}
             ]},

            {'title': 'WATCHTIME (HORAS)', 'variant': 'accent', 'total': data.iloc[14, 7], 'formatter': lambda x: f"{float(x):,.0f} h".replace(",", "."), 'plot_index': '4f', 
             'breakdown': [
                 {'label': 'VOD', 'value': data.iloc[66, 7], 'deviation': data2.iloc[66, 6], 'value_x_pos': 0.28}, 
                 {'label': 'LIVE', 'value': data.iloc[67, 7], 'deviation': data2.iloc[67, 6], 'value_x_pos': 0.28}, 
                 {'label': 'SHORT', 'value': data.iloc[68, 7], 'deviation': data2.iloc[68, 6], 'value_x_pos': 0.34}
             ]}
        ]
        
        for config in metricas_config:
            card_data = {
                'title': config['title'], 'mainValue': config['formatter'](config['total']),
                'variant': config.get('variant', 'default'), 'breakdown': []
            }
            for item in config['breakdown']:
                dev_percent = item['deviation'] * 100
                change_sign = '+' if dev_percent >= 0 else ''
                formatted_change = f"{change_sign}{dev_percent:.0f}%"
                if item['label'] == 'Perdidos':
                    change_type = 'negative' if dev_percent >= 0 else 'positive'
                else:
                    change_type = 'positive' if dev_percent >= 0 else 'negative'
                
                # Passa a configuração de posição para a função de desenho
                item_data = {
                    'label': item['label'], 'value': config['formatter'](item['value']),
                    'change': formatted_change, 'changeType': change_type
                }
                if 'value_x_pos' in item:
                    item_data['value_x_pos'] = item['value_x_pos']
                
                card_data['breakdown'].append(item_data)
            
            safe_title = config['title'].replace(' (USD)', '').replace(' ', '_')
            output_path = f"dados_full/{artista}/plots/{config['plot_index']} - Card_{safe_title}_v2.png"
            criar_big_number_card(card_data, output_path)
            
    except Exception as e: 
        print(f"❌ Erro em gerar_cards_detalhados para '{artista}': {e}")
        

def publicated_table(artista, file_path):
    """
    Gera a tabela de vídeos publicados com design avançado, mas sem os indicadores de tendência (setas).
    """
    try:
        # --- 1. CARREGAMENTO E PREPARAÇÃO DOS DADOS ---
        df_full_data = pd.read_excel(file_path.format(artista=artista), sheet_name='Resultado', index_col=0)
        
        # Seleciona apenas as linhas de Vídeos e Lives
        df = df_full_data.iloc[11:13].copy()
        
        # Remove a coluna 'Média' e transpõe para ter meses como colunas
        if 'Média' in df.columns:
            df.drop('Média', axis=1, inplace=True)
        
        df.columns = pd.to_datetime(df.columns).strftime('%b')

        # --- 2. CRIAÇÃO DA TABELA VAZIA ---
        fig, ax = plt.subplots(figsize=(10, 2.5)) # Tamanho otimizado para 2 linhas
        ax.axis('off')
        
        tabela = ax.table(
            cellText=[[''] * len(df.columns)] * len(df.index), # Células vazias
            rowLabels=df.index,
            colLabels=df.columns,
            loc='center',
            cellLoc='right' # Alinha o texto à direita
        )
        tabela.auto_set_font_size(False)
        tabela.scale(1, 2.5) # Ajusta a altura das células

        # --- 3. ESTILIZAÇÃO E PREENCHIMENTO ---
        for (row, col), cell in tabela.get_celld().items():
            cell.set_edgecolor('none')

            # Estilo do Cabeçalho
            if row == 0:
                cell.set_text_props(ha='center', color='white', **STYLE_CONFIG['table_font_props'])
                cell.set_facecolor(STYLE_CONFIG['colors']['primary_blue'])
                cell.set_height(0.2)
            # Estilo da Coluna de Métricas
            elif col == -1:
                cell.set_text_props(ha='left', va='center', color=STYLE_CONFIG['colors']['text_dark'], **STYLE_CONFIG['table_font_props'])
                cell.get_text().set_text(f"  {df.index[row-1]}")
                cell.set_facecolor('#F0F4FF')
                cell.set_edgecolor('#E5E7EB')
                cell.set_linewidth(1)
                cell.set_width(0.4)
            # Estilo das Células de Dados
            else:
                cell.set_facecolor('#FFFFFF' if row % 2 != 0 else '#F0F4FF')
                
                current_value = df.iloc[row-1, col]
                display_val = f"{int(current_value)}" if pd.notna(current_value) else '-'
                
                # ALTERAÇÃO: A lógica de ícone e tendência foi removida daqui.
                # O texto é definido diretamente, sem as setas.
                cell.get_text().set_text(display_val)
                cell.get_text().set_color(STYLE_CONFIG['colors']['text_dark']) # Cor única para todos os dados
                cell.get_text().set_font_properties(STYLE_CONFIG['table_font_props']['fontproperties'])
                cell.get_text().set_size(STYLE_CONFIG['table_font_props']['size'])

        fig.suptitle('', y=0.5, **STYLE_CONFIG['font_props_title'])
        plt.tight_layout(rect=[0, 0, 1, 1])
        
        plt.savefig(
            f'dados_full/{artista}/plots/5 - Publicados.png', 
            dpi=STYLE_CONFIG['dpi'], 
            bbox_inches='tight',
            pad_inches=0.1,
            transparent=STYLE_CONFIG['transparent']
        )
        plt.close(fig)
        
    except Exception as e:
        print(f"❌ Erro em publicated_table para '{artista}': {e}")


def analyze_initial_updated(artista, file_path):
    try:
        df = pd.read_excel(file_path, sheet_name='Resultado'); 
        meses_com_media = pd.to_datetime(df.columns[1:], errors='coerce'); 
        visualizacoes = df.iloc[7, 1:].values; 
        receita_total = df.iloc[0, 1:].values
        df_anomalias = pd.DataFrame({'Meses': meses_com_media.values, 'Visualizacoes': visualizacoes, 'ReceitaTotal': receita_total}).dropna(subset=['Meses'])
        df_anomalias['Visualizacoes'] = pd.to_numeric(df_anomalias['Visualizacoes']); 
        df_anomalias['ReceitaTotal'] = pd.to_numeric(df_anomalias['ReceitaTotal'])

        fig, ax1 = plt.subplots(figsize=STYLE_CONFIG['figsize_standard']); 
        color_receita = STYLE_CONFIG['colors']['primary_purple']; 
        color_views = STYLE_CONFIG['colors']['secondary_yellow']

        ax1.set_xlabel('Mês', **STYLE_CONFIG['font_props_label']); 
        ax1.set_ylabel('Receita Total ($)', color=color_receita, **STYLE_CONFIG['font_props_label'])

        linha_receita, = ax1.plot(df_anomalias['Meses'], df_anomalias['ReceitaTotal'], color=color_receita, marker='o', label='Receita Total');

        ax1.tick_params(axis='y', labelcolor=color_receita); 
        ax1.set_ylim(bottom=0); ax2 = ax1.twinx(); 
        ax2.set_ylabel('Visualizações', color=color_views, **STYLE_CONFIG['font_props_label'])

        linha_visualizacao, = ax2.plot(df_anomalias['Meses'], df_anomalias['Visualizacoes'], color=color_views, marker='o', label='Visualizações');

        ax2.tick_params(axis='y', labelcolor=color_views); 
        ax2.set_ylim(bottom=0); ax2.yaxis.set_major_formatter(FuncFormatter(formatar_eixo_numeros))
        ax1.legend(handles=[linha_receita], loc='upper left'); 
        ax2.legend(handles=[linha_visualizacao], loc='upper right')

        plt.title('Receita x Visualização', **STYLE_CONFIG['font_props_title']); plt.tight_layout(); 
        plt.savefig(f'dados_full/{artista}/plots/6 - Análise_Inicial.png', dpi=STYLE_CONFIG['dpi'], transparent=STYLE_CONFIG['transparent'], bbox_inches='tight'); 
        plt.close(fig)

    except Exception as e: print(f"❌ Erro em analyze_initial_updated para '{artista}': {e}")


def watch_table(artista, file_path):
    """
    Gera a tabela de WatchTime, CPM e Taxa de Preenchimento com o design unificado
    e a formatação de dados corrigida.
    """
    try:
        # --- 1. CARREGAMENTO E PREPARAÇÃO DOS DADOS ---
        df = pd.read_excel(file_path.format(artista=artista), sheet_name="Resultado")
        
        # Define as linhas a serem selecionadas
        ROW_TAXA_PREENCHIMENTO = 62 
        linhas_selecionadas = [14, 57, ROW_TAXA_PREENCHIMENTO]
        df_watch = df.iloc[linhas_selecionadas].copy()
        
        df_watch.set_index(df.columns[0], inplace=True)
        
        if 'Média' in df_watch.columns:
            df_watch = df_watch.drop(columns='Média')
        
        df_watch.columns = pd.to_datetime(df_watch.columns, errors='coerce').strftime('%b')

        # --- 2. CRIAÇÃO DA TABELA VAZIA ---
        fig, ax = plt.subplots(figsize=(10, 3)) 
        ax.axis('off')
        
        tabela = ax.table(
            cellText=[[''] * len(df_watch.columns)] * len(df_watch.index),
            rowLabels=df_watch.index,
            colLabels=df_watch.columns,
            loc='center',
            cellLoc='right'
        )
        tabela.auto_set_font_size(False)
        tabela.scale(1, 2.8)

        # --- 3. ESTILIZAÇÃO E PREENCHIMENTO CÉLULA POR CÉLULA ---
        for (row, col), cell in tabela.get_celld().items():
            cell.set_edgecolor('none')

            # Estilo do Cabeçalho
            if row == 0:
                cell.set_text_props(ha='center', color='white', **STYLE_CONFIG['table_font_props'])
                cell.set_facecolor(STYLE_CONFIG['colors']['primary_blue'])
                cell.set_height(0.18)
            # Estilo da Coluna de Métricas
            elif col == -1:
                cell.set_text_props(ha='left', va='center', color=STYLE_CONFIG['colors']['text_dark'], **STYLE_CONFIG['table_font_props'])
                cell.get_text().set_text(f"  {df_watch.index[row-1]}")
                cell.set_facecolor('#F0F4FF')
                cell.set_edgecolor('#E5E7EB')
                cell.set_linewidth(1)
                cell.set_width(0.4)
            # Estilo das Células de Dados
            else:
                cell.set_facecolor('#FFFFFF' if row % 2 != 0 else '#F0F4FF')
                
                metric_name = df_watch.index[row-1]
                current_value = df_watch.iloc[row-1, col]
                
                # Formatação contextual do texto da célula
                display_val = ""
                if pd.notna(current_value):
                    # CORREÇÃO APLICADA AQUI:
                    # Verifica o nome da métrica para aplicar a formatação correta.
                    if "Watch Time Total" in metric_name: # Corresponde a "WatchTime"
                        display_val = f"{float(current_value):,.0f} h".replace(",", ".")
                    elif "CPM" in metric_name:
                        display_val = f"${dec_format(current_value)}"
                    else: # A terceira métrica (Taxa de Preenchimento) será formatada como porcentagem
                        display_val = f"{dec_format(current_value)}%"
                else:
                    display_val = '-'
                
                # Define o texto com cor e fonte padronizadas
                cell.get_text().set_text(display_val)
                cell.get_text().set_color(STYLE_CONFIG['colors']['text_dark'])
                cell.get_text().set_font_properties(STYLE_CONFIG['table_font_props']['fontproperties'])
                cell.get_text().set_size(STYLE_CONFIG['table_font_props']['size'])

        fig.suptitle("", y=0.5, **STYLE_CONFIG['font_props_title'])
        plt.tight_layout(rect=[0, 0, 1, 1])
        
        plt.savefig(
            f"dados_full/{artista}/plots/7 - Watchtime.png", 
            transparent=STYLE_CONFIG['transparent'], 
            bbox_inches="tight", 
            dpi=STYLE_CONFIG['dpi']
        )
        plt.close(fig)

    except Exception as e:
        print(f"❌ Erro em watch_table para '{artista}': {e}")


def monetization_graph(artista, file_path):
    """
    Gera o gráfico de monetização com posicionamento vertical garantido para os rótulos,
    evitando sobreposição.
    """
    try:
        # --- 1. CARREGAMENTO DOS DADOS (sem alterações) ---
        df = pd.read_excel(file_path, sheet_name="Resultado")
        meses = pd.to_datetime(df.columns[2:], errors='coerce').dropna().to_series().dt.strftime('%b')
        receita_vod_velho = df.iloc[58, 2:].astype(float)
        receita_vod_novo = df.iloc[55, 2:].astype(float)
        receita_lives_velho = df.iloc[60, 2:].astype(float)
        receita_lives_novo = df.iloc[56, 2:].astype(float)
        receita_total = df.iloc[0, 2:].astype(float)

        # --- 2. CONFIGURAÇÃO DO GRÁFICO (sem alterações) ---
        fig, ax1 = plt.subplots(figsize=STYLE_CONFIG['figsize_standard'])
        index = np.arange(len(meses))
        bar_width = 0.35
        chart_colors = {
            'vod_novo': '#6B7FFF', 'vod_velho': '#3157F7',
            'lives_novo': '#FCA5A5', 'lives_velho': '#EF4444',
            'total_line': '#22C55E'
        }

        # --- 3. DESENHO DAS BARRAS (sem alterações) ---
        rects_vod_velho = ax1.bar(index - bar_width/2, receita_vod_velho, bar_width, label='Receita VOD Velho', color=chart_colors['vod_velho'])
        rects_vod_novo = ax1.bar(index - bar_width/2, receita_vod_novo, bar_width, bottom=receita_vod_velho, label='Receita VOD Novo', color=chart_colors['vod_novo'])
        rects_lives_velho = ax1.bar(index + bar_width/2, receita_lives_velho, bar_width, label='Receita Live Velho', color=chart_colors['lives_velho'])
        rects_lives_novo = ax1.bar(index + bar_width/2, receita_lives_novo, bar_width, bottom=receita_lives_velho, label='Receita Live Novo', color=chart_colors['lives_novo'])

        # --- 4. LÓGICA DE RÓTULOS COM POSICIONAMENTO VERTICAL ---
        
        # Função auxiliar para adicionar os rótulos de forma hierárquica
        def add_stacked_labels(bottom_rects, top_rects):
            for i, (rect_bottom, rect_top) in enumerate(zip(bottom_rects, top_rects)):
                height_bottom = rect_bottom.get_height()
                height_top = rect_top.get_height()

                # --- RÓTULO DA BARRA DE BAIXO ---
                if height_bottom > 0:
                    # Posiciona o rótulo na metade inferior da barra de baixo
                    y_pos_bottom = rect_bottom.get_y() + height_bottom * 0.01
                    ax1.text(rect_bottom.get_x() + rect_bottom.get_width() / 2., y_pos_bottom,
                             f'${height_bottom:,.0f}', ha='center', va='center', color='white',
                             bbox=dict(boxstyle='round,pad=0.3', lw=0, facecolor=rect_bottom.get_facecolor(), alpha=0.9),
                             **STYLE_CONFIG['label_font_props'])

                # --- RÓTULO DA BARRA DE CIMA ---
                if height_top > 0:
                    # Posiciona o rótulo na metade superior da barra de cima
                    y_pos_top = rect_top.get_y() + height_top * 1.01
                    text_color = 'white' if rect_top.get_facecolor() != chart_colors['lives_novo'] else 'black'
                    ax1.text(rect_top.get_x() + rect_top.get_width() / 2., y_pos_top,
                             f'${height_top:,.0f}', ha='center', va='center', color=text_color,
                             bbox=dict(boxstyle='round,pad=0.3', lw=0, facecolor=rect_top.get_facecolor(), alpha=0.9),
                             **STYLE_CONFIG['label_font_props'])

        # Aplica a lógica para os dois grupos de barras
        add_stacked_labels(rects_vod_velho, rects_vod_novo)
        add_stacked_labels(rects_lives_velho, rects_lives_novo)

        # --- 5. CONFIGURAÇÃO FINAL E SALVAMENTO (sem alterações) ---
        ax1.set_xlabel('Meses', **STYLE_CONFIG['font_props_label']); ax1.set_ylabel('Receita por Categoria (USD)', **STYLE_CONFIG['font_props_label']); ax1.set_title('Receita de VODs e Lives (Novo vs Velho)', **STYLE_CONFIG['font_props_title']); ax1.set_xticks(index); ax1.set_xticklabels(meses); ax1.yaxis.set_major_formatter(FuncFormatter(formatar_eixo_numeros))
        ax2 = ax1.twinx(); ax2.plot(meses, receita_total, marker='o', color=chart_colors['total_line'], label='Receita Total (USD)', linewidth=2.5)
        ax2.set_ylabel('Receita Total (USD)', color=chart_colors['total_line'], **STYLE_CONFIG['font_props_label']); ax2.tick_params(axis='y', labelcolor=chart_colors['total_line']); ax2.set_ylim(bottom=0, top=receita_total.max() * 1.2); ax2.yaxis.set_major_formatter(FuncFormatter(formatar_eixo_numeros))
        
        lines, labels = ax1.get_legend_handles_labels(); lines2, labels2 = ax2.get_legend_handles_labels()
        fig.legend(handles=lines + lines2, labels=labels + labels2, loc='upper center', bbox_to_anchor=(0.5, 0.03), ncol=len(labels + labels2), prop={'size': 9})
        
        fig.tight_layout(rect=[0, 0, 1, 0.95])
        plt.savefig(f'dados_full/{artista}/plots/8 - Monetizacao_v2.png', dpi=STYLE_CONFIG['dpi'], transparent=STYLE_CONFIG['transparent'], bbox_inches='tight')
        plt.close(fig)

    except Exception as e:
        print(f"❌ Erro em monetization_graph para '{artista}': {e}")


def revenue_per_type_chart(artista, file_path):
    """
    Gera o gráfico de RPM e Receita, agora com bordas cinzas nas linhas de receita.
    """
    try:
        df = pd.read_excel(file_path, sheet_name="Resultado")
        meses = pd.to_datetime(df.columns[2:], errors='coerce').dropna().to_series().dt.strftime('%b')
        rpm_vod_novo = df.iloc[21, 2:].astype(float)
        rpm_live_novo = df.iloc[22, 2:].astype(float)
        rpm_shorts_novo = df.iloc[23, 2:].astype(float)
        receita_vod = df.iloc[1, 2:].astype(float)
        receita_live = df.iloc[2, 2:].astype(float)
        receita_shorts = df.iloc[3, 2:].astype(float)

        fig, ax1 = plt.subplots(figsize=STYLE_CONFIG['figsize_standard'])
        width = 0.25
        x = np.arange(len(meses))

        # Desenha as barras de RPM
        bars1 = ax1.bar(x - width, rpm_vod_novo, width, label='RPM VOD', color=STYLE_CONFIG['colors']['vod'])
        bars2 = ax1.bar(x, rpm_live_novo, width, label='RPM Live', color='#22C55E')
        bars3 = ax1.bar(x + width, rpm_shorts_novo, width, label='RPM Shorts', color=STYLE_CONFIG['colors']['shorts'])

        # Adiciona os rótulos às barras
        for bars in [bars1, bars2, bars3]:
            for bar in bars:
                ax1.text(
                    bar.get_x() + bar.get_width() / 2,
                    bar.get_height(),
                    f'${bar.get_height():.2f}',
                    ha='center',
                    va='bottom',
                    bbox=STYLE_CONFIG['label_bbox_props'],
                    # CORREÇÃO: Usando a chave correta do STYLE_CONFIG
                    **STYLE_CONFIG['label_font_props']
                )

        ax1.set_title('RPM Novo e Receita por Tipo de Conteúdo', **STYLE_CONFIG['font_props_title'])
        ax1.set_ylabel('RPM Novo (USD)', **STYLE_CONFIG['font_props_label'])
        ax1.set_xlabel('Meses', **STYLE_CONFIG['font_props_label'])
        ax1.set_xticks(x)
        ax1.set_xticklabels(meses)

        # Eixo secundário para a receita
        ax2 = ax1.twinx()
        
        # --- ALTERAÇÃO APLICADA AQUI ---
        # Define o efeito de borda cinza para as linhas
        line_effect = [path_effects.withStroke(linewidth=3, foreground='#808080')]

        # Aplica o efeito de borda em cada linha
        ax2.plot(x, receita_vod, color=STYLE_CONFIG['colors']['vod'], marker='o', label='Receita VOD', path_effects=line_effect)
        ax2.plot(x, receita_live, color='#22C55E', marker='o', label='Receita Live', path_effects=line_effect)
        ax2.plot(x, receita_shorts, color=STYLE_CONFIG['colors']['shorts'], marker='o', label='Receita Shorts', path_effects=line_effect)
        
        ax2.set_ylabel('Receita (USD)', **STYLE_CONFIG['font_props_label'])
        ax2.set_ylim(bottom=0)
        ax2.yaxis.set_major_formatter(FuncFormatter(formatar_eixo_numeros))
        
        # Legenda e salvamento
        lines, labels = ax1.get_legend_handles_labels()
        lines2, labels2 = ax2.get_legend_handles_labels()
        fig.legend(handles=lines + lines2, labels=labels + labels2, loc='upper center', bbox_to_anchor=(0.5, 0.03), ncol=len(labels + labels2), prop={'size': 9})

        plt.tight_layout(rect=[0, 0, 1, 0.95]) # Ajusta para dar espaço à legenda
        plt.savefig(f'dados_full/{artista}/plots/9 - Monetização por formatos.png', dpi=STYLE_CONFIG['dpi'], transparent=STYLE_CONFIG['transparent'], bbox_inches='tight')
        plt.close(fig)

    except Exception as e:
        print(f"❌ Erro em revenue_per_type_chart para '{artista}': {e}")


def conversion_graph(artista, file_path):
    """
    Gera o gráfico de Conversão (CTR vs Views) com o novo design padronizado.
    """
    try:
        # --- 1. CARREGAMENTO E PREPARAÇÃO DOS DADOS ---
        df = pd.read_excel(file_path, sheet_name='Resultado')
        meses = pd.to_datetime(df.columns[2:], errors='coerce').dropna()
        
        # Garante que os dados numéricos sejam lidos corretamente
        df_numeric = df.iloc[:, 2:].apply(pd.to_numeric, errors='coerce')

        ctr_vod = df_numeric.iloc[24].astype(float)
        ctr_lives = df_numeric.iloc[25].astype(float)
        views_sem_shorts = df_numeric.iloc[7].astype(float)
        views_vod = df_numeric.iloc[9].astype(float)
        views_lives = df_numeric.iloc[10].astype(float)

        # --- 2. CRIAÇÃO DO GRÁFICO ---
        fig, ax = plt.subplots(figsize=STYLE_CONFIG['figsize_standard'])
        index = np.arange(len(meses))
        bar_width = 0.35

        # Barras para CTR com as novas cores
        bars1 = ax.bar(index - bar_width/2, ctr_vod, bar_width, color=STYLE_CONFIG['colors']['vod'], label='CTR (%) VOD Novo')
        bars2 = ax.bar(index + bar_width/2, ctr_lives, bar_width, color=STYLE_CONFIG['colors']['live'], label='CTR (%) Live Novo')

        # Adiciona os rótulos de dados às barras com o novo estilo
        for bar in bars1:
            yval = bar.get_height()
            ax.text(bar.get_x() + bar.get_width()/2.0, yval, f'{yval:.1f}%',
                    ha='center', va='bottom', bbox=STYLE_CONFIG['label_bbox_props'],
                    **STYLE_CONFIG['label_font_props'])

        for bar in bars2:
            yval = bar.get_height()
            ax.text(bar.get_x() + bar.get_width()/2.0, yval, f'{yval:.1f}%',
                    ha='center', va='bottom', bbox=STYLE_CONFIG['label_bbox_props'],
                    **STYLE_CONFIG['label_font_props'])
        
        # Ajusta o limite do eixo para dar espaço aos rótulos
        ax.set_ylim(top=max(ctr_vod.max(), ctr_lives.max()) * 1.2)

        # --- 3. CONFIGURAÇÃO DOS EIXOS E TÍTULO ---
        ax.set_xticks(index)
        ax.set_xticklabels(meses.strftime("%b/%Y"), rotation=45)
        ax.set_ylabel('CTR (%)', **STYLE_CONFIG['font_props_label'])
        ax.set_xlabel('Meses', **STYLE_CONFIG['font_props_label'])
        ax.set_title("Taxa de Cliques (CTR) e Visualizações",pad=25, **STYLE_CONFIG['font_props_title'])

        efeito_contorno = [path_effects.withStroke(linewidth=3, foreground='gray')]

        # --- 4. EIXO SECUNDÁRIO COM AS LINHAS DE VISUALIZAÇÕES ---
        ax2 = ax.twinx()
        ax2.plot(index, views_sem_shorts, color=STYLE_CONFIG['colors']['positive'], marker='o', label='Views sem Shorts', path_effects=efeito_contorno)
        ax2.plot(index, views_vod, color=STYLE_CONFIG['colors']['primary_blue'], marker='s', label='Views VOD', path_effects=efeito_contorno)
        ax2.plot(index, views_lives, color=STYLE_CONFIG['colors']['primary_yellow'], marker='^', label='Views Lives', path_effects=efeito_contorno)
        ax2.set_ylabel('Visualizações', **STYLE_CONFIG['font_props_label'])
        ax2.set_ylim(0, views_sem_shorts.max() * 1.15)
        ax2.yaxis.set_major_formatter(FuncFormatter(formatar_eixo_numeros))

        # --- 5. LEGENDA E SALVAMENTO ---
        lines, labels = ax.get_legend_handles_labels()
        lines2, labels2 = ax2.get_legend_handles_labels()
        fig.legend(handles=lines + lines2, labels=labels + labels2, loc='upper center', bbox_to_anchor=(0.5, 0.03), ncol=5, prop={'size': 9})

        plt.tight_layout(rect=[0, 0, 1, 0.95])
        plt.savefig(f'dados_full/{artista}/plots/10 - Conversao.png', dpi=STYLE_CONFIG['dpi'], transparent=STYLE_CONFIG['transparent'], bbox_inches='tight')
        plt.close(fig)
        
    except Exception as e:
        print(f"❌ Erro ao gerar o gráfico 'conversion_graph': {e}")


def gerar_grafico_views(artista, file_path):
    """
    Cria um gráfico de visualizações com rótulos de dados que não se sobrepõem.
    """
    try:
        # --- 1. CARREGAMENTO DOS DADOS ---
        df = pd.read_excel(file_path, sheet_name="Resultado")
        meses = pd.to_datetime(df.columns[2:], errors='coerce').dropna().to_series().dt.strftime('%b')
        
        idx_views_vod_novo = 80; idx_views_lives_novo = 81; idx_views_shorts_novo = 82
        idx_views_vod_velho = 83; idx_views_lives_velho = 84; idx_views_shorts_velho = 85
        idx_views_total = 86
        
        views_vod_novo = df.iloc[idx_views_vod_novo, 2:].astype(float)
        views_lives_novo = df.iloc[idx_views_lives_novo, 2:].astype(float)
        views_shorts_novo = df.iloc[idx_views_shorts_novo, 2:].astype(float)
        views_vod_velho = df.iloc[idx_views_vod_velho, 2:].astype(float)
        views_lives_velho = df.iloc[idx_views_lives_velho, 2:].astype(float)
        views_shorts_velho = df.iloc[idx_views_shorts_velho, 2:].astype(float)
        views_total = df.iloc[idx_views_total, 2:].astype(float)

        # --- 2. CONFIGURAÇÃO DO GRÁFICO ---
        fig, ax1 = plt.subplots(figsize=STYLE_CONFIG['figsize_standard'])
        index = np.arange(len(meses))
        bar_width = 0.25
        
        chart_colors = {
            'vod_novo': '#6B7FFF', 'vod_velho': '#3157F7',
            'lives_novo': '#FCA5A5', 'lives_velho': '#EF4444',
            'shorts_novo': '#a78bfa', 'shorts_velho': '#7c3aed',
            'total_line': '#22C55E'
        }

        # --- 3. DESENHO DAS BARRAS ---
        rects_vod_velho = ax1.bar(index - bar_width, views_vod_velho, bar_width, label='Views VOD Velho', color=chart_colors['vod_velho'])
        rects_vod_novo = ax1.bar(index - bar_width, views_vod_novo, bar_width, bottom=views_vod_velho, label='Views VOD Novo', color=chart_colors['vod_novo'])
        rects_lives_velho = ax1.bar(index, views_lives_velho, bar_width, label='Views Live Velho', color=chart_colors['lives_velho'])
        rects_lives_novo = ax1.bar(index, views_lives_novo, bar_width, bottom=views_lives_velho, label='Views Live Novo', color=chart_colors['lives_novo'])
        rects_shorts_velho = ax1.bar(index + bar_width, views_shorts_velho, bar_width, label='Views Shorts Velho', color=chart_colors['shorts_velho'])
        rects_shorts_novo = ax1.bar(index + bar_width, views_shorts_novo, bar_width, bottom=views_shorts_velho, label='Views Shorts Novo', color=chart_colors['shorts_novo'])

        # --- LÓGICA DE RÓTULOS ATUALIZADA ---
        def add_stacked_labels(bottom_rects, top_rects):
            for rect_bottom, rect_top in zip(bottom_rects, top_rects):
                height_bottom = rect_bottom.get_height()
                height_top = rect_top.get_height()

                # Rótulo da barra de baixo
                if height_bottom > 0:
                    # ALTERAÇÃO: Posição ajustada para a metade inferior da barra
                    y_pos = rect_bottom.get_y() + height_bottom * 0.25 
                    ax1.text(rect_bottom.get_x() + rect_bottom.get_width() / 2., y_pos,
                             formatar_eixo_numeros(height_bottom, None), ha='center', va='center', color='white',
                             bbox=dict(boxstyle='round,pad=0.3', lw=0, facecolor=rect_bottom.get_facecolor(), alpha=0.9),
                             **STYLE_CONFIG['label_font_props'])

                # Rótulo da barra de cima
                if height_top > 0:
                    # ALTERAÇÃO: Posição ajustada para a metade superior da barra
                    y_pos = rect_top.get_y() + height_top * 0.75
                    ax1.text(rect_top.get_x() + rect_top.get_width() / 2., y_pos,
                             formatar_eixo_numeros(height_top, None), ha='center', va='center', color='white',
                             bbox=dict(boxstyle='round,pad=0.3', lw=0, facecolor=rect_top.get_facecolor(), alpha=0.9),
                             **STYLE_CONFIG['label_font_props'])

        # Aplica a lógica para cada grupo de barras
        add_stacked_labels(rects_vod_velho, rects_vod_novo)
        add_stacked_labels(rects_lives_velho, rects_lives_novo)
        add_stacked_labels(rects_shorts_velho, rects_shorts_novo)
        
        # --- 4. CONFIGURAÇÃO FINAL ---
        ax1.set_xlabel('Meses', **STYLE_CONFIG['font_props_label']); ax1.set_ylabel('Visualizações', **STYLE_CONFIG['font_props_label']); ax1.set_title('Visualizações Velho vs Novo', **STYLE_CONFIG['font_props_title']); ax1.set_xticks(index); ax1.set_xticklabels(meses); ax1.yaxis.set_major_formatter(FuncFormatter(formatar_eixo_numeros))
        ax2 = ax1.twinx(); ax2.plot(meses, views_total, marker='o', color=chart_colors['total_line'], label='Views Totais', linewidth=2.5)
        ax2.set_ylabel('Visualizações Totais', color=chart_colors['total_line'], **STYLE_CONFIG['font_props_label']); ax2.tick_params(axis='y', labelcolor=chart_colors['total_line']); ax2.set_ylim(bottom=0, top=views_total.max() * 1.2); ax2.yaxis.set_major_formatter(FuncFormatter(formatar_eixo_numeros))
        lines, labels = ax1.get_legend_handles_labels(); lines2, labels2 = ax2.get_legend_handles_labels()
        fig.legend(handles=lines + lines2, labels=labels + labels2, loc='upper center', bbox_to_anchor=(0.5, 0.03), ncol=4, prop={'size': 9})
        fig.tight_layout(rect=[0, 0, 1, 0.95])
        plt.savefig(f'dados_full/{artista}/plots/19 - Views_Novo_vs_Velho.png', dpi=STYLE_CONFIG['dpi'], transparent=STYLE_CONFIG['transparent'], bbox_inches='tight')
        plt.close(fig)
        
    except Exception as e:
        print(f"❌ Erro em gerar_grafico_views para '{artista}': {e}")


def gerar_grafico_qualidade(artista, file_path, tipo_conteudo, plot_index):
    """
    Gera o gráfico de qualidade, com background em todos os rótulos de dados das barras.
    Se o tipo for 'shorts', exibe os valores de tempo em segundos.
    """
    try:
        df = pd.read_excel(file_path, sheet_name='Resultado')
        
        meses = df.columns[2:].to_list()
        
        # ALTERAÇÃO 1: Definir variáveis de configuração com base no tipo de conteúdo
        unidade_tempo = "min"
        formato_tempo = ".1f" # Formato para minutos (uma casa decimal)
        conversor_tempo = converter_excel_time_para_minutos

        if tipo_conteudo == 'vod':
            metric_rows = {'tamanho': 28, 'porcentagem': 26, 'tempo_assistido': 30, 'impressoes': 15}
            title, label_barra, label_linha = "Qualidade dos Novos VODs", "Tamanho Médio do Vídeo (min)", "Média de Impressões por VOD"
        elif tipo_conteudo == 'shorts':
            metric_rows = {'tamanho': 77, 'porcentagem': 78, 'tempo_assistido': 79, 'impressoes': 17}
            # ALTERAÇÃO 2: Mudar a unidade e o conversor para segundos
            unidade_tempo = "s"
            formato_tempo = ".0f" # Formato para segundos (número inteiro)
            conversor_tempo = converter_excel_time_para_segundos # << USA A NOVA FUNÇÃO
            title, label_barra, label_linha = "Qualidade dos Novos Shorts", f"Tamanho Médio do Short ({unidade_tempo})", "Média de Impressões por Short"
        else: # live
            metric_rows = {'tamanho': 29, 'porcentagem': 27, 'tempo_assistido': 31, 'impressoes': 16}
            title, label_barra, label_linha = "Qualidade das Novas Lives", "Tamanho Médio da Live (min)", "Média de Impressões por Live"

        tamanho_medio_data = df.iloc[metric_rows['tamanho'], 2:].values
        porcentagem_media_data = df.iloc[metric_rows['porcentagem'], 2:].values
        tempo_medio_assistido_data = df.iloc[metric_rows['tempo_assistido'], 2:].values
        impressoes_media_data = df.iloc[metric_rows['impressoes'], 2:].values

        # ALTERAÇÃO 3: Usar a função de conversão definida dinamicamente
        tamanho_medio = [conversor_tempo(val) for val in tamanho_medio_data]
        porcentagem_media = [float(val) if val is not None else 0.0 for val in porcentagem_media_data]
        tempo_medio_assistido = [conversor_tempo(val) for val in tempo_medio_assistido_data]
        impressoes_media = [float(val) if val is not None else 0.0 for val in impressoes_media_data]
        
        # O cálculo da altura visual não muda
        altura_porcentagem_visual = [t * (p / 100) for t, p in zip(tamanho_medio, porcentagem_media)]
        
        fig, ax1 = plt.subplots(figsize=STYLE_CONFIG['figsize_standard'])
        positions = list(range(len(meses)))
        
        ax1.bar(positions, tamanho_medio, color=STYLE_CONFIG['colors']['secondary_blue'], label=label_barra)
        ax1.bar(positions, altura_porcentagem_visual, color=STYLE_CONFIG['colors']['primary_blue'], label="Tempo Médio Assistido")
        
        for i, (x, total, parcial, perc, tempo) in enumerate(zip(positions, tamanho_medio, altura_porcentagem_visual, porcentagem_media, tempo_medio_assistido)):
            # ALTERAÇÃO 4: Usar as variáveis de unidade e formato nos rótulos
            # Rótulo dentro da barra azul escura
            ax1.annotate(f"{tempo:{formato_tempo}} {unidade_tempo}\n({perc:.1f}%)",  
                         xy=(x, parcial / 2),  
                         ha="center", va="center", color='#ffffff',  
                         bbox=dict(boxstyle='round,pad=0.3', lw=0, facecolor=STYLE_CONFIG['colors']['primary_blue'], alpha=0.9),
                         **STYLE_CONFIG['label_font_props'])
            
            # Rótulo acima da barra azul clara
            ax1.annotate(f"{total:{formato_tempo}} {unidade_tempo}",  
                         xy=(x, total + (max(tamanho_medio) * 0.02)), # Pequeno ajuste de posição
                         ha="center", va="bottom", color=STYLE_CONFIG['colors']['text_dark'],
                         bbox=dict(boxstyle='round,pad=0.3', lw=0, facecolor=STYLE_CONFIG['colors']['secondary_blue'], alpha=0.9),
                         **STYLE_CONFIG['label_font_props'])
        
        # O resto da função permanece o mesmo
        ax1.set_ylabel(label_barra, **STYLE_CONFIG['font_props_label'])
        ax1.set_ylim(0, max(tamanho_medio) * 1.3 if tamanho_medio else 1)
        ax1.set_xticks(positions)
        ax1.set_xticklabels(meses)
        
        ax2 = ax1.twinx()
        ax2.plot(positions, impressoes_media, color=STYLE_CONFIG['colors']['primary_red'], marker='o', linewidth=2, label=label_linha)
        ax2.set_ylabel("Impressões por Envio", color=STYLE_CONFIG['colors']['primary_red'], **STYLE_CONFIG['font_props_label'])
        ax2.tick_params(axis='y', colors=STYLE_CONFIG['colors']['primary_purple'])
        ax2.set_ylim(0, max(impressoes_media) * 1.2 if impressoes_media else 1)
        ax2.yaxis.set_major_formatter(FuncFormatter(formatar_eixo_numeros))
        
        lines1, labels1 = ax1.get_legend_handles_labels()
        lines2, labels2 = ax2.get_legend_handles_labels()
        ax2.legend(lines1 + lines2, labels1 + labels2, loc=(0.075, 1), fontsize=9, ncol=len(labels1 + labels2))
        
        fig.suptitle(title, x=0.035, ha='left', **STYLE_CONFIG['font_props_subtitle'], y=0.73, rotation=90)       
        plt.tight_layout(rect=[1, 1, 1, 0.96])
        fig.savefig(f'dados_full/{artista}/plots/{plot_index} - Qualidade_{tipo_conteudo.capitalize()}.png', dpi=STYLE_CONFIG['dpi'], bbox_inches="tight", transparent=STYLE_CONFIG['transparent'])
        plt.close(fig)
        
    except Exception as e:
        print(f"❌ Erro em gerar_grafico_qualidade para '{artista}' ({tipo_conteudo}): {e}")
        

def dfOrigem(nomeArquivo1, nomeArquivo2):
    df1 = pd.read_csv(nomeArquivo1 + ".csv"); df2 = pd.read_csv(nomeArquivo2 + ".csv"); df = pd.concat([df1, df2], ignore_index=True); df['Data'] = pd.to_datetime(df['Data']); df.sort_values('Data', inplace=True); df["Mês"] = df['Data'].dt.strftime('%Y-%m'); df = df[['Origem do tráfego', 'Visualizações', 'Mês']].groupby(by=['Mês', 'Origem do tráfego']).sum().reset_index(); origensImportantes = ("Recursos de navegação","Vídeos sugeridos","Páginas do canal","Externa","Notificações","Pesquisa do YouTube","Playlists","Publicidade no YouTube"); b = df[~df['Origem do tráfego'].isin(origensImportantes)].groupby(by=['Mês']).sum(numeric_only=True); b["Origem do tráfego"] = 'Outros'; b.reset_index(inplace=True); df = pd.concat([b, df[df['Origem do tráfego'].isin(origensImportantes)]]); total = df.groupby('Origem do tráfego')['Visualizações'].sum().sort_values(ascending=False).index.tolist(); df = df.pivot(index="Mês", columns="Origem do tráfego", values="Visualizações").fillna(0); df = df.reset_index().sort_values(by="Mês").reset_index(drop=True)
    return df, total


def traficSorce_graph(df, x_labels, total, title, color_map, artista):
    try:
        fig, ax = plt.subplots(figsize=STYLE_CONFIG['figsize_wide']); bottom = np.zeros(len(x_labels))
        for n, origem in enumerate(total):
            if origem in df.columns: y_values = df[origem].values; ax.bar(x=x_labels, height=y_values, color=color_map[n % len(color_map)], bottom=bottom, label=origem); bottom += y_values
        ax.legend(loc='upper left', bbox_to_anchor=(1, 1)); ax.yaxis.set_major_formatter(FuncFormatter(formatar_eixo_numeros)); ax.set_title(title.title(), **STYLE_CONFIG['font_props_title']); ax.tick_params(axis='x', rotation=45); ax.set_ylabel('Visualizações', **STYLE_CONFIG['font_props_label']); plt.tight_layout()
        plt.savefig(f'dados_full/{artista}/plots/13 - Origem_do_trafego.png', dpi=STYLE_CONFIG['dpi'], transparent=STYLE_CONFIG['transparent']); plt.close(fig)
    except Exception as e: print(f"❌ Erro em traficSorce_graph para '{artista}': {e}")
    
    
def subscription_growth(artista, file_path):
    """
    Gera o gráfico de crescimento de inscrições com espaçamento ajustado,
    rótulos com background e posicionamento vertical aprimorado.
    """
    try:
        # --- 1. CARREGAMENTO E PREPARAÇÃO DOS DADOS (sem alterações) ---
        df = pd.read_excel(file_path, sheet_name="Resultado")
        meses = pd.to_datetime(df.columns[2:], errors='coerce').dropna().to_series().dt.strftime('%b')
        insc_novo_vod = df.iloc[41, 2:].astype(float)
        insc_novo_live = df.iloc[42, 2:].astype(float)
        insc_novo_shorts = df.iloc[43, 2:].astype(float)
        insc_velho_vod = df.iloc[44, 2:].astype(float)
        insc_velho_live = df.iloc[45, 2:].astype(float)
        insc_velho_shorts = df.iloc[46, 2:].astype(float)
        insc_total = df.iloc[50, 2:].astype(float)
        
        # --- 2. CONFIGURAÇÃO DO GRÁFICO ---
        fig, ax1 = plt.subplots(figsize=STYLE_CONFIG['figsize_standard'])
        width = 0.25 
        x = np.arange(len(meses))
        
        # ALTERAÇÃO 1: Aumentado o espaçamento lateral entre os grupos de barras
        spacing_factor = 1.2 # Aumentado de 1.1 para 1.2 para mais espaço

        # --- 3. DESENHO DAS BARRAS COM NOVO ESPAÇAMENTO ---
        # As barras agora usam o novo 'spacing_factor' para se afastarem mais
        bars_vod_novo = ax1.bar(x - width * spacing_factor, insc_novo_vod, width, label='VOD Novo', color=STYLE_CONFIG['colors']['vod'])
        bars_vod_velho = ax1.bar(x - width * spacing_factor, insc_velho_vod, width, bottom=insc_novo_vod, label='VOD Velho', color=STYLE_CONFIG['colors']['secondary_blue'])
        
        bars_live_novo = ax1.bar(x, insc_novo_live, width, label='Live Novo', color=STYLE_CONFIG['colors']['live'])
        bars_live_velho = ax1.bar(x, insc_velho_live, width, bottom=insc_novo_live, label='Live Velho', color=STYLE_CONFIG['colors']['accent_purple'])
        
        bars_shorts_novo = ax1.bar(x + width * spacing_factor, insc_novo_shorts, width, label='Shorts Novo', color=STYLE_CONFIG['colors']['shorts'])
        bars_shorts_velho = ax1.bar(x + width * spacing_factor, insc_velho_shorts, width, bottom=insc_novo_shorts, label='Shorts Velho', color=f"{STYLE_CONFIG['colors']['shorts']}80")

        # --- 4. LÓGICA DE RÓTULOS COM BACKGROUND E ESPAÇAMENTO VERTICAL ---
        def add_stacked_labels(rects_novo, rects_velho):
            for rect_n, rect_v in zip(rects_novo, rects_velho):
                height_novo = rect_n.get_height()
                height_velho = rect_v.get_height()

                # Rótulo para a barra de CIMA ("Novo")
                if height_novo > 0:
                    # ALTERAÇÃO 2: Posição do rótulo de cima ajustada para 80% da sua altura, empurrando-o para cima
                    y_pos_novo = rect_n.get_y() + height_novo * 0.80
                    ax1.text(rect_n.get_x() + rect_n.get_width() / 2, y_pos_novo,
                             f"{formatar_eixo_numeros(height_novo, None)}", ha='center', va='center',
                             color='white', bbox=dict(boxstyle='round,pad=0.3', lw=0, facecolor=rect_n.get_facecolor(), alpha=0.9),
                             **STYLE_CONFIG['label_font_props'])

                # Rótulo para a barra de BAIXO ("Velho")
                if height_velho > 0:
                    # ALTERAÇÃO 3: Posição do rótulo de baixo ajustada para 20% da sua altura, empurrando-o para baixo
                    y_pos_velho = rect_v.get_y() + height_velho * 0.20
                    ax1.text(rect_v.get_x() + rect_v.get_width() / 2, y_pos_velho,
                             f"{formatar_eixo_numeros(height_velho, None)}", ha='center', va='center',
                             color='white', bbox=dict(boxstyle='round,pad=0.3', lw=0, facecolor=rect_v.get_facecolor(), alpha=0.9),
                             **STYLE_CONFIG['label_font_props'])

        # Aplica a lógica para cada grupo de barras
        add_stacked_labels(bars_vod_novo, bars_vod_velho)
        add_stacked_labels(bars_live_novo, bars_live_velho)
        add_stacked_labels(bars_shorts_novo, bars_shorts_velho)

        # --- 5. CONFIGURAÇÃO FINAL E SALVAMENTO (sem alterações) ---
        ax1.set_title('Inscrições por Tipo de Conteúdo', **STYLE_CONFIG['font_props_title'])
        ax1.set_ylabel('Número de Inscrições', **STYLE_CONFIG['font_props_label'])
        ax1.set_xlabel('Meses', **STYLE_CONFIG['font_props_label'])
        ax1.set_xticks(x)
        ax1.set_xticklabels(meses)
        ax1.yaxis.set_major_formatter(FuncFormatter(formatar_eixo_numeros))
        
        ax2 = ax1.twinx()
        ax2.plot(x, insc_total, color=STYLE_CONFIG['colors']['positive'], marker='o', linestyle='-', label='Total de Inscrições')
        ax2.set_ylabel('Total de Inscrições', **STYLE_CONFIG['font_props_label'], color=STYLE_CONFIG['colors']['positive'])
        ax2.tick_params(axis='y', labelcolor=STYLE_CONFIG['colors']['positive'])
        ax2.set_ylim(bottom=0, top=max(insc_total) * 1.2 if not insc_total.empty else 1)
        ax2.yaxis.set_major_formatter(FuncFormatter(formatar_eixo_numeros))
        
        fig.legend(loc='upper right', bbox_to_anchor=(0.95, 0.03), ncol=7)
        plt.tight_layout(rect=[0, 0, 0.9, 1])
        plt.savefig(f'dados_full/{artista}/plots/14 - Inscricoes_por_Tipo_de_Conteudo.png', dpi=STYLE_CONFIG['dpi'], transparent=STYLE_CONFIG['transparent'], bbox_inches='tight')
        plt.close(fig)
        
    except Exception as e:
        print(f"❌ Erro em subscription_growth para '{artista}': {e}")


def gerar_grafico_engajamento_tipo(artista, file_path, tipo_conteudo, plot_index):
    """
    Gera o gráfico de engajamento com cores de barra padronizadas e uma linha estilizada.
    """
    try:
        df = pd.read_excel(file_path, sheet_name='Resultado')
        meses = pd.to_datetime(df.columns[2:], errors='coerce').dropna()
        
        if tipo_conteudo == 'vod':
            engajamento_row, impressoes_row, title = 63, 15, "Engajamento VODs"
        elif tipo_conteudo == 'shorts':
            engajamento_row, impressoes_row, title = 76, 17, "Engajamento Shorts"   
        else: # live
            engajamento_row, impressoes_row, title = 64, 16, "Engajamento Lives"
            
        engajamento = df.iloc[engajamento_row, 2:].astype(float)
        impressoes_media = df.iloc[impressoes_row, 2:].astype(float)
        
        fig, ax = plt.subplots(figsize=(10, 6))
        index = np.arange(len(meses))
        
        # --- ALTERAÇÃO 1: Cor da barra padronizada ---
        # Todas as barras agora usarão a cor 'primary_blue'
        bars = ax.bar(index, engajamento, 0.4, color=STYLE_CONFIG['colors']['primary_blue'], label=f'Média de engajamento {tipo_conteudo.capitalize()}')
        
        for bar in bars:
            yval = bar.get_height()
            ax.text(bar.get_x() + bar.get_width() / 2, yval, f'{yval:.1f}', 
                    ha='center', va='bottom', bbox=STYLE_CONFIG['label_bbox_props'], 
                    **STYLE_CONFIG['label_font_props'])
            
        ax.set_xticks(index)
        ax.set_xticklabels(meses.strftime("%b/%Y"), rotation=0)
        ax.set_ylabel('Média de engajamento', **STYLE_CONFIG['font_props_label'])
        ax.set_xlabel('Meses', **STYLE_CONFIG['font_props_label'])
                
        ax2 = ax.twinx()

        # --- ALTERAÇÃO 2: Estilo da linha ---
        # Define a cor da linha como a mesma dos 'shorts'
        cor_linha = STYLE_CONFIG['colors']['primary_red']
        # Define o efeito de contorno cinza
        efeito_contorno = [path_effects.withStroke(linewidth=3, foreground='gray')]
        
        # Aplica a nova cor e o contorno à linha
        ax2.plot(index, impressoes_media, color=cor_linha, marker='o', 
                 label=f"Média de Impressões por {tipo_conteudo.capitalize()} Novo",
                 path_effects=efeito_contorno)
        
        ax2.set_ylabel('Impressões', **STYLE_CONFIG['font_props_label'], color=cor_linha)
        ax2.tick_params(axis='y', labelcolor=cor_linha)
        ax2.set_ylim(0, max(impressoes_media) * 1.2 if not impressoes_media.empty else 1)
        ax2.yaxis.set_major_formatter(FuncFormatter(formatar_eixo_numeros))
        
        fig.suptitle(title, x=0.055, y=0.475, ha='left', va='center', rotation='vertical', **STYLE_CONFIG['font_props_title'])
        
        plt.subplots_adjust(left=0.15)
        fig.legend(loc='upper right', bbox_to_anchor=(0.95, 0.95), ncol=2)

        plt.savefig(f'dados_full/{artista}/plots/{plot_index} - Engajamento_{tipo_conteudo.capitalize()}.png', dpi=STYLE_CONFIG['dpi'], transparent=STYLE_CONFIG['transparent'], bbox_inches='tight')
        plt.close(fig)
        
    except Exception as e:
        print(f"❌ Erro em gerar_grafico_engajamento_tipo para '{artista}': {e}")


def gerar_grafico_comunidade(artista):
    """
    Gera um gráfico de barras e linha para analisar publicações e impressões da comunidade.
    A função usa a data mais recente do arquivo 'total.csv' como referência para
    filtrar os últimos 6 meses de dados.
    """
    try:
        # --- 1. Carregar e Preparar Dados (total.csv) ---
        data_meses = pd.read_csv(f'dados_full/{artista}/total.csv').drop(0)
        data_meses["Data"] = pd.to_datetime(data_meses["Data"])

        # --- 2. Lógica para Definir o Período de 6 Meses ---
        
        # *** CORREÇÃO DEFINITIVA APLICADA AQUI ***
        # Encontra a data mais recente (que o pandas interpreta como o dia 1 do mês)
        data_mais_recente_inicio_mes = data_meses["Data"].max()
        # Ajusta a data para ser o ÚLTIMO dia do mês, garantindo que todo o mês seja incluído.
        data_mais_recente = data_mais_recente_inicio_mes + pd.offsets.MonthEnd(0)
        
        # Define o limite inferior da janela de 6 meses
        primeiro_dia_mes_recente = data_mais_recente.replace(day=1)
        data_limite_inferior = primeiro_dia_mes_recente - pd.DateOffset(months=5)

        # Filtra os dados de 'total.csv' para o período desejado
        data_meses_filtrada = data_meses[
            (data_meses["Data"] >= data_limite_inferior) &
            (data_meses["Data"] <= data_mais_recente)
        ].copy()

        # --- 3. Carregar e Preparar Dados (comunidade.csv) ---
        data_comu = pd.read_csv(f'dados_full/{artista}/comunidade.csv')
        data_comu["Horário de publicação da postagem"] = pd.to_datetime(
            data_comu["Horário de publicação da postagem"], errors="coerce"
        )
        data_comu.dropna(subset=["Horário de publicação da postagem"], inplace=True)

        # Filtra os dados de 'comunidade.csv' para o mesmo período
        data_comu_filtrada = data_comu[
            (data_comu["Horário de publicação da postagem"] >= data_limite_inferior) &
            (data_comu["Horário de publicação da postagem"] <= data_mais_recente)
        ].copy()

        # --- 4. Agrupar e Juntar os Dados ---
        data_comu_filtrada["Mes"] = data_comu_filtrada["Horário de publicação da postagem"].dt.month
        comuCount = data_comu_filtrada.groupby("Mes").size().reset_index(name='Publicações na Comunidade')

        df_graf_com = data_meses_filtrada.copy()
        df_graf_com["Mes"] = df_graf_com["Data"].dt.month
        df_graf_com = df_graf_com.merge(comuCount, on="Mes", how="left").fillna(0)

        # --- 5. Ordenar e Formatar o DataFrame para o Gráfico ---
        df_graf_com['AnoMesTemp'] = df_graf_com['Data'].dt.to_period('M')
        df_graf_com = df_graf_com.sort_values(by='AnoMesTemp').drop(columns='AnoMesTemp')

        meses_ordem = {
            1: "Jan", 2: "Fev", 3: "Mar", 4: "Abr", 5: "Mai", 6: "Jun",
            7: "Jul", 8: "Ago", 9: "Set", 10: "Out", 11: "Nov", 12: "Dez"
        }
        df_graf_com['MesAbreviado'] = df_graf_com['Mes'].map(meses_ordem)
        df_graf_com.set_index('MesAbreviado', inplace=True)
        
        # --- 6. Geração do Gráfico ---
        fig, ax1 = plt.subplots(figsize=STYLE_CONFIG['figsize_standard'])
        ax2 = ax1.twinx()

        x = np.arange(len(df_graf_com.index))
        bars = ax1.bar(x, df_graf_com["Publicações na Comunidade"], width=0.4, color=STYLE_CONFIG['colors']['primary_blue'], label="Publicações na Comunidade")
        
        efeito_contorno = [path_effects.withStroke(linewidth=3, foreground='gray')]
        ax2.plot(x, df_graf_com["Impressões da postagem"], marker="o", color=STYLE_CONFIG['colors']['primary_yellow'], label="Impressões da postagem", path_effects=efeito_contorno)

        for bar in bars:
            height = bar.get_height()
            ax1.text(bar.get_x() + bar.get_width() / 2, height, f'{int(height)}', ha='center', va='bottom', bbox=STYLE_CONFIG['label_bbox_props'], **STYLE_CONFIG['label_font_props'])

        # --- 7. Estilização e Configurações Finais do Gráfico ---
        ax1.set_ylim(0, max(df_graf_com["Publicações na Comunidade"]) * 1.2 if not df_graf_com["Publicações na Comunidade"].empty and df_graf_com["Publicações na Comunidade"].max() > 0 else 1)
        ax2.set_ylim(0, max(df_graf_com["Impressões da postagem"]) * 1.2 if not df_graf_com["Impressões da postagem"].empty and df_graf_com["Impressões da postagem"].max() > 0 else 1)

        ax1.set_xlabel("Mês", **STYLE_CONFIG['font_props_label'])
        ax1.set_ylabel("Publicações", **STYLE_CONFIG['font_props_label'])
        ax2.set_ylabel("Impressões na Comunidade", **STYLE_CONFIG['font_props_label'])
        plt.title("Comunidade: Publicações X Impressões", **STYLE_CONFIG['font_props_title'])

        ax1.set_xticks(x)
        ax1.set_xticklabels(df_graf_com.index, rotation=45, ha='right')
        ax1.grid(axis="y", linestyle="--", alpha=0.5)

        fig.legend(loc='lower center', bbox_to_anchor=(0.5, 0.05), ncol=2)
        plt.tight_layout(rect=[0, 0.1, 1, 1])
        
        # Exemplo para exibir o gráfico. Descomente se precisar.
        plt.savefig(f'dados_full/{artista}/plots/17 - Comunidade.png', dpi=STYLE_CONFIG['dpi'], transparent=STYLE_CONFIG['transparent'], bbox_inches='tight')
        plt.close(fig)

    except Exception as e:
        print(f"❌ Erro em gerar_grafico_engajamento para '{artista}': {e}")


def gerar_tabela_inscritos_avancada(artista):
     """
     Gera a tabela final de inscritos, com design avançado, lógica de 7 para 6 meses,
     cores de tendência e agora com os ícones de seta (↑/↓).
     """
     try:
         # --- 1. CARREGAMENTO E PREPARAÇÃO DOS DADOS ---
         df_full = pd.read_csv(f'dados_full/{artista}/total.csv').drop(0)
         df_full["Data"] = pd.to_datetime(df_full["Data"])
         df_full.sort_values('Data', ascending=True, inplace=True, ignore_index=True)

         if len(df_full) >= 7:
             df_full = df_full.tail(7).reset_index(drop=True)

         with open(f"dados_full/{artista}/sub.txt", "r", encoding='latin-1') as f: #adicionando encoding. Samuel caso continue o erro mudar 'latin-1' para 'windows-1252'
             inscant = int(''.join(f.readline().split('.')))

         inscant_inicial_periodo = inscant - df_full['Inscritos'].sum()
         total_acumulado = (df_full['Inscritos'].cumsum() + inscant_inicial_periodo)

         ganhos = pd.to_numeric(df_full['Inscrições obtidas'], errors='coerce').fillna(0)
         perdidos = pd.to_numeric(df_full['Inscrições perdidas'], errors='coerce').fillna(0)

         saldo = ganhos - perdidos
         crescimento_pct = (total_acumulado.pct_change() * 100).fillna(0).round(2)

         # --- 2. MONTAGEM DO DATAFRAME FINAL ---
         df_display = pd.DataFrame({
             "Inscritos Totais": total_acumulado.tail(6).values,
             "% de Crescimento": crescimento_pct.tail(6).values,
             "Inscritos Obtidos": ganhos.tail(6).values,
             "Inscritos Perdidos": perdidos.tail(6).values,
             "Saldo de Inscritos": saldo.tail(6).values
         })
         df_display = df_display.T
         df_display.columns = pd.to_datetime(df_full['Data'].tail(6)).dt.strftime('%b')

         df_full_trend = pd.DataFrame({
             "Inscritos Totais": total_acumulado, "% de Crescimento": crescimento_pct,
             "Inscritos Obtidos": ganhos, "Inscritos Perdidos": perdidos, "Saldo de Inscritos": saldo
         }).T
         df_full_trend.columns = pd.to_datetime(df_full['Data']).dt.strftime('%b')

         # --- 3. CRIAÇÃO E ESTILIZAÇÃO DA TABELA ---
         fig, ax = plt.subplots(figsize=(12, 4.5))
         ax.axis('off')

         tabela = ax.table(
             cellText=[[''] * len(df_display.columns)] * len(df_display.index),
             rowLabels=df_display.index, colLabels=df_display.columns,
             loc='center', cellLoc='right'
         )
         tabela.auto_set_font_size(False); tabela.scale(1, 2.8)

         # Ponto-chave da alteração: extrair a tendência de Inscritos Totais
         # para usá-la na lógica da métrica % de Crescimento
         total_inscritos_trend = get_trend(df_full_trend.iloc[0, -1], df_full_trend.iloc[0, -2])

         for (row, col_display), cell in tabela.get_celld().items():
             col_full = col_display + 1
             cell.set_edgecolor('none')
             if row == 0:
                 cell.set_text_props(ha='center', color='white', **STYLE_CONFIG['table_font_props']); cell.set_facecolor(STYLE_CONFIG['colors']['primary_blue'])
             elif col_display == -1:
                 cell.set_text_props(ha='left', va='center', color=STYLE_CONFIG['colors']['text_dark'], **STYLE_CONFIG['table_font_props']); cell.get_text().set_text(f"  {df_display.index[row-1]}"); cell.set_facecolor('#F0F4FF'); cell.set_edgecolor('#E5E7EB'); cell.set_linewidth(1); cell.set_width(0.4)
             else:
                 cell.set_facecolor('#FFFFFF' if row % 2 == 0 else '#F0F4FF')
                 metric_name = df_full_trend.index[row-1]; current_value = df_full_trend.iloc[row-1, col_full]; prev_value = df_full_trend.iloc[row-1, col_full - 1]
                 trend = get_trend(current_value, prev_value); color = STYLE_CONFIG['colors']['text_dark']
                 
                 # --- ALTERAÇÃO APLICADA AQUI ---
                 # Define a tendência a ser usada para a cor e o ícone
                 trend_to_use = trend
                 if metric_name == "% de Crescimento":
                    trend_to_use = total_inscritos_trend

                 # Lógica para definir a cor
                 if metric_name == "Inscritos Perdidos":
                     if trend_to_use == "up": color = STYLE_CONFIG['colors']['negative']
                     elif trend_to_use == "down": color = STYLE_CONFIG['colors']['positive']
                 elif trend_to_use == "up":
                     color = STYLE_CONFIG['colors']['positive']
                 elif trend_to_use == "down":
                     color = STYLE_CONFIG['colors']['negative']
                  
                 # Lógica para definir o ícone
                 icon = ''
                 if trend_to_use == "up": icon = '↑ '
                 elif trend_to_use == "down": icon = '↓ '
                  
                 display_val = ""
                 if pd.notna(current_value):
                     if metric_name == "% de Crescimento": display_val = f"{current_value:,.2f}%".replace(",", "X").replace(".", ",").replace("X", ".")
                     else: display_val = f"{int(current_value):,}".replace(",", ".")
                  
                 cell.get_text().set_text(f"{icon}{display_val}")
                 cell.get_text().set_color(color); cell.get_text().set_font_properties(STYLE_CONFIG['table_font_props']['fontproperties']); cell.get_text().set_size(STYLE_CONFIG['table_font_props']['size'])

         fig.suptitle("", y=0.5, **STYLE_CONFIG['font_props_title'])
         plt.tight_layout(rect=[0, 0, 1, 1])
          
         plt.savefig(f'dados_full/{artista}/plots/18 - Tabela de Inscritos.png', transparent=STYLE_CONFIG['transparent'], bbox_inches="tight", dpi=STYLE_CONFIG['dpi'])
         plt.close(fig)

     except Exception as e:
         print(f"Erro em gerar_tabela_inscritos_avancada para '{artista}': {e}")


def run(artista):
    # O loop 'for' e a chamada 'buscar_lista_artistas' foram REMOVIDOS.
    # 'artista' agora é recebido como argumento.
    print(f"\n--- Generating reports for: {artista} ---")
    os.makedirs(f'dados_full/{artista}/plots', exist_ok=True)
    file_path_4_1 = f'exports_tabelas/tabela_4.1_{artista}.xlsx'

    gerar_tabela_metricas_avancada(artista, 'VOD', 'videos.csv', 1)
    gerar_tabela_metricas_avancada(artista, 'Lives', 'lives.csv', 2)
    gerar_tabela_metricas_avancada(artista, 'Shorts', 'shorts.csv', 3)
    gerar_cards_detalhados(artista, file_path_4_1)
    publicated_table(artista, file_path_4_1)
    analyze_initial_updated(artista, file_path_4_1)
    watch_table(artista, file_path_4_1)
    monetization_graph(artista, file_path_4_1)
    revenue_per_type_chart(artista, file_path_4_1)
    conversion_graph(artista, file_path_4_1)
    gerar_grafico_qualidade(artista, file_path_4_1, 'vod', 11)
    gerar_grafico_qualidade(artista, file_path_4_1, 'live', 12)
    gerar_grafico_qualidade(artista, file_path_4_1, 'shorts', 12.5)
    try:
        semestralOrigem, total = dfOrigem(f'dados_full/{artista}/origem_lives', f'dados_full/{artista}/origem_vods')
        traficSorce_graph(semestralOrigem, semestralOrigem['Mês'], total, 'Views by Traffic Source', plt.cm.get_cmap('tab20').colors, artista)
    except Exception as e: print(f"❌ Error generating Traffic Source chart for '{artista}': {e}")
    subscription_growth(artista, file_path_4_1)
    gerar_grafico_engajamento_tipo(artista, file_path_4_1, 'vod', 15)
    gerar_grafico_engajamento_tipo(artista, file_path_4_1, 'live', 16)
    gerar_grafico_engajamento_tipo(artista, file_path_4_1, 'shorts', 16.5)
    gerar_grafico_comunidade(artista)
    gerar_tabela_inscritos_avancada(artista)
    gerar_grafico_views(artista, file_path_4_1)
    
    # Corrigindo o print que daria erro de encoding
    print(f"OK - Reports for '{artista}' completed.")


if __name__ == "__main__":
    # Pega o nome do artista do argumento passado pelo main.py
    if len(sys.argv) < 2:
        print("Error: No artist provided. This script must be called by main.py")
        sys.exit(1) # Sai com erro
    
    artista_argumento = sys.argv[1]
    
    # Executa a função run APENAS para esse artista
    run(artista_argumento)
