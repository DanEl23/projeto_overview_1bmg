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
    print(f"⚠️ Warning: 'DM Sans' font not found. Using default font. Error: {e}")
    mpl.rcParams["font.family"] = "sans-serif"
    dm_sans_bold = fm.FontProperties(weight='bold')


STYLE_CONFIG = {
    'colors': {
        'primary_blue': '#4f46e5', 'secondary_blue': '#c7d2fe', 'primary_purple': '#7c3aed',
        'primary_yellow':'#E2FC51', 'secondary_yellow': '#3157F7', 'accent_blue': '#60a5fa',
        'accent_purple': '#a78bfa', 'text_dark': '#1f2937', 'text_light': '#6b7280',
        'background': '#f9fafb', 'positive': '#10b981', 'negative': '#ef4444',
        'neutral': '#d1d5db', 'vod': '#4f46e5', 'live': '#3157F7', 'shorts': '#E2FC51',
    },
    'font_bold': dm_sans_bold,
    'font_props_title': {'fontproperties': dm_sans_bold, 'size': 20},
    'font_props_subtitle': {'fontproperties': dm_sans_bold, 'size': 16},
    'font_props_label': {'size': 12},
    'label_font_props': { # Bold font
        'fontproperties': dm_sans_bold,
        'size': 10
        # The 'color' key was removed to avoid conflicts
    },
    'label_bbox_props': { # Rounded background
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
    "Published videos": "▶",
    "Impressions": "👁",
    "Impressions click-through rate (%)": "🖱",
    "Views": "📊",
    "Average view duration": "⏰",
    "Average percentage viewed (%)": "📈",
    "Subscribers": "👥",
    "RPM (USD)": "$",
    "Estimated revenue (USD)": "💰"
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
    if abs(num) >= 1_000_000: return f'{num/1_000_000:.2f}M'
    if abs(num) >= 1_000: return f'{num/1_000:.0f}K'
    return str(int(num))


def converter_excel_time_para_minutos(time_val):
    """
    Converts an Excel time value to decimal minutes.
    Works with floats (Excel time format), strings ('HH:MM:SS')
    and datetime.time objects.
    """
    # If it's a number (common Excel format for time)
    if isinstance(time_val, (int, float)):
        # Excel stores time as a fraction of 1 day (24 hours * 60 minutes)
        return float(time_val) * 24 * 60

    # If it's a string like 'HH:MM:SS' or 'MM:SS'
    if isinstance(time_val, str) and ':' in time_val:
        parts = time_val.split(':')
        try:
            if len(parts) == 3: return (int(parts[0]) * 60) + int(parts[1]) + (int(parts[2]) / 60.0)
            if len(parts) == 2: return int(parts[0]) + (int(parts[1]) / 60.0)
        except (ValueError, TypeError):
            return 0.0 # Returns 0 if the string is not a valid format

    # If it's a Python time object
    if hasattr(time_val, 'hour'):
        return time_val.hour * 60 + time_val.minute + time_val.second / 60.0

    # If it's none of the expected formats, return 0
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
    """Determines the trend (up, down, neutral) by comparing two numeric values."""
    if not all(isinstance(v, (int, float)) for v in [current_val, prev_val]):
        return "neutral"
    if current_val > prev_val:
        return "up"
    if current_val < prev_val:
        return "down"
    return "neutral"


def get_performance_color(value, row_data):
    """
    Calculates the cell background color based on the relative performance of the metric in its own row.
    This is the Python version of your getPerformanceColor function.
    """
    # Filters only numeric values to avoid distorting the scale
    numeric_values = [v for v in row_data if isinstance(v, (int, float)) and v > 0]
    if not numeric_values:
        return {'facecolor': '#f9fafb', 'textcolor': '#1f2937'} # Neutral color

    max_val, min_val = max(numeric_values), min(numeric_values)
    range_val = max_val - min_val

    if range_val == 0 or not isinstance(value, (int, float)):
        return {'facecolor': '#f9fafb', 'textcolor': '#1f2937'} # Neutral color

    # Normalizes the value between 0 and 1
    normalized = (value - min_val) / range_val

    # Returns background and text colors based on performance
    if normalized >= 0.75:
        return {'facecolor': '#d1fae5', 'textcolor': '#065f46'} # Green
    if normalized >= 0.35:
        return {'facecolor': '#dbeafe', 'textcolor': '#1e40af'} # Blue
    return {'facecolor': '#fee2e2', 'textcolor': '#991b1b'}     # Red


def format_for_table(value):
    """Formats numbers for table display (K for thousand, M for million)."""
    if not isinstance(value, (int, float)):
        return str(value)
    if abs(value) >= 1_000_000:
        return f'{value/1_000_000:.1f}M'
    if abs(value) >= 1_000:
        return f'{value/1_000:.1f}K'
    if isinstance(value, float) and value < 100: # For percentages and small values
        return f'{value:.2f}'
    return str(int(value))


def gerar_tabela_metricas_avancada(artista, tipo_conteudo, nome_arquivo, plot_index):
    """
    Final and corrected version that generates the metrics table with functional font control.
    """
    try:
        # --- 1. DATA LOADING AND PREPARATION (COLUMN NAMES IN PORTUGUESE) ---
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

        # --- 2. TABLE CREATION ---
        fig, ax = plt.subplots(figsize=(16, 7)); ax.axis('off')

        tabela = ax.table(
            cellText=[[''] * len(df_display.columns)] * len(df_display.index),
            rowLabels=df_display.index, colLabels=df_display.columns,
            loc='center', cellLoc='right'
        )
        tabela.auto_set_font_size(False) # Prevents automatic adjustment
        tabela.scale(1, 3) # Adjusts cell height

        # --- 3. STYLING AND FILLING (CORRECTED LOGIC) ---
        for (row, col_display), cell in tabela.get_celld().items():
            col_full = col_display + 1; cell.set_edgecolor('none')

            # Gets font properties from STYLE_CONFIG
            font_properties = STYLE_CONFIG['table_font_props']

            if row == 0: # Header
                cell.set_text_props(**font_properties, color='white', ha='center')
                cell.set_facecolor('#3157F7'); cell.set_height(0.12)
            elif col_display == -1: # Metrics Column
                cell.set_text_props(**font_properties, ha='left', va='center', color='#142544')
                # Display metric names in English for the table row labels
                english_metric_names = {
                    'Vídeos publicados': 'Published videos',
                    'Impressões': 'Impressions',
                    'Taxa de cliques de impressões (%)': 'Impressions click-through rate (%)',
                    'Visualizações': 'Views',
                    'Duração média da visualização': 'Average view duration',
                    'Porcentagem visualizada média (%)': 'Average percentage viewed (%)',
                    'Inscritos': 'Subscribers',
                    'RPM (USD)': 'RPM (USD)',
                    'Receita estimada (USD)': 'Estimated revenue (USD)'
                }
                original_metric_name = df_full.index[row-1]
                display_metric_name = english_metric_names.get(original_metric_name, original_metric_name) # Fallback to original if not found
                cell.get_text().set_text(f"  {display_metric_name}")
                cell.set_facecolor('#F0F4FF'); cell.set_edgecolor('#E5E7EB'); cell.set_linewidth(1); cell.set_width(0.35)
            else: # Data Cells
                cell.set_facecolor('#FFFFFF' if row % 2 != 0 else '#F0F4FF')
                metric_name = df_full.index[row-1]; current_value = df_full.iloc[row-1, col_full]; prev_value = df_full.iloc[row-1, col_full - 1]
                trend = get_trend(current_value, prev_value)

                display_val = ""
                if pd.notna(current_value):
                    if metric_name in ["Impressões", "Visualizações", "Inscritos", "Vídeos publicados"]: display_val = custom_format(current_value)
                    elif metric_name == "Duração média da visualização":
                        minutes, seconds = divmod(int(current_value), 60); display_val = f"{minutes}:{seconds:02d}"
                    elif metric_name == "Receita estimada (USD)": display_val = format_currency(current_value)
                    else: display_val = dec_format(current_value)
                    if metric_name in ["Taxa de cliques de impressões (%)", "Porcentagem visualizada média (%)"]: display_val += "%"

                icon = ''; color = STYLE_CONFIG['colors']['text_dark']
                if trend == "up": icon = '↑ '; color = STYLE_CONFIG['colors']['positive']
                elif trend == "down": icon = '↓ '; color = STYLE_CONFIG['colors']['negative']

                cell.get_text().set_text(f"{icon}{display_val}")
                # APPLIES PROPERTIES CORRECTLY
                cell.get_text().set_font_properties(font_properties['fontproperties'])
                cell.get_text().set_size(font_properties['size'])
                cell.get_text().set_color(color)

        fig.suptitle("", y=0.5, **STYLE_CONFIG['font_props_title'])
        plt.tight_layout(rect=[0, 0, 1, 0.95])

        plt.savefig(
            f'dados_full/{artista}/plots/{plot_index} - Metrics_{tipo_conteudo}_Advanced.png',
            dpi=STYLE_CONFIG['dpi'], bbox_inches='tight', pad_inches=0.1, transparent=STYLE_CONFIG['transparent']
        )
        plt.close(fig)
    except Exception as e:
        print(f"❌ Error in gerar_tabela_metricas_avancada for '{artista}' ({tipo_conteudo}): {e}")


def formatar_numero_card(value):
    """Formats large numbers for the new card standard (thousand, million)."""
    if not isinstance(value, (int, float)):
        return str(value)
    if abs(value) >= 1_000_000:
        return f'{value / 1_000_000:.1f}M'
    if abs(value) >= 1_000:
        return f'{value / 1_000:.0f}K'
    return str(int(value))


def criar_big_number_card(config, output_filename):
    """
    Creates an image of a metric card, now with customizable value positioning.
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

        # --- ALTERATION APPLIED HERE ---
        # Gets the custom position if it exists, otherwise uses the default 0.35
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
    Main function that reads the data and adds a custom position configuration
    for the Subscribers card.
    """
    try:
        data = pd.read_excel(file_path, sheet_name='Resultado')
        data2 = pd.read_excel(file_path, sheet_name='Desvio')

        metricas_config = [
            {'title': 'VIEWS', 'variant': 'primary', 'total': (data.iloc[8, 7] + data.iloc[9, 7] + data.iloc[10, 7]), 'formatter': formatar_numero_card, 'plot_index': '4a',
             'breakdown': [
                 {'label': 'VOD', 'value': data.iloc[9, 7], 'deviation': data2.iloc[9, 6], 'value_x_pos': 0.28},
                 {'label': 'LIVE', 'value': data.iloc[10, 7], 'deviation': data2.iloc[10, 6], 'value_x_pos': 0.28},
                 {'label': 'SHORT', 'value': data.iloc[8, 7], 'deviation': data2.iloc[8, 6], 'value_x_pos': 0.34}
             ]},

            {'title': 'REVENUE (USD)', 'variant': 'secondary', 'total': (data.iloc[1, 7] + data.iloc[2, 7] + data.iloc[3, 7]), 'formatter': format_currency, 'plot_index': '4b',
             'breakdown': [
                 {'label': 'VOD', 'value': data.iloc[1, 7], 'deviation': data2.iloc[1, 6], 'value_x_pos': 0.29},
                 {'label': 'LIVE', 'value': data.iloc[2, 7], 'deviation': data2.iloc[2, 6], 'value_x_pos': 0.29},
                 {'label': 'SHORT', 'value': data.iloc[3, 7], 'deviation': data2.iloc[3, 6], 'value_x_pos': 0.34}
             ]},

            # --- ALTERATION APPLIED HERE ---
            {'title': 'SUBSCRIBERS', 'variant': 'accent', 'total': data.iloc[50, 7], 'formatter': custom_format, 'plot_index': '4c',
             'breakdown': [
                 {'label': 'Gained', 'value': data.iloc[73, 7], 'deviation': data2.iloc[73, 6]},
                 # Added 'value_x_pos' key to adjust only this item
                 {'label': 'Lost', 'value': data.iloc[74, 7], 'deviation': data2.iloc[74, 6], 'value_x_pos': 0.38},
             ]},

            {'title': 'RPM (USD)', 'variant': 'primary', 'total': data.iloc[72, 7], 'formatter': lambda x: f"${dec_format(x)}", 'plot_index': '4d',
             'breakdown': [
                 {'label': 'VOD', 'value': data.iloc[69, 7], 'deviation': data2.iloc[69, 6], 'value_x_pos': 0.29},
                 {'label': 'LIVE', 'value': data.iloc[70, 7], 'deviation': data2.iloc[70, 6], 'value_x_pos': 0.29},
                 {'label': 'SHORT', 'value': data.iloc[71, 7], 'deviation': data2.iloc[71, 6], 'value_x_pos': 0.34}
             ]},

            {'title': 'IMPRESSIONS', 'variant': 'secondary', 'total': (data.iloc[5, 7] + data.iloc[6, 7] + data.iloc[65, 7]), 'formatter': formatar_numero_card, 'plot_index': '4e',
             'breakdown': [
                 {'label': 'VOD', 'value': data.iloc[5, 7], 'deviation': data2.iloc[5, 6], 'value_x_pos': 0.28},
                 {'label': 'LIVE', 'value': data.iloc[6, 7], 'deviation': data2.iloc[6, 6], 'value_x_pos': 0.28},
                 {'label': 'SHORT', 'value': data.iloc[65, 7], 'deviation': data2.iloc[65, 6], 'value_x_pos': 0.34}
             ]},

            {'title': 'WATCHTIME (HOURS)', 'variant': 'accent', 'total': data.iloc[14, 7], 'formatter': lambda x: f"{float(x):,.0f} h".replace(",", "."), 'plot_index': '4f',
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
                if item['label'] == 'Lost':
                    change_type = 'negative' if dev_percent >= 0 else 'positive'
                else:
                    change_type = 'positive' if dev_percent >= 0 else 'negative'

                # Passes the position configuration to the drawing function
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
        print(f"❌ Error in gerar_cards_detalhados for '{artista}': {e}")


def gerar_tabela_metricas(artista, tipo_conteudo, nome_arquivo, plot_index):
    try:
        data = pd.read_csv(f'dados_full/{artista}/{nome_arquivo}'); data.drop(0, inplace=True); data['Duração média da visualização'] = data['Duração média da visualização'].apply(lambda x: int(x[-5:-3]) * 60 + int(x[-2:]) if isinstance(x, str) else x); data["Data"] = data['Data'].apply(lambda i: datetime.strptime(i, '%Y-%m')); data.sort_values('Data', ascending=True, inplace=True, ignore_index=True); data["Mes"] = data.Data.apply(lambda i: i.strftime("%b"))
        nova_ordem_colunas = ['Mes', 'Vídeos publicados', 'Impressões', 'Taxa de cliques de impressões (%)', 'Visualizações', 'Duração média da visualização', 'Porcentagem visualizada média (%)', 'Inscritos', 'RPM (USD)', 'Receita estimada (USD)']; df = data[nova_ordem_colunas].copy().set_index('Mes').T
        pct_change_df = df.T.pct_change().iloc[1:]; colors = [[STYLE_CONFIG['colors']['positive'] if val >= 0 else STYLE_CONFIG['colors']['negative'] for val in row] for row in pct_change_df.values.T]; df.drop(df.columns[0], axis=1, inplace=True); alpha = 0.4; to_rgba = lambda hex_color: (*tuple(int(hex_color.lstrip('#')[i:i+2], 16) / 255 for i in (0, 2, 4)), alpha); cell_colors = [list(map(to_rgba, row)) for row in colors]; X = np.array(df).tolist()
        X[0] = [str(int(float(val))) if not pd.isna(val) else '0' for val in X[0]]; X[1] = [custom_format(val) for val in X[1]]; X[2] = [dec_format(val) for val in X[2]]; X[3] = [custom_format(val) for val in X[3]]; X[4] = [f"{dec_format(round(float(val) / 60, 2))} min" for val in X[4]]; X[5] = [dec_format(val) for val in X[5]]; X[6] = [custom_format(val) if not pd.isna(val) else '0' for val in X[6]]; X[7] = [f"${dec_format(val)}" for val in X[7]]; X[8] = [f"${dec_format(val)}" for val in X[8]]
        fig, ax = plt.subplots(figsize=STYLE_CONFIG['figsize_table_metricas']); ax.set_title(f"Growth Metrics ({tipo_conteudo.capitalize()})", **STYLE_CONFIG['font_props_title'], y=1.1); ax.axis('off'); tabela = ax.table(cellText=X, colLabels=df.columns, loc='center', cellColours=cell_colors, rowLabels=df.index); tabela.scale(1, 3.5); tabela.set_fontsize(20); fig.savefig(f'dados_full/{artista}/plots/{plot_index} - Metrics {tipo_conteudo.capitalize()}.png', transparent=STYLE_CONFIG['transparent'], bbox_inches='tight', dpi=STYLE_CONFIG['dpi']); plt.close(fig)
    except Exception as e: print(f"❌ Error in gerar_tabela_metricas for '{artista}' ({tipo_conteudo}): {e}")


def publicated_table(artista, file_path):
    """
    Generates the published videos table with advanced design, but without trend indicators (arrows).
    """
    try:
        # --- 1. DATA LOADING AND PREPARATION ---
        df_full_data = pd.read_excel(file_path.format(artista=artista), sheet_name='Resultado', index_col=0)

        # Selects only the Video and Live rows
        # The row labels are not read from the file but hardcoded for display purposes.
        df = df_full_data.iloc[11:13].copy()

        # Removes the 'Média' column and transposes to have months as columns
        if 'Média' in df.columns:
            df.drop('Média', axis=1, inplace=True)

        # Column names (months) are kept as they are.
        df.columns = pd.to_datetime(df.columns).strftime('%b')

        # --- 2. EMPTY TABLE CREATION ---
        fig, ax = plt.subplots(figsize=(10, 2.5)) # Optimized size for 2 rows
        ax.axis('off')

        tabela = ax.table(
            cellText=[[''] * len(df.columns)] * len(df.index), # Empty cells
            rowLabels=['Published Videos', 'Published Lives'], # Hardcoded English labels for display
            colLabels=df.columns,
            loc='center',
            cellLoc='right' # Aligns text to the right
        )
        tabela.auto_set_font_size(False)
        tabela.scale(1, 2.5) # Adjusts cell height

        # --- 3. STYLING AND FILLING ---
        for (row, col), cell in tabela.get_celld().items():
            cell.set_edgecolor('none')

            # Header Style
            if row == 0:
                cell.set_text_props(ha='center', color='white', **STYLE_CONFIG['table_font_props'])
                cell.set_facecolor(STYLE_CONFIG['colors']['primary_blue'])
                cell.set_height(0.2)
            # Metrics Column Style
            elif col == -1:
                cell.set_text_props(ha='left', va='center', color=STYLE_CONFIG['colors']['text_dark'], **STYLE_CONFIG['table_font_props'])
                cell.get_text().set_text(f"  {['Published Videos', 'Published Lives'][row-1]}") # English labels for display
                cell.set_facecolor('#F0F4FF')
                cell.set_edgecolor('#E5E7EB')
                cell.set_linewidth(1)
                cell.set_width(0.4)
            # Data Cells Style
            else:
                cell.set_facecolor('#FFFFFF' if row % 2 != 0 else '#F0F4FF')

                current_value = df.iloc[row-1, col]
                display_val = f"{int(current_value)}" if pd.notna(current_value) else '-'

                # CHANGE: Icon and trend logic removed from here.
                # The text is set directly, without arrows.
                cell.get_text().set_text(display_val)
                cell.get_text().set_color(STYLE_CONFIG['colors']['text_dark']) # Single color for all data
                cell.get_text().set_font_properties(STYLE_CONFIG['table_font_props']['fontproperties'])
                cell.get_text().set_size(STYLE_CONFIG['table_font_props']['size'])

        fig.suptitle('', y=0.5, **STYLE_CONFIG['font_props_title'])
        plt.tight_layout(rect=[0, 0, 1, 1])

        plt.savefig(
            f'dados_full/{artista}/plots/5 - Published.png',
            dpi=STYLE_CONFIG['dpi'],
            bbox_inches='tight',
            pad_inches=0.1,
            transparent=STYLE_CONFIG['transparent']
        )
        plt.close(fig)

    except Exception as e:
        print(f"❌ Error in publicated_table for '{artista}': {e}")


def analyze_initial_updated(artista, file_path):
    try:
        df = pd.read_excel(file_path, sheet_name='Resultado');
        meses_com_media = pd.to_datetime(df.columns[1:], errors='coerce');
        # Column names kept in Portuguese as per file
        visualizacoes = df.iloc[7, 1:].values;
        receita_total = df.iloc[0, 1:].values
        df_anomalias = pd.DataFrame({'Meses': meses_com_media.values, 'Visualizacoes': visualizacoes, 'ReceitaTotal': receita_total}).dropna(subset=['Meses'])
        df_anomalias['Visualizacoes'] = pd.to_numeric(df_anomalias['Visualizacoes']);
        df_anomalias['ReceitaTotal'] = pd.to_numeric(df_anomalias['ReceitaTotal'])

        fig, ax1 = plt.subplots(figsize=STYLE_CONFIG['figsize_standard']);
        color_receita = STYLE_CONFIG['colors']['primary_purple'];
        color_views = STYLE_CONFIG['colors']['secondary_yellow']

        ax1.set_xlabel('Month', **STYLE_CONFIG['font_props_label']);
        ax1.set_ylabel('Total Revenue ($)', color=color_receita, **STYLE_CONFIG['font_props_label'])

        linha_receita, = ax1.plot(df_anomalias['Meses'], df_anomalias['ReceitaTotal'], color=color_receita, marker='o', label='Total Revenue');

        ax1.tick_params(axis='y', labelcolor=color_receita);
        ax1.set_ylim(bottom=0); ax2 = ax1.twinx();
        ax2.set_ylabel('Views', color=color_views, **STYLE_CONFIG['font_props_label'])

        linha_visualizacao, = ax2.plot(df_anomalias['Meses'], df_anomalias['Visualizacoes'], color=color_views, marker='o', label='Views');

        ax2.tick_params(axis='y', labelcolor=color_views);
        ax2.set_ylim(bottom=0); ax2.yaxis.set_major_formatter(FuncFormatter(formatar_eixo_numeros))
        ax1.legend(handles=[linha_receita], loc='upper left');
        ax2.legend(handles=[linha_visualizacao], loc='upper right')

        plt.title('Revenue x Views', **STYLE_CONFIG['font_props_title']); plt.tight_layout();
        plt.savefig(f'dados_full/{artista}/plots/6 - Initial_Analysis.png', dpi=STYLE_CONFIG['dpi'], transparent=STYLE_CONFIG['transparent'], bbox_inches='tight');
        plt.close(fig)

    except Exception as e: print(f"❌ Error in analyze_initial_updated for '{artista}': {e}")


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
            rowLabels=['Total WatchTime', 'CPM (USD)', 'Fill Rate'],
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
                cell.get_text().set_text(f"  {['Total WatchTime', 'CPM (USD)', 'Fill Rate'][row-1]}")
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
    Generates the monetization graph with guaranteed vertical positioning for labels,
    avoiding overlap.
    """
    try:
        # --- 1. DATA LOADING (column names in Portuguese) ---
        df = pd.read_excel(file_path, sheet_name="Resultado")
        meses = pd.to_datetime(df.columns[2:], errors='coerce').dropna().to_series().dt.strftime('%b')
        receita_vod_velho = df.iloc[58, 2:].astype(float)
        receita_vod_novo = df.iloc[55, 2:].astype(float)
        receita_lives_velho = df.iloc[60, 2:].astype(float)
        receita_lives_novo = df.iloc[56, 2:].astype(float)
        receita_total = df.iloc[0, 2:].astype(float)

        # --- 2. CHART CONFIGURATION (unchanged) ---
        fig, ax1 = plt.subplots(figsize=STYLE_CONFIG['figsize_standard'])
        index = np.arange(len(meses))
        bar_width = 0.35
        chart_colors = {
            'vod_novo': '#6B7FFF', 'vod_velho': '#3157F7',
            'lives_novo': '#FCA5A5', 'lives_velho': '#EF4444',
            'total_line': '#22C55E'
        }

        # --- 3. DRAWING THE BARS (unchanged) ---
        rects_vod_velho = ax1.bar(index - bar_width/2, receita_vod_velho, bar_width, label='Old VOD Revenue', color=chart_colors['vod_velho'])
        rects_vod_novo = ax1.bar(index - bar_width/2, receita_vod_novo, bar_width, bottom=receita_vod_velho, label='New VOD Revenue', color=chart_colors['vod_novo'])
        rects_lives_velho = ax1.bar(index + bar_width/2, receita_lives_velho, bar_width, label='Old Live Revenue', color=chart_colors['lives_velho'])
        rects_lives_novo = ax1.bar(index + bar_width/2, receita_lives_novo, bar_width, bottom=receita_lives_velho, label='New Live Revenue', color=chart_colors['lives_novo'])

        # --- 4. LABEL LOGIC WITH VERTICAL POSITIONING ---

        # Helper function to add labels hierarchically
        def add_stacked_labels(bottom_rects, top_rects):
            for i, (rect_bottom, rect_top) in enumerate(zip(bottom_rects, top_rects)):
                height_bottom = rect_bottom.get_height()
                height_top = rect_top.get_height()

                # --- BOTTOM BAR LABEL ---
                if height_bottom > 0:
                    # Positions the label in the lower half of the bottom bar
                    y_pos_bottom = rect_bottom.get_y() + height_bottom * 0.01
                    ax1.text(rect_bottom.get_x() + rect_bottom.get_width() / 2., y_pos_bottom,
                             f'${height_bottom:,.0f}', ha='center', va='center', color='white',
                             bbox=dict(boxstyle='round,pad=0.3', lw=0, facecolor=rect_bottom.get_facecolor(), alpha=0.9),
                             **STYLE_CONFIG['label_font_props'])

                # --- TOP BAR LABEL ---
                if height_top > 0:
                    # Positions the label in the upper half of the top bar
                    y_pos_top = rect_top.get_y() + height_top * 1.01
                    text_color = 'white' if rect_top.get_facecolor() != chart_colors['lives_novo'] else 'black'
                    ax1.text(rect_top.get_x() + rect_top.get_width() / 2., y_pos_top,
                             f'${height_top:,.0f}', ha='center', va='center', color=text_color,
                             bbox=dict(boxstyle='round,pad=0.3', lw=0, facecolor=rect_top.get_facecolor(), alpha=0.9),
                             **STYLE_CONFIG['label_font_props'])

        # Applies the logic to both groups of bars
        add_stacked_labels(rects_vod_velho, rects_vod_novo)
        add_stacked_labels(rects_lives_velho, rects_lives_novo)

        # --- 5. FINAL CONFIGURATION AND SAVING (unchanged) ---
        ax1.set_xlabel('Months', **STYLE_CONFIG['font_props_label']); ax1.set_ylabel('Revenue by Category (USD)', **STYLE_CONFIG['font_props_label']); ax1.set_title('VOD and Live Revenue (New vs Old)', **STYLE_CONFIG['font_props_title']); ax1.set_xticks(index); ax1.set_xticklabels(meses); ax1.yaxis.set_major_formatter(FuncFormatter(formatar_eixo_numeros))
        ax2 = ax1.twinx(); ax2.plot(meses, receita_total, marker='o', color=chart_colors['total_line'], label='Total Revenue (USD)', linewidth=2.5)
        ax2.set_ylabel('Total Revenue (USD)', color=chart_colors['total_line'], **STYLE_CONFIG['font_props_label']); ax2.tick_params(axis='y', labelcolor=chart_colors['total_line']); ax2.set_ylim(bottom=0, top=receita_total.max() * 1.2); ax2.yaxis.set_major_formatter(FuncFormatter(formatar_eixo_numeros))

        lines, labels = ax1.get_legend_handles_labels(); lines2, labels2 = ax2.get_legend_handles_labels()
        fig.legend(handles=lines + lines2, labels=labels + labels2, loc='upper center', bbox_to_anchor=(0.5, 0.03), ncol=len(labels + labels2), prop={'size': 9})

        fig.tight_layout(rect=[0, 0, 1, 0.95])
        plt.savefig(f'dados_full/{artista}/plots/8 - Monetization_v2.png', dpi=STYLE_CONFIG['dpi'], transparent=STYLE_CONFIG['transparent'], bbox_inches='tight')
        plt.close(fig)

    except Exception as e:
        print(f"❌ Error in monetization_graph for '{artista}': {e}")


def revenue_per_type_chart(artista, file_path):
    """
    Generates the RPM and Revenue chart, now with gray borders on the revenue lines.
    """
    try:
        df = pd.read_excel(file_path, sheet_name="Resultado")
        meses = pd.to_datetime(df.columns[2:], errors='coerce').dropna().to_series().dt.strftime('%b')
        # Column names kept in Portuguese as per file
        rpm_vod_novo = df.iloc[21, 2:].astype(float)
        rpm_live_novo = df.iloc[22, 2:].astype(float)
        rpm_shorts_novo = df.iloc[23, 2:].astype(float)
        receita_vod = df.iloc[1, 2:].astype(float)
        receita_live = df.iloc[2, 2:].astype(float)
        receita_shorts = df.iloc[3, 2:].astype(float)

        fig, ax1 = plt.subplots(figsize=STYLE_CONFIG['figsize_standard'])
        width = 0.25
        x = np.arange(len(meses))

        # Draws RPM bars
        bars1 = ax1.bar(x - width, rpm_vod_novo, width, label='RPM VOD', color=STYLE_CONFIG['colors']['vod'])
        bars2 = ax1.bar(x, rpm_live_novo, width, label='RPM Live', color='#22C55E')
        bars3 = ax1.bar(x + width, rpm_shorts_novo, width, label='RPM Shorts', color=STYLE_CONFIG['colors']['shorts'])

        # Adds labels to the bars
        for bars in [bars1, bars2, bars3]:
            for bar in bars:
                ax1.text(
                    bar.get_x() + bar.get_width() / 2,
                    bar.get_height(),
                    f'${bar.get_height():.2f}',
                    ha='center',
                    va='bottom',
                    bbox=STYLE_CONFIG['label_bbox_props'],
                    # CORRECTION: Using the correct key from STYLE_CONFIG
                    **STYLE_CONFIG['label_font_props']
                )

        ax1.set_title('New RPM and Revenue by Content Type', **STYLE_CONFIG['font_props_title'])
        ax1.set_ylabel('New RPM (USD)', **STYLE_CONFIG['font_props_label'])
        ax1.set_xlabel('Months', **STYLE_CONFIG['font_props_label'])
        ax1.set_xticks(x)
        ax1.set_xticklabels(meses)

        # Secondary axis for revenue
        ax2 = ax1.twinx()

        # --- ALTERATION APPLIED HERE ---
        # Defines the gray border effect for the lines
        line_effect = [path_effects.withStroke(linewidth=3, foreground='#808080')]

        # Applies the border effect to each line
        ax2.plot(x, receita_vod, color=STYLE_CONFIG['colors']['vod'], marker='o', label='VOD Revenue', path_effects=line_effect)
        ax2.plot(x, receita_live, color='#22C55E', marker='o', label='Live Revenue', path_effects=line_effect)
        ax2.plot(x, receita_shorts, color=STYLE_CONFIG['colors']['shorts'], marker='o', label='Shorts Revenue', path_effects=line_effect)

        ax2.set_ylabel('Revenue (USD)', **STYLE_CONFIG['font_props_label'])
        ax2.set_ylim(bottom=0)
        ax2.yaxis.set_major_formatter(FuncFormatter(formatar_eixo_numeros))

        # Legenda e salvamento
        lines, labels = ax1.get_legend_handles_labels()
        lines2, labels2 = ax2.get_legend_handles_labels()
        fig.legend(handles=lines + lines2, labels=labels + labels2, loc='upper center', bbox_to_anchor=(0.5, 0.03), ncol=len(labels + labels2), prop={'size': 9})

        plt.tight_layout(rect=[0, 0, 1, 0.95]) # Adjusts to make space for the legend
        plt.savefig(f'dados_full/{artista}/plots/9 - Monetization by formats.png', dpi=STYLE_CONFIG['dpi'], transparent=STYLE_CONFIG['transparent'], bbox_inches='tight')
        plt.close(fig)

    except Exception as e:
        print(f"❌ Error in revenue_per_type_chart for '{artista}': {e}")


def conversion_graph(artista, file_path):
    """
    Generates the Conversion (CTR vs Views) chart with the new standardized design.
    """
    try:
        df = pd.read_excel(file_path, sheet_name='Resultado')
        meses = pd.to_datetime(df.columns[2:], errors='coerce').dropna()

        # Ensures numerical data is read correctly
        df_numeric = df.iloc[:, 2:].apply(pd.to_numeric, errors='coerce')

        # Column names kept in Portuguese as per file
        ctr_vod = df_numeric.iloc[24].astype(float)
        ctr_lives = df_numeric.iloc[25].astype(float)
        views_sem_shorts = df_numeric.iloc[7].astype(float)
        views_vod = df_numeric.iloc[9].astype(float)
        views_lives = df_numeric.iloc[10].astype(float)

        # --- 2. CHART CREATION ---
        fig, ax = plt.subplots(figsize=STYLE_CONFIG['figsize_standard'])
        index = np.arange(len(meses))
        bar_width = 0.35

        # CTR bars with new colors
        bars1 = ax.bar(index - bar_width/2, ctr_vod, bar_width, color=STYLE_CONFIG['colors']['vod'], label='CTR (%) New VOD')
        bars2 = ax.bar(index + bar_width/2, ctr_lives, bar_width, color=STYLE_CONFIG['colors']['live'], label='CTR (%) New Live')

        # Adds data labels to the bars with the new style
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

        # Adjusts axis limit to make space for labels
        ax.set_ylim(top=max(ctr_vod.max(), ctr_lives.max()) * 1.2)

        # --- 3. AXIS AND TITLE CONFIGURATION ---
        ax.set_xticks(index)
        ax.set_xticklabels(meses.strftime("%b/%Y"), rotation=45)
        ax.set_ylabel('CTR (%)', **STYLE_CONFIG['font_props_label'])
        ax.set_xlabel('Months', **STYLE_CONFIG['font_props_label'])
        ax.set_title("Click-Through Rate (CTR) and Views", **STYLE_CONFIG['font_props_title'])

        efeito_contorno = [path_effects.withStroke(linewidth=3, foreground='gray')]

        # --- 4. SECONDARY AXIS WITH VIEWS LINES ---
        ax2 = ax.twinx()
        ax2.plot(index, views_sem_shorts, color=STYLE_CONFIG['colors']['positive'], marker='o', label='Views without Shorts', path_effects=efeito_contorno)
        ax2.plot(index, views_vod, color=STYLE_CONFIG['colors']['primary_blue'], marker='s', label='VOD Views', path_effects=efeito_contorno)
        ax2.plot(index, views_lives, color=STYLE_CONFIG['colors']['primary_yellow'], marker='^', label='Live Views', path_effects=efeito_contorno)
        ax2.set_ylabel('Views', **STYLE_CONFIG['font_props_label'])
        ax2.set_ylim(0, views_sem_shorts.max() * 1.15)
        ax2.yaxis.set_major_formatter(FuncFormatter(formatar_eixo_numeros))

        # --- 5. LEGEND AND SAVING ---
        lines, labels = ax.get_legend_handles_labels()
        lines2, labels2 = ax2.get_legend_handles_labels()
        fig.legend(handles=lines + lines2, labels=labels + labels2, loc='upper center', bbox_to_anchor=(0.5, 0.03), ncol=5, prop={'size': 9})

        plt.tight_layout(rect=[0, 0, 1, 0.95])
        plt.savefig(f'dados_full/{artista}/plots/10 - Conversion.png', dpi=STYLE_CONFIG['dpi'], transparent=STYLE_CONFIG['transparent'], bbox_inches='tight')
        plt.close(fig)

    except Exception as e:
        print(f"❌ Error generating 'conversion_graph' chart: {e}")


def gerar_grafico_views(artista, file_path):
    """
    Creates a views chart with non-overlapping data labels.
    """
    try:
        # --- 1. DATA LOADING (column names in Portuguese) ---
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

        # --- 2. CHART CONFIGURATION ---
        fig, ax1 = plt.subplots(figsize=STYLE_CONFIG['figsize_standard'])
        index = np.arange(len(meses))
        bar_width = 0.25

        chart_colors = {
            'vod_novo': '#6B7FFF', 'vod_velho': '#3157F7',
            'lives_novo': '#FCA5A5', 'lives_velho': '#EF4444',
            'shorts_novo': '#a78bfa', 'shorts_velho': '#7c3aed',
            'total_line': '#22C55E'
        }

        # --- 3. DRAWING THE BARS ---
        rects_vod_velho = ax1.bar(index - bar_width, views_vod_velho, bar_width, label='Old VOD Views', color=chart_colors['vod_velho'])
        rects_vod_novo = ax1.bar(index - bar_width, views_vod_novo, bar_width, bottom=views_vod_velho, label='New VOD Views', color=chart_colors['vod_novo'])
        rects_lives_velho = ax1.bar(index, views_lives_velho, bar_width, label='Old Live Views', color=chart_colors['lives_velho'])
        rects_lives_novo = ax1.bar(index, views_lives_novo, bar_width, bottom=views_lives_velho, label='New Live Views', color=chart_colors['lives_novo'])
        rects_shorts_velho = ax1.bar(index + bar_width, views_shorts_velho, bar_width, label='Old Shorts Views', color=chart_colors['shorts_velho'])
        rects_shorts_novo = ax1.bar(index + bar_width, views_shorts_novo, bar_width, bottom=views_shorts_velho, label='New Shorts Views', color=chart_colors['shorts_novo'])

        # --- UPDATED LABEL LOGIC ---
        def add_stacked_labels(bottom_rects, top_rects):
            for rect_bottom, rect_top in zip(bottom_rects, top_rects):
                height_bottom = rect_bottom.get_height()
                height_top = rect_top.get_height()

                # Label for the bottom bar
                if height_bottom > 0:
                    # CHANGE: Position adjusted to the lower half of the bar
                    y_pos = rect_bottom.get_y() + height_bottom * 0.25
                    ax1.text(rect_bottom.get_x() + rect_bottom.get_width() / 2., y_pos,
                             formatar_eixo_numeros(height_bottom, None), ha='center', va='center', color='white',
                             bbox=dict(boxstyle='round,pad=0.3', lw=0, facecolor=rect_bottom.get_facecolor(), alpha=0.9),
                             **STYLE_CONFIG['label_font_props'])

                # Label for the top bar
                if height_top > 0:
                    # CHANGE: Position adjusted to the upper half of the bar
                    y_pos = rect_top.get_y() + height_top * 0.75
                    ax1.text(rect_top.get_x() + rect_top.get_width() / 2., y_pos,
                             formatar_eixo_numeros(height_top, None), ha='center', va='center', color='white',
                             bbox=dict(boxstyle='round,pad=0.3', lw=0, facecolor=rect_top.get_facecolor(), alpha=0.9),
                             **STYLE_CONFIG['label_font_props'])

        # Applies the logic to each group of bars
        add_stacked_labels(rects_vod_velho, rects_vod_novo)
        add_stacked_labels(rects_lives_velho, rects_lives_novo)
        add_stacked_labels(rects_shorts_velho, rects_shorts_novo)

        # --- 4. FINAL CONFIGURATION ---
        ax1.set_xlabel('Months', **STYLE_CONFIG['font_props_label']); ax1.set_ylabel('Views', **STYLE_CONFIG['font_props_label']); ax1.set_title('Views Old vs New', **STYLE_CONFIG['font_props_title']); ax1.set_xticks(index); ax1.set_xticklabels(meses); ax1.yaxis.set_major_formatter(FuncFormatter(formatar_eixo_numeros))
        ax2 = ax1.twinx(); ax2.plot(meses, views_total, marker='o', color=chart_colors['total_line'], label='Total Views', linewidth=2.5)
        ax2.set_ylabel('Total Views', color=chart_colors['total_line'], **STYLE_CONFIG['font_props_label']); ax2.tick_params(axis='y', labelcolor=chart_colors['total_line']); ax2.set_ylim(bottom=0, top=views_total.max() * 1.2); ax2.yaxis.set_major_formatter(FuncFormatter(formatar_eixo_numeros))
        lines, labels = ax1.get_legend_handles_labels(); lines2, labels2 = ax2.get_legend_handles_labels()
        fig.legend(handles=lines + lines2, labels=labels + labels2, loc='upper center', bbox_to_anchor=(0.5, 0.03), ncol=4, prop={'size': 9})
        fig.tight_layout(rect=[0, 0, 1, 0.95])
        plt.savefig(f'dados_full/{artista}/plots/19 - Views_New_vs_Old.png', dpi=STYLE_CONFIG['dpi'], transparent=STYLE_CONFIG['transparent'], bbox_inches='tight')
        plt.close(fig)

    except Exception as e:
        print(f"❌ Error in gerar_grafico_views for '{artista}': {e}")


def gerar_grafico_qualidade(artista, file_path, tipo_conteudo, plot_index):
    """
    Generates the quality chart, with backgrounds on all data labels of the bars.
    """
    try:
        df = pd.read_excel(file_path, sheet_name='Resultado')

        meses = df.columns[2:].to_list()

        if tipo_conteudo == 'vod':
            metric_rows = {'tamanho': 28, 'porcentagem': 26, 'tempo_assistido': 30, 'impressoes': 15}
            title, label_barra, label_linha = "Quality of New VODs", "Average Video Size (min)", "Average Impressions per VOD"
        elif tipo_conteudo == 'shorts':
            metric_rows = {'tamanho': 77, 'porcentagem': 78, 'tempo_assistido': 79, 'impressoes': 17}
            title, label_barra, label_linha = "Quality of New Shorts", "Average Short Size (min)", "Average Impressions per Short"
        else: # live
            metric_rows = {'tamanho': 29, 'porcentagem': 27, 'tempo_assistido': 31, 'impressoes': 16}
            title, label_barra, label_linha = "Quality of New Lives", "Average Live Size (min)", "Average Impressions per Live"

        # Data retrieval (assuming original Portuguese column names for indices 2: values)
        tamanho_medio_data = df.iloc[metric_rows['tamanho'], 2:].values
        porcentagem_media_data = df.iloc[metric_rows['porcentagem'], 2:].values
        tempo_medio_assistido_data = df.iloc[metric_rows['tempo_assistido'], 2:].values
        impressoes_media_data = df.iloc[metric_rows['impressoes'], 2:].values

        tamanho_medio = [converter_excel_time_para_minutos(val) for val in tamanho_medio_data]
        porcentagem_media = [float(val) if val is not None else 0.0 for val in porcentagem_media_data]
        tempo_medio_assistido = [converter_excel_time_para_minutos(val) for val in tempo_medio_assistido_data]
        impressoes_media = [float(val) if val is not None else 0.0 for val in impressoes_media_data]

        altura_porcentagem_visual = [t * (p / 100) for t, p in zip(tamanho_medio, porcentagem_media)]

        fig, ax1 = plt.subplots(figsize=STYLE_CONFIG['figsize_standard'])
        positions = list(range(len(meses)))

        ax1.bar(positions, tamanho_medio, color=STYLE_CONFIG['colors']['secondary_blue'], label=label_barra)
        ax1.bar(positions, altura_porcentagem_visual, color=STYLE_CONFIG['colors']['primary_blue'], label="Average Watch Time")

        for i, (x, total, parcial, perc, tempo) in enumerate(zip(positions, tamanho_medio, altura_porcentagem_visual, porcentagem_media, tempo_medio_assistido)):
            # Label inside the dark blue bar
            ax1.annotate(f"{tempo:.1f} min\n({perc:.1f}%)",
                         xy=(x, parcial / 2),
                         ha="center", va="center", color='#ffffff',
                         bbox=dict(boxstyle='round,pad=0.3', lw=0, facecolor=STYLE_CONFIG['colors']['primary_blue'], alpha=0.9),
                         **STYLE_CONFIG['label_font_props'])

            # --- ALTERATION APPLIED HERE ---
            # Added bbox for the label of the bottom bar (light blue)
            ax1.annotate(f"{total:.1f} min",
                         xy=(x, total + 0.03),
                         ha="center", va="bottom", color=STYLE_CONFIG['colors']['text_dark'],
                         bbox=dict(boxstyle='round,pad=0.3', lw=0, facecolor=STYLE_CONFIG['colors']['secondary_blue'], alpha=0.9),
                         **STYLE_CONFIG['label_font_props'])

        # The rest of the function remains the same
        ax1.set_ylabel(label_barra, **STYLE_CONFIG['font_props_label'])
        ax1.set_ylim(0, max(tamanho_medio) * 1.3 if tamanho_medio else 1)
        ax1.set_xticks(positions)
        ax1.set_xticklabels(meses)

        ax2 = ax1.twinx()
        ax2.plot(positions, impressoes_media, color=STYLE_CONFIG['colors']['primary_purple'], marker='o', linewidth=2, label=label_linha)
        ax2.set_ylabel("Impressions per Upload", color=STYLE_CONFIG['colors']['primary_purple'], **STYLE_CONFIG['font_props_label'])
        ax2.tick_params(axis='y', colors=STYLE_CONFIG['colors']['primary_purple'])
        ax2.set_ylim(0, max(impressoes_media) * 1.2 if impressoes_media else 1)
        ax2.yaxis.set_major_formatter(FuncFormatter(formatar_eixo_numeros))

        lines1, labels1 = ax1.get_legend_handles_labels()
        lines2, labels2 = ax2.get_legend_handles_labels()
        ax2.legend(lines1 + lines2, labels1 + labels2, loc=(0.075, 1), fontsize=9, ncol=len(labels1 + labels2))

        fig.suptitle(title, x=0.035, ha='left', **STYLE_CONFIG['font_props_subtitle'], y=0.73, rotation=90)
        plt.tight_layout(rect=[1, 1, 1, 0.96])
        fig.savefig(f'dados_full/{artista}/plots/{plot_index} - Quality_{tipo_conteudo.capitalize()}.png', dpi=STYLE_CONFIG['dpi'], bbox_inches="tight", transparent=STYLE_CONFIG['transparent'])
        plt.close(fig)

    except Exception as e:
        print(f"❌ Error in gerar_grafico_qualidade for '{artista}' ({tipo_conteudo}): {e}")


def dfOrigem(nomeArquivo1, nomeArquivo2):
    df1 = pd.read_csv(nomeArquivo1 + ".csv"); df2 = pd.read_csv(nomeArquivo2 + ".csv"); df = pd.concat([df1, df2], ignore_index=True); df['Data'] = pd.to_datetime(df['Data']); df.sort_values('Data', inplace=True); df["Mês"] = df['Data'].dt.strftime('%Y-%m');
    # Keep original Portuguese column names for filtering and grouping
    origensImportantes = ("Recursos de navegação","Vídeos sugeridos","Páginas do canal","Externa","Notificações","Pesquisa do YouTube","Playlists","Publicidade no YouTube");
    df = df[['Origem do tráfego', 'Visualizações', 'Mês']].groupby(by=['Mês', 'Origem do tráfego']).sum().reset_index();
    b = df[~df['Origem do tráfego'].isin(origensImportantes)].groupby(by=['Mês']).sum(numeric_only=True);
    b["Origem do tráfego"] = 'Outros'; # Keep 'Outros' for internal logic then translate for plot
    b.reset_index(inplace=True); df = pd.concat([b, df[df['Origem do tráfego'].isin(origensImportantes)]]);
    total = df.groupby('Origem do tráfego')['Visualizações'].sum().sort_values(ascending=False).index.tolist();
    df = df.pivot(index="Mês", columns="Origem do tráfego", values="Visualizações").fillna(0);
    df = df.reset_index().sort_values(by="Mês").reset_index(drop=True)
    return df, total


def traficSorce_graph(df, x_labels, total, title, color_map, artista):
    try:
        fig, ax = plt.subplots(figsize=STYLE_CONFIG['figsize_wide']); bottom = np.zeros(len(x_labels))
        # Map Portuguese origins to English for the legend
        translated_origins = {
            "Recursos de navegação": "Browse features",
            "Vídeos sugeridos": "Suggested videos",
            "Páginas do canal": "Channel pages",
            "Externa": "External",
            "Notificações": "Notifications",
            "Pesquisa do YouTube": "Youtube",
            "Playlists": "Playlists",
            "Publicidade no YouTube": "YouTube advertising",
            "Outros": "Others"
        }
        for n, origem in enumerate(total):
            if origem in df.columns:
                y_values = df[origem].values
                display_label = translated_origins.get(origem, origem) # Get translated label or use original
                ax.bar(x=x_labels, height=y_values, color=color_map[n % len(color_map)], bottom=bottom, label=display_label);
                bottom += y_values
        ax.legend(loc='upper left', bbox_to_anchor=(1, 1)); ax.yaxis.set_major_formatter(FuncFormatter(formatar_eixo_numeros)); ax.set_title(title.title(), **STYLE_CONFIG['font_props_title']); ax.tick_params(axis='x', rotation=45); ax.set_ylabel('Views', **STYLE_CONFIG['font_props_label']); plt.tight_layout()
        plt.savefig(f'dados_full/{artista}/plots/13 - Traffic_Source.png', dpi=STYLE_CONFIG['dpi'], transparent=STYLE_CONFIG['transparent']); plt.close(fig)
    except Exception as e: print(f"❌ Error in traficSorce_graph for '{artista}': {e}")


def subscription_growth(artista, file_path):
    """
    Generates the subscription growth chart with adjusted spacing,
    labels with background, and improved vertical positioning.
    """
    try:
        # --- 1. DATA LOADING AND PREPARATION (column names in Portuguese) ---
        df = pd.read_excel(file_path, sheet_name="Resultado")
        meses = pd.to_datetime(df.columns[2:], errors='coerce').dropna().to_series().dt.strftime('%b')
        # Original Portuguese column names for data extraction
        insc_novo_vod = df.iloc[41, 2:].astype(float)
        insc_novo_live = df.iloc[42, 2:].astype(float)
        insc_novo_shorts = df.iloc[43, 2:].astype(float)
        insc_velho_vod = df.iloc[44, 2:].astype(float)
        insc_velho_live = df.iloc[45, 2:].astype(float)
        insc_velho_shorts = df.iloc[46, 2:].astype(float)
        insc_total = df.iloc[50, 2:].astype(float)

        # --- 2. CHART CONFIGURATION ---
        fig, ax1 = plt.subplots(figsize=STYLE_CONFIG['figsize_standard'])
        width = 0.25
        x = np.arange(len(meses))

        # ALTERATION 1: Increased side spacing between bar groups
        spacing_factor = 1.2 # Increased from 1.1 to 1.2 for more space

        # --- 3. DRAWING THE BARS WITH NEW SPACING ---
        # The bars now use the new 'spacing_factor' to move further apart
        bars_vod_novo = ax1.bar(x - width * spacing_factor, insc_novo_vod, width, label='New VOD', color=STYLE_CONFIG['colors']['vod'])
        bars_vod_velho = ax1.bar(x - width * spacing_factor, insc_velho_vod, width, bottom=insc_novo_vod, label='Old VOD', color=STYLE_CONFIG['colors']['secondary_blue'])

        bars_live_novo = ax1.bar(x, insc_novo_live, width, label='New Live', color=STYLE_CONFIG['colors']['live'])
        bars_live_velho = ax1.bar(x, insc_velho_live, width, bottom=insc_novo_live, label='Old Live', color=STYLE_CONFIG['colors']['accent_purple'])

        bars_shorts_novo = ax1.bar(x + width * spacing_factor, insc_novo_shorts, width, label='New Shorts', color=STYLE_CONFIG['colors']['shorts'])
        bars_shorts_velho = ax1.bar(x + width * spacing_factor, insc_velho_shorts, width, bottom=insc_novo_shorts, label='Old Shorts', color=f"{STYLE_CONFIG['colors']['shorts']}80")

        # --- 4. LABEL LOGIC WITH BACKGROUND AND VERTICAL SPACING ---
        def add_stacked_labels(rects_novo, rects_velho):
            for rect_n, rect_v in zip(rects_novo, rects_velho):
                height_novo = rect_n.get_height()
                height_velho = rect_v.get_height()

                # Label for the TOP bar ("New")
                if height_novo > 0:
                    # ALTERATION 2: Top label position adjusted to 80% of its height, pushing it up
                    y_pos_novo = rect_n.get_y() + height_novo * 0.80
                    ax1.text(rect_n.get_x() + rect_n.get_width() / 2, y_pos_novo,
                             f"{formatar_eixo_numeros(height_novo, None)}", ha='center', va='center',
                             color='white', bbox=dict(boxstyle='round,pad=0.3', lw=0, facecolor=rect_n.get_facecolor(), alpha=0.9),
                             **STYLE_CONFIG['label_font_props'])

                # Label for the BOTTOM bar ("Old")
                if height_velho > 0:
                    # ALTERATION 3: Bottom label position adjusted to 20% of its height, pushing it down
                    y_pos_velho = rect_v.get_y() + height_velho * 0.20
                    ax1.text(rect_v.get_x() + rect_v.get_width() / 2, y_pos_velho,
                             f"{formatar_eixo_numeros(height_velho, None)}", ha='center', va='center',
                             color='white', bbox=dict(boxstyle='round,pad=0.3', lw=0, facecolor=rect_v.get_facecolor(), alpha=0.9),
                             **STYLE_CONFIG['label_font_props'])

        # Applies the logic to each group of bars
        add_stacked_labels(bars_vod_novo, bars_vod_velho)
        add_stacked_labels(bars_live_novo, bars_live_velho)
        add_stacked_labels(bars_shorts_novo, bars_shorts_velho)

        # --- 5. FINAL CONFIGURATION AND SAVING (unchanged) ---
        ax1.set_title('Subscribers by Content Type', **STYLE_CONFIG['font_props_title'])
        ax1.set_ylabel('Number of Subscribers', **STYLE_CONFIG['font_props_label'])
        ax1.set_xlabel('Months', **STYLE_CONFIG['font_props_label'])
        ax1.set_xticks(x)
        ax1.set_xticklabels(meses)
        ax1.yaxis.set_major_formatter(FuncFormatter(formatar_eixo_numeros))

        ax2 = ax1.twinx()
        ax2.plot(x, insc_total, color=STYLE_CONFIG['colors']['positive'], marker='o', linestyle='-', label='Total Subscribers')
        ax2.set_ylabel('Total Subscribers', **STYLE_CONFIG['font_props_label'], color=STYLE_CONFIG['colors']['positive'])
        ax2.tick_params(axis='y', labelcolor=STYLE_CONFIG['colors']['positive'])
        ax2.set_ylim(bottom=0, top=max(insc_total) * 1.2 if not insc_total.empty else 1)
        ax2.yaxis.set_major_formatter(FuncFormatter(formatar_eixo_numeros))

        fig.legend(loc='upper right', bbox_to_anchor=(0.95, 0.03), ncol=7)
        plt.tight_layout(rect=[0, 0, 0.9, 1])
        plt.savefig(f'dados_full/{artista}/plots/14 - Subscribers_by_Content_Type.png', dpi=STYLE_CONFIG['dpi'], transparent=STYLE_CONFIG['transparent'], bbox_inches='tight')
        plt.close(fig)

    except Exception as e:
        print(f"❌ Error in subscription_growth for '{artista}': {e}")


def gerar_grafico_engajamento_tipo(artista, file_path, tipo_conteudo, plot_index):
    """
    Gera o gráfico de engajamento com cores de barra padronizadas e uma linha estilizada.
    """
    try:
        df = pd.read_excel(file_path, sheet_name='Resultado')
        meses = pd.to_datetime(df.columns[2:], errors='coerce').dropna()

        if tipo_conteudo == 'vod':
            engajamento_row, impressoes_row, title = 63, 15, "VOD Engagement"
        elif tipo_conteudo == 'shorts':
            engajamento_row, impressoes_row, title = 76, 17, "Shorts Engagement"
        else: # live
            engajamento_row, impressoes_row, title = 64, 16, "Lives Engagement"

        # Data retrieval (assuming original Portuguese column names for values)
        engajamento = df.iloc[engajamento_row, 2:].astype(float)
        impressoes_media = df.iloc[impressoes_row, 2:].astype(float)

        fig, ax = plt.subplots(figsize=(10, 6))
        index = np.arange(len(meses))

        # --- ALTERATION 1: Standardized bar color ---
        # All bars now use 'primary_blue' color
        bars = ax.bar(index, engajamento, 0.4, color=STYLE_CONFIG['colors']['primary_blue'], label=f'Average {tipo_conteudo.capitalize()} Engagement')

        for bar in bars:
            yval = bar.get_height()
            ax.text(bar.get_x() + bar.get_width() / 2, yval, f'{yval:.1f}',
                    ha='center', va='bottom', bbox=STYLE_CONFIG['label_bbox_props'],
                    **STYLE_CONFIG['label_font_props'])

        ax.set_xticks(index)
        ax.set_xticklabels(meses.strftime("%b/%Y"), rotation=0)
        ax.set_ylabel('Average Engagement', **STYLE_CONFIG['font_props_label'])
        ax.set_xlabel('Months', **STYLE_CONFIG['font_props_label'])

        ax2 = ax.twinx()

        # --- ALTERATION 2: Line style ---
        # Sets the line color to be the same as 'shorts'
        cor_linha = STYLE_CONFIG['colors']['shorts']
        # Defines the gray outline effect
        efeito_contorno = [path_effects.withStroke(linewidth=3, foreground='gray')]

        # Applies the new color and outline to the line
        ax2.plot(index, impressoes_media, color=cor_linha, marker='o',
                 label=f"Average Impressions per New {tipo_conteudo.capitalize()}",
                 path_effects=efeito_contorno)

        ax2.set_ylabel('Impressions', **STYLE_CONFIG['font_props_label'], color=cor_linha)
        ax2.tick_params(axis='y', labelcolor=cor_linha)
        ax2.set_ylim(0, max(impressoes_media) * 1.2 if not impressoes_media.empty else 1)
        ax2.yaxis.set_major_formatter(FuncFormatter(formatar_eixo_numeros))

        fig.suptitle(title, x=0.055, y=0.475, ha='left', va='center', rotation='vertical', **STYLE_CONFIG['font_props_title'])

        plt.subplots_adjust(left=0.15)
        fig.legend(loc='upper right', bbox_to_anchor=(0.95, 0.95), ncol=2)

        plt.savefig(f'dados_full/{artista}/plots/{plot_index} - Engagement_{tipo_conteudo.capitalize()}.png', dpi=STYLE_CONFIG['dpi'], transparent=STYLE_CONFIG['transparent'], bbox_inches='tight')
        plt.close(fig)

    except Exception as e:
        print(f"❌ Error in gerar_grafico_engajamento_tipo for '{artista}': {e}")


def generate_comunity_chart(artista):
    """
    Generates a bar and line chart to analyze community posts and impressions.
    The function uses the most recent date from the 'total.csv' file as a reference
    to filter the last 6 months of data.
    """
    try:
        # --- 1. Load and Prepare Data (total.csv) ---
        data_meses = pd.read_csv(f'dados_full/{artista}/total.csv').drop(0)
        data_meses["Data"] = pd.to_datetime(data_meses["Data"])

        # --- 2. Logic to Define the 6-Month Period ---
        # Find the latest date (which pandas interprets as the 1st of the month)
        latest_date_start_of_month = data_meses["Data"].max()
        # Adjust the date to be the LAST day of that month to ensure the whole month is included.
        latest_date = latest_date_start_of_month + pd.offsets.MonthEnd(0)
        
        # Define the lower bound of the 6-month window
        first_day_of_latest_month = latest_date.replace(day=1)
        lower_bound_date = first_day_of_latest_month - pd.DateOffset(months=5)

        # Filter the 'total.csv' data for the desired period
        data_meses_filtrada = data_meses[
            (data_meses["Data"] >= lower_bound_date) &
            (data_meses["Data"] <= latest_date)
        ].copy()

        # --- 3. Load and Prepare Data (comunidade.csv) ---
        data_comu = pd.read_csv(f'dados_full/{artista}/comunidade.csv')
        data_comu["Horário de publicação da postagem"] = pd.to_datetime(
            data_comu["Horário de publicação da postagem"], errors="coerce"
        )
        data_comu.dropna(subset=["Horário de publicação da postagem"], inplace=True)

        # Filter the 'comunidade.csv' data for the same period
        data_comu_filtrada = data_comu[
            (data_comu["Horário de publicação da postagem"] >= lower_bound_date) &
            (data_comu["Horário de publicação da postagem"] <= latest_date)
        ].copy()

        # --- 4. Group and Merge Data ---
        # Extract the month and count the number of publications
        data_comu_filtrada["Mes"] = data_comu_filtrada["Horário de publicação da postagem"].dt.month
        # The new column name is now in English
        comuCount = data_comu_filtrada.groupby("Mes").size().reset_index(name='Community Posts')

        # Prepare the final DataFrame for the plot
        df_graf_com = data_meses_filtrada.copy()
        df_graf_com["Mes"] = df_graf_com["Data"].dt.month
        df_graf_com = df_graf_com.merge(comuCount, on="Mes", how="left").fillna(0)

        # --- 5. Sort and Format the DataFrame for Plotting ---
        df_graf_com['AnoMesTemp'] = df_graf_com['Data'].dt.to_period('M')
        df_graf_com = df_graf_com.sort_values(by='AnoMesTemp').drop(columns='AnoMesTemp')

        # Map month number to its English abbreviation
        month_order = {
            1: "Jan", 2: "Feb", 3: "Mar", 4: "Apr", 5: "May", 6: "Jun",
            7: "Jul", 8: "Aug", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dec"
        }
        df_graf_com['MesAbreviado'] = df_graf_com['Mes'].map(month_order)
        df_graf_com.set_index('MesAbreviado', inplace=True)
        
        # --- 6. Chart Generation ---
        fig, ax1 = plt.subplots(figsize=STYLE_CONFIG['figsize_standard'])
        ax2 = ax1.twinx()

        x = np.arange(len(df_graf_com.index))
        
        # Bar Chart: Community Posts (using the English column name and label)
        bars = ax1.bar(x, df_graf_com["Community Posts"], width=0.4, color=STYLE_CONFIG['colors']['primary_blue'], label="Community Posts")
        
        # Line Chart: Post Impressions (using the original column name but an English label)
        contour_effect = [path_effects.withStroke(linewidth=3, foreground='gray')]
        ax2.plot(x, df_graf_com["Impressões da postagem"], marker="o", color=STYLE_CONFIG['colors']['primary_yellow'], label="Post Impressions", path_effects=contour_effect)

        # Add data labels to the bars
        for bar in bars:
            height = bar.get_height()
            ax1.text(bar.get_x() + bar.get_width() / 2, height, f'{int(height)}', ha='center', va='bottom', bbox=STYLE_CONFIG['label_bbox_props'], **STYLE_CONFIG['label_font_props'])

        # --- 7. Styling and Final Chart Settings ---
        # Set axis limits for better visualization
        ax1.set_ylim(0, max(df_graf_com["Community Posts"]) * 1.2 if not df_graf_com["Community Posts"].empty and df_graf_com["Community Posts"].max() > 0 else 1)
        ax2.set_ylim(0, max(df_graf_com["Impressões da postagem"]) * 1.2 if not df_graf_com["Impressões da postagem"].empty and df_graf_com["Impressões da postagem"].max() > 0 else 1)

        # Axis labels and title in English
        ax1.set_xlabel("Month", **STYLE_CONFIG['font_props_label'])
        ax1.set_ylabel("Posts", **STYLE_CONFIG['font_props_label'])
        ax2.set_ylabel("Community Impressions", **STYLE_CONFIG['font_props_label'])
        plt.title("Community: Posts vs. Impressions", **STYLE_CONFIG['font_props_title'])

        # X-axis settings
        ax1.set_xticks(x)
        ax1.set_xticklabels(df_graf_com.index, rotation=45, ha='right')
        ax1.grid(axis="y", linestyle="--", alpha=0.5)

        # Legend and Layout
        fig.legend(loc='lower center', bbox_to_anchor=(0.5, 0.05), ncol=2)
        plt.tight_layout(rect=[0, 0.1, 1, 1])
        
        # Save the chart (uncomment to use)
        plt.savefig(f'dados_full/{artista}/plots/17 - Community.png', dpi=STYLE_CONFIG['dpi'], transparent=STYLE_CONFIG['transparent'], bbox_inches='tight');
        plt.close(fig)

    except Exception as e:
        # Error message in English
        print(f"❌ Error in generate_engagement_chart for '{artista}': {e}")


def gerar_tabela_inscritos_avancada(artista):
    """
    Generates the final subscriber table, with advanced design, 7 to 6 months logic,
    trend colors and now with arrow icons (↑/↓).
    """
    try:
        # --- 1. DATA LOADING AND PREPARATION ---
        df_full = pd.read_csv(f'dados_full/{artista}/total.csv').drop(0)
        df_full["Data"] = pd.to_datetime(df_full["Data"])
        df_full.sort_values('Data', ascending=True, inplace=True, ignore_index=True)

        if len(df_full) >= 7:
            df_full = df_full.tail(7).reset_index(drop=True)

        with open(f"dados_full/{artista}/sub.txt", "r", encoding='latin-1') as f:
            inscant = int(''.join(f.readline().split('.')))

        # Access columns using ORIGINAL PORTUGUESE NAMES from df_full
        inscant_inicial_periodo = inscant - df_full['Inscritos'].sum()
        total_acumulado = (df_full['Inscritos'].cumsum() + inscant_inicial_periodo)

        ganhos = pd.to_numeric(df_full['Inscrições obtidas'], errors='coerce').fillna(0)
        perdidos = pd.to_numeric(df_full['Inscrições perdidas'], errors='coerce').fillna(0)

        saldo = ganhos - perdidos
        crescimento_pct = (total_acumulado.pct_change() * 100).fillna(0).round(2)

        # --- 2. FINAL DATAFRAME ASSEMBLY (Use English names for display in df_display) ---
        df_display = pd.DataFrame({
            "Total Subscribers": total_acumulado.tail(6).values,
            "Growth (%)": crescimento_pct.tail(6).values,
            "Subscribers Gained": ganhos.tail(6).values,
            "Subscribers Lost": perdidos.tail(6).values,
            "Subscriber Balance": saldo.tail(6).values
        })
        df_display = df_display.T # Transpose to make metric names row labels
        df_display.columns = pd.to_datetime(df_full['Data'].tail(6)).dt.strftime('%b')

        # df_full_trend is created with English names for consistent logic in the loop
        df_full_trend = pd.DataFrame({
            "Total Subscribers": total_acumulado,
            "Growth (%)": crescimento_pct,
            "Subscribers Gained": ganhos,
            "Subscribers Lost": perdidos,
            "Subscriber Balance": saldo
        }).T
        df_full_trend.columns = pd.to_datetime(df_full['Data']).dt.strftime('%b') # Keep original column names for months

        # --- 3. TABLE CREATION AND STYLIZATION ---
        fig, ax = plt.subplots(figsize=(12, 4.5))
        ax.axis('off')

        tabela = ax.table(
            cellText=[[''] * len(df_display.columns)] * len(df_display.index),
            rowLabels=df_display.index, # These are already English from df_display
            colLabels=df_display.columns,
            loc='center', cellLoc='right'
        )
        tabela.auto_set_font_size(False); tabela.scale(1, 2.8)

        for (row, col_display), cell in tabela.get_celld().items():
            col_full = col_display + 1
            cell.set_edgecolor('none')
            if row == 0: # Header row
                cell.set_text_props(ha='center', color='white', **STYLE_CONFIG['table_font_props']); cell.set_facecolor(STYLE_CONFIG['colors']['primary_blue'])
            elif col_display == -1: # Row labels column (metric names)
                cell.set_text_props(ha='left', va='center', color=STYLE_CONFIG['colors']['text_dark'], **STYLE_CONFIG['table_font_props']);
                # The row labels (df_display.index) are already English, use them directly
                display_row_label = df_display.index[row-1]
                cell.get_text().set_text(f"  {display_row_label}");
                cell.set_facecolor('#F0F4FF'); cell.set_edgecolor('#E5E7EB'); cell.set_linewidth(1); cell.set_width(0.4)
            else: # Data cells
                cell.set_facecolor('#FFFFFF' if row % 2 == 0 else '#F0F4FF')

                # Use the English metric name from df_full_trend's index for logic and display
                metric_name_for_logic = df_full_trend.index[row-1]
                current_value = df_full_trend.iloc[row-1, col_full];
                # Ensure prev_value is accessed safely, especially for the first data column
                prev_value = df_full_trend.iloc[row-1, col_full - 1] if col_full - 1 >= 0 else None

                trend = get_trend(current_value, prev_value); # get_trend handles None
                color = STYLE_CONFIG['colors']['text_dark']

                # CORRECTED: Use English metric names for color logic
                if metric_name_for_logic == "Subscribers Lost": # If 'Subscribers Lost'
                    if trend == "up": color = STYLE_CONFIG['colors']['negative'] # More lost is negative
                    elif trend == "down": color = STYLE_CONFIG['colors']['positive'] # Less lost is positive
                elif metric_name_for_logic in ["Total Subscribers", "Growth (%)", "Subscribers Gained", "Subscriber Balance"]: # For these metrics
                     if trend == "up": color = STYLE_CONFIG['colors']['positive'] # Up is positive
                     elif trend == "down": color = STYLE_CONFIG['colors']['negative'] # Down is negative

                icon = ''
                if trend == "up": icon = '↑ '
                elif trend == "down": icon = '↓ '

                display_val = ""
                if pd.notna(current_value):
                    if metric_name_for_logic == "Growth (%)": display_val = f"{current_value:,.2f}%".replace(",", "X").replace(".", ",").replace("X", ".")
                    else: display_val = f"{int(current_value):,}".replace(",", ".")

                # The icon is added to the final text
                cell.get_text().set_text(f"{icon}{display_val}")
                cell.get_text().set_color(color); cell.get_text().set_font_properties(STYLE_CONFIG['table_font_props']['fontproperties']); cell.get_text().set_size(STYLE_CONFIG['table_font_props']['size'])

        fig.suptitle("", y=0.5, **STYLE_CONFIG['font_props_title']) # Title remains empty as before
        plt.tight_layout(rect=[0, 0, 1, 1])

        plt.savefig(f'dados_full/{artista}/plots/18 - Subscribers Table.png', transparent=STYLE_CONFIG['transparent'], bbox_inches="tight", dpi=STYLE_CONFIG['dpi'])
        plt.close(fig)

    except Exception as e:
        print(f" Error in generate_advanced_subscriber_table for '{artista}': {e}") # Translated error message


def run(artista):     
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
    generate_comunity_chart(artista)
    gerar_tabela_inscritos_avancada(artista)
    gerar_grafico_views(artista, file_path_4_1)

    # Corrigindo o print com emoji
    print(f"OK - Reports for '{artista}' completed.")


if __name__ == "__main__":
    # Pega o nome do artista do argumento passado pelo main.py
    if len(sys.argv) < 2:
        print("Error: No artist provided. This script must be called by main.py")
        sys.exit(1) # Sai com erro

    artista_argumento = sys.argv[1]

    # Executa a função run APENAS para esse artista
    run(artista_argumento)
