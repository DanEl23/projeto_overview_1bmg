import os
import pandas as pd



colunas_data = {
    'Impressões':'Impressões (soma)',
    'Visualizações':'Visualizações (soma)',
    'Taxa de cliques de impressões (%)':'CTR (%) (média)',
    'Duração média da visualização':'Tempo Médio Assistido (média)',
    'Porcentagem visualizada média (%)':'Porcentagem Média Assistida (média)',
    'RPM (BRL)':'RPM (R$) (média)',
    'Marcações "Não gostei"':'Porcentagem de "Não Gostei" (média)',
    'Inscritos':'Saldo de Inscritos (média)',
}

colunas_conteudo = {
    'Vídeos publicados':'Número de conteúdos',
    'Impressões':'Impressões',
    'Visualizações':'Visualizações',
    'Taxa de cliques de impressões (%)':'CTR (%)',
    'Duração média da visualização':'Tempo Médio Assistido',
    'Porcentagem visualizada média (%)':'Porcentagem Média Assistida',
    'RPM (BRL)':'RPM (R$)',
    'Marcações "Não gostei"':'Porcentagem de "Não Gostei"',
    'Inscritos':'Saldo de Inscritos',
    'Minutagem média':'Tamanho médio do vídeo (min)'
    
}

colunas_receita = {
    'Receita estimada (BRL)_x':'Total',
    'Receita estimada (BRL)_y':'Conteúdo novo (Soma)',
    'Conteúdo Velho':'Conteúdo velho (Soma)'

}

artista = ''
with open('dados_full/exports.txt') as f:
    lines = f.readlines()
lines = [i.rstrip() for i in lines]


for artista in lines:
    f = open("dados_full/"+artista+"/sub.txt", "r")
    subs = int(''.join((f.readline()).split('.')))
    f.close()

    data_1 = pd.read_csv(f'dados_full/{artista}/subs.csv')
    data_1 = data_1.drop(data_1.index[0])
    data_1['Data'] = pd.to_datetime(data_1['Data'])
    data_1['Data'] = data_1['Data'].dt.to_period('M')
    data = data_1.copy()
    data = data[['Data','Impressões', 'Visualizações','Taxa de cliques de impressões (%)', 'Duração média da visualização', 'Porcentagem visualizada média (%)', 'RPM (BRL)', 'Marcações "Não gostei"', 'Inscritos']]
    data['Marcações "Não gostei"'] = data['Marcações "Não gostei"'] / data['Visualizações']
    data['Duração média da visualização'] = pd.to_timedelta(data['Duração média da visualização'])


    data_impressoes = data['Impressões'].mean()
    data_visualizacoes = data['Visualizações'].mean()
    data_ctr = data['Taxa de cliques de impressões (%)'].mean()
    data_tempo_assistido = data['Duração média da visualização'].mean()
    data_tempo_assistido = data_tempo_assistido.seconds
    data_tempo_assistido = '{:02d}:{:02d}'.format(data_tempo_assistido // 60, data_tempo_assistido % 60)
    data_porcentagem_assistida = data['Porcentagem visualizada média (%)'].mean()
    data_rpm = data['RPM (BRL)'].mean()
    data_nao_gostei = data['Marcações "Não gostei"'].mean()


    data['Duração média da visualização'] = pd.to_timedelta(data['Duração média da visualização']).dt.components['minutes'].astype(str).str.zfill(2) + ':' + pd.to_timedelta(data['Duração média da visualização']).dt.components['seconds'].astype(str).str.zfill(2)
    data.sort_values(by='Data', inplace=True, ascending=False)
    data_cumulative = subs - data['Inscritos'].cumsum()
    data['Inscritos'] = data_cumulative.shift(1).fillna(subs).astype(int)
    data_inscritos = data['Inscritos'].mean()
    data.sort_values(by='Data', inplace=True, ascending=True)
    data.rename(columns=colunas_data, inplace=True)
    data = data.T
    data.insert(0, 'Média (últimos 6 meses)', ['',data_impressoes, data_visualizacoes, data_ctr, data_tempo_assistido, data_porcentagem_assistida, data_rpm, data_nao_gostei, data_inscritos])
    data.insert(0, 'Mês Atual', data.iloc[:,-1])
    data.reset_index(inplace=True)
    data

    conteudo_1 = pd.read_csv(f'dados_full/{artista}/videos.csv')
    conteudo_1['Horário de publicação do vídeo'] = pd.to_datetime(conteudo_1['Horário de publicação do vídeo'])
    conteudo_1['Data'] = conteudo_1['Horário de publicação do vídeo'].dt.to_period('M')
    conteudo_1['Duração média da visualização'] = pd.to_timedelta(conteudo_1['Duração média da visualização'])
    conteudo_1['Duração média da visualização'] = conteudo_1['Duração média da visualização'].apply(lambda x: x.total_seconds())
    conteudo_1.set_index('Data', inplace=True)
    conteudo = conteudo_1.copy()
    conteudo_1 = conteudo.groupby(pd.Grouper(freq='M')).agg({'Receita estimada (BRL)' : 'sum'})
    conteudo_1 = conteudo_1.reset_index()


    conteudo = conteudo.groupby(pd.Grouper(freq='M')).agg({'Impressões':'sum', 'Visualizações':'sum', 'Taxa de cliques de impressões (%)':'mean', 'Duração média da visualização':'mean', 'Porcentagem visualizada média (%)':'mean', 'RPM (BRL)':'mean', 'Marcações "Não gostei"':'mean', 'Inscritos':'sum'})
    conteudo['Minutagem média'] = (conteudo['Duração média da visualização'] * 100) / conteudo['Porcentagem visualizada média (%)']
    conteudo['Marcações "Não gostei"'] = conteudo['Marcações "Não gostei"'] / conteudo['Visualizações']

    conteudo_impressoes = conteudo['Impressões'].mean()
    conteudo_visualizacoes = conteudo['Visualizações'].mean()
    conteudo_ctr = conteudo['Taxa de cliques de impressões (%)'].mean()
    conteudo_tempo_assistido = conteudo['Duração média da visualização'].mean()
    conteudo_tempo_assistido = '{:02d}:{:02d}'.format(int(conteudo_tempo_assistido // 60), int(conteudo_tempo_assistido % 60))
    conteudo_porcentagem_assistida = conteudo['Porcentagem visualizada média (%)'].mean()
    conteudo_rpm = conteudo['RPM (BRL)'].mean()
    conteudo_nao_gostei = conteudo['Marcações "Não gostei"'].mean()
    conteudo_inscritos = conteudo['Inscritos'].mean()



    conteudo['Duração média da visualização'] = conteudo['Duração média da visualização'].apply(lambda x: int(x))
    conteudo['Duração média da visualização'] = conteudo['Duração média da visualização'].apply(lambda x: pd.to_timedelta(x, unit = 's'))
    conteudo['Duração média da visualização'] = pd.to_timedelta(conteudo['Duração média da visualização']).dt.components['minutes'].astype(str).str.zfill(2) + ':' + pd.to_timedelta(conteudo['Duração média da visualização']).dt.components['seconds'].astype(str).str.zfill(2)
    conteudo['Minutagem média'] = conteudo['Minutagem média'] / 60
    conteudo_minutagem = conteudo['Minutagem média'].mean()
    conteudo = conteudo.reset_index()
    conteudo = pd.merge(conteudo, data_1[['Data', 'Vídeos publicados']], on = 'Data', how = 'left')
    conteudo_pubs = conteudo['Vídeos publicados'].mean()
    cols = conteudo.columns.tolist()
    cols = cols[:1] + cols[-1:] + cols[1:-1]
    conteudo = conteudo[cols]
    conteudo.sort_values(by = 'Data', inplace = True)
    conteudo.rename(columns=colunas_conteudo, inplace=True)
    conteudo = conteudo.T
    conteudo.insert(0, 'Média (últimos 6 meses)', ['',conteudo_pubs,conteudo_impressoes, conteudo_visualizacoes, conteudo_ctr, conteudo_tempo_assistido, conteudo_porcentagem_assistida, conteudo_rpm, conteudo_nao_gostei, conteudo_inscritos, conteudo_minutagem])
    conteudo.insert(0, 'Mês Atual', conteudo.iloc[:,-1])
    conteudo.reset_index(inplace=True)
    conteudo

    receita = data_1[['Data', 'Receita estimada (BRL)']]
    receita = pd.merge(receita, conteudo_1[['Data', 'Receita estimada (BRL)']], on = 'Data', how = 'left')
    receita['Conteúdo Velho'] = receita['Receita estimada (BRL)_x'] - receita['Receita estimada (BRL)_y']

    receita_total = receita['Receita estimada (BRL)_x'].mean()
    receita_novo = receita['Receita estimada (BRL)_y'].mean()
    receita_velho = receita['Conteúdo Velho'].mean()

    receita.sort_values(by = 'Data', inplace = True)
    receita.rename(columns=colunas_receita, inplace=True)
    receita = receita.T
    receita.insert(0, 'Média (últimos 6 meses)', ['',receita_total, receita_novo, receita_velho])
    receita.insert(0, 'Mês Atual', receita.iloc[:,-1])
    receita.reset_index(inplace=True)
    receita

    with pd.ExcelWriter(f'exports/{artista}.xlsx') as writer:
        # Salvando o primeiro DataFrame
        data.to_excel(writer, sheet_name='Sheet1', index=False)

        # Obtendo o número de linhas do primeiro DataFrame
        num_rows_df1 = data.shape[0]

        # Salvando o segundo DataFrame abaixo do primeiro com uma linha vazia de separação
        conteudo.to_excel(writer, sheet_name='Sheet1', startrow=num_rows_df1 + 2, index=False)

        # Obtendo o número de linhas do segundo DataFrame
        num_rows_df2 = conteudo.shape[0]

        # Salvando o terceiro DataFrame abaixo do segundo com uma linha vazia de separação
        receita.to_excel(writer, sheet_name='Sheet1', startrow=num_rows_df1 + num_rows_df2 + 3, index=False)
