import pandas as pd
import matplotlib.pyplot as plt
import os
from io import BytesIO


ANO = 2025
DIRETORIO_DADOS = 'ArchivesDA'
DIRETORIO_SAIDA = os.path.join(DIRETORIO_DADOS, 'Output')

# Confirma a existência do diretório
os.makedirs(DIRETORIO_SAIDA, exist_ok=True)

# Caminhos dos arquivos
caminho_dados = os.path.join(DIRETORIO_DADOS, f'{ANO}_Viagem.csv')
caminho_saida_tabela = os.path.join(DIRETORIO_SAIDA, f'tabela_{ANO}.xlsx')
caminho_saida_grafico = os.path.join(DIRETORIO_SAIDA, f'grafico_{ANO}.png')

pd.set_option('display.max_columns', None)

# Leitura dos dados
df_viagens = pd.read_csv(caminho_dados, encoding='Windows-1252', sep=';', decimal=',')

# Criando nova coluna de despesas
df_viagens['Despesas'] = df_viagens[['Valor diárias', 'Valor passagens', 'Valor outros gastos']].sum(axis=1)

# Tratamento de valores nulos
df_viagens['Cargo'] = df_viagens['Cargo'].fillna('NÃO IDENTIFICADO')

# Convertendo colunas de datas
df_viagens['Período - Data de início'] = pd.to_datetime(df_viagens['Período - Data de início'], format='%d/%m/%Y')
df_viagens['Período - Data de fim'] = pd.to_datetime(df_viagens['Período - Data de fim'], format='%d/%m/%Y')

# Criando novas colunas
df_viagens['Mês da viagem'] = df_viagens['Período - Data de início'].dt.month_name()
df_viagens['Dias de viagem'] = (df_viagens['Período - Data de fim'] - df_viagens['Período - Data de início']).dt.days

# Criando tabela consolidada
df_viagens_consolidado = (
    df_viagens.groupby('Cargo').agg(
        despesa_media=('Despesas', 'mean'),
        duracao_media=('Dias de viagem', 'mean'),
        despesas_totais=('Despesas', 'sum'),
        destinos_mais_frequentes=('Destinos', pd.Series.mode),
        n_viagens=('Nome', 'count')
    ).reset_index().sort_values(by='despesas_totais', ascending=False)
)

# Filtro por cargos relevantes (> 1% das viagens)
df_cargos = df_viagens['Cargo'].value_counts(normalize=True).reset_index()
df_cargos.columns = ['Cargo', 'Proporcao']
cargos_relevantes = df_cargos.loc[df_cargos['Proporcao'] > 0.01, 'Cargo']
df_final = df_viagens_consolidado[df_viagens_consolidado['Cargo'].isin(cargos_relevantes)]

# Criando e configurando o gráfico
fig, ax = plt.subplots(figsize=(16, 6))
ax.barh(df_final['Cargo'], df_final['n_viagens'], color='#12c9a8')
ax.invert_yaxis()
ax.set_facecolor('#48464f')
fig.suptitle(f'Viagens por cargo público ({ANO})')
plt.figtext(0.65, 0.89, 'Fonte: Portal da Transparência', fontsize=8)
plt.grid(color='gray', linestyle='--', linewidth=0.5)
plt.yticks(fontsize=8)
plt.xlabel('Número de viagens')
plt.ylabel('Cargos')

# Salvando o gráfico em memória
img_data = BytesIO()
plt.savefig(img_data, format='png', bbox_inches='tight')
plt.savefig(caminho_saida_grafico, bbox_inches='tight')
plt.close()
img_data.seek(0)


# Salvando a tabela e inserindo o gráfico no Excel
with pd.ExcelWriter(caminho_saida_tabela, engine='xlsxwriter') as writer:
    df_final.to_excel(writer, sheet_name='Resumo', index=False)
    workbook = writer.book
    worksheet = writer.sheets['Resumo']
    worksheet.insert_image('H2', 'grafico.png', {'image_data': img_data})

print(f'Tabela e gráfico salvos em: {caminho_saida_tabela}')
