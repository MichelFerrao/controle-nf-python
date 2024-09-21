import pandas as pd

# Carregar o arquivo Excel
arquivo_excel = 'notas_fiscais.xlsx'

# Ler as planilhas no Excel (a primeira planilha)
df = pd.read_excel(arquivo_excel, sheet_name=0)

# Exibir as primeiras linhas do DataFrame
print("Dados da planilha:")
print(df.head())

# Função para filtrar notas recebidas
def filtrar_notas_recebidas(df):
    return df[df['Recebido'] == 'Sim']

# Função para filtrar notas não recebidas
def filtrar_notas_nao_recebidas(df):
    return df[df['Recebido'] == 'Não']

# Função para somar o total de notas fiscais
def somar_valores(df):
    return df[' Valor '].sum()

# Filtrar notas recebidas
notas_recebidas = filtrar_notas_recebidas(df)
total_recebido = somar_valores(notas_recebidas)

# Exibir resultado
print(f"\nTotal Recebido: R$ {total_recebido:.2f}")
print(notas_recebidas)

# Filtrar notas não recebidas
notas_nao_recebidas = filtrar_notas_nao_recebidas(df)
total_nao_recebido = somar_valores(notas_nao_recebidas)

# Exibir resultado
print(f"\nTotal a Receber: R$ {total_nao_recebido:.2f}")
print(notas_nao_recebidas)

# Gerar relatório final
def gerar_relatorio(df):
    print("\nRelatório Completo de Notas Fiscais:")
    print(df[['Data (NF)', 'Empresa', 'Cliente', 'Produto/Serviço', ' Tipo ', ' Valor ', 'Recebido']])
    print("\nResumo:")
    print(f"Total Recebido: R$ {total_recebido:.2f}")
    print(f"Total a Receber: R$ {total_nao_recebido:.2f}")

# Gerar relatório
gerar_relatorio(df)
