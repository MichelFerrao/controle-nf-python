import pandas as pd

# Carregar o arquivo Excel
arquivo_excel = 'notas_fiscais.xlsx'

try:
    # Ler as planilhas no Excel (a primeira planilha)
    df = pd.read_excel(arquivo_excel, sheet_name=0)

    # Remover espaços dos nomes das colunas
    df.columns = df.columns.str.strip()

    # Verificar se a coluna 'Valor' existe
    if 'Valor' in df.columns:
        # Converter a coluna de 'Valor' para numérico (remover 'R$' e ',')
        df['Valor'] = df['Valor'].replace({'R\$': '', ',': ''}, regex=True).astype(float)
    else:
        raise ValueError("A coluna 'Valor' não foi encontrada no DataFrame.")

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
        return df['Valor'].sum()

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
        print(df[['Data (NF)', 'Empresa', 'Cliente', 'Produto/Serviço', 'Tipo', 'Valor', 'Recebido']])
        print("\nResumo:")
        print(f"Total Recebido: R$ {total_recebido:.2f}")
        print(f"Total a Receber: R$ {total_nao_recebido:.2f}")

    # Gerar relatório
    gerar_relatorio(df)

except Exception as e:
    print(f"Ocorreu um erro: {e}")
import pandas as pd

# Carregar o arquivo Excel
arquivo_excel = 'notas_fiscais.xlsx'

try:
    # Ler as planilhas no Excel (a primeira planilha)
    df = pd.read_excel(arquivo_excel, sheet_name=0)

    # Remover espaços dos nomes das colunas
    df.columns = df.columns.str.strip()

    # Verificar se a coluna 'Valor' existe
    if 'Valor' in df.columns:
        # Converter a coluna de 'Valor' para numérico (remover 'R$' e ',')
        df['Valor'] = df['Valor'].replace({'R\$': '', ',': ''}, regex=True).astype(float)
    else:
        raise ValueError("A coluna 'Valor' não foi encontrada no DataFrame.")

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
        return df['Valor'].sum()

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
        print(df[['Data (NF)', 'Empresa', 'Cliente', 'Produto/Serviço', 'Tipo', 'Valor', 'Recebido']])
        print("\nResumo:")
        print(f"Total Recebido: R$ {total_recebido:.2f}")
        print(f"Total a Receber: R$ {total_nao_recebido:.2f}")

    # Gerar relatório
    gerar_relatorio(df)

except Exception as e:
    print(f"Ocorreu um erro: {e}")

def gerar_relatorio(df):
    # Contagem de notas
    total_notas = len(df)
    total_recebidas = len(filtrar_notas_recebidas(df))
    total_nao_recebidas = len(filtrar_notas_nao_recebidas(df))

    # Resumo por empresa
    resumo_empresa = df.groupby('Empresa')['Valor'].sum().reset_index()
    resumo_empresa.columns = ['Empresa', 'Total']

    # Exibir relatório
    print("\nRelatório Completo de Notas Fiscais:")
    print(df[['Data (NF)', 'Empresa', 'Cliente', 'Produto/Serviço', 'Tipo', 'Valor', 'Recebido']])
    
    print("\nResumo:")
    print(f"Total de Notas: {total_notas}")
    print(f"Notas Recebidas: {total_recebidas} (R$ {total_recebido:.2f})")
    print(f"Notas Não Recebidas: {total_nao_recebidas} (R$ {total_nao_recebido:.2f})")

    if total_notas > 0:
        percentual_recebido = (total_recebidas / total_notas) * 100
        percentual_nao_recebido = (total_nao_recebidas / total_notas) * 100
        print(f"Percentual de Notas Recebidas: {percentual_recebido:.2f}%")
        print(f"Percentual de Notas Não Recebidas: {percentual_nao_recebido:.2f}%")

    print("\nResumo por Empresa:")
    print(resumo_empresa)
    print("\nTotal Geral: R$ {:.2f}".format(total_recebido + total_nao_recebido))

# Gerar relatório
gerar_relatorio(df)

import pandas as pd

def gerar_relatorio_debitos_recebimentos(df):
    # Filtrar notas recebidas e não recebidas
    notas_recebidas = filtrar_notas_recebidas(df)
    notas_nao_recebidas = filtrar_notas_nao_recebidas(df)

    # Totais
    total_recebido = somar_valores(notas_recebidas)
    total_nao_recebido = somar_valores(notas_nao_recebidas)

    # Exibir relatório de recebimentos
    print("\nRelatório de Recebimentos:")
    print(notas_recebidas[['Data (NF)', 'Empresa', 'Cliente', 'Produto/Serviço', 'Valor']])
    print(f"\nTotal Recebido: R$ {total_recebido:.2f}")
    print(f"Número de Notas Recebidas: {len(notas_recebidas)}")

    # Exibir relatório de débitos
    print("\nRelatório de Débitos:")
    print(notas_nao_recebidas[['Data (NF)', 'Empresa', 'Cliente', 'Produto/Serviço', 'Valor']])
    print(f"\nTotal a Receber: R$ {total_nao_recebido:.2f}")
    print(f"Número de Notas Não Recebidas: {len(notas_nao_recebidas)}")

    # Resumo
    print("\nResumo Geral:")
    print(f"Total de Notas: {len(df)}")
    print(f"Total Recebido: R$ {total_recebido:.2f}")
    print(f"Total a Receber: R$ {total_nao_recebido:.2f}")
    print(f"Total Geral: R$ {total_recebido + total_nao_recebido:.2f}")

    # Análise por Empresa
    print("\nAnálise por Empresa:")
    resumo_empresa = df.groupby('Empresa').agg(
        Total_Recebido=('Valor', lambda x: x[df['Recebido'] == 'Sim'].sum()),
        Total_Nao_Recebido=('Valor', lambda x: x[df['Recebido'] == 'Não'].sum())
    ).reset_index()

    print(resumo_empresa)

# Gerar relatório
gerar_relatorio_debitos_recebimentos(df)
