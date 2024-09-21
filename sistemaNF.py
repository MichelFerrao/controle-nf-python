import pandas as pd

# Função para carregar a planilha
def carregar_planilha(arquivo):
    try:
        planilha = pd.read_excel(arquivo)
        return planilha
    except FileNotFoundError:
        print("Arquivo não encontrado!")
        return None

# Função para exibir notas fiscais
def exibir_notas_fiscais(planilha):
    if planilha is not None:
        print(planilha)
    else:
        print("Erro ao carregar planilha.")

# Função para calcular o total de notas
def calcular_total(planilha):
    if planilha is not None:
        total = planilha['Valor'].sum()
        print(f"Total em Notas Fiscais: R$ {total:.2f}")
    else:
        print("Erro ao carregar planilha.")

# Executando o código
arquivo_excel = 'controle_notas_fiscais.xlsx'  # Caminho para seu arquivo Excel
planilha = carregar_planilha(arquivo_excel)
exibir_notas_fiscais(planilha)
calcular_total(planilha)
def filtrar_notas_recebidas(planilha):
    if planilha is not None:
        recebidas = planilha[planilha['Recebido'] == 'Sim']
        print("Notas Fiscais Recebidas:")
        print(recebidas)
    else:
        print("Erro ao carregar planilha.")

filtrar_notas_recebidas(planilha)
def salvar_planilha(planilha, arquivo):
    try:
        planilha.to_excel(arquivo, index=False)
        print("Planilha salva com sucesso!")
    except Exception as e:
        print(f"Erro ao salvar planilha: {e}")

# Exemplo de uso
salvar_planilha(planilha, 'controle_notas_fiscais_atualizado.xlsx')
