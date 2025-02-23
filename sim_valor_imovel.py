import pandas as pd
import numpy as np

# Função para converter taxa anual para mensal
def calcular_taxa_mensal(taxa_anual):
    return ((1 + taxa_anual) ** (1 / 12) - 1)

def calcula_taxa_anual(taxa_mensal):
    return ((1 + taxa_mensal) ** 12 - 1)*100

# Função para calcular o valor financiado máximo usando SAC
def calcular_valor_financiado(pmt_max, taxa_mensal, prazo_meses):
    try:
        return pmt_max / ((1 / prazo_meses) + taxa_mensal)
    except ZeroDivisionError:
        print(f"Erro: Divisão por zero detectada ao calcular o valor financiado. Verifique o prazo do financiamento.")
        return None

# Ler a planilha de entrada
entrada_excel = "dados_entrada_simulação.xlsx"  # Nome do arquivo Excel de entrada
try:
    df = pd.read_excel(entrada_excel)
except FileNotFoundError:
    print(f"Erro: O arquivo '{entrada_excel}' não foi encontrado.")
    exit()

# Extrair os valores da planilha
try:
    aluguel = df.loc[0, "Aluguel"]
    iptu_min = df.loc[0, "IPTU anual mínimo"]
    iptu_max = df.loc[0, "IPTU anual máximo"]
    condominio_min = df.loc[0, "Condomínio mínimo"]
    condominio_max = df.loc[0, "Condomínio máximo"]
    taxa_min = df.loc[0, "Taxa anual efetiva mínima do financiamento"]
    taxa_max = df.loc[0, "Taxa anual efetiva máxima do financiamento"]
    entrada = df.loc[0, "Entrada"]
    prazo_min = df.loc[0, "Prazo mínimo do financiamento (anos)"]
    prazo_medio = df.loc[0, "Prazo médio do financiamento (anos)"]
    prazo_max = df.loc[0, "Prazo máximo do financiamento (anos)"]
except KeyError as e:
    print(f"Erro: A coluna '{e}' está faltando na planilha de entrada.")
    exit()

# Gerar 10 valores aleatórios para IPTU, Condomínio e Taxa Anual com distribuição normal
np.random.seed(42)  # Fixar semente para reprodutibilidade
iptu_simulado = np.random.uniform(iptu_min, iptu_max, 10)
condominio_simulado = np.random.uniform(condominio_min, condominio_max, 10)
taxa_simulada = np.random.uniform(taxa_min, taxa_max, 10)

# Criar lista para armazenar resultados
resultados = []

# Iterar sobre todas as combinações de IPTU, Condomínio, Taxa e Prazos
for iptu in iptu_simulado:
    for condominio in condominio_simulado:
        for taxa in taxa_simulada:
            for prazo in [prazo_min, prazo_medio, prazo_max]:
                # Calcular o valor disponível para prestações mensais
                iptu_mensal = iptu / 12
                pmt_max = aluguel - condominio - iptu_mensal

                # Converter taxa anual para mensal
                taxa_mensal = calcular_taxa_mensal(taxa)

                # Calcular o valor financiado máximo
                prazo_meses = prazo * 12
                valor_financiado = calcular_valor_financiado(pmt_max, taxa_mensal, prazo_meses)

                if valor_financiado is not None:  # Ignorar cálculos inválidos
                    valor_total_imovel = entrada + valor_financiado
                    resultados.append({
                        "Valor do Aluguel Simulado (R$)": round(aluguel, 2),
                        "IPTU anual Simulado (R$)": round(iptu, 2),
                        "Condomínio Simulado (R$)": round(condominio, 2),
                        "Entrada Simulada (R$)": round(entrada, 2),
                        "Taxa Anual (%)": f"{round(taxa*100, 2)}%",
                        "Prazo (Anos)": prazo,
                        "Valor Financiado Máximo (R$)": round(valor_financiado, 2),
                        "Valor Total Máximo do Imóvel (R$)": round(valor_total_imovel, 2)
                    })

# Criar DataFrame com os resultados
df_resultados = pd.DataFrame(resultados)

# Ordenar os resultados pelo Valor Total Máximo do Imóvel
df_resultados = df_resultados.sort_values(by="Valor Total Máximo do Imóvel (R$)", ascending=False)

# Adicionar coluna de percentil
df_resultados['Percentil'] = pd.qcut(df_resultados['Valor Total Máximo do Imóvel (R$)'], q=100, labels=False)
df_resultados['Percentil'] = df_resultados['Percentil'].apply(lambda x: f"P{x}")

# Filtrar apenas os percentis desejados
percentis_desejados = ["P0", "P10", "P50", "P90", "P100"]
df_filtrado = df_resultados[df_resultados['Percentil'].isin(percentis_desejados)]

# Salvar os resultados filtrados em uma nova planilha Excel
saida_excel = "resultados_calculo.xlsx"
if not df_filtrado.empty:
    df_filtrado.to_excel(saida_excel, index=False)
    print(f"Os resultados foram salvos no arquivo '{saida_excel}'.")
else:
    print("Nenhum resultado válido foi gerado devido a erros nos cálculos.")