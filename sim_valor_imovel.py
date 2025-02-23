import pandas as pd

# Função para converter taxa anual para mensal
def calcular_taxa_mensal(taxa_anual):
    return ((1 + taxa_anual / 100) ** (1 / 12) - 1)

def calcula_taxa_anual(taxa_mensal):
    return ((1 + taxa_mensal) ** 12 - 1) * 100


# Função para calcular o valor financiado máximo usando SAC
def calcular_valor_financiado(pmt_max, taxa_mensal, prazo_meses):
    try:
        return pmt_max / ((1 / prazo_meses) + taxa_mensal)
    except ZeroDivisionError:
        print(f"Erro: Divisão por zero detectada ao calcular o valor financiado. Verifique o prazo do financiamento.")
        return None

# Ler a planilha de entrada
entrada_excel = "dados_entrada.xlsx"  # Nome do arquivo Excel de entrada
try:
    df = pd.read_excel(entrada_excel)
except FileNotFoundError:
    print(f"Erro: O arquivo '{entrada_excel}' não foi encontrado.")
    exit()

# Extrair os valores da planilha
try:
    aluguel = df.loc[0, "Aluguel"]
    iptu_anual = df.loc[0, "IPTU anual"]
    condominio = df.loc[0, "Condomínio"]
    entrada = df.loc[0, "Entrada"]
    taxa_min = df.loc[0, "Taxa anual efetiva mínima do financiamento"] 
    taxa_media = df.loc[0, "Taxa anual efetiva média do financiamento"]
    taxa_max = df.loc[0, "Taxa anual efetiva máxima do financiamento"]
    prazo_min = df.loc[0, "Prazo mínimo do financiamento (anos)"]
    prazo_medio = df.loc[0, "Prazo médio do financiamento (anos)"]
    prazo_max = df.loc[0, "Prazo máximo do financiamento (anos)"]
except KeyError as e:
    print(f"Erro: A coluna '{e}' está faltando na planilha de entrada.")
    exit()

# Calcular o valor disponível para prestações mensais
iptu_mensal = iptu_anual / 12
pmt_max = aluguel - condominio - iptu_mensal

# Converter taxas anuais para mensais
taxa_min_mensal = calcular_taxa_mensal(taxa_min * 100)
taxa_media_mensal = calcular_taxa_mensal(taxa_media * 100)
taxa_max_mensal = calcular_taxa_mensal(taxa_max * 100)

# Calcular o valor financiado máximo e o valor total máximo do imóvel
resultados = []
for taxa in [taxa_min_mensal, taxa_media_mensal, taxa_max_mensal]:
    for prazo in [prazo_min, prazo_medio, prazo_max]:
        prazo_meses = prazo * 12
        valor_financiado = calcular_valor_financiado(pmt_max, taxa, prazo_meses)
        if valor_financiado is not None:  # Ignorar cálculos inválidos
            valor_total_imovel = entrada + valor_financiado
            resultados.append({
                "Valor do Aluguel Simulado (R$)": round(aluguel, 2),
                "IPTU Simulado (R$)": round(iptu_mensal, 2),
                "Condomínio Simulado (R$)": round(condominio, 2),
                "Entrada Simulada (R$)": round(entrada, 2),
                "|": "|",
                "Taxa Anual (%)": round(calcula_taxa_anual(taxa), 2),
                "Prazo (Anos)": prazo,
                "Valor Financiado Máximo (R$)": round(valor_financiado, 2),
                "Valor Total Máximo do Imóvel (R$)": round(valor_total_imovel, 2)
            })

# Criar um DataFrame com os resultados
df_resultados = pd.DataFrame(resultados)

# Salvar os resultados em uma nova planilha Excel
saida_excel = "resultados_calculo.xlsx"
if not df_resultados.empty:  # Salvar apenas se houver resultados válidos
    df_resultados.to_excel(saida_excel, index=False)
    print(f"Os resultados foram salvos no arquivo '{saida_excel}'.")
else:
    print("Nenhum resultado válido foi gerado devido a erros nos cálculos.")