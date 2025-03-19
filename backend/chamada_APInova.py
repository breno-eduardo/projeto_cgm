import os
import requests
import pandas as pd

# Definir o ano para 2024
ano = "2024"

# Construir a URL da API com o parâmetro an_referencia=2024
url = f"https://apidatalake.tesouro.gov.br/ords/siconfi/tt/extrato_entregas?id_ente=3304557&an_referencia={ano}"

# Requisição para a API
response = requests.get(url)
response.raise_for_status()  # Levanta exceção se ocorrer erro
json_data = response.json()

# Extrair a lista de itens do JSON
items = json_data["items"]

# Converter a lista de dicionários para um DataFrame
df = pd.DataFrame(items)

# Substituir os valores na coluna "status_relatorio"
df["status_relatorio"] = df["status_relatorio"].replace({
    "HO": "homologado",
    "RE": "retificado"
})

# Adicionar a coluna "org_entregavel" concatenando "instituicao" e "entregavel"
df["org_entregavel"] = df["instituicao"] + " - " + df["entregavel"]

# Excluir as colunas indesejadas
colunas_excluir = ["exercicio", "cod_ibge", "populacao", "forma_envio", "tipo_relatorio"]
df.drop(columns=colunas_excluir, inplace=True, errors="ignore")

# Definir os nomes dos arquivos de saída no mesmo diretório de execução
output_csv = os.path.join(os.getcwd(), f"Siconfi_{ano}_output.csv")
output_xlsx = os.path.join(os.getcwd(), f"Siconfi_{ano}_output.xlsx")

# Salvar o DataFrame em um arquivo CSV com codificação UTF-8 (sem o índice)
df.to_csv(output_csv, index=False, encoding="utf-8")
print("Arquivo CSV gerado com sucesso:", output_csv)

# Salvar o DataFrame em um arquivo Excel (XLSX), sem o índice
df.to_excel(output_xlsx, index=False, engine="openpyxl")
print("Arquivo XLSX gerado com sucesso:", output_xlsx)
