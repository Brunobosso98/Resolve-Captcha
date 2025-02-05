import re
import pandas as pd
import os

# Ler o conteúdo do arquivo Markdown
with open("dados.txt", "r", encoding="utf-8") as f:
    content = f.read()

# Expressão regular para capturar CNPJ e o valor após "Pegar Valores Simples Nacional"
pattern = r"CNPJ:\s*(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}).*?Pegar Valores Simples Nacional:\s*R?\$?([\d.,]+)"

# Encontrar todas as ocorrências
matches = re.findall(pattern, content, re.DOTALL)

# Criar DataFrame para melhor visualização
df = pd.DataFrame(matches, columns=["CNPJ", "Valor Simples Nacional"])

# Salvar os dados extraídos em um arquivo Excel, se não existir
if not os.path.exists("dados_extraidos.xlsx"):
    df.to_excel("dados_extraidos.xlsx", index=False)

# Exibir os dados extraídos no terminal
print(df)
