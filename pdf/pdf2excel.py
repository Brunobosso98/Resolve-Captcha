import fitz  # PyMuPDF
import pandas as pd

# Carregar o PDF
pdf_document = fitz.open(r"C:\Users\bruno.martins\Desktop\Automações\ResolveCaptcha\livro fiscal\teste.pdf")

# Extrair texto e converter em DataFrame
data = []
for page in pdf_document:
    text = page.get_text("text")
    linhas = text.split("\n")
    data.extend([linha.split() for linha in linhas])

df = pd.DataFrame(data)

# Salvar como Excel
df.to_excel(r"C:\Users\bruno.martins\Desktop\Automações\ResolveCaptcha\livro fiscal\resultado.xlsx", index=False)
