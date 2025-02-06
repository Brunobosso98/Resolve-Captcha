import camelot
import pandas as pd

# Caminho do PDF
pdf_path = r"C:\Users\bruno.martins\Desktop\Automações\ResolveCaptcha\livro fiscal\teste.pdf"

# Extraindo tabelas do PDF
tables = camelot.read_pdf(pdf_path, pages="all")  # 'all' processa todas as páginas

# Verifica se encontrou tabelas
if tables.n > 0:
    with pd.ExcelWriter("output.xlsx", engine="openpyxl") as writer:
        for i, table in enumerate(tables):
            df = table.df  # Converter tabela para DataFrame
            df.to_excel(writer, sheet_name=f"Pagina_{i+1}", index=False)

    print("Conversão concluída! Salvo como 'output.xlsx'")
else:
    print("Nenhuma tabela detectada no PDF.")
