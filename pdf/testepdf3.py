import pdfplumber
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
import pandas as pd

def extrair_dados_pdf(pdf_path):
    dados_extraidos = {
        "CCM": "", "CNPJ": "", "Mês Referência": "", "Situação": "", "Encerramento": "",
        "Razão Social": "", "Endereço": "", "Número": "", "Complemento": "", "Bairro": "", "Cidade": "", "Estado": "",
        "Lançamentos Válidos": [], "Lançamentos Substituídos": []
    }
    
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            lines = text.split("\n")

            is_valid = None  # Inicializa a variável is_valid
            
            for line in lines:
                if "CCM:" in line:
                    parts = line.split()
                    if len(parts) > 9:  # Verifica se há pelo menos 10 elementos
                        dados_extraidos["CCM"] = parts[1]
                        dados_extraidos["CNPJ"] = parts[3]
                        dados_extraidos["Mês Referência"] = parts[5]
                        dados_extraidos["Situação"] = parts[7]
                        dados_extraidos["Encerramento"] = parts[9]
                
                elif "Razão Social:" in line:
                    dados_extraidos["Razão Social"] = line.split(":")[1].strip()
                elif "Endereço:" in line:
                    dados_extraidos["Endereço"] = line.split(":")[1].strip()
                elif "Número:" in line:
                    dados_extraidos["Número"] = line.split(":")[1].strip()
                elif "Complemento:" in line:
                    dados_extraidos["Complemento"] = line.split(":")[1].strip()
                elif "Bairro:" in line:
                    dados_extraidos["Bairro"] = line.split(":")[1].strip()
                elif "Cidade:" in line:
                    dados_extraidos["Cidade"] = line.split(":")[1].strip()
                elif "Estado:" in line:
                    dados_extraidos["Estado"] = line.split(":")[1].strip()
                
                elif "LANÇAMENTOS VÁLIDOS" in line:
                    is_valid = True
                elif "LANÇAMENTOS SUBSTITUÍDOS" in line:
                    is_valid = False
                
                elif is_valid is True and len(line.split()) > 5:
                    dados_extraidos["Lançamentos Válidos"].append(line.split())
                elif is_valid is False and len(line.split()) > 5:
                    dados_extraidos["Lançamentos Substituídos"].append(line.split())
    
    return dados_extraidos

def criar_planilha(dados, output_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Livro Fiscal"
    
    bold_font = Font(bold=True)
    center_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thick_border = Border(
        left=Side(style="thick"), right=Side(style="thick"),
        top=Side(style="thick"), bottom=Side(style="thick")
    )
    
    ws.append(["Prefeitura Municipal de Itapira"])
    ws.append(["SECRETARIA MUNICIPAL DA FAZENDA"])
    ws.append(["Divisão de ISSQN"])
    ws.append([""])
    ws.append(["Livro Fiscal Serviços Prestados Mensal"])
    ws.append(["Data impressão:", "05/02/2025 13:35:40"])
    ws.append([""])
    
    for campo in ["CCM", "CNPJ", "Mês Referência", "Situação", "Encerramento", "Razão Social", "Endereço", "Número", "Complemento", "Bairro", "Cidade", "Estado"]:
        ws.append([f"{campo}:"])
        ws.append([dados[campo]])
    
    ws.append([""])
    ws.append(["LANÇAMENTOS VÁLIDOS"])
    ws.append(["Dia", "Número", "Série", "Tipo", "Situação", "Cod.", "Aliq.(%)", "Base(R$)", "ISS(R$)", "CNPJ Tomador", "Razão Tomador", "Lanç."])
    for row in dados["Lançamentos Válidos"]:
        ws.append(row)
    
    ws.append([""])
    ws.append(["LANÇAMENTOS SUBSTITUÍDOS E CANCELADOS"])
    ws.append(["Dia", "Número", "Série", "Tipo", "Situação", "Cod.", "Aliq.(%)", "Base(R$)", "ISS(R$)", "CNPJ Tomador", "Razão Tomador", "Evento"])
    for row in dados["Lançamentos Substituídos"]:
        ws.append(row)
    
    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = center_alignment
            if cell.row <= 12:
                cell.font = bold_font
            if cell.value:
                cell.border = thick_border
    
    wb.save(output_path)
    print(f"Arquivo Excel salvo em: {output_path}")

def main(pdf_path, output_path):
    dados = extrair_dados_pdf(pdf_path)
    criar_planilha(dados, output_path)

if __name__ == "__main__":
    pdf_file = r"C:\Users\bruno.martins\Desktop\ResolveCaptcha\livro fiscal\teste.pdf"
    excel_file = r"C:\Users\bruno.martins\Desktop\ResolveCaptcha\livro fiscal\resultado.xlsx"
    main(pdf_file, excel_file)
