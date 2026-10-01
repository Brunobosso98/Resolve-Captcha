"""
Validação de NFS-e emitidas.

Cruza a planilha 'notasEmitir.xlsx' (notas planejadas) com o log de emissões
do SIGISS (notas executadas) e identifica divergências por CNPJ + Valor.

Como usar:
    1. Exporte o relatório de emissões do SIGISS e cole as linhas no arquivo
       'notas_emitidas.txt' (uma nota por linha). Formato esperado (campos
       separados por TAB ou múltiplos espaços):
         NUMERO  DATA_HORA  DATA  TIPO  SERVICO  VALOR  CNPJ  STATUS  CHAVE

    2. Execute a partir da raiz do projeto:
         python NOTAS/validar_nfse.py

    3. Veja o relatório no terminal e em 'validacao_resultado.txt'.

Argumentos opcionais:
    -p/--planilha  Caminho da planilha (padrão: notasEmitir.xlsx)
    -l/--log       Caminho do log de emissões (padrão: notas_emitidas.txt)
    -o/--output    Caminho do relatório gerado (padrão: validacao_resultado.txt)
"""

import os
import re
import sys
import argparse
from datetime import datetime
from collections import defaultdict

import pandas as pd


PLANILHA_PADRAO = "notasEmitir.xlsx"
LOG_PADRAO = "notas_emitidas.txt"
RELATORIO_PADRAO = "validacao_resultado.txt"

COLUNAS_CNPJ = ("CNPJ Tomador", "CNPJ", "CPF/CNPJ", "Documento", "CPF")
COLUNAS_VALOR = ("Valor", "valor", "Valor NF", "Valor Nota")


def normalizar_cnpj(valor):
    """Remove tudo que não é dígito."""
    if valor is None:
        return ""
    return re.sub(r"\D", "", str(valor)).strip()


def normalizar_valor(valor):
    """
    Converte representações de valor monetário para float.
    Aceita: float, int, "1500,00", "1.500,00", "1500.00", "R$ 1.500,00".
    Retorna None se não conseguir converter.
    """
    if valor is None:
        return None
    if isinstance(valor, float) and pd.isna(valor):
        return None
    if isinstance(valor, (int, float)):
        return round(float(valor), 2)
    texto = str(valor).strip()
    if not texto:
        return None
    texto = texto.replace("R$", "").replace(" ", "")
    # Se tem vírgula, é formato brasileiro: ponto = milhar, vírgula = decimal
    if "," in texto:
        texto = texto.replace(".", "").replace(",", ".")
    try:
        return round(float(texto), 2)
    except ValueError:
        return None


def parse_linha_emissao(linha):
    """
    Faz o parse de uma linha de emissão do SIGISS.

    Formato esperado (separador TAB ou múltiplos espaços):
        NUMERO  DATA_HORA  DATA  TIPO  SERVICO  VALOR  CNPJ  STATUS  CHAVE
    Retorna dict normalizado ou None se a linha for inválida.
    """
    partes = re.split(r"\s{2,}|\t", linha.strip())
    if len(partes) < 9:
        return None
    try:
        valor = normalizar_valor(partes[5])
        cnpj = normalizar_cnpj(partes[6])
        if not cnpj or valor is None:
            return None
        return {
            "numero": partes[0].strip(),
            "data_hora": partes[1].strip(),
            "data": partes[2].strip(),
            "tipo": partes[3].strip(),
            "servico": partes[4].strip(),
            "valor": valor,
            "cnpj": cnpj,
            "status": partes[7].strip(),
            "chave": partes[8].strip(),
            "valor_original": partes[5].strip(),
            "cnpj_original": partes[6].strip(),
        }
    except (ValueError, IndexError):
        return None


def carregar_emitidas(caminho):
    """Lê o arquivo de emissões (uma nota por linha)."""
    if not os.path.exists(caminho):
        return None
    with open(caminho, "r", encoding="utf-8") as f:
        texto = f.read()
    resultados = []
    for linha in texto.splitlines():
        linha = linha.strip()
        if not linha or linha.startswith("#"):
            continue
        item = parse_linha_emissao(linha)
        if item:
            resultados.append(item)
    return resultados


def encontrar_coluna(df, candidatos):
    """Encontra a primeira coluna cujo nome bate com algum candidato."""
    for col in df.columns:
        if str(col).strip() in candidatos:
            return col
    cols_lower = {str(c).strip().lower(): c for c in df.columns}
    for cand in candidatos:
        if cand.lower() in cols_lower:
            return cols_lower[cand.lower()]
    return None


def carregar_planejadas(caminho):
    """Lê a planilha de notas a emitir e retorna lista de planejadas."""
    if not os.path.exists(caminho):
        return None
    df = pd.read_excel(caminho, engine="openpyxl")
    col_cnpj = encontrar_coluna(df, COLUNAS_CNPJ)
    col_valor = encontrar_coluna(df, COLUNAS_VALOR)
    if not col_cnpj:
        print(f"ERRO: nenhuma coluna de CNPJ encontrada. Colunas: {list(df.columns)}")
        return None
    if not col_valor:
        print(f"ERRO: nenhuma coluna de Valor encontrada. Colunas: {list(df.columns)}")
        return None

    planejadas = []
    for idx, row in df.iterrows():
        cnpj = normalizar_cnpj(row[col_cnpj])
        valor = normalizar_valor(row[col_valor])
        if not cnpj or valor is None or valor <= 0:
            continue
        planejadas.append({
            "linha_excel": idx + 2,  # +1 do índice 0-based, +1 do cabeçalho
            "cnpj": cnpj,
            "valor": valor,
            "cnpj_original": str(row[col_cnpj]).strip(),
            "valor_original": str(row[col_valor]).strip(),
        })
    return planejadas


def cruzar(planejadas, emitidas):
    """
    Faz o cruzamento por (CNPJ, Valor).
    Apenas notas com status 'Valida' são consideradas como emitidas.
    """
    planejadas_por_chave = defaultdict(list)
    for p in planejadas:
        planejadas_por_chave[(p["cnpj"], p["valor"])].append(p)

    emitidas_validas = [e for e in emitidas if e["status"].lower() == "valida"]
    emitidas_por_chave = defaultdict(list)
    for e in emitidas_validas:
        emitidas_por_chave[(e["cnpj"], e["valor"])].append(e)

    usadas = set()
    conferidas = []
    faltantes = []

    for chave, lista_p in planejadas_por_chave.items():
        disponiveis = list(emitidas_por_chave.get(chave, []))
        for p in lista_p:
            if disponiveis:
                emitida = disponiveis.pop(0)
                usadas.add(id(emitida))
                conferidas.append({"planejada": p, "emitida": emitida})
            else:
                faltantes.append(p)

    extras = []
    duplicadas = []
    for chave, lista_e in emitidas_por_chave.items():
        qtd_planejada = len(planejadas_por_chave.get(chave, []))
        if chave not in planejadas_por_chave:
            extras.extend(lista_e)
        elif len(lista_e) > qtd_planejada:
            excesso = len(lista_e) - qtd_planejada
            duplicadas.extend(lista_e[-excesso:])

    return {
        "conferidas": conferidas,
        "faltantes": faltantes,
        "extras": extras,
        "duplicadas": duplicadas,
    }


def formatar_valor_br(valor):
    """Formata float como 1.500,00."""
    return f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def gerar_relatorio(resultados, args):
    agora = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    L = []
    L.append("=" * 100)
    L.append("RELATÓRIO DE VALIDAÇÃO DE NFS-e")
    L.append(f"Gerado em:    {agora}")
    L.append(f"Planilha:     {args.planilha}")
    L.append(f"Log emissão:  {args.log}")
    L.append("=" * 100)

    qtd_conferidas = len(resultados["conferidas"])
    qtd_faltantes = len(resultados["faltantes"])
    qtd_extras = len(resultados["extras"])
    qtd_duplicadas = len(resultados["duplicadas"])
    qtd_planejadas = qtd_conferidas + qtd_faltantes
    qtd_emitidas = qtd_conferidas + qtd_extras + qtd_duplicadas

    L.append("")
    L.append(f"📋  Planejadas na planilha:  {qtd_planejadas}")
    L.append(f"📤  Emitidas (válidas):      {qtd_emitidas}")
    L.append(f"✅  Conferidas (match):      {qtd_conferidas}")
    L.append(f"❌  Faltam emitir:           {qtd_faltantes}")
    L.append(f"⚠️   Duplicadas no log:      {qtd_duplicadas}")
    L.append(f"❓  Extras (não planejadas): {qtd_extras}")

    # Conferidas
    L.append("")
    L.append("-" * 100)
    L.append(f"✅ NOTAS CONFERIDAS ({qtd_conferidas})")
    L.append("-" * 100)
    if resultados["conferidas"]:
        L.append(
            f"{'Plan.':>5}  {'CNPJ':>14}  {'Valor Plan.':>13}  "
            f"{'Nota':>6}  {'Valor Emit.':>13}  {'Data/Hora':>19}  {'Chave (final)':>12}"
        )
        for item in resultados["conferidas"]:
            p = item["planejada"]
            e = item["emitida"]
            L.append(
                f"{p['linha_excel']:>5}  {p['cnpj']:>14}  {formatar_valor_br(p['valor']):>13}  "
                f"{e['numero']:>6}  {formatar_valor_br(e['valor']):>13}  "
                f"{e['data_hora']:>19}  ...{e['chave'][-10:]:>12}"
            )
    else:
        L.append("(nenhuma)")

    # Faltantes
    L.append("")
    L.append("-" * 100)
    L.append(f"❌ PLANEJADAS MAS NÃO ENCONTRADAS ({qtd_faltantes})")
    L.append("-" * 100)
    if resultados["faltantes"]:
        L.append(f"{'Plan.':>5}  {'CNPJ':>14}  {'Valor':>13}  {'CNPJ original':>20}  {'Valor original':>15}")
        for p in resultados["faltantes"]:
            L.append(
                f"{p['linha_excel']:>5}  {p['cnpj']:>14}  {formatar_valor_br(p['valor']):>13}  "
                f"{p['cnpj_original']:>20}  {p['valor_original']:>15}"
            )
    else:
        L.append("(nenhuma — todas as planejadas foram emitidas)")

    # Duplicadas
    L.append("")
    L.append("-" * 100)
    L.append(f"⚠️  DUPLICADAS NO LOG DE EMISSÕES ({qtd_duplicadas})")
    L.append("-" * 100)
    if resultados["duplicadas"]:
        L.append(f"{'Nota':>6}  {'CNPJ':>14}  {'Valor':>13}  {'Data/Hora':>19}  {'Chave':>44}")
        for e in resultados["duplicadas"]:
            L.append(
                f"{e['numero']:>6}  {e['cnpj']:>14}  {formatar_valor_br(e['valor']):>13}  "
                f"{e['data_hora']:>19}  {e['chave']}"
            )
    else:
        L.append("(nenhuma)")

    # Extras
    L.append("")
    L.append("-" * 100)
    L.append(f"❓ EXTRAS NO LOG (NÃO PLANEJADAS) ({qtd_extras})")
    L.append("-" * 100)
    if resultados["extras"]:
        L.append(f"{'Nota':>6}  {'CNPJ':>14}  {'Valor':>13}  {'Data/Hora':>19}  {'Chave':>44}")
        for e in resultados["extras"]:
            L.append(
                f"{e['numero']:>6}  {e['cnpj']:>14}  {formatar_valor_br(e['valor']):>13}  "
                f"{e['data_hora']:>19}  {e['chave']}"
            )
    else:
        L.append("(nenhuma)")

    L.append("")
    L.append("=" * 100)
    return "\n".join(L) + "\n"


def main():
    parser = argparse.ArgumentParser(
        description="Valida se as NFS-e planejadas na planilha foram realmente emitidas no SIGISS."
    )
    parser.add_argument("-p", "--planilha", default=PLANILHA_PADRAO,
                        help=f"Planilha com notas planejadas (padrão: {PLANILHA_PADRAO})")
    parser.add_argument("-l", "--log", default=LOG_PADRAO,
                        help=f"Log de emissões do SIGISS (padrão: {LOG_PADRAO})")
    parser.add_argument("-o", "--output", default=RELATORIO_PADRAO,
                        help=f"Arquivo de saída do relatório (padrão: {RELATORIO_PADRAO})")
    args = parser.parse_args()

    if not os.path.exists(args.planilha):
        print(f"❌ ERRO: planilha '{args.planilha}' não encontrada.")
        sys.exit(1)

    planejadas = carregar_planejadas(args.planilha)
    if planejadas is None:
        sys.exit(1)

    if not os.path.exists(args.log):
        print(f"❌ ERRO: log de emissões '{args.log}' não encontrado.")
        print("   Crie o arquivo colando as linhas exportadas do SIGISS.")
        print("   Formato (campos separados por TAB ou múltiplos espaços):")
        print("     NUMERO  DATA_HORA  DATA  TIPO  SERVICO  VALOR  CNPJ  STATUS  CHAVE")
        sys.exit(1)

    emitidas = carregar_emitidas(args.log)
    if emitidas is None:
        sys.exit(1)

    print(f"📋  {len(planejadas)} nota(s) planejada(s) carregada(s) de '{args.planilha}'")
    print(f"📤  {len(emitidas)} nota(s) emitida(s) carregada(s) de '{args.log}'")

    if not planejadas or not emitidas:
        print("Nada para cruzar.")
        sys.exit(0)

    resultados = cruzar(planejadas, emitidas)
    texto = gerar_relatorio(resultados, args)

    print(texto)

    with open(args.output, "w", encoding="utf-8") as f:
        f.write(texto)
    print(f"📝  Relatório salvo em '{args.output}'")


if __name__ == "__main__":
    main()
