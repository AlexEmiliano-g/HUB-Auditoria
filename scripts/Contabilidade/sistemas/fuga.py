import os
import re
try:
    from scripts.Contabilidade.tabulador_comum import transformar_balancete
except ModuleNotFoundError:
    from tabulador_comum import transformar_balancete

def processar(lista_arquivos):
    """Regras exclusivas para o sistema Fuga"""
    resultados = {}
    arquivos_balancete = [f for f in lista_arquivos if os.path.basename(f).upper().startswith('B_')]

    for arquivo in arquivos_balancete:
        nome_base = os.path.basename(arquivo)
        match = re.match(r"B_(\d{2}).*\.xlsx?", nome_base, re.IGNORECASE)
        nome_aba = match.group(1) if match else os.path.splitext(nome_base)[0]
        
        df = transformar_balancete(arquivo)
        resultados[nome_aba] = df

    return resultados