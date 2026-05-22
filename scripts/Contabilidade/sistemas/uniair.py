import os
import re
try:
    from scripts.Contabilidade.tabulador_comum import transformar_balancete, transformar_plano_de_contas
except ModuleNotFoundError:
    from tabulador_comum import transformar_balancete, transformar_plano_de_contas

def processar(lista_arquivos):
    """Regras exclusivas para o sistema Uniair"""
    resultados = {}
    arquivos_plano_contas = [f for f in lista_arquivos if os.path.basename(f).upper().startswith('P_')]
    arquivos_balancete = [f for f in lista_arquivos if os.path.basename(f).upper().startswith('B_')]

    # Processa os Balancetes
    for arquivo in arquivos_balancete:
        nome_base = os.path.basename(arquivo)
        match = re.match(r"B_(\d{2}).*\.xlsx?", nome_base, re.IGNORECASE)
        nome_aba = match.group(1) if match else os.path.splitext(nome_base)[0]
        
        df = transformar_balancete(arquivo)
        resultados[nome_aba] = df

    # O Uniair possui a regra adicional do Plano de Contas
    if arquivos_plano_contas:
        df_plano = transformar_plano_de_contas(arquivos_plano_contas[0])
        resultados["Plano de Contas"] = df_plano
        
    return resultados