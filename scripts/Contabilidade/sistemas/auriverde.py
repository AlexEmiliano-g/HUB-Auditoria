import os
import warnings

try:
    from scripts.Contabilidade.tabulador_comum import (
        transformar_balancete_auriverde,
        obter_nome_aba_seguro
    )
except ModuleNotFoundError:
    from tabulador_comum import (
        transformar_balancete_auriverde,
        obter_nome_aba_seguro
    )

def processar(lista_arquivos):
    """
    Regras exclusivas para o sistema Auriverde:
    - Lê balancetes de planilhas Excel (XLS, XLSX).
    - Gera nome de abas dinamicamente, tratando colisões (regra B_XX).
    """
    resultados = {}
    abas_existentes = set() # Controle global de abas usadas neste processamento
    
    arquivos_balancete = [f for f in lista_arquivos if f.lower().endswith(('.xls', '.xlsx'))]
    arquivos_ignorados = [f for f in lista_arquivos if not f.lower().endswith(('.xls', '.xlsx'))]

    if arquivos_ignorados:
        nomes_ignorados = ", ".join([os.path.basename(f) for f in arquivos_ignorados])
        warnings.warn(
            f"Os seguintes arquivos foram ignorados no sistema Auriverde pois não são Excel (.xls, .xlsx): {nomes_ignorados}"
        )

    for arquivo in arquivos_balancete:
        # 1. Pega o nome seguro com base na regra geral do HUB passando as abas já processadas
        nome_aba = obter_nome_aba_seguro(arquivo, abas_existentes)
        
        # 2. Registra que essa aba acabou de ser utilizada
        abas_existentes.add(nome_aba)
        
        # 3. Faz o processamento do dataframe
        df = transformar_balancete_auriverde(arquivo)
        resultados[nome_aba] = df

    if not resultados:
        raise ValueError("Nenhum arquivo válido (.xls ou .xlsx) foi processado para o sistema Auriverde.")

    return resultados