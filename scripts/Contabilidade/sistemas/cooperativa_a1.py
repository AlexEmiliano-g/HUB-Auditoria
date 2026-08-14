import os
import warnings

try:
    from scripts.Contabilidade.tabulador_comum import (
        transformar_balancete_cooperativa_a1,
        obter_nome_aba_seguro
    )
except ModuleNotFoundError:
    from tabulador_comum import (
        transformar_balancete_cooperativa_a1,
        obter_nome_aba_seguro
    )

def processar(lista_arquivos):
    """
    Regras exclusivas para o sistema Cooperativa A1:
    - Lê balancetes originados de relatórios em Texto (.txt).
    - Gera nome de abas dinamicamente, tratando colisões (regra B_XX).
    """
    resultados = {}
    abas_existentes = set()
    
    # Filtra os arquivos válidos pelo formato do sistema
    arquivos_balancete = [f for f in lista_arquivos if f.lower().endswith('.txt')]
    arquivos_ignorados = [f for f in lista_arquivos if not f.lower().endswith('.txt')]

    if arquivos_ignorados:
        nomes_ignorados = ", ".join([os.path.basename(f) for f in arquivos_ignorados])
        warnings.warn(
            f"Os seguintes arquivos foram ignorados no sistema Cooperativa A1 pois não são TXT (.txt): {nomes_ignorados}"
        )

    for arquivo in arquivos_balancete:
        nome_aba = obter_nome_aba_seguro(arquivo, abas_existentes)
        abas_existentes.add(nome_aba)
        
        df = transformar_balancete_cooperativa_a1(arquivo)
        resultados[nome_aba] = df

    if not resultados:
        raise ValueError("Nenhum arquivo válido (.txt) foi processado para a Cooperativa A1.")

    return resultados