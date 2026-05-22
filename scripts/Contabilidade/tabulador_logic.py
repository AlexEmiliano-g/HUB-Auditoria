# Tenta importar os sistemas (suporta tanto o uso pelo HUB quanto isolado)
try:
    from scripts.Contabilidade.sistemas import fuga, uniair
except ModuleNotFoundError:
    from sistemas import fuga, uniair

# ==============================================================================
# REGISTRO DE SISTEMAS
# Quando você criar um novo sistema no futuro (ex: dominio.py),
# basta importar ali em cima e adicionar o nome dele neste dicionário abaixo!
# ==============================================================================
SISTEMAS_REGISTRADOS = {
    "Fuga": fuga.processar,
    "Uniair": uniair.processar
}

def obter_nomes_sistemas():
    """Retorna a lista de sistemas para preencher o Menu Dropdown na Interface."""
    return list(SISTEMAS_REGISTRADOS.keys())

def processar_arquivos_selecionados(lista_arquivos, sistema):
    """Pega a função correta do dicionário e a executa."""
    if sistema not in SISTEMAS_REGISTRADOS:
        raise ValueError(f"O sistema '{sistema}' não possui uma regra de tabulação configurada.")
    
    # Busca a função no dicionário e a executa
    funcao_processamento = SISTEMAS_REGISTRADOS[sistema]
    return funcao_processamento(lista_arquivos)