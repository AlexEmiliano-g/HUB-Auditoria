# ==============================================================================
# DIRETRIZES DE ARQUITETURA E MANUTENÇÃO
# Os blocos a seguir contêm regras estruturais estritas para o funcionamento
# do sistema e para a mitigação de conflitos de versionamento (Git).
# ==============================================================================

# REGRA DE IMPORTAÇÃO: SUPORTE A MÚLTIPLOS AMBIENTES
# O bloco try/except abaixo garante que o módulo funcione corretamente tanto 
# quando integrado ao HUB principal, quanto executado isoladamente para testes. 
# É imperativo manter a estrutura de uma importação por linha.

try:
    from scripts.Contabilidade.sistemas import (
        agropan,
        auriverde,
        camnpal,
        Coagril,
        coasul,
        cocpell,
        cooabriel,
        coopatrigo,
        cooperativa_a1,
        Coopercargo,
        Cooperlate,
        Cooperoque,
        cootravale,
        cotrijuc,
        cotrimaio,
        cotripal,
        cotrisal,
        fecoagroRS,
        FecoagroSC,
        fuga,
        girandosol,
        grupo_CCGL,
        languiru,
        ouro_do_sul,
        paradiso_giovanella,
        SAP,
        serramar,
        soli3,
        uniair,
        uniodontofederacao,
        useall,
    )
except ModuleNotFoundError:
    from sistemas import (
        agropan,
        auriverde,
        camnpal,
        Coagril,
        coasul,
        cocpell,
        cooabriel,
        coopatrigo,
        cooperativa_a1,
        Coopercargo,
        Cooperlate,
        Cooperoque,
        cootravale,
        cotrijuc,
        cotrimaio,
        cotripal,
        cotrisal,
        fecoagroRS,
        FecoagroSC,
        fuga,
        girandosol,
        grupo_CCGL,
        languiru,
        ouro_do_sul,
        paradiso_giovanella,
        SAP,
        serramar,
        soli3,
        uniair,
        uniodontofederacao,
        useall,
    )

# ==============================================================================
# REGISTRO DE SISTEMAS CONTÁBEIS
# 
# PADRÃO OBRIGATÓRIO (PREVENÇÃO DE CONFLITOS DE MERGE):
# 1. Todo novo sistema deve ser inserido estritamente em ordem alfabética (A-Z).
# 2. Mantenha o formato de um sistema por linha.
# 3. É obrigatório manter a vírgula (,) após o último item do dicionário.
# ==============================================================================
SISTEMAS_REGISTRADOS = {
    "Agropan": agropan.processar,
    "Auriverde":auriverde.processar,
    "Camnpal": camnpal.processar,
    "Coagril": Coagril.processar,
    "Coasul": coasul.processar,
    "Cocpell": cocpell.processar,
    "Cooabriel": cooabriel.processar,
    "Coopatrigo": coopatrigo.processar,
    "Cooperativa A1": cooperativa_a1.processar,
    "Coopercargo": Coopercargo.processar,
    "Cooperlate": Cooperlate.processar,
    "Cooperoque": Cooperoque.processar,
    "Cootravale": cootravale.processar,
    "Cotrijuc": cotrijuc.processar,
    "Cotrimaio": cotrimaio.processar,
    "Cotripal": cotripal.processar,
    "Cotrisal": cotrisal.processar,
    "Fecoagro RS": fecoagroRS.processar,
    "Fecoagro SC": FecoagroSC.processar,
    "Fuga": fuga.processar,
    "Girando Sol": girandosol.processar,
    "Grupo CCGL": grupo_CCGL.processar,
    "Languiru": languiru.processar,
    "Ouro do Sul": ouro_do_sul.processar,
    "Paradiso Giovanella": paradiso_giovanella.processar,
    "SAP": SAP.processar,
    "Serramar": serramar.processar,
    "Soli3": soli3.processar,
    "Uniair": uniair.processar,
    "UniOdonto Federação": uniodontofederacao.processar,
    "Useall": useall.processar,
}

# ==============================================================================
# MOTOR CENTRAL DE ROTEAMENTO
# Camada de comunicação entre a Interface de Usuário (UI) e as regras de negócio.
# Alterações na assinatura destas funções causarão quebra na integração com a interface.
# ==============================================================================

def obter_nomes_sistemas():
    """Retorna a lista de sistemas para preenchimento do componente visual (Dropdown)."""
    return list(SISTEMAS_REGISTRADOS.keys())

def processar_arquivos_selecionados(lista_arquivos, sistema):
    """Encaminha os arquivos selecionados para o módulo de processamento correspondente."""
    if sistema not in SISTEMAS_REGISTRADOS:
        raise ValueError(f"O sistema '{sistema}' não possui uma regra de tabulação configurada no dicionário.")
    
    funcao_processamento = SISTEMAS_REGISTRADOS[sistema]
    return funcao_processamento(lista_arquivos)