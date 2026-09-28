import os
import re
import unicodedata

import pandas as pd


# ==============================================================================
# CONFIGURACOES DO SISTEMA FecoagroRS
# ==============================================================================

EXTENSOES_VALIDAS_fecoagroRS = {".xls", ".xlsx"}

COLUNAS_DESTINO_fecoagroRS = [
    "Atividade",
    "Conta",
    "Nome",
    "Cód. Reduzido",
    "Saldo Anterior",
    "Débito",
    "Crédito",
    "Movimento",
    "Saldo Acumulado",
]

# Mapeamento literal informado pelo usuario.
# Indices Python iniciam em zero.
INDICE_CODIGO_REDUZIDO_fecoagroRS = 0   # Coluna A
INDICE_CLASSIFICACAO_fecoagroRS = 1     # Coluna B
INDICE_NOME_fecoagroRS = 3              # Coluna D
INDICE_SALDO_ANTERIOR_fecoagroRS = 5    # Coluna F
INDICE_DEBITO_fecoagroRS = 6            # Coluna G
INDICE_CREDITO_fecoagroRS = 8           # Coluna I
INDICE_SALDO_ACUMULADO_fecoagroRS = 10  # Coluna K


# ==============================================================================
# NOMES DAS ABAS
# ==============================================================================

def _limpar_nome_aba_fecoagroRS(nome):
    """Cria um nome valido de aba, limitado a 31 caracteres."""
    nome_limpo = re.sub(
        r'[\\/\x2a?:\[\]]',
        "_",
        str(nome).strip(),
    )
    return (nome_limpo or "Sem nome")[:31]


def _obter_nome_sem_extensao_fecoagroRS(caminho_arquivo):
    """Retorna o nome do arquivo sem a extensao."""
    nome_arquivo = os.path.basename(caminho_arquivo)
    return os.path.splitext(nome_arquivo)[0].strip()


def _obter_mes_fecoagroRS(caminho_arquivo):
    """Extrai XX quando o nome do arquivo comeca com B_XX."""
    nome_arquivo = os.path.basename(caminho_arquivo)
    correspondencia = re.match(
        r"^B_(0[1-9]|1[0-2])",
        nome_arquivo,
        re.IGNORECASE,
    )
    return correspondencia.group(1) if correspondencia else None


def _registrar_nome_aba_fecoagroRS(nome, nomes_utilizados):
    """Registra um nome de aba se ainda nao estiver em uso."""
    nome_aba = _limpar_nome_aba_fecoagroRS(nome)
    chave = nome_aba.casefold()

    if chave in nomes_utilizados:
        return None

    nomes_utilizados.add(chave)
    return nome_aba


def _gerar_nome_aba_fecoagroRS(caminho_arquivo, nomes_utilizados):
    """Gera nome unico, priorizando o mes de arquivos B_XX."""
    nome_completo = _obter_nome_sem_extensao_fecoagroRS(caminho_arquivo)
    mes = _obter_mes_fecoagroRS(caminho_arquivo)

    if mes:
        nome_aba = _registrar_nome_aba_fecoagroRS(mes, nomes_utilizados)
        if nome_aba is not None:
            return nome_aba

    nome_aba = _registrar_nome_aba_fecoagroRS(
        nome_completo,
        nomes_utilizados,
    )
    if nome_aba is not None:
        return nome_aba

    contador = 2
    while True:
        sufixo = "_" + str(contador)
        limite = 31 - len(sufixo)
        candidato = nome_completo[:limite] + sufixo
        nome_aba = _registrar_nome_aba_fecoagroRS(
            candidato,
            nomes_utilizados,
        )

        if nome_aba is not None:
            return nome_aba

        contador += 1


# ==============================================================================
# NORMALIZACAO DE TEXTOS E CODIGOS
# ==============================================================================

def _normalizar_texto_fecoagroRS(valor):
    """Normaliza espacos e converte valores ausentes em texto vazio."""
    if valor is None:
        return ""

    try:
        if pd.isna(valor):
            return ""
    except (TypeError, ValueError):
        pass

    texto = str(valor).replace("\xa0", " ")
    return re.sub(r"\s{1,}", " ", texto).strip()


def _normalizar_texto_comparacao_fecoagroRS(valor):
    """Converte para maiusculas e remove acentos para comparacao."""
    texto = _normalizar_texto_fecoagroRS(valor).upper()
    if not texto:
        return ""

    texto = unicodedata.normalize("NFKD", texto)
    return "".join(
        caractere
        for caractere in texto
        if not unicodedata.combining(caractere)
    )


def _normalizar_classificacao_fecoagroRS(valor):
    """Remove pontos e espacos da classificacao da coluna B."""
    if valor is None:
        return ""

    try:
        if pd.isna(valor):
            return ""
    except (TypeError, ValueError):
        pass

    if isinstance(valor, int):
        texto = str(valor)
    elif isinstance(valor, float):
        if valor.is_integer():
            texto = str(int(valor))
        else:
            texto = format(valor, "f").rstrip("0").rstrip(".")
    else:
        texto = _normalizar_texto_fecoagroRS(valor)

    texto = (
        texto.replace("\xa0", "")
        .replace(" ", "")
        .replace(",", ".")
        .replace(".", "")
    )

    return texto if texto.isdigit() else ""


def _normalizar_codigo_reduzido_fecoagroRS(valor):
    """Normaliza o codigo reduzido da coluna A e o mantem como texto."""
    if valor is None:
        return ""

    try:
        if pd.isna(valor):
            return ""
    except (TypeError, ValueError):
        pass

    if isinstance(valor, int):
        return str(valor)

    if isinstance(valor, float):
        if valor.is_integer():
            return str(int(valor))
        return format(valor, "f").rstrip("0").rstrip(".")

    texto = _normalizar_texto_fecoagroRS(valor)
    if re.fullmatch(r"\d+\.0+", texto):
        return texto.split(".", maxsplit=1)[0]

    return texto


# ==============================================================================
# CONVERSAO DOS VALORES E TRATAMENTO D/C
# ==============================================================================

def _extrair_natureza_fecoagroRS(valor):
    """Extrai D ou C existente no final da propria celula de valor."""
    texto = _normalizar_texto_fecoagroRS(valor).upper()
    if not texto:
        return ""

    correspondencia = re.search(
        r"([DC])\s{0,}$",
        texto,
        re.IGNORECASE,
    )
    return correspondencia.group(1).upper() if correspondencia else ""


def _converter_numero_fecoagroRS(valor):
    """Converte valor monetario e remove D/C do final da celula."""
    if valor is None:
        return 0.0

    try:
        if pd.isna(valor):
            return 0.0
    except (TypeError, ValueError):
        pass

    if isinstance(valor, (int, float)):
        return round(float(valor), 2)

    texto = str(valor).strip().upper()
    if texto in {"", "-", "--"}:
        return 0.0

    texto = re.sub(
        r"\s{0,}[DC]\s{0,}$",
        "",
        texto,
        flags=re.IGNORECASE,
    )

    texto = (
        texto.replace("\xa0", "")
        .replace(" ", "")
        .replace("R$", "")
        .replace("$", "")
    )

    negativo_parenteses = texto.startswith("(") and texto.endswith(")")
    negativo_final = texto.endswith("-") and texto != "-"

    if negativo_parenteses:
        texto = texto[1:-1]
    if negativo_final:
        texto = texto[:-1]

    if "," in texto:
        texto = texto.replace(".", "").replace(",", ".")

    try:
        numero = float(texto)
    except (ValueError, TypeError):
        return 0.0

    if negativo_parenteses or negativo_final:
        numero = -abs(numero)

    return round(numero, 2)


def _aplicar_natureza_fecoagroRS(valor, natureza_padrao=""):
    """
    Aplica a natureza existente na celula ou a natureza padrao.

    A natureza da celula tem prioridade. Quando nao houver D/C:
        natureza_padrao D -> positivo
        natureza_padrao C -> negativo
        natureza vazia    -> preserva o sinal original
    """
    numero = _converter_numero_fecoagroRS(valor)
    natureza_celula = _extrair_natureza_fecoagroRS(valor)
    natureza_padrao = _normalizar_texto_comparacao_fecoagroRS(
        natureza_padrao
    ).strip()

    if natureza_celula in {"D", "C"}:
        natureza_final = natureza_celula
    elif natureza_padrao in {"D", "C"}:
        natureza_final = natureza_padrao
    else:
        natureza_final = ""

    if natureza_final == "D":
        return round(abs(numero), 2)
    if natureza_final == "C":
        return round(-abs(numero), 2)

    return round(numero, 2)


# ==============================================================================
# LEITURA DO EXCEL
# ==============================================================================

def _obter_engine_excel_fecoagroRS(caminho_arquivo):
    """Retorna o mecanismo adequado para XLS ou XLSX."""
    extensao = os.path.splitext(caminho_arquivo)[1].lower()
    return "xlrd" if extensao == ".xls" else "openpyxl"


def _ler_excel_fecoagroRS(caminho_arquivo):
    """Le a primeira aba sem considerar uma linha fixa de cabecalho."""
    nome_arquivo = os.path.basename(caminho_arquivo)

    try:
        dataframe = pd.read_excel(
            caminho_arquivo,
            sheet_name=0,
            header=None,
            dtype=object,
            engine=_obter_engine_excel_fecoagroRS(caminho_arquivo),
        )
    except Exception as erro:
        raise ValueError(
            f"Nao foi possivel ler o arquivo FecoagroRS "
            f"'{nome_arquivo}'. Erro: {erro}"
        ) from erro

    if dataframe.empty:
        raise ValueError(f"O arquivo FecoagroRS '{nome_arquivo}' esta vazio.")

    if dataframe.shape[1] < 11:
        raise ValueError(
            f"O arquivo FecoagroRS '{nome_arquivo}' possui "
            f"{dataframe.shape[1]} coluna(s), mas sao necessarias "
            "pelo menos 11 colunas, de A ate K."
        )

    return dataframe


# ==============================================================================
# IDENTIFICACAO DO CABECALHO E DAS LINHAS CONTABEIS
# ==============================================================================

def _linha_eh_cabecalho_fecoagroRS(linha):
    """Identifica a linha de titulos do balancete."""
    valores = [
        _normalizar_texto_comparacao_fecoagroRS(valor)
        for valor in linha.tolist()
    ]
    texto_completo = " | ".join(valor for valor in valores if valor)

    termos = [
        "CODIGO",
        "CLASSIFICACAO",
        "DESCRICAO DA CONTA",
        "SALDO ANTERIOR",
        "DEBITO",
        "CREDITO",
        "SALDO ATUAL",
    ]

    return sum(termo in texto_completo for termo in termos) >= 4


def _linha_possui_dados_fecoagroRS(linha):
    """Valida codigo, classificacao e descricao nas colunas informadas."""
    if len(linha) < 11:
        return False

    codigo = _normalizar_codigo_reduzido_fecoagroRS(
        linha.iloc[INDICE_CODIGO_REDUZIDO_fecoagroRS]
    )
    classificacao = _normalizar_classificacao_fecoagroRS(
        linha.iloc[INDICE_CLASSIFICACAO_fecoagroRS]
    )
    nome = _normalizar_texto_fecoagroRS(
        linha.iloc[INDICE_NOME_fecoagroRS]
    )

    return bool(codigo and classificacao and nome)


def _localizar_inicio_dados_fecoagroRS(dataframe, nome_arquivo):
    """Localiza a linha seguinte ao cabecalho ou a primeira conta valida."""
    for indice in range(dataframe.shape[0]):
        if _linha_eh_cabecalho_fecoagroRS(dataframe.iloc[indice]):
            return indice + 1

    for indice in range(dataframe.shape[0]):
        if _linha_possui_dados_fecoagroRS(dataframe.iloc[indice]):
            return indice

    raise ValueError(
        f"Nao foi possivel localizar o inicio do balancete "
        f"no arquivo fecoagroRS '{nome_arquivo}'."
    )


# ==============================================================================
# EXTRACAO E MONTAGEM
# ==============================================================================

def _extrair_registro_fecoagroRS(linha):
    """Extrai A, B, D, F, G, I e K conforme o de-para informado."""
    if not _linha_possui_dados_fecoagroRS(linha):
        return None

    codigo_reduzido = _normalizar_codigo_reduzido_fecoagroRS(
        linha.iloc[INDICE_CODIGO_REDUZIDO_fecoagroRS]
    )
    classificacao = _normalizar_classificacao_fecoagroRS(
        linha.iloc[INDICE_CLASSIFICACAO_fecoagroRS]
    )
    nome = _normalizar_texto_fecoagroRS(
        linha.iloc[INDICE_NOME_fecoagroRS]
    )

    saldo_anterior = _aplicar_natureza_fecoagroRS(
        linha.iloc[INDICE_SALDO_ANTERIOR_fecoagroRS]
    )
    debito = _aplicar_natureza_fecoagroRS(
        linha.iloc[INDICE_DEBITO_fecoagroRS],
        natureza_padrao="D",
    )
    credito = _aplicar_natureza_fecoagroRS(
        linha.iloc[INDICE_CREDITO_fecoagroRS],
        natureza_padrao="C",
    )
    saldo_acumulado = _aplicar_natureza_fecoagroRS(
        linha.iloc[INDICE_SALDO_ACUMULADO_fecoagroRS]
    )
    movimento = round(debito + credito, 2)

    return {
        "Conta": classificacao,
        "Nome": nome,
        "Cód. Reduzido": codigo_reduzido,
        "Saldo Anterior": saldo_anterior,
        "Débito": debito,
        "Crédito": credito,
        "Movimento": movimento,
        "Saldo Acumulado": saldo_acumulado,
    }


def _montar_dataframe_fecoagroRS(registros):
    """Monta o DataFrame final na ordem padrao do tabulador."""
    if not registros:
        return pd.DataFrame(columns=COLUNAS_DESTINO_fecoagroRS)

    dataframe = pd.DataFrame(registros)
    dataframe.insert(0, "Atividade", "Geral")

    for coluna in [
        "Saldo Anterior",
        "Débito",
        "Crédito",
        "Movimento",
        "Saldo Acumulado",
    ]:
        dataframe[coluna] = pd.to_numeric(
            dataframe[coluna],
            errors="coerce",
        ).fillna(0.0).round(2)

    dataframe["Movimento"] = (
        dataframe["Débito"] + dataframe["Crédito"]
    ).round(2)

    return dataframe[COLUNAS_DESTINO_fecoagroRS].copy()


# ==============================================================================
# TRANSFORMACAO E FUNCAO PUBLICA
# ==============================================================================

def transformar_balancete_fecoagroRS(caminho_arquivo):
    """Transforma um arquivo Excel do sistema FecoagroRS."""
    nome_arquivo = os.path.basename(caminho_arquivo)
    extensao = os.path.splitext(caminho_arquivo)[1].lower()

    if extensao not in EXTENSOES_VALIDAS_fecoagroRS:
        raise ValueError(f"O arquivo '{nome_arquivo}' nao e um Excel valido.")

    dataframe_origem = _ler_excel_fecoagroRS(caminho_arquivo)
    indice_inicio = _localizar_inicio_dados_fecoagroRS(
        dataframe_origem,
        nome_arquivo,
    )
    dataframe_dados = dataframe_origem.iloc[indice_inicio:].copy()
    registros = []

    for _, linha in dataframe_dados.iterrows():
        if _linha_eh_cabecalho_fecoagroRS(linha):
            continue

        registro = _extrair_registro_fecoagroRS(linha)
        if registro is not None:
            registros.append(registro)

    if not registros:
        raise ValueError(
            f"Nenhuma conta valida foi encontrada no arquivo FecoagroRS "
            f"'{nome_arquivo}'. Foram utilizadas as colunas A, B, D, "
            "F, G, I e K."
        )

    return _montar_dataframe_fecoagroRS(registros)


def _chave_ordenacao_arquivo_fecoagroRS(caminho_arquivo):
    """Ordena arquivos B_XX pelo mes e os demais pelo nome."""
    mes = _obter_mes_fecoagroRS(caminho_arquivo)
    numero_mes = int(mes) if mes else 99
    return numero_mes, os.path.basename(caminho_arquivo).casefold()


def processar(lista_arquivos):
    """Processa todos os arquivos selecionados para o FecoagroRS."""
    if not lista_arquivos:
        raise ValueError(
            "Nenhum arquivo foi selecionado para o sistema FecoagroRS."
        )

    arquivos_excel = [
        arquivo
        for arquivo in lista_arquivos
        if os.path.splitext(str(arquivo))[1].lower()
        in EXTENSOES_VALIDAS_fecoagroRS
    ]

    if not arquivos_excel:
        raise ValueError(
            "Nenhum arquivo Excel valido foi encontrado para o FecoagroRS."
        )

    resultados = {}
    nomes_utilizados = set()

    for arquivo in sorted(
        arquivos_excel,
        key=_chave_ordenacao_arquivo_fecoagroRS,
    ):
        nome_aba = _gerar_nome_aba_fecoagroRS(
            arquivo,
            nomes_utilizados,
        )
        dataframe = transformar_balancete_fecoagroRS(arquivo)

        if dataframe is None or dataframe.empty:
            raise ValueError(
                f"O arquivo '{os.path.basename(arquivo)}' nao retornou "
                "dados validos."
            )

        resultados[nome_aba] = dataframe

    if not resultados:
        raise ValueError(
            "Nenhum resultado foi gerado para o sistema FecoagroRS."
        )

    return resultados
