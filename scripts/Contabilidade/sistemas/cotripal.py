import os
import re
import unicodedata

import pandas as pd


# ==============================================================================
# CONFIGURACOES DO SISTEMA Cotripal
# ==============================================================================

EXTENSOES_VALIDAS_cotripal = {".xls", ".xlsx"}

COLUNAS_DESTINO_cotripal = [
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

# O arquivo possui uma coluna A vazia. Os dados começam na coluna B.
INDICE_CLASSIFICACAO_cotripal = 1       # B - NAT-CONTAB
INDICE_CODIGO_REDUZIDO_cotripal = 2     # C - CT-CONTAB
INDICE_NOME_cotripal = 3                # D - NOME
INDICE_SALDO_ANTERIOR_cotripal = 4      # E - SALDO ANTERIOR
INDICE_NATUREZA_ANTERIOR_cotripal = 5   # F - D/C SL.ANTERIOR
INDICE_DEBITO_cotripal = 6              # G - DEBITOS
INDICE_CREDITO_cotripal = 7             # H - CREDITOS
INDICE_SALDO_FINAL_cotripal = 8         # I - SALDO FINAL
INDICE_NATUREZA_FINAL_cotripal = 9      # J - D/C SL.FINAL


# ==============================================================================
# NOMES DAS ABAS DE DESTINO
# ==============================================================================

def _limpar_nome_aba_cotripal(nome):
    """Ajusta um texto para utilização como nome de aba do Excel."""
    nome_limpo = re.sub(
        r'[\\/\x2a?:\[\]]',
        "_",
        str(nome).strip(),
    )
    return (nome_limpo or "Sem nome")[:31]


def _obter_nome_sem_extensao_cotripal(caminho_arquivo):
    """Retorna o nome do arquivo sem a extensão."""
    nome_arquivo = os.path.basename(caminho_arquivo)
    return os.path.splitext(nome_arquivo)[0].strip()


def _obter_mes_cotripal(caminho_arquivo):
    """Extrai o mês quando o arquivo começa com B_XX."""
    nome_arquivo = os.path.basename(caminho_arquivo)
    correspondencia = re.match(
        r"^B_(0[1-9]|1[0-2])",
        nome_arquivo,
        re.IGNORECASE,
    )
    return correspondencia.group(1) if correspondencia else None


def _registrar_nome_aba_cotripal(nome, nomes_utilizados):
    """Registra um nome de aba se ainda não estiver em uso."""
    nome_aba = _limpar_nome_aba_cotripal(nome)
    chave = nome_aba.casefold()

    if chave in nomes_utilizados:
        return None

    nomes_utilizados.add(chave)
    return nome_aba


def _gerar_nome_aba_cotripal(caminho_arquivo, nomes_utilizados):
    """Gera nome exclusivo, priorizando XX em arquivos B_XX."""
    nome_completo = _obter_nome_sem_extensao_cotripal(caminho_arquivo)
    mes = _obter_mes_cotripal(caminho_arquivo)

    if mes:
        nome_aba = _registrar_nome_aba_cotripal(mes, nomes_utilizados)
        if nome_aba is not None:
            return nome_aba

    nome_aba = _registrar_nome_aba_cotripal(
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
        nome_aba = _registrar_nome_aba_cotripal(
            candidato,
            nomes_utilizados,
        )

        if nome_aba is not None:
            return nome_aba

        contador += 1


# ==============================================================================
# NORMALIZACAO DE TEXTOS E CODIGOS
# ==============================================================================

def _normalizar_texto_cotripal(valor):
    """Normaliza espaços e converte valores ausentes em texto vazio."""
    if valor is None:
        return ""

    try:
        if pd.isna(valor):
            return ""
    except (TypeError, ValueError):
        pass

    texto = str(valor).replace("\xa0", " ")
    return re.sub(r"\s{1,}", " ", texto).strip()


def _normalizar_texto_comparacao_cotripal(valor):
    """Converte para maiúsculas e remove acentos."""
    texto = _normalizar_texto_cotripal(valor).upper()
    if not texto:
        return ""

    texto = unicodedata.normalize("NFKD", texto)
    return "".join(
        caractere
        for caractere in texto
        if not unicodedata.combining(caractere)
    )


def _normalizar_codigo_com_pontos_cotripal(valor):
    """
    Remove pontos e espaços de classificações ou códigos reduzidos.

    Exemplos:
        1.2.02.01 -> 120201
        1010.3    -> 10103
        1.9       -> 19
    """
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
        texto = _normalizar_texto_cotripal(valor)

    texto = (
        texto.replace("\xa0", "")
        .replace(" ", "")
        .replace(",", ".")
        .replace(".", "")
    )

    return texto if texto.isdigit() else ""


def _normalizar_classificacao_cotripal(valor):
    """Normaliza a classificação NAT-CONTAB."""
    return _normalizar_codigo_com_pontos_cotripal(valor)


def _normalizar_codigo_reduzido_cotripal(valor):
    """Normaliza o código reduzido CT-CONTAB."""
    return _normalizar_codigo_com_pontos_cotripal(valor)


# ==============================================================================
# CONVERSAO DE VALORES E NATUREZA D/C
# ==============================================================================

def _normalizar_natureza_cotripal(natureza):
    """Normaliza a natureza separada para D, C ou vazio."""
    texto = (
        _normalizar_texto_comparacao_cotripal(natureza)
        .replace(".", "")
        .replace("/", "")
        .replace("\\", "")
        .strip()
    )

    if texto in {"D", "DEBITO", "DEVEDOR"}:
        return "D"
    if texto in {"C", "CREDITO", "CREDOR"}:
        return "C"
    return ""


def _converter_numero_cotripal(valor):
    """Converte números brasileiros e números reais do Excel para float."""
    if valor is None:
        return 0.0

    try:
        if pd.isna(valor):
            return 0.0
    except (TypeError, ValueError):
        pass

    if isinstance(valor, (int, float)):
        return round(float(valor), 2)

    texto = str(valor).strip()
    if texto in {"", "-", "--"}:
        return 0.0

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


def _aplicar_natureza_cotripal(valor, natureza):
    """Aplica D como positivo e C como negativo."""
    numero = _converter_numero_cotripal(valor)
    natureza_normalizada = _normalizar_natureza_cotripal(natureza)

    if natureza_normalizada == "D":
        return round(abs(numero), 2)
    if natureza_normalizada == "C":
        return round(-abs(numero), 2)

    return round(numero, 2)


# ==============================================================================
# LEITURA DO EXCEL
# ==============================================================================

def _obter_engine_excel_cotripal(caminho_arquivo):
    """Retorna o mecanismo adequado para XLS ou XLSX."""
    extensao = os.path.splitext(caminho_arquivo)[1].lower()
    return "xlrd" if extensao == ".xls" else "openpyxl"


def _ler_excel_cotripal(caminho_arquivo):
    """Lê a primeira aba sem uma linha fixa de cabeçalho."""
    nome_arquivo = os.path.basename(caminho_arquivo)

    try:
        dataframe = pd.read_excel(
            caminho_arquivo,
            sheet_name=0,
            header=None,
            dtype=object,
            engine=_obter_engine_excel_cotripal(caminho_arquivo),
        )
    except Exception as erro:
        raise ValueError(
            f"Não foi possível ler o arquivo Cotripal "
            f"'{nome_arquivo}'. Erro: {erro}"
        ) from erro

    if dataframe.empty:
        raise ValueError(f"O arquivo Cotripal '{nome_arquivo}' está vazio.")

    if dataframe.shape[1] < 10:
        raise ValueError(
            f"O arquivo Cotripal '{nome_arquivo}' possui "
            f"{dataframe.shape[1]} coluna(s), mas são necessárias "
            "pelo menos 10 colunas, de A até J."
        )

    return dataframe


# ==============================================================================
# IDENTIFICACAO DO CABECALHO E DAS LINHAS
# ==============================================================================

def _linha_eh_cabecalho_cotripal(linha):
    """Identifica a linha NAT-CONTAB, CT-CONTAB, NOME e saldos."""
    valores = [
        _normalizar_texto_comparacao_cotripal(valor)
        for valor in linha.tolist()
    ]
    texto_completo = " | ".join(valor for valor in valores if valor)

    termos = [
        "NAT-CONTAB",
        "CT-CONTAB",
        "NOME",
        "SALDO ANTERIOR",
        "DEBITOS",
        "CREDITOS",
        "SALDO FINAL",
    ]

    return sum(termo in texto_completo for termo in termos) >= 4


def _linha_possui_dados_cotripal(linha):
    """Valida classificação, código reduzido e nome."""
    if len(linha) < 10:
        return False

    classificacao = _normalizar_classificacao_cotripal(
        linha.iloc[INDICE_CLASSIFICACAO_cotripal]
    )
    codigo_reduzido = _normalizar_codigo_reduzido_cotripal(
        linha.iloc[INDICE_CODIGO_REDUZIDO_cotripal]
    )
    nome = _normalizar_texto_cotripal(
        linha.iloc[INDICE_NOME_cotripal]
    )

    return bool(classificacao and codigo_reduzido and nome)


def _localizar_inicio_dados_cotripal(dataframe, nome_arquivo):
    """Localiza o cabeçalho ou a primeira conta válida."""
    for indice in range(dataframe.shape[0]):
        if _linha_eh_cabecalho_cotripal(dataframe.iloc[indice]):
            return indice + 1

    for indice in range(dataframe.shape[0]):
        if _linha_possui_dados_cotripal(dataframe.iloc[indice]):
            return indice

    raise ValueError(
        f"Não foi possível localizar o início do balancete "
        f"no arquivo Cotripal '{nome_arquivo}'."
    )


# ==============================================================================
# EXTRACAO E MONTAGEM
# ==============================================================================

def _extrair_registro_cotripal(linha):
    """Extrai classificação, reduzido, nome, saldos e movimentos."""
    if not _linha_possui_dados_cotripal(linha):
        return None

    classificacao = _normalizar_classificacao_cotripal(
        linha.iloc[INDICE_CLASSIFICACAO_cotripal]
    )
    codigo_reduzido = _normalizar_codigo_reduzido_cotripal(
        linha.iloc[INDICE_CODIGO_REDUZIDO_cotripal]
    )
    nome = _normalizar_texto_cotripal(
        linha.iloc[INDICE_NOME_cotripal]
    )

    saldo_anterior = _aplicar_natureza_cotripal(
        linha.iloc[INDICE_SALDO_ANTERIOR_cotripal],
        linha.iloc[INDICE_NATUREZA_ANTERIOR_cotripal],
    )

    debito = round(
        abs(_converter_numero_cotripal(linha.iloc[INDICE_DEBITO_cotripal])),
        2,
    )

    credito = round(
        -abs(_converter_numero_cotripal(linha.iloc[INDICE_CREDITO_cotripal])),
        2,
    )

    movimento = round(debito + credito, 2)

    saldo_acumulado = _aplicar_natureza_cotripal(
        linha.iloc[INDICE_SALDO_FINAL_cotripal],
        linha.iloc[INDICE_NATUREZA_FINAL_cotripal],
    )

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


def _montar_dataframe_cotripal(registros):
    """Monta o DataFrame final na ordem padrão do tabulador."""
    if not registros:
        return pd.DataFrame(columns=COLUNAS_DESTINO_cotripal)

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

    dataframe["Débito"] = dataframe["Débito"].abs().round(2)
    dataframe["Crédito"] = (-dataframe["Crédito"].abs()).round(2)
    dataframe["Movimento"] = (
        dataframe["Débito"] + dataframe["Crédito"]
    ).round(2)

    return dataframe[COLUNAS_DESTINO_cotripal].copy()


# ==============================================================================
# TRANSFORMACAO E FUNCAO PUBLICA
# ==============================================================================

def transformar_balancete_cotripal(caminho_arquivo):
    """Transforma um arquivo Excel do sistema Cotripal."""
    nome_arquivo = os.path.basename(caminho_arquivo)
    extensao = os.path.splitext(caminho_arquivo)[1].lower()

    if extensao not in EXTENSOES_VALIDAS_cotripal:
        raise ValueError(f"O arquivo '{nome_arquivo}' não é um Excel válido.")

    dataframe_origem = _ler_excel_cotripal(caminho_arquivo)
    indice_inicio = _localizar_inicio_dados_cotripal(
        dataframe_origem,
        nome_arquivo,
    )
    dataframe_dados = dataframe_origem.iloc[indice_inicio:].copy()
    registros = []

    for _, linha in dataframe_dados.iterrows():
        if _linha_eh_cabecalho_cotripal(linha):
            continue

        registro = _extrair_registro_cotripal(linha)
        if registro is not None:
            registros.append(registro)

    if not registros:
        raise ValueError(
            f"Nenhuma conta válida foi encontrada no arquivo Cotripal "
            f"'{nome_arquivo}'."
        )

    return _montar_dataframe_cotripal(registros)


def _chave_ordenacao_arquivo_cotripal(caminho_arquivo):
    """Ordena arquivos B_XX pelo mês e os demais pelo nome."""
    mes = _obter_mes_cotripal(caminho_arquivo)
    numero_mes = int(mes) if mes else 99
    return numero_mes, os.path.basename(caminho_arquivo).casefold()


def processar(lista_arquivos):
    """Processa todos os arquivos selecionados para o Cotripal."""
    if not lista_arquivos:
        raise ValueError(
            "Nenhum arquivo foi selecionado para o sistema Cotripal."
        )

    arquivos_excel = [
        arquivo
        for arquivo in lista_arquivos
        if os.path.splitext(str(arquivo))[1].lower()
        in EXTENSOES_VALIDAS_cotripal
    ]

    if not arquivos_excel:
        raise ValueError(
            "Nenhum arquivo Excel válido foi encontrado para o Cotripal."
        )

    resultados = {}
    nomes_utilizados = set()

    for arquivo in sorted(
        arquivos_excel,
        key=_chave_ordenacao_arquivo_cotripal,
    ):
        nome_aba = _gerar_nome_aba_cotripal(
            arquivo,
            nomes_utilizados,
        )
        dataframe = transformar_balancete_cotripal(arquivo)

        if dataframe is None or dataframe.empty:
            raise ValueError(
                f"O arquivo '{os.path.basename(arquivo)}' não retornou "
                "dados válidos."
            )

        resultados[nome_aba] = dataframe

    if not resultados:
        raise ValueError(
            "Nenhum resultado foi gerado para o sistema Cotripal."
        )

    return resultados
