import os
import re
import unicodedata

import pandas as pd


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

# Índices Python iniciados em zero.
INDICE_CODIGO_ORIGEM_fecoagroRS = 0       # Coluna A
INDICE_CLASSIFICACAO_fecoagroRS = 1       # Coluna B
INDICE_NOME_fecoagroRS = 3                # Coluna D
INDICE_SALDO_ANTERIOR_fecoagroRS = 5      # Coluna F
INDICE_DEBITO_fecoagroRS = 6              # Coluna G
INDICE_CREDITO_fecoagroRS = 8             # Coluna I
INDICE_SALDO_ACUMULADO_fecoagroRS = 10    # Coluna K


def _limpar_nome_aba_fecoagroRS(nome):
    """Ajusta um texto para utilização como nome de aba do Excel."""
    nome_limpo = re.sub(
        r'[\\/\x2a?:\[\]]',
        "_",
        str(nome).strip(),
    )
    return (nome_limpo or "Sem nome")[:31]


def _obter_nome_sem_extensao_fecoagroRS(caminho_arquivo):
    """Retorna o nome do arquivo sem a extensão."""
    nome_arquivo = os.path.basename(caminho_arquivo)
    return os.path.splitext(nome_arquivo)[0].strip()


def _obter_mes_fecoagroRS(caminho_arquivo):
    """Extrai XX quando o arquivo começa com B_XX."""
    nome_arquivo = os.path.basename(caminho_arquivo)
    correspondencia = re.match(
        r"^B_(0[1-9]|1[0-2])",
        nome_arquivo,
        re.IGNORECASE,
    )
    return correspondencia.group(1) if correspondencia else None


def _registrar_nome_aba_fecoagroRS(nome, nomes_utilizados):
    """Registra um nome de aba se ainda não estiver em uso."""
    nome_aba = _limpar_nome_aba_fecoagroRS(nome)
    chave = nome_aba.casefold()

    if chave in nomes_utilizados:
        return None

    nomes_utilizados.add(chave)
    return nome_aba


def _gerar_nome_aba_fecoagroRS(caminho_arquivo, nomes_utilizados):
    """Gera um nome exclusivo, priorizando XX em arquivos B_XX."""
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


def _normalizar_texto_fecoagroRS(valor):
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


def _normalizar_texto_comparacao_fecoagroRS(valor):
    """Converte para maiúsculas e remove acentos para comparação."""
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
    """Normaliza a classificação da coluna B preservando os pontos."""
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

    texto = (
        _normalizar_texto_fecoagroRS(valor)
        .replace("\xa0", "")
        .replace(" ", "")
        .replace(",", ".")
    )

    if re.fullmatch(r"\d+\.0+", texto):
        return texto.split(".", maxsplit=1)[0]

    if not re.fullmatch(r"\d+(?:\.\d+){0,}", texto):
        return ""

    return texto


def _normalizar_codigo_origem_fecoagroRS(valor):
    """Normaliza o código da coluna A usado nas contas analíticas."""
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


def _extrair_natureza_fecoagroRS(valor):
    """Extrai D ou C existente no final da célula monetária."""
    texto = _normalizar_texto_fecoagroRS(valor).upper()
    correspondencia = re.search(
        r"([DC])\s{0,}$",
        texto,
        re.IGNORECASE,
    )
    return correspondencia.group(1).upper() if correspondencia else ""


def _converter_numero_fecoagroRS(valor):
    """Converte valor monetário e remove D/C do final da célula."""
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
    """Aplica a natureza da célula ou a natureza padrão da coluna."""
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


def _obter_engine_excel_fecoagroRS(caminho_arquivo):
    """Retorna o mecanismo adequado para XLS ou XLSX."""
    extensao = os.path.splitext(caminho_arquivo)[1].lower()
    return "xlrd" if extensao == ".xls" else "openpyxl"


def _ler_excel_fecoagroRS(caminho_arquivo):
    """Lê a primeira aba sem uma linha fixa de cabeçalho."""
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
            f"Não foi possível ler o arquivo fecoagroRS "
            f"'{nome_arquivo}'. Erro: {erro}"
        ) from erro

    if dataframe.empty:
        raise ValueError(f"O arquivo fecoagroRS '{nome_arquivo}' está vazio.")

    if dataframe.shape[1] < 11:
        raise ValueError(
            f"O arquivo fecoagroRS '{nome_arquivo}' possui "
            f"{dataframe.shape[1]} coluna(s), mas são necessárias "
            "pelo menos 11 colunas, de A até K."
        )

    return dataframe


def _linha_eh_cabecalho_fecoagroRS(linha):
    """Identifica a linha de títulos do balancete."""
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
    """Valida código A, classificação B e nome D."""
    if len(linha) < 11:
        return False

    codigo_origem = _normalizar_codigo_origem_fecoagroRS(
        linha.iloc[INDICE_CODIGO_ORIGEM_fecoagroRS]
    )
    classificacao = _normalizar_classificacao_fecoagroRS(
        linha.iloc[INDICE_CLASSIFICACAO_fecoagroRS]
    )
    nome = _normalizar_texto_fecoagroRS(
        linha.iloc[INDICE_NOME_fecoagroRS]
    )

    return bool(codigo_origem and classificacao and nome)


def _localizar_inicio_dados_fecoagroRS(dataframe, nome_arquivo):
    """Localiza a linha seguinte ao cabeçalho ou a primeira conta válida."""
    for indice in range(dataframe.shape[0]):
        if _linha_eh_cabecalho_fecoagroRS(dataframe.iloc[indice]):
            return indice + 1

    for indice in range(dataframe.shape[0]):
        if _linha_possui_dados_fecoagroRS(dataframe.iloc[indice]):
            return indice

    raise ValueError(
        f"Não foi possível localizar o início do balancete "
        f"no arquivo fecoagroRS '{nome_arquivo}'."
    )


def _extrair_registro_fecoagroRS(linha):
    """Extrai os dados e guarda temporariamente o código da coluna A."""
    if not _linha_possui_dados_fecoagroRS(linha):
        return None

    codigo_origem = _normalizar_codigo_origem_fecoagroRS(
        linha.iloc[INDICE_CODIGO_ORIGEM_fecoagroRS]
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
        "Cód. Reduzido": classificacao,
        "Saldo Anterior": saldo_anterior,
        "Débito": debito,
        "Crédito": credito,
        "Movimento": movimento,
        "Saldo Acumulado": saldo_acumulado,
        "_Código Origem": codigo_origem,
    }


def _conta_eh_sintetica_fecoagroRS(
    classificacao,
    classificacoes_existentes,
):
    """Uma conta é sintética quando possui classificação subordinada."""
    classificacao = _normalizar_classificacao_fecoagroRS(classificacao)

    if not classificacao:
        return False

    prefixo_subconta = classificacao + "."

    return any(
        outra_classificacao != classificacao
        and outra_classificacao.startswith(prefixo_subconta)
        for outra_classificacao in classificacoes_existentes
    )


def _ajustar_contas_analiticas_fecoagroRS(registros):
    """
    Mantém as contas sintéticas e concatena o código A nas analíticas.

    Sintética:
        Conta = classificação original.

    Analítica:
        Conta = classificação + " - " + código da coluna A.

    Cód. Reduzido:
        permanece com a classificação original da coluna B.
    """
    if not registros:
        return registros

    classificacoes_existentes = [
        _normalizar_classificacao_fecoagroRS(
            registro.get("Cód. Reduzido", "")
        )
        for registro in registros
    ]
    classificacoes_existentes = [
        classificacao
        for classificacao in classificacoes_existentes
        if classificacao
    ]

    for registro in registros:
        classificacao = _normalizar_classificacao_fecoagroRS(
            registro.get("Cód. Reduzido", "")
        )
        codigo_origem = _normalizar_codigo_origem_fecoagroRS(
            registro.get("_Código Origem", "")
        )

        if not classificacao:
            registro["Conta"] = ""
        elif _conta_eh_sintetica_fecoagroRS(
            classificacao,
            classificacoes_existentes,
        ):
            registro["Conta"] = classificacao
        elif codigo_origem:
            registro["Conta"] = classificacao + " - " + codigo_origem
        else:
            registro["Conta"] = classificacao

    return registros


def _montar_dataframe_fecoagroRS(registros):
    """Monta o DataFrame final e remove o campo auxiliar."""
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

    if "_Código Origem" in dataframe.columns:
        dataframe.drop(columns=["_Código Origem"], inplace=True)

    return dataframe[COLUNAS_DESTINO_fecoagroRS].copy()


def transformar_balancete_fecoagroRS(caminho_arquivo):
    """Transforma um arquivo Excel do sistema fecoagroRS."""
    nome_arquivo = os.path.basename(caminho_arquivo)
    extensao = os.path.splitext(caminho_arquivo)[1].lower()

    if extensao not in EXTENSOES_VALIDAS_fecoagroRS:
        raise ValueError(f"O arquivo '{nome_arquivo}' não é um Excel válido.")

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
            f"Nenhuma conta válida foi encontrada no arquivo fecoagroRS "
            f"'{nome_arquivo}'. Foram utilizadas as colunas A, B, D, "
            "F, G, I e K."
        )

    registros = _ajustar_contas_analiticas_fecoagroRS(registros)
    return _montar_dataframe_fecoagroRS(registros)


def _chave_ordenacao_arquivo_fecoagroRS(caminho_arquivo):
    """Ordena arquivos B_XX pelo mês e os demais pelo nome."""
    mes = _obter_mes_fecoagroRS(caminho_arquivo)
    numero_mes = int(mes) if mes else 99
    return numero_mes, os.path.basename(caminho_arquivo).casefold()


def processar(lista_arquivos):
    """Processa todos os arquivos selecionados para o fecoagroRS."""
    if not lista_arquivos:
        raise ValueError(
            "Nenhum arquivo foi selecionado para o sistema fecoagroRS."
        )

    arquivos_excel = [
        arquivo
        for arquivo in lista_arquivos
        if os.path.splitext(str(arquivo))[1].lower()
        in EXTENSOES_VALIDAS_fecoagroRS
    ]

    if not arquivos_excel:
        raise ValueError(
            "Nenhum arquivo Excel válido foi encontrado para o fecoagroRS."
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
                f"O arquivo '{os.path.basename(arquivo)}' não retornou "
                "dados válidos."
            )

        resultados[nome_aba] = dataframe

    if not resultados:
        raise ValueError(
            "Nenhum resultado foi gerado para o sistema fecoagroRS."
        )

    return resultados
