import os

import pandas as pd


# ==============================================================================
# FUNÇÕES AUXILIARES
# ==============================================================================

def _converter_valor_monetario(valor):
    """
    Converte um valor monetário para float.

    Reconhece indicadores contábeis:
        D = valor positivo
        C = valor negativo

    Exemplos:
        1.250,50D -> 1250.50
        1.250,50C -> -1250.50
        vazio     -> 0.0
    """
    if pd.isna(valor) or str(valor).strip() == "":
        return 0.0

    texto = str(valor).strip().upper()
    sinal = 1

    if texto.endswith("D"):
        numero_str = texto[:-1].strip()
        sinal = 1

    elif texto.endswith("C"):
        numero_str = texto[:-1].strip()
        sinal = -1

    else:
        numero_str = texto

    try:
        numero_str = (
            numero_str
            .replace(".", "")
            .replace(",", ".")
        )

        return float(numero_str) * sinal

    except (ValueError, TypeError):
        return 0.0


def _converter_serie_decimal_brasileiro(serie):
    """
    Converte uma série do Pandas com números no formato brasileiro.

    Exemplos:
        388788112,3 -> 388788112.3
        1.250,50    -> 1250.50
        vazio       -> 0.0
    """
    texto = (
        serie
        .fillna("")
        .astype(str)
        .str.strip()
        .str.replace(".", "", regex=False)
        .str.replace(",", ".", regex=False)
    )

    return pd.to_numeric(
        texto,
        errors="coerce"
    ).fillna(0.0)


def _obter_engine_excel(caminho_arquivo):
    """
    Define o mecanismo de leitura conforme a extensão do arquivo.

    Arquivos .xls utilizam xlrd.
    Arquivos .xlsx utilizam openpyxl.
    """
    extensao = os.path.splitext(caminho_arquivo)[1].lower()

    if extensao == ".xls":
        return "xlrd"

    return "openpyxl"


def _ler_csv_useall(caminho_arquivo):
    """
    Lê um arquivo CSV exportado pelo Useall.

    O arquivo deve utilizar ponto e vírgula como separador.

    A função tenta as seguintes codificações:
        1. UTF-8 com possível BOM;
        2. Latin-1;
        3. Windows-1252.
    """
    nome_arquivo = os.path.basename(caminho_arquivo)

    configuracao_leitura = {
        "sep": ";",
        "dtype": str,
        "keep_default_na": False,
    }

    codificacoes = [
        "utf-8-sig",
        "latin-1",
        "cp1252",
    ]

    ultimo_erro = None

    for codificacao in codificacoes:
        try:
            return pd.read_csv(
                caminho_arquivo,
                encoding=codificacao,
                **configuracao_leitura
            )

        except UnicodeDecodeError as erro:
            ultimo_erro = erro

        except Exception as erro:
            raise ValueError(
                f"Não foi possível ler o arquivo CSV "
                f"'{nome_arquivo}'. Erro: {erro}"
            ) from erro

    raise ValueError(
        f"Não foi possível identificar a codificação do arquivo CSV "
        f"'{nome_arquivo}'. Erro: {ultimo_erro}"
    )


# ==============================================================================
# TRANSFORMAÇÃO COMUM DOS BALANCETES EM EXCEL
# ==============================================================================

def transformar_balancete(caminho_arquivo):
    """
    Transforma um balancete no layout comum utilizado pelos sistemas
    que trabalham com arquivos Excel.
    """
    nome_arquivo = os.path.basename(caminho_arquivo)

    try:
        df_origem = pd.read_excel(
            caminho_arquivo,
            sheet_name=0,
            dtype=str,
            engine=_obter_engine_excel(caminho_arquivo)
        )

        df_origem.dropna(
            how="all",
            inplace=True
        )

    except Exception as erro:
        raise ValueError(
            f"Não foi possível ler o arquivo {nome_arquivo}. "
            f"Erro: {erro}"
        ) from erro

    if df_origem.shape[1] < 6:
        raise ValueError(
            f"O arquivo '{nome_arquivo}' possui "
            f"{df_origem.shape[1]} coluna(s), mas são necessárias "
            f"pelo menos 6 colunas para transformar o balancete."
        )

    df_origem.reset_index(
        drop=True,
        inplace=True
    )

    df_destino = pd.DataFrame(
        index=df_origem.index
    )

    df_destino["Atividade"] = "Geral"

    df_destino["Conta"] = (
        df_origem.iloc[:, 0]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    df_destino["Nome"] = (
        df_origem.iloc[:, 1]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    df_destino["Cód. Reduzido"] = (
        df_origem.iloc[:, 0]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    saldo_anterior = pd.to_numeric(
        df_origem.iloc[:, 2]
        .fillna("")
        .astype(str)
        .str.replace(",", ".", regex=False),
        errors="coerce"
    ).fillna(0.0)

    debito = pd.to_numeric(
        df_origem.iloc[:, 3]
        .fillna("")
        .astype(str)
        .str.replace(",", ".", regex=False),
        errors="coerce"
    ).fillna(0.0)

    credito = pd.to_numeric(
        df_origem.iloc[:, 4]
        .fillna("")
        .astype(str)
        .str.replace(",", ".", regex=False),
        errors="coerce"
    ).fillna(0.0)

    saldo_acumulado_origem = pd.to_numeric(
        df_origem.iloc[:, 5]
        .fillna("")
        .astype(str)
        .str.replace(",", ".", regex=False),
        errors="coerce"
    ).fillna(0.0)

    df_destino["Saldo Anterior"] = -saldo_anterior
    df_destino["Débito"] = debito
    df_destino["Crédito"] = credito
    df_destino["Movimento"] = debito - credito
    df_destino["Saldo Acumulado"] = -saldo_acumulado_origem

    return df_destino


# ==============================================================================
# TRANSFORMAÇÃO DO PLANO DE CONTAS
# ==============================================================================

def transformar_plano_de_contas(caminho_arquivo):
    """
    Transforma um plano de contas no layout esperado pelo HUB.
    """
    nome_arquivo = os.path.basename(caminho_arquivo)

    try:
        df_origem = pd.read_excel(
            caminho_arquivo,
            sheet_name=0,
            skiprows=5,
            dtype=str,
            engine=_obter_engine_excel(caminho_arquivo)
        )

        df_origem.dropna(
            how="all",
            inplace=True
        )

    except Exception as erro:
        raise ValueError(
            f"Não foi possível ler o plano de contas "
            f"{nome_arquivo}. Erro: {erro}"
        ) from erro

    if df_origem.shape[1] < 4:
        raise ValueError(
            f"O plano de contas '{nome_arquivo}' possui "
            f"{df_origem.shape[1]} coluna(s), mas são necessárias "
            f"pelo menos 4 colunas para realizar a transformação."
        )

    df_origem.reset_index(
        drop=True,
        inplace=True
    )

    df_destino = pd.DataFrame(
        index=df_origem.index
    )

    classif = (
        df_origem.iloc[:, 2]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    df_destino["Chave Cliente"] = classif

    df_destino["Chave D&M"] = (
        classif
        .str[0]
        .where(
            classif.str[0].isin(["1", "2", "3"]),
            ""
        )
    )

    df_destino["Classificação"] = (
        df_origem.iloc[:, 0]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    df_destino["Descrição"] = (
        df_origem.iloc[:, 3]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    df_destino["Sint./An."] = (
        df_origem.iloc[:, 1]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.upper()
        .apply(
            lambda valor: "S" if valor == "T" else "A"
        )
    )

    def definir_tipo_conta(classificacao):
        """
        Define o tipo da conta com base no primeiro caractere
        da classificação.
        """
        if (
            not isinstance(classificacao, str)
            or not classificacao
        ):
            return "R"

        primeiro_caractere = classificacao[0]

        if primeiro_caractere == "1":
            return "A"

        if primeiro_caractere == "2":
            return "P"

        return "R"

    df_destino["At/Pas/Res"] = classif.apply(
        definir_tipo_conta
    )

    df_destino["Indice"] = range(
        1,
        len(df_destino) + 1
    )

    return df_destino


# ==============================================================================
# TRANSFORMAÇÃO DO BALANCETE USEALL
# ==============================================================================

def transformar_balancete_useall(caminho_arquivo):
    """
    Transforma o balancete CSV exportado pelo sistema Useall.

    Layout final:
        Coluna A = Atividade
        Coluna B = Conta
        Coluna C = Nome
        Coluna D = Cód. Reduzido
        Coluna E = Saldo Anterior
        Coluna F = Débito
        Coluna G = Crédito
        Coluna H = Movimento
        Coluna I = Saldo Acumulado

    Mapeamento:
        Valor fixo "Geral"        -> Atividade
        Origem B - ClasMasc       -> Conta
        Origem C - NomeConta      -> Nome
        Origem A - CodigoConta    -> Cód. Reduzido
        Origem D - SaldoAnterior  -> Saldo Anterior
        Origem E - Debitos        -> Débito
        Origem F - Creditos       -> Crédito
        Débito menos Crédito      -> Movimento
        Origem G - SaldoFinal     -> Saldo Acumulado
    """
    nome_arquivo = os.path.basename(caminho_arquivo)
    extensao = os.path.splitext(caminho_arquivo)[1].lower()

    if extensao != ".csv":
        raise ValueError(
            f"O arquivo '{nome_arquivo}' não possui extensão CSV."
        )

    df_origem = _ler_csv_useall(caminho_arquivo)

    # Remove espaços e uma possível marca BOM dos cabeçalhos.
    df_origem.columns = [
        str(coluna)
        .replace("\ufeff", "")
        .strip()
        for coluna in df_origem.columns
    ]

    colunas_obrigatorias = [
        "CodigoConta",
        "ClasMasc",
        "NomeConta",
        "SaldoAnterior",
        "Debitos",
        "Creditos",
        "SaldoFinal",
    ]

    colunas_ausentes = [
        coluna
        for coluna in colunas_obrigatorias
        if coluna not in df_origem.columns
    ]

    if colunas_ausentes:
        colunas_encontradas = ", ".join(
            str(coluna)
            for coluna in df_origem.columns
        )

        raise ValueError(
            f"O arquivo '{nome_arquivo}' não possui todas as "
            f"colunas obrigatórias do Useall. "
            f"Colunas ausentes: {', '.join(colunas_ausentes)}. "
            f"Colunas encontradas: {colunas_encontradas}."
        )

    # Converte campos contendo somente espaços em valores ausentes.
    df_origem = df_origem.replace(
        r"^\s*$",
        pd.NA,
        regex=True
    )

    # Remove linhas completamente vazias nas colunas utilizadas.
    df_origem.dropna(
        how="all",
        subset=colunas_obrigatorias,
        inplace=True
    )

    df_origem.reset_index(
        drop=True,
        inplace=True
    )

    # Inicializa o DataFrame com o mesmo índice da origem.
    # Isso garante que o valor "Geral" seja repetido em todas as linhas.
    df_destino = pd.DataFrame(
        index=df_origem.index
    )

    # Coluna A: valor fixo
    df_destino["Atividade"] = "Geral"

    # Coluna B: origem B, ClasMasc
    df_destino["Conta"] = (
        df_origem["ClasMasc"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    # Coluna C: origem C, NomeConta
    df_destino["Nome"] = (
        df_origem["NomeConta"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    # Coluna D: origem A, CodigoConta
    df_destino["Cód. Reduzido"] = (
        df_origem["CodigoConta"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    # Converte os campos monetários antes de montar as colunas.
    saldo_anterior = _converter_serie_decimal_brasileiro(
        df_origem["SaldoAnterior"]
    )

    debito = _converter_serie_decimal_brasileiro(
        df_origem["Debitos"]
    )

    credito = _converter_serie_decimal_brasileiro(
        df_origem["Creditos"]
    )

    saldo_acumulado = _converter_serie_decimal_brasileiro(
        df_origem["SaldoFinal"]
    )

    # Coluna E: origem D, SaldoAnterior
    df_destino["Saldo Anterior"] = saldo_anterior

    # Coluna F: origem E, Debitos
    df_destino["Débito"] = debito

    # Coluna G: origem F, Creditos
    df_destino["Crédito"] = credito

    # Coluna H: Débito menos Crédito
    df_destino["Movimento"] = debito - credito

    # Coluna I: origem G, SaldoFinal
    df_destino["Saldo Acumulado"] = saldo_acumulado

    return df_destino