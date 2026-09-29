import os
import re

import pandas as pd


# ==============================================================================
# CONFIGURACOES DO SISTEMA serramar
# ==============================================================================

EXTENSAO_VALIDA_serramar = ".txt"

COLUNAS_DESTINO_serramar = [
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


# ==============================================================================
# NOMES DAS ABAS
# ==============================================================================

def _limpar_nome_aba_serramar(nome):
    """Ajusta um texto para utilizacao como nome de aba do Excel."""
    nome_limpo = re.sub(
        r'[\\/\x2a?:\[\]]',
        "_",
        str(nome).strip(),
    )
    return (nome_limpo or "Sem nome")[:31]


def _obter_nome_sem_extensao_serramar(caminho_arquivo):
    """Retorna o nome do arquivo sem a extensao."""
    nome_arquivo = os.path.basename(caminho_arquivo)
    return os.path.splitext(nome_arquivo)[0].strip()


def _obter_mes_serramar(caminho_arquivo):
    """Extrai XX quando o nome do arquivo comeca com B_XX."""
    nome_arquivo = os.path.basename(caminho_arquivo)
    correspondencia = re.match(
        r"^B_(0[1-9]|1[0-2])",
        nome_arquivo,
        re.IGNORECASE,
    )
    return correspondencia.group(1) if correspondencia else None


def _registrar_nome_aba_serramar(nome, nomes_utilizados):
    """Registra um nome de aba caso ainda nao esteja em uso."""
    nome_aba = _limpar_nome_aba_serramar(nome)
    chave = nome_aba.casefold()

    if chave in nomes_utilizados:
        return None

    nomes_utilizados.add(chave)
    return nome_aba


def _gerar_nome_aba_serramar(caminho_arquivo, nomes_utilizados):
    """Gera um nome exclusivo, priorizando XX para arquivos B_XX."""
    nome_completo = _obter_nome_sem_extensao_serramar(caminho_arquivo)
    mes = _obter_mes_serramar(caminho_arquivo)

    if mes:
        nome_aba = _registrar_nome_aba_serramar(mes, nomes_utilizados)
        if nome_aba is not None:
            return nome_aba

    nome_aba = _registrar_nome_aba_serramar(
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
        nome_aba = _registrar_nome_aba_serramar(
            candidato,
            nomes_utilizados,
        )

        if nome_aba is not None:
            return nome_aba

        contador += 1


# ==============================================================================
# NORMALIZACAO DOS TEXTOS E CODIGOS
# ==============================================================================

def _normalizar_texto_serramar(valor):
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


def _normalizar_classificacao_serramar(valor):
    """Normaliza a classificacao encontrada antes do primeiro hifen."""
    texto = _normalizar_texto_serramar(valor).replace(" ", "")

    if not texto:
        return ""

    if not re.fullmatch(r"\d+(?:[.]\d+){0,}", texto):
        return ""

    return texto


def _remover_zeros_conta_sintetica_serramar(classificacao):
    """
    Remove zeros finais somente de contas sinteticas.

    Os zeros internos sao preservados. Contas analiticas nao passam por
    esta funcao e, portanto, mantem inclusive o zero existente no final.
    """
    classificacao = _normalizar_classificacao_serramar(classificacao)

    if not classificacao or not classificacao.isdigit():
        return classificacao

    classificacao_sem_zeros = classificacao.rstrip("0")
    return classificacao_sem_zeros or classificacao


# ==============================================================================
# CONVERSAO DOS VALORES E TRATAMENTO D/C
# ==============================================================================

def _extrair_natureza_serramar(valor):
    """Extrai D ou C existente no final do saldo."""
    texto = _normalizar_texto_serramar(valor).upper()
    correspondencia = re.search(
        r"([DC])\s{0,}$",
        texto,
        re.IGNORECASE,
    )
    return correspondencia.group(1).upper() if correspondencia else ""


def _converter_numero_serramar(valor):
    """Converte valor monetario brasileiro e remove D/C do final."""
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


def _aplicar_natureza_serramar(valor):
    """Aplica D como positivo e C como negativo."""
    numero = _converter_numero_serramar(valor)
    natureza = _extrair_natureza_serramar(valor)

    if natureza == "D":
        return round(abs(numero), 2)
    if natureza == "C":
        return round(-abs(numero), 2)

    return round(numero, 2)


# ==============================================================================
# LEITURA DO TXT
# ==============================================================================

def _ler_linhas_txt_serramar(caminho_arquivo):
    """Le o TXT testando as codificacoes mais comuns."""
    nome_arquivo = os.path.basename(caminho_arquivo)
    codificacoes = ["utf-8-sig", "cp1252", "latin-1"]
    ultimo_erro = None

    for codificacao in codificacoes:
        try:
            with open(
                caminho_arquivo,
                mode="r",
                encoding=codificacao,
            ) as arquivo:
                return arquivo.readlines()
        except UnicodeDecodeError as erro:
            ultimo_erro = erro
        except OSError as erro:
            raise ValueError(
                f"Nao foi possivel abrir o arquivo TXT "
                f"'{nome_arquivo}'. Erro: {erro}"
            ) from erro

    raise ValueError(
        f"Nao foi possivel identificar a codificacao do arquivo "
        f"TXT '{nome_arquivo}'. Erro: {ultimo_erro}"
    )


# ==============================================================================
# IDENTIFICACAO DO CABECALHO E DOS CAMPOS
# ==============================================================================

def _linha_eh_cabecalho_serramar(linha):
    """Identifica cabecalhos, separadores e linhas vazias."""
    texto = _normalizar_texto_serramar(linha).upper()

    if not texto:
        return True

    if re.fullmatch(r"[-_= ]+", texto):
        return True

    termos = [
        "TITULO",
        "SALDO INICIAL",
        "DEBITO",
        "CREDITO",
        "SALDO FINAL",
    ]

    return sum(termo in texto for termo in termos) >= 2


def _obter_correspondencias_monetarias_serramar(linha):
    """Localiza valores monetarios brasileiros, com D/C opcional."""
    padrao = re.compile(
        r"(?<![\d.,])"
        r"[+-]?"
        r"(?:\d{1,3}(?:\.\d{3}){0,}|\d+)"
        r",\d{2}"
        r"[DCdc]?"
        r"-?"
        r"(?![\d.,])"
    )

    return list(padrao.finditer(str(linha)))


def _separar_classificacao_nome_serramar(prefixo):
    """Separa classificacao e nome e preserva a indentacao da linha."""
    prefixo_original = str(prefixo).replace("\xa0", " ").rstrip()
    prefixo_sem_tab = prefixo_original.expandtabs(2)
    quantidade_espacos_inicio = len(prefixo_sem_tab) - len(
        prefixo_sem_tab.lstrip(" ")
    )

    correspondencia = re.fullmatch(
        r"\s{0,}"
        r"(?P<classificacao>\d+(?:[.]\d+){0,})"
        r"\s{0,}-\s{0,}"
        r"(?P<nome>.+?)"
        r"\s{0,}",
        prefixo_sem_tab,
    )

    if correspondencia is None:
        return "", "", 0

    classificacao = _normalizar_classificacao_serramar(
        correspondencia.group("classificacao")
    )
    nome = _normalizar_texto_serramar(
        correspondencia.group("nome")
    )

    return classificacao, nome, quantidade_espacos_inicio


def _extrair_registro_serramar(linha):
    """Extrai classificacao, nome e os quatro valores monetarios."""
    if linha is None:
        return None

    linha_limpa = (
        str(linha)
        .replace("\xa0", " ")
        .replace("\r", "")
        .replace("\n", "")
        .rstrip()
    )

    if _linha_eh_cabecalho_serramar(linha_limpa):
        return None

    valores_encontrados = _obter_correspondencias_monetarias_serramar(
        linha_limpa
    )

    if len(valores_encontrados) < 4:
        return None

    valores = valores_encontrados[-4:]
    prefixo = linha_limpa[:valores[0].start()].rstrip()
    classificacao, nome, indentacao = _separar_classificacao_nome_serramar(
        prefixo
    )

    if not classificacao or not nome:
        return None

    saldo_anterior = _aplicar_natureza_serramar(valores[0].group())
    debito = round(abs(_converter_numero_serramar(valores[1].group())), 2)
    credito = round(-abs(_converter_numero_serramar(valores[2].group())), 2)
    saldo_acumulado = _aplicar_natureza_serramar(valores[3].group())
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
        "_Classificação Original": classificacao,
        "_Indentação": indentacao,
    }


def _ajustar_classificacoes_serramar(registros):
    """
    Remove zeros finais apenas das contas sinteticas.

    Uma conta e sintetica quando a proxima conta valida possui indentacao
    maior. Contas analiticas preservam integralmente a classificacao,
    inclusive quando terminam em zero, como 1121110.
    """
    if not registros:
        return registros

    for indice, registro in enumerate(registros):
        classificacao_original = registro.get(
            "_Classificação Original",
            registro.get("Conta", ""),
        )
        indentacao_atual = int(registro.get("_Indentação", 0))

        proxima_indentacao = -1
        if indice + 1 < len(registros):
            proxima_indentacao = int(
                registros[indice + 1].get("_Indentação", 0)
            )

        conta_sintetica = proxima_indentacao > indentacao_atual

        if conta_sintetica:
            classificacao_final = _remover_zeros_conta_sintetica_serramar(
                classificacao_original
            )
        else:
            classificacao_final = classificacao_original

        registro["Conta"] = classificacao_final
        registro["Cód. Reduzido"] = classificacao_final

    return registros


# ==============================================================================
# MONTAGEM E TRANSFORMACAO
# ==============================================================================

def _montar_dataframe_serramar(registros):
    """Monta o DataFrame final na ordem padrao do tabulador."""
    if not registros:
        return pd.DataFrame(columns=COLUNAS_DESTINO_serramar)

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

    colunas_auxiliares = [
        coluna
        for coluna in ["_Classificação Original", "_Indentação"]
        if coluna in dataframe.columns
    ]
    if colunas_auxiliares:
        dataframe.drop(columns=colunas_auxiliares, inplace=True)

    return dataframe[COLUNAS_DESTINO_serramar].copy()


def transformar_balancete_serramar(caminho_arquivo):
    """Transforma um arquivo TXT do sistema serramar."""
    nome_arquivo = os.path.basename(caminho_arquivo)
    extensao = os.path.splitext(caminho_arquivo)[1].lower()

    if extensao != EXTENSAO_VALIDA_serramar:
        raise ValueError(
            f"O arquivo '{nome_arquivo}' nao possui extensao TXT."
        )

    linhas = _ler_linhas_txt_serramar(caminho_arquivo)

    if not linhas:
        raise ValueError(f"O arquivo TXT '{nome_arquivo}' esta vazio.")

    registros = []

    for linha in linhas:
        registro = _extrair_registro_serramar(linha)
        if registro is not None:
            registros.append(registro)

    if not registros:
        raise ValueError(
            f"Nenhuma conta valida foi encontrada no arquivo serramar "
            f"'{nome_arquivo}'."
        )

    registros = _ajustar_classificacoes_serramar(registros)
    return _montar_dataframe_serramar(registros)


def _chave_ordenacao_arquivo_serramar(caminho_arquivo):
    """Ordena arquivos B_XX pelo mes e os demais pelo nome."""
    mes = _obter_mes_serramar(caminho_arquivo)
    numero_mes = int(mes) if mes else 99
    return numero_mes, os.path.basename(caminho_arquivo).casefold()


def processar(lista_arquivos):
    """Processa todos os arquivos selecionados para o serramar."""
    if not lista_arquivos:
        raise ValueError(
            "Nenhum arquivo foi selecionado para o sistema serramar."
        )

    arquivos_txt = [
        arquivo
        for arquivo in lista_arquivos
        if os.path.splitext(str(arquivo))[1].lower()
        == EXTENSAO_VALIDA_serramar
    ]

    if not arquivos_txt:
        raise ValueError(
            "Nenhum arquivo TXT valido foi encontrado para o serramar."
        )

    resultados = {}
    nomes_utilizados = set()

    for arquivo in sorted(
        arquivos_txt,
        key=_chave_ordenacao_arquivo_serramar,
    ):
        nome_aba = _gerar_nome_aba_serramar(
            arquivo,
            nomes_utilizados,
        )
        dataframe = transformar_balancete_serramar(arquivo)

        if dataframe is None or dataframe.empty:
            raise ValueError(
                f"O arquivo '{os.path.basename(arquivo)}' nao retornou "
                "dados validos."
            )

        resultados[nome_aba] = dataframe

    if not resultados:
        raise ValueError(
            "Nenhum resultado foi gerado para o sistema serramar."
        )

    return resultados
