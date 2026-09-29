import os
import re
import unicodedata
from collections import OrderedDict

import pandas as pd


# ==============================================================================
# CONFIGURACOES DO SISTEMA ouro_do_sul
# ==============================================================================

EXTENSOES_VALIDAS_ouro_do_sul = {".xls", ".xlsx"}

COLUNAS_DESTINO_ouro_do_sul = [
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

# As colunas A ate E podem conter partes da classificacao.
# Em alguns arquivos, uma dessas colunas pode conter a nomenclatura.
QUANTIDADE_COLUNAS_CLASSIFICACAO_ouro_do_sul = 5

MESES_ouro_do_sul = {
    "JAN": "01",
    "JANEIRO": "01",
    "FEV": "02",
    "FEVEREIRO": "02",
    "MAR": "03",
    "MARCO": "03",
    "ABR": "04",
    "ABRIL": "04",
    "MAI": "05",
    "MAIO": "05",
    "JUN": "06",
    "JUNHO": "06",
    "JUL": "07",
    "JULHO": "07",
    "AGO": "08",
    "AGOSTO": "08",
    "SET": "09",
    "SETEMBRO": "09",
    "OUT": "10",
    "OUTUBRO": "10",
    "NOV": "11",
    "NOVEMBRO": "11",
    "DEZ": "12",
    "DEZEMBRO": "12",
}


# ==============================================================================
# NORMALIZACAO GERAL
# ==============================================================================

def _valor_vazio_ouro_do_sul(valor):
    """Verifica se um valor esta vazio ou ausente."""
    if valor is None:
        return True

    try:
        return bool(pd.isna(valor))
    except (TypeError, ValueError):
        return False


def _normalizar_texto_ouro_do_sul(valor):
    """Normaliza espacos e converte valores ausentes em texto vazio."""
    if _valor_vazio_ouro_do_sul(valor):
        return ""

    texto = str(valor).replace("\xa0", " ")
    return re.sub(r"\s{1,}", " ", texto).strip()


def _normalizar_texto_comparacao_ouro_do_sul(valor):
    """Converte para maiusculas e remove acentos para comparacao."""
    texto = _normalizar_texto_ouro_do_sul(valor).upper()
    if not texto:
        return ""

    texto = unicodedata.normalize("NFKD", texto)
    return "".join(
        caractere
        for caractere in texto
        if not unicodedata.combining(caractere)
    )


def _texto_eh_parte_classificacao_ouro_do_sul(valor):
    """
    Verifica se uma celula de A ate E faz parte da classificacao.

    Sao aceitos somente grupos numericos, com zeros preservados.
    Exemplos validos: 1, 01, 001, 1.0.
    """
    texto = _normalizar_texto_ouro_do_sul(valor)
    if not texto:
        return False

    if isinstance(valor, int):
        return True

    if isinstance(valor, float):
        return valor.is_integer()

    texto_sem_espacos = texto.replace(" ", "")

    if re.fullmatch(r"\d+", texto_sem_espacos):
        return True

    return bool(re.fullmatch(r"\d+\.0+", texto_sem_espacos))


def _normalizar_parte_classificacao_ouro_do_sul(valor):
    """Converte uma parte numerica da classificacao em texto."""
    if not _texto_eh_parte_classificacao_ouro_do_sul(valor):
        return ""

    if isinstance(valor, int):
        return str(valor)

    if isinstance(valor, float):
        return str(int(valor))

    texto = _normalizar_texto_ouro_do_sul(valor).replace(" ", "")

    if re.fullmatch(r"\d+\.0+", texto):
        return texto.split(".", maxsplit=1)[0]

    return texto


def _converter_numero_ouro_do_sul(valor):
    """Converte valores monetarios brasileiros para float."""
    if _valor_vazio_ouro_do_sul(valor):
        return 0.0

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


def _valor_parece_monetario_ouro_do_sul(valor):
    """Verifica se a celula aparenta conter um valor monetario."""
    if _valor_vazio_ouro_do_sul(valor):
        return False

    if isinstance(valor, (int, float)):
        return True

    texto = _normalizar_texto_ouro_do_sul(valor)

    return bool(
        re.fullmatch(
            r"[+-]?(?:\d{1,3}(?:\.\d{3}){0,}|\d+),\d{2}-?",
            texto.replace(" ", ""),
        )
    )


# ==============================================================================
# IDENTIFICACAO DO MES PELO NOME DA ABA
# ==============================================================================

def _limpar_nome_aba_ouro_do_sul(nome):
    """Cria um nome valido para a aba de destino."""
    nome_limpo = re.sub(
        r'[\\/\x2a?:\[\]]',
        "_",
        str(nome).strip(),
    )
    return (nome_limpo or "Sem nome")[:31]


def _obter_mes_nome_aba_ouro_do_sul(nome_aba):
    """
    Identifica o mes pelo nome da aba da origem.

    Prioridades:
        B_XX
        numero isolado entre 01 e 12
        nome ou abreviacao do mes
    """
    texto = _normalizar_texto_comparacao_ouro_do_sul(nome_aba)

    correspondencia = re.search(
        r"(?:^|[^A-Z0-9])B[_\- ]?(0[1-9]|1[0-2])(?:[^0-9]|$)",
        texto,
        re.IGNORECASE,
    )
    if correspondencia:
        return correspondencia.group(1)

    correspondencia = re.search(
        r"(?:^|[^0-9])(0[1-9]|1[0-2])(?:[^0-9]|$)",
        texto,
    )
    if correspondencia:
        return correspondencia.group(1)

    palavras = re.findall(r"[A-Z]+", texto)
    for palavra in palavras:
        if palavra in MESES_ouro_do_sul:
            return MESES_ouro_do_sul[palavra]

    return None


def _gerar_nome_destino_ouro_do_sul(nome_aba, nomes_destino):
    """Retorna o mes identificado ou um nome de aba exclusivo."""
    mes = _obter_mes_nome_aba_ouro_do_sul(nome_aba)
    if mes:
        return mes

    nome_base = _limpar_nome_aba_ouro_do_sul(nome_aba)
    candidato = nome_base
    contador = 2

    while candidato.casefold() in nomes_destino:
        sufixo = "_" + str(contador)
        candidato = nome_base[:31 - len(sufixo)] + sufixo
        contador += 1

    return candidato


# ==============================================================================
# LEITURA DO EXCEL
# ==============================================================================

def _obter_engine_excel_ouro_do_sul(caminho_arquivo):
    """Define o mecanismo apropriado para XLS ou XLSX."""
    extensao = os.path.splitext(caminho_arquivo)[1].lower()
    return "xlrd" if extensao == ".xls" else "openpyxl"


def _ler_abas_excel_ouro_do_sul(caminho_arquivo):
    """Le todas as abas da pasta de trabalho sem cabecalho fixo."""
    nome_arquivo = os.path.basename(caminho_arquivo)
    engine = _obter_engine_excel_ouro_do_sul(caminho_arquivo)

    try:
        pasta_trabalho = pd.ExcelFile(caminho_arquivo, engine=engine)
    except Exception as erro:
        raise ValueError(
            f"Nao foi possivel abrir o arquivo ouro_do_sul "
            f"'{nome_arquivo}'. Erro: {erro}"
        ) from erro

    abas = OrderedDict()

    for nome_aba in pasta_trabalho.sheet_names:
        try:
            dataframe = pd.read_excel(
                pasta_trabalho,
                sheet_name=nome_aba,
                header=None,
                dtype=object,
            )
        except Exception as erro:
            raise ValueError(
                f"Nao foi possivel ler a aba '{nome_aba}' do arquivo "
                f"'{nome_arquivo}'. Erro: {erro}"
            ) from erro

        if not dataframe.empty:
            abas[nome_aba] = dataframe

    if not abas:
        raise ValueError(
            f"O arquivo ouro_do_sul '{nome_arquivo}' nao possui abas com dados."
        )

    return abas


# ==============================================================================
# IDENTIFICACAO DE CABECALHOS E EXTRACAO DAS LINHAS
# ==============================================================================

def _linha_eh_cabecalho_ouro_do_sul(linha):
    """Identifica linhas de cabecalho, mes e titulos do balancete."""
    valores = [
        _normalizar_texto_comparacao_ouro_do_sul(valor)
        for valor in linha.tolist()
    ]
    texto_completo = " | ".join(valor for valor in valores if valor)

    if not texto_completo:
        return True

    if texto_completo.startswith("MES:"):
        return True

    termos = [
        "CONTA",
        "N O M E N C L A T U R A",
        "NOMENCLATURA",
        "SALDO ANTERIOR",
        "DEBITO",
        "CREDITO",
        "SALDO ATUAL",
    ]

    return sum(termo in texto_completo for termo in termos) >= 2


def _obter_classificacao_nome_ouro_do_sul(linha):
    """
    Concatena as partes numéricas das colunas A até E e identifica
    a nomenclatura da conta.

    Regras:
        - colunas A e B preservam a quantidade original de algarismos;
        - a partir da coluna C, cada grupo numérico deverá possuir
          no mínimo dois algarismos;
        - quando o grupo tiver apenas um algarismo, será acrescentado
          zero à esquerda;
        - valores que já possuam dois ou mais algarismos serão
          preservados;
        - a primeira célula textual será considerada a nomenclatura;
        - quando a coluna E for numérica, o sistema procurará a
          nomenclatura nas colunas seguintes.

    Exemplos:
        A=1
            Classificação: 1

        A=1, B=1
            Classificação: 11

        A=1, B=1, C=1
            Classificação: 1101

        A=1, B=1, C=1, D=01
            Classificação: 110101

        A=1, B=1, C=1, D=01, E=1
            Classificação: 11010101
    """
    partes_classificacao = []
    nome = ""

    limite = min(
        QUANTIDADE_COLUNAS_CLASSIFICACAO_ouro_do_sul,
        len(linha)
    )

    for indice in range(
        limite
    ):
        valor = linha.iloc[
            indice
        ]

        if _valor_vazio_ouro_do_sul(
            valor
        ):
            continue

        if _texto_eh_parte_classificacao_ouro_do_sul(
            valor
        ):
            parte = _normalizar_parte_classificacao_ouro_do_sul(
                valor
            )

            # A = índice 0
            # B = índice 1
            # C = índice 2
            # D = índice 3
            # E = índice 4
            #
            # A partir da coluna C, cada parte numérica deverá
            # possuir no mínimo dois algarismos.
            if indice >= 2:
                parte = parte.zfill(
                    2
                )

            partes_classificacao.append(
                parte
            )

        elif not nome:
            nome = _normalizar_texto_ouro_do_sul(
                valor
            )

    # Se não foi encontrado texto entre as colunas A e E,
    # significa que a coluna E também pode fazer parte da
    # classificação. Nesse caso, procura a nomenclatura nas
    # colunas posteriores.
    if not nome:
        for indice in range(
            QUANTIDADE_COLUNAS_CLASSIFICACAO_ouro_do_sul,
            len(linha)
        ):
            valor = linha.iloc[
                indice
            ]

            if _valor_vazio_ouro_do_sul(
                valor
            ):
                continue

            # Ao encontrar o primeiro valor monetário, encerra
            # a procura, pois a nomenclatura deveria estar antes.
            if _valor_parece_monetario_ouro_do_sul(
                valor
            ):
                break

            texto = _normalizar_texto_ouro_do_sul(
                valor
            )

            if texto:
                nome = texto
                break

    classificacao = "".join(
        parte
        for parte in partes_classificacao
        if parte
    )

    return classificacao, nome


def _obter_valores_monetarios_ouro_do_sul(linha):
    """Localiza, na ordem, saldo anterior, debito, credito e saldo atual."""
    valores = []

    for valor in linha.tolist():
        if _valor_parece_monetario_ouro_do_sul(valor):
            valores.append(_converter_numero_ouro_do_sul(valor))

    if len(valores) < 4:
        return None

    return valores[-4:]


def _extrair_registro_ouro_do_sul(linha):
    """Extrai classificacao, nome e quatro valores monetarios da linha."""
    if _linha_eh_cabecalho_ouro_do_sul(linha):
        return None

    classificacao, nome = _obter_classificacao_nome_ouro_do_sul(linha)

    if not classificacao or not nome:
        return None

    valores = _obter_valores_monetarios_ouro_do_sul(linha)
    if valores is None:
        return None

    saldo_anterior, debito, credito, saldo_acumulado = valores
    movimento = round(debito + credito, 2)

    return {
        "Conta": classificacao,
        "Nome": nome,
        "Cód. Reduzido": classificacao,
        "Saldo Anterior": round(saldo_anterior, 2),
        "Débito": round(debito, 2),
        "Crédito": round(credito, 2),
        "Movimento": movimento,
        "Saldo Acumulado": round(saldo_acumulado, 2),
    }


def _transformar_aba_ouro_do_sul(dataframe, nome_arquivo, nome_aba):
    """Transforma uma aba individual do arquivo Excel."""
    registros = []

    for _, linha in dataframe.iterrows():
        registro = _extrair_registro_ouro_do_sul(linha)
        if registro is not None:
            registros.append(registro)

    if not registros:
        print(
            f"[ouro_do_sul] A aba '{nome_aba}' do arquivo "
            f"'{nome_arquivo}' nao possui contas validas."
        )
        return pd.DataFrame(columns=COLUNAS_DESTINO_ouro_do_sul)

    destino = pd.DataFrame(registros)
    destino.insert(0, "Atividade", "Geral")

    for coluna in [
        "Saldo Anterior",
        "Débito",
        "Crédito",
        "Movimento",
        "Saldo Acumulado",
    ]:
        destino[coluna] = pd.to_numeric(
            destino[coluna],
            errors="coerce",
        ).fillna(0.0).round(2)

    destino["Movimento"] = (
        destino["Débito"] + destino["Crédito"]
    ).round(2)

    return destino[COLUNAS_DESTINO_ouro_do_sul].copy()


# ==============================================================================
# CONSOLIDACAO DE ARQUIVOS E ABAS
# ==============================================================================

def transformar_arquivo_ouro_do_sul(caminho_arquivo):
    """Transforma todas as abas de um arquivo e retorna por mes/aba."""
    nome_arquivo = os.path.basename(caminho_arquivo)
    extensao = os.path.splitext(caminho_arquivo)[1].lower()

    if extensao not in EXTENSOES_VALIDAS_ouro_do_sul:
        raise ValueError(
            f"O arquivo '{nome_arquivo}' nao e um Excel valido."
        )

    abas_origem = _ler_abas_excel_ouro_do_sul(caminho_arquivo)
    resultados = OrderedDict()
    nomes_destino = set()

    for nome_aba, dataframe in abas_origem.items():
        nome_destino = _gerar_nome_destino_ouro_do_sul(
            nome_aba,
            nomes_destino,
        )
        nomes_destino.add(nome_destino.casefold())

        dataframe_destino = _transformar_aba_ouro_do_sul(
            dataframe,
            nome_arquivo,
            nome_aba,
        )

        if dataframe_destino.empty:
            continue

        if nome_destino in resultados:
            resultados[nome_destino] = pd.concat(
                [resultados[nome_destino], dataframe_destino],
                ignore_index=True,
            )
        else:
            resultados[nome_destino] = dataframe_destino

    return resultados


def _chave_ordenacao_aba_ouro_do_sul(nome_aba):
    """Ordena primeiro as abas mensais de 01 a 12."""
    if re.fullmatch(r"0[1-9]|1[0-2]", str(nome_aba)):
        return int(nome_aba), str(nome_aba).casefold()

    return 99, str(nome_aba).casefold()


def processar(lista_arquivos):
    """
    Processa varios arquivos Excel e consolida abas do mesmo mes.

    Dois arquivos com uma aba B_01 geram uma unica aba 01 na saida.
    O mes e sempre identificado pelo nome da aba da origem.
    """
    if not lista_arquivos:
        raise ValueError(
            "Nenhum arquivo foi selecionado para o sistema ouro_do_sul."
        )

    arquivos_excel = [
        arquivo
        for arquivo in lista_arquivos
        if os.path.splitext(str(arquivo))[1].lower()
        in EXTENSOES_VALIDAS_ouro_do_sul
    ]

    if not arquivos_excel:
        raise ValueError(
            "Nenhum arquivo Excel valido foi encontrado para o ouro_do_sul."
        )

    consolidados = OrderedDict()

    for arquivo in arquivos_excel:
        resultados_arquivo = transformar_arquivo_ouro_do_sul(arquivo)

        for nome_aba, dataframe in resultados_arquivo.items():
            if nome_aba in consolidados:
                consolidados[nome_aba] = pd.concat(
                    [consolidados[nome_aba], dataframe],
                    ignore_index=True,
                )
            else:
                consolidados[nome_aba] = dataframe.copy()

    if not consolidados:
        raise ValueError(
            "Nenhum resultado foi gerado para o sistema ouro_do_sul."
        )

    resultados_ordenados = OrderedDict()

    for nome_aba in sorted(
        consolidados.keys(),
        key=_chave_ordenacao_aba_ouro_do_sul,
    ):
        dataframe = consolidados[nome_aba]
        dataframe.reset_index(drop=True, inplace=True)
        resultados_ordenados[nome_aba] = dataframe

    return resultados_ordenados
