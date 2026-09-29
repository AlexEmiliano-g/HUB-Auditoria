import os
import re
import unicodedata
from collections import OrderedDict
from datetime import date, datetime

import pandas as pd


EXTENSOES_VALIDAS_cotravale = {".xls", ".xlsx"}

COLUNAS_DESTINO_cotravale = [
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

MESES_cotravale = {
    "JAN": 1, "JANEIRO": 1,
    "FEV": 2, "FEVEREIRO": 2,
    "MAR": 3, "MARCO": 3,
    "ABR": 4, "ABRIL": 4,
    "MAI": 5, "MAIO": 5,
    "JUN": 6, "JUNHO": 6,
    "JUL": 7, "JULHO": 7,
    "AGO": 8, "AGOSTO": 8,
    "SET": 9, "SEPT": 9, "SETEMBRO": 9,
    "OUT": 10, "OUTUBRO": 10,
    "NOV": 11, "NOVEMBRO": 11,
    "DEZ": 12, "DEZEMBRO": 12,
}

TERMOS_RESULTADO_cotravale = {
    "RESULTADO",
    "RESULTADOS",
    "RECEITA",
    "RECEITAS",
    "DESPESA",
    "DESPESAS",
    "CUSTO",
    "CUSTOS",
}


def _valor_vazio_cotravale(valor):
    if valor is None:
        return True
    try:
        return bool(pd.isna(valor))
    except (TypeError, ValueError):
        return False


def _normalizar_texto_cotravale(valor):
    if _valor_vazio_cotravale(valor):
        return ""
    texto = str(valor).replace("\xa0", " ")
    return re.sub(r"\s+", " ", texto).strip()


def _normalizar_texto_comparacao_cotravale(valor):
    texto = _normalizar_texto_cotravale(valor).upper()
    texto = unicodedata.normalize("NFKD", texto)
    return "".join(c for c in texto if not unicodedata.combining(c))


def _normalizar_classificacao_cotravale(valor):
    if _valor_vazio_cotravale(valor):
        return ""

    if isinstance(valor, int):
        return str(valor)

    if isinstance(valor, float):
        if valor.is_integer():
            return str(int(valor))
        return format(valor, "f").rstrip("0").rstrip(".")

    texto = _normalizar_texto_cotravale(valor).replace(" ", "")

    if re.fullmatch(r"\d+\.0+", texto):
        return texto.split(".", maxsplit=1)[0]

    if not re.fullmatch(r"\d+(?:\.\d+)*", texto):
        return ""

    return texto


def _converter_numero_cotravale(valor):
    if _valor_vazio_cotravale(valor):
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


def _ajustar_ano_cotravale(ano):
    ano = int(ano)
    return 2000 + ano if ano < 100 else ano


def _interpretar_mes_cotravale(valor):
    if _valor_vazio_cotravale(valor):
        return None

    if isinstance(valor, (datetime, date, pd.Timestamp)):
        return int(valor.year), int(valor.month)

    texto = _normalizar_texto_comparacao_cotravale(valor)
    texto = texto.replace(".", "/").replace("-", "/")
    texto = re.sub(r"\s+", "", texto)

    correspondencia = re.fullmatch(r"([A-Z]+)/(\d{2}|\d{4})", texto)
    if correspondencia:
        numero_mes = MESES_cotravale.get(correspondencia.group(1))
        if numero_mes:
            return _ajustar_ano_cotravale(correspondencia.group(2)), numero_mes

    correspondencia = re.fullmatch(
        r"(0?[1-9]|1[0-2])/(\d{2}|\d{4})",
        texto,
    )
    if correspondencia:
        return (
            _ajustar_ano_cotravale(correspondencia.group(2)),
            int(correspondencia.group(1)),
        )

    return None


def _obter_engine_excel_cotravale(caminho_arquivo):
    extensao = os.path.splitext(caminho_arquivo)[1].lower()
    return "xlrd" if extensao == ".xls" else "openpyxl"


def _ler_primeira_aba_cotravale(caminho_arquivo):
    nome_arquivo = os.path.basename(caminho_arquivo)
    try:
        dataframe = pd.read_excel(
            caminho_arquivo,
            sheet_name=0,
            header=None,
            dtype=object,
            engine=_obter_engine_excel_cotravale(caminho_arquivo),
        )
    except Exception as erro:
        raise ValueError(
            f"Nao foi possivel ler o arquivo cotravale "
            f"'{nome_arquivo}'. Erro: {erro}"
        ) from erro

    if dataframe.empty:
        raise ValueError(f"O arquivo cotravale '{nome_arquivo}' esta vazio.")

    return dataframe


def _pontuacao_cabecalho_cotravale(linha):
    textos = [
        _normalizar_texto_comparacao_cotravale(valor)
        for valor in linha.tolist()
    ]

    termos = {"GRUPO", "CLASSIFICACAO", "DESCRICAO DA CONTA"}
    quantidade = sum(
        any(termo == texto or termo in texto for texto in textos)
        for termo in termos
    )
    meses = sum(
        _interpretar_mes_cotravale(valor) is not None
        for valor in linha.tolist()
    )

    return quantidade, meses


def _localizar_cabecalho_cotravale(dataframe, nome_arquivo):
    for indice in range(dataframe.shape[0]):
        quantidade_titulos, quantidade_meses = _pontuacao_cabecalho_cotravale(
            dataframe.iloc[indice]
        )
        if quantidade_titulos >= 3 and quantidade_meses >= 2:
            return indice

    raise ValueError(
        f"Nao foi possivel localizar o cabecalho mensal no arquivo "
        f"cotravale '{nome_arquivo}'."
    )


def _localizar_colunas_fixadas_cotravale(linha_cabecalho):
    indice_grupo = None
    indice_classificacao = None
    indice_descricao = None

    for indice, valor in enumerate(linha_cabecalho.tolist()):
        texto = _normalizar_texto_comparacao_cotravale(valor)

        if texto == "GRUPO":
            indice_grupo = indice
        elif texto == "CLASSIFICACAO":
            indice_classificacao = indice
        elif texto in {"DESCRICAO DA CONTA", "DESCRICAO CONTA"}:
            indice_descricao = indice

    if indice_classificacao is None or indice_descricao is None:
        raise ValueError(
            "O cotravale nao encontrou as colunas Classificacao e "
            "Descricao da Conta."
        )

    return indice_grupo, indice_classificacao, indice_descricao


def _localizar_colunas_mensais_cotravale(linha_cabecalho):
    colunas = []

    for indice, valor in enumerate(linha_cabecalho.tolist()):
        periodo = _interpretar_mes_cotravale(valor)
        if periodo is not None:
            ano, mes = periodo
            colunas.append({
                "indice": indice,
                "ano": ano,
                "mes": mes,
                "rotulo": _normalizar_texto_cotravale(valor),
            })

    colunas.sort(key=lambda item: (item["ano"], item["mes"], item["indice"]))

    if len(colunas) < 2:
        raise ValueError(
            "O cotravale precisa de pelo menos duas colunas mensais."
        )

    return colunas


def _conta_eh_resultado_cotravale(
    linha,
    indice_grupo,
    indice_classificacao,
    indice_descricao,
):
    """
    Identifica a natureza da conta exclusivamente pelo primeiro algarismo
    da classificacao contábil.

    Regras:
        classificacao iniciada por 1 = ativo, conta patrimonial;
        classificacao iniciada por 2 = passivo, conta patrimonial;
        classificacao iniciada por 3 ou algarismo superior = resultado.

    As colunas Grupo e Descricao da Conta nao interferem na classificacao
    entre patrimonial e resultado.
    """
    classificacao = _normalizar_classificacao_cotravale(
        linha.iloc[indice_classificacao]
    )

    if not classificacao:
        return False

    primeiro_algarismo = classificacao[0]

    if not primeiro_algarismo.isdigit():
        return False

    return int(primeiro_algarismo) >= 3


def _extrair_registros_periodo_cotravale(
    dataframe_dados,
    indice_grupo,
    indice_classificacao,
    indice_descricao,
    coluna_anterior,
    coluna_atual,
    saldos_resultado_anteriores,
):
    """
    Extrai os registros de um periodo mensal.

    Contas patrimoniais:
        saldo anterior = saldo informado no mes anterior;
        movimento = saldo atual informado - saldo anterior;
        saldo acumulado = saldo atual informado.

    Contas de resultado:
        o valor informado em cada coluna mensal representa o movimento;
        janeiro desconsidera dezembro e inicia com saldo anterior zero;
        de fevereiro em diante, o saldo anterior corresponde ao saldo
        final calculado para a mesma conta no mes anterior;
        saldo acumulado = saldo anterior + movimento do mes.
    """
    registros = []
    periodo_e_janeiro = coluna_atual["mes"] == 1

    for _, linha in dataframe_dados.iterrows():
        classificacao = _normalizar_classificacao_cotravale(
            linha.iloc[indice_classificacao]
        )
        nome = _normalizar_texto_cotravale(
            linha.iloc[indice_descricao]
        )

        if not classificacao or not nome:
            continue

        valor_mes_anterior = _converter_numero_cotravale(
            linha.iloc[coluna_anterior["indice"]]
        )
        valor_mes_atual = _converter_numero_cotravale(
            linha.iloc[coluna_atual["indice"]]
        )

        conta_resultado = _conta_eh_resultado_cotravale(
            linha,
            indice_grupo,
            indice_classificacao,
            indice_descricao,
        )

        if conta_resultado:
            # Para contas de resultado, o valor da coluna mensal e o
            # movimento do proprio mes, e nao um saldo final recebido.
            movimento = round(valor_mes_atual, 2)

            if periodo_e_janeiro:
                # Janeiro nunca utiliza o valor apresentado em dezembro.
                saldo_anterior = 0.0
            else:
                # Nos demais meses, utiliza o saldo final calculado no
                # processamento do mes anterior para a mesma conta.
                saldo_anterior = round(
                    saldos_resultado_anteriores.get(classificacao, 0.0),
                    2,
                )

            saldo_acumulado = round(saldo_anterior + movimento, 2)
            saldos_resultado_anteriores[classificacao] = saldo_acumulado
        else:
            # Contas patrimoniais continuam recebendo saldos finais.
            saldo_anterior = round(valor_mes_anterior, 2)
            saldo_acumulado = round(valor_mes_atual, 2)
            movimento = round(saldo_acumulado - saldo_anterior, 2)

        registros.append({
            "Atividade": "Geral",
            "Conta": classificacao,
            "Nome": nome,
            "Cód. Reduzido": classificacao,
            "Saldo Anterior": saldo_anterior,
            "Débito": 0.0,
            "Crédito": 0.0,
            "Movimento": movimento,
            "Saldo Acumulado": saldo_acumulado,
        })

    return registros


def _montar_dataframe_cotravale(registros):
    if not registros:
        return pd.DataFrame(columns=COLUNAS_DESTINO_cotravale)

    dataframe = pd.DataFrame(registros)

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

    dataframe["Débito"] = 0.0
    dataframe["Crédito"] = 0.0

    return dataframe[COLUNAS_DESTINO_cotravale].copy()


def transformar_balancete_cotravale(caminho_arquivo):
    nome_arquivo = os.path.basename(caminho_arquivo)
    extensao = os.path.splitext(caminho_arquivo)[1].lower()

    if extensao not in EXTENSOES_VALIDAS_cotravale:
        raise ValueError(f"O arquivo '{nome_arquivo}' nao e um Excel valido.")

    origem = _ler_primeira_aba_cotravale(caminho_arquivo)
    indice_cabecalho = _localizar_cabecalho_cotravale(origem, nome_arquivo)
    linha_cabecalho = origem.iloc[indice_cabecalho]

    indice_grupo, indice_classificacao, indice_descricao = (
        _localizar_colunas_fixadas_cotravale(linha_cabecalho)
    )
    colunas_mensais = _localizar_colunas_mensais_cotravale(linha_cabecalho)
    dados = origem.iloc[indice_cabecalho + 1:].copy()
    resultados = OrderedDict()
    saldos_resultado_anteriores = {}

    for posicao in range(1, len(colunas_mensais)):
        anterior = colunas_mensais[posicao - 1]
        atual = colunas_mensais[posicao]

        registros = _extrair_registros_periodo_cotravale(
            dados,
            indice_grupo,
            indice_classificacao,
            indice_descricao,
            anterior,
            atual,
            saldos_resultado_anteriores,
        )

        if registros:
            nome_aba = f"{atual['mes']:02d}"
            resultados[nome_aba] = _montar_dataframe_cotravale(registros)

    if not resultados:
        raise ValueError(
            f"Nenhuma aba mensal foi gerada no arquivo cotravale "
            f"'{nome_arquivo}'."
        )

    return resultados


def _chave_ordenacao_aba_cotravale(nome_aba):
    correspondencia = re.match(r"^(0[1-9]|1[0-2])", str(nome_aba))
    if correspondencia:
        return int(correspondencia.group(1)), str(nome_aba).casefold()
    return 99, str(nome_aba).casefold()


def processar(lista_arquivos):
    if not lista_arquivos:
        raise ValueError(
            "Nenhum arquivo foi selecionado para o sistema cotravale."
        )

    arquivos_excel = [
        arquivo
        for arquivo in lista_arquivos
        if os.path.splitext(str(arquivo))[1].lower()
        in EXTENSOES_VALIDAS_cotravale
    ]

    if not arquivos_excel:
        raise ValueError(
            "Nenhum arquivo Excel valido foi encontrado para o cotravale."
        )

    consolidados = OrderedDict()

    for arquivo in arquivos_excel:
        resultados_arquivo = transformar_balancete_cotravale(arquivo)

        for nome_aba, dataframe in resultados_arquivo.items():
            if nome_aba in consolidados:
                consolidados[nome_aba] = pd.concat(
                    [consolidados[nome_aba], dataframe],
                    ignore_index=True,
                )
            else:
                consolidados[nome_aba] = dataframe.copy()

    resultados = OrderedDict()

    for nome_aba in sorted(consolidados, key=_chave_ordenacao_aba_cotravale):
        dataframe = consolidados[nome_aba]
        dataframe.reset_index(drop=True, inplace=True)
        resultados[nome_aba] = dataframe

    if not resultados:
        raise ValueError(
            "Nenhum resultado foi gerado para o sistema cotravale."
        )

    return resultados
