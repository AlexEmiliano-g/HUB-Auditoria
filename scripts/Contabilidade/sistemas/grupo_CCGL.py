import os
import re
import unicodedata

import pandas as pd


# ==============================================================================
# CONFIGURAÇÕES DO Grupo_CCGL
# ==============================================================================

EXTENSOES_VALIDAS_grupo_CCGL = {
    ".xls",
    ".xlsx",
}

COLUNAS_DESTINO_grupo_CCGL = [
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

PREFIXOS_CONTAS_RESULTADO_grupo_CCGL = (
    "3",
    "4",
    "5",
)

CONTAS_ADICIONAIS_RESULTADO_grupo_CCGL = {
    "240205",
    "240205001",
}


# ==============================================================================
# NOMES DAS ABAS
# ==============================================================================

def _limpar_nome_aba_grupo_CCGL(nome):
    """
    Ajusta um texto para utilização como nome de aba do Excel.

    Caracteres inválidos são substituídos por sublinhado.
    O nome final é limitado a 31 caracteres.
    """
    nome_limpo = re.sub(
        r'[\\/\x2a?:\[\]]',
        "_",
        str(nome).strip()
    )

    if not nome_limpo:
        nome_limpo = "Sem nome"

    return nome_limpo[:31]


def _obter_nome_sem_extensao_grupo_CCGL(caminho_arquivo):
    """
    Retorna o nome do arquivo sem a extensão.
    """
    nome_arquivo = os.path.basename(
        caminho_arquivo
    )

    return os.path.splitext(
        nome_arquivo
    )[0].strip()


def _obter_mes_grupo_CCGL(caminho_arquivo):
    """
    Extrai o mês quando o nome do arquivo começa com B_XX.

    Exemplos:
        B_01_BALANCETE.xlsx -> 01
        B_05.2026.xls       -> 05
        BALANCETE.xlsx      -> None
    """
    nome_arquivo = os.path.basename(
        caminho_arquivo
    )

    correspondencia = re.match(
        r"^B_(0[1-9]|1[0-2])",
        nome_arquivo,
        re.IGNORECASE
    )

    if correspondencia is None:
        return None

    return correspondencia.group(1)


def _registrar_nome_aba_grupo_CCGL(
    nome,
    nomes_utilizados
):
    """
    Registra um nome de aba quando ainda não estiver em uso.
    """
    nome_aba = _limpar_nome_aba_grupo_CCGL(
        nome
    )

    chave_nome = nome_aba.casefold()

    if chave_nome in nomes_utilizados:
        return None

    nomes_utilizados.add(
        chave_nome
    )

    return nome_aba


def _gerar_nome_aba_grupo_CCGL(
    caminho_arquivo,
    nomes_utilizados
):
    """
    Gera um nome exclusivo para a aba.

    Regras:
    1. Arquivo B_XX tenta utilizar XX;
    2. Se XX já estiver em uso, utiliza o nome completo;
    3. Arquivo fora do padrão utiliza o nome sem extensão;
    4. Nomes repetidos recebem um sufixo numérico.
    """
    nome_completo = _obter_nome_sem_extensao_grupo_CCGL(
        caminho_arquivo
    )

    mes = _obter_mes_grupo_CCGL(
        caminho_arquivo
    )

    if mes:
        nome_aba = _registrar_nome_aba_grupo_CCGL(
            mes,
            nomes_utilizados
        )

        if nome_aba is not None:
            return nome_aba

    nome_aba = _registrar_nome_aba_grupo_CCGL(
        nome_completo,
        nomes_utilizados
    )

    if nome_aba is not None:
        return nome_aba

    contador = 2

    while True:
        sufixo = "_" + str(contador)
        limite_nome = 31 - len(sufixo)

        nome_candidato = (
            nome_completo[:limite_nome]
            + sufixo
        )

        nome_aba = _registrar_nome_aba_grupo_CCGL(
            nome_candidato,
            nomes_utilizados
        )

        if nome_aba is not None:
            return nome_aba

        contador += 1


# ==============================================================================
# NORMALIZAÇÃO DOS TEXTOS
# ==============================================================================

def _normalizar_texto_grupo_CCGL(valor):
    """
    Normaliza um texto extraído do Excel.

    A função:
    - converte valores ausentes em texto vazio;
    - substitui espaços especiais;
    - reduz espaços consecutivos;
    - remove espaços das extremidades.
    """
    if valor is None:
        return ""

    try:
        if pd.isna(valor):
            return ""
    except (TypeError, ValueError):
        pass

    texto = str(valor).replace(
        "\xa0",
        " "
    )

    texto = re.sub(
        r"\s{1,}",
        " ",
        texto
    )

    return texto.strip()


def _normalizar_texto_comparacao_grupo_CCGL(valor):
    """
    Normaliza um texto para comparação.

    A função:
    - converte para letras maiúsculas;
    - remove acentos;
    - normaliza espaços.
    """
    texto = _normalizar_texto_grupo_CCGL(
        valor
    ).upper()

    if not texto:
        return ""

    texto_decomposto = unicodedata.normalize(
        "NFKD",
        texto
    )

    texto_sem_acentos = "".join(
        caractere
        for caractere in texto_decomposto
        if not unicodedata.combining(
            caractere
        )
    )

    return texto_sem_acentos


# ==============================================================================
# NORMALIZAÇÃO DA CLASSIFICAÇÃO
# ==============================================================================

def _normalizar_classificacao_grupo_CCGL(valor):
    """
    Remove os pontos existentes na classificação.

    Exemplos:
        1          -> 1
        1.1        -> 11
        1.01       -> 101
        1.1.01     -> 1101
        1.01.02.03 -> 1010203
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
            texto = str(
                int(valor)
            )
        else:
            texto = format(
                valor,
                "f"
            ).rstrip("0").rstrip(".")

    else:
        texto = _normalizar_texto_grupo_CCGL(
            valor
        )

    texto = (
        texto
        .replace("\xa0", "")
        .replace(" ", "")
        .replace(",", ".")
    )

    if re.fullmatch(
        r"\d+\.0+",
        texto
    ):
        texto = texto.split(
            ".",
            maxsplit=1
        )[0]

    classificacao = texto.replace(
        ".",
        ""
    )

    if not classificacao.isdigit():
        return ""

    return classificacao


def _classificacao_valida_grupo_CCGL(valor):
    """
    Verifica se a classificação é válida.
    """
    classificacao = _normalizar_classificacao_grupo_CCGL(
        valor
    )

    return bool(
        classificacao
        and classificacao.isdigit()
    )


# ==============================================================================
# CONVERSÃO DOS VALORES
# ==============================================================================

def _extrair_natureza_grupo_CCGL(valor):
    """
    Extrai a natureza contábil existente ao final do valor.

    Exemplos:
        1.000,00 D -> D
        1.000,00D  -> D
        1.000,00 C -> C
        1.000,00C  -> C
        1.000,00   -> vazio
    """
    if valor is None:
        return ""

    try:
        if pd.isna(valor):
            return ""
    except (TypeError, ValueError):
        pass

    texto = str(valor).strip().upper()

    correspondencia = re.search(
        r"([DC])\s{0,}$",
        texto,
        re.IGNORECASE
    )

    if correspondencia is None:
        return ""

    return correspondencia.group(1).upper()


def _converter_numero_grupo_CCGL(valor):
    """
    Converte um valor monetário para float.

    Formatos aceitos:
        2.063.141.576,53
        -2.063.141.576,53
        2063141576.53
        (2.063.141.576,53)
        2.063.141.576,53-
        R$ 2.063.141.576,53

    O indicador D/C será removido antes da conversão.
    """
    if valor is None:
        return 0.0

    try:
        if pd.isna(valor):
            return 0.0
    except (TypeError, ValueError):
        pass

    if isinstance(
        valor,
        (int, float)
    ):
        return round(
            float(valor),
            2
        )

    texto = str(valor).strip().upper()

    if texto in {
        "",
        "-",
        "--",
    }:
        return 0.0

    texto = re.sub(
        r"\s{0,}[DC]\s{0,}$",
        "",
        texto,
        flags=re.IGNORECASE
    )

    texto = (
        texto
        .replace("\xa0", "")
        .replace(" ", "")
        .replace("R$", "")
        .replace("$", "")
    )

    negativo_parenteses = (
        texto.startswith("(")
        and texto.endswith(")")
    )

    negativo_final = (
        texto.endswith("-")
        and texto != "-"
    )

    if negativo_parenteses:
        texto = texto[1:-1]

    if negativo_final:
        texto = texto[:-1]

    if "," in texto:
        texto = (
            texto
            .replace(".", "")
            .replace(",", ".")
        )

    try:
        numero = float(
            texto
        )
    except (ValueError, TypeError):
        return 0.0

    if negativo_parenteses or negativo_final:
        numero = -abs(
            numero
        )

    return round(
        numero,
        2
    )


def _aplicar_natureza_grupo_CCGL(valor):
    """
    Aplica a natureza contábil incorporada ao valor.

    Regras:
        D = positivo
        C = negativo

    Quando não houver indicador D/C, o sinal original será mantido.
    """
    numero_original = _converter_numero_grupo_CCGL(
        valor
    )

    natureza = _extrair_natureza_grupo_CCGL(
        valor
    )

    if natureza == "D":
        return round(
            abs(numero_original),
            2
        )

    if natureza == "C":
        return round(
            -abs(numero_original),
            2
        )

    return round(
        numero_original,
        2
    )


# ==============================================================================
# LEITURA DO EXCEL
# ==============================================================================

def _obter_engine_excel_grupo_CCGL(caminho_arquivo):
    """
    Define o mecanismo de leitura do arquivo Excel.
    """
    extensao = os.path.splitext(
        caminho_arquivo
    )[1].lower()

    if extensao == ".xls":
        return "xlrd"

    if extensao == ".xlsx":
        return "openpyxl"

    raise ValueError(
        f"A extensão '{extensao}' não é compatível com o grupo_CCGL."
    )


def _ler_excel_grupo_CCGL(caminho_arquivo):
    """
    Lê a primeira aba do arquivo sem considerar cabeçalho fixo.
    """
    nome_arquivo = os.path.basename(
        caminho_arquivo
    )

    try:
        df_origem = pd.read_excel(
            caminho_arquivo,
            sheet_name=0,
            header=None,
            dtype=object,
            engine=_obter_engine_excel_grupo_CCGL(
                caminho_arquivo
            )
        )
    except Exception as erro:
        raise ValueError(
            f"Não foi possível ler o arquivo do Grupo_CCGL "
            f"'{nome_arquivo}'. Erro: {erro}"
        ) from erro

    if df_origem.empty:
        raise ValueError(
            f"O arquivo do Grupo_CCGL '{nome_arquivo}' está vazio."
        )

    if df_origem.shape[1] < 7:
        raise ValueError(
            f"O arquivo do Grupo_CCGL '{nome_arquivo}' possui "
            f"{df_origem.shape[1]} coluna(s), mas são necessárias "
            "pelo menos 7 colunas, de A até G."
        )

    return df_origem


# ==============================================================================
# IDENTIFICAÇÃO DO CABEÇALHO
# ==============================================================================

def _linha_eh_cabecalho_grupo_CCGL(linha):
    """
    Identifica uma linha de cabeçalho do balancete.
    """
    textos_linha = [
        _normalizar_texto_comparacao_grupo_CCGL(
            valor
        )
        for valor in linha.tolist()
    ]

    texto_completo = " | ".join(
        texto
        for texto in textos_linha
        if texto
    )

    termos_cabecalho = [
        "CONTA",
        "DESCRICAO",
        "SALDO ANTERIOR",
        "DEBITO",
        "CREDITO",
        "MOV PERIODO",
        "MOV. PERIODO",
        "MOVIMENTO PERIODO",
        "SALDO ATUAL",
    ]

    quantidade_encontrada = sum(
        termo in texto_completo
        for termo in termos_cabecalho
    )

    return quantidade_encontrada >= 4


def _linha_possui_dados_grupo_CCGL(linha):
    """
    Verifica se uma linha representa uma conta válida.

    Layout:
        A = Conta
        B = Descrição
        C até G = valores
    """
    if len(linha) < 7:
        return False

    classificacao = _normalizar_classificacao_grupo_CCGL(
        linha.iloc[0]
    )

    descricao = _normalizar_texto_grupo_CCGL(
        linha.iloc[1]
    )

    if not classificacao:
        return False

    if not descricao:
        return False

    return True


def _localizar_inicio_dados_grupo_CCGL(
    df_origem,
    nome_arquivo
):
    """
    Localiza o início dos dados.

    Primeiro procura o cabeçalho.
    Se o cabeçalho não for localizado, procura a primeira conta válida.
    """
    for indice_linha in range(
        df_origem.shape[0]
    ):
        linha = df_origem.iloc[
            indice_linha
        ]

        if _linha_eh_cabecalho_grupo_CCGL(
            linha
        ):
            return indice_linha + 1

    for indice_linha in range(
        df_origem.shape[0]
    ):
        linha = df_origem.iloc[
            indice_linha
        ]

        if _linha_possui_dados_grupo_CCGL(
            linha
        ):
            return indice_linha

    raise ValueError(
        f"Não foi possível localizar o cabeçalho ou a primeira "
        f"conta válida no arquivo do Grupo_CCGL '{nome_arquivo}'."
    )


# ==============================================================================
# EXTRAÇÃO DOS REGISTROS
# ==============================================================================

def _extrair_registro_grupo_CCGL(linha):
    """
    Extrai uma conta do arquivo do Grupo_CCGL.

    Origem:
        A = Conta
        B = Descrição
        C = Saldo anterior
        D = Débito
        E = Crédito
        F = Movimento do período
        G = Saldo atual
    """
    if not _linha_possui_dados_grupo_CCGL(
        linha
    ):
        return None

    classificacao = _normalizar_classificacao_grupo_CCGL(
        linha.iloc[0]
    )

    descricao = _normalizar_texto_grupo_CCGL(
        linha.iloc[1]
    )

    saldo_anterior = _aplicar_natureza_grupo_CCGL(
        linha.iloc[2]
    )

    debito = _aplicar_natureza_grupo_CCGL(
        linha.iloc[3]
    )

    credito = _aplicar_natureza_grupo_CCGL(
        linha.iloc[4]
    )

    movimento = _aplicar_natureza_grupo_CCGL(
        linha.iloc[5]
    )

    saldo_acumulado = _aplicar_natureza_grupo_CCGL(
        linha.iloc[6]
    )

    return {
        "Conta": classificacao,
        "Nome": descricao,
        "Cód. Reduzido": classificacao,
        "Saldo Anterior": round(
            saldo_anterior,
            2
        ),
        "Débito": round(
            debito,
            2
        ),
        "Crédito": round(
            credito,
            2
        ),
        "Movimento": round(
            movimento,
            2
        ),
        "Saldo Acumulado": round(
            saldo_acumulado,
            2
        ),
    }


# ==============================================================================
# CONFERÊNCIA DOS REGISTROS
# ==============================================================================

def _calcular_movimento_esperado_grupo_CCGL(registro):
    """
    Calcula o movimento esperado usando débito e crédito.

    Se o crédito já estiver negativo:
        Movimento = Débito + Crédito

    Se o crédito estiver positivo:
        Movimento = Débito - Crédito
    """
    debito = round(
        registro["Débito"],
        2
    )

    credito = round(
        registro["Crédito"],
        2
    )

    if credito < 0:
        return round(
            debito + credito,
            2
        )

    return round(
        debito - credito,
        2
    )


def _conferir_registro_grupo_CCGL(registro):
    """
    Compara os saldos e movimentos informados.

    A conferência é apenas informativa e não altera o resultado.
    """
    movimento_esperado = _calcular_movimento_esperado_grupo_CCGL(
        registro
    )

    movimento_informado = round(
        registro["Movimento"],
        2
    )

    diferenca_movimento = round(
        movimento_informado
        - movimento_esperado,
        2
    )

    saldo_esperado = round(
        registro["Saldo Anterior"]
        + movimento_informado,
        2
    )

    diferenca_saldo = round(
        registro["Saldo Acumulado"]
        - saldo_esperado,
        2
    )

    return {
        "Diferença Movimento": diferenca_movimento,
        "Diferença Saldo": diferenca_saldo,
    }


# ==============================================================================
# MONTAGEM DO DATAFRAME
# ==============================================================================

def _montar_dataframe_grupo_CCGL(
    registros_validos
):
    """
    Monta o DataFrame final do Grupo_CCGL.
    """
    if not registros_validos:
        return pd.DataFrame(
            columns=COLUNAS_DESTINO_grupo_CCGL
        )

    df_destino = pd.DataFrame(
        registros_validos
    )

    df_destino.insert(
        0,
        "Atividade",
        "Geral"
    )

    colunas_monetarias = [
        "Saldo Anterior",
        "Débito",
        "Crédito",
        "Movimento",
        "Saldo Acumulado",
    ]

    for coluna in colunas_monetarias:
        df_destino[coluna] = (
            pd.to_numeric(
                df_destino[coluna],
                errors="coerce"
            )
            .fillna(0.0)
            .round(2)
        )

    return df_destino[
        COLUNAS_DESTINO_grupo_CCGL
    ].copy()


# ==============================================================================
# TRANSFORMAÇÃO DO BALANCETE do Grupo_CCGL
# ==============================================================================

def transformar_balancete_grupo_CCGL(caminho_arquivo):
    """
    Transforma o balancete Excel do sistema do Grupo_CCGL.

    Layout de destino:
        A = Atividade
        B = Conta
        C = Nome
        D = Cód. Reduzido
        E = Saldo Anterior
        F = Débito
        G = Crédito
        H = Movimento
        I = Saldo Acumulado
    """
    nome_arquivo = os.path.basename(
        caminho_arquivo
    )

    extensao = os.path.splitext(
        caminho_arquivo
    )[1].lower()

    if extensao not in EXTENSOES_VALIDAS_grupo_CCGL:
        raise ValueError(
            f"O arquivo '{nome_arquivo}' não é um arquivo Excel válido."
        )

    df_origem = _ler_excel_grupo_CCGL(
        caminho_arquivo
    )

    indice_inicio = _localizar_inicio_dados_grupo_CCGL(
        df_origem,
        nome_arquivo
    )

    df_dados = df_origem.iloc[
        indice_inicio:
    ].copy()

    df_dados.reset_index(
        drop=True,
        inplace=True
    )

    registros_validos = []
    linhas_ignoradas = []
    divergencias = []

    for indice_relativo, linha in df_dados.iterrows():
        if _linha_eh_cabecalho_grupo_CCGL(
            linha
        ):
            continue

        registro = _extrair_registro_grupo_CCGL(
            linha
        )

        if registro is None:
            if _classificacao_valida_grupo_CCGL(
                linha.iloc[0]
            ):
                linhas_ignoradas.append({
                    "Linha": (
                        indice_inicio
                        + indice_relativo
                        + 1
                    ),
                    "Conta": (
                        _normalizar_texto_grupo_CCGL(
                            linha.iloc[0]
                        )
                    ),
                    "Descrição": (
                        _normalizar_texto_grupo_CCGL(
                            linha.iloc[1]
                        )
                    ),
                })

            continue

        conferencia = _conferir_registro_grupo_CCGL(
            registro
        )

        if (
            abs(conferencia["Diferença Movimento"]) > 0.02
            or abs(conferencia["Diferença Saldo"]) > 0.02
        ):
            divergencias.append({
                "Linha": (
                    indice_inicio
                    + indice_relativo
                    + 1
                ),
                "Conta": registro["Conta"],
                "Nome": registro["Nome"],
                "Diferença Movimento": (
                    conferencia["Diferença Movimento"]
                ),
                "Diferença Saldo": (
                    conferencia["Diferença Saldo"]
                ),
            })

        registros_validos.append(
            registro
        )

    if not registros_validos:
        detalhe = ""

        if linhas_ignoradas:
            exemplos = linhas_ignoradas[:5]

            detalhe = (
                " Exemplos de linhas não processadas: "
                + " | ".join(
                    (
                        f"Linha {item['Linha']}: "
                        f"Conta={item['Conta']}; "
                        f"Descrição={item['Descrição']}"
                    )
                    for item in exemplos
                )
            )

        raise ValueError(
            f"Nenhuma conta válida foi encontrada no arquivo "
            f"do Grupo_CCGL '{nome_arquivo}'. Verifique se a conta está "
            "na coluna A, a descrição na coluna B e os valores "
            "nas colunas C até G."
            + detalhe
        )

    if linhas_ignoradas:
        print(
            f"[Grupo_CCGL] "
            f"{len(linhas_ignoradas)} linha(s) foram ignoradas "
            f"no arquivo '{nome_arquivo}'."
        )

        for item in linhas_ignoradas:
            print(
                f"[Grupo_CCGL] Linha {item['Linha']}: "
                f"Conta={item['Conta']}; "
                f"Descrição={item['Descrição']}"
            )

    if divergencias:
        print(
            f"[Grupo_CCGL] "
            f"{len(divergencias)} linha(s) apresentaram divergência "
            f"de movimento ou saldo no arquivo '{nome_arquivo}'."
        )

        for item in divergencias:
            print(
                f"[Grupo_CCGL] Linha {item['Linha']}: "
                f"Conta={item['Conta']}; "
                f"Nome={item['Nome']}; "
                f"Diferença movimento="
                f"{item['Diferença Movimento']:.2f}; "
                f"Diferença saldo="
                f"{item['Diferença Saldo']:.2f}"
            )

    df_destino = _montar_dataframe_grupo_CCGL(
        registros_validos
    )

    if df_destino.empty:
        raise ValueError(
            f"O arquivo do Grupo_CCGL '{nome_arquivo}' não gerou "
            "nenhum registro no layout de destino."
        )

    return df_destino


# ==============================================================================
# FUNÇÃO PÚBLICA DO Grupo_CCGL
# ==============================================================================

def _conta_eh_resultado_grupo_CCGL(conta):
    """
    Identifica as contas que devem receber o ajuste sequencial
    dos saldos de resultado.

    São consideradas:
        - contas iniciadas por 3;
        - contas iniciadas por 4;
        - contas iniciadas por 5;
        - conta 240205;
        - conta 240205001.

    A classificação é normalizada antes da comparação, removendo
    pontos e espaços.
    """
    conta_normalizada = (
        _normalizar_texto_grupo_CCGL(conta)
        .replace(".", "")
        .replace(" ", "")
    )

    if not conta_normalizada:
        return False

    if conta_normalizada in CONTAS_ADICIONAIS_RESULTADO_grupo_CCGL:
        return True

    return conta_normalizada.startswith(
        PREFIXOS_CONTAS_RESULTADO_grupo_CCGL
    )


def _obter_mes_ordenacao_grupo_CCGL(nome_aba):
    """
    Obtém o número do mês utilizado para ordenar as abas.

    Exemplos:
        01                -> 1
        02                -> 2
        B_03_BALANCETE    -> 3
        B_12.2026         -> 12

    Abas cujo mês não puder ser identificado serão colocadas
    depois das abas mensais.
    """
    nome_texto = str(nome_aba).strip()

    if re.fullmatch(
        r"0[1-9]|1[0-2]",
        nome_texto
    ):
        return int(nome_texto)

    correspondencia = re.match(
        r"^B_(0[1-9]|1[0-2])",
        nome_texto,
        re.IGNORECASE
    )

    if correspondencia is not None:
        return int(
            correspondencia.group(1)
        )

    return 99


def _arredondar_valor_grupo_CCGL(valor):
    """
    Converte um valor para número e arredonda para duas casas.

    Valores vazios ou inválidos são convertidos para zero.
    """
    numero = pd.to_numeric(
        pd.Series([valor]),
        errors="coerce"
    ).iloc[0]

    if pd.isna(numero):
        return 0.0

    return round(
        float(numero),
        2
    )


def _ajustar_saldos_resultado_grupo_CCGL(resultados):
    """
    Ajusta os saldos das contas de resultado em todos os meses.

    Regras:
        Janeiro:
            Saldo Anterior = 0,00
            Saldo Acumulado = Movimento de janeiro

        Meses posteriores:
            Saldo Anterior = Saldo Acumulado ajustado do mês anterior
            Saldo Acumulado = Saldo Anterior + Movimento do mês

    O movimento já existente no DataFrame é preservado.

    As contas patrimoniais não são alteradas.
    """
    if not resultados:
        return resultados

    itens_ordenados = sorted(
        resultados.items(),
        key=lambda item: (
            _obter_mes_ordenacao_grupo_CCGL(
                item[0]
            ),
            str(item[0]).casefold()
        )
    )

    saldos_resultado = {}
    resultados_ajustados = {}

    for nome_aba, dataframe_original in itens_ordenados:
        if dataframe_original is None:
            continue

        dataframe = dataframe_original.copy()

        colunas_obrigatorias = {
            "Conta",
            "Saldo Anterior",
            "Movimento",
            "Saldo Acumulado",
        }

        colunas_ausentes = (
            colunas_obrigatorias
            - set(dataframe.columns)
        )

        if colunas_ausentes:
            raise ValueError(
                f"Não foi possível ajustar as contas de resultado "
                f"na aba '{nome_aba}'. Colunas ausentes: "
                f"{', '.join(sorted(colunas_ausentes))}."
            )

        mes_atual = _obter_mes_ordenacao_grupo_CCGL(
            nome_aba
        )

        if mes_atual == 99:
            raise ValueError(
                f"Não foi possível identificar o mês da aba "
                f"'{nome_aba}'. Utilize arquivos iniciados por B_XX "
                f"para permitir o ajuste sequencial dos saldos."
            )

        for indice in dataframe.index:
            conta = _normalizar_texto_grupo_CCGL(
                dataframe.at[
                    indice,
                    "Conta"
                ]
            )

            if not _conta_eh_resultado_grupo_CCGL(
                conta
            ):
                continue

            movimento = _arredondar_valor_grupo_CCGL(
                dataframe.at[
                    indice,
                    "Movimento"
                ]
            )

            if mes_atual == 1:
                saldo_anterior_ajustado = 0.0

            else:
                saldo_anterior_ajustado = round(
                    saldos_resultado.get(
                        conta,
                        0.0
                    ),
                    2
                )

            saldo_acumulado_ajustado = round(
                saldo_anterior_ajustado
                + movimento,
                2
            )

            dataframe.at[
                indice,
                "Saldo Anterior"
            ] = saldo_anterior_ajustado

            dataframe.at[
                indice,
                "Saldo Acumulado"
            ] = saldo_acumulado_ajustado

            saldos_resultado[conta] = (
                saldo_acumulado_ajustado
            )

        dataframe["Saldo Anterior"] = (
            pd.to_numeric(
                dataframe["Saldo Anterior"],
                errors="coerce"
            )
            .fillna(0.0)
            .round(2)
        )

        dataframe["Movimento"] = (
            pd.to_numeric(
                dataframe["Movimento"],
                errors="coerce"
            )
            .fillna(0.0)
            .round(2)
        )

        dataframe["Saldo Acumulado"] = (
            pd.to_numeric(
                dataframe["Saldo Acumulado"],
                errors="coerce"
            )
            .fillna(0.0)
            .round(2)
        )

        resultados_ajustados[
            nome_aba
        ] = dataframe

    return resultados_ajustados

def processar(lista_arquivos):
    """
    Processa os arquivos selecionados para o sistema do Grupo_CCGL.

    Regras:
        - aceita arquivos XLS e XLSX;
        - ignora outras extensões;
        - cada arquivo gera uma aba;
        - arquivos B_XX tentam usar XX;
        - meses repetidos usam o nome completo do arquivo;
        - contas de resultado têm seus saldos reconstruídos
          sequencialmente a partir de janeiro;
        - janeiro considera saldo anterior igual a zero;
        - o saldo final de cada mês é transportado como saldo
          anterior do mês seguinte.
    """
    resultados = {}
    nomes_utilizados = set()

    if not lista_arquivos:
        raise ValueError(
            "Nenhum arquivo foi selecionado para o sistema do Grupo_CCGL."
        )

    arquivos_excel = [
        arquivo
        for arquivo in lista_arquivos
        if os.path.splitext(
            str(arquivo)
        )[1].lower() in EXTENSOES_VALIDAS_grupo_CCGL
    ]

    if not arquivos_excel:
        raise ValueError(
            "Nenhum arquivo Excel válido foi encontrado para o "
            "sistema do Grupo_CCGL. Selecione arquivos XLS ou XLSX."
        )

    arquivos_ordenados = sorted(
        arquivos_excel,
        key=lambda arquivo: (
            int(
                _obter_mes_grupo_CCGL(
                    arquivo
                )
            )
            if _obter_mes_grupo_CCGL(
                arquivo
            ) is not None
            else 99,
            os.path.basename(
                arquivo
            ).casefold()
        )
    )

    for arquivo in arquivos_ordenados:
        nome_aba = _gerar_nome_aba_grupo_CCGL(
            caminho_arquivo=arquivo,
            nomes_utilizados=nomes_utilizados
        )

        dataframe = transformar_balancete_grupo_CCGL(
            arquivo
        )

        if dataframe is None:
            raise ValueError(
                f"O arquivo '{os.path.basename(arquivo)}' não "
                "retornou um DataFrame."
            )

        if dataframe.empty:
            raise ValueError(
                f"O arquivo '{os.path.basename(arquivo)}' não "
                "retornou dados válidos para tabulação."
            )

        resultados[nome_aba] = dataframe

    if not resultados:
        raise ValueError(
            "Nenhum resultado foi gerado para o sistema do Grupo_CCGL."
        )

    resultados = _ajustar_saldos_resultado_grupo_CCGL(
        resultados
    )

    return resultados