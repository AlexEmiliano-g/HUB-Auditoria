import os
import re
import unicodedata

import pandas as pd


# ==============================================================================
# CONFIGURAÇÕES DO SISTEMA coopatrigo
# ==============================================================================

EXTENSOES_VALIDAS_coopatrigo = {
    ".xls",
    ".xlsx",
}

COLUNAS_DESTINO_coopatrigo = [
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

def _limpar_nome_aba_coopatrigo(nome):
    """
    Ajusta um texto para utilização como nome de aba do Excel.

    Regras do Excel:
    - não permite os caracteres: barra, barra invertida, asterisco,
      interrogação, dois-pontos e colchetes;
    - não permite nome vazio;
    - permite no máximo 31 caracteres.
    """
    nome_texto = str(nome).strip()

    nome_limpo = re.sub(
        r'[\\/*?:\[\]]',
        "_",
        nome_texto
    )

    if not nome_limpo:
        nome_limpo = "Sem nome"

    return nome_limpo[:31]


def _obter_nome_sem_extensao_coopatrigo(caminho_arquivo):
    """
    Retorna o nome do arquivo sem a extensão.
    """
    nome_arquivo = os.path.basename(
        caminho_arquivo
    )

    nome_sem_extensao = os.path.splitext(
        nome_arquivo
    )[0]

    return nome_sem_extensao.strip()


def _obter_mes_nome_arquivo_coopatrigo(caminho_arquivo):
    """
    Obtém o mês quando o arquivo começa com B_XX.

    Exemplos:
        B_01_BALANCETE.xlsx -> 01
        B_05.2026.xls       -> 05
        b_12-RELATORIO.xlsx -> 12
        BALANCETE.xlsx      -> None

    Somente os meses de 01 a 12 são aceitos.
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


def _registrar_nome_aba_coopatrigo(
    nome,
    nomes_utilizados
):
    """
    Registra um nome de aba quando ele ainda não está em uso.

    A comparação não diferencia letras maiúsculas e minúsculas.
    """
    nome_aba = _limpar_nome_aba_coopatrigo(
        nome
    )

    chave_comparacao = nome_aba.casefold()

    if chave_comparacao in nomes_utilizados:
        return None

    nomes_utilizados.add(
        chave_comparacao
    )

    return nome_aba


def _gerar_nome_aba_coopatrigo(
    caminho_arquivo,
    nomes_utilizados
):
    """
    Gera um nome exclusivo para a aba do arquivo tabulado.

    Regras:
    1. Se o arquivo começar com B_XX, tenta utilizar XX;
    2. Se o mês já estiver em uso, utiliza o nome completo do arquivo;
    3. Arquivos fora do padrão usam o nome sem extensão;
    4. Se o nome completo estiver duplicado, adiciona um sufixo.
    """
    nome_completo = _obter_nome_sem_extensao_coopatrigo(
        caminho_arquivo
    )

    mes = _obter_mes_nome_arquivo_coopatrigo(
        caminho_arquivo
    )

    if mes:
        nome_aba = _registrar_nome_aba_coopatrigo(
            mes,
            nomes_utilizados
        )

        if nome_aba is not None:
            return nome_aba

    nome_aba = _registrar_nome_aba_coopatrigo(
        nome_completo,
        nomes_utilizados
    )

    if nome_aba is not None:
        return nome_aba

    contador = 2

    while True:
        sufixo = f"_{contador}"
        limite_nome = 31 - len(sufixo)

        nome_candidato = (
            f"{nome_completo[:limite_nome]}{sufixo}"
        )

        nome_aba = _registrar_nome_aba_coopatrigo(
            nome_candidato,
            nomes_utilizados
        )

        if nome_aba is not None:
            return nome_aba

        contador += 1


# ==============================================================================
# NORMALIZAÇÃO DE TEXTOS
# ==============================================================================

def _normalizar_texto_coopatrigo(valor):
    """
    Normaliza um texto extraído do Excel.

    A função:
    - converte valores ausentes em texto vazio;
    - substitui espaços não separáveis;
    - reduz espaços consecutivos;
    - remove espaços nas extremidades.
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
        r"\s+",
        " ",
        texto
    )

    return texto.strip()


def _normalizar_texto_comparacao_coopatrigo(valor):
    """
    Normaliza um texto para comparação.

    A função:
    - normaliza espaços;
    - converte para letras maiúsculas;
    - remove acentos.

    Exemplos:
        Classificação -> CLASSIFICACAO
        Débito        -> DEBITO
        Crédito       -> CREDITO
        Código        -> CODIGO
    """
    texto = _normalizar_texto_coopatrigo(
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
# NORMALIZAÇÃO DA CLASSIFICAÇÃO E DO CÓDIGO
# ==============================================================================

def _normalizar_classificacao_coopatrigo(valor):
    """
    Normaliza a classificação contábil.

    Exemplos:
        1       -> 1
        1.0     -> 1
        1.01    -> 1.01
        1.01.01 -> 1.01.01

    Observação:
    Quando o Excel armazena uma classificação como número, pode haver
    perda dos zeros após o ponto. O ideal é que classificações como
    1.01 estejam armazenadas como texto no arquivo de origem.
    """
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
            return str(
                int(valor)
            )

        texto_float = format(
            valor,
            "f"
        )

        return (
            texto_float
            .rstrip("0")
            .rstrip(".")
        )

    texto = _normalizar_texto_coopatrigo(
        valor
    )

    if not texto:
        return ""

    texto = texto.replace(
        ",",
        "."
    )

    if re.fullmatch(
        r"\d+\.0+",
        texto
    ):
        return texto.split(
            ".",
            maxsplit=1
        )[0]

    return texto


def _normalizar_codigo_reduzido_coopatrigo(valor):
    """
    Normaliza o código reduzido.

    Exemplos:
        1       -> 1
        1.0     -> 1
        10000   -> 10000
        10000.0 -> 10000
    """
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
            return str(
                int(valor)
            )

        return str(valor)

    texto = _normalizar_texto_coopatrigo(
        valor
    )

    if not texto:
        return ""

    if re.fullmatch(
        r"\d+\.0+",
        texto
    ):
        return texto.split(
            ".",
            maxsplit=1
        )[0]

    return texto


# ==============================================================================
# VALIDAÇÃO DAS CLASSIFICAÇÕES
# ==============================================================================

def _classificacao_valida_coopatrigo(valor):
    """
    Verifica se o conteúdo representa uma classificação válida.

    Exemplos aceitos:
        1
        1.01
        1.01.01
        2.01.03.001
    """
    classificacao = _normalizar_classificacao_coopatrigo(
        valor
    )

    if not classificacao:
        return False

    return bool(
        re.fullmatch(
            r"\d+(?:\.\d+)*",
            classificacao
        )
    )


def _codigo_reduzido_valido_coopatrigo(valor):
    """
    Verifica se o código reduzido possui formato numérico.
    """
    codigo = _normalizar_codigo_reduzido_coopatrigo(
        valor
    )

    if not codigo:
        return False

    return bool(
        re.fullmatch(
            r"\d+",
            codigo
        )
    )


# ==============================================================================
# CONVERSÃO DOS VALORES MONETÁRIOS
# ==============================================================================

def _converter_numero_coopatrigo(valor):
    """
    Converte valores numéricos para float.

    Formatos aceitos:
        1394888784
        1394888784.00
        1.394.888.784,00
        -1.394.888.784,00
        (1.394.888.784,00)
        1.394.888.784,00-
        R$ 1.394.888.784,00

    Valores vazios ou inválidos são convertidos para zero.

    O resultado é arredondado para duas casas decimais.
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

    texto = str(valor).strip()

    if texto in {
        "",
        "-",
        "--",
    }:
        return 0.0

    texto = (
        texto
        .replace("\xa0", "")
        .replace(" ", "")
        .replace("R$", "")
        .replace("$", "")
    )

    negativo_por_parenteses = (
        texto.startswith("(")
        and texto.endswith(")")
    )

    negativo_no_final = (
        texto.endswith("-")
        and texto != "-"
    )

    if negativo_por_parenteses:
        texto = texto[1:-1]

    if negativo_no_final:
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

    if (
        negativo_por_parenteses
        or negativo_no_final
    ):
        numero = -abs(
            numero
        )

    return round(
        numero,
        2
    )


def _normalizar_natureza_coopatrigo(natureza):
    """
    Normaliza a natureza contábil.

    Valores aceitos:
        D
        C
        Débito
        Crédito
        Devedor
        Credor
        C/D
    """
    natureza_normalizada = (
        _normalizar_texto_comparacao_coopatrigo(
            natureza
        )
        .replace(".", "")
        .replace("/", "")
        .replace("\\", "")
        .strip()
    )

    if natureza_normalizada in {
        "D",
        "DEBITO",
        "DEVEDOR",
    }:
        return "D"

    if natureza_normalizada in {
        "C",
        "CREDITO",
        "CREDOR",
    }:
        return "C"

    return ""


def _aplicar_natureza_coopatrigo(
    valor,
    natureza
):
    """
    Aplica a natureza contábil ao saldo.

    Regras:
        D = valor positivo
        C = valor negativo

    Quando houver natureza válida, o sinal originalmente existente
    no valor será ignorado.

    Exemplos:
        1000 com D  -> 1000.00
        1000 com C  -> -1000.00
        -1000 com D -> 1000.00
        -1000 com C -> -1000.00
    """
    numero_original = _converter_numero_coopatrigo(
        valor
    )

    numero_absoluto = round(
        abs(numero_original),
        2
    )

    natureza_normalizada = _normalizar_natureza_coopatrigo(
        natureza
    )

    if natureza_normalizada == "D":
        return numero_absoluto

    if natureza_normalizada == "C":
        return round(
            -numero_absoluto,
            2
        )

    return round(
        numero_original,
        2
    )


# ==============================================================================
# LEITURA DO EXCEL
# ==============================================================================

def _obter_engine_excel_coopatrigo(caminho_arquivo):
    """
    Define o mecanismo de leitura do arquivo Excel.

    Arquivos XLS utilizam xlrd.
    Arquivos XLSX utilizam openpyxl.
    """
    extensao = os.path.splitext(
        caminho_arquivo
    )[1].lower()

    if extensao == ".xls":
        return "xlrd"

    if extensao == ".xlsx":
        return "openpyxl"

    raise ValueError(
        f"A extensão '{extensao}' não é compatível com o Coopatrigo."
    )


def _ler_excel_coopatrigo(caminho_arquivo):
    """
    Lê a primeira planilha do arquivo Excel sem considerar
    uma linha fixa de cabeçalho.
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
            engine=_obter_engine_excel_coopatrigo(
                caminho_arquivo
            )
        )

    except Exception as erro:
        raise ValueError(
            f"Não foi possível ler o arquivo Coopatrigo "
            f"'{nome_arquivo}'. Erro: {erro}"
        ) from erro

    if df_origem.empty:
        raise ValueError(
            f"O arquivo Coopatrigo '{nome_arquivo}' está vazio."
        )

    if df_origem.shape[1] < 9:
        raise ValueError(
            f"O arquivo Coopatrigo '{nome_arquivo}' possui "
            f"{df_origem.shape[1]} coluna(s), mas são necessárias "
            "pelo menos 9 colunas, de A até I."
        )

    return df_origem


# ==============================================================================
# IDENTIFICAÇÃO DO CABEÇALHO
# ==============================================================================

def _linha_eh_cabecalho_coopatrigo(linha):
    """
    Verifica se uma linha parece ser o cabeçalho do balancete.

    Títulos esperados:
        Classificação
        Código
        Nome
        Saldo Anterior
        C/D
        Saldo Débito
        Saldo Crédito
        Saldo Atual
    """
    textos_linha = [
        _normalizar_texto_comparacao_coopatrigo(
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
        "CLASSIFICACAO",
        "CODIGO",
        "NOME",
        "SALDO ANTERIOR",
        "SALDO DEBITO",
        "SALDO CREDITO",
        "SALDO ATUAL",
    ]

    quantidade_encontrada = sum(
        termo in texto_completo
        for termo in termos_cabecalho
    )

    return quantidade_encontrada >= 4


def _linha_possui_dados_coopatrigo(linha):
    """
    Verifica se uma linha representa uma conta contábil válida.

    Posições utilizadas:
        A = classificação
        B = código reduzido
        C = descrição
    """
    if len(linha) < 9:
        return False

    classificacao = linha.iloc[0]
    codigo_reduzido = linha.iloc[1]

    descricao = _normalizar_texto_coopatrigo(
        linha.iloc[2]
    )

    if not _classificacao_valida_coopatrigo(
        classificacao
    ):
        return False

    if not _codigo_reduzido_valido_coopatrigo(
        codigo_reduzido
    ):
        return False

    if not descricao:
        return False

    return True


def _localizar_inicio_dados_coopatrigo(
    df_origem,
    nome_arquivo
):
    """
    Localiza o início dos dados do balancete.

    Estratégia 1:
        localiza a linha de cabeçalho e inicia na linha seguinte.

    Estratégia 2:
        caso o cabeçalho esteja distribuído em diversas linhas,
        localiza diretamente a primeira conta válida.
    """
    for indice_linha in range(
        df_origem.shape[0]
    ):
        linha = df_origem.iloc[
            indice_linha
        ]

        if _linha_eh_cabecalho_coopatrigo(
            linha
        ):
            return indice_linha + 1

    for indice_linha in range(
        df_origem.shape[0]
    ):
        linha = df_origem.iloc[
            indice_linha
        ]

        if _linha_possui_dados_coopatrigo(
            linha
        ):
            return indice_linha

    raise ValueError(
        f"Não foi possível localizar o cabeçalho ou a primeira "
        f"conta válida no arquivo Coopatrigo '{nome_arquivo}'."
    )


# ==============================================================================
# EXTRAÇÃO DAS LINHAS CONTÁBEIS
# ==============================================================================

def _extrair_registro_coopatrigo(linha):
    """
    Extrai e transforma uma linha contábil do Coopatrigo.

    Origem:
        A = Classificação
        B = Código
        C = Nome
        D = Saldo Anterior
        E = C/D do Saldo Anterior
        F = Saldo Débito
        G = Saldo Crédito
        H = Saldo Atual
        I = C/D do Saldo Atual
    """
    if not _linha_possui_dados_coopatrigo(
        linha
    ):
        return None

    classificacao = _normalizar_classificacao_coopatrigo(
        linha.iloc[0]
    )

    codigo_reduzido = _normalizar_codigo_reduzido_coopatrigo(
        linha.iloc[1]
    )

    descricao = _normalizar_texto_coopatrigo(
        linha.iloc[2]
    )

    saldo_anterior = _aplicar_natureza_coopatrigo(
        linha.iloc[3],
        linha.iloc[4]
    )

    debito = round(
        abs(
            _converter_numero_coopatrigo(
                linha.iloc[5]
            )
        ),
        2
    )

    credito = round(
        -abs(
            _converter_numero_coopatrigo(
                linha.iloc[6]
            )
        ),
        2
    )

    movimento = round(
        debito + credito,
        2
    )

    saldo_acumulado = _aplicar_natureza_coopatrigo(
        linha.iloc[7],
        linha.iloc[8]
    )

    return {
        "Conta": classificacao,
        "Nome": descricao,
        "Cód. Reduzido": codigo_reduzido,
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
# CONSTRUÇÃO DO DATAFRAME DE DESTINO
# ==============================================================================

def _montar_dataframe_destino_coopatrigo(
    registros_validos
):
    """
    Monta o DataFrame final do Coopatrigo na ordem exigida pelo HUB.
    """
    if not registros_validos:
        return pd.DataFrame(
            columns=COLUNAS_DESTINO_coopatrigo
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

    df_destino["Débito"] = (
        df_destino["Débito"]
        .abs()
        .round(2)
    )

    df_destino["Crédito"] = (
        -df_destino["Crédito"].abs()
    ).round(2)

    df_destino["Movimento"] = (
        df_destino["Débito"]
        + df_destino["Crédito"]
    ).round(2)

    return df_destino[
        COLUNAS_DESTINO_coopatrigo
    ].copy()


# ==============================================================================
# TRANSFORMAÇÃO DO BALANCETE COOPATRIGO
# ==============================================================================

def transformar_balancete_coopatrigo(caminho_arquivo):
    """
    Transforma um arquivo Excel do cliente Coopatrigo.

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

    Regras:
        - ignora informações anteriores ao balancete;
        - ignora linhas vazias e cabeçalhos repetidos;
        - natureza D gera saldo positivo;
        - natureza C gera saldo negativo;
        - débito fica positivo;
        - crédito fica negativo;
        - Movimento = Débito + Crédito;
        - valores são arredondados para duas casas decimais.
    """
    nome_arquivo = os.path.basename(
        caminho_arquivo
    )

    extensao = os.path.splitext(
        caminho_arquivo
    )[1].lower()

    if extensao not in EXTENSOES_VALIDAS_coopatrigo:
        raise ValueError(
            f"O arquivo '{nome_arquivo}' não é um arquivo Excel válido."
        )

    df_origem = _ler_excel_coopatrigo(
        caminho_arquivo
    )

    indice_inicio = _localizar_inicio_dados_coopatrigo(
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

    for indice_relativo, linha in df_dados.iterrows():
        if _linha_eh_cabecalho_coopatrigo(
            linha
        ):
            continue

        registro = _extrair_registro_coopatrigo(
            linha
        )

        if registro is not None:
            registros_validos.append(
                registro
            )

            continue

        if _classificacao_valida_coopatrigo(
            linha.iloc[0]
        ):
            linhas_ignoradas.append({
                "Linha": (
                    indice_inicio
                    + indice_relativo
                    + 1
                ),
                "Classificação": (
                    _normalizar_texto_coopatrigo(
                        linha.iloc[0]
                    )
                ),
                "Código": (
                    _normalizar_texto_coopatrigo(
                        linha.iloc[1]
                    )
                ),
                "Nome": (
                    _normalizar_texto_coopatrigo(
                        linha.iloc[2]
                    )
                ),
            })

    if not registros_validos:
        detalhe = ""

        if linhas_ignoradas:
            exemplos = linhas_ignoradas[:5]

            detalhe = (
                " Exemplos de linhas não processadas: "
                + " | ".join(
                    (
                        f"Linha {item['Linha']}: "
                        f"Classificação={item['Classificação']}; "
                        f"Código={item['Código']}; "
                        f"Nome={item['Nome']}"
                    )
                    for item in exemplos
                )
            )

        raise ValueError(
            f"Nenhuma conta válida foi encontrada no arquivo "
            f"Coopatrigo '{nome_arquivo}'. Verifique se a classificação "
            "está na coluna A, o código na coluna B, o nome na coluna C "
            "e os valores estão nas colunas D até I."
            + detalhe
        )

    if linhas_ignoradas:
        print(
            f"[Coopatrigo] "
            f"{len(linhas_ignoradas)} linha(s) com classificação "
            f"foram ignoradas no arquivo '{nome_arquivo}'."
        )

        for item in linhas_ignoradas:
            print(
                f"[Coopatrigo] Linha {item['Linha']}: "
                f"Classificação={item['Classificação']}; "
                f"Código={item['Código']}; "
                f"Nome={item['Nome']}"
            )

    df_destino = _montar_dataframe_destino_coopatrigo(
        registros_validos
    )

    if df_destino.empty:
        raise ValueError(
            f"O arquivo Coopatrigo '{nome_arquivo}' não gerou "
            "nenhum registro no layout de destino."
        )

    return df_destino


# ==============================================================================
# FUNÇÃO PÚBLICA DO SISTEMA COOPATRIGO
# ==============================================================================

def processar(lista_arquivos):
    """
    Processa os arquivos selecionados para o Cliente Coopatrigo.

    Regras:
        - processa apenas XLS e XLSX;
        - ignora outras extensões;
        - cada arquivo válido gera uma aba;
        - arquivos B_XX utilizam inicialmente XX;
        - meses repetidos utilizam o nome completo do arquivo;
        - nomes totalmente repetidos recebem sufixo numérico.

    Retorno:
        Dicionário no formato:

        {
            "Nome da aba": DataFrame
        }
    """
    resultados = {}
    nomes_utilizados = set()

    if not lista_arquivos:
        raise ValueError(
            "Nenhum arquivo foi selecionado para o cliente Coopatrigo."
        )

    arquivos_excel = [
        arquivo
        for arquivo in lista_arquivos
        if os.path.splitext(
            str(arquivo)
        )[1].lower() in EXTENSOES_VALIDAS_coopatrigo
    ]

    if not arquivos_excel:
        raise ValueError(
            "Nenhum arquivo Excel válido foi encontrado para o "
            "cliente Coopatrigo. Selecione arquivos XLS ou XLSX."
        )

    for arquivo in arquivos_excel:
        nome_aba = _gerar_nome_aba_coopatrigo(
            caminho_arquivo=arquivo,
            nomes_utilizados=nomes_utilizados
        )

        dataframe = transformar_balancete_coopatrigo(
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
            "Nenhum resultado foi gerado para o cliente Coopatrigo."
        )

    return resultados