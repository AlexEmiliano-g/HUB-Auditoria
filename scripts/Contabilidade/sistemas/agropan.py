import os
import re
import unicodedata

import pandas as pd


# ==============================================================================
# FUNÇÕES PARA OS NOMES DAS ABAS
# ==============================================================================

def _limpar_nome_aba_agropan(nome):
    """
    Ajusta um texto para utilização como nome de aba do Excel.

    Regras:
    - substitui caracteres inválidos por sublinhado;
    - impede nomes vazios;
    - limita o nome a 31 caracteres.
    """
    nome_limpo = re.sub(
        r'[\\/*?:\[\]]',
        "_",
        str(nome).strip()
    )

    if not nome_limpo:
        nome_limpo = "Sem nome"

    return nome_limpo[:31]


def _obter_nome_sem_extensao_agropan(caminho_arquivo):
    """
    Retorna o nome do arquivo sem a extensão.
    """
    nome_arquivo = os.path.basename(caminho_arquivo)
    nome_sem_extensao = os.path.splitext(nome_arquivo)[0]

    return nome_sem_extensao.strip()


def _obter_mes_agropan(caminho_arquivo):
    """
    Extrai o mês quando o nome do arquivo começa com B_XX.

    Exemplos:
        B_01_BALANCETE.xlsx -> 01
        B_05.2026.xls       -> 05
        BALANCETE.xlsx      -> None
    """
    nome_arquivo = os.path.basename(caminho_arquivo)

    correspondencia = re.match(
        r"^B_(0[1-9]|1[0-2])",
        nome_arquivo,
        re.IGNORECASE
    )

    if correspondencia is None:
        return None

    return correspondencia.group(1)


def _registrar_nome_aba_agropan(nome, nomes_utilizados):
    """
    Registra um nome de aba caso ainda não esteja em uso.

    A comparação não diferencia letras maiúsculas e minúsculas.
    """
    nome_aba = _limpar_nome_aba_agropan(nome)
    chave_nome = nome_aba.casefold()

    if chave_nome in nomes_utilizados:
        return None

    nomes_utilizados.add(chave_nome)

    return nome_aba


def _gerar_nome_aba_agropan(caminho_arquivo, nomes_utilizados):
    """
    Gera um nome exclusivo para a aba de destino.

    Regras:
    1. Arquivos B_XX tentam utilizar XX;
    2. Se XX já estiver em uso, utiliza o nome completo do arquivo;
    3. Arquivos fora do padrão utilizam o nome sem extensão;
    4. Duplicidades recebem um sufixo numérico.
    """
    nome_completo = _obter_nome_sem_extensao_agropan(
        caminho_arquivo
    )

    mes = _obter_mes_agropan(
        caminho_arquivo
    )

    if mes:
        nome_aba = _registrar_nome_aba_agropan(
            mes,
            nomes_utilizados
        )

        if nome_aba is not None:
            return nome_aba

    nome_aba = _registrar_nome_aba_agropan(
        nome_completo,
        nomes_utilizados
    )

    if nome_aba is not None:
        return nome_aba

    contador = 2

    while True:
        sufixo = f"_{contador}"
        limite_nome = 31 - len(sufixo)

        nome_com_sufixo = (
            f"{nome_completo[:limite_nome]}{sufixo}"
        )

        nome_aba = _registrar_nome_aba_agropan(
            nome_com_sufixo,
            nomes_utilizados
        )

        if nome_aba is not None:
            return nome_aba

        contador += 1


# ==============================================================================
# FUNÇÕES DE NORMALIZAÇÃO
# ==============================================================================

def _normalizar_texto_agropan(valor):
    """
    Normaliza os textos extraídos do cliente Agropan.

    A função:
    - converte valores ausentes em texto vazio;
    - substitui espaços não separáveis;
    - reduz espaços consecutivos;
    - remove espaços das extremidades.
    """
    if valor is None or pd.isna(valor):
        return ""

    texto = str(valor).replace("\xa0", " ")
    texto = re.sub(r"\s+", " ", texto)

    return texto.strip()


def _normalizar_texto_comparacao_agropan(valor):
    """
    Normaliza textos usados na localização do cabeçalho.

    Além da limpeza de espaços, a função:
    - converte para letras maiúsculas;
    - remove acentos.
    """
    texto = _normalizar_texto_agropan(valor).upper()

    if not texto:
        return ""

    texto = unicodedata.normalize(
        "NFKD",
        texto
    )

    texto = "".join(
        caractere
        for caractere in texto
        if not unicodedata.combining(caractere)
    )

    return texto


def _normalizar_codigo_agropan(valor):
    """
    Normaliza classificações e códigos reduzidos.

    Exemplos:
        10000.0 -> 10000
        1.01    -> 1.01
        1       -> 1
    """
    texto = _normalizar_texto_agropan(valor)

    if not texto:
        return ""

    if re.fullmatch(r"\d+\.0+", texto):
        return texto.split(
            ".",
            maxsplit=1
        )[0]

    return texto


# ==============================================================================
# CONVERSÃO DE VALORES
# ==============================================================================

def _converter_numero_agropan(valor):
    """
    Converte um valor monetário para float.

    Exemplos:
        905.904.511,18  -> 905904511.18
        -849.264.526,46 -> -849264526.46
        0,00            -> 0.0
        -               -> 0.0
        (1.250,50)      -> -1250.50

    O resultado é arredondado para duas casas decimais.
    """
    if valor is None or pd.isna(valor):
        return 0.0

    if isinstance(valor, (int, float)):
        return round(
            float(valor),
            2
        )

    texto = str(valor).strip()

    if texto in {"", "-", "--"}:
        return 0.0

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
        numero = float(texto)

        if negativo_parenteses or negativo_final:
            numero = -abs(numero)

        return round(
            numero,
            2
        )

    except (ValueError, TypeError):
        return 0.0


def _converter_saldo_agropan(valor):
    """
    Converte um saldo com natureza D ou C incorporada ao valor.

    Regras:
        D = positivo
        C = negativo

    Exemplos:
        1.218.650.132,64D -> 1218650132.64
        1.218.650.132,64C -> -1218650132.64
        -1.000,00D        -> 1000.00
        -1.000,00C        -> -1000.00

    Quando houver D ou C, o sinal numérico original será ignorado.
    """
    if valor is None or pd.isna(valor):
        return 0.0

    if isinstance(valor, (int, float)):
        return round(
            float(valor),
            2
        )

    texto = str(valor).strip().upper()

    if texto in {"", "-", "--"}:
        return 0.0

    natureza = ""

    correspondencia_natureza = re.search(
        r"([DC])\s*$",
        texto,
        re.IGNORECASE
    )

    if correspondencia_natureza is not None:
        natureza = (
            correspondencia_natureza
            .group(1)
            .upper()
        )

        texto = texto[
            :correspondencia_natureza.start()
        ].strip()

    numero = _converter_numero_agropan(
        texto
    )

    if natureza == "D":
        return round(
            abs(numero),
            2
        )

    if natureza == "C":
        return round(
            -abs(numero),
            2
        )

    return round(
        numero,
        2
    )


# ==============================================================================
# LEITURA E LOCALIZAÇÃO DO CABEÇALHO
# ==============================================================================

def _obter_engine_excel_agropan(caminho_arquivo):
    """
    Define o mecanismo utilizado para leitura do arquivo Excel.

    XLS utiliza xlrd.
    XLSX utiliza openpyxl.
    """
    extensao = os.path.splitext(
        caminho_arquivo
    )[1].lower()

    if extensao == ".xls":
        return "xlrd"

    return "openpyxl"


def _localizar_cabecalho_agropan(df_origem, nome_arquivo):
    """
    Localiza a linha do cabeçalho do balancete.

    O layout esperado possui os títulos nas posições:
        A = Classificação
        C = Conta
        D = Nome
        H = Saldo Anterior
        J = Débito
        K = Crédito
        M = Saldo Atual

    A função tenta primeiro validar as posições conhecidas.
    Se o cabeçalho estiver dividido em mais de uma linha, localiza
    a primeira linha de dados por meio da classificação.
    """
    for indice_linha in range(df_origem.shape[0]):
        linha = df_origem.iloc[indice_linha]

        classificacao = _normalizar_texto_comparacao_agropan(
            linha.iloc[0]
        )

        conta = _normalizar_texto_comparacao_agropan(
            linha.iloc[2]
        )

        nome = _normalizar_texto_comparacao_agropan(
            linha.iloc[3]
        )

        saldo_anterior = _normalizar_texto_comparacao_agropan(
            linha.iloc[7]
        )

        debito = _normalizar_texto_comparacao_agropan(
            linha.iloc[9]
        )

        credito = _normalizar_texto_comparacao_agropan(
            linha.iloc[10]
        )

        saldo_atual = _normalizar_texto_comparacao_agropan(
            linha.iloc[12]
        )

        classificacao_valida = (
            "CLASSIFICACAO" in classificacao
        )

        conta_valida = (
            conta == "CONTA"
            or "CONTA" in conta
        )

        nome_valido = (
            nome == "NOME"
            or "NOME" in nome
        )

        saldo_anterior_valido = (
            "SALDO ANTERIOR" in saldo_anterior
        )

        debito_valido = (
            "DEBITO" in debito
        )

        credito_valido = (
            "CREDITO" in credito
        )

        saldo_atual_valido = (
            "SALDO ATUAL" in saldo_atual
        )

        quantidade_validacoes = sum([
            classificacao_valida,
            conta_valida,
            nome_valido,
            saldo_anterior_valido,
            debito_valido,
            credito_valido,
            saldo_atual_valido,
        ])

        if quantidade_validacoes >= 4:
            return indice_linha

    # Segunda estratégia:
    # procura a primeira linha que tenha classificação, código reduzido,
    # descrição e pelo menos um valor monetário.
    for indice_linha in range(df_origem.shape[0]):
        linha = df_origem.iloc[indice_linha]

        classificacao = _normalizar_codigo_agropan(
            linha.iloc[0]
        )

        codigo_reduzido = _normalizar_codigo_agropan(
            linha.iloc[2]
        )

        descricao = _normalizar_texto_agropan(
            linha.iloc[3]
        )

        saldo_anterior = _normalizar_texto_agropan(
            linha.iloc[7]
        )

        classificacao_valida = bool(
            re.fullmatch(
                r"\d+(?:\.\d+)*",
                classificacao
            )
        )

        codigo_valido = bool(
            re.fullmatch(
                r"\d+",
                codigo_reduzido
            )
        )

        descricao_valida = bool(
            descricao
        )

        saldo_valido = (
            bool(saldo_anterior)
            and bool(
                re.search(
                    r"\d",
                    saldo_anterior
                )
            )
        )

        if (
            classificacao_valida
            and codigo_valido
            and descricao_valida
            and saldo_valido
        ):
            # Retorna a linha anterior porque a transformação inicia
            # após o índice retornado.
            return indice_linha - 1

    raise ValueError(
        f"Não foi possível localizar o início do balancete "
        f"no arquivo '{nome_arquivo}'."
    )


def _linha_eh_cabecalho_agropan(linha):
    """
    Identifica linhas de cabeçalho repetidas no decorrer do arquivo.
    """
    textos = [
        _normalizar_texto_comparacao_agropan(valor)
        for valor in linha.tolist()
    ]

    texto_completo = " | ".join(
        texto
        for texto in textos
        if texto
    )

    termos_cabecalho = [
        "CLASSIFICACAO",
        "SALDO ANTERIOR",
        "SALDO ATUAL",
        "DEBITO",
        "CREDITO",
    ]

    quantidade_encontrada = sum(
        termo in texto_completo
        for termo in termos_cabecalho
    )

    return quantidade_encontrada >= 2


def _linha_possui_dados_agropan(linha):
    """
    Verifica se uma linha representa uma conta contábil.

    A classificação deve estar na coluna A, o código reduzido
    na coluna C e a descrição na coluna D.
    """
    classificacao = _normalizar_codigo_agropan(
        linha.iloc[0]
    )

    codigo_reduzido = _normalizar_codigo_agropan(
        linha.iloc[2]
    )

    descricao = _normalizar_texto_agropan(
        linha.iloc[3]
    )

    classificacao_valida = bool(
        re.fullmatch(
            r"\d+(?:\.\d+)*",
            classificacao
        )
    )

    codigo_valido = bool(
        re.fullmatch(
            r"\d+",
            codigo_reduzido
        )
    )

    if not classificacao_valida:
        return False

    if not codigo_valido:
        return False

    if not descricao:
        return False

    return True


# ==============================================================================
# TRANSFORMAÇÃO DO BALANCETE AGROPAN
# ==============================================================================

def transformar_balancete_agropan(caminho_arquivo):
    """
    Transforma um balancete Excel do Cliente Agropan.

    Layout da origem:
        A = Classificação
        B = vazia
        C = Conta ou código reduzido
        D = Nome
        E = vazia
        F = vazia
        G = vazia
        H = Saldo Anterior
        I = vazia
        J = Débito
        K = Crédito
        L = vazia
        M = Saldo Atual

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

    Mapeamento:
        Valor fixo Geral      -> Atividade
        Origem A              -> Conta
        Origem D              -> Nome
        Origem C              -> Cód. Reduzido
        Origem H              -> Saldo Anterior
        Origem J              -> Débito
        Origem K              -> Crédito
        Débito mais Crédito   -> Movimento
        Origem M              -> Saldo Acumulado

    Regras:
        - ignora o cabeçalho inicial;
        - preserva as colunas vazias da origem;
        - saldo terminado em D é positivo;
        - saldo terminado em C é negativo;
        - débito é positivo;
        - crédito é negativo;
        - Movimento = Débito + Crédito;
        - valores são arredondados para duas casas decimais.
    """
    nome_arquivo = os.path.basename(
        caminho_arquivo
    )

    extensao = os.path.splitext(
        caminho_arquivo
    )[1].lower()

    if extensao not in {".xls", ".xlsx"}:
        raise ValueError(
            f"O arquivo '{nome_arquivo}' não é um arquivo Excel válido."
        )

    try:
        df_origem = pd.read_excel(
            caminho_arquivo,
            sheet_name=0,
            header=None,
            dtype=object,
            engine=_obter_engine_excel_agropan(
                caminho_arquivo
            )
        )

    except Exception as erro:
        raise ValueError(
            f"Não foi possível ler o arquivo Agropan "
            f"'{nome_arquivo}'. Erro: {erro}"
        ) from erro

    if df_origem.empty:
        raise ValueError(
            f"O arquivo Agropan '{nome_arquivo}' está vazio."
        )

    if df_origem.shape[1] < 13:
        raise ValueError(
            f"O arquivo Agropan '{nome_arquivo}' possui "
            f"{df_origem.shape[1]} coluna(s), mas são necessárias "
            "pelo menos 13 colunas, de A até M."
        )

    indice_cabecalho = _localizar_cabecalho_agropan(
        df_origem,
        nome_arquivo
    )

    df_dados = df_origem.iloc[
        max(indice_cabecalho + 1, 0):
    ].copy()

    df_dados.reset_index(
        drop=True,
        inplace=True
    )

    registros_validos = []

    for _, linha in df_dados.iterrows():
        if _linha_eh_cabecalho_agropan(
            linha
        ):
            continue

        if not _linha_possui_dados_agropan(
            linha
        ):
            continue

        classificacao = _normalizar_codigo_agropan(
            linha.iloc[0]
        )

        codigo_reduzido = _normalizar_codigo_agropan(
            linha.iloc[2]
        )

        descricao = _normalizar_texto_agropan(
            linha.iloc[3]
        )

        saldo_anterior = _converter_saldo_agropan(
            linha.iloc[7]
        )

        debito = round(
            abs(
                _converter_numero_agropan(
                    linha.iloc[9]
                )
            ),
            2
        )

        credito = round(
            -abs(
                _converter_numero_agropan(
                    linha.iloc[10]
                )
            ),
            2
        )

        movimento = round(
            debito + credito,
            2
        )

        saldo_acumulado = _converter_saldo_agropan(
            linha.iloc[12]
        )

        registros_validos.append({
            "Conta": classificacao,
            "Nome": descricao,
            "Cód. Reduzido": codigo_reduzido,
            "Saldo Anterior": saldo_anterior,
            "Débito": debito,
            "Crédito": credito,
            "Movimento": movimento,
            "Saldo Acumulado": saldo_acumulado,
        })

    if not registros_validos:
        raise ValueError(
            f"Nenhuma conta válida foi encontrada no arquivo "
            f"Agropan '{nome_arquivo}'. Verifique se a classificação "
            "está na coluna A, o código reduzido na coluna C, o nome "
            "na coluna D e os valores nas colunas H, J, K e M."
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

    df_destino = df_destino[
        [
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
    ]

    return df_destino


# ==============================================================================
# FUNÇÃO PÚBLICA DO CLIENTE AGROPAN
# ==============================================================================

def processar(lista_arquivos):
    """
    Processa os arquivos selecionados para o cliente Agropan.

    Regras:
    - processa somente arquivos XLS e XLSX;
    - ignora arquivos de outras extensões;
    - cada arquivo válido gera uma aba;
    - arquivos B_XX utilizam inicialmente XX como nome da aba;
    - meses repetidos utilizam o nome completo do arquivo;
    - arquivos fora do padrão utilizam o nome sem extensão.
    """
    resultados = {}
    nomes_utilizados = set()

    if not lista_arquivos:
        raise ValueError(
            "Nenhum arquivo foi selecionado para o cliente Agropan."
        )

    arquivos_excel = [
        arquivo
        for arquivo in lista_arquivos
        if os.path.splitext(str(arquivo))[1].lower()
        in {".xls", ".xlsx"}
    ]

    if not arquivos_excel:
        raise ValueError(
            "Nenhum arquivo Excel válido foi encontrado para o "
            "cliente Agropan. Selecione arquivos XLS ou XLSX."
        )

    for arquivo in arquivos_excel:
        nome_aba = _gerar_nome_aba_agropan(
            caminho_arquivo=arquivo,
            nomes_utilizados=nomes_utilizados
        )

        dataframe = transformar_balancete_agropan(
            arquivo
        )

        if dataframe is None or dataframe.empty:
            raise ValueError(
                f"O arquivo '{os.path.basename(arquivo)}' não "
                "retornou dados válidos para tabulação."
            )

        resultados[nome_aba] = dataframe

    if not resultados:
        raise ValueError(
            "Nenhum resultado foi gerado para o cliente Agropan."
        )

    return resultados