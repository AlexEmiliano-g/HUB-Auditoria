import os
import re
import pandas as pd

def _limpar_nome_aba(nome):
    """
    Ajusta um nome para utilização como aba do Excel.

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


def _obter_mes_nome_planilha(nome_planilha):
    """
    Extrai o mês quando o nome da planilha começa com B_XX.

    Exemplos:
        B_01              -> 01
        B_02.2026         -> 02
        B_03_BALANCETE    -> 03
        b_04              -> 04
        MAIO              -> None

    Somente meses entre 01 e 12 são considerados válidos.
    """
    nome_planilha = str(nome_planilha).strip()

    correspondencia = re.match(
        r"^B_(0[1-9]|1[0-2])",
        nome_planilha,
        re.IGNORECASE
    )

    if correspondencia is None:
        return None

    return correspondencia.group(1)


def _registrar_nome_aba(nome, nomes_utilizados):
    """
    Registra um nome de aba caso ainda não esteja em uso.

    A comparação não diferencia letras maiúsculas e minúsculas.
    """
    nome_aba = _limpar_nome_aba(nome)
    chave_nome = nome_aba.casefold()

    if chave_nome in nomes_utilizados:
        return None

    nomes_utilizados.add(chave_nome)

    return nome_aba


def _gerar_nome_aba_unico(
    nome_planilha,
    nome_arquivo,
    nomes_utilizados
):
    """
    Gera o nome da aba do arquivo tabulado.

    Regras:
    1. Se a planilha começar com B_XX, tenta utilizar XX;
    2. Se o mês já estiver em uso, utiliza o nome completo da planilha;
    3. Se o nome da planilha também estiver em uso, combina o nome da
       planilha com o nome do arquivo;
    4. Se ainda houver duplicidade, acrescenta um sufixo numérico.

    Exemplos:
        B_01                -> 01
        B_01_AJUSTE         -> B_01_AJUSTE, quando 01 já existe
        B_02.2026           -> 02
        MAIO                -> MAIO
    """
    nome_planilha = str(nome_planilha).strip()
    mes = _obter_mes_nome_planilha(nome_planilha)

    # Se seguir o padrão B_XX, tenta primeiro utilizar somente o mês.
    if mes:
        nome_aba = _registrar_nome_aba(
            mes,
            nomes_utilizados
        )

        if nome_aba is not None:
            return nome_aba

    # Se o mês já estiver utilizado ou a planilha não seguir B_XX,
    # tenta utilizar o nome completo da planilha.
    nome_aba = _registrar_nome_aba(
        nome_planilha,
        nomes_utilizados
    )

    if nome_aba is not None:
        return nome_aba

    # Se o nome da planilha também estiver repetido, combina o nome da
    # planilha com o nome do arquivo de origem.
    nome_arquivo_sem_extensao = os.path.splitext(
        os.path.basename(nome_arquivo)
    )[0]

    nome_completo = (
        f"{nome_planilha}_{nome_arquivo_sem_extensao}"
    )

    nome_aba = _registrar_nome_aba(
        nome_completo,
        nomes_utilizados
    )

    if nome_aba is not None:
        return nome_aba

    # Caso extremo: mesmo nome de arquivo e mesma planilha.
    contador = 2

    while True:
        sufixo = f"_{contador}"
        limite_nome = 31 - len(sufixo)

        nome_com_sufixo = (
            f"{nome_completo[:limite_nome]}{sufixo}"
        )

        nome_aba = _registrar_nome_aba(
            nome_com_sufixo,
            nomes_utilizados
        )

        if nome_aba is not None:
            return nome_aba

        contador += 1


def processar(lista_arquivos):
    """
    Processa os arquivos Excel do sistema SAP.

    Regras:
    - processa arquivos XLS e XLSX;
    - ignora arquivos de outras extensões;
    - processa todas as planilhas válidas de cada arquivo;
    - planilhas iniciadas por B_XX geram inicialmente uma aba XX;
    - planilhas com o mesmo mês não substituem resultados anteriores;
    - planilhas fora do padrão usam o próprio nome;
    - cada planilha válida gera uma aba na pasta de trabalho final.
    """
    resultados = {}
    nomes_utilizados = set()
    arquivos_sem_resultados = []

    if not lista_arquivos:
        raise ValueError(
            "Nenhum arquivo foi selecionado para o sistema SAP."
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
            "sistema SAP. Selecione arquivos XLS ou XLSX."
        )

    for arquivo in arquivos_excel:
        nome_arquivo = os.path.basename(arquivo)

        resultados_arquivo = transformar_pasta_trabalho_sap(
            arquivo
        )

        if not resultados_arquivo:
            arquivos_sem_resultados.append(
                nome_arquivo
            )
            continue

        for nome_planilha, dataframe in resultados_arquivo.items():
            if dataframe is None or dataframe.empty:
                continue

            nome_aba = _gerar_nome_aba_unico(
                nome_planilha=nome_planilha,
                nome_arquivo=nome_arquivo,
                nomes_utilizados=nomes_utilizados
            )

            resultados[nome_aba] = dataframe

    if not resultados:
        detalhe = ""

        if arquivos_sem_resultados:
            detalhe = (
                " Arquivos sem planilhas válidas: "
                + ", ".join(arquivos_sem_resultados)
                + "."
            )

        raise ValueError(
            "Nenhuma planilha com classificações numéricas válidas foi "
            f"encontrada para o sistema SAP.{detalhe}"
        )

    return resultados

# ==============================================================================
# TRANSFORMAÇÃO DO BALANCETE SAP
# ==============================================================================

def _normalizar_texto_sap(valor):
    """
    Normaliza textos extraídos das planilhas do sistema SAP.

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


def _normalizar_conta_sap(valor):
    """
    Identifica a classificação contábil no final da célula.

    Todo texto, prefixo ou separador existente antes da classificação
    será desconsiderado.

    A função procura o último bloco formado exclusivamente por números.

    Exemplos:
        PC01/11110001       -> 11110001
        CONTA/11110003      -> 11110003
        ABC-11110004        -> 11110004
        EMPRESA 11110005    -> 11110005
        XPTO: 11110006      -> 11110006
        11110007            -> 11110007

    Retorna uma string vazia quando nenhum número for encontrado.
    """
    texto = _normalizar_texto_sap(valor)

    if not texto:
        return ""

    # Procura o último bloco numérico existente no final da célula.
    #
    # O uso do final da string evita capturar números que pertençam
    # ao prefixo, como o "01" de PC01.
    correspondencia = re.search(
        r"(?P<conta>\d+)\s*$",
        texto
    )

    if correspondencia is None:
        return ""

    conta = correspondencia.group(
        "conta"
    ).strip()

    return conta


def _converter_numero_sap(valor):
    """
    Converte valores monetários do sistema SAP para float.

    Exemplos:
        9.321,13 BRL    -> 9321.13
        -136,60 BRL     -> -136.60
        24.116,85 BRL   -> 24116.85
        0,00 BRL        -> 0.0
        -               -> 0.0

    Todos os resultados são arredondados para duas casas decimais.
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
        .replace("BRL", "")
        .replace("brl", "")
        .replace("Brl", "")
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


def _obter_engine_sap(caminho_arquivo):
    """
    Define o mecanismo de leitura do arquivo Excel.
    """
    extensao = os.path.splitext(
        caminho_arquivo
    )[1].lower()

    if extensao == ".xls":
        return "xlrd"

    return "openpyxl"


def _linha_tem_conta_sap(linha):
    """
    Verifica se uma linha possui uma classificação numérica válida
    na coluna C.
    """
    if len(linha) < 3:
        return False

    conta = _normalizar_conta_sap(
        linha.iloc[2]
    )

    return bool(conta)


def _localizar_primeira_conta_sap(df_origem):
    """
    Localiza a primeira linha que possui uma conta iniciada por PC01
    na coluna C.

    Dessa forma, todas as linhas anteriores são consideradas cabeçalho
    e são automaticamente desconsideradas.
    """
    for indice, linha in df_origem.iterrows():
        if _linha_tem_conta_sap(linha):
            return indice

    return None


def _linha_eh_cabecalho_sap(linha):
    """
    Identifica linhas do cabeçalho repetido.

    Essa verificação é complementar. A principal validação continua
    sendo a existência de uma conta PC01 na coluna C.
    """
    valores = [
        _normalizar_texto_sap(valor).upper()
        for valor in linha.tolist()
        if _normalizar_texto_sap(valor)
    ]

    if not valores:
        return False

    texto_linha = " | ".join(valores)

    termos_cabecalho = [
        "DT.LANÇAMENTO DE",
        "DT.LANCAMENTO DE",
        "LEDGER FONTE",
        "DATA LÇTO.ATÉ",
        "DATA LCTO.ATE",
        "INDICADORES",
        "CONTA DO RAZÃO",
        "CONTA DO RAZAO",
        "SALDO INICIAL EM MOEDA DA EMPRESA",
        "SALDO DEVEDOR EM MOEDA DA EMPRESA",
        "SALDO CREDOR EM MOEDA DA EMPRESA",
        "SALDO FINAL NA MOEDA DA EMPRESA",
    ]

    return any(
        termo in texto_linha
        for termo in termos_cabecalho
    )


def _transformar_planilha_sap(
    df_origem,
    nome_arquivo,
    nome_planilha
):
    """
    Transforma uma única planilha do sistema SAP.

    Layout da origem:
        A = Empresa
        B = Empresa
        C = Conta do Razão
        D = Descrição da conta
        E = Saldo inicial
        F = Saldo devedor
        G = Saldo credor
        H = Saldo final

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
    if df_origem is None or df_origem.empty:
        return None

    if df_origem.shape[1] < 8:
        return None

    indice_inicio = _localizar_primeira_conta_sap(
        df_origem
    )

    if indice_inicio is None:
        return None

    df_dados = df_origem.loc[
        indice_inicio:
    ].copy()

    df_dados.reset_index(
        drop=True,
        inplace=True
    )

    registros_validos = []

    for _, linha in df_dados.iterrows():
        # Ignora cabeçalhos que eventualmente se repitam entre
        # blocos ou páginas da mesma planilha.
        if _linha_eh_cabecalho_sap(linha):
            continue

        conta = _normalizar_conta_sap(
            linha.iloc[2]
        )

        # Somente linhas cuja coluna C termine com uma classificação
        # numérica válida são processadas.
        if not conta:
            continue

        descricao = _normalizar_texto_sap(
            linha.iloc[3]
        )

        if not descricao:
            continue

        saldo_anterior = _converter_numero_sap(
            linha.iloc[4]
        )

        debito = _converter_numero_sap(
            linha.iloc[5]
        )

        credito = _converter_numero_sap(
            linha.iloc[6]
        )

        saldo_acumulado = _converter_numero_sap(
            linha.iloc[7]
        )

        movimento = round(
            debito + credito,
            2
        )

        registros_validos.append({
            "Conta": conta,
            "Nome": descricao,
            "Cód. Reduzido": conta,
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
            "Movimento": movimento,
            "Saldo Acumulado": round(
                saldo_acumulado,
                2
            ),
        })

    if not registros_validos:
        return None

    df_registros = pd.DataFrame(
        registros_validos
    )

    df_destino = pd.DataFrame(
        index=df_registros.index
    )

    # Coluna A
    df_destino["Atividade"] = "Geral"

    # Coluna B
    df_destino["Conta"] = (
        df_registros["Conta"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    # Coluna C
    df_destino["Nome"] = (
        df_registros["Nome"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    # Coluna D
    df_destino["Cód. Reduzido"] = (
        df_registros["Cód. Reduzido"]
        .fillna("")
        .astype(str)
        .str.strip()
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
                df_registros[coluna],
                errors="coerce"
            )
            .fillna(0.0)
            .round(2)
        )

    # Recalcula o movimento depois do arredondamento.
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


def transformar_pasta_trabalho_sap(caminho_arquivo):
    """
    Processa todas as planilhas de uma pasta de trabalho do SAP.

    Cada planilha que contiver pelo menos uma conta válida iniciada
    por PC01 será transformada em um DataFrame.

    Retorno:
        {
            "Nome da planilha": dataframe
        }
    """
    nome_arquivo = os.path.basename(
        caminho_arquivo
    )

    extensao = os.path.splitext(
        caminho_arquivo
    )[1].lower()

    if extensao not in {".xls", ".xlsx"}:
        raise ValueError(
            f"O arquivo '{nome_arquivo}' não é um Excel válido."
        )

    engine = _obter_engine_sap(
        caminho_arquivo
    )

    try:
        pasta_trabalho = pd.ExcelFile(
            caminho_arquivo,
            engine=engine
        )

    except Exception as erro:
        raise ValueError(
            f"Não foi possível abrir a pasta de trabalho "
            f"'{nome_arquivo}'. Erro: {erro}"
        ) from erro

    if not pasta_trabalho.sheet_names:
        raise ValueError(
            f"A pasta de trabalho '{nome_arquivo}' não possui planilhas."
        )

    resultados = {}

    for nome_planilha in pasta_trabalho.sheet_names:
        try:
            df_origem = pd.read_excel(
                pasta_trabalho,
                sheet_name=nome_planilha,
                header=None,
                dtype=object
            )

        except Exception as erro:
            print(
                f"[SAP] Não foi possível ler a planilha "
                f"'{nome_planilha}' do arquivo '{nome_arquivo}'. "
                f"Erro: {erro}"
            )
            continue

        dataframe = _transformar_planilha_sap(
            df_origem=df_origem,
            nome_arquivo=nome_arquivo,
            nome_planilha=nome_planilha
        )

        if dataframe is None or dataframe.empty:
            print(
                f"[SAP] Planilha '{nome_planilha}' ignorada no "
                f"arquivo '{nome_arquivo}': nenhuma conta PC01 válida."
            )
            continue

        resultados[nome_planilha] = dataframe

    if not resultados:
        raise ValueError(
            f"Nenhuma planilha válida foi encontrada no arquivo "
            f"SAP '{nome_arquivo}'. Verifique se as contas estão "
            f"na coluna C e terminam com uma classificação numérica."
)

    return resultados