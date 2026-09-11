import os
import re
import pandas as pd
import re

def _limpar_nome_aba(nome):
    """
    Ajusta um texto para utilização como nome de aba do Excel.

    O Excel:
    - não permite os caracteres: \\ / ? * [ ] :
    - permite no máximo 31 caracteres.
    """
    nome_limpo = re.sub(
        r'[\\/*?:\[\]]',
        "_",
        str(nome).strip()
    )

    if not nome_limpo:
        nome_limpo = "Sem nome"

    return nome_limpo[:31]


def _obter_nome_completo(caminho_arquivo):
    """Retorna o nome completo do arquivo sem a extensão."""
    nome_arquivo = os.path.basename(caminho_arquivo)
    return os.path.splitext(nome_arquivo)[0].strip()


def _obter_mes(caminho_arquivo):
    """
    Extrai o mês quando o arquivo começa com B_XX.

    Exemplos:
        B_01.2026.xlsx  -> 01
        B_02_FILIAL.xls -> 02
        Balancete.xlsx  -> None
    """
    nome_arquivo = os.path.basename(caminho_arquivo)

    correspondencia = re.match(
        r"^B_(\d{2})",
        nome_arquivo,
        re.IGNORECASE
    )

    if correspondencia:
        return correspondencia.group(1)

    return None


def _registrar_nome_aba(nome, nomes_utilizados):
    """
    Registra um nome de aba caso ainda não esteja em uso.

    A comparação não diferencia letras maiúsculas e minúsculas.
    """
    nome_aba = _limpar_nome_aba(nome)
    nome_comparacao = nome_aba.casefold()

    if nome_comparacao in nomes_utilizados:
        return None

    nomes_utilizados.add(nome_comparacao)
    return nome_aba


def _gerar_nome_aba(caminho_arquivo, nomes_utilizados):
    """
    Gera um nome único para a aba.

    Regras:
    - arquivos B_XX tentam utilizar XX;
    - se o mês já estiver em uso, utiliza o nome completo;
    - arquivos fora do padrão usam o nome completo sem extensão;
    - duplicidades totais recebem um sufixo numérico.
    """
    nome_completo = _obter_nome_completo(caminho_arquivo)
    mes = _obter_mes(caminho_arquivo)

    if mes:
        nome_aba = _registrar_nome_aba(
            mes,
            nomes_utilizados
        )

        if nome_aba is not None:
            return nome_aba

    nome_aba = _registrar_nome_aba(
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

        nome_aba = _registrar_nome_aba(
            nome_com_sufixo,
            nomes_utilizados
        )

        if nome_aba is not None:
            return nome_aba

        contador += 1


def processar(lista_arquivos):
    """
    Processa os balancetes Excel do cliente Cooperlate.

    Regras:
    - processa somente arquivos .xls e .xlsx;
    - arquivos de outras extensões são ignorados;
    - cada arquivo gera uma aba;
    - arquivos B_XX tentam utilizar XX como nome da aba;
    - meses repetidos usam o nome completo do arquivo;
    - arquivos fora do padrão usam o nome sem extensão.
    """
    resultados = {}
    nomes_utilizados = set()

    if not lista_arquivos:
        raise ValueError(
            "Nenhum arquivo foi selecionado para o cliente Cooperlate."
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
            "cliente Cooperlate. Selecione arquivos .xls ou .xlsx."
        )

    for arquivo in arquivos_excel:
        nome_aba = _gerar_nome_aba(
            caminho_arquivo=arquivo,
            nomes_utilizados=nomes_utilizados
        )

        dataframe = transformar_balancete_cooperlate(arquivo)

        if dataframe is None or dataframe.empty:
            raise ValueError(
                f"O arquivo '{os.path.basename(arquivo)}' não retornou "
                "dados válidos para tabulação."
            )

        resultados[nome_aba] = dataframe

    if not resultados:
        raise ValueError(
            "Nenhum resultado foi gerado para o cliente Cooperlate."
        )

    return resultados
# ==============================================================================
# TRANSFORMAÇÃO DO BALANCETE COOPERLATE
# ==============================================================================

def _normalizar_texto_cooperlate(valor):
    """
    Normaliza textos utilizados no processamento do cliente Cooperlate.

    Remove espaços especiais, espaços repetidos e espaços nas
    extremidades.
    """
    if pd.isna(valor):
        return ""

    return re.sub(
        r"\s+",
        " ",
        str(valor).replace("\xa0", " ")
    ).strip()


def _converter_numero_cooperlate(valor):
    """
    Converte valores monetários do cliente Cooperlate para float.

    Formatos aceitos:
        145.430.616,65
        145430616,65
        145430616.65
        -145.430.616,65
        (145.430.616,65)
        R$ 145.430.616,65

    Valores vazios ou inválidos são convertidos para 0.0.
    """
    if pd.isna(valor):
        return 0.0

    if isinstance(valor, (int, float)):
        return float(valor)

    texto = str(valor).strip()

    if not texto:
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

    if negativo_por_parenteses:
        texto = texto[1:-1]

    if "," in texto:
        texto = (
            texto
            .replace(".", "")
            .replace(",", ".")
        )

    try:
        numero = float(texto)

        if negativo_por_parenteses:
            numero = -abs(numero)

        return numero

    except (ValueError, TypeError):
        return 0.0


def _aplicar_natureza_cooperlate(valor, natureza):
    """
    Aplica a natureza contábil ao valor.

    Regras:
        D = positivo
        C = negativo

    Naturezas vazias ou desconhecidas mantêm o sinal original.
    """
    numero = _converter_numero_cooperlate(valor)

    natureza_normalizada = (
        _normalizar_texto_cooperlate(natureza)
        .upper()
    )

    if natureza_normalizada == "D":
        return abs(numero)

    if natureza_normalizada == "C":
        return -abs(numero)

    return numero


def _separar_conta_descricao_cooperlate(valor):
    """
    Separa a classificação e a descrição existentes na coluna A.

    Exemplos:
        "1 ATIVO"
            -> classificação: "1"
            -> descrição: "ATIVO"

        "101 ATIVO CIRCULANTE"
            -> classificação: "101"
            -> descrição: "ATIVO CIRCULANTE"

        "10101 DISPONIVEL"
            -> classificação: "10101"
            -> descrição: "DISPONIVEL"

    A classificação deve estar no início da célula e ser seguida
    por pelo menos um espaço antes da descrição.
    """
    texto = _normalizar_texto_cooperlate(valor)

    if not texto:
        return "", ""

    correspondencia = re.match(
        r"^(?P<conta>\d+)\s+(?P<descricao>.+?)$",
        texto
    )

    if correspondencia is None:
        return "", ""

    conta = correspondencia.group("conta").strip()
    descricao = correspondencia.group("descricao").strip()

    return conta, descricao


def _localizar_cabecalho_cooperlate(df_origem, nome_arquivo):
    """
    Localiza a linha de cabeçalho do balancete.

    O cabeçalho deve apresentar:
        Coluna A = Conta
        Coluna B = Cta red

    As linhas anteriores, como a indicação do período, são ignoradas.
    """
    for indice in df_origem.index:
        valor_coluna_a = (
            _normalizar_texto_cooperlate(
                df_origem.iloc[indice, 0]
            )
            .upper()
        )

        valor_coluna_b = (
            _normalizar_texto_cooperlate(
                df_origem.iloc[indice, 1]
            )
            .upper()
        )

        coluna_a_valida = valor_coluna_a == "CONTA"

        coluna_b_valida = valor_coluna_b in {
            "CTA RED",
            "CTA. RED",
            "CTA RED.",
            "CTA. RED.",
        }

        if coluna_a_valida and coluna_b_valida:
            return indice

    raise ValueError(
        f"Não foi possível localizar o cabeçalho com as colunas "
        f"'Conta' e 'Cta red' no arquivo '{nome_arquivo}'."
    )


def transformar_balancete_cooperlate(caminho_arquivo):
    """
    Transforma o balancete Excel do cliente Cooperlate.

    O arquivo é lido sem cabeçalho para permitir a identificação
    dinâmica da linha que contém as colunas Conta e Cta red.

    Layout de origem:
        A = Classificação e descrição
        B = Cta red
        C = Saldo Anterior
        D = Natureza do Saldo Anterior
        E = Débitos Mês
        F = Natureza dos Débitos
        G = Créditos Mês
        H = Natureza dos Créditos
        I = Saldo Mês
        J = Natureza do Saldo Mês
        K = Saldo Atual
        L = Natureza do Saldo Atual

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
    nome_arquivo = os.path.basename(caminho_arquivo)
    extensao = os.path.splitext(caminho_arquivo)[1].lower()

    if extensao not in {".xls", ".xlsx"}:
        raise ValueError(
            f"O arquivo '{nome_arquivo}' não é um arquivo Excel válido."
        )

    engine = "xlrd" if extensao == ".xls" else "openpyxl"

    try:
        df_origem = pd.read_excel(
            caminho_arquivo,
            sheet_name=0,
            header=None,
            dtype=object,
            engine=engine
        )

    except Exception as erro:
        raise ValueError(
            f"Não foi possível ler o arquivo do cliente Cooperlate "
            f"'{nome_arquivo}'. Erro: {erro}"
        ) from erro

    if df_origem.shape[1] < 12:
        raise ValueError(
            f"O arquivo do cliente Cooperlate '{nome_arquivo}' possui "
            f"{df_origem.shape[1]} coluna(s), mas são necessárias "
            "pelo menos 12 colunas, de A até L."
        )

    indice_cabecalho = _localizar_cabecalho_cooperlate(
        df_origem,
        nome_arquivo
    )

    # Mantém somente as linhas posteriores ao cabeçalho.
    df_origem = df_origem.iloc[
        indice_cabecalho + 1:
    ].copy()

    # Remove linhas totalmente vazias, inclusive as linhas em branco
    # existentes entre as contas.
    df_origem.dropna(
        how="all",
        inplace=True
    )

    df_origem.reset_index(
        drop=True,
        inplace=True
    )

    registros_validos = []

    for _, linha in df_origem.iterrows():
        conta, descricao = _separar_conta_descricao_cooperlate(
            linha.iloc[0]
        )

        # Linhas sem classificação e descrição válidas são ignoradas.
        # Isso também evita processar cabeçalhos repetidos, totais sem
        # conta e outras informações complementares.
        if not conta or not descricao:
            continue

        codigo_reduzido = _normalizar_texto_cooperlate(
            linha.iloc[1]
        )

        registros_validos.append({
            "Conta": conta,
            "Nome": descricao,
            "Cód. Reduzido": codigo_reduzido,
            "Saldo Anterior": _aplicar_natureza_cooperlate(
                linha.iloc[2],
                linha.iloc[3]
            ),
            "Débito": abs(
                _converter_numero_cooperlate(
                    linha.iloc[4]
                )
            ),
            "Crédito": abs(
                _converter_numero_cooperlate(
                    linha.iloc[6]
                )
            ),
            "Saldo Acumulado": _aplicar_natureza_cooperlate(
                linha.iloc[10],
                linha.iloc[11]
            ),
        })

    if not registros_validos:
        raise ValueError(
            f"Nenhuma conta válida foi encontrada no arquivo "
            f"do cliente Cooperlate '{nome_arquivo}'."
        )

    df_registros = pd.DataFrame(
        registros_validos
    )

    df_destino = pd.DataFrame(
        index=df_registros.index
    )

    # Coluna A: valor fixo.
    df_destino["Atividade"] = "Geral"

    # Coluna B: classificação extraída da origem A.
    df_destino["Conta"] = df_registros["Conta"]

    # Coluna C: descrição extraída da origem A.
    df_destino["Nome"] = df_registros["Nome"]

    # Coluna D: origem B.
    df_destino["Cód. Reduzido"] = (
        df_registros["Cód. Reduzido"]
    )

    # Coluna E: origem C com natureza da origem D.
    df_destino["Saldo Anterior"] = (
        df_registros["Saldo Anterior"]
        .astype(float)
    )

    # Coluna F: origem E.
    df_destino["Débito"] = (
        df_registros["Débito"]
        .astype(float)
    )

    # Coluna G: origem G.
    df_destino["Crédito"] = (
        df_registros["Crédito"]
        .astype(float)
    )

    # Coluna H: Débito menos Crédito.
    df_destino["Movimento"] = (
        df_destino["Débito"]
        - df_destino["Crédito"]
    )

    # Coluna I: origem K com natureza da origem L.
    df_destino["Saldo Acumulado"] = (
        df_registros["Saldo Acumulado"]
        .astype(float)
    )
    
    return df_destino