import os
import re
import unicodedata

import pandas as pd


# ==============================================================================
# CONFIGURAÇÕES DO SISTEMA Cotrijuc
# ==============================================================================

EXTENSOES_VALIDAS_cotrijuc = {
    ".xls",
    ".xlsx",
}

NOME_ABA_ORIGEM_cotrijuc = "2-PLANO DE CONTAS"

COLUNAS_DESTINO_cotrijuc = [
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
# NOMES DAS ABAS DE DESTINO
# ==============================================================================

def _limpar_nome_aba_cotrijuc(nome):
    """
    Ajusta um texto para utilização como nome de aba do Excel.

    Regras:
    - substitui caracteres inválidos por sublinhado;
    - impede nomes vazios;
    - limita o nome a 31 caracteres.
    """
    nome_limpo = re.sub(
        r'[\\/\x2a?:\[\]]',
        "_",
        str(nome).strip()
    )

    if not nome_limpo:
        nome_limpo = "Sem nome"

    return nome_limpo[:31]


def _obter_nome_sem_extensao_cotrijuc(caminho_arquivo):
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


def _obter_mes_cotrijuc(caminho_arquivo):
    """
    Identifica o mês quando o arquivo começa com B_XX.

    Exemplos:
        B_01_BALANCETE.xlsx -> 01
        B_05.2026.xls       -> 05
        BALANCETE.xlsx      -> None

    Somente meses entre 01 e 12 são aceitos.
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


def _registrar_nome_aba_cotrijuc(
    nome,
    nomes_utilizados
):
    """
    Registra um nome de aba caso ainda não esteja em uso.

    A comparação não diferencia letras maiúsculas e minúsculas.
    """
    nome_aba = _limpar_nome_aba_cotrijuc(
        nome
    )

    chave_nome = nome_aba.casefold()

    if chave_nome in nomes_utilizados:
        return None

    nomes_utilizados.add(
        chave_nome
    )

    return nome_aba


def _gerar_nome_aba_cotrijuc(
    caminho_arquivo,
    nomes_utilizados
):
    """
    Gera um nome exclusivo para a aba de destino.

    Regras:
    1. Arquivos B_XX tentam utilizar XX;
    2. Se XX já estiver em uso, utiliza o nome completo do arquivo;
    3. Arquivos fora do padrão utilizam o nome sem extensão;
    4. Nomes repetidos recebem um sufixo numérico.
    """
    nome_completo = _obter_nome_sem_extensao_cotrijuc(
        caminho_arquivo
    )

    mes = _obter_mes_cotrijuc(
        caminho_arquivo
    )

    if mes:
        nome_aba = _registrar_nome_aba_cotrijuc(
            mes,
            nomes_utilizados
        )

        if nome_aba is not None:
            return nome_aba

    nome_aba = _registrar_nome_aba_cotrijuc(
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

        nome_aba = _registrar_nome_aba_cotrijuc(
            nome_candidato,
            nomes_utilizados
        )

        if nome_aba is not None:
            return nome_aba

        contador += 1


# ==============================================================================
# NORMALIZAÇÃO DOS TEXTOS
# ==============================================================================

def _normalizar_texto_cotrijuc(valor):
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


def _normalizar_texto_comparacao_cotrijuc(valor):
    """
    Normaliza um texto para comparação.

    A função:
    - converte para letras maiúsculas;
    - remove acentos;
    - normaliza os espaços.
    """
    texto = _normalizar_texto_cotrijuc(
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


def _normalizar_nome_aba_origem_cotrijuc(nome_aba):
    """
    Normaliza o nome de uma aba da pasta de trabalho.

    Exemplos reconhecidos como equivalentes:
        2-Plano de contas
        2-PLANO DE CONTAS
        2-plano de contas
        2-Plano de Contas
        2-Plano de contas
    """
    texto = _normalizar_texto_comparacao_cotrijuc(
        nome_aba
    )

    texto = re.sub(
        r"\s{0,}-\s{0,}",
        "-",
        texto
    )

    return texto.strip()


# ==============================================================================
# LOCALIZAÇÃO DA ABA 2-PLANO DE CONTAS
# ==============================================================================

def _localizar_aba_plano_contas_cotrijuc(pasta_trabalho):
    """
    Localiza exclusivamente a aba 2-Plano de contas.

    A comparação ignora:
    - maiúsculas e minúsculas;
    - acentos;
    - espaços excedentes próximos ao hífen;
    - espaços no início ou no final.

    A primeira aba, 1-Parametros, será ignorada.
    """
    nome_procurado = _normalizar_nome_aba_origem_cotrijuc(
        NOME_ABA_ORIGEM_cotrijuc
    )

    for nome_aba in pasta_trabalho.sheet_names:
        nome_normalizado = (
            _normalizar_nome_aba_origem_cotrijuc(
                nome_aba
            )
        )

        if nome_normalizado == nome_procurado:
            return nome_aba

    nomes_disponiveis = ", ".join(
        str(nome_aba)
        for nome_aba in pasta_trabalho.sheet_names
    )

    raise ValueError(
        "Não foi encontrada a aba '2-Plano de contas'. "
        f"Abas disponíveis: {nomes_disponiveis}."
    )


# ==============================================================================
# NORMALIZAÇÃO DA CLASSIFICAÇÃO
# ==============================================================================

def _normalizar_classificacao_cotrijuc(valor):
    """
    Remove os pontos existentes na classificação.

    Exemplos:
        1          -> 1
        1.1        -> 11
        1.01       -> 101
        1.1.01     -> 1101
        1.01.02.03 -> 1010203

    A classificação será mantida como texto.
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
        texto = _normalizar_texto_cotrijuc(
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


def _classificacao_valida_cotrijuc(valor):
    """
    Verifica se o conteúdo representa uma classificação válida.
    """
    classificacao = _normalizar_classificacao_cotrijuc(
        valor
    )

    return bool(
        classificacao
        and classificacao.isdigit()
    )


# ==============================================================================
# CONVERSÃO DOS VALORES MONETÁRIOS
# ==============================================================================

def _extrair_natureza_cotrijuc(valor):
    """
    Extrai a natureza contábil incorporada ao final do valor.

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


def _converter_numero_cotrijuc(valor):
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
    O resultado será arredondado para duas casas decimais.
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


def _aplicar_natureza_cotrijuc(valor):
    """
    Converte o valor e aplica sua natureza contábil.

    Regras:
        D = positivo
        C = negativo

    Quando não houver D/C, mantém o sinal original.

    Exemplos:
        1.000,00 D  -> 1000.00
        1.000,00 C  -> -1000.00
        -1.000,00 D -> 1000.00
        -1.000,00 C -> -1000.00
        -1.000,00   -> -1000.00
    """
    numero_original = _converter_numero_cotrijuc(
        valor
    )

    natureza = _extrair_natureza_cotrijuc(
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

def _obter_engine_excel_cotrijuc(caminho_arquivo):
    """
    Define o mecanismo utilizado para ler o arquivo Excel.

    XLS utiliza xlrd.
    XLSX utiliza openpyxl.
    """
    extensao = os.path.splitext(
        caminho_arquivo
    )[1].lower()

    if extensao == ".xls":
        return "xlrd"

    if extensao == ".xlsx":
        return "openpyxl"

    raise ValueError(
        f"A extensão '{extensao}' não é compatível com o balancte Cotrijuc."
    )


def _ler_aba_plano_contas_cotrijuc(caminho_arquivo):
    """
    Abre a pasta de trabalho e lê exclusivamente a aba
    2-Plano de contas.

    A aba 1-Parametros será ignorada.
    """
    nome_arquivo = os.path.basename(
        caminho_arquivo
    )

    engine = _obter_engine_excel_cotrijuc(
        caminho_arquivo
    )

    try:
        pasta_trabalho = pd.ExcelFile(
            caminho_arquivo,
            engine=engine
        )
    except Exception as erro:
        raise ValueError(
            f"Não foi possível abrir o arquivo Cotrijuc "
            f"'{nome_arquivo}'. Erro: {erro}"
        ) from erro

    if len(pasta_trabalho.sheet_names) == 0:
        raise ValueError(
            f"O arquivo Cotrijuc '{nome_arquivo}' não possui abas."
        )

    nome_aba_plano = _localizar_aba_plano_contas_cotrijuc(
        pasta_trabalho
    )

    try:
        df_origem = pd.read_excel(
            pasta_trabalho,
            sheet_name=nome_aba_plano,
            header=None,
            dtype=object
        )
    except Exception as erro:
        raise ValueError(
            f"Não foi possível ler a aba '{nome_aba_plano}' "
            f"do arquivo Cotrijuc '{nome_arquivo}'. Erro: {erro}"
        ) from erro

    if df_origem.empty:
        raise ValueError(
            f"A aba '{nome_aba_plano}' do arquivo "
            f"'{nome_arquivo}' está vazia."
        )

    if df_origem.shape[1] < 7:
        raise ValueError(
            f"A aba '{nome_aba_plano}' possui "
            f"{df_origem.shape[1]} coluna(s), mas são necessárias "
            "pelo menos 7 colunas, de A até G."
        )

    return df_origem


# ==============================================================================
# IDENTIFICAÇÃO DO CABEÇALHO
# ==============================================================================

def _linha_eh_cabecalho_cotrijuc(linha):
    """
    Verifica se uma linha corresponde ao cabeçalho do plano de contas.

    Títulos esperados:
        Conta
        Descrição
        Saldo anterior
        Débito
        Crédito
        Movimento do período
        Saldo atual
    """
    textos_linha = [
        _normalizar_texto_comparacao_cotrijuc(
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


def _linha_possui_dados_cotrijuc(linha):
    """
    Verifica se uma linha representa uma conta contábil válida.

    Layout esperado:
        A = Conta
        B = Descrição
        C = Saldo anterior
        D = Débito
        E = Crédito
        F = Movimento do período
        G = Saldo atual
    """
    if len(linha) < 7:
        return False

    classificacao = _normalizar_classificacao_cotrijuc(
        linha.iloc[0]
    )

    descricao = _normalizar_texto_cotrijuc(
        linha.iloc[1]
    )

    if not classificacao:
        return False

    if not descricao:
        return False

    return True


def _localizar_inicio_dados_cotrijuc(
    df_origem,
    nome_arquivo
):
    """
    Localiza o início dos dados na aba 2-Plano de contas.

    Primeira estratégia:
        localiza o cabeçalho e inicia na linha seguinte.

    Segunda estratégia:
        localiza diretamente a primeira conta válida.
    """
    for indice_linha in range(
        df_origem.shape[0]
    ):
        linha = df_origem.iloc[
            indice_linha
        ]

        if _linha_eh_cabecalho_cotrijuc(
            linha
        ):
            return indice_linha + 1

    for indice_linha in range(
        df_origem.shape[0]
    ):
        linha = df_origem.iloc[
            indice_linha
        ]

        if _linha_possui_dados_cotrijuc(
            linha
        ):
            return indice_linha

    raise ValueError(
        f"Não foi possível localizar o cabeçalho ou a primeira "
        f"conta válida na aba '2-Plano de contas' do arquivo "
        f"Cotrijuc '{nome_arquivo}'."
    )


# ==============================================================================
# EXTRAÇÃO DOS REGISTROS
# ==============================================================================

def _extrair_registro_cotrijuc(linha):
    """
    Extrai uma conta da aba 2-Plano de contas.

    Origem:
        A = Conta
        B = Descrição
        C = Saldo anterior
        D = Débito
        E = Crédito
        F = Movimento do período
        G = Saldo atual
    """
    if not _linha_possui_dados_cotrijuc(
        linha
    ):
        return None

    classificacao = _normalizar_classificacao_cotrijuc(
        linha.iloc[0]
    )

    descricao = _normalizar_texto_cotrijuc(
        linha.iloc[1]
    )

    saldo_anterior = _aplicar_natureza_cotrijuc(
        linha.iloc[2]
    )

    debito = _aplicar_natureza_cotrijuc(
        linha.iloc[3]
    )

    credito = _aplicar_natureza_cotrijuc(
        linha.iloc[4]
    )

    movimento = _aplicar_natureza_cotrijuc(
        linha.iloc[5]
    )

    saldo_acumulado = _aplicar_natureza_cotrijuc(
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
# MONTAGEM DO DATAFRAME
# ==============================================================================

def _montar_dataframe_cotrijuc(
    registros_validos
):
    """
    Monta o DataFrame final do Cotrijuc.
    """
    if not registros_validos:
        return pd.DataFrame(
            columns=COLUNAS_DESTINO_cotrijuc
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
        COLUNAS_DESTINO_cotrijuc
    ].copy()


# ==============================================================================
# TRANSFORMAÇÃO DO BALANCETE Cotrijuc
# ==============================================================================

def transformar_balancete_cotrijuc(caminho_arquivo):
    """
    Transforma exclusivamente a aba 2-Plano de contas.

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
        - ignora completamente a aba 1-Parametros;
        - processa exclusivamente a aba 2-Plano de contas;
        - ignora as informações anteriores ao cabeçalho;
        - remove os pontos da classificação;
        - aplica os indicadores D/C existentes nos valores;
        - mantém os saldos exatamente conforme cada arquivo;
        - não ajusta ou acumula saldos de contas de resultado;
        - arredonda os valores para duas casas decimais.
    """
    nome_arquivo = os.path.basename(
        caminho_arquivo
    )

    extensao = os.path.splitext(
        caminho_arquivo
    )[1].lower()

    if extensao not in EXTENSOES_VALIDAS_cotrijuc:
        raise ValueError(
            f"O arquivo '{nome_arquivo}' não é um arquivo Excel válido."
        )

    df_origem = _ler_aba_plano_contas_cotrijuc(
        caminho_arquivo
    )

    indice_inicio = _localizar_inicio_dados_cotrijuc(
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
        if _linha_eh_cabecalho_cotrijuc(
            linha
        ):
            continue

        registro = _extrair_registro_cotrijuc(
            linha
        )

        if registro is not None:
            registros_validos.append(
                registro
            )

            continue

        if _classificacao_valida_cotrijuc(
            linha.iloc[0]
        ):
            linhas_ignoradas.append({
                "Linha": (
                    indice_inicio
                    + indice_relativo
                    + 1
                ),
                "Conta": (
                    _normalizar_texto_cotrijuc(
                        linha.iloc[0]
                    )
                ),
                "Descrição": (
                    _normalizar_texto_cotrijuc(
                        linha.iloc[1]
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
                        f"Conta={item['Conta']}; "
                        f"Descrição={item['Descrição']}"
                    )
                    for item in exemplos
                )
            )

        raise ValueError(
            f"Nenhuma conta válida foi encontrada na aba "
            f"'2-Plano de contas' do arquivo Cotrijuc "
            f"'{nome_arquivo}'. Verifique se a conta está na coluna A, "
            "a descrição na coluna B e os valores nas colunas C até G."
            + detalhe
        )

    if linhas_ignoradas:
        print(
            f"[Cotrijuc] "
            f"{len(linhas_ignoradas)} linha(s) foram ignoradas "
            f"na aba '2-Plano de contas' do arquivo "
            f"'{nome_arquivo}'."
        )

        for item in linhas_ignoradas:
            print(
                f"[Cotrijuc] Linha {item['Linha']}: "
                f"Conta={item['Conta']}; "
                f"Descrição={item['Descrição']}"
            )

    df_destino = _montar_dataframe_cotrijuc(
        registros_validos
    )

    if df_destino.empty:
        raise ValueError(
            f"O arquivo Cotrijuc '{nome_arquivo}' não gerou "
            "nenhum registro no layout de destino."
        )

    return df_destino


# ==============================================================================
# FUNÇÃO PÚBLICA DO Cotrijuc
# ==============================================================================

def processar(lista_arquivos):
    """
    Processa os arquivos selecionados para o sistema Cotrijuc.

    Regras:
        - processa somente XLS e XLSX;
        - ignora arquivos de outras extensões;
        - processa somente a aba 2-Plano de contas;
        - cada arquivo gera uma aba no resultado;
        - arquivos B_XX tentam utilizar XX;
        - meses repetidos utilizam o nome completo do arquivo;
        - não aplica ajustes sequenciais às contas de resultado.
    """
    resultados = {}
    nomes_utilizados = set()

    if not lista_arquivos:
        raise ValueError(
            "Nenhum arquivo foi selecionado para o sistema Cotrijuc."
        )

    arquivos_excel = [
        arquivo
        for arquivo in lista_arquivos
        if os.path.splitext(
            str(arquivo)
        )[1].lower() in EXTENSOES_VALIDAS_cotrijuc
    ]

    if not arquivos_excel:
        raise ValueError(
            "Nenhum arquivo Excel válido foi encontrado para o "
            "sistema Cotrijuc. Selecione arquivos XLS ou XLSX."
        )

    for arquivo in arquivos_excel:
        nome_aba = _gerar_nome_aba_cotrijuc(
            caminho_arquivo=arquivo,
            nomes_utilizados=nomes_utilizados
        )

        dataframe = transformar_balancete_cotrijuc(
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
            "Nenhum resultado foi gerado para o sistema Cotrijuc."
        )

    return resultados