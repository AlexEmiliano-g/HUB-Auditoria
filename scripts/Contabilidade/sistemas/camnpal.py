import csv
import os
import re
import unicodedata

import pandas as pd


# ==============================================================================
# NOMES DAS ABAS
# ==============================================================================

def _limpar_nome_aba_camnpal(nome):
    """
    Ajusta o texto para utilização como nome de aba do Excel.

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


def _obter_nome_sem_extensao_camnpal(caminho_arquivo):
    """
    Retorna o nome do arquivo sem a extensão.
    """
    nome_arquivo = os.path.basename(caminho_arquivo)
    nome_sem_extensao = os.path.splitext(nome_arquivo)[0]

    return nome_sem_extensao.strip()


def _obter_mes_camnpal(caminho_arquivo):
    """
    Extrai o mês quando o nome do arquivo começa com B_XX.

    Exemplos:
        B_01_BALANCETE.csv -> 01
        B_05.2026.csv      -> 05
        BALANCETE.csv      -> None
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


def _registrar_nome_aba_camnpal(nome, nomes_utilizados):
    """
    Registra um nome de aba caso ainda não esteja em uso.

    A comparação não diferencia letras maiúsculas e minúsculas.
    """
    nome_aba = _limpar_nome_aba_camnpal(nome)
    chave_nome = nome_aba.casefold()

    if chave_nome in nomes_utilizados:
        return None

    nomes_utilizados.add(chave_nome)

    return nome_aba


def _gerar_nome_aba_camnpal(caminho_arquivo, nomes_utilizados):
    """
    Gera um nome exclusivo para a aba de destino.

    Regras:
    1. Arquivos B_XX tentam utilizar XX;
    2. Se XX já estiver em uso, utiliza o nome completo do arquivo;
    3. Arquivos fora do padrão utilizam o nome sem extensão;
    4. Duplicidades recebem um sufixo numérico.
    """
    nome_completo = _obter_nome_sem_extensao_camnpal(
        caminho_arquivo
    )

    mes = _obter_mes_camnpal(
        caminho_arquivo
    )

    if mes:
        nome_aba = _registrar_nome_aba_camnpal(
            mes,
            nomes_utilizados
        )

        if nome_aba is not None:
            return nome_aba

    nome_aba = _registrar_nome_aba_camnpal(
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

        nome_aba = _registrar_nome_aba_camnpal(
            nome_com_sufixo,
            nomes_utilizados
        )

        if nome_aba is not None:
            return nome_aba

        contador += 1


# ==============================================================================
# NORMALIZAÇÃO DE TEXTOS E CÓDIGOS
# ==============================================================================

def _normalizar_texto_camnpal(valor):
    """
    Normaliza textos extraídos do CSV.

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


def _normalizar_texto_comparacao_camnpal(valor):
    """
    Normaliza um texto para comparação.

    Além da limpeza de espaços:
    - converte para letras maiúsculas;
    - remove acentos.
    """
    texto = _normalizar_texto_camnpal(valor).upper()

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


def _normalizar_codigo_reduzido_camnpal(valor):
    """
    Normaliza o código reduzido, preservando-o como texto.

    Exemplos:
        1000   -> 1000
        1000.0 -> 1000
    """
    texto = _normalizar_texto_camnpal(valor)

    if not texto:
        return ""

    if re.fullmatch(r"\d+\.0+", texto):
        return texto.split(
            ".",
            maxsplit=1
        )[0]

    return texto


def _separar_conta_descricao_camnpal(valor):
    """
    Separa a classificação e a descrição contidas na coluna A.

    Exemplos:
        1 ATIVO
            Conta: 1
            Nome: ATIVO

        101 ATIVO CIRCULANTE
            Conta: 101
            Nome: ATIVO CIRCULANTE

        10101 DISPONIVEL
            Conta: 10101
            Nome: DISPONIVEL

        1.01 ATIVO CIRCULANTE
            Conta: 1.01
            Nome: ATIVO CIRCULANTE

        1 01 01 DISPONIVEL
            Conta: 1 01 01
            Nome: DISPONIVEL
    """
    texto = _normalizar_texto_camnpal(valor)

    if not texto:
        return "", ""

    correspondencia = re.match(
        r"^\s*"
        r"(?P<conta>\d+(?:(?:\s+|[.\-/])\d+)*)"
        r"\s+"
        r"(?P<descricao>.+?)"
        r"\s*$",
        texto
    )

    if correspondencia is None:
        return "", ""

    conta = re.sub(
        r"\s+",
        " ",
        correspondencia.group("conta")
    ).strip()

    descricao = re.sub(
        r"\s+",
        " ",
        correspondencia.group("descricao")
    ).strip()

    if not conta or not descricao:
        return "", ""

    return conta, descricao


# ==============================================================================
# CONVERSÃO DOS VALORES
# ==============================================================================

def _converter_numero_camnpal(valor):
    """
    Converte valores monetários brasileiros para float.

    Exemplos:
        1.009.340.704,47  -> 1009340704.47
        -676.505.705,37   -> -676505705.37
        0,00              -> 0.0
        -                 -> 0.0
        (1.250,50)        -> -1250.50

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


def _aplicar_natureza_camnpal(valor, natureza):
    """
    Aplica a natureza contábil ao valor.

    Regras:
        D = positivo
        C = negativo

    O sinal originalmente existente no valor será desconsiderado
    quando a natureza D ou C for informada.

    Exemplos:
        1.000,00 + D  -> 1000.00
        1.000,00 + C  -> -1000.00
        -1.000,00 + D -> 1000.00
        -1.000,00 + C -> -1000.00
    """
    numero = round(
        abs(
            _converter_numero_camnpal(
                valor
            )
        ),
        2
    )

    natureza_normalizada = (
        _normalizar_texto_comparacao_camnpal(
            natureza
        )
        .replace(".", "")
    )

    if natureza_normalizada == "D":
        return numero

    if natureza_normalizada == "C":
        return round(
            -numero,
            2
        )

    # Caso a natureza esteja ausente, preserva o sinal original.
    return round(
        _converter_numero_camnpal(
            valor
        ),
        2
    )


# ==============================================================================
# LEITURA DO CSV
# ==============================================================================

def _ler_csv_camnpal(caminho_arquivo):
    """
    Lê o CSV do cliente Camnpal.

    Características:
    - campos separados por ponto e vírgula;
    - pode possuir linhas iniciais com quantidades diferentes de campos;
    - pode possuir linhas vazias entre as contas;
    - pode conter ponto e vírgula excedente ao final das linhas.

    A leitura é feita com o módulo csv para evitar erros como:
        Expected X fields in line Y, saw Z
    """
    nome_arquivo = os.path.basename(
        caminho_arquivo
    )

    codificacoes = [
        "utf-8-sig",
        "cp1252",
        "latin-1",
    ]

    linhas_lidas = None
    ultimo_erro = None

    for codificacao in codificacoes:
        try:
            with open(
                caminho_arquivo,
                mode="r",
                encoding=codificacao,
                newline=""
            ) as arquivo:
                leitor = csv.reader(
                    arquivo,
                    delimiter=";",
                    quotechar='"'
                )

                linhas_lidas = [
                    list(linha)
                    for linha in leitor
                ]

            break

        except UnicodeDecodeError as erro:
            ultimo_erro = erro

        except (OSError, csv.Error) as erro:
            raise ValueError(
                f"Não foi possível ler o arquivo CSV "
                f"'{nome_arquivo}'. Erro: {erro}"
            ) from erro

    if linhas_lidas is None:
        raise ValueError(
            f"Não foi possível identificar a codificação do arquivo "
            f"CSV '{nome_arquivo}'. Erro: {ultimo_erro}"
        )

    if not linhas_lidas:
        raise ValueError(
            f"O arquivo CSV '{nome_arquivo}' está vazio."
        )

    # Remove campos vazios excedentes apenas no final das linhas.
    # Os campos vazios existentes entre as colunas são preservados.
    for linha in linhas_lidas:
        while linha and not str(linha[-1]).strip():
            linha.pop()

    maior_quantidade_colunas = max(
        len(linha)
        for linha in linhas_lidas
    )

    if maior_quantidade_colunas == 0:
        raise ValueError(
            f"O arquivo CSV '{nome_arquivo}' não possui conteúdo válido."
        )

    linhas_padronizadas = []

    for linha in linhas_lidas:
        quantidade_faltante = (
            maior_quantidade_colunas
            - len(linha)
        )

        linha_padronizada = linha + (
            [""] * quantidade_faltante
        )

        linhas_padronizadas.append(
            linha_padronizada
        )

    return pd.DataFrame(
        linhas_padronizadas,
        dtype=str
    )


# ==============================================================================
# IDENTIFICAÇÃO DO CABEÇALHO E DAS LINHAS
# ==============================================================================

def _linha_eh_cabecalho_camnpal(linha):
    """
    Verifica se a linha corresponde ao cabeçalho do balancete.

    O cabeçalho esperado contém:
        Conta
        Saldo Anterior
        Débitos Mês
        Créditos Mês
        Saldo Atual
    """
    valores = [
        _normalizar_texto_comparacao_camnpal(
            valor
        )
        for valor in linha.tolist()
    ]

    texto_completo = " | ".join(
        valor
        for valor in valores
        if valor
    )

    termos_cabecalho = [
        "CONTA",
        "SALDO ANTERIOR",
        "DEBITOS MES",
        "CREDITOS MES",
        "SALDO ATUAL",
    ]

    quantidade_encontrada = sum(
        termo in texto_completo
        for termo in termos_cabecalho
    )

    return quantidade_encontrada >= 3


def _localizar_cabecalho_camnpal(df_origem, nome_arquivo):
    """
    Localiza dinamicamente o cabeçalho do balancete.

    A primeira coluna do cabeçalho deve conter o termo Conta.
    """
    for indice, linha in df_origem.iterrows():
        primeira_coluna = (
            _normalizar_texto_comparacao_camnpal(
                linha.iloc[0]
            )
        )

        if primeira_coluna != "CONTA":
            continue

        if _linha_eh_cabecalho_camnpal(
            linha
        ):
            return indice

    raise ValueError(
        f"Não foi possível localizar o cabeçalho do balancete "
        f"no arquivo Camnpal '{nome_arquivo}'."
    )


def _linha_possui_dados_camnpal(linha):
    """
    Verifica se uma linha representa uma conta válida.

    A coluna A deve conter classificação e descrição.
    """
    conta, descricao = _separar_conta_descricao_camnpal(
        linha.iloc[0]
    )

    if not conta:
        return False

    if not descricao:
        return False

    return True


# ==============================================================================
# TRANSFORMAÇÃO DO BALANCETE CAMNPAL
# ==============================================================================

def transformar_balancete_camnpal(caminho_arquivo):
    """
    Transforma o balancete CSV do cliente Camnpal.

    Layout da origem:
        A = Classificação e descrição
        B = Código reduzido
        C = Saldo Anterior
        D = Natureza do Saldo Anterior
        E = Débitos do Mês
        F = Natureza dos Débitos
        G = Créditos do Mês
        H = Natureza dos Créditos
        I = Saldo do Mês
        J = Natureza do Saldo do Mês
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

    Mapeamento:
        Valor fixo Geral     -> Atividade
        Origem A, conta      -> Conta
        Origem A, descrição  -> Nome
        Origem B             -> Cód. Reduzido
        Origem C + D         -> Saldo Anterior
        Origem E + F         -> Débito
        Origem G + H         -> Crédito
        Débito + Crédito     -> Movimento
        Origem K + L         -> Saldo Acumulado

    Regras:
        - ignora informações anteriores ao cabeçalho;
        - ignora linhas vazias entre as contas;
        - separa classificação e descrição;
        - D gera valor positivo;
        - C gera valor negativo;
        - todos os valores são arredondados para duas casas decimais.
    """
    nome_arquivo = os.path.basename(
        caminho_arquivo
    )

    extensao = os.path.splitext(
        caminho_arquivo
    )[1].lower()

    if extensao != ".csv":
        raise ValueError(
            f"O arquivo '{nome_arquivo}' não possui extensão CSV."
        )

    df_origem = _ler_csv_camnpal(
        caminho_arquivo
    )

    if df_origem.empty:
        raise ValueError(
            f"O arquivo CSV '{nome_arquivo}' está vazio."
        )

    if df_origem.shape[1] < 12:
        raise ValueError(
            f"O arquivo Camnpal '{nome_arquivo}' possui no máximo "
            f"{df_origem.shape[1]} coluna(s), mas são necessárias "
            "pelo menos 12 colunas, de A até L."
        )

    indice_cabecalho = _localizar_cabecalho_camnpal(
        df_origem,
        nome_arquivo
    )

    df_dados = df_origem.iloc[
        indice_cabecalho + 1:
    ].copy()

    df_dados.reset_index(
        drop=True,
        inplace=True
    )

    registros_validos = []

    for _, linha in df_dados.iterrows():
        # Ignora cabeçalhos repetidos.
        if _linha_eh_cabecalho_camnpal(
            linha
        ):
            continue

        # Ignora linhas vazias, informações complementares e rodapés.
        if not _linha_possui_dados_camnpal(
            linha
        ):
            continue

        conta, descricao = _separar_conta_descricao_camnpal(
            linha.iloc[0]
        )

        codigo_reduzido = _normalizar_codigo_reduzido_camnpal(
            linha.iloc[1]
        )

        saldo_anterior = _aplicar_natureza_camnpal(
            linha.iloc[2],
            linha.iloc[3]
        )

        debito = _aplicar_natureza_camnpal(
            linha.iloc[4],
            linha.iloc[5]
        )

        credito = _aplicar_natureza_camnpal(
            linha.iloc[6],
            linha.iloc[7]
        )

        saldo_acumulado = _aplicar_natureza_camnpal(
            linha.iloc[10],
            linha.iloc[11]
        )

        # A natureza já foi aplicada. Portanto:
        # Débito D será positivo e Crédito C será negativo.
        movimento = round(
            debito + credito,
            2
        )

        registros_validos.append({
            "Conta": conta,
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
            "Movimento": movimento,
            "Saldo Acumulado": round(
                saldo_acumulado,
                2
            ),
        })

    if not registros_validos:
        raise ValueError(
            f"Nenhuma conta válida foi encontrada no arquivo "
            f"Camnpal '{nome_arquivo}'. Verifique se a classificação "
            "e a descrição estão na coluna A e se os valores estão "
            "dispostos nas colunas C até L."
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

    # Recalcula o movimento após o arredondamento dos valores.
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
# FUNÇÃO PÚBLICA DO CLIENTE CAMNPAL
# ==============================================================================

def processar(lista_arquivos):
    """
    Processa os arquivos selecionados para o cliente Camnpal.

    Regras:
    - processa somente arquivos CSV;
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
            "Nenhum arquivo foi selecionado para o cliente Camnpal."
        )

    arquivos_csv = [
        arquivo
        for arquivo in lista_arquivos
        if os.path.splitext(str(arquivo))[1].lower() == ".csv"
    ]

    if not arquivos_csv:
        raise ValueError(
            "Nenhum arquivo CSV válido foi encontrado para o "
            "cliente Camnpal."
        )

    for arquivo in arquivos_csv:
        nome_aba = _gerar_nome_aba_camnpal(
            caminho_arquivo=arquivo,
            nomes_utilizados=nomes_utilizados
        )

        dataframe = transformar_balancete_camnpal(
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
            "Nenhum resultado foi gerado para o cliente Camnpal."
        )

    return resultados