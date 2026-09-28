import csv
import os
import re
import unicodedata

import pandas as pd


# ==============================================================================
# CONFIGURAÇÕES DO SISTEMA C
# ==============================================================================

EXTENSAO_VALIDA_cotrimaio = ".csv"

COLUNAS_DESTINO_cotrimaio = [
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

def _limpar_nome_aba_cotrimaio(nome):
    """
    Ajusta um texto para utilização como nome de aba do Excel.

    Regras:
    - substitui caracteres inválidos por sublinhado;
    - impede nomes vazios;
    - limita o nome da aba a 31 caracteres.
    """
    nome_limpo = re.sub(
        r'[\\/\x2a?:\[\]]',
        "_",
        str(nome).strip()
    )

    if not nome_limpo:
        nome_limpo = "Sem nome"

    return nome_limpo[:31]


def _obter_nome_sem_extensao_cotrimaio(caminho_arquivo):
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


def _obter_mes_cotrimaio(caminho_arquivo):
    """
    Obtém o mês quando o arquivo começa com B_XX.

    Exemplos:
        B_01_BALANCETE.csv -> 01
        B_02.2026.csv      -> 02
        B_12 AJUSTE.csv    -> 12
        BALANCETE.csv      -> None
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


def _registrar_nome_aba_cotrimaio(
    nome,
    nomes_utilizados
):
    """
    Registra um nome de aba caso ainda não esteja em uso.

    A comparação não diferencia letras maiúsculas e minúsculas.
    """
    nome_aba = _limpar_nome_aba_cotrimaio(
        nome
    )

    chave_nome = nome_aba.casefold()

    if chave_nome in nomes_utilizados:
        return None

    nomes_utilizados.add(
        chave_nome
    )

    return nome_aba


def _gerar_nome_aba_cotrimaio(
    caminho_arquivo,
    nomes_utilizados
):
    """
    Gera um nome exclusivo para a aba de destino.

    Regras:
    1. Arquivos B_XX tentam utilizar XX;
    2. Se o mês já estiver em uso, utiliza o nome completo;
    3. Arquivos fora do padrão usam o nome sem extensão;
    4. Duplicidades recebem um sufixo numérico.
    """
    nome_completo = _obter_nome_sem_extensao_cotrimaio(
        caminho_arquivo
    )

    mes = _obter_mes_cotrimaio(
        caminho_arquivo
    )

    if mes:
        nome_aba = _registrar_nome_aba_cotrimaio(
            mes,
            nomes_utilizados
        )

        if nome_aba is not None:
            return nome_aba

    nome_aba = _registrar_nome_aba_cotrimaio(
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

        nome_aba = _registrar_nome_aba_cotrimaio(
            nome_candidato,
            nomes_utilizados
        )

        if nome_aba is not None:
            return nome_aba

        contador += 1


# ==============================================================================
# NORMALIZAÇÃO DOS TEXTOS
# ==============================================================================

def _normalizar_texto_cotrimaio(valor):
    """
    Normaliza um texto extraído do CSV.

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


def _normalizar_texto_comparacao_cotrimaio(valor):
    """
    Normaliza um texto para comparação.

    A função:
    - converte para letras maiúsculas;
    - remove acentos;
    - normaliza espaços.

    Exemplos:
        Débitos Mês  -> DEBITOS MES
        Créditos Mês -> CREDITOS MES
        Saldo Atual  -> SALDO ATUAL
    """
    texto = _normalizar_texto_cotrimaio(
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
# SEPARAÇÃO DA CLASSIFICAÇÃO E DA DESCRIÇÃO
# ==============================================================================

def _separar_conta_descricao_cotrimaio(valor):
    """
    Separa a classificação e a descrição presentes na coluna A.

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
    """
    texto = _normalizar_texto_cotrimaio(
        valor
    )

    if not texto:
        return "", ""

    correspondencia = re.match(
        r"^\s{0,}"
        r"(?P<conta>\d+(?:(?:\s+|[.\-/])\d+){0,})"
        r"\s+"
        r"(?P<descricao>.+?)"
        r"\s{0,}$",
        texto
    )

    if correspondencia is None:
        return "", ""

    conta = re.sub(
        r"\s{1,}",
        " ",
        correspondencia.group(
            "conta"
        )
    ).strip()

    descricao = re.sub(
        r"\s{1,}",
        " ",
        correspondencia.group(
            "descricao"
        )
    ).strip()

    if not conta:
        return "", ""

    if not descricao:
        return "", ""

    return conta, descricao


def _normalizar_codigo_reduzido_cotrimaio(valor):
    """
    Normaliza o código reduzido.

    Exemplos:
        1000   -> 1000
        1000.0 -> 1000
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

    texto = _normalizar_texto_cotrimaio(
        valor
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


# ==============================================================================
# CONVERSÃO DOS VALORES
# ==============================================================================

def _converter_numero_cotrimaio(valor):
    """
    Converte valores monetários brasileiros para float.

    Exemplos:
        244.661.126,28  -> 244661126.28
        -75.377.909,30  -> -75377909.30
        0,00            -> 0.0
        -               -> 0.0
        (1.250,50)      -> -1250.50

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

    if (
        negativo_parenteses
        or negativo_final
    ):
        numero = -abs(
            numero
        )

    return round(
        numero,
        2
    )


def _normalizar_natureza_cotrimaio(natureza):
    """
    Normaliza a natureza contábil.

    Valores reconhecidos:
        D
        C
        Débito
        Crédito
        Devedor
        Credor
    """
    natureza_normalizada = (
        _normalizar_texto_comparacao_cotrimaio(
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


def _aplicar_natureza_cotrimaio(
    valor,
    natureza
):
    """
    Aplica a natureza D/C ao valor.

    Regras:
        D = positivo
        C = negativo

    O sinal originalmente existente no valor é ignorado quando
    uma natureza válida for informada.
    """
    numero_original = _converter_numero_cotrimaio(
        valor
    )

    numero_absoluto = round(
        abs(numero_original),
        2
    )

    natureza_normalizada = _normalizar_natureza_cotrimaio(
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
# LEITURA DO CSV
# ==============================================================================

def _ler_csv_cotrimaio(caminho_arquivo):
    """
    Lê o CSV da Cotrimaio separado por ponto e vírgula.

    A leitura é feita linha a linha para permitir:
    - cabeçalhos com poucas colunas;
    - linhas vazias;
    - quantidades diferentes de campos;
    - delimitador excedente no final;
    - diferentes codificações.

    Codificações testadas:
        utf-8-sig
        cp1252
        latin-1
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

    # Remove somente campos vazios excedentes no final.
    # Os campos vazios posicionais entre as colunas são preservados.
    for linha in linhas_lidas:
        while (
            len(linha) > 0
            and not str(linha[-1]).strip()
        ):
            linha.pop()

    maior_quantidade_colunas = max(
        len(linha)
        for linha in linhas_lidas
    )

    if maior_quantidade_colunas == 0:
        raise ValueError(
            f"O arquivo CSV '{nome_arquivo}' não possui "
            "conteúdo válido."
        )

    linhas_padronizadas = []

    for linha in linhas_lidas:
        quantidade_faltante = (
            maior_quantidade_colunas
            - len(linha)
        )

        complementos = [
            ""
            for _ in range(
                quantidade_faltante
            )
        ]

        linha_padronizada = (
            linha
            + complementos
        )

        linhas_padronizadas.append(
            linha_padronizada
        )

    return pd.DataFrame(
        linhas_padronizadas,
        dtype=str
    )


# ==============================================================================
# IDENTIFICAÇÃO DO CABEÇALHO
# ==============================================================================

def _linha_eh_cabecalho_cotrimaio(linha):
    """
    Identifica a linha de cabeçalho do balancete.

    Títulos esperados:
        Conta
        Saldo Anterior
        Débitos Mês
        Créditos Mês
        Saldo Mês
        Saldo Atual
    """
    textos_linha = [
        _normalizar_texto_comparacao_cotrimaio(
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
        "SALDO ANTERIOR",
        "DEBITOS MES",
        "CREDITOS MES",
        "SALDO MES",
        "SALDO ATUAL",
    ]

    quantidade_encontrada = sum(
        termo in texto_completo
        for termo in termos_cabecalho
    )

    return quantidade_encontrada >= 3


def _linha_possui_dados_cotrimaio(linha):
    """
    Verifica se a linha contém conta e descrição válidas.
    """
    if len(linha) < 12:
        return False

    conta, descricao = (
        _separar_conta_descricao_cotrimaio(
            linha.iloc[0]
        )
    )

    if not conta:
        return False

    if not descricao:
        return False

    return True


def _localizar_inicio_dados_cotrimaio(
    df_origem,
    nome_arquivo
):
    """
    Localiza o início dos dados.

    Primeira estratégia:
        localiza o cabeçalho iniciado por Conta.

    Segunda estratégia:
        localiza diretamente a primeira linha contábil válida.
    """
    for indice_linha in range(
        df_origem.shape[0]
    ):
        linha = df_origem.iloc[
            indice_linha
        ]

        primeira_coluna = (
            _normalizar_texto_comparacao_cotrimaio(
                linha.iloc[0]
            )
        )

        if (
            primeira_coluna == "CONTA"
            and _linha_eh_cabecalho_cotrimaio(
                linha
            )
        ):
            return indice_linha + 1

    for indice_linha in range(
        df_origem.shape[0]
    ):
        linha = df_origem.iloc[
            indice_linha
        ]

        if _linha_possui_dados_cotrimaio(
            linha
        ):
            return indice_linha

    raise ValueError(
        f"Não foi possível localizar o cabeçalho ou a primeira "
        f"conta válida no arquivo Cotrimaio '{nome_arquivo}'."
    )


# ==============================================================================
# EXTRAÇÃO DO REGISTRO
# ==============================================================================

def _extrair_registro_cotrimaio(linha):
    """
    Extrai e transforma uma linha contábil da Cotrimaio.

    Origem:
        A = Classificação e descrição
        B = Código reduzido
        C = Saldo Anterior
        D = Natureza do Saldo Anterior
        E = Débito do Mês
        F = Natureza do Débito
        G = Crédito do Mês
        H = Natureza do Crédito
        I = Movimento do Mês
        J = Natureza do Movimento
        K = Saldo Atual
        L = Natureza do Saldo Atual
    """
    if not _linha_possui_dados_cotrimaio(
        linha
    ):
        return None

    conta, descricao = (
        _separar_conta_descricao_cotrimaio(
            linha.iloc[0]
        )
    )

    codigo_reduzido = (
        _normalizar_codigo_reduzido_cotrimaio(
            linha.iloc[1]
        )
    )

    saldo_anterior = _aplicar_natureza_cotrimaio(
        linha.iloc[2],
        linha.iloc[3]
    )

    debito = _aplicar_natureza_cotrimaio(
        linha.iloc[4],
        linha.iloc[5]
    )

    credito = _aplicar_natureza_cotrimaio(
        linha.iloc[6],
        linha.iloc[7]
    )

    movimento = _aplicar_natureza_cotrimaio(
        linha.iloc[8],
        linha.iloc[9]
    )

    saldo_acumulado = _aplicar_natureza_cotrimaio(
        linha.iloc[10],
        linha.iloc[11]
    )

    return {
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

def _conferir_registro_cotrimaio(registro):
    """
    Confere o movimento e o saldo atual.

    Movimento calculado:
        Débito + Crédito

    Saldo calculado:
        Saldo Anterior + Movimento informado

    A conferência é informativa e não altera os valores.
    """
    movimento_calculado = round(
        registro["Débito"]
        + registro["Crédito"],
        2
    )

    diferenca_movimento = round(
        registro["Movimento"]
        - movimento_calculado,
        2
    )

    saldo_calculado = round(
        registro["Saldo Anterior"]
        + registro["Movimento"],
        2
    )

    diferenca_saldo = round(
        registro["Saldo Acumulado"]
        - saldo_calculado,
        2
    )

    return {
        "Diferença Movimento": (
            diferenca_movimento
        ),
        "Diferença Saldo": (
            diferenca_saldo
        ),
    }


# ==============================================================================
# MONTAGEM DO DATAFRAME FINAL
# ==============================================================================

def _montar_dataframe_cotrimaio(
    registros_validos
):
    """
    Monta o DataFrame final na ordem exigida pelo tabulador.
    """
    if not registros_validos:
        return pd.DataFrame(
            columns=COLUNAS_DESTINO_cotrimaio
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
        COLUNAS_DESTINO_cotrimaio
    ].copy()


# ==============================================================================
# TRANSFORMAÇÃO DO BALANCETE
# ==============================================================================

def transformar_balancete_cotrimaio(caminho_arquivo):
    """
    Transforma um balancete CSV do sistema Cotrimaio.

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

    if extensao != EXTENSAO_VALIDA_cotrimaio:
        raise ValueError(
            f"O arquivo '{nome_arquivo}' não possui extensão CSV."
        )

    df_origem = _ler_csv_cotrimaio(
        caminho_arquivo
    )

    if df_origem.empty:
        raise ValueError(
            f"O arquivo Cotrimaio '{nome_arquivo}' está vazio."
        )

    if df_origem.shape[1] < 12:
        raise ValueError(
            f"O arquivo Cotrimaio '{nome_arquivo}' possui no máximo "
            f"{df_origem.shape[1]} coluna(s), mas são necessárias "
            "pelo menos 12 colunas, de A até L."
        )

    indice_inicio = _localizar_inicio_dados_cotrimaio(
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
        if _linha_eh_cabecalho_cotrimaio(
            linha
        ):
            continue

        registro = _extrair_registro_cotrimaio(
            linha
        )

        if registro is None:
            conteudo = _normalizar_texto_cotrimaio(
                linha.iloc[0]
            )

            if conteudo:
                conta, descricao = (
                    _separar_conta_descricao_cotrimaio(
                        conteudo
                    )
                )

                if conta or descricao:
                    linhas_ignoradas.append({
                        "Linha": (
                            indice_inicio
                            + indice_relativo
                            + 1
                        ),
                        "Conteúdo": conteudo,
                    })

            continue

        conferencia = _conferir_registro_cotrimaio(
            registro
        )

        if (
            abs(
                conferencia[
                    "Diferença Movimento"
                ]
            ) > 0.02
            or abs(
                conferencia[
                    "Diferença Saldo"
                ]
            ) > 0.02
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
                    conferencia[
                        "Diferença Movimento"
                    ]
                ),
                "Diferença Saldo": (
                    conferencia[
                        "Diferença Saldo"
                    ]
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
                        f"{item['Conteúdo']}"
                    )
                    for item in exemplos
                )
            )

        raise ValueError(
            f"Nenhuma conta válida foi encontrada no arquivo "
            f"Cotrimaio '{nome_arquivo}'. Verifique se o separador "
            "é ponto e vírgula e se os dados estão nas colunas "
            "A até L."
            + detalhe
        )

    if linhas_ignoradas:
        print(
            f"[Cotrimaio] "
            f"{len(linhas_ignoradas)} linha(s) foram ignoradas "
            f"no arquivo '{nome_arquivo}'."
        )

        for item in linhas_ignoradas:
            print(
                f"[Cotrimaio] Linha {item['Linha']}: "
                f"{item['Conteúdo']}"
            )

    if divergencias:
        print(
            f"[Cotrimaio] "
            f"{len(divergencias)} linha(s) apresentaram "
            f"divergência no arquivo '{nome_arquivo}'."
        )

        for item in divergencias:
            print(
                f"[Cotrimaio] Linha {item['Linha']}: "
                f"Conta={item['Conta']}; "
                f"Nome={item['Nome']}; "
                f"Diferença movimento="
                f"{item['Diferença Movimento']:.2f}; "
                f"Diferença saldo="
                f"{item['Diferença Saldo']:.2f}"
            )

    df_destino = _montar_dataframe_cotrimaio(
        registros_validos
    )

    if df_destino.empty:
        raise ValueError(
            f"O arquivo Cotrimaio '{nome_arquivo}' não gerou "
            "nenhum registro no layout de destino."
        )

    return df_destino


# ==============================================================================
# FUNÇÃO PÚBLICA DO Cotrimaio
# ==============================================================================

def processar(lista_arquivos):
    """
    Processa os arquivos selecionados para o sistema Cotrimaio.

    Regras:
    - processa somente arquivos CSV;
    - ignora outras extensões;
    - cada arquivo gera uma aba;
    - arquivos B_XX tentam utilizar XX;
    - meses repetidos usam o nome completo;
    - nomes repetidos recebem sufixo numérico.

    Retorno:
        {
            "Nome da aba": DataFrame
        }
    """
    resultados = {}
    nomes_utilizados = set()

    if not lista_arquivos:
        raise ValueError(
            "Nenhum arquivo foi selecionado para o sistema Cotrimaio."
        )

    arquivos_csv = [
        arquivo
        for arquivo in lista_arquivos
        if os.path.splitext(
            str(arquivo)
        )[1].lower() == EXTENSAO_VALIDA_cotrimaio
    ]

    if not arquivos_csv:
        raise ValueError(
            "Nenhum arquivo CSV válido foi encontrado para o "
            "sistema Cotrimaio."
        )

    for arquivo in arquivos_csv:
        nome_aba = _gerar_nome_aba_cotrimaio(
            caminho_arquivo=arquivo,
            nomes_utilizados=nomes_utilizados
        )

        dataframe = transformar_balancete_cotrimaio(
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
            "Nenhum resultado foi gerado para o sistema Cotrimaio."
        )

    return resultados