import os
import re

import pandas as pd


# ==============================================================================
# CONFIGURAÇÕES DO SISTEMA Cotrisal
# ==============================================================================

EXTENSAO_VALIDA_cotrisal = ".txt"


COLUNAS_DESTINO_cotrisal = [
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

def _limpar_nome_aba_cotrisal(nome):
    """
    Ajusta o texto para utilização como nome de aba do Excel.

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


def _obter_nome_sem_extensao_cotrisal(caminho_arquivo):
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


def _obter_mes_cotrisal(caminho_arquivo):
    """
    Extrai o mês quando o nome do arquivo começa com B_XX.

    Exemplos:
        B_01_BALANCETE.txt -> 01
        B_05.2026.txt      -> 05
        BALANCETE.txt      -> None

    Somente meses entre 01 e 12 são considerados válidos.
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


def _registrar_nome_aba_cotrisal(
    nome,
    nomes_utilizados
):
    """
    Registra um nome de aba caso ainda não esteja sendo utilizado.

    A comparação não diferencia letras maiúsculas e minúsculas.
    """
    nome_aba = _limpar_nome_aba_cotrisal(
        nome
    )

    chave_nome = nome_aba.casefold()

    if chave_nome in nomes_utilizados:
        return None

    nomes_utilizados.add(
        chave_nome
    )

    return nome_aba


def _gerar_nome_aba_cotrisal(
    caminho_arquivo,
    nomes_utilizados
):
    """
    Gera um nome exclusivo para a aba de destino.

    Regras:
    1. Arquivos B_XX tentam utilizar XX;
    2. Se XX já estiver utilizado, usa o nome completo do arquivo;
    3. Arquivos fora do padrão usam o nome sem extensão;
    4. Nomes repetidos recebem um sufixo numérico.
    """
    nome_completo = _obter_nome_sem_extensao_cotrisal(
        caminho_arquivo
    )

    mes = _obter_mes_cotrisal(
        caminho_arquivo
    )

    if mes:
        nome_aba = _registrar_nome_aba_cotrisal(
            mes,
            nomes_utilizados
        )

        if nome_aba is not None:
            return nome_aba

    nome_aba = _registrar_nome_aba_cotrisal(
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

        nome_aba = _registrar_nome_aba_cotrisal(
            nome_candidato,
            nomes_utilizados
        )

        if nome_aba is not None:
            return nome_aba

        contador += 1


# ==============================================================================
# NORMALIZAÇÃO DOS TEXTOS
# ==============================================================================

def _normalizar_texto_cotrisal(valor):
    """
    Normaliza textos extraídos do TXT.

    A função:
    - converte valores ausentes em texto vazio;
    - substitui espaços especiais;
    - reduz espaços consecutivos;
    - remove espaços das extremidades.

    A função não junta as letras da nomenclatura.
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


def _normalizar_classificacao_cotrisal(valor):
    """
    Remove os espaços existentes na classificação contábil.

    Exemplos:
        1           -> 1
        1 01        -> 101
        1 01 01     -> 10101
        1 01 01 001 -> 10101001

    Somente os espaços são removidos. Os números são preservados.
    """
    texto = _normalizar_texto_cotrisal(
        valor
    )

    if not texto:
        return ""

    classificacao = re.sub(
        r"\s{1,}",
        "",
        texto
    )

    if not classificacao.isdigit():
        return ""

    return classificacao


# ==============================================================================
# CONVERSÃO DOS VALORES
# ==============================================================================

def _converter_numero_cotrisal(valor):
    """
    Converte valores monetários para float.

    Formatos reconhecidos:
        2791.129.658,74
        -2397.660.743,17
        0,00
        -1.000,00
        (1.000,00)
        1.000,00-

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


# ==============================================================================
# LEITURA DO TXT
# ==============================================================================

def _ler_linhas_txt_cotrisal(caminho_arquivo):
    """
    Lê as linhas do arquivo TXT.

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

    ultimo_erro = None

    for codificacao in codificacoes:
        try:
            with open(
                caminho_arquivo,
                mode="r",
                encoding=codificacao
            ) as arquivo:
                return arquivo.readlines()

        except UnicodeDecodeError as erro:
            ultimo_erro = erro

        except OSError as erro:
            raise ValueError(
                f"Não foi possível abrir o arquivo TXT "
                f"'{nome_arquivo}'. Erro: {erro}"
            ) from erro

    raise ValueError(
        f"Não foi possível identificar a codificação do arquivo "
        f"TXT '{nome_arquivo}'. Erro: {ultimo_erro}"
    )


# ==============================================================================
# IDENTIFICAÇÃO DE CABEÇALHOS
# ==============================================================================

def _linha_eh_cabecalho_cotrisal(linha):
    """
    Identifica cabeçalhos, parâmetros e quebras de página.

    Uma linha contábil não será descartada apenas porque sua
    nomenclatura contém o texto C O N T A.

    Exemplo que deve ser processado:
        3 C O N T A S DE R E S U L T A D O

    Exemplo que deve ser ignorado:
        C O N T A N O M E N C L A T U R A
        SALDO ANTERIOR DEBITO CREDITO SALDO ATUAL
    """
    texto = _normalizar_texto_cotrisal(
        linha
    ).upper()

    if not texto:
        return True

    # Linhas de parâmetros.
    if texto.startswith("MES:"):
        return True

    if (
        "MES:" in texto
        and "LOCAL:" in texto
    ):
        return True

    if "CENTRO DE CUSTOS:" in texto:
        return True

    # Cabeçalho principal.
    possui_conta = (
        "C O N T A" in texto
    )

    possui_nomenclatura = (
        "N O M E N C L A T U R A" in texto
    )

    possui_saldo_anterior = (
        "SALDO ANTERIOR" in texto
    )

    possui_debito = (
        "DEBITO" in texto
        or "DÉBITO" in texto
    )

    possui_credito = (
        "CREDITO" in texto
        or "CRÉDITO" in texto
    )

    possui_saldo_atual = (
        "SALDO ATUAL" in texto
    )

    quantidade_termos = sum([
        possui_conta,
        possui_nomenclatura,
        possui_saldo_anterior,
        possui_debito,
        possui_credito,
        possui_saldo_atual,
    ])

    if quantidade_termos >= 3:
        return True

    # Cabeçalhos de páginas intermediárias.
    if (
        "LISTA BALANCETE CONTABIL" in texto
        or "DATA EMISSAO:" in texto
        or "HORA EMISSAO:" in texto
    ):
        return True

    # Totais finais do relatório.
    if texto.startswith(
        "T O T A L"
    ):
        return True

    return False


# ==============================================================================
# EXTRAÇÃO DOS CAMPOS
# ==============================================================================

def _obter_correspondencias_monetarias_cotrisal(linha):
    """
    Localiza valores monetários brasileiros em uma linha.

    Exemplos reconhecidos:
        2791.129.658,74
        3001.218.345,25
        -2397.660.743,17
        3394.687.260,82
        1.000,00
        0,00

    O primeiro grupo pode possuir qualquer quantidade de dígitos.
    """
    if linha is None:
        return []

    padrao_monetario = re.compile(
        r"(?<![\d.,])"
        r"[+-]?"
        r"\d{1,}"
        r"(?:\.\d{3}){0,}"
        r",\d{2}"
        r"-?"
        r"(?![\d.,])"
    )

    return list(
        padrao_monetario.finditer(
            str(linha)
        )
    )


def _separar_classificacao_nome_cotrisal(prefixo):
    """
    Separa a classificação e a nomenclatura antes dos valores.

    A classificação pode conter espaços entre seus grupos.

    Exemplos:
        1     A T I V O
            Conta: 1
            Nome: A T I V O

        1 01     ATIVO CIRCULANTE
            Conta: 101
            Nome: ATIVO CIRCULANTE

        1 01 01     CAIXA E BANCOS
            Conta: 10101
            Nome: CAIXA E BANCOS

        3 CONTAS DE RESULTADO
            Conta: 3
            Nome: CONTAS DE RESULTADO

    Somente os espaços da classificação são removidos.
    Os espaços da nomenclatura são preservados, com retirada apenas
    dos espaços excedentes utilizados para alinhamento.
    """
    if prefixo is None:
        return "", ""

    prefixo_limpo = (
        str(prefixo)
        .replace("\xa0", " ")
        .replace("\r", "")
        .replace("\n", "")
        .strip()
    )

    if not prefixo_limpo:
        return "", ""

    correspondencia = re.fullmatch(
        r"\s{0,}"
        r"(?P<classificacao>\d+(?:\s+\d+){0,})"
        r"\s+"
        r"(?P<nome>.+?)"
        r"\s{0,}",
        prefixo_limpo
    )

    if correspondencia is None:
        return "", ""

    classificacao_original = correspondencia.group(
        "classificacao"
    )

    nome_original = correspondencia.group(
        "nome"
    )

    classificacao = _normalizar_classificacao_cotrisal(
        classificacao_original
    )

    nome = _normalizar_texto_cotrisal(
        nome_original
    )

    if not classificacao:
        return "", ""

    if not nome:
        return "", ""

    return classificacao, nome


def _extrair_registro_cotrisal(linha):
    """
    Extrai uma linha contábil do TXT do sistema Cotrisal.

    Estrutura esperada:
        classificação;
        nomenclatura;
        saldo anterior;
        débito;
        crédito;
        saldo atual.

    Regras:
        - utiliza os quatro últimos valores monetários da linha;
        - considera todo o conteúdo anterior como conta e nome;
        - remove espaços somente da classificação;
        - preserva a nomenclatura;
        - mantém o sinal dos valores da origem;
        - calcula Movimento como Débito mais Crédito;
        - arredonda os valores para duas casas decimais.
    """
    if linha is None:
        return None

    linha_limpa = (
        str(linha)
        .replace("\xa0", " ")
        .replace("\r", "")
        .replace("\n", "")
        .rstrip()
    )

    if not linha_limpa.strip():
        return None

    if _linha_eh_cabecalho_cotrisal(
        linha_limpa
    ):
        return None

    valores_encontrados = (
        _obter_correspondencias_monetarias_cotrisal(
            linha_limpa
        )
    )

    if len(valores_encontrados) < 4:
        return None

    # Utiliza os quatro últimos valores para evitar que números
    # eventualmente existentes no nome interfiram na extração.
    valores_utilizados = valores_encontrados[
        -4:
    ]

    primeiro_valor = valores_utilizados[0]

    prefixo = linha_limpa[
        :primeiro_valor.start()
    ].rstrip()

    classificacao, nome = (
        _separar_classificacao_nome_cotrisal(
            prefixo
        )
    )

    if not classificacao:
        return None

    if not nome:
        return None

    saldo_anterior = round(
        _converter_numero_cotrisal(
            valores_utilizados[0].group()
        ),
        2
    )

    debito = round(
        _converter_numero_cotrisal(
            valores_utilizados[1].group()
        ),
        2
    )

    credito = round(
        _converter_numero_cotrisal(
            valores_utilizados[2].group()
        ),
        2
    )

    saldo_acumulado = round(
        _converter_numero_cotrisal(
            valores_utilizados[3].group()
        ),
        2
    )

    movimento = round(
        debito + credito,
        2
    )

    return {
        "Conta": classificacao,
        "Nome": nome,
        "Cód. Reduzido": classificacao,
        "Saldo Anterior": saldo_anterior,
        "Débito": debito,
        "Crédito": credito,
        "Movimento": movimento,
        "Saldo Acumulado": saldo_acumulado,
    }


# ==============================================================================
# MONTAGEM DO DATAFRAME
# ==============================================================================

def _montar_dataframe_cotrisal(registros):
    """
    Monta o DataFrame final na ordem padrão do tabulador.
    """
    if not registros:
        return pd.DataFrame(
            columns=COLUNAS_DESTINO_cotrisal
        )

    dataframe = pd.DataFrame(
        registros
    )

    dataframe.insert(
        0,
        "Atividade",
        "Geral"
    )

    colunas_texto = [
        "Atividade",
        "Conta",
        "Nome",
        "Cód. Reduzido",
    ]

    for coluna in colunas_texto:
        dataframe[coluna] = (
            dataframe[coluna]
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
        dataframe[coluna] = (
            pd.to_numeric(
                dataframe[coluna],
                errors="coerce"
            )
            .fillna(0.0)
            .round(2)
        )

    dataframe["Movimento"] = (
        dataframe["Débito"]
        + dataframe["Crédito"]
    ).round(2)

    dataframe = dataframe[
        COLUNAS_DESTINO_cotrisal
    ].copy()

    dataframe.reset_index(
        drop=True,
        inplace=True
    )

    return dataframe


# ==============================================================================
# TRANSFORMAÇÃO DO BALANCETE
# ==============================================================================

def transformar_balancete_cotrisal(caminho_arquivo):
    """
    Transforma um arquivo TXT do sistema Cotrisal.

    Layout de origem:
        A = Classificação
        B = Nomenclatura
        C = Saldo Anterior
        D = Débito
        E = Crédito
        F = Saldo Atual

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

    if extensao != EXTENSAO_VALIDA_cotrisal:
        raise ValueError(
            f"O arquivo '{nome_arquivo}' não possui extensão TXT."
        )

    linhas = _ler_linhas_txt_cotrisal(
        caminho_arquivo
    )

    if not linhas:
        raise ValueError(
            f"O arquivo TXT '{nome_arquivo}' está vazio."
        )

    registros = []
    linhas_contabeis_ignoradas = []

    for numero_linha, linha in enumerate(
        linhas,
        start=1
    ):
        registro = _extrair_registro_cotrisal(
            linha
        )

        if registro is not None:
            registros.append(
                registro
            )

            continue

        linha_limpa = (
            str(linha)
            .replace("\xa0", " ")
            .strip()
        )

        if re.match(
            r"^\s{0,}\d",
            linha_limpa
        ):
            valores = (
                _obter_correspondencias_monetarias_cotrisal(
                    linha_limpa
                )
            )

            if len(valores) >= 4:
                linhas_contabeis_ignoradas.append({
                    "Linha": numero_linha,
                    "Conteúdo": linha_limpa,
                })

    if not registros:
        detalhe = ""

        if linhas_contabeis_ignoradas:
            exemplos = linhas_contabeis_ignoradas[
                :5
            ]

            detalhe = (
                " Exemplos de linhas não reconhecidas: "
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
            f"Cotrisal '{nome_arquivo}'. Verifique se cada linha "
            "possui classificação, nomenclatura, saldo anterior, "
            "débito, crédito e saldo atual."
            + detalhe
        )

    if linhas_contabeis_ignoradas:
        print(
            f"[Cotrisal] "
            f"{len(linhas_contabeis_ignoradas)} linha(s) "
            f"aparentemente contábil(eis) foram ignoradas "
            f"no arquivo '{nome_arquivo}'."
        )

        for item in linhas_contabeis_ignoradas:
            print(
                f"[Cotrisal] Linha {item['Linha']}: "
                f"{item['Conteúdo']}"
            )

    dataframe = _montar_dataframe_cotrisal(
        registros
    )

    if dataframe.empty:
        raise ValueError(
            f"O arquivo Cotrisal '{nome_arquivo}' não gerou "
            "registros válidos."
        )

    return dataframe


# ==============================================================================
# ORDENAÇÃO DOS ARQUIVOS
# ==============================================================================

def _chave_ordenacao_arquivo_cotrisal(caminho_arquivo):
    """
    Ordena arquivos B_XX pelo mês.

    Arquivos fora do padrão são posicionados depois dos arquivos
    mensais e ordenados pelo nome.
    """
    mes = _obter_mes_cotrisal(
        caminho_arquivo
    )

    if mes is None:
        numero_mes = 99
    else:
        numero_mes = int(
            mes
        )

    nome_arquivo = os.path.basename(
        caminho_arquivo
    ).casefold()

    return (
        numero_mes,
        nome_arquivo
    )


# ==============================================================================
# FUNÇÃO PÚBLICA DO SISTEMA Cotrisal
# ==============================================================================

def processar(lista_arquivos):
    """
    Processa os arquivos selecionados para o Cotrisal.

    Regras:
    - processa somente arquivos TXT;
    - ignora arquivos de outras extensões;
    - cada arquivo válido gera uma aba;
    - arquivos B_XX tentam utilizar XX;
    - meses repetidos utilizam o nome completo do arquivo;
    - arquivos mensais são processados em ordem cronológica.
    """
    if not lista_arquivos:
        raise ValueError(
            "Nenhum arquivo foi selecionado para o sistema Cotrisal."
        )

    arquivos_txt = [
        arquivo
        for arquivo in lista_arquivos
        if os.path.splitext(
            str(arquivo)
        )[1].lower() == EXTENSAO_VALIDA_cotrisal
    ]

    if not arquivos_txt:
        raise ValueError(
            "Nenhum arquivo TXT válido foi encontrado para o "
            "sistema Cotrisal."
        )

    arquivos_ordenados = sorted(
        arquivos_txt,
        key=_chave_ordenacao_arquivo_cotrisal
    )

    resultados = {}
    nomes_utilizados = set()

    for arquivo in arquivos_ordenados:
        nome_aba = _gerar_nome_aba_cotrisal(
            caminho_arquivo=arquivo,
            nomes_utilizados=nomes_utilizados
        )

        dataframe = transformar_balancete_cotrisal(
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
                "retornou dados válidos."
            )

        resultados[nome_aba] = dataframe

    if not resultados:
        raise ValueError(
            "Nenhum resultado foi gerado para o sistema Cotrisal."
        )

    return resultados