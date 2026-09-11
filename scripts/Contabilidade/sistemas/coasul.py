import os
import re
import pandas as pd


def _limpar_nome_aba(nome):
    """
    Ajusta o nome para utilização como aba do Excel.
    """
    nome_limpo = re.sub(
        r'[\\/*?:\[\]]',
        "_",
        str(nome).strip()
    )

    if not nome_limpo:
        nome_limpo = "Sem nome"

    return nome_limpo[:31]


def _obter_nome_sem_extensao(caminho_arquivo):
    """
    Retorna o nome completo do arquivo sem sua extensão.
    """
    nome_arquivo = os.path.basename(caminho_arquivo)
    nome_sem_extensao = os.path.splitext(nome_arquivo)[0]

    return nome_sem_extensao.strip()


def _obter_mes(caminho_arquivo):
    """
    Extrai o mês quando o nome do arquivo começa com B_XX.

    Exemplos:
        B_01_BALANCETE.txt -> 01
        B_12.2026.txt      -> 12
        BALANCETE.txt      -> None
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
    Registra um nome de aba se ele ainda não estiver em uso.
    """
    nome_aba = _limpar_nome_aba(nome)
    chave_nome = nome_aba.casefold()

    if chave_nome in nomes_utilizados:
        return None

    nomes_utilizados.add(chave_nome)

    return nome_aba


def _gerar_nome_aba(caminho_arquivo, nomes_utilizados):
    """
    Gera um nome exclusivo para cada aba.

    Regras:
    - arquivos iniciados por B_XX tentam usar XX;
    - se XX já estiver em uso, utiliza o nome completo;
    - arquivos fora do padrão usam o nome sem extensão;
    - nomes totalmente repetidos recebem um sufixo numérico.
    """
    nome_completo = _obter_nome_sem_extensao(caminho_arquivo)
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
    Processa os balancetes TXT do cliente Coasul.

    Somente arquivos TXT são processados.
    Arquivos com outras extensões são ignorados.
    """
    resultados = {}
    nomes_utilizados = set()

    if not lista_arquivos:
        raise ValueError(
            "Nenhum arquivo foi selecionado para o cliente Coasul."
        )

    arquivos_txt = [
        arquivo
        for arquivo in lista_arquivos
        if os.path.splitext(str(arquivo))[1].lower() == ".txt"
    ]

    if not arquivos_txt:
        raise ValueError(
            "Nenhum arquivo TXT válido foi encontrado para o "
            "cliente Coasul."
        )

    for arquivo in arquivos_txt:
        nome_aba = _gerar_nome_aba(
            caminho_arquivo=arquivo,
            nomes_utilizados=nomes_utilizados
        )

        dataframe = transformar_balancete_coasul(
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
            "Nenhum resultado foi gerado para o cliente Coasul."
        )

    return resultados
# ==============================================================================
# TRANSFORMAÇÃO DO BALANCETE COASUL
# ==============================================================================

def _normalizar_texto_coasul(valor):
    """
    Normaliza textos extraídos do relatório do cliente Coasul.
    """
    if valor is None or pd.isna(valor):
        return ""

    texto = str(valor).replace("\xa0", " ")
    texto = re.sub(r"\s+", " ", texto)

    return texto.strip()


def _normalizar_descricao_coasul(valor):
    """
    Normaliza a descrição da conta.

    Quando todas as partes da descrição forem letras isoladas,
    as letras serão unidas.

    Exemplos:
        A  T  I  V  O       -> ATIVO
        C I R C U L A N T E -> CIRCULANTE
    """
    texto = _normalizar_texto_coasul(valor)

    if not texto:
        return ""

    partes = texto.split()

    todas_partes_sao_letras_isoladas = (
        len(partes) > 1
        and all(
            len(parte) == 1 and parte.isalpha()
            for parte in partes
        )
    )

    if todas_partes_sao_letras_isoladas:
        return "".join(partes)

    return texto


def _normalizar_classificacao_coasul(valor):
    """
    Normaliza uma classificação com 10 dígitos.

    A classificação é dividida em grupos de dois dígitos.
    Somente os grupos finais iguais a 00 são removidos.

    Exemplos:
        0100000000 -> 01
        0101000000 -> 0101
        0101010000 -> 010101
        0101011000 -> 01010110
        0101011010 -> 0101011010
    """
    if valor is None or pd.isna(valor):
        return ""

    classificacao = str(valor).strip()

    classificacao = (
        classificacao
        .replace("\xa0", "")
        .replace(" ", "")
    )

    if not classificacao:
        return ""

    if not classificacao.isdigit():
        return classificacao

    if len(classificacao) != 10:
        return classificacao

    grupos = [
        classificacao[indice:indice + 2]
        for indice in range(0, 10, 2)
    ]

    while len(grupos) > 1 and grupos[-1] == "00":
        grupos.pop()

    return "".join(grupos)


def _converter_numero_coasul(valor):
    """
    Converte valores monetários da Coasul para float.

    Reconhece:
        1.250,50
        -1.250,50
        1.250,50-
        0,00
        -
        vazio
    """
    if valor is None or pd.isna(valor):
        return 0.0

    if isinstance(valor, (int, float)):
        return float(valor)

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

        return numero

    except (ValueError, TypeError):
        return 0.0


def _aplicar_natureza_coasul(valor, natureza):
    """
    Aplica o sinal ao valor conforme a natureza contábil.

    D = positivo
    C = negativo
    """
    numero = _converter_numero_coasul(valor)

    natureza_normalizada = (
        _normalizar_texto_coasul(natureza)
        .upper()
        .replace(".", "")
    )

    if natureza_normalizada == "D":
        return abs(numero)

    if natureza_normalizada == "C":
        return -abs(numero)

    return numero


def _ler_linhas_txt_coasul(caminho_arquivo):
    """
    Lê o TXT utilizando diferentes codificações.
    """
    nome_arquivo = os.path.basename(caminho_arquivo)

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


def _extrair_linha_contabil_coasul(linha):
    """
    Extrai uma linha contábil válida do TXT do cliente Coasul.

    Regras:

    1. A conta deve começar com uma classificação de 10 dígitos.

    2. Os quatro últimos valores monetários são:
       - saldo anterior;
       - débito;
       - crédito;
       - saldo final.

    3. O saldo anterior considera a natureza D/C:
       - D = positivo;
       - C = negativo.

    4. Quando o saldo anterior vier negativo na origem, a natureza
       D/C será invertida:
       - valor negativo com D = saldo negativo;
       - valor negativo com C = saldo positivo.

    5. O saldo final sempre será positivo. Qualquer sinal negativo
       existente na origem será ignorado.

    6. O crédito será convertido para negativo.

    7. O movimento será calculado por:
       Movimento = Débito + Crédito.

    8. Todos os cálculos serão arredondados para duas casas decimais.
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

    correspondencia_conta = re.match(
        r"^\s*(?P<classificacao>\d{10})(?P<restante>.*)$",
        linha_limpa
    )

    if correspondencia_conta is None:
        return None

    classificacao_original = correspondencia_conta.group(
        "classificacao"
    )

    restante = correspondencia_conta.group(
        "restante"
    )

    if not restante.strip():
        return None

    # Reconhece valores monetários brasileiros.
    #
    # Exemplos:
    # 1.250,50
    # -1.250,50
    # 1.250,50-
    # 0,00
    padrao_monetario = re.compile(
        r"(?<!\d)"
        r"[+-]?"
        r"(?:\d{1,3}(?:\.\d{3})+|\d+)"
        r",\d{2}"
        r"-?"
        r"(?!\d)"
    )

    valores_encontrados = list(
        padrao_monetario.finditer(restante)
    )

    if len(valores_encontrados) < 4:
        return None

    # Utiliza os quatro últimos valores encontrados na linha.
    valores_utilizados = valores_encontrados[-4:]

    saldo_anterior_texto = valores_utilizados[0].group()
    debito_texto = valores_utilizados[1].group()
    credito_texto = valores_utilizados[2].group()
    saldo_final_texto = valores_utilizados[3].group()

    # Todo o conteúdo anterior ao saldo inicial contém a descrição
    # e a natureza contábil CT.
    prefixo = restante[
        :valores_utilizados[0].start()
    ].rstrip()

    # Procura a natureza D/C no final do prefixo.
    correspondencia_natureza = re.search(
        r"(?:^|\s)(?P<natureza>[DC])\s*$",
        prefixo,
        re.IGNORECASE
    )

    if correspondencia_natureza is not None:
        natureza = (
            correspondencia_natureza
            .group("natureza")
            .strip()
            .upper()
        )

        descricao_bruta = prefixo[
            :correspondencia_natureza.start()
        ]

    else:
        # Caso o alinhamento da linha seja diferente, procura a última
        # ocorrência isolada de D ou C antes dos valores monetários.
        naturezas_encontradas = list(
            re.finditer(
                r"(?<![A-Za-zÀ-ÿ])"
                r"(?P<natureza>[DC])"
                r"(?![A-Za-zÀ-ÿ])",
                prefixo,
                re.IGNORECASE
            )
        )

        if not naturezas_encontradas:
            return None

        ultima_natureza = naturezas_encontradas[-1]

        natureza = (
            ultima_natureza
            .group("natureza")
            .strip()
            .upper()
        )

        descricao_bruta = (
            prefixo[:ultima_natureza.start()]
            + " "
            + prefixo[ultima_natureza.end():]
        )

    classificacao = _normalizar_classificacao_coasul(
        classificacao_original
    )

    descricao = _normalizar_descricao_coasul(
        descricao_bruta
    )

    if not classificacao or not descricao:
        return None

    # ------------------------------------------------------------------
    # SALDO ANTERIOR
    # ------------------------------------------------------------------

    saldo_anterior_origem = round(
        _converter_numero_coasul(
            saldo_anterior_texto
        ),
        2
    )

    saldo_anterior_absoluto = round(
        abs(saldo_anterior_origem),
        2
    )

    saldo_anterior_veio_negativo = (
        saldo_anterior_origem < 0
    )

    if natureza == "D":
        if saldo_anterior_veio_negativo:
            saldo_anterior = -saldo_anterior_absoluto
        else:
            saldo_anterior = saldo_anterior_absoluto

    elif natureza == "C":
        if saldo_anterior_veio_negativo:
            saldo_anterior = saldo_anterior_absoluto
        else:
            saldo_anterior = -saldo_anterior_absoluto

    else:
        return None

    saldo_anterior = round(
        saldo_anterior,
        2
    )

    # ------------------------------------------------------------------
    # DÉBITO
    # ------------------------------------------------------------------

    debito = round(
        abs(
            _converter_numero_coasul(
                debito_texto
            )
        ),
        2
    )

    # ------------------------------------------------------------------
    # CRÉDITO
    # ------------------------------------------------------------------

    # Qualquer sinal existente no crédito da origem é ignorado.
    # O crédito será sempre gravado como negativo.
    credito = round(
        -abs(
            _converter_numero_coasul(
                credito_texto
            )
        ),
        2
    )

    # ------------------------------------------------------------------
    # MOVIMENTO
    # ------------------------------------------------------------------

    movimento = round(
        debito + credito,
        2
    )

    # ------------------------------------------------------------------
    # SALDO FINAL
    # ------------------------------------------------------------------

    # Ignora o sinal apresentado no saldo final da origem.
    saldo_final_absoluto = round(
        abs(
            _converter_numero_coasul(
                saldo_final_texto
            )
        ),
        2
    )

    # Calcula o sinal esperado do saldo final com base na equação
    # contábil: Saldo Anterior + Movimento.
    saldo_final_calculado = round(
        saldo_anterior + movimento,
        2
    )

    # Mantém o valor final informado no relatório, mas aplica o sinal
    # determinado pelo resultado da equação contábil.
    if saldo_final_calculado < 0:
        saldo_acumulado = round(
            -saldo_final_absoluto,
            2
        )
    else:
        saldo_acumulado = saldo_final_absoluto


    return {
        "Conta": classificacao,
        "Nome": descricao,
        "Cód. Reduzido": classificacao,
        "Saldo Anterior": saldo_anterior,
        "Débito": debito,
        "Crédito": credito,
        "Movimento": movimento,
        "Saldo Acumulado": saldo_acumulado,
    }


def transformar_balancete_coasul(caminho_arquivo):
    """
    Transforma o balancete TXT do cliente Coasul.

    Layout final:
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
        - processa somente arquivos TXT;
        - ignora cabeçalhos, rodapés e quebras de página;
        - remove grupos finais "00" excedentes da classificação;
        - mantém os débitos positivos;
        - transforma os créditos em negativos;
        - calcula Movimento como Débito + Crédito;
        - trata o saldo anterior conforme a natureza D/C;
        - inverte a natureza quando o saldo anterior vier negativo;
        - ignora o sinal original do saldo final;
        - define o sinal do saldo final pelo resultado:
          Saldo Anterior + Movimento;
        - arredonda todos os valores monetários para duas casas.
    """
    nome_arquivo = os.path.basename(caminho_arquivo)
    extensao = os.path.splitext(caminho_arquivo)[1].lower()

    if extensao != ".txt":
        raise ValueError(
            f"O arquivo '{nome_arquivo}' não possui extensão TXT."
        )

    linhas = _ler_linhas_txt_coasul(
        caminho_arquivo
    )

    if not linhas:
        raise ValueError(
            f"O arquivo TXT '{nome_arquivo}' está vazio."
        )

    registros_validos = []
    linhas_contabeis_ignoradas = []

    for numero_linha, linha in enumerate(linhas, start=1):
        registro = _extrair_linha_contabil_coasul(
            linha
        )

        if registro is not None:
            registros_validos.append(registro)
            continue

        linha_limpa = (
            str(linha)
            .replace("\xa0", " ")
            .replace("\r", "")
            .replace("\n", "")
            .strip()
        )

        # Registra apenas linhas que aparentam ser contábeis,
        # mas não foram reconhecidas.
        if re.match(r"^\d{10}(?:\s|$)", linha_limpa):
            linhas_contabeis_ignoradas.append({
                "Linha": numero_linha,
                "Conteúdo": linha_limpa,
            })

    if not registros_validos:
        detalhe = ""

        if linhas_contabeis_ignoradas:
            exemplos = linhas_contabeis_ignoradas[:5]

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
            f"Nenhuma linha contábil válida foi encontrada no arquivo "
            f"Coasul '{nome_arquivo}'.{detalhe}"
        )

    if linhas_contabeis_ignoradas:
        print(
            f"[COASUL] {len(linhas_contabeis_ignoradas)} linha(s) "
            f"contábil(eis) não reconhecida(s) no arquivo "
            f"'{nome_arquivo}'."
        )

        for item in linhas_contabeis_ignoradas:
            print(
                f"[COASUL] Linha {item['Linha']}: "
                f"{item['Conteúdo']}"
            )

    df_registros = pd.DataFrame(
        registros_validos
    )

    colunas_obrigatorias = [
        "Conta",
        "Nome",
        "Cód. Reduzido",
        "Saldo Anterior",
        "Débito",
        "Crédito",
        "Movimento",
        "Saldo Acumulado",
    ]

    colunas_ausentes = [
        coluna
        for coluna in colunas_obrigatorias
        if coluna not in df_registros.columns
    ]

    if colunas_ausentes:
        raise ValueError(
            f"O processamento do arquivo '{nome_arquivo}' não retornou "
            f"todas as colunas obrigatórias. Colunas ausentes: "
            f"{', '.join(colunas_ausentes)}."
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

    # Coluna E
    df_destino["Saldo Anterior"] = (
        pd.to_numeric(
            df_registros["Saldo Anterior"],
            errors="coerce"
        )
        .fillna(0.0)
        .round(2)
    )

    # Coluna F
    df_destino["Débito"] = (
        pd.to_numeric(
            df_registros["Débito"],
            errors="coerce"
        )
        .fillna(0.0)
        .abs()
        .round(2)
    )

    # Coluna G
    df_destino["Crédito"] = (
        pd.to_numeric(
            df_registros["Crédito"],
            errors="coerce"
        )
        .fillna(0.0)
        .abs()
        .mul(-1)
        .round(2)
    )

    # Coluna H
    df_destino["Movimento"] = (
        df_destino["Débito"]
        + df_destino["Crédito"]
    ).round(2)

    # Valor absoluto informado como saldo final.
    saldo_final_absoluto = (
        pd.to_numeric(
            df_registros["Saldo Acumulado"],
            errors="coerce"
        )
        .fillna(0.0)
        .abs()
        .round(2)
    )

    # Resultado da equação contábil que determinará o sinal.
    saldo_final_calculado = (
        df_destino["Saldo Anterior"]
        + df_destino["Movimento"]
    ).round(2)

    # Coluna I:
    # mantém o valor absoluto informado, mas aplica o sinal calculado.
    df_destino["Saldo Acumulado"] = saldo_final_absoluto.where(
        saldo_final_calculado >= 0,
        -saldo_final_absoluto
    ).round(2)

    # Garante a ordem definitiva das colunas.
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