import os
import re
import shutil
import subprocess
import sys
import tempfile
import unicodedata
from collections import OrderedDict, defaultdict

import pandas as pd
from openpyxl import Workbook


# ==============================================================================
# CONFIGURACOES DO SISTEMA soli3
# ==============================================================================

EXTENSAO_VALIDA_soli3 = ".pdf"

COLUNAS_DESTINO_soli3 = [
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

COLUNAS_EXCEL_INTERMEDIARIO_soli3 = [
    "Página",
    "Linha",
    "Texto Extraído",
    "Conta PDF",
    "Cód. Reduzido PDF",
    "Nome PDF",
    "Saldo Anterior PDF",
    "Débito PDF",
    "Crédito PDF",
    "Saldo Final PDF",
    "Conta",
    "Cód. Reduzido",
    "Nome",
    "Saldo Anterior Bruto",
    "Natureza Anterior",
    "Débito Bruto",
    "Crédito Bruto",
    "Saldo Final Bruto",
    "Natureza Final",
]

# Parametros de analise de leiaute do pdfminer.six.
LA_PARAMS_soli3 = {
    "line_overlap": 0.50,
    "char_margin": 2.00,
    "line_margin": 0.40,
    "word_margin": 0.10,
    "boxes_flow": None,
    "detect_vertical": False,
    "all_texts": True,
}

# Limites proporcionais das colunas no leiaute fixo do balancete.
# Os valores foram definidos como percentuais da largura da pagina para
# manter o funcionamento quando o PDF for redimensionado ou passar por OCR.
LIMITES_COLUNAS_PDF_soli3 = OrderedDict([
    ("Conta PDF", (0.125, 0.260)),
    ("Cód. Reduzido PDF", (0.260, 0.335)),
    ("Nome PDF", (0.335, 0.558)),
    ("Saldo Anterior PDF", (0.558, 0.663)),
    ("Débito PDF", (0.663, 0.767)),
    ("Crédito PDF", (0.767, 0.877)),
    ("Saldo Final PDF", (0.877, 1.010)),
])


# ==============================================================================
# NOMES DAS ABAS
# ==============================================================================

def _limpar_nome_aba_soli3(nome):
    """Ajusta um texto para utilizacao como nome de aba do Excel."""
    nome_limpo = re.sub(
        r'[\\/\x2a?:\[\]]',
        "_",
        str(nome).strip(),
    )
    return (nome_limpo or "Sem nome")[:31]


def _obter_nome_sem_extensao_soli3(caminho_arquivo):
    nome_arquivo = os.path.basename(caminho_arquivo)
    return os.path.splitext(nome_arquivo)[0].strip()


def _obter_mes_soli3(caminho_arquivo):
    nome_arquivo = os.path.basename(caminho_arquivo)
    correspondencia = re.match(
        r"^B_(0[1-9]|1[0-2])",
        nome_arquivo,
        re.IGNORECASE,
    )
    return correspondencia.group(1) if correspondencia else None


def _registrar_nome_aba_soli3(nome, nomes_utilizados):
    nome_aba = _limpar_nome_aba_soli3(nome)
    chave = nome_aba.casefold()

    if chave in nomes_utilizados:
        return None

    nomes_utilizados.add(chave)
    return nome_aba


def _gerar_nome_aba_soli3(caminho_arquivo, nomes_utilizados):
    nome_completo = _obter_nome_sem_extensao_soli3(caminho_arquivo)
    mes = _obter_mes_soli3(caminho_arquivo)

    if mes:
        nome_aba = _registrar_nome_aba_soli3(mes, nomes_utilizados)
        if nome_aba is not None:
            return nome_aba

    nome_aba = _registrar_nome_aba_soli3(
        nome_completo,
        nomes_utilizados,
    )
    if nome_aba is not None:
        return nome_aba

    contador = 2
    while True:
        sufixo = "_" + str(contador)
        candidato = nome_completo[:31 - len(sufixo)] + sufixo
        nome_aba = _registrar_nome_aba_soli3(
            candidato,
            nomes_utilizados,
        )
        if nome_aba is not None:
            return nome_aba
        contador += 1


# ==============================================================================
# NORMALIZACAO
# ==============================================================================

def _normalizar_texto_soli3(valor):
    if valor is None:
        return ""

    try:
        if pd.isna(valor):
            return ""
    except (TypeError, ValueError):
        pass

    texto = str(valor).replace("\xa0", " ")
    return re.sub(r"\s+", " ", texto).strip()


def _normalizar_texto_comparacao_soli3(valor):
    texto = _normalizar_texto_soli3(valor).upper()
    texto = unicodedata.normalize("NFKD", texto)
    return "".join(
        caractere
        for caractere in texto
        if not unicodedata.combining(caractere)
    )


def _normalizar_classificacao_soli3(valor):
    texto = (
        _normalizar_texto_soli3(valor)
        .replace(" ", "")
        .replace(",", ".")
    )

    if not re.fullmatch(r"\d+(?:\.\d+)*", texto):
        return ""

    return texto


def _normalizar_natureza_soli3(natureza):
    texto = _normalizar_texto_comparacao_soli3(natureza)

    if texto in {"D", "DEBITO", "DEVEDOR"}:
        return "D"
    if texto in {"C", "CREDITO", "CREDOR"}:
        return "C"

    return ""


def _converter_numero_soli3(valor):
    if valor is None:
        return 0.0

    try:
        if pd.isna(valor):
            return 0.0
    except (TypeError, ValueError):
        pass

    if isinstance(valor, (int, float)):
        return round(float(valor), 2)

    texto = str(valor).strip()
    if texto in {"", "-", "--"}:
        return 0.0

    texto = (
        texto.replace("\xa0", "")
        .replace(" ", "")
        .replace("R$", "")
        .replace("$", "")
    )

    negativo_parenteses = texto.startswith("(") and texto.endswith(")")
    negativo_final = texto.endswith("-") and texto != "-"

    if negativo_parenteses:
        texto = texto[1:-1]
    if negativo_final:
        texto = texto[:-1]

    if "," in texto:
        texto = texto.replace(".", "").replace(",", ".")

    try:
        numero = float(texto)
    except (ValueError, TypeError):
        return 0.0

    if negativo_parenteses or negativo_final:
        numero = -abs(numero)

    return round(numero, 2)


def _aplicar_natureza_soli3(valor, natureza):
    numero = _converter_numero_soli3(valor)
    natureza_normalizada = _normalizar_natureza_soli3(natureza)

    if natureza_normalizada == "D":
        return round(abs(numero), 2)
    if natureza_normalizada == "C":
        return round(-abs(numero), 2)

    return round(numero, 2)


# ==============================================================================
# OCR AUTOMATICO
# ==============================================================================

def _obter_comando_ocr_soli3():
    """Localiza OCRmyPDF no PATH ou nos interpretadores Python comuns."""
    candidatos = []
    executavel_ocr = shutil.which("ocrmypdf")

    if executavel_ocr:
        candidatos.append([executavel_ocr])

    if sys.executable and not getattr(sys, "frozen", False):
        candidatos.append([sys.executable, "-m", "ocrmypdf"])

    executavel_py = shutil.which("py")
    if executavel_py:
        candidatos.append([executavel_py, "-m", "ocrmypdf"])

    executavel_python = shutil.which("python")
    if executavel_python:
        candidatos.append([executavel_python, "-m", "ocrmypdf"])

    comandos_testados = set()

    for comando in candidatos:
        chave = tuple(comando)
        if chave in comandos_testados:
            continue
        comandos_testados.add(chave)

        try:
            resultado = subprocess.run(
                comando + ["--version"],
                capture_output=True,
                text=True,
                check=False,
                timeout=20,
            )
        except (OSError, subprocess.TimeoutExpired):
            continue

        if resultado.returncode == 0:
            return comando

    return None


def _executar_ocr_soli3(caminho_arquivo, caminho_saida):
    """Cria um PDF pesquisavel temporario utilizando OCRmyPDF."""
    nome_arquivo = os.path.basename(caminho_arquivo)
    comando_base = _obter_comando_ocr_soli3()

    if comando_base is None:
        raise ValueError(
            f"O PDF '{nome_arquivo}' nao possui texto pesquisavel "
            "suficiente e o OCRmyPDF nao foi localizado."
        )

    comando = comando_base + [
        "--language", "por",
        "--rotate-pages",
        "--deskew",
        "--skip-text",
        "--output-type", "pdf",
        "--optimize", "1",
        "--jobs", "2",
        caminho_arquivo,
        caminho_saida,
    ]

    try:
        resultado = subprocess.run(
            comando,
            capture_output=True,
            text=True,
            check=False,
            timeout=1800,
        )
    except subprocess.TimeoutExpired as erro:
        raise ValueError(
            f"O OCR de '{nome_arquivo}' ultrapassou 30 minutos."
        ) from erro
    except OSError as erro:
        raise ValueError(
            f"Nao foi possivel iniciar o OCR de '{nome_arquivo}'. "
            f"Erro: {erro}"
        ) from erro

    if resultado.returncode != 0:
        mensagem = (
            resultado.stderr.strip()
            or resultado.stdout.strip()
            or "Erro nao identificado."
        )
        raise ValueError(
            f"Falha no OCR de '{nome_arquivo}'. Detalhes: {mensagem}"
        )

    if not os.path.exists(caminho_saida) or os.path.getsize(caminho_saida) == 0:
        raise ValueError(
            f"O OCR de '{nome_arquivo}' nao gerou um PDF valido."
        )

    return caminho_saida


# ==============================================================================
# PDFMINER.SIX - EXTRACAO POSICIONAL
# ==============================================================================

def _carregar_pdfminer_soli3():
    """Importa pdfminer.six e apresenta mensagem clara se estiver ausente."""
    try:
        from pdfminer.high_level import extract_pages
        from pdfminer.layout import LAParams, LTChar, LTTextContainer
    except ImportError as erro:
        raise ImportError(
            "Instale pdfminer.six para processar o soli3: "
            "python -m pip install pdfminer.six"
        ) from erro

    return extract_pages, LAParams, LTChar, LTTextContainer


def _extrair_caracteres_elemento_soli3(elemento, classe_char):
    """Percorre recursivamente um objeto de leiaute e retorna LTChar."""
    caracteres = []

    if isinstance(elemento, classe_char):
        return [elemento]

    if hasattr(elemento, "__iter__"):
        for filho in elemento:
            caracteres.extend(
                _extrair_caracteres_elemento_soli3(filho, classe_char)
            )

    return caracteres


def _agrupar_caracteres_em_linhas_soli3(caracteres, tolerancia_y=2.5):
    """Agrupa caracteres visualmente alinhados na mesma linha."""
    grupos = []

    for caractere in sorted(
        caracteres,
        key=lambda item: (-item.y0, item.x0),
    ):
        centro_y = (caractere.y0 + caractere.y1) / 2
        grupo_encontrado = None

        for grupo in grupos:
            if abs(grupo["centro_y"] - centro_y) <= tolerancia_y:
                grupo_encontrado = grupo
                break

        if grupo_encontrado is None:
            grupo_encontrado = {
                "centro_y": centro_y,
                "caracteres": [],
            }
            grupos.append(grupo_encontrado)

        grupo_encontrado["caracteres"].append(caractere)
        quantidade = len(grupo_encontrado["caracteres"])
        grupo_encontrado["centro_y"] = (
            grupo_encontrado["centro_y"] * (quantidade - 1) + centro_y
        ) / quantidade

    return sorted(grupos, key=lambda item: -item["centro_y"])


def _montar_texto_linha_soli3(caracteres):
    """Reconstroi a linha inserindo espacos de acordo com as coordenadas."""
    caracteres = sorted(caracteres, key=lambda item: item.x0)

    if not caracteres:
        return ""

    resultado = []
    x_anterior = None
    largura_media = []

    for caractere in caracteres:
        texto = caractere.get_text()

        if not texto or texto in {"\n", "\r"}:
            continue

        largura = max(caractere.x1 - caractere.x0, 0.1)
        largura_media.append(largura)
        media = sum(largura_media[-20:]) / len(largura_media[-20:])

        if x_anterior is not None:
            distancia = caractere.x0 - x_anterior

            if distancia > media * 0.85:
                quantidade_espacos = max(1, min(8, round(distancia / media)))
                resultado.append(" " * quantidade_espacos)

        resultado.append(texto)
        x_anterior = caractere.x1

    return "".join(resultado).strip()


def _segmentar_colunas_fixadas_soli3(caracteres, largura_pagina):
    """Separa uma linha visual nas colunas fixas do relatorio."""
    celulas = {nome: [] for nome in LIMITES_COLUNAS_PDF_soli3}

    for caractere in caracteres:
        texto = caractere.get_text()
        if not texto or texto in {"\n", "\r"}:
            continue

        centro_x = (caractere.x0 + caractere.x1) / 2
        proporcao_x = centro_x / max(largura_pagina, 1)

        for nome_coluna, (inicio, fim) in LIMITES_COLUNAS_PDF_soli3.items():
            if inicio <= proporcao_x < fim:
                celulas[nome_coluna].append(caractere)
                break

    resultado = {}
    for nome_coluna, caracteres_celula in celulas.items():
        resultado[nome_coluna] = _normalizar_texto_soli3(
            _montar_texto_linha_soli3(caracteres_celula)
        )

    return resultado


def _extrair_linhas_pdfminer_soli3(caminho_pdf):
    """Extrai linhas com pagina e ordem visual utilizando pdfminer.six."""
    extract_pages, LAParams, LTChar, _ = _carregar_pdfminer_soli3()
    parametros = LAParams(**LA_PARAMS_soli3)
    linhas_extraidas = []

    for numero_pagina, pagina in enumerate(
        extract_pages(caminho_pdf, laparams=parametros),
        start=1,
    ):
        caracteres = []

        for elemento in pagina:
            caracteres.extend(
                _extrair_caracteres_elemento_soli3(elemento, LTChar)
            )

        grupos = _agrupar_caracteres_em_linhas_soli3(caracteres)

        largura_pagina = float(getattr(pagina, "width", pagina.x1))

        for numero_linha, grupo in enumerate(grupos, start=1):
            texto = _montar_texto_linha_soli3(grupo["caracteres"])

            if texto:
                celulas = _segmentar_colunas_fixadas_soli3(
                    grupo["caracteres"],
                    largura_pagina,
                )
                item = {
                    "Página": numero_pagina,
                    "Linha": numero_linha,
                    "Texto Extraído": texto,
                }
                item.update(celulas)
                linhas_extraidas.append(item)

    return linhas_extraidas


def _pdf_possui_texto_suficiente_soli3(linhas_extraidas):
    """Verifica se a extracao direta possui texto util suficiente."""
    quantidade = sum(
        len(_normalizar_texto_soli3(item.get("Texto Extraído", "")))
        for item in linhas_extraidas
    )
    return quantidade >= 20


def _extrair_linhas_com_ocr_soli3(caminho_arquivo):
    """Extrai com pdfminer e aciona OCR quando o PDF for digitalizado."""
    linhas = _extrair_linhas_pdfminer_soli3(caminho_arquivo)

    if _pdf_possui_texto_suficiente_soli3(linhas):
        return linhas

    diretorio_temporario = tempfile.mkdtemp(prefix="soli3_ocr_")
    caminho_ocr = os.path.join(diretorio_temporario, "documento_ocr.pdf")

    try:
        _executar_ocr_soli3(caminho_arquivo, caminho_ocr)
        linhas_ocr = _extrair_linhas_pdfminer_soli3(caminho_ocr)

        if not _pdf_possui_texto_suficiente_soli3(linhas_ocr):
            raise ValueError(
                "O OCR foi realizado, mas nao produziu texto suficiente."
            )

        return linhas_ocr
    finally:
        shutil.rmtree(diretorio_temporario, ignore_errors=True)


# ==============================================================================
# INTERPRETACAO DAS LINHAS CONTABEIS
# ==============================================================================

def _linha_eh_cabecalho_soli3(linha):
    texto = _normalizar_texto_comparacao_soli3(linha)

    if not texto:
        return True

    if re.fullmatch(r"[-_= ]+", texto):
        return True

    termos = [
        "BALANCETE ANALITICO",
        "SITUACAO EM",
        "CONTA CONTABIL",
        "SALDO ANTERIOR",
        "DEBITOS",
        "CREDITOS",
        "SALDO ATUAL",
        "HORA:",
        "PAGINA",
    ]

    return sum(termo in texto for termo in termos) >= 2


def _padrao_linha_contabil_soli3():
    numero = r"[+-]?(?:\d{1,3}(?:\.\d{3})*|\d+),\d{2}"

    return re.compile(
        r"^\s*"
        r"(?P<conta>\d+(?:\.\d+)*)"
        r"\s+"
        r"(?P<reduzido>\d+(?:\.\d+)*)"
        r"\s+"
        r"(?P<nome>.+?)"
        r"\s+"
        r"(?P<saldo_anterior>" + numero + r")"
        r"\s*(?P<natureza_anterior>[DCdc])"
        r"\s+"
        r"(?P<debito>" + numero + r")"
        r"\s+"
        r"(?P<credito>" + numero + r")"
        r"\s+"
        r"(?P<saldo_final>" + numero + r")"
        r"\s*(?P<natureza_final>[DCdc])"
        r"\s*$"
    )


def _interpretar_linha_contabil_soli3(texto):
    """Separa os campos de uma linha contabil completa."""
    if _linha_eh_cabecalho_soli3(texto):
        return None

    correspondencia = _padrao_linha_contabil_soli3().match(
        _normalizar_texto_soli3(texto)
    )

    if correspondencia is None:
        return None

    dados = correspondencia.groupdict()

    return {
        "Conta": _normalizar_classificacao_soli3(dados["conta"]),
        "Cód. Reduzido": _normalizar_classificacao_soli3(
            dados["reduzido"]
        ),
        "Nome": _normalizar_texto_soli3(dados["nome"]),
        "Saldo Anterior Bruto": dados["saldo_anterior"],
        "Natureza Anterior": dados["natureza_anterior"].upper(),
        "Débito Bruto": dados["debito"],
        "Crédito Bruto": dados["credito"],
        "Saldo Final Bruto": dados["saldo_final"],
        "Natureza Final": dados["natureza_final"].upper(),
    }


def _separar_valor_natureza_celula_soli3(texto):
    """Separa valor monetario e natureza D/C de uma celula fixa."""
    texto = _normalizar_texto_soli3(texto)
    if not texto:
        return "", ""

    correspondencia = re.search(
        r"(?P<valor>[+-]?(?:(?:\d{1,3}(?:\.\d{3})*)|\d+)?[,]\d{2})"
        r"\s*(?P<natureza>[DCdc])?",
        texto,
    )
    if correspondencia is None:
        return "", ""

    valor = correspondencia.group("valor")
    if valor.startswith(","):
        valor = "0" + valor
    elif valor.startswith("-,"):
        valor = "-0" + valor[1:]

    natureza = (correspondencia.group("natureza") or "").upper()
    return valor, natureza


def _interpretar_celulas_fixadas_soli3(item):
    """Interpreta cada coluna do PDF sem depender dos espacos da linha."""
    conta = _normalizar_classificacao_soli3(item.get("Conta PDF", ""))
    reduzido = _normalizar_classificacao_soli3(
        item.get("Cód. Reduzido PDF", "")
    )
    nome = _normalizar_texto_soli3(item.get("Nome PDF", ""))

    if not conta or not reduzido or not nome:
        return None

    saldo_anterior, natureza_anterior = _separar_valor_natureza_celula_soli3(
        item.get("Saldo Anterior PDF", "")
    )
    debito, _ = _separar_valor_natureza_celula_soli3(
        item.get("Débito PDF", "")
    )
    credito, _ = _separar_valor_natureza_celula_soli3(
        item.get("Crédito PDF", "")
    )
    saldo_final, natureza_final = _separar_valor_natureza_celula_soli3(
        item.get("Saldo Final PDF", "")
    )

    # Ao menos um campo monetario deve existir para a linha ser contabil.
    if not any([saldo_anterior, debito, credito, saldo_final]):
        return None

    return {
        "Conta": conta,
        "Cód. Reduzido": reduzido,
        "Nome": nome,
        "Saldo Anterior Bruto": saldo_anterior,
        "Natureza Anterior": natureza_anterior,
        "Débito Bruto": debito,
        "Crédito Bruto": credito,
        "Saldo Final Bruto": saldo_final,
        "Natureza Final": natureza_final,
    }


def _combinar_linhas_quebradas_soli3(linhas_extraidas):
    """Combina ate tres linhas consecutivas da mesma pagina."""
    resultado = []
    por_pagina = defaultdict(list)

    for item in linhas_extraidas:
        por_pagina[item["Página"]].append(item)

    for pagina in sorted(por_pagina):
        linhas = sorted(por_pagina[pagina], key=lambda item: item["Linha"])
        indice = 0

        while indice < len(linhas):
            item_atual = linhas[indice]
            texto_atual = item_atual["Texto Extraído"]

            if (
                _interpretar_celulas_fixadas_soli3(item_atual)
                or _interpretar_linha_contabil_soli3(texto_atual)
            ):
                resultado.append(item_atual)
                indice += 1
                continue

            combinado = None
            quantidade_usada = 1

            for quantidade in (2, 3):
                if indice + quantidade > len(linhas):
                    continue

                candidato = " ".join(
                    _normalizar_texto_soli3(
                        linhas[posicao]["Texto Extraído"]
                    )
                    for posicao in range(indice, indice + quantidade)
                )

                if _interpretar_linha_contabil_soli3(candidato):
                    combinado = {
                        "Página": pagina,
                        "Linha": item_atual["Linha"],
                        "Texto Extraído": candidato,
                    }
                    quantidade_usada = quantidade
                    break

            resultado.append(combinado or item_atual)
            indice += quantidade_usada

    return resultado

def _conta_eh_resultado_soli3(classificacao):
    """
    Identifica contas de resultado pelo primeiro algarismo
    da classificação contábil.

    Regras:
        1 = Ativo;
        2 = Passivo;
        3 = Passivo;
        4 ou superior = Resultado.
    """
    classificacao = _normalizar_classificacao_soli3(
        classificacao
    )

    if not classificacao:
        return False

    primeiro_algarismo = classificacao[0]

    if not primeiro_algarismo.isdigit():
        return False

    return int(primeiro_algarismo) >= 4


# ==============================================================================
# CONVERSAO PDF -> EXCEL INTERMEDIARIO
# ==============================================================================

def _converter_pdf_para_excel_soli3(caminho_pdf, caminho_excel):
    """
    Converte o PDF em uma planilha intermediaria estruturada.

    A planilha contem todas as linhas extraidas e, quando a linha for
    contabil, tambem contem os campos separados. Essa etapa permite que a
    tabulacao leia uma estrutura Excel em vez de interpretar o PDF duas vezes.
    """
    linhas = _extrair_linhas_com_ocr_soli3(caminho_pdf)
    linhas = _combinar_linhas_quebradas_soli3(linhas)

    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "PDF Extraído"
    worksheet.append(COLUNAS_EXCEL_INTERMEDIARIO_soli3)

    quantidade_registros = 0

    for item in linhas:
        interpretado = _interpretar_celulas_fixadas_soli3(item)

        if interpretado is None:
            interpretado = _interpretar_linha_contabil_soli3(
                item["Texto Extraído"]
            )

        linha_excel = {
            "Página": item["Página"],
            "Linha": item["Linha"],
            "Texto Extraído": item["Texto Extraído"],
            "Conta PDF": item.get("Conta PDF", ""),
            "Cód. Reduzido PDF": item.get("Cód. Reduzido PDF", ""),
            "Nome PDF": item.get("Nome PDF", ""),
            "Saldo Anterior PDF": item.get("Saldo Anterior PDF", ""),
            "Débito PDF": item.get("Débito PDF", ""),
            "Crédito PDF": item.get("Crédito PDF", ""),
            "Saldo Final PDF": item.get("Saldo Final PDF", ""),
            "Conta": "",
            "Cód. Reduzido": "",
            "Nome": "",
            "Saldo Anterior Bruto": "",
            "Natureza Anterior": "",
            "Débito Bruto": "",
            "Crédito Bruto": "",
            "Saldo Final Bruto": "",
            "Natureza Final": "",
        }

        if interpretado is not None:
            linha_excel.update(interpretado)
            quantidade_registros += 1

        worksheet.append([
            linha_excel[coluna]
            for coluna in COLUNAS_EXCEL_INTERMEDIARIO_soli3
        ])

    workbook.save(caminho_excel)

    if quantidade_registros == 0:
        raise ValueError(
            "O PDF foi convertido para Excel, mas nenhuma linha contabil "
            "compativel com o layout do soli3 foi identificada."
        )

    return caminho_excel


# ==============================================================================
# EXTRACAO DO EXCEL INTERMEDIARIO E TABULACAO
# ==============================================================================

def _ler_excel_intermediario_soli3(caminho_excel):
    """Le a planilha gerada na etapa PDF para Excel."""
    dataframe = pd.read_excel(
        caminho_excel,
        sheet_name="PDF Extraído",
        dtype=object,
        engine="openpyxl",
    )

    if dataframe.empty:
        raise ValueError("A planilha intermediaria do soli3 esta vazia.")

    return dataframe


def _extrair_registros_excel_soli3(dataframe):
    """Converte os campos brutos da planilha intermediaria em registros."""
    registros = []

    for _, linha in dataframe.iterrows():
        classificacao = _normalizar_classificacao_soli3(
            linha.get("Conta", "")
        )
        codigo_reduzido = _normalizar_classificacao_soli3(
            linha.get("Cód. Reduzido", "")
        )
        nome = _normalizar_texto_soli3(linha.get("Nome", ""))

        if not classificacao or not codigo_reduzido or not nome:
            continue

        saldo_anterior = _aplicar_natureza_soli3(
            linha.get("Saldo Anterior Bruto", ""),
            linha.get("Natureza Anterior", ""),
        )
        debito = round(
            abs(_converter_numero_soli3(linha.get("Débito Bruto", ""))),
            2,
        )
        credito = round(
            -abs(_converter_numero_soli3(linha.get("Crédito Bruto", ""))),
            2,
        )
        saldo_acumulado = _aplicar_natureza_soli3(
            linha.get("Saldo Final Bruto", ""),
            linha.get("Natureza Final", ""),
        )
        movimento = round(debito + credito, 2)

        registros.append({
            "Atividade": "Geral",
            "Conta": classificacao,
            "Nome": nome,
            "Cód. Reduzido": codigo_reduzido,
            "Saldo Anterior": saldo_anterior,
            "Débito": debito,
            "Crédito": credito,
            "Movimento": movimento,
            "Saldo Acumulado": saldo_acumulado,
        })

    return registros


def _montar_dataframe_soli3(registros):
    if not registros:
        return pd.DataFrame(columns=COLUNAS_DESTINO_soli3)

    dataframe = pd.DataFrame(registros)

    for coluna in [
        "Saldo Anterior",
        "Débito",
        "Crédito",
        "Movimento",
        "Saldo Acumulado",
    ]:
        dataframe[coluna] = pd.to_numeric(
            dataframe[coluna],
            errors="coerce",
        ).fillna(0.0).round(2)

    dataframe["Débito"] = dataframe["Débito"].abs().round(2)
    dataframe["Crédito"] = (-dataframe["Crédito"].abs()).round(2)
    dataframe["Movimento"] = (
        dataframe["Débito"] + dataframe["Crédito"]
    ).round(2)

    return dataframe[COLUNAS_DESTINO_soli3].copy()


def transformar_balancete_soli3(caminho_arquivo):
    """
    Executa o fluxo PDF -> Excel temporario -> DataFrame tabulado.
    """
    nome_arquivo = os.path.basename(caminho_arquivo)
    extensao = os.path.splitext(caminho_arquivo)[1].lower()

    if extensao != EXTENSAO_VALIDA_soli3:
        raise ValueError(
            f"O arquivo '{nome_arquivo}' nao possui extensao PDF."
        )

    diretorio_temporario = tempfile.mkdtemp(prefix="soli3_excel_")
    caminho_excel = os.path.join(
        diretorio_temporario,
        "pdf_convertido.xlsx",
    )

    try:
        _converter_pdf_para_excel_soli3(
            caminho_arquivo,
            caminho_excel,
        )
        dataframe_intermediario = _ler_excel_intermediario_soli3(
            caminho_excel
        )
        registros = _extrair_registros_excel_soli3(
            dataframe_intermediario
        )

        if not registros:
            raise ValueError(
                f"Nenhuma conta valida foi extraida do PDF soli3 "
                f"'{nome_arquivo}' apos a conversao para Excel."
            )

        return _montar_dataframe_soli3(registros)
    finally:
        shutil.rmtree(diretorio_temporario, ignore_errors=True)


# ==============================================================================
# FUNCAO PUBLICA
# ==============================================================================

def _chave_ordenacao_arquivo_soli3(caminho_arquivo):
    mes = _obter_mes_soli3(caminho_arquivo)
    numero_mes = int(mes) if mes else 99
    return numero_mes, os.path.basename(caminho_arquivo).casefold()


def processar(lista_arquivos):
    """Processa todos os PDFs selecionados para o soli3."""
    if not lista_arquivos:
        raise ValueError(
            "Nenhum arquivo foi selecionado para o sistema soli3."
        )

    arquivos_pdf = [
        arquivo
        for arquivo in lista_arquivos
        if os.path.splitext(str(arquivo))[1].lower()
        == EXTENSAO_VALIDA_soli3
    ]

    if not arquivos_pdf:
        raise ValueError(
            "Nenhum arquivo PDF valido foi encontrado para o soli3."
        )

    resultados = {}
    nomes_utilizados = set()

    for arquivo in sorted(
        arquivos_pdf,
        key=_chave_ordenacao_arquivo_soli3,
    ):
        nome_aba = _gerar_nome_aba_soli3(
            arquivo,
            nomes_utilizados,
        )
        dataframe = transformar_balancete_soli3(arquivo)

        if dataframe is None or dataframe.empty:
            raise ValueError(
                f"O arquivo '{os.path.basename(arquivo)}' nao retornou "
                "dados validos."
            )

        resultados[nome_aba] = dataframe

    if not resultados:
        raise ValueError(
            "Nenhum resultado foi gerado para o sistema soli3."
        )

    return resultados
