import os
import re
import shutil
import subprocess
import tempfile
import unicodedata

import pandas as pd


EXTENSOES_VALIDAS_cocpell = {".xls", ".xlsx"}

COLUNAS_DESTINO_cocpell = [
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

CABECALHOS_cocpell = {
    "CONTA": "conta",
    "DESCRICAO": "nome",
    "DESCRICAO DA CONTA": "nome",
    "SALDO ANTERIOR": "saldo_anterior",
    "DEBITO": "debito",
    "CREDITO": "credito",
    "MOV PERIODO": "movimento_origem",
    "MOVIMENTO PERIODO": "movimento_origem",
    "SALDO ATUAL": "saldo_atual",
    "SALDO FINAL": "saldo_atual",
}


# ==============================================================================
# NORMALIZACAO
# ==============================================================================

def _valor_vazio_cocpell(valor):
    if valor is None:
        return True
    try:
        return bool(pd.isna(valor))
    except (TypeError, ValueError):
        return False


def _normalizar_texto_cocpell(valor):
    if _valor_vazio_cocpell(valor):
        return ""
    texto = str(valor).replace("\xa0", " ")
    return re.sub(r"\s+", " ", texto).strip()


def _normalizar_texto_comparacao_cocpell(valor):
    texto = _normalizar_texto_cocpell(valor).upper()
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", texto).strip()


def _normalizar_classificacao_cocpell(valor):
    if _valor_vazio_cocpell(valor):
        return ""

    if isinstance(valor, int):
        return str(valor)

    if isinstance(valor, float):
        if valor.is_integer():
            return str(int(valor))
        return format(valor, "f").rstrip("0").rstrip(".")

    texto = _normalizar_texto_cocpell(valor).replace(" ", "")

    if re.fullmatch(r"\d+\.0+", texto):
        return texto.split(".", maxsplit=1)[0]

    if not re.fullmatch(r"\d+(?:\.\d+)*", texto):
        return ""

    return texto


def _extrair_natureza_cocpell(valor):
    texto = _normalizar_texto_cocpell(valor).upper()
    correspondencia = re.search(r"([DC])\s*$", texto)
    return correspondencia.group(1) if correspondencia else ""


def _converter_numero_cocpell(valor):
    if _valor_vazio_cocpell(valor):
        return 0.0

    if isinstance(valor, (int, float)):
        return round(float(valor), 2)

    texto = _normalizar_texto_cocpell(valor).upper()
    if texto in {"", "-", "--"}:
        return 0.0

    texto = re.sub(r"\s*[DC]\s*$", "", texto)
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
    except (TypeError, ValueError):
        return 0.0

    if negativo_parenteses or negativo_final:
        numero = -abs(numero)

    return round(numero, 2)


def _aplicar_natureza_cocpell(valor):
    numero = _converter_numero_cocpell(valor)
    natureza = _extrair_natureza_cocpell(valor)

    if natureza == "D":
        return round(abs(numero), 2)
    if natureza == "C":
        return round(-abs(numero), 2)

    return round(numero, 2)


# ==============================================================================
# LEITURA ROBUSTA DE EXCEL
# ==============================================================================

def _erro_parece_estilo_openpyxl_cocpell(erro):
    texto = str(erro).lower()
    termos = [
        "openpyxl.styles",
        "borders.border",
        "should be <class 'openpyxl.styles",
        "styles.xml",
        "style",
        "border",
        "fill",
        "font",
    ]
    return any(termo in texto for termo in termos)


def _ler_com_engine_cocpell(caminho, sheet_name, engine):
    return pd.read_excel(
        caminho,
        sheet_name=sheet_name,
        header=None,
        dtype=object,
        engine=engine,
    )


def _normalizar_com_libreoffice_cocpell(caminho_arquivo):
    executavel = shutil.which("soffice") or shutil.which("libreoffice")

    if executavel is None:
        return None, None

    diretorio = tempfile.mkdtemp(prefix="cocpell_excel_")
    comando = [
        executavel,
        "--headless",
        "--convert-to",
        "xlsx",
        "--outdir",
        diretorio,
        caminho_arquivo,
    ]

    try:
        resultado = subprocess.run(
            comando,
            capture_output=True,
            text=True,
            check=False,
            timeout=300,
        )
    except Exception:
        shutil.rmtree(diretorio, ignore_errors=True)
        return None, None

    if resultado.returncode != 0:
        shutil.rmtree(diretorio, ignore_errors=True)
        return None, None

    nome_saida = os.path.splitext(os.path.basename(caminho_arquivo))[0] + ".xlsx"
    caminho_saida = os.path.join(diretorio, nome_saida)

    if not os.path.exists(caminho_saida):
        shutil.rmtree(diretorio, ignore_errors=True)
        return None, None

    return caminho_saida, diretorio


def _ler_excel_robusto_cocpell(caminho_arquivo, sheet_name=0):
    """
    Ordem de leitura:
        XLS  -> xlrd, depois calamine.
        XLSX -> openpyxl, depois calamine, depois LibreOffice.
    """
    nome = os.path.basename(caminho_arquivo)
    extensao = os.path.splitext(caminho_arquivo)[1].lower()
    erros = []

    if extensao == ".xls":
        engines = ["xlrd", "calamine"]
    elif extensao == ".xlsx":
        engines = ["openpyxl", "calamine"]
    else:
        raise ValueError(f"O arquivo '{nome}' nao e um Excel valido.")

    for engine in engines:
        try:
            return _ler_com_engine_cocpell(caminho_arquivo, sheet_name, engine)
        except ImportError as erro:
            erros.append(f"{engine}: dependencia ausente ({erro})")
        except Exception as erro:
            erros.append(f"{engine}: {erro}")

            if engine == "openpyxl" and not _erro_parece_estilo_openpyxl_cocpell(erro):
                # Ainda tenta calamine, mas preserva o diagnostico original.
                pass

    caminho_normalizado, diretorio = _normalizar_com_libreoffice_cocpell(
        caminho_arquivo
    )

    if caminho_normalizado:
        try:
            return _ler_com_engine_cocpell(
                caminho_normalizado,
                sheet_name,
                "openpyxl",
            )
        except Exception as erro:
            erros.append(f"LibreOffice/openpyxl: {erro}")
        finally:
            shutil.rmtree(diretorio, ignore_errors=True)

    detalhes = " | ".join(erros)
    raise ValueError(
        f"Nao foi possivel ler o arquivo Cocpell '{nome}'. "
        "O arquivo pode possuir estilos internos invalidos. "
        "Instale o fallback com 'python -m pip install python-calamine' "
        "ou abra e salve uma nova copia no Microsoft Excel. "
        f"Detalhes: {detalhes}"
    )


# ==============================================================================
# CABECALHO E MAPEAMENTO
# ==============================================================================

def _mapear_cabecalho_cocpell(linha):
    mapeamento = {}

    for indice, valor in enumerate(linha.tolist()):
        texto = _normalizar_texto_comparacao_cocpell(valor)
        texto = re.sub(r"\s+", " ", texto)

        for titulo, campo in CABECALHOS_cocpell.items():
            if texto == titulo:
                mapeamento[campo] = indice
                break

    return mapeamento


def _localizar_cabecalho_cocpell(dataframe, nome_arquivo):
    campos_obrigatorios = {
        "conta",
        "nome",
        "saldo_anterior",
        "debito",
        "credito",
        "saldo_atual",
    }

    for indice in range(dataframe.shape[0]):
        mapeamento = _mapear_cabecalho_cocpell(dataframe.iloc[indice])

        if campos_obrigatorios.issubset(mapeamento):
            return indice, mapeamento

    raise ValueError(
        f"Nao foi possivel localizar o cabecalho do balancete Cocpell "
        f"no arquivo '{nome_arquivo}'."
    )


def _extrair_registro_cocpell(linha, colunas):
    classificacao = _normalizar_classificacao_cocpell(
        linha.iloc[colunas["conta"]]
    )
    nome = _normalizar_texto_cocpell(linha.iloc[colunas["nome"]])

    if not classificacao or not nome:
        return None

    saldo_anterior = _aplicar_natureza_cocpell(
        linha.iloc[colunas["saldo_anterior"]]
    )
    debito = round(
        abs(_converter_numero_cocpell(linha.iloc[colunas["debito"]])),
        2,
    )
    credito = round(
        -abs(_converter_numero_cocpell(linha.iloc[colunas["credito"]])),
        2,
    )
    movimento = round(debito + credito, 2)
    saldo_acumulado = _aplicar_natureza_cocpell(
        linha.iloc[colunas["saldo_atual"]]
    )

    return {
        "Atividade": "Geral",
        "Conta": classificacao,
        "Nome": nome,
        "Cód. Reduzido": classificacao,
        "Saldo Anterior": saldo_anterior,
        "Débito": debito,
        "Crédito": credito,
        "Movimento": movimento,
        "Saldo Acumulado": saldo_acumulado,
    }


def _montar_dataframe_cocpell(registros):
    if not registros:
        return pd.DataFrame(columns=COLUNAS_DESTINO_cocpell)

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

    return dataframe[COLUNAS_DESTINO_cocpell].copy()


# ==============================================================================
# TRANSFORMACAO E FUNCAO PUBLICA
# ==============================================================================

def transformar_balancete_cocpell(caminho_arquivo):
    nome_arquivo = os.path.basename(caminho_arquivo)
    dataframe_origem = _ler_excel_robusto_cocpell(caminho_arquivo, sheet_name=0)
    indice_cabecalho, colunas = _localizar_cabecalho_cocpell(
        dataframe_origem,
        nome_arquivo,
    )

    registros = []

    for _, linha in dataframe_origem.iloc[indice_cabecalho + 1:].iterrows():
        # Ignora cabecalhos repetidos no corpo do relatorio.
        if _mapear_cabecalho_cocpell(linha):
            continue

        registro = _extrair_registro_cocpell(linha, colunas)
        if registro is not None:
            registros.append(registro)

    if not registros:
        raise ValueError(
            f"Nenhuma conta valida foi encontrada no arquivo Cocpell "
            f"'{nome_arquivo}'."
        )

    return _montar_dataframe_cocpell(registros)


def _obter_mes_cocpell(caminho_arquivo):
    nome = os.path.basename(caminho_arquivo)
    correspondencia = re.match(r"^B_(0[1-9]|1[0-2])", nome, re.IGNORECASE)
    return correspondencia.group(1) if correspondencia else None


def _limpar_nome_aba_cocpell(nome):
    nome = re.sub(r'[\\/\x2a?:\[\]]', "_", str(nome).strip())
    return (nome or "Sem nome")[:31]


def _gerar_nome_aba_cocpell(caminho_arquivo, nomes_usados):
    mes = _obter_mes_cocpell(caminho_arquivo)
    nome_base = mes or os.path.splitext(os.path.basename(caminho_arquivo))[0]
    nome_base = _limpar_nome_aba_cocpell(nome_base)
    candidato = nome_base
    contador = 2

    while candidato.casefold() in nomes_usados:
        sufixo = "_" + str(contador)
        candidato = nome_base[:31 - len(sufixo)] + sufixo
        contador += 1

    nomes_usados.add(candidato.casefold())
    return candidato


def _chave_ordenacao_cocpell(caminho_arquivo):
    mes = _obter_mes_cocpell(caminho_arquivo)
    return (
        int(mes) if mes else 99,
        os.path.basename(caminho_arquivo).casefold(),
    )


def processar(lista_arquivos):
    if not lista_arquivos:
        raise ValueError(
            "Nenhum arquivo foi selecionado para o sistema cocpell."
        )

    arquivos_excel = [
        arquivo
        for arquivo in lista_arquivos
        if os.path.splitext(str(arquivo))[1].lower()
        in EXTENSOES_VALIDAS_cocpell
    ]

    if not arquivos_excel:
        raise ValueError(
            "Nenhum arquivo Excel valido foi encontrado para o cocpell."
        )

    resultados = {}
    nomes_usados = set()

    for arquivo in sorted(arquivos_excel, key=_chave_ordenacao_cocpell):
        nome_aba = _gerar_nome_aba_cocpell(arquivo, nomes_usados)
        resultados[nome_aba] = transformar_balancete_cocpell(arquivo)

    return resultados
