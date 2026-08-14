import os
import re

try:
    from scripts.Contabilidade.tabulador_comum import (
        transformar_balancete_cooperlate,
    )
except ModuleNotFoundError:
    from tabulador_comum import transformar_balancete_cooperlate


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