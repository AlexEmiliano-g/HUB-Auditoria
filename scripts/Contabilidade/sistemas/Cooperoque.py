import os
import re

try:
    from scripts.Contabilidade.tabulador_comum import (
        transformar_balancete_cooperoque
    )
except ModuleNotFoundError:
    from tabulador_comum import transformar_balancete_cooperoque


def _limpar_nome_aba(nome):
    """
    Ajusta um texto para ser utilizado como nome de aba do Excel.

    O Excel não permite os caracteres:
        \ / ? * [ ] :

    O nome também é limitado a 31 caracteres.
    """
    nome = str(nome).strip()

    nome_limpo = re.sub(
        r'[\\/*?:\[\]]',
        "_",
        nome
    )

    if not nome_limpo:
        nome_limpo = "Sem nome"

    return nome_limpo[:31]


def _obter_nome_completo(caminho_arquivo):
    """Retorna o nome completo do arquivo sem a extensão."""
    nome_arquivo = os.path.basename(caminho_arquivo)
    nome_sem_extensao = os.path.splitext(nome_arquivo)[0]

    return nome_sem_extensao.strip()


def _obter_mes(caminho_arquivo):
    """
    Extrai o período quando o arquivo começa com B_XX.

    Exemplos:
        B_01.2026.xlsx     -> 01
        B_02_FILIAL.xls    -> 02
        2026.xlsx          -> None
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
    Registra um nome de aba se ainda não estiver em uso.

    A comparação não diferencia letras maiúsculas de minúsculas.
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
    - Arquivo B_XX tenta usar XX;
    - Se XX já estiver ocupado, usa o nome completo do arquivo;
    - Arquivo fora do padrão usa o nome completo sem a extensão;
    - Em uma duplicidade total, acrescenta um número ao final.
    """
    nome_completo = _obter_nome_completo(caminho_arquivo)
    mes = _obter_mes(caminho_arquivo)

    if mes:
        nome_mes = _registrar_nome_aba(
            mes,
            nomes_utilizados
        )

        if nome_mes is not None:
            return nome_mes

    nome_arquivo = _registrar_nome_aba(
        nome_completo,
        nomes_utilizados
    )

    if nome_arquivo is not None:
        return nome_arquivo

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
    Processa os balancetes do cliente Cooperoque.

    Regras:
    - Somente arquivos .xls e .xlsx são processados;
    - Arquivos de outras extensões são ignorados;
    - Cada arquivo Excel gera uma aba;
    - Arquivos B_XX usam inicialmente XX como nome da aba;
    - Meses repetidos usam o nome completo do arquivo;
    - Arquivos fora do padrão usam o nome completo sem a extensão.
    """
    resultados = {}
    nomes_utilizados = set()

    if not lista_arquivos:
        raise ValueError(
            "Nenhum arquivo foi selecionado para o cliente Cooperoque."
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
            "cliente Cooperoque. Selecione arquivos .xls ou .xlsx."
        )

    for arquivo in arquivos_excel:
        nome_aba = _gerar_nome_aba(
            caminho_arquivo=arquivo,
            nomes_utilizados=nomes_utilizados
        )

        dataframe = transformar_balancete_cooperoque(
            arquivo
        )

        if dataframe is None:
            raise ValueError(
                f"O processamento do arquivo "
                f"'{os.path.basename(arquivo)}' não retornou dados."
            )

        resultados[nome_aba] = dataframe

    if not resultados:
        raise ValueError(
            "Nenhum resultado foi gerado para o cliente Cooperoque."
        )

    return resultados