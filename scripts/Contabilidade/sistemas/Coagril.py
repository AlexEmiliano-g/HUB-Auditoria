import os
import re

try:
    from scripts.Contabilidade.tabulador_comum import (
        transformar_balancete_coagril,
    )
except ModuleNotFoundError:
    from tabulador_comum import (
        transformar_balancete_coagril,
    )


def _limpar_nome_aba(nome):
    """
    Ajusta um texto para que possa ser utilizado como nome de aba no Excel.

    O Excel:
    - não permite os caracteres: \ / ? * [ ] :
    - limita o nome da aba a 31 caracteres.
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
    """
    Retorna o nome completo do arquivo sem a extensão.
    """
    nome_arquivo = os.path.basename(caminho_arquivo)
    nome_sem_extensao = os.path.splitext(nome_arquivo)[0]

    return nome_sem_extensao.strip()


def _obter_mes(caminho_arquivo):
    """
    Obtém o mês quando o nome do arquivo começa com B_XX.

    Exemplos:
        B_01.2026.txt       -> 01
        B_02_FILIAL.txt     -> 02
        Balancete.txt       -> None
        2026.txt            -> None
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
    - arquivo iniciado por B_XX tenta utilizar XX;
    - se XX já estiver em uso, utiliza o nome completo do arquivo;
    - arquivo fora do padrão utiliza o nome completo sem extensão;
    - se o nome completo também estiver repetido, adiciona um sufixo.
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
    Processa os arquivos do cliente Coagril.

    Regras:
    - arquivos TXT são tratados como balancetes;
    - arquivos Excel são tratados como planos de contas;
    - a regra especial de "(-) RATEIO" é aplicada somente aos
      planos de contas processados pelo Coagril;
    - arquivos de outras extensões são ignorados.
    """
    resultados = {}
    nomes_utilizados = set()

    if not lista_arquivos:
        raise ValueError(
            "Nenhum arquivo foi selecionado para o cliente Coagril."
        )

    arquivos_validos = [
        arquivo
        for arquivo in lista_arquivos
        if os.path.splitext(str(arquivo))[1].lower()
        in {".txt", ".xls", ".xlsx"}
    ]

    if not arquivos_validos:
        raise ValueError(
            "Nenhum arquivo compatível foi encontrado para o cliente "
            "Coagril. Selecione arquivos TXT, XLS ou XLSX."
        )

    for arquivo in arquivos_validos:
        extensao = os.path.splitext(str(arquivo))[1].lower()

        nome_aba = _gerar_nome_aba(
            caminho_arquivo=arquivo,
            nomes_utilizados=nomes_utilizados
        )

        dataframe = transformar_balancete_coagril(arquivo)

        if dataframe is None or dataframe.empty:
            raise ValueError(
                f"O arquivo '{os.path.basename(arquivo)}' não retornou "
                "dados válidos para tabulação."
            )

        resultados[nome_aba] = dataframe

    if not resultados:
        raise ValueError(
            "Nenhum resultado foi gerado para o cliente Coagril."
        )

    return resultados