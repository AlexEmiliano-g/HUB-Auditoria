import os
import re

try:
    from scripts.Contabilidade.tabulador_comum import transformar_balancete_useall
except ModuleNotFoundError:
    from tabulador_comum import transformar_balancete_useall


def _limpar_nome_aba(nome):
    """
    Ajusta um texto para que possa ser utilizado como nome de aba no Excel.

    O Excel:
    - Não permite os caracteres: \\ / ? * [ ] :
    - Limita o nome da aba a 31 caracteres.
    """
    nome_limpo = re.sub(r'[\\/*?:\[\]]', "_", str(nome)).strip()

    if not nome_limpo:
        nome_limpo = "Sem nome"

    return nome_limpo[:31]


def _nome_arquivo_sem_extensao(caminho_arquivo):
    """Retorna o nome completo do arquivo sem a extensão."""
    nome_arquivo = os.path.basename(caminho_arquivo)
    return os.path.splitext(nome_arquivo)[0]


def _obter_mes_do_arquivo(caminho_arquivo):
    """
    Extrai o mês quando o arquivo começa com B_XX.

    Exemplos:
        B_01.2026.csv      -> 01
        B_02_FILIAL.csv    -> 02
        2026.csv           -> None
        B_JANEIRO.csv      -> None
    """
    nome_arquivo = os.path.basename(caminho_arquivo)
    match = re.match(r"^B_(\d{2})", nome_arquivo, re.IGNORECASE)

    return match.group(1) if match else None


def _gerar_nome_aba_unico(nome_desejado, nome_arquivo, nomes_utilizados):
    """
    Gera um nome de aba único.

    Se o nome desejado já estiver sendo utilizado, passa a usar o nome
    completo do arquivo sem a extensão.

    Se até mesmo o nome completo já estiver em uso, acrescenta um sufixo
    numérico para impedir a sobrescrita.
    """
    nome_aba = _limpar_nome_aba(nome_desejado)

    if nome_aba.lower() not in nomes_utilizados:
        nomes_utilizados.add(nome_aba.lower())
        return nome_aba

    nome_completo = _limpar_nome_aba(nome_arquivo)

    if nome_completo.lower() not in nomes_utilizados:
        nomes_utilizados.add(nome_completo.lower())
        return nome_completo

    contador = 2

    while True:
        sufixo = f"_{contador}"
        limite_nome = 31 - len(sufixo)

        nome_com_sufixo = (
            f"{nome_completo[:limite_nome]}{sufixo}"
        )

        if nome_com_sufixo.lower() not in nomes_utilizados:
            nomes_utilizados.add(nome_com_sufixo.lower())
            return nome_com_sufixo

        contador += 1


def processar(lista_arquivos):
    """
    Processa os arquivos selecionados para o sistema Useall.

    Regras:
    - Somente arquivos CSV são processados;
    - Arquivos de outras extensões são ignorados;
    - Arquivos iniciados com B_XX usam XX como nome da aba;
    - Se o mês já tiver sido utilizado, usa o nome completo do arquivo;
    - Arquivos fora do padrão usam o nome completo sem a extensão;
    - Cada arquivo CSV gera uma aba diferente.
    """
    resultados = {}
    nomes_utilizados = set()

    arquivos_csv = [
        arquivo
        for arquivo in lista_arquivos
        if os.path.splitext(arquivo)[1].lower() == ".csv"
    ]

    for arquivo in arquivos_csv:
        nome_completo = _nome_arquivo_sem_extensao(arquivo)
        mes = _obter_mes_do_arquivo(arquivo)

        if mes:
            nome_desejado = mes
        else:
            nome_desejado = nome_completo

        nome_aba = _gerar_nome_aba_unico(
            nome_desejado=nome_desejado,
            nome_arquivo=nome_completo,
            nomes_utilizados=nomes_utilizados
        )

        df = transformar_balancete_useall(arquivo)
        resultados[nome_aba] = df

    return resultados