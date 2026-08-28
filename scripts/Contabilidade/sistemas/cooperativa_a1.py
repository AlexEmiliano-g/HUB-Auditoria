import os
import warnings
import pandas as pd
import re

# Importa a regra global para nomeação de abas
try:
    from scripts.Contabilidade.tabulador_comum import obter_nome_aba_seguro
except ModuleNotFoundError:
    from tabulador_comum import obter_nome_aba_seguro

# ==============================================================================
# FUNÇÕES PRIVADAS DA COOPERATIVA A1
# ==============================================================================

def _converter_numero_cooperativa_a1(valor):
    """
    Converte valores monetários do TXT da Cooperativa A1 para float.
    Trata formatações atípicas como '1948.655.803,10' ou negativos '-37.224.930,08'.
    """
    if not valor:
        return 0.0
    
    texto = str(valor).strip()
    if not texto:
        return 0.0

    # Remove o ponto de milhar e troca a vírgula decimal por ponto
    texto = texto.replace(".", "").replace(",", ".")
    
    try:
        return float(texto)
    except (ValueError, TypeError):
        return 0.0


def transformar_balancete_cooperativa_a1(caminho_arquivo):
    """
    Transforma o balancete TXT do cliente Cooperativa A1.
    
    Lê o arquivo texto linha a linha utilizando expressões regulares
    para driblar cabeçalhos e quebras de página. Calcula o Saldo Anterior
    dinamicamente (Saldo Atual - Movimento).
    """
    nome_arquivo = os.path.basename(caminho_arquivo)
    
    # Tenta múltiplas codificações para abrir o TXT legado
    codificacoes = ["utf-8-sig", "cp1252", "latin-1"]
    linhas = []
    
    for codificacao in codificacoes:
        try:
            with open(caminho_arquivo, mode="r", encoding=codificacao) as arquivo:
                linhas = arquivo.readlines()
            break
        except UnicodeDecodeError:
            continue
        except OSError as erro:
            raise ValueError(f"Não foi possível abrir o arquivo TXT '{nome_arquivo}'. Erro: {erro}")
            
    if not linhas:
        raise ValueError(f"Não foi possível ler o arquivo '{nome_arquivo}' ou ele está vazio.")

    # Regex para capturar: Conta, Nome, Debito, Credito, Movimento, Saldo Atual
    # Ex: ' 01 01 01   DISPONIVEL   593.578.129,84   642.748.012,78   -49.169.882,94   632.598.142,41'
    regex_linha_contabil = re.compile(
        r"^\s*([\d\s]+)\s+(.*?)\s+([-\d.,]+)\s+([-\d.,]+)\s+([-\d.,]+)\s+([-\d.,]+)\s*$"
    )
    
    registros = []
    
    for linha in linhas:
        # Ignora as quebras de página (form feed) comuns no layout
        linha_limpa = linha.replace("\x0c", "")
        
        match = regex_linha_contabil.match(linha_limpa)
        if match:
            conta_com_espacos = match.group(1).strip()
            nome_conta = match.group(2).strip()
            
            # Converte as colunas monetárias
            debito = _converter_numero_cooperativa_a1(match.group(3))
            credito = _converter_numero_cooperativa_a1(match.group(4))
            movimento = _converter_numero_cooperativa_a1(match.group(5))
            saldo_atual = _converter_numero_cooperativa_a1(match.group(6))
            
            # O Saldo Anterior é oculto no relatório, fazemos a engenharia reversa
            saldo_anterior = saldo_atual - movimento
            
            # O Código Reduzido será a conta sem os espaços
            codigo_reduzido = conta_com_espacos.replace(" ", "")
            
            registros.append({
                "Atividade": "Geral",
                "Conta": conta_com_espacos,
                "Nome": nome_conta,
                "Cód. Reduzido": codigo_reduzido,
                "Saldo Anterior": round(saldo_anterior, 2),
                "Débito": round(debito, 2),
                "Crédito": round(credito, 2),
                "Movimento": round(movimento, 2),
                "Saldo Acumulado": round(saldo_atual, 2)
            })

    if not registros:
        raise ValueError(f"Nenhuma conta contábil válida foi encontrada no arquivo '{nome_arquivo}'.")

    df_destino = pd.DataFrame(registros)
    return df_destino

# ==============================================================================
# ROTEADOR DO SISTEMA COOPERATIVA A1
# ==============================================================================

def processar(lista_arquivos):
    """
    Regras exclusivas para o sistema Cooperativa A1:
    - Lê balancetes APENAS de arquivos de texto (.txt).
    - Gera nome de abas dinamicamente, tratando colisões (regra B_XX).
    """
    resultados = {}
    abas_existentes = set()
    
    arquivos_balancete = [f for f in lista_arquivos if f.lower().endswith('.txt')]
    arquivos_ignorados = [f for f in lista_arquivos if not f.lower().endswith('.txt')]

    if arquivos_ignorados:
        nomes_ignorados = ", ".join([os.path.basename(f) for f in arquivos_ignorados])
        warnings.warn(
            f"Os seguintes arquivos foram ignorados na Cooperativa A1 (espera-se apenas .TXT): {nomes_ignorados}"
        )

    for arquivo in arquivos_balancete:
        nome_aba = obter_nome_aba_seguro(arquivo, abas_existentes)
        abas_existentes.add(nome_aba)
        
        # Chama a função que agora é interna
        df = transformar_balancete_cooperativa_a1(arquivo)
        resultados[nome_aba] = df

    if not resultados:
        raise ValueError("Nenhum arquivo de texto (.txt) válido foi processado para o sistema Cooperativa A1.")

    return resultados