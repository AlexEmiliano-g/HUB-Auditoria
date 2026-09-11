import os
import warnings
import pandas as pd
import re

# 1. IMPORTA APENAS A FERRAMENTA GLOBAL NECESSÁRIA
try:
    from scripts.Contabilidade.tabulador_comum import obter_nome_aba_seguro
except ModuleNotFoundError:
    from tabulador_comum import obter_nome_aba_seguro

# ==============================================================================
# FUNÇÕES AUXILIARES PRIVADAS DA COOABRIEL
# ==============================================================================

def _converter_numero_cooabriel(valor):
    """
    Converte valores monetários do cliente Cooabriel para float.
    Trata formatações como '2.892.239.571,50' ou '0,00'.
    """
    if pd.isna(valor):
        return 0.0

    if isinstance(valor, (int, float)):
        return float(valor)

    texto = str(valor).strip()

    if not texto:
        return 0.0

    texto = texto.replace("\xa0", "").replace(" ", "")

    if "," in texto:
        texto = texto.replace(".", "").replace(",", ".")

    try:
        return float(texto)
    except (ValueError, TypeError):
        return 0.0


def _aplicar_quebra_numero(base, complemento):
    """
    Junta a base do número com a parte que 'vazou' para a linha de baixo.
    Usa formatação matemática para evitar erros catastróficos de leitura 
    quando o Excel omite o ",0".
    """
    base_str = str(base).strip() if pd.notna(base) else ""
    comp_str = str(complemento).strip() if pd.notna(complemento) else ""
    
    # Garante que o complemento contenha apenas os números que vazaram
    comp_str = re.sub(r"\D", "", comp_str)
    
    if not comp_str:
        return base_str
    if not base_str:
        return comp_str
        
    # Converte a base incompleta ignorando as formatações nocivas do Excel
    val = _converter_numero_cooabriel(base_str)
    
    # Força o formato com 2 casas decimais. Ex: 1887918607.0 -> "1887918607.00"
    val_fmt = f"{val:.2f}"
    
    # Se o complemento é 1 dígito (ex: '4'), injeta e substitui o último zero do centavo.
    if len(comp_str) == 1:
        return val_fmt[:-1] + comp_str
    # Se o complemento for 2 dígitos (ex: '45'), substitui as duas casas decimais.
    elif len(comp_str) == 2:
        return val_fmt[:-2] + comp_str
    else:
        return val_fmt + comp_str


def _aplicar_natureza_cooabriel(valor, natureza):
    """
    Aplica o sinal do saldo conforme a natureza contábil.
    Regras: D = positivo, C = negativo.
    """
    numero = _converter_numero_cooabriel(valor)

    if pd.isna(natureza):
        return numero

    natureza_texto = str(natureza).strip().upper()

    if natureza_texto == "D":
        return abs(numero)
    if natureza_texto == "C":
        return -abs(numero)

    return numero

# ==============================================================================
# LÓGICA PRINCIPAL DE TRANSFORMAÇÃO DA COOABRIEL
# ==============================================================================

def transformar_balancete_cooabriel(caminho_arquivo):
    """
    Transforma o balancete Excel do cliente Cooabriel.
    Reconstrói os registros "órfãos" colando-os na linha principal.
    """
    nome_arquivo = os.path.basename(caminho_arquivo)
    extensao = os.path.splitext(caminho_arquivo)[1].lower()

    if extensao not in {".xls", ".xlsx"}:
        raise ValueError(f"O arquivo '{nome_arquivo}' não é um arquivo Excel válido.")

    engine = "xlrd" if extensao == ".xls" else "openpyxl"

    try:
        df_raw = pd.read_excel(
            caminho_arquivo, 
            sheet_name=0, 
            header=None, 
            dtype=str, 
            engine=engine
        )
    except Exception as erro:
        raise ValueError(f"Não foi possível ler o arquivo '{nome_arquivo}'. Erro: {erro}")

    if df_raw.shape[1] < 9:
        raise ValueError(
            f"O arquivo '{nome_arquivo}' possui {df_raw.shape[1]} coluna(s), "
            "mas o layout Cooabriel exige 9 colunas (Conta, Chave, Descrição, SA, Nat, Déb, Créd, SF, Nat)."
        )

    registros_corrigidos = []
    ultimo_registro = None

    # Algoritmo de varredura e reconstrução de quebras de linha
    for i, row in df_raw.iterrows():
        conta = str(row[0]).strip() if pd.notna(row[0]) else ""
        chave = str(row[1]).strip() if pd.notna(row[1]) else ""
        desc = str(row[2]).strip() if pd.notna(row[2]) else ""

        # Ignora cabeçalhos principais
        if conta in ["Conta", "COOP AGRARIA DOS CAFEICULTORES DE SAO GABRIEL", 
                     "Balancete de Verificação"] or conta.startswith("CNPJ:"):
            continue

        # Verifica se é rodapé (Folha/Paginação)
        texto_linha = " ".join([str(x) for x in row.values if pd.notna(x)]).upper()
        if "FOLHA:" in texto_linha or "MÊS/ANO:" in texto_linha or "MES/ANO:" in texto_linha:
            continue

        # IDENTIFICAÇÃO DE LINHA QUEBRADA (CONTINUAÇÃO)
        # Como toda conta contábil real no Cooabriel tem uma Chave, se a chave estiver vazia,
        # significa que a linha é um fragmento (continuação) da conta anterior.
        if not chave:
            if ultimo_registro is not None:
                # 1. Junta pedaço da Conta (removendo vírgulas ou pontos indesejados no início)
                if conta:
                    clean_conta = re.sub(r"^[.,]", "", conta)
                    ultimo_registro["Conta"] += clean_conta
                    
                # 2. Junta pedaço da Descrição
                if desc:
                    if ultimo_registro["Nome"]:
                        ultimo_registro["Nome"] += " " + desc
                    else:
                        ultimo_registro["Nome"] = desc

                # 3. Junta Valores com injeção matemática exata
                if pd.notna(row[3]) and str(row[3]).strip():
                    ultimo_registro["Saldo Anterior"] = _aplicar_quebra_numero(ultimo_registro["Saldo Anterior"], row[3])
                if pd.notna(row[5]) and str(row[5]).strip():
                    ultimo_registro["Débitos"] = _aplicar_quebra_numero(ultimo_registro["Débitos"], row[5])
                if pd.notna(row[6]) and str(row[6]).strip():
                    ultimo_registro["Créditos"] = _aplicar_quebra_numero(ultimo_registro["Créditos"], row[6])
                if pd.notna(row[7]) and str(row[7]).strip():
                    ultimo_registro["Saldo Final"] = _aplicar_quebra_numero(ultimo_registro["Saldo Final"], row[7])
            continue

        # Se tem chave, é um novo registro válido
        record = {
            "Conta": conta,
            "Chave": chave,
            "Nome": desc,
            "Saldo Anterior": str(row[3]).strip() if pd.notna(row[3]) else "0",
            "Nat SA": str(row[4]).strip() if pd.notna(row[4]) else "",
            "Débitos": str(row[5]).strip() if pd.notna(row[5]) else "0",
            "Créditos": str(row[6]).strip() if pd.notna(row[6]) else "0",
            "Saldo Final": str(row[7]).strip() if pd.notna(row[7]) else "0",
            "Nat SF": str(row[8]).strip() if pd.notna(row[8]) else ""
        }
        registros_corrigidos.append(record)
        ultimo_registro = record

    if not registros_corrigidos:
        raise ValueError(f"Nenhuma conta contábil foi encontrada no arquivo '{nome_arquivo}'.")

    # Transforma os dicionários limpos em DataFrame
    df_registros = pd.DataFrame(registros_corrigidos)
    df_destino = pd.DataFrame(index=df_registros.index)

    df_destino["Atividade"] = "Geral"
    df_destino["Conta"] = df_registros["Conta"]
    df_destino["Nome"] = df_registros["Nome"]
    df_destino["Cód. Reduzido"] = df_registros["Chave"]

    # Converte os números e aplica a natureza (Positivo = D, Negativo = C)
    df_destino["Saldo Anterior"] = pd.Series([
        _aplicar_natureza_cooabriel(val, nat) 
        for val, nat in zip(df_registros["Saldo Anterior"], df_registros["Nat SA"])
    ])
    
    debito = df_registros["Débitos"].apply(_converter_numero_cooabriel)
    credito = df_registros["Créditos"].apply(_converter_numero_cooabriel)
    
    df_destino["Débito"] = debito
    df_destino["Crédito"] = credito
    df_destino["Movimento"] = debito - credito
    
    df_destino["Saldo Acumulado"] = pd.Series([
        _aplicar_natureza_cooabriel(val, nat) 
        for val, nat in zip(df_registros["Saldo Final"], df_registros["Nat SF"])
    ])

    return df_destino

# ==============================================================================
# MOTOR DO SISTEMA (ROTEAMENTO)
# ==============================================================================

def processar(lista_arquivos):
    """
    Regras exclusivas para o sistema Cooabriel:
    - Lê balancetes de planilhas Excel (XLS, XLSX).
    - Gera nome de abas dinamicamente, tratando colisões (regra B_XX).
    """
    resultados = {}
    abas_existentes = set()
    
    arquivos_balancete = [f for f in lista_arquivos if f.lower().endswith(('.xls', '.xlsx'))]
    arquivos_ignorados = [f for f in lista_arquivos if not f.lower().endswith(('.xls', '.xlsx'))]

    if arquivos_ignorados:
        nomes_ignorados = ", ".join([os.path.basename(f) for f in arquivos_ignorados])
        warnings.warn(
            f"Os arquivos a seguir foram ignorados na Cooabriel (apenas Excel é aceito): {nomes_ignorados}"
        )

    for arquivo in arquivos_balancete:
        nome_aba = obter_nome_aba_seguro(arquivo, abas_existentes)
        abas_existentes.add(nome_aba)
        
        # Chama a função que agora é interna deste módulo
        df = transformar_balancete_cooabriel(arquivo)
        resultados[nome_aba] = df

    if not resultados:
        raise ValueError("Nenhum arquivo válido (.xls ou .xlsx) foi processado para o sistema Cooabriel.")

    return resultados