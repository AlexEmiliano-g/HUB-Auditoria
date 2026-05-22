import pandas as pd
import os

def _converter_valor_monetario(valor):
    if pd.isna(valor) or valor == '':
        return 0.0
    texto = str(valor).strip().upper()
    sinal = 1
    if texto.endswith('D'):
        sinal = 1
        numero_str = texto[:-1].strip()
    elif texto.endswith('C'):
        sinal = -1
        numero_str = texto[:-1].strip()
    else:
        numero_str = texto
    try:
        numero_str = numero_str.replace('.', '').replace(',', '.')
        return float(numero_str) * sinal
    except (ValueError, TypeError):
        return 0.0

def transformar_balancete(caminho_arquivo):
    try:
        df_origem = pd.read_excel(caminho_arquivo, sheet_name=0, dtype=str, engine='xlrd' if caminho_arquivo.endswith('.xls') else 'openpyxl')
        df_origem.dropna(how='all', inplace=True)
    except Exception as e:
        raise ValueError(f"Não foi possível ler o arquivo {os.path.basename(caminho_arquivo)}. Erro: {e}")

    df_destino = pd.DataFrame()
    df_destino['Atividade'] = 'Geral'
    df_destino['Conta'] = df_origem.iloc[:, 0].astype(str)
    df_destino['Nome'] = df_origem.iloc[:, 1]
    df_destino['Cód. Reduzido'] = df_origem.iloc[:, 0].astype(str)
    
    saldo_anterior = pd.to_numeric(df_origem.iloc[:, 2].astype(str).str.replace(',', '.'), errors='coerce').fillna(0)
    debito = pd.to_numeric(df_origem.iloc[:, 3].astype(str).str.replace(',', '.'), errors='coerce').fillna(0)
    credito = pd.to_numeric(df_origem.iloc[:, 4].astype(str).str.replace(',', '.'), errors='coerce').fillna(0)
    saldo_acumulado_origem = pd.to_numeric(df_origem.iloc[:, 5].astype(str).str.replace(',', '.'), errors='coerce').fillna(0)
    
    df_destino['Saldo Anterior'] = -saldo_anterior
    df_destino['Débito'] = debito
    df_destino['Crédito'] = credito
    df_destino['Movimento'] = debito - credito
    df_destino['Saldo Acumulado'] = -saldo_acumulado_origem

    return df_destino

def transformar_plano_de_contas(caminho_arquivo):
    try:
        df_origem = pd.read_excel(caminho_arquivo, sheet_name=0, skiprows=5, dtype=str, engine='xlrd' if caminho_arquivo.endswith('.xls') else 'openpyxl')
        df_origem.dropna(how='all', inplace=True)
    except Exception as e:
        raise ValueError(f"Não foi possível ler o plano de contas {os.path.basename(caminho_arquivo)}. Erro: {e}")
        
    df_destino = pd.DataFrame()
    classif = df_origem.iloc[:, 2].str.strip()
    df_destino['Chave Cliente'] = classif
    df_destino['Chave D&M'] = classif.str[0].where(classif.str[0].isin(['1', '2', '3']), '')
    df_destino['Classificação'] = df_origem.iloc[:, 0]
    df_destino['Descrição'] = df_origem.iloc[:, 3].str.strip()
    df_destino['Sint./An.'] = df_origem.iloc[:, 1].str.strip().str.upper().apply(lambda x: 'S' if x == 'T' else 'A')
    
    def definir_tipo_conta(c):
        if not isinstance(c, str) or len(c) == 0: return 'R'
        primeiro_char = c[0]
        if primeiro_char == '1': return 'A'
        if primeiro_char == '2': return 'P'
        return 'R'
        
    df_destino['At/Pas/Res'] = classif.apply(definir_tipo_conta)
    df_destino['Indice'] = range(1, len(df_destino) + 1)
    return df_destino