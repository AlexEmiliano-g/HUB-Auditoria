import os
import re
import pandas as pd

def ler_txt_seguro(caminho_arquivo):
    """
    Tenta ler o arquivo TXT com diferentes encodings (utf-8 e latin1).
    Utiliza a vírgula como separador e as aspas duplas como encapsulador de texto,
    que é o formato padrão do Paradiso Giovanella.
    """
    nome_arquivo = os.path.basename(caminho_arquivo)
    
    # Parâmetros de leitura para o padrão do TXT fornecido
    config = {
        'sep': ',',
        'quotechar': '"',
        'header': None,
        'dtype': str
    }
    
    try:
        df = pd.read_csv(caminho_arquivo, encoding='utf-8', **config)
    except UnicodeDecodeError:
        try:
            df = pd.read_csv(caminho_arquivo, encoding='latin1', **config)
        except Exception as e:
            raise ValueError(f"Não foi possível ler o TXT em latin1 '{nome_arquivo}'. Erro: {e}")
    except Exception as e:
        raise ValueError(f"Não foi possível ler o arquivo '{nome_arquivo}'. Erro: {e}")
    
    df.dropna(how='all', inplace=True)
    return df

def processar(lista_arquivos):
    """
    Processa os arquivos TXT selecionados para o sistema Paradiso Giovanella.
    """
    resultados = {}
    
    # Filtra apenas os arquivos TXT
    arquivos_txt = [f for f in lista_arquivos if f.lower().endswith('.txt')]
    
    for arquivo in arquivos_txt:
        nome_base = os.path.basename(arquivo)
        
        # Aplica a regra B_xx para definir o nome da aba
        # Captura exatamente os 2 primeiros dígitos após "B_" (ex: B_01.2026.txt -> 01)
        match = re.match(r"^B_(\d{2})", nome_base, re.IGNORECASE)
        
        if match:
            nome_aba = match.group(1)
        else:
            # Fallback de segurança: se o arquivo não tiver o padrão, pega o nome dele
            nome_sem_extensao = os.path.splitext(nome_base)[0]
            nome_aba = nome_sem_extensao[:31].strip()
        
        df_origem = ler_txt_seguro(arquivo)
        df_destino = pd.DataFrame()
        num_cols = len(df_origem.columns)
        
        # Mapeamento de Texto
        df_destino['Atividade'] = 'Geral'
        df_destino['Conta'] = df_origem.iloc[:, 1].astype(str).str.strip() if num_cols > 1 else ''
        df_destino['Nome'] = df_origem.iloc[:, 2].astype(str).str.strip() if num_cols > 2 else ''
        df_destino['Cód. Reduzido'] = df_origem.iloc[:, 0].astype(str).str.strip() if num_cols > 0 else ''
        
        # Função para limpar e converter os valores monetários
        def extrair_valor(index_coluna):
            if num_cols > index_coluna:
                return pd.to_numeric(
                    df_origem.iloc[:, index_coluna].astype(str)
                    .str.replace('.', '', regex=False) # Remove o ponto de milhar
                    .str.replace(',', '.', regex=False), # Troca a vírgula decimal por ponto
                    errors='coerce'
                ).fillna(0.0)
            return 0.0

        # Mapeamento de Valores Numéricos
        df_destino['Saldo Anterior'] = extrair_valor(3)
        
        debito = extrair_valor(4)
        df_destino['Débito'] = debito
        
        credito = extrair_valor(5)
        df_destino['Crédito'] = credito
        
        # Cálculo do Movimento
        df_destino['Movimento'] = debito - credito
        
        df_destino['Saldo Acumulado'] = extrair_valor(6)
        
        # Tratamento para impedir sobreposição caso existam dois arquivos com a mesma aba
        contador = 2
        aba_original = nome_aba
        while nome_aba in resultados:
            sufixo = f"_{contador}"
            nome_aba = f"{aba_original[:31-len(sufixo)]}{sufixo}"
            contador += 1
            
        resultados[nome_aba] = df_destino
        
    return resultados