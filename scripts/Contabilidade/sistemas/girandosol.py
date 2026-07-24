import os
import re
import copy
import pandas as pd

def ler_csv_seguro(caminho_arquivo):
    """Tenta ler o arquivo CSV com diferentes encodings para evitar quebra de caracteres."""
    try:
        df = pd.read_csv(caminho_arquivo, sep=';', dtype=str, encoding='latin1')
    except UnicodeDecodeError:
        try:
            df = pd.read_csv(caminho_arquivo, sep=';', dtype=str, encoding='utf-8')
        except Exception as e:
            raise ValueError(f"Não foi possível ler o CSV em utf-8 {os.path.basename(caminho_arquivo)}. Erro: {e}")
    except Exception as e:
        raise ValueError(f"Não foi possível ler o arquivo {os.path.basename(caminho_arquivo)}. Erro: {e}")
    
    df.dropna(how='all', inplace=True)
    return df

def processar(lista_arquivos):
    """Regras exclusivas para o sistema Girando Sol (Arquivos CSV)"""
    resultados = {}
    
    # 1. Ignora arquivos Excel (prevenção de erro)
    arquivos_validos = [f for f in lista_arquivos if not f.lower().endswith(('.xls', '.xlsx'))]
    
    # 2. Leitura e organização das abas
    arquivos_plano = [f for f in arquivos_validos if os.path.basename(f).upper().startswith('P_')]
    qtd_plano = len(arquivos_plano)
    
    for arquivo in arquivos_validos:
        nome_base = os.path.basename(arquivo)
        nome_sem_extensao = os.path.splitext(nome_base)[0]
        
        # Checa se é um arquivo de Plano de Contas
        if nome_base.upper().startswith('P_'):
            if qtd_plano == 1:
                nome_aba = "Plano de Contas"
            else:
                nome_aba = nome_sem_extensao
                
            if nome_aba in resultados:
                nome_aba = f"{nome_aba}_duplicado"
                
            resultados[nome_aba] = _transformar_plano_de_contas_girandosol(arquivo)
            
        else:
            # Se não for P_, trata como balancete (B_)
            prefixo = nome_base[:4]
            match = re.match(r"^B_(\d{2})", prefixo, re.IGNORECASE)
            
            if match:
                nome_aba = match.group(1)
                if nome_aba in resultados:
                    nome_aba = nome_sem_extensao
            else:
                nome_aba = nome_sem_extensao
                
            if nome_aba in resultados:
                nome_aba = f"{nome_aba}_duplicado"

            resultados[nome_aba] = _transformar_balancete_girandosol(arquivo)

    # 3. Lógica de Recálculo (A partir de Abril para contas de Resultado)
    resultados_recalculados = copy.deepcopy(resultados)
    
    meses_disponiveis = [int(aba) for aba in resultados_recalculados.keys() if re.match(r"^\d{2}$", aba)]
    meses_validos = []
    
    if 1 in meses_disponiveis:
        esperado = 1
        for m in sorted(meses_disponiveis):
            if m == esperado:
                meses_validos.append(m)
                esperado += 1
            else:
                break 
                
    for m in meses_validos:
        if m >= 4:
            aba_atual = f"{m:02d}"
            aba_anterior = f"{(m-1):02d}"
            
            df_curr = resultados_recalculados[aba_atual]
            df_prev = resultados_recalculados[aba_anterior]
            
            # --- PARTE A: Injeção das contas sem movimento ---
            mask_prev_resultado = df_prev['Conta'].astype(str).str[0].isin(['3', '4', '5', '6', '7', '8', '9'])
            contas_curr = set(df_curr['Conta'].astype(str))
            
            # Pega os índices das contas que existem no mês anterior mas sumiram no atual
            missing_indices = df_prev[mask_prev_resultado & (~df_prev['Conta'].astype(str).isin(contas_curr))].index
            
            if not missing_indices.empty:
                curr_records = df_curr.to_dict('records')
                # Mapeamento dinâmico {Conta: Índice na lista}
                curr_conta_to_idx = {str(row['Conta']): i for i, row in enumerate(curr_records)}
                
                for idx in missing_indices:
                    row_prev = df_prev.loc[idx]
                    conta = str(row_prev['Conta'])
                    
                    # Busca reversa para achar a conta âncora que existe no mês atual
                    anchor_idx_in_curr = -1
                    for prev_idx in range(idx - 1, -1, -1):
                        anchor_conta = str(df_prev.loc[prev_idx, 'Conta'])
                        if anchor_conta in curr_conta_to_idx:
                            anchor_idx_in_curr = curr_conta_to_idx[anchor_conta]
                            break
                            
                    # Nova linha formatada com valores zerados e o acumulado do mês anterior
                    new_record = {
                        'Atividade': 'Geral',
                        'Conta': conta,
                        'Nome': row_prev['Nome'],
                        'Cód. Reduzido': row_prev['Cód. Reduzido'],
                        'Saldo Anterior': row_prev['Saldo Acumulado'],
                        'Débito': 0.0,
                        'Crédito': 0.0,
                        'Movimento': 0.0,
                        'Saldo Acumulado': row_prev['Saldo Acumulado']
                    }
                    
                    # Insere abaixo da âncora
                    insert_pos = anchor_idx_in_curr + 1 if anchor_idx_in_curr != -1 else 0
                    curr_records.insert(insert_pos, new_record)
                    
                    # Atualiza os índices no dicionário
                    curr_conta_to_idx[conta] = insert_pos
                    for k, v in list(curr_conta_to_idx.items()):
                        if v >= insert_pos and k != conta:
                            curr_conta_to_idx[k] = v + 1
                
                # Reconstrói o DataFrame com as linhas injetadas na ordem correta
                df_curr = pd.DataFrame(curr_records)
                
            # --- PARTE B: Recálculo dos Valores ---
            mask_resultado = df_curr['Conta'].astype(str).str[0].isin(['3', '4', '5', '6', '7', '8', '9'])
            mapping_prev = df_prev.groupby('Conta')['Saldo Acumulado'].last().to_dict()
            
            df_curr.loc[mask_resultado, 'Saldo Anterior'] = df_curr.loc[mask_resultado, 'Conta'].map(mapping_prev).fillna(0.0)
            df_curr.loc[mask_resultado, 'Saldo Acumulado'] = df_curr.loc[mask_resultado, 'Saldo Anterior'] + df_curr.loc[mask_resultado, 'Movimento']
            
            resultados_recalculados[aba_atual] = df_curr

    return resultados, resultados_recalculados


def _transformar_balancete_girandosol(caminho_arquivo):
    df_origem = ler_csv_seguro(caminho_arquivo)
    df_destino = pd.DataFrame()
    num_cols = len(df_origem.columns)
    
    df_destino['Conta'] = df_origem.iloc[:, 0].astype(str).str.strip() if num_cols > 0 else ''
    df_destino.insert(0, 'Atividade', 'Geral')
    df_destino['Nome'] = df_origem.iloc[:, 1].astype(str).str.strip() if num_cols > 1 else ''
    df_destino['Cód. Reduzido'] = df_origem.iloc[:, 2].astype(str).str.strip() if num_cols > 2 else ''
    
    def extrair_valor(index_coluna):
        if num_cols > index_coluna:
            return pd.to_numeric(
                df_origem.iloc[:, index_coluna].astype(str)
                .str.replace('.', '', regex=False)
                .str.replace(',', '.', regex=False), 
                errors='coerce'
            ).fillna(0)
        return 0.0

    df_destino['Saldo Anterior'] = extrair_valor(4)
    df_destino['Débito'] = extrair_valor(5)
    df_destino['Crédito'] = extrair_valor(6)
    df_destino['Movimento'] = extrair_valor(7)
    df_destino['Saldo Acumulado'] = extrair_valor(8)

    return df_destino


def _transformar_plano_de_contas_girandosol(caminho_arquivo):
    df_origem = ler_csv_seguro(caminho_arquivo)
    df_destino = pd.DataFrame()
    num_cols = len(df_origem.columns)
    
    # Mapeamento do novo plano de contas
    df_destino['Chave Cliente'] = df_origem.iloc[:, 0].astype(str).str.strip() if num_cols > 0 else ''
    df_destino['Chave D&M'] = ""
    df_destino['Classificação'] = df_origem.iloc[:, 1].astype(str).str.strip() if num_cols > 1 else ''
    df_destino['Descrição'] = df_origem.iloc[:, 2].astype(str).str.strip() if num_cols > 2 else ''
    df_destino['Sint./An.'] = ""
    df_destino['At/Pas/Res'] = ""
    df_destino['Indice'] = range(1, len(df_destino) + 1)
    
    return df_destino