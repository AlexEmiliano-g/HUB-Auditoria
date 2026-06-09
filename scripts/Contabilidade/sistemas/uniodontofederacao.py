import os
import pandas as pd

def processar(lista_arquivos):
    """Regras exclusivas para o sistema UniOdontoFederacao com Tabela de Check Dinâmica"""
    resultados = {}
    avisos = []
    
    abas_esperadas = ['Ativo', 'Passivo', 'Receitas', 'Despesas', 'Custos']

    for arquivo in lista_arquivos:
        nome_base = os.path.splitext(os.path.basename(arquivo))[0]
        
        try:
            xls = pd.ExcelFile(arquivo)
            abas_existentes = xls.sheet_names
            
            abas_faltantes = [aba for aba in abas_esperadas if aba not in abas_existentes]
            if abas_faltantes:
                avisos.append({
                    'Arquivo original': os.path.basename(arquivo),
                    'Status': 'Aviso',
                    'Detalhe': f"Abas esperadas ausentes: {', '.join(abas_faltantes)}. As demais foram empilhadas."
                })
            
            dfs_empilhados = []
            
            for aba in abas_esperadas:
                if aba in abas_existentes:
                    try:
                        df_raw = pd.read_excel(xls, sheet_name=aba, header=None)
                        
                        idx_cabecalho = -1
                        for i, row in df_raw.iterrows():
                            linha_str = [str(x).strip().upper() for x in row.values]
                            if 'CLASSIFICADOR' in linha_str and 'NOME DA CONTA' in linha_str:
                                idx_cabecalho = i
                                break
                        
                        if idx_cabecalho == -1:
                            avisos.append({
                                'Arquivo original': os.path.basename(arquivo),
                                'Status': 'Erro',
                                'Detalhe': f"Aba '{aba}' ignorada. Cabeçalho padrão não encontrado."
                            })
                            continue
                            
                        idx_inicio = idx_cabecalho + 2
                        if idx_inicio >= len(df_raw): continue 
                            
                        idx_fim = idx_inicio
                        while idx_fim < len(df_raw):
                            val_col_A = df_raw.iloc[idx_fim, 0] 
                            val_col_B = df_raw.iloc[idx_fim, 1] 
                            
                            if pd.isna(val_col_A) and pd.isna(val_col_B): break
                            if str(val_col_A).strip() == '' and str(val_col_B).strip() == '': break
                            idx_fim += 1
                            
                        if idx_inicio == idx_fim: continue 
                            
                        df_bloco = df_raw.iloc[idx_inicio:idx_fim].copy()
                        
                        df_aba = pd.DataFrame()
                        df_aba['(sem preenchimento)'] = ""
                        df_aba['Conta'] = df_bloco.iloc[:, 1]           
                        df_aba['Nome'] = df_bloco.iloc[:, 2]            
                        df_aba['Cód. Reduzido'] = df_bloco.iloc[:, 0]   
                        df_aba['Saldo Anterior'] = df_bloco.iloc[:, 4]  
                        df_aba['Débito'] = df_bloco.iloc[:, 5]          
                        df_aba['Crédito'] = df_bloco.iloc[:, 6]         
                        
                        debito_num = pd.to_numeric(df_aba['Débito'], errors='coerce').fillna(0)
                        credito_num = pd.to_numeric(df_aba['Crédito'], errors='coerce').fillna(0)
                        df_aba['Movimento'] = debito_num - credito_num
                        df_aba['Saldo Acumulado'] = df_bloco.iloc[:, 7]
                        
                        # Criando a nova coluna em branco para o D/C (Será a Coluna J)
                        df_aba['D/C'] = "" 
                        
                        dfs_empilhados.append(df_aba)
                        
                    except Exception as e:
                        avisos.append({
                            'Arquivo original': os.path.basename(arquivo),
                            'Status': 'Erro',
                            'Detalhe': f"Falha ao processar a aba '{aba}': {str(e)}"
                        })
                        
            if dfs_empilhados:
                df_final_arquivo = pd.concat(dfs_empilhados, ignore_index=True)
                
                # --- PROCESSAMENTO DA TABELA DE CHECK COM FÓRMULAS ---
                N = len(df_final_arquivo)
                
                def achar_linha_excel(df, termo_exato, termo_contem):
                    for idx, nome in enumerate(df['Nome']):
                        if str(nome).strip().upper() == termo_exato.upper(): return idx + 2
                    for idx, nome in enumerate(df['Nome']):
                        if termo_contem.upper() in str(nome).strip().upper(): return idx + 2
                    return None

                row_ativo = achar_linha_excel(df_final_arquivo, 'ATIVO', 'ATIVO')
                row_passivo = achar_linha_excel(df_final_arquivo, 'PASSIVO', 'PASSIVO')
                row_destinacao = achar_linha_excel(df_final_arquivo, 'CONTAS DE DESTINAÇÃO / APURAÇÃO DE RESULTADO', 'DESTINAÇÃO')
                if not row_destinacao: row_destinacao = achar_linha_excel(df_final_arquivo, 'APURAÇÃO DE RESULTADO', 'APURAÇÃO')
                row_despesas = achar_linha_excel(df_final_arquivo, 'DESPESAS', 'DESPESA')
                row_receitas = achar_linha_excel(df_final_arquivo, 'RECEITAS', 'RECEITA')

                cell_ativo = f"I{row_ativo}" if row_ativo else "0"
                cell_passivo = f"I{row_passivo}" if row_passivo else "0"
                cell_destinacao = f"I{row_destinacao}" if row_destinacao else "0"
                cell_despesas = f"I{row_despesas}" if row_despesas else "0"
                cell_receitas = f"I{row_receitas}" if row_receitas else "0"

                # Estrutura do Check com CUSTOS removido do cálculo e coluna D/C preenchida
                check_rows = [
                    {col: "" for col in df_final_arquivo.columns}, # Linha em branco
                    {'(sem preenchimento)': '', 'Conta': '', 'Nome': 'ATIVO', 'Cód. Reduzido': '', 'Saldo Anterior': '', 'Débito': '', 'Crédito': '', 'Movimento': '', 'Saldo Acumulado': f"={cell_ativo}", 'D/C': 'D'},
                    {'(sem preenchimento)': '', 'Conta': '', 'Nome': 'PASSIVO', 'Cód. Reduzido': '', 'Saldo Anterior': '', 'Débito': '', 'Crédito': '', 'Movimento': '', 'Saldo Acumulado': f"={cell_passivo}", 'D/C': 'C'},
                    {'(sem preenchimento)': '', 'Conta': '', 'Nome': 'CONTAS DE DESTINAÇÃO / APURAÇÃO DE RESULTADO', 'Cód. Reduzido': '', 'Saldo Anterior': '', 'Débito': '', 'Crédito': '', 'Movimento': '', 'Saldo Acumulado': f"={cell_destinacao}", 'D/C': 'D'},
                    {'(sem preenchimento)': '', 'Conta': '', 'Nome': 'DESPESAS', 'Cód. Reduzido': '', 'Saldo Anterior': '', 'Débito': '', 'Crédito': '', 'Movimento': '', 'Saldo Acumulado': f"={cell_despesas}", 'D/C': 'D'},
                    {'(sem preenchimento)': '', 'Conta': '', 'Nome': 'RECEITAS', 'Cód. Reduzido': '', 'Saldo Anterior': '', 'Débito': '', 'Crédito': '', 'Movimento': '', 'Saldo Acumulado': f"={cell_receitas}", 'D/C': 'C'},
                    {'(sem preenchimento)': '', 'Conta': '', 'Nome': 'Diferença', 'Cód. Reduzido': '', 'Saldo Anterior': '', 'Débito': '', 'Crédito': '', 'Movimento': '', 'Saldo Acumulado': f"=(I{N+3}+I{N+5}+I{N+6})-(I{N+4}+I{N+7})", 'D/C': ''},
                    {'(sem preenchimento)': '', 'Conta': '', 'Nome': 'Resultado do Período', 'Cód. Reduzido': '', 'Saldo Anterior': '', 'Débito': '', 'Crédito': '', 'Movimento': '', 'Saldo Acumulado': f"=I{N+7}-I{N+6}-I{N+5}", 'D/C': 'C'}
                ]
                
                df_check = pd.DataFrame(check_rows)
                df_final_arquivo = pd.concat([df_final_arquivo, df_check], ignore_index=True)
                # -----------------------------------------------------------------
                
                resultados[nome_base] = df_final_arquivo
            else:
                avisos.append({
                    'Arquivo original': os.path.basename(arquivo),
                    'Status': 'Aviso',
                    'Detalhe': "Nenhum dado capturado nas abas. O arquivo não foi tabulado."
                })
                
        except Exception as e:
            avisos.append({
                'Arquivo original': os.path.basename(arquivo),
                'Status': 'Erro Crítico',
                'Detalhe': f"Impossível ler o arquivo: {str(e)}"
            })

    if avisos:
        resultados['Avisos'] = pd.DataFrame(avisos)
        
    return resultados