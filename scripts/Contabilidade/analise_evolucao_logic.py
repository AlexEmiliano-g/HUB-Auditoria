import os
import pandas as pd
import numpy as np
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

def parse_perc(val_str):
    if not val_str: return float('inf')
    v = val_str.replace('%', '').replace(',', '.').strip()
    try:
        return float(v) / 100.0
    except:
        return 0.0

def executar_analise_evolucao(caminho_entrada, opcoes, regras=None):
    if regras is None: regras = {} # CORREÇÃO AQUI
    
    log_erros = []
    inclusao_inteligente = opcoes.get('inclusao_inteligente', True)
    is_reprocess = opcoes.get('is_reprocess', False)
    caminho_saida_desejado = opcoes.get('caminho_saida')
    
    base_saida = caminho_entrada.replace(".xlsx", "_Analise_Evolucao_PTA").replace(".xlsb", "_Analise_Evolucao_PTA")
    
    # =========================================================================
    # 1. LEITURA DOS PARÂMETROS
    # =========================================================================
    if is_reprocess:
        caminho_pta_atual = opcoes.get('caminho_pta_reprocess')
        try:
            with pd.ExcelFile(caminho_pta_atual) as xl_params:
                if "Parametros" not in xl_params.sheet_names:
                    raise ValueError(f"A aba 'Parametros' não foi encontrada no arquivo {os.path.basename(caminho_pta_atual)}.")
                df_plano_raw = xl_params.parse("Parametros")
            log_erros.append(f"Informação: Leitura de parâmetros realizada a partir do arquivo {os.path.basename(caminho_pta_atual)}.")
        except Exception as e:
            raise ValueError(f"Erro ao ler os parâmetros do PTA editado. Detalhe: {e}")
            
    else:
        with pd.ExcelFile(caminho_entrada) as xl_params:
            nomes_possiveis = ['planodecontas', 'planodeconta', 'planoconta', 'parametros', 'cadastroparametros']
            abas_encontradas = [sht for sht in xl_params.sheet_names if str(sht).lower().replace(" ", "").replace("_", "").split("(")[0].strip() in nomes_possiveis]
            
            if not abas_encontradas:
                raise ValueError("CRÍTICO: Nenhuma aba de Plano de Contas ou Parâmetros encontrada no arquivo original.")
            df_plano_raw = xl_params.parse(abas_encontradas[-1])

    if len(df_plano_raw.columns) >= 6:
        df_plano = df_plano_raw.iloc[:, :7].copy()
        df_plano.columns = ['Chave Cliente', 'Chave D&M', 'Classificação', 'Descrição', 'Sint./An.', 'At/Pas/Res', 'Indice'][:len(df_plano.columns)]
        
        for col in ['Chave Cliente', 'Chave D&M', 'Classificação', 'Descrição', 'Sint./An.', 'At/Pas/Res', 'Indice']:
            if col not in df_plano.columns: df_plano[col] = ""
                
        df_plano['Chave Cliente'] = df_plano['Chave Cliente'].astype(str).str.strip().replace(['nan', 'NAN', 'None'], '')
        df_plano = df_plano[df_plano['Chave Cliente'] != '']
        df_plano['Chave_Clean'] = df_plano['Chave Cliente'].str.replace('.', '', regex=False)
        
        df_plano['Sint./An.'] = df_plano['Sint./An.'].astype(str).str.upper().str.strip().replace(['NAN', 'NULL', 'NONE'], '')
        df_plano['At/Pas/Res'] = df_plano['At/Pas/Res'].astype(str).str.upper().str.strip().replace(['NAN', 'NULL', 'NONE'], '')
        
        if 'Observação' not in df_plano_raw.columns:
            df_plano['Observação'] = ""
        else:
            df_plano['Observação'] = df_plano_raw['Observação'].astype(str).replace('nan', '')
    else:
        raise ValueError("CRÍTICO: A aba de Plano de Contas possui colunas insuficientes.")

    # =========================================================================
    # 2. LEITURA DOS BALANCETES
    # =========================================================================
    df_lista = []
    saldo_dezembro_dit = {}
    
    with pd.ExcelFile(caminho_entrada) as xl_dados:
        abas_meses = [sheet for sheet in xl_dados.sheet_names if str(sheet).isdigit() and 1 <= int(sheet) <= 12]
        if not abas_meses: raise ValueError("Nenhum balancete válido encontrado no arquivo original (01 a 12).")

        for mes in abas_meses:
            df_mes_raw = xl_dados.parse(mes)
            if df_mes_raw.empty and len(df_mes_raw.columns) < 2: continue
                
            if len(df_mes_raw.columns) >= 9:
                df_mes = df_mes_raw.iloc[:, :9].copy()
                df_mes.columns = ['Atividade', 'Conta', 'Descrição', 'Cod. Reduzido', 'Saldo Anterior', 'Débito', 'Crédito', 'Movimento', 'Saldo Acumulado']
                
                for col_num in ['Saldo Anterior', 'Débito', 'Crédito', 'Movimento', 'Saldo Acumulado']:
                    df_mes[col_num] = pd.to_numeric(df_mes[col_num], errors='coerce').fillna(0.0)
                
                df_mes['Mês'] = int(mes)
                df_lista.append(df_mes)
                
                if mes == min(abas_meses, key=int):
                    df_mes['Conta'] = df_mes['Conta'].astype(str).str.strip()
                    saldo_dezembro_dit = dict(zip(df_mes['Conta'], df_mes['Saldo Anterior']))

    df_consolidado = pd.concat(df_lista, ignore_index=True)
    df_consolidado['Conta'] = df_consolidado['Conta'].astype(str).str.strip()
    df_consolidado['Conta_Clean'] = df_consolidado['Conta'].str.replace('.', '', regex=False)
    meses_disponiveis = sorted(df_consolidado['Mês'].unique())
    ultimo_mes = meses_disponiveis[-1]
    num_meses_total = len(meses_disponiveis)

    # =========================================================================
    # 3. VERIFICAÇÃO DE CONTAS ÓRFÃS
    # =========================================================================
    contas_balancete = df_consolidado[['Cod. Reduzido', 'Conta', 'Conta_Clean', 'Descrição']].drop_duplicates(subset=['Conta_Clean'])
    contas_balancete = contas_balancete[contas_balancete['Conta_Clean'] != '']
    
    orfao_mask = ~contas_balancete['Conta_Clean'].isin(df_plano['Chave_Clean'])
    contas_orfas = contas_balancete[orfao_mask].copy()
    
    if not contas_orfas.empty:
        if inclusao_inteligente:
            log_erros.append(f"Informação: {len(contas_orfas)} conta(s) presentes no balancete foram adicionadas ao plano de contas.")
            novas_linhas = pd.DataFrame({
                'Chave Cliente': contas_orfas['Conta'],
                'Chave_Clean': contas_orfas['Conta_Clean'],
                'Chave D&M': '',
                'Classificação': contas_orfas['Cod. Reduzido'], 
                'Descrição': contas_orfas['Descrição'],
                'Sint./An.': '', 
                'At/Pas/Res': '',
                'Indice': 0,
                'Observação': 'Adicionado por classificação automática'
            })
            df_plano = pd.concat([df_plano, novas_linhas], ignore_index=True)

    # =========================================================================
    # 4. ORDENAÇÃO E INTELIGÊNCIA HIERÁRQUICA
    # =========================================================================
    df_plano = df_plano.sort_values('Chave_Clean').reset_index(drop=True)

    if inclusao_inteligente or is_reprocess:
        chaves = df_plano['Chave_Clean']
        proxima_chave = chaves.shift(-1).fillna('')
        
        condicao_sintetica = [(str(prox).startswith(str(atual))) and (str(atual) != '') for atual, prox in zip(chaves, proxima_chave)]
        sint_an_motor = np.where(condicao_sintetica, 'S', 'A')
        
        prefixo = df_plano['Chave_Clean'].str[0]
        apr_motor = np.where(prefixo == '1', 'A', np.where(prefixo == '2', 'P', 'R'))
        
        novos_sa, novos_apr, novas_obs = [], [], []
        
        for u_sa, m_sa, u_apr, m_apr, obs in zip(df_plano['Sint./An.'], sint_an_motor, df_plano['At/Pas/Res'], apr_motor, df_plano['Observação']):
            obs_atual = str(obs) if obs else ""
            
            if u_sa == '': 
                final_sa = m_sa
            elif u_sa == 'A' and m_sa == 'S':
                final_sa = u_sa
                if "Sugestão do sistema: Sintética(S)" not in obs_atual:
                    obs_atual += f" | Sugestão do sistema: Sintética(S)"
            elif u_sa != m_sa and "S/A Apontado:" not in obs_atual:
                final_sa = u_sa
                obs_atual += f" | S/A Apontado: {m_sa}"
            else: 
                final_sa = u_sa
                
            if u_apr == '': final_apr = m_apr
            elif u_apr != m_apr and "A/P/R Apontado:" not in obs_atual:
                final_apr = u_apr; obs_atual += f" | A/P/R Apontado: {m_apr}"
            else: final_apr = u_apr

            novos_sa.append(final_sa)
            novos_apr.append(final_apr)
            novas_obs.append(obs_atual.strip(" | "))
            
        df_plano['Sint./An.'] = novos_sa
        df_plano['At/Pas/Res'] = novos_apr
        df_plano['Observação'] = novas_obs

    df_plano['Indice'] = range(1, len(df_plano) + 1)

    # =========================================================================
    # 5. DIAGNÓSTICO DE MALHA FINA E CALCULO DE BASES (FATURAMENTO)
    # =========================================================================
    df_pivot_mov = df_consolidado.pivot_table(index='Conta', columns='Mês', values='Movimento', aggfunc='sum').fillna(0)
    df_pivot_sld = df_consolidado.pivot_table(index='Conta', columns='Mês', values='Saldo Acumulado', aggfunc='sum').fillna(0)

    def calc_base_val(chaves_input, col_tipo):
        if not chaves_input: return 0.0
        total = 0.0
        chaves_list = [c.strip() for c in str(chaves_input).split(',')]
        for ch in chaves_list:
            contas = df_plano[df_plano[col_tipo] == ch]['Chave Cliente'].tolist()
            for c in contas:
                if c in df_pivot_sld.index:
                    total += df_pivot_sld.loc[c, ultimo_mes]
        return abs(total)

    bases_ativo = [calc_base_val(r['chave'], 'Chave D&M') for r in regras.get('ativo', [])]
    bases_passivo = [calc_base_val(r['chave'], 'Chave Cliente') for r in regras.get('passivo', [])]
    bases_resultado = [calc_base_val(r['chave'], 'Chave D&M') for r in regras.get('resultado', [])]

    df_plano_a = df_plano[df_plano['Sint./An.'] == 'A']
    chaves_a = sorted(df_plano_a['Chave_Clean'].astype(str).tolist())
    pais_conflitantes = []
    
    for i in range(len(chaves_a) - 1):
        c_atual = chaves_a[i]
        c_prox = chaves_a[i+1]
        if c_prox.startswith(c_atual) and c_atual != c_prox:
            pais_conflitantes.append(c_atual)
            
    pais_conflitantes = sorted(list(set(pais_conflitantes)))
    
    if pais_conflitantes:
        log_erros.append(f"Conflito de Parametrização S/A: Foram identificadas {len(pais_conflitantes)} conta(s) marcadas como Analíticas (A) que possuem contas filhas também Analíticas. Ocorrência de duplicidade matemática na soma do balancete.")
        for pai in pais_conflitantes[:5]:
            conta_original = df_plano[df_plano['Chave_Clean'] == pai]['Chave Cliente'].values[0]
            desc_original = df_plano[df_plano['Chave_Clean'] == pai]['Descrição'].values[0]
            log_erros.append(f"Detalhe de Duplicidade S/A: Conta raiz [{conta_original} - {desc_original}].")

    todas_chaves_geral = df_plano['Chave_Clean'].tolist()
    contas_s_vazias = []
    
    for _, row_s in df_plano[df_plano['Sint./An.'] == 'S'].iterrows():
        chave_s = row_s['Chave_Clean']
        filhas_s = [c for c in todas_chaves_geral if c.startswith(chave_s)]
        if len(filhas_s) <= 1:
            contas_s_vazias.append(row_s['Chave Cliente'])
            
    if contas_s_vazias:
        log_erros.append(f"Furo de Estrutura S/A: Identificadas {len(contas_s_vazias)} conta(s) parametrizada(s) como Sintética (S) sem subcontas cadastradas. O saldo atrelado a estas raízes não será computado no Balancete Histórico.")
        for cv in contas_s_vazias[:5]:
            desc_cv = df_plano[df_plano['Chave Cliente'] == cv]['Descrição'].values[0]
            log_erros.append(f"Detalhe de Estrutura S/A: Conta [{cv} - {desc_cv}].")

    df_consolidado['Diff_Intramensal'] = df_consolidado['Saldo Anterior'] + df_consolidado['Movimento'] - df_consolidado['Saldo Acumulado']
    erros_math = df_consolidado[~np.isclose(df_consolidado['Diff_Intramensal'], 0, atol=0.01)]
    if not erros_math.empty:
        log_erros.append(f"Inconsistência Matemática Interna: {len(erros_math)} registro(s) apresentam divergência na equação (Saldo Anterior + Movimento != Saldo Acumulado).")
        erros_math_sorted = erros_math.reindex(erros_math['Diff_Intramensal'].abs().sort_values(ascending=False).index)
        for _, row_e in erros_math_sorted.head(5).iterrows():
            log_erros.append(f"Detalhe de Inconsistência Interna: Mês {row_e['Mês']} | Conta [{row_e['Conta']}] | Diferença de R$ {row_e['Diff_Intramensal']:,.2f}.")

    for i in range(len(meses_disponiveis) - 1):
        m_ant = meses_disponiveis[i]
        m_atu = meses_disponiveis[i+1]
        
        df_ant = df_consolidado[df_consolidado['Mês'] == m_ant][['Conta', 'Descrição', 'Saldo Acumulado']]
        df_atu = df_consolidado[df_consolidado['Mês'] == m_atu][['Conta', 'Saldo Anterior']]
        
        df_merge = pd.merge(df_ant, df_atu, on='Conta', how='inner')
        df_merge['Diff_Intermensal'] = df_merge['Saldo Acumulado'] - df_merge['Saldo Anterior']
        
        quebras = df_merge[~np.isclose(df_merge['Diff_Intermensal'], 0, atol=0.01)]
        if not quebras.empty:
            diff_liquida = quebras['Diff_Intermensal'].sum()
            log_erros.append(f"Quebra de Continuidade Intermensal (Mês {m_ant} -> Mês {m_atu}): Identificadas {len(quebras)} conta(s) com Saldo Inicial diferente do Saldo Final anterior. Divergência líquida consolidada de R$ {diff_liquida:,.2f}.")
            quebras_sorted = quebras.reindex(quebras['Diff_Intermensal'].abs().sort_values(ascending=False).index)
            for _, row_q in quebras_sorted.head(3).iterrows():
                log_erros.append(f"Detalhe de Continuidade: Conta [{row_q['Conta']} - {row_q['Descrição']}] apresenta divergência de R$ {row_q['Diff_Intermensal']:,.2f}.")

    for m in meses_disponiveis:
        soma_balancete = df_consolidado[df_consolidado['Mês'] == m]['Saldo Acumulado'].sum()
        if not np.isclose(soma_balancete, 0, atol=0.01):
            log_erros.append(f"Desbalanceamento Global da Origem: O balancete bruto do mês {m} não totaliza zero. Diferença total apurada: R$ {soma_balancete:,.2f}.")

    df_plano = df_plano.drop(columns=['Chave_Clean'])

    # =========================================================================
    # 6. EXPORTAÇÃO E APLICAÇÃO DAS REGRAS NBC TA 530
    # =========================================================================
    caminho_saida = caminho_saida_desejado
    base_saida_arq, ext = os.path.splitext(caminho_saida)
    counter = 2
    
    while os.path.exists(caminho_saida):
        try:
            with open(caminho_saida, 'r+'): break 
        except PermissionError:
            caminho_saida = f"{base_saida_arq}_v{counter}{ext}"
            counter += 1

    wb = Workbook()
    wb.remove(wb.active)
    fonte_padrao = Font(name='Calibri', size=11)
    fonte_negrito = Font(name='Calibri', size=11, bold=True)
    borda_fina = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    fundo_inclusao = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid") 
    fundo_amarelo = PatternFill(start_color="FFFFFF00", end_color="FFFFFF00", fill_type="solid")

    # ABA 1: Parametros
    ws_param = wb.create_sheet("Parametros")
    colunas_param = ['Chave Cliente', 'Chave D&M', 'Classificação', 'Descrição', 'Sint./An.', 'At/Pas/Res', 'Indice', 'Observação']
    
    for c_idx, col_name in enumerate(colunas_param, 1):
        cel = ws_param.cell(row=1, column=c_idx, value=col_name)
        cel.font = fonte_negrito; cel.border = borda_fina

    for r_idx, row in df_plano.iterrows():
        linha = r_idx + 2
        foi_ad = "Adicionado" in str(row.get('Observação', ''))
        for c_idx, col_name in enumerate(colunas_param, 1):
            val = row.get(col_name, "")
            cel = ws_param.cell(row=linha, column=c_idx, value=val)
            cel.font = fonte_padrao
            if foi_ad: cel.fill = fundo_inclusao

    for col in ws_param.columns: ws_param.column_dimensions[col[0].column_letter].width = 15
    ws_param.column_dimensions['D'].width = 35 
    ws_param.column_dimensions['H'].width = 65 

    # ABA 2: Balancete_Histórico
    ws_hist = wb.create_sheet("Balancete_Histórico")
    ws_hist['G2'] = "Ativo Acumulado"; ws_hist['G3'] = "Passivo Acumulado"
    ws_hist['G4'] = "Resultado Mensal"; ws_hist['G5'] = "Resultado Acumulado"
    ws_hist['G6'] = "Check"
    for r in range(2, 7): ws_hist[f'G{r}'].font = fonte_negrito

    meses_nomes_dict = {1:"Janeiro", 2:"Fevereiro", 3:"Março", 4:"Abril", 5:"Maio", 6:"Junho", 7:"Julho", 8:"Agosto", 9:"Setembro", 10:"Outubro", 11:"Novembro", 12:"Dezembro"}
    
    col_cursor = 8
    for idx, m in enumerate(meses_disponiveis):
        l_mov = get_column_letter(col_cursor)
        l_sld = get_column_letter(col_cursor + 1)
        
        ws_hist[f'{l_mov}1'] = meses_nomes_dict.get(m, f"Mês {m}")
        ws_hist[f'{l_mov}1'].font = fonte_negrito
        
        ws_hist[f'{l_mov}2'] = f'=SUMIFS({l_sld}:{l_sld}, $BB:$BB, "A", $BC:$BC, "A")'
        ws_hist[f'{l_mov}3'] = f'=SUMIFS({l_sld}:{l_sld}, $BB:$BB, "A", $BC:$BC, "P")'
        ws_hist[f'{l_mov}4'] = f'=SUMIFS({l_mov}:{l_mov}, $BB:$BB, "A", $BC:$BC, "R")'
        
        if idx == 0: 
            ws_hist[f'{l_mov}5'] = f'={l_mov}4'
        else: 
            ws_hist[f'{l_mov}5'] = f'={get_column_letter(col_cursor - 2)}5+{l_mov}4'
            
        ws_hist[f'{l_mov}6'] = f'={l_mov}2+{l_mov}3+{l_mov}5'
        
        for r in range(2, 7):
            ws_hist[f'{l_mov}{r}'].font = fonte_negrito
            ws_hist[f'{l_mov}{r}'].number_format = '#,##0.00'
        col_cursor += 2

    headers_hist = ['Seleção', 'Atividade', 'Chave', 'Conta', 'Descrição', 'Cod. Reduzido', 'Acum. Dez']
    for m in meses_disponiveis: headers_hist.extend(['Movimento', 'Saldo Acumulado'])
    
    for c_idx, h_text in enumerate(headers_hist, 1):
        cel = ws_hist.cell(row=8, column=c_idx, value=h_text)
        cel.font = fonte_negrito; cel.border = borda_fina

    linha_inicio = 9
    for r_idx, row_p in df_plano.iterrows():
        current_row = linha_inicio + r_idx
        conta_cod = str(row_p['Chave Cliente']).strip()
        
        # Inteligência de Amostragem NBC TA 530
        selecao_val = ""
        cols_amarelas = []
        
        if row_p['Sint./An.'] == 'A':
            saldo_final_abs = abs(df_pivot_sld.loc[conta_cod, ultimo_mes]) if conta_cod in df_pivot_sld.index else 0.0
            
            if row_p['At/Pas/Res'] == 'A' and opcoes.get('ativo', False):
                for i, r_dict in enumerate(regras.get('ativo', [])):
                    if not r_dict['min']: continue
                    p_min = parse_perc(r_dict['min'])
                    p_max = parse_perc(r_dict['max'])
                    base = bases_ativo[i]
                    if (base * p_min) < saldo_final_abs <= (base * p_max if p_max != float('inf') else float('inf')):
                        selecao_val = f"X-A{i+1}"
                        cols_amarelas.append(1)
                        break
                        
            elif row_p['At/Pas/Res'] == 'P' and opcoes.get('passivo', False):
                for i, r_dict in enumerate(regras.get('passivo', [])):
                    if not r_dict['min']: continue
                    p_min = parse_perc(r_dict['min'])
                    p_max = parse_perc(r_dict['max'])
                    base = bases_passivo[i]
                    if (base * p_min) < saldo_final_abs <= (base * p_max if p_max != float('inf') else float('inf')):
                        selecao_val = f"X-P{i+1}"
                        cols_amarelas.append(1)
                        break
                        
            elif row_p['At/Pas/Res'] == 'R' and opcoes.get('resultado', False):
                movimentos = [df_pivot_mov.loc[conta_cod, m] if conta_cod in df_pivot_mov.index and m in df_pivot_mov.columns else 0.0 for m in meses_disponiveis]
                
                # Inteligência de Média: Conta quando a conta "nasceu" (primeiro saldo/mov != 0)
                idx_primeiro_mes = -1
                for i_m, m in enumerate(meses_disponiveis):
                    sld = df_pivot_sld.loc[conta_cod, m] if conta_cod in df_pivot_sld.index else 0.0
                    mov = df_pivot_mov.loc[conta_cod, m] if conta_cod in df_pivot_mov.index else 0.0
                    if abs(sld) > 0.01 or abs(mov) > 0.01:
                        idx_primeiro_mes = i_m
                        break
                
                meses_ativos = num_meses_total - idx_primeiro_mes if idx_primeiro_mes != -1 else 0
                media = sum(movimentos) / meses_ativos if meses_ativos > 0 else 0.0
                
                for i, r_dict in enumerate(regras.get('resultado', [])):
                    if not r_dict['min']: continue
                    p_min = parse_perc(r_dict['min'])
                    p_max = parse_perc(r_dict['max'])
                    p_desvio = parse_perc(r_dict['perc'])
                    base = bases_resultado[i]
                    
                    if (base * p_min) < saldo_final_abs <= (base * p_max if p_max != float('inf') else float('inf')):
                        if media != 0:
                            matched = False
                            col_cursor_temp = 8
                            for mov_val in movimentos:
                                desvio = (mov_val / media) - 1
                                if abs(desvio) >= p_desvio:
                                    matched = True
                                    cols_amarelas.append(col_cursor_temp)
                                col_cursor_temp += 2
                            if matched:
                                selecao_val = f"X-R{i+1}"
                                if 1 not in cols_amarelas: cols_amarelas.append(1)
                                break
        
        c_sel = ws_hist.cell(row=current_row, column=1, value=selecao_val)
        if 1 in cols_amarelas: c_sel.fill = fundo_amarelo
        
        ws_hist.cell(row=current_row, column=2, value="Geral") 
        ws_hist.cell(row=current_row, column=3, value=conta_cod) 
        ws_hist.cell(row=current_row, column=4, value=row_p['Classificação']) 
        ws_hist.cell(row=current_row, column=5, value=row_p['Descrição']) 
        ws_hist.cell(row=current_row, column=6, value=row_p['Chave D&M']) 
        ws_hist.cell(row=current_row, column=7, value=saldo_dezembro_dit.get(conta_cod, 0.0)).number_format = '#,##0.00'
        
        col_cursor = 8
        for m in meses_disponiveis:
            mov_val = df_pivot_mov.loc[conta_cod, m] if conta_cod in df_pivot_mov.index and m in df_pivot_mov.columns else 0.0
            sld_val = df_pivot_sld.loc[conta_cod, m] if conta_cod in df_pivot_sld.index and m in df_pivot_sld.columns else 0.0
            
            c_mov = ws_hist.cell(row=current_row, column=col_cursor, value=mov_val)
            c_mov.number_format = '#,##0.00'
            if col_cursor in cols_amarelas: c_mov.fill = fundo_amarelo
                
            ws_hist.cell(row=current_row, column=col_cursor+1, value=sld_val).number_format = '#,##0.00'
            col_cursor += 2

        ws_hist.cell(row=current_row, column=54, value=row_p['Sint./An.'])
        ws_hist.cell(row=current_row, column=55, value=row_p['At/Pas/Res'])

    for col in ws_hist.columns:
        if col[0].column < 54: ws_hist.column_dimensions[get_column_letter(col[0].column)].width = 15
    ws_hist.column_dimensions['E'].width = 35 

    # ABA 3: Log
    ws_log = wb.create_sheet("Log_Diagnostico")
    ws_log['A1'] = "Relatório de Diagnóstico Analítico"
    ws_log['A1'].font = fonte_negrito
    if log_erros:
        for i, erro in enumerate(log_erros, start=3): ws_log.cell(row=i, column=1, value=erro).font = fonte_padrao
    else:
         ws_log.cell(row=3, column=1, value="Nenhuma inconsistência primária identificada no processamento estrutural.").font = fonte_padrao

    wb.save(caminho_saida)
    wb.close() 

    # =========================================================================
    # 7. ATUALIZA A PLANILHA ORIGINAL
    # =========================================================================
    if inclusao_inteligente or is_reprocess:
        try:
            wb_orig = load_workbook(caminho_entrada)
            base_name = "Parametros"
            sheet_name = base_name
            counter = 1
            while sheet_name in wb_orig.sheetnames:
                sheet_name = f"{base_name} ({counter})"
                counter += 1
                
            ws_orig_param = wb_orig.create_sheet(title=sheet_name)
            for c_idx, col_name in enumerate(colunas_param, 1):
                ws_orig_param.cell(row=1, column=c_idx, value=col_name).font = fonte_negrito
                
            for r_idx, row in df_plano.iterrows():
                for c_idx, col_name in enumerate(colunas_param, 1):
                    ws_orig_param.cell(row=r_idx+2, column=c_idx, value=row.get(col_name, ""))
                    
            wb_orig.save(caminho_entrada)
            wb_orig.close()
        except Exception:
            pass
    
    return caminho_saida, log_erros