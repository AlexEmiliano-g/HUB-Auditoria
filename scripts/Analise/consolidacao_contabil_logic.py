# ==============================================================================
# scripts/Analise/consolidacao_contabil_logic.py
# Módulo de Regras de Negócio: Consolidação e Evolução Contábil
# ==============================================================================

import pandas as pd
import os
import traceback
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.formula.translate import Translator
from openpyxl.styles import Font
import copy
import re

# Nomes de abas reservadas para configuração do sistema
ABAS_SISTEMA = ['Cadastro_Empresas']

def higienizar_nome_aba(nome):
    """ Remove caracteres proibidos pelo Excel e limita a 31 caracteres """
    nome_limpo = re.sub(r'[\\/\?\*\[\]]', '', str(nome))
    return nome_limpo[:31].strip()

def copiar_estilo(celula_origem, celula_destino):
    """ Copia todas as propriedades de formatação visual de uma célula para outra """
    if celula_origem.font:
        celula_destino.font = copy.copy(celula_origem.font)
    if celula_origem.border:
        celula_destino.border = copy.copy(celula_origem.border)
    if celula_origem.fill:
        celula_destino.fill = copy.copy(celula_origem.fill)
    if celula_origem.number_format:
        celula_destino.number_format = copy.copy(celula_origem.number_format)
    if celula_origem.protection:
        celula_destino.protection = copy.copy(celula_origem.protection)
    if celula_origem.alignment:
        celula_destino.alignment = copy.copy(celula_origem.alignment)

def exportar_templates(pasta_destino):
    """ Gera o modelo de Template Mestre para preenchimento do usuário """
    try:
        caminho_params = os.path.join(pasta_destino, "Template_Mestre_Consolidacao.xlsx")
        
        df_cad = pd.DataFrame([
            {"Empresa": "Empresa Alfa Ltda", "CNPJ": "11.111.111/0001-11", "Arquivo_Balancete": "Balancete_Exemplo_EmpresaAlfa.xlsx", "Plano_Utilizado": "Plano_1"},
            {"Empresa": "Empresa Beta S/A", "CNPJ": "22.222.222/0001-22", "Arquivo_Balancete": "Balancete_Exemplo_EmpresaBeta.xlsx", "Plano_Utilizado": "Plano_1"}
        ])
        
        df_p1 = pd.DataFrame([
            {"Conta_Analitica": "11601.01", "Descricao_Conta": "ICMS a Recuperar", "Codigo_Sintetico_BP": "11601"},
            {"Conta_Analitica": "12101.01", "Descricao_Conta": "Outros Créditos LP", "Codigo_Sintetico_BP": "12101"},
            {"Conta_Analitica": "12102.01", "Descricao_Conta": "Tributos a Recuperar LP", "Codigo_Sintetico_BP": "12102"}
        ])

        with pd.ExcelWriter(caminho_params, engine='openpyxl') as writer:
            df_cad.to_excel(writer, sheet_name="Cadastro_Empresas", index=False)
            df_p1.to_excel(writer, sheet_name="Plano_1", index=False)

        wb = openpyxl.load_workbook(caminho_params)
        ws_balanco = wb.create_sheet("Balanço Patrimonial")
        
        ws_balanco.append(["Código", "Nome da Conta", "Mês Base (O robô copia daqui)"])
        ws_balanco.append(["", "Realizável a Longo Prazo", "=SUM(C4:C8)"])
        ws_balanco.append(["12101", "Outros Créditos", ""])
        ws_balanco.append(["12102", "Tributos a Recuperar", ""])
        ws_balanco.append(["12103", "Depósitos e Bloqueios Judiciais", ""])
        ws_balanco.append(["12105", "Tributos Diferidos", ""])
        ws_balanco.append(["12106", "Outros Valores a Receber", ""])

        font_negrito = Font(bold=True)
        for col_idx in range(1, 4):
            ws_balanco.cell(row=1, column=col_idx).font = font_negrito
            ws_balanco.cell(row=2, column=col_idx).font = font_negrito

        for cell in ws_balanco["B"]: cell.number_format = '@'
        ws_balanco.column_dimensions['B'].width = 35
        ws_balanco.column_dimensions['C'].width = 20
        wb.save(caminho_params)

        caminho_balancete = os.path.join(pasta_destino, "Balancete_Exemplo_EmpresaAlfa.xlsx")
        df_balancete = pd.DataFrame([
            {"Atividade": "Geral", "Conta": "11601.01", "Nome": "ICMS a Recuperar", "Cód Reduzido": "101", "Saldo Anterior": 0.0, "Débito": 645899.24, "Crédito": 0.0, "Movimento": 645899.24, "Saldo Acumulado": 645899.24},
            {"Atividade": "Geral", "Conta": "12101.01", "Nome": "Outros Créditos LP", "Cód Reduzido": "102", "Saldo Anterior": 1000.00, "Débito": 4797164.60, "Crédito": 0.0, "Movimento": 4797164.60, "Saldo Acumulado": 4798164.60}
        ])
        with pd.ExcelWriter(caminho_balancete, engine='openpyxl') as writer:
            df_balancete.to_excel(writer, sheet_name="01", index=False) 
            df_balancete.to_excel(writer, sheet_name="02", index=False) 

        return True, f"Templates gerados com sucesso."
    except Exception as e:
        return False, f"Ocorreu um erro: {str(e)}"


def executar_validacao_contas(pasta_balancetes, arquivo_parametros, pasta_destino):
    """ Valida a estrutura inicial de arquivos contábeis """
    try:
        xls_params = pd.ExcelFile(arquivo_parametros)
        if 'Cadastro_Empresas' not in xls_params.sheet_names:
            return False, "ERRO: Aba 'Cadastro_Empresas' não localizada."
        return True, "Validação concluída com sucesso! Estrutura pronta para consolidação."
    except Exception as e:
        return False, f"Erro na validação: {str(e)}"


def gerar_relatorios_consolidacao(pasta_balancetes, arquivo_parametros, pasta_destino):
    """
    PROCESSAMENTO DAS 3 PLANILHAS DE ANÁLISE HORIZONTAL (COM 3 COLUNAS POR PERÍODO).
    Extrai Anterior, Movimento e Acumulado, ocultando o Anterior e Acumulado visualmente.
    """
    try:
        # 1. Carregar Parâmetros de Cadastro e Mapeamentos De-Para (Planos)
        xls_params = pd.ExcelFile(arquivo_parametros)
        df_cad = pd.read_excel(xls_params, sheet_name='Cadastro_Empresas')
        
        lista_empresas = df_cad['Empresa'].dropna().unique().tolist()
        
        mapeamento_geral = {}
        for plano in df_cad['Plano_Utilizado'].dropna().unique():
            if plano in xls_params.sheet_names:
                df_plano = pd.read_excel(xls_params, sheet_name=plano, dtype={'Conta_Analitica': str, 'Codigo_Sintetico_BP': str})
                df_plano['Conta_Analitica'] = df_plano['Conta_Analitica'].str.strip()
                df_plano['Codigo_Sintetico_BP'] = df_plano['Codigo_Sintetico_BP'].str.strip()
                mapeamento_geral[plano] = dict(zip(df_plano['Conta_Analitica'], df_plano['Codigo_Sintetico_BP']))

        # 2. Varredura e Consolidação dos Balancetes na Memória
        # Estrutura: dados[empresa][mes][codigo_mestre] = {'SA': 0.0, 'MOV': 0.0, 'SAC': 0.0}
        dados_multidimensionais = {}
        meses_detectados = set()

        for idx, row in df_cad.iterrows():
            empresa = str(row.get('Empresa')).strip()
            arquivo_balancete = row.get('Arquivo_Balancete')
            plano_nome = row.get('Plano_Utilizado')
            if pd.isna(arquivo_balancete) or pd.isna(plano_nome): continue
            
            caminho_bal = os.path.join(pasta_balancetes, str(arquivo_balancete))
            if not os.path.exists(caminho_bal): continue
            
            if empresa not in dados_multidimensionais:
                dados_multidimensionais[empresa] = {}
                
            xls_bal = pd.ExcelFile(caminho_bal)
            abas_meses = [aba for aba in xls_bal.sheet_names if str(aba).strip().isdigit() and len(str(aba).strip()) == 2]
            map_plano = mapeamento_geral.get(plano_nome, {})
            
            for mes in abas_meses:
                mes_str = str(mes).strip()
                meses_detectados.add(mes_str)
                if mes_str not in dados_multidimensionais[empresa]:
                    dados_multidimensionais[empresa][mes_str] = {}
                
                df_mes = pd.read_excel(xls_bal, sheet_name=mes, dtype={'Conta': str})
                
                # Garante que as 3 colunas vitais existam e as converte para números
                for col in ['Saldo Anterior', 'Movimento', 'Saldo Acumulado']:
                    if col in df_mes.columns:
                        df_mes[col] = pd.to_numeric(df_mes[col], errors='coerce').fillna(0.0)
                    else:
                        df_mes[col] = 0.0
                
                for _, row_dados in df_mes.iterrows():
                    conta_analitica = str(row_dados.get('Conta', '')).strip()
                    if not conta_analitica or conta_analitica == 'nan': continue

                    codigo_mestre = map_plano.get(conta_analitica, conta_analitica)
                    
                    if codigo_mestre not in dados_multidimensionais[empresa][mes_str]:
                        dados_multidimensionais[empresa][mes_str][codigo_mestre] = {'SA': 0.0, 'MOV': 0.0, 'SAC': 0.0}
                        
                    dados_multidimensionais[empresa][mes_str][codigo_mestre]['SA'] += row_dados['Saldo Anterior']
                    dados_multidimensionais[empresa][mes_str][codigo_mestre]['MOV'] += row_dados['Movimento']
                    dados_multidimensionais[empresa][mes_str][codigo_mestre]['SAC'] += row_dados['Saldo Acumulado']

        lista_meses = sorted(list(meses_detectados))
        if not lista_meses:
            return False, "Nenhum dado válido de mês foi localizado para consolidar."

        wb_template = openpyxl.load_workbook(arquivo_parametros)
        abas_relatorios = [aba for aba in wb_template.sheetnames if aba not in ABAS_SISTEMA and not aba.startswith("Plano_")]

        # =====================================================================
        # PLANILHA 1: COMPARATIVO ENTRE EMPRESAS POR MÊS
        # =====================================================================
        wb_rel1 = openpyxl.Workbook()
        wb_rel1.remove(wb_rel1.active)

        for mes in lista_meses:
            for nome_aba_tmpl in abas_relatorios:
                ws_origem = wb_template[nome_aba_tmpl]
                ws_destino = wb_rel1.create_sheet(title=higienizar_nome_aba(f"{mes}_{nome_aba_tmpl}"))
                
                for col_letter in ['A', 'B', 'C']:
                    if col_letter in ws_origem.column_dimensions:
                        ws_destino.column_dimensions[col_letter].width = ws_origem.column_dimensions[col_letter].width

                for row_idx, row in enumerate(ws_origem.iter_rows(values_only=False), start=1):
                    cell_a_orig, cell_b_orig = row[0], row[1]
                    cell_c_orig = row[2] if len(row) > 2 else None
                    codigo_conta = str(cell_a_orig.value).strip() if cell_a_orig.value else ""
                    
                    copiar_estilo(cell_a_orig, ws_destino.cell(row=row_idx, column=1, value=cell_a_orig.value))
                    copiar_estilo(cell_b_orig, ws_destino.cell(row=row_idx, column=2, value=cell_b_orig.value))

                    # Cria 3 colunas para cada Empresa
                    for emp_offset, empresa in enumerate(lista_empresas):
                        col_sa = 3 + (emp_offset * 3)
                        col_mov = col_sa + 1
                        col_sac = col_sa + 2

                        cell_sa = ws_destino.cell(row=row_idx, column=col_sa)
                        cell_mov = ws_destino.cell(row=row_idx, column=col_mov)
                        cell_sac = ws_destino.cell(row=row_idx, column=col_sac)

                        if cell_c_orig:
                            copiar_estilo(cell_c_orig, cell_sa)
                            copiar_estilo(cell_c_orig, cell_mov)
                            copiar_estilo(cell_c_orig, cell_sac)

                        if row_idx == 1:
                            # Ajustar a largura e Ocultar as colunas
                            ws_destino.column_dimensions[get_column_letter(col_sa)].width = 18
                            ws_destino.column_dimensions[get_column_letter(col_mov)].width = 18
                            ws_destino.column_dimensions[get_column_letter(col_sac)].width = 18
                            
                            ws_destino.column_dimensions[get_column_letter(col_sa)].hidden = True
                            ws_destino.column_dimensions[get_column_letter(col_sac)].hidden = True

                            cell_sa.value = f"{empresa} - Anterior"
                            cell_mov.value = f"{empresa} - Movimento"
                            cell_sac.value = f"{empresa} - Final"
                            continue
                            
                        tem_formula = cell_c_orig and cell_c_orig.data_type == 'f'
                        if not tem_formula and codigo_conta and codigo_conta != "None":
                            dados_conta = dados_multidimensionais.get(empresa, {}).get(mes, {}).get(codigo_conta, {})
                            if dados_conta.get('SA', 0) != 0: cell_sa.value = dados_conta['SA']
                            if dados_conta.get('MOV', 0) != 0: cell_mov.value = dados_conta['MOV']
                            if dados_conta.get('SAC', 0) != 0: cell_sac.value = dados_conta['SAC']
                        
                        elif tem_formula:
                            cell_sa.value = Translator(cell_c_orig.value, origin=f"C{row_idx}").translate_formula(f"{get_column_letter(col_sa)}{row_idx}")
                            cell_mov.value = Translator(cell_c_orig.value, origin=f"C{row_idx}").translate_formula(f"{get_column_letter(col_mov)}{row_idx}")
                            cell_sac.value = Translator(cell_c_orig.value, origin=f"C{row_idx}").translate_formula(f"{get_column_letter(col_sac)}{row_idx}")

        wb_rel1.save(os.path.join(pasta_destino, "Relatorio_1_Comparativo_Empresas_Por_Mes.xlsx"))

        # =====================================================================
        # PLANILHA 2: CONSOLIDAÇÃO GLOBAL
        # =====================================================================
        wb_rel2 = openpyxl.Workbook()
        wb_rel2.remove(wb_rel2.active)

        for nome_aba_tmpl in abas_relatorios:
            ws_origem = wb_template[nome_aba_tmpl]
            ws_destino = wb_rel2.create_sheet(title=higienizar_nome_aba(nome_aba_tmpl))
            
            for col_letter in ['A', 'B', 'C']:
                if col_letter in ws_origem.column_dimensions:
                    ws_destino.column_dimensions[col_letter].width = ws_origem.column_dimensions[col_letter].width

            for row_idx, row in enumerate(ws_origem.iter_rows(values_only=False), start=1):
                cell_a_orig, cell_b_orig = row[0], row[1]
                cell_c_orig = row[2] if len(row) > 2 else None
                codigo_conta = str(cell_a_orig.value).strip() if cell_a_orig.value else ""
                
                copiar_estilo(cell_a_orig, ws_destino.cell(row=row_idx, column=1, value=cell_a_orig.value))
                copiar_estilo(cell_b_orig, ws_destino.cell(row=row_idx, column=2, value=cell_b_orig.value))

                # Cria 3 colunas para cada Mês
                for m_offset, mes in enumerate(lista_meses):
                    col_sa = 3 + (m_offset * 3)
                    col_mov = col_sa + 1
                    col_sac = col_sa + 2

                    cell_sa = ws_destino.cell(row=row_idx, column=col_sa)
                    cell_mov = ws_destino.cell(row=row_idx, column=col_mov)
                    cell_sac = ws_destino.cell(row=row_idx, column=col_sac)

                    if cell_c_orig:
                        copiar_estilo(cell_c_orig, cell_sa)
                        copiar_estilo(cell_c_orig, cell_mov)
                        copiar_estilo(cell_c_orig, cell_sac)

                    if row_idx == 1:
                        ws_destino.column_dimensions[get_column_letter(col_sa)].width = 18
                        ws_destino.column_dimensions[get_column_letter(col_mov)].width = 18
                        ws_destino.column_dimensions[get_column_letter(col_sac)].width = 18
                        
                        ws_destino.column_dimensions[get_column_letter(col_sa)].hidden = True
                        ws_destino.column_dimensions[get_column_letter(col_sac)].hidden = True

                        cell_sa.value = f"Mês {mes} - Anterior"
                        cell_mov.value = f"Mês {mes} - Movimento"
                        cell_sac.value = f"Mês {mes} - Final"
                        continue
                        
                    tem_formula = cell_c_orig and cell_c_orig.data_type == 'f'
                    if not tem_formula and codigo_conta and codigo_conta != "None":
                        soma_sa = sum(dados_multidimensionais[emp][mes].get(codigo_conta, {}).get('SA', 0.0) for emp in dados_multidimensionais if mes in dados_multidimensionais[emp])
                        soma_mov = sum(dados_multidimensionais[emp][mes].get(codigo_conta, {}).get('MOV', 0.0) for emp in dados_multidimensionais if mes in dados_multidimensionais[emp])
                        soma_sac = sum(dados_multidimensionais[emp][mes].get(codigo_conta, {}).get('SAC', 0.0) for emp in dados_multidimensionais if mes in dados_multidimensionais[emp])

                        if soma_sa != 0: cell_sa.value = soma_sa
                        if soma_mov != 0: cell_mov.value = soma_mov
                        if soma_sac != 0: cell_sac.value = soma_sac
                    
                    elif tem_formula:
                        cell_sa.value = Translator(cell_c_orig.value, origin=f"C{row_idx}").translate_formula(f"{get_column_letter(col_sa)}{row_idx}")
                        cell_mov.value = Translator(cell_c_orig.value, origin=f"C{row_idx}").translate_formula(f"{get_column_letter(col_mov)}{row_idx}")
                        cell_sac.value = Translator(cell_c_orig.value, origin=f"C{row_idx}").translate_formula(f"{get_column_letter(col_sac)}{row_idx}")

        wb_rel2.save(os.path.join(pasta_destino, "Relatorio_2_Consolidado_Global.xlsx"))

        # =====================================================================
        # PLANILHA 3: EVOLUÇÃO INDIVIDUAL POR EMPRESA
        # =====================================================================
        wb_rel3 = openpyxl.Workbook()
        wb_rel3.remove(wb_rel3.active)

        for empresa in lista_empresas:
            for nome_aba_tmpl in abas_relatorios:
                ws_origem = wb_template[nome_aba_tmpl]
                ws_destino = wb_rel3.create_sheet(title=higienizar_nome_aba(f"{empresa}_{nome_aba_tmpl}"))
                
                for col_letter in ['A', 'B', 'C']:
                    if col_letter in ws_origem.column_dimensions:
                        ws_destino.column_dimensions[col_letter].width = ws_origem.column_dimensions[col_letter].width

                for row_idx, row in enumerate(ws_origem.iter_rows(values_only=False), start=1):
                    cell_a_orig, cell_b_orig = row[0], row[1]
                    cell_c_orig = row[2] if len(row) > 2 else None
                    codigo_conta = str(cell_a_orig.value).strip() if cell_a_orig.value else ""
                    
                    copiar_estilo(cell_a_orig, ws_destino.cell(row=row_idx, column=1, value=cell_a_orig.value))
                    copiar_estilo(cell_b_orig, ws_destino.cell(row=row_idx, column=2, value=cell_b_orig.value))

                    for m_offset, mes in enumerate(lista_meses):
                        col_sa = 3 + (m_offset * 3)
                        col_mov = col_sa + 1
                        col_sac = col_sa + 2

                        cell_sa = ws_destino.cell(row=row_idx, column=col_sa)
                        cell_mov = ws_destino.cell(row=row_idx, column=col_mov)
                        cell_sac = ws_destino.cell(row=row_idx, column=col_sac)

                        if cell_c_orig:
                            copiar_estilo(cell_c_orig, cell_sa)
                            copiar_estilo(cell_c_orig, cell_mov)
                            copiar_estilo(cell_c_orig, cell_sac)

                        if row_idx == 1:
                            ws_destino.column_dimensions[get_column_letter(col_sa)].width = 18
                            ws_destino.column_dimensions[get_column_letter(col_mov)].width = 18
                            ws_destino.column_dimensions[get_column_letter(col_sac)].width = 18
                            
                            ws_destino.column_dimensions[get_column_letter(col_sa)].hidden = True
                            ws_destino.column_dimensions[get_column_letter(col_sac)].hidden = True

                            cell_sa.value = f"Mês {mes} - Anterior"
                            cell_mov.value = f"Mês {mes} - Movimento"
                            cell_sac.value = f"Mês {mes} - Final"
                            continue
                            
                        tem_formula = cell_c_orig and cell_c_orig.data_type == 'f'
                        if not tem_formula and codigo_conta and codigo_conta != "None":
                            dados_conta = dados_multidimensionais.get(empresa, {}).get(mes, {}).get(codigo_conta, {})
                            if dados_conta.get('SA', 0) != 0: cell_sa.value = dados_conta['SA']
                            if dados_conta.get('MOV', 0) != 0: cell_mov.value = dados_conta['MOV']
                            if dados_conta.get('SAC', 0) != 0: cell_sac.value = dados_conta['SAC']
                        
                        elif tem_formula:
                            cell_sa.value = Translator(cell_c_orig.value, origin=f"C{row_idx}").translate_formula(f"{get_column_letter(col_sa)}{row_idx}")
                            cell_mov.value = Translator(cell_c_orig.value, origin=f"C{row_idx}").translate_formula(f"{get_column_letter(col_mov)}{row_idx}")
                            cell_sac.value = Translator(cell_c_orig.value, origin=f"C{row_idx}").translate_formula(f"{get_column_letter(col_sac)}{row_idx}")

        wb_rel3.save(os.path.join(pasta_destino, "Relatorio_3_Evolucao_Individual_Empresas.xlsx"))

        return True, f"As 3 planilhas analíticas horizontais (com Colunas Ocultas e Visuais Mantidos) foram geradas com sucesso!"

    except Exception as e:
        return False, f"Erro técnico na consolidação: {str(e)}\n\nDetalhes:\n{traceback.format_exc()}"