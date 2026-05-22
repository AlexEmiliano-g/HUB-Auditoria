# Caminho do Código: /scripts/Analise/analise_fiscal_logic.py
# (Sem alterações nesta etapa)

import pandas as pd
from datetime import datetime
import os
import numpy as np # Importado para np.cumsum

# ==============================================================================
# ETAPA 1: LÓGICA DE PROCESSAMENTO DO ARQUIVO SPED EFD CONTRIBUIÇÕES (BLOCO C)
# (Baseado no script relatorio_C100_EFD_C_app.py)
# ==============================================================================

def parse_line(line):
    """Divide uma linha do SPED em campos."""
    line = line.strip()
    if line and line.startswith('|') and line.endswith('|'):
        return line[1:-1].split('|')
    return []

def safe_get(data_list, index, default=''):
    """Acessa um índice de lista com segurança, retornando default se falhar."""
    return data_list[index] if len(data_list) > index else default

def safe_float_convert(value_str):
    """Converte string para float com segurança (tratando ',', '.' e vazios)."""
    if not value_str: return 0.0
    try:
        # Primeiro remove pontos de milhar, depois substitui vírgula decimal por ponto
        return float(value_str.replace('.', '').replace(',', '.'))
    except (ValueError, TypeError):
        return 0.0

def format_date(date_str):
    """Formata data DDMMAAAA para DD/MM/AAAA."""
    if date_str and len(date_str) == 8 and date_str.isdigit():
        try:
            return datetime.strptime(date_str, '%d%m%Y').strftime('%d/%m/%Y')
        except ValueError:
            return '' # Retorna vazio se a data for inválida
    return ''

# Dicionário para traduzir códigos de tipo de item
TIPO_ITEM_MAP = {
    '00': '00 - Mercadoria para Revenda', '01': '01 - Matéria-Prima',
    '02': '02 - Embalagem', '03': '03 - Produto em Processo',
    '04': '04 - Produto Acabado', '05': '05 - Subproduto',
    '06': '06 - Produto Intermediário', '07': '07 - Material de Uso e Consumo',
    '08': '08 - Ativo Imobilizado', '09': '09 - Serviços',
    '10': '10 - Outros Insumos', '99': '99 - Outras'
}

def processar_bloco_nota_contrib(c100_data, temp_c170_list, temp_c190_list, mapas):
    """
    Processa os itens (C170) ou o resumo analítico (C190) de UMA nota fiscal (C100).
    Prioriza C170 se ambos existirem.
    Retorna uma lista de dicionários, cada um representando uma linha do relatório final.
    """
    linhas_geradas = []
    mapa_produtos = mapas.get('produtos', {})

    if temp_c170_list:
        for campos_c170 in temp_c170_list:
            cod_item = safe_get(campos_c170, 2)
            produto = mapa_produtos.get(cod_item, {})
            tipo_item_cod = produto.get('Tipo Item', '')

            # Monta a linha com dados do C100 e C170
            row = {
                **c100_data, # Inclui todos os dados já extraídos do C100
                'Origem': 'C170 (Item)',
                'Código Item': cod_item,
                'Descrição Item': produto.get('Descrição Item', ''),
                'NCM': produto.get('NCM', ''),
                'Tipo Item': TIPO_ITEM_MAP.get(tipo_item_cod, tipo_item_cod),
                'CFOP': safe_get(campos_c170, 10),
                'Vlr Item': safe_float_convert(safe_get(campos_c170, 6)),
                # Adicione outros campos do C170 se necessário para análise futura
            }
            row['Vlr Operação'] = row.get('Vlr Operação', 0.0)
            linhas_geradas.append(row)

    elif temp_c190_list:
        for campos_c190 in temp_c190_list:
            row = {
                **c100_data,
                'Origem': 'C190 (Resumo)',
                'CFOP': safe_get(campos_c190, 2),
                'Vlr Operação': safe_float_convert(safe_get(campos_c190, 4)),
                'Código Item': '', 'Descrição Item': '', 'NCM': '', 'Tipo Item': '', 'Vlr Item': 0.0,
            }
            linhas_geradas.append(row)

    elif not temp_c170_list and not temp_c190_list and c100_data:
         row = {
                **c100_data,
                'Origem': 'C100 (Cabeçalho)',
                'CFOP': '', 'Vlr Item': 0.0, 'Vlr Operação': 0.0,
                'Código Item': '', 'Descrição Item': '', 'NCM': '', 'Tipo Item': '',
            }
         linhas_geradas.append(row)

    return linhas_geradas

def gerar_relatorio_efd_contrib(lista_arquivos):
    """
    Função principal que lê múltiplos arquivos SPED EFD Contribuições (.txt),
    extrai dados dos registros 0000, 0150, 0200, C010, C100, C170 e C190,
    e retorna um DataFrame pandas consolidado focado no Bloco C.
    """
    if not lista_arquivos:
        return pd.DataFrame()

    mapa_empresa, mapa_participantes, mapa_produtos = {}, {}, {}
    all_report_rows = []

    print("Iniciando Etapa 1: Mapeamento de dados de apoio (0150, 0200)...")
    for filepath in lista_arquivos:
        nome_arquivo = os.path.basename(filepath)
        print(f"  - Mapeando arquivo: {nome_arquivo}")
        try:
            with open(filepath, 'r', encoding='latin-1') as f:
                for line in f:
                    campos = parse_line(line)
                    if not campos: continue
                    reg = campos[0]

                    if reg == '0000' and nome_arquivo not in mapa_empresa:
                         mapa_empresa[nome_arquivo] = {
                             'Nome Empresa Declarante': safe_get(campos, 6),
                             'Período': f"{format_date(safe_get(campos, 4))} a {format_date(safe_get(campos, 5))}"
                         }

                    elif reg == '0150':
                        cod_part = safe_get(campos, 1) # Código interno (índice 1)
                        if cod_part and cod_part not in mapa_participantes:
                            # --- (CORREÇÃO CNPJ/CPF - V3 - ÍNDICES CORRETOS) ---
                            cnpj = safe_get(campos, 4) # CNPJ é o campo de índice 4
                            cpf = safe_get(campos, 5)  # CPF é o campo de índice 5
                            # --- (FIM DA CORREÇÃO V3) ---
                            identificador_real = cnpj if cnpj else (cpf if cpf else cod_part)
                            mapa_participantes[cod_part] = {
                                'Nome Participante': safe_get(campos, 2), # Nome (índice 2)
                                'CNPJ/CPF Participante': identificador_real
                            }

                    elif reg == '0200':
                        cod_item = safe_get(campos, 1)
                        if cod_item and cod_item not in mapa_produtos:
                            mapa_produtos[cod_item] = {
                                'Descrição Item': safe_get(campos, 2),
                                'Tipo Item': safe_get(campos, 6),
                                'NCM': safe_get(campos, 7)
                            }
        except Exception as e:
            print(f"Erro ao mapear o arquivo {nome_arquivo}: {e}")
            continue

    mapas_gerais = {'produtos': mapa_produtos}

    print("\nIniciando Etapa 2: Processamento do Bloco C (C100, C170, C190)...")
    for filepath in lista_arquivos:
        nome_arquivo = os.path.basename(filepath)
        print(f"  - Processando arquivo: {nome_arquivo}")
        info_empresa = mapa_empresa.get(nome_arquivo, {})
        report_rows_arquivo = []

        try:
            with open(filepath, 'r', encoding='latin-1') as f:
                current_c100_data = {}
                temp_c170_list = []
                temp_c190_list = []
                cnpj_estabelecimento_atual = ""

                for line in f:
                    campos = parse_line(line)
                    if not campos: continue
                    reg = campos[0]

                    if reg == 'C010':
                        cnpj_estabelecimento_atual = safe_get(campos, 1)

                    if (reg == 'C100' or (current_c100_data and not reg.startswith('C1'))) and current_c100_data:
                        report_rows_arquivo.extend(
                            processar_bloco_nota_contrib(current_c100_data, temp_c170_list, temp_c190_list, mapas_gerais)
                        )
                        current_c100_data, temp_c170_list, temp_c190_list = {}, [], []

                    if reg == 'C100':
                        cod_part_interno = safe_get(campos, 3) # Índice 3
                        participante = mapa_participantes.get(cod_part_interno, {})

                        cnpj_cpf_final = participante.get('CNPJ/CPF Participante')
                        if not cnpj_cpf_final:
                            cnpj_cpf_final = cod_part_interno

                        current_c100_data = {
                            'CNPJ Estabelecimento': cnpj_estabelecimento_atual,
                            'Nome Empresa Declarante': info_empresa.get('Nome Empresa Declarante', ''),
                            'Período': info_empresa.get('Período', ''),
                            'Tipo Operação': 'Entrada' if safe_get(campos, 1) == '0' else 'Saída', # Índice 1
                            'CNPJ/CPF Participante': cnpj_cpf_final,
                            'Nome Participante': participante.get('Nome Participante', 'Participante não encontrado'),
                            'Número Documento': safe_get(campos, 7), # Índice 7
                            'Chave NF-e': safe_get(campos, 8),     # Índice 8
                            'Data Documento': format_date(safe_get(campos, 10)), # Índice 10
                            'Vlr Documento': safe_float_convert(safe_get(campos, 12)), # Índice 12
                            'Origem': '', 'Código Item': '', 'Descrição Item': '', 'NCM': '',
                            'Tipo Item': '', 'CFOP': '', 'Vlr Item': 0.0, 'Vlr Operação': 0.0
                        }

                    elif reg == 'C170' and current_c100_data:
                        temp_c170_list.append(campos)

                    elif reg == 'C190' and current_c100_data:
                        temp_c190_list.append(campos)

                if current_c100_data:
                    report_rows_arquivo.extend(
                        processar_bloco_nota_contrib(current_c100_data, temp_c170_list, temp_c190_list, mapas_gerais)
                    )

            all_report_rows.extend(report_rows_arquivo)

        except Exception as e:
            print(f"Erro ao processar o arquivo {nome_arquivo}: {e}")
            continue

    if not all_report_rows:
        print("Aviso: Nenhum registro C100/C170/C190 válido foi encontrado nos arquivos processados.")
        return pd.DataFrame()

    df_final = pd.DataFrame(all_report_rows)
    print(f"Processamento concluído. {len(df_final)} linhas geradas.")

    colunas_finais = [
        'CNPJ Estabelecimento', 'Nome Empresa Declarante', 'Período', 'Data Documento',
        'Tipo Operação', 'Número Documento', 'Chave NF-e',
        'CNPJ/CPF Participante', 'Nome Participante', 'Vlr Documento',
        'Origem', 'Código Item', 'Descrição Item', 'NCM', 'Tipo Item', 'CFOP',
        'Vlr Item', 'Vlr Operação'
    ]
    for col in colunas_finais:
        if col not in df_final.columns:
            default_value = 0.0 if 'Vlr' in col else ''
            df_final[col] = default_value

    return df_final[colunas_finais]


# ==============================================================================
# ETAPA 2: LÓGICA DE PROCESSAMENTO DO DASHBOARD
# ==============================================================================

def process_dataframe_for_dashboard(df_raw):
    """Prepara o DataFrame bruto (do CSV ou SPED) para o dashboard."""

    if df_raw is None:
        print("Erro crítico: process_dataframe_for_dashboard recebeu None. Retornando DF vazio.")
        return pd.DataFrame()
    if df_raw.empty:
        print("Aviso: DataFrame bruto está vazio. Nenhum dado para processar.")
        return pd.DataFrame()

    df = df_raw.copy()

    # --- (NOVO) Coluna 'CFOP' adicionada ---
    required_cols = ['Tipo Operação', 'CNPJ/CPF Participante', 'Nome Participante',
                     'NCM', 'CFOP', 'Data Documento', 'Vlr Item', 'Vlr Operação']
    
    for col in required_cols:
        if col not in df.columns:
            print(f"Aviso: Coluna '{col}' não encontrada no DataFrame. Adicionando coluna vazia/padrão.")
            default_value = 0.0 if 'Vlr' in col else ''
            df[col] = default_value

    df_entrada = df[df['Tipo Operação'] == 'Entrada'].copy()

    if df_entrada.empty:
        print("Aviso: Não foram encontradas operações de 'Entrada' no DataFrame.")
        # Retorna um DF vazio com as colunas esperadas
        return pd.DataFrame(columns=['CNPJ/CPF Participante', 'Nome Participante', 'Valor_Total', 'NCM', 'CFOP', 'Data Documento'])

    df_entrada['Vlr Item'] = pd.to_numeric(df_entrada['Vlr Item'], errors='coerce').fillna(0.0)
    df_entrada['Vlr Operação'] = pd.to_numeric(df_entrada['Vlr Operação'], errors='coerce').fillna(0.0)

    # Define o 'Valor_Total' com base na melhor informação disponível (Item > Operação)
    df_entrada['Valor_Total'] = df_entrada.apply(
        lambda row: row['Vlr Item'] if row['Vlr Item'] > 0 else row['Vlr Operação'],
        axis=1
    )

    try:
        # Converte 'Data Documento' para datetime
        df_entrada['Data Documento'] = pd.to_datetime(df_entrada['Data Documento'], format='%d/%m/%Y', errors='coerce')
        if df_entrada['Data Documento'].isnull().any(): # Se a primeira tentativa deixou NaTs
             df_entrada['Data Documento'] = pd.to_datetime(df_entrada['Data Documento'], errors='coerce') # Tenta inferir
    except Exception as e:
        print(f"Aviso: Falha ao converter 'Data Documento' para data: {e}. A coluna será mantida como está.")

    # --- Limpeza das colunas de agrupamento ---
    df_entrada['NCM'] = df_entrada['NCM'].fillna('Não Informado').astype(str)
    df_entrada['NCM'] = df_entrada['NCM'].str.strip()
    df_entrada.loc[df_entrada['NCM'] == '', 'NCM'] = 'Não Informado'

    # --- (NOVO) Limpeza da coluna CFOP ---
    df_entrada['CFOP'] = df_entrada['CFOP'].fillna('Não Informado').astype(str)
    df_entrada['CFOP'] = df_entrada['CFOP'].str.strip()
    df_entrada.loc[df_entrada['CFOP'] == '', 'CFOP'] = 'Não Informado'

    df_entrada['Nome Participante'] = df_entrada['Nome Participante'].fillna('Não Informado').astype(str)
    df_entrada.loc[df_entrada['Nome Participante'] == '', 'Nome Participante'] = 'Não Informado'

    df_entrada['CNPJ/CPF Participante'] = df_entrada['CNPJ/CPF Participante'].fillna('Não Informado').astype(str)
    df_entrada.loc[df_entrada['CNPJ/CPF Participante'] == '', 'CNPJ/CPF Participante'] = 'Não Informado'

    # --- (NOVO) Coluna 'CFOP' adicionada ao DF final ---
    df_final = df_entrada[[
        'CNPJ/CPF Participante',
        'Nome Participante',
        'Valor_Total',
        'NCM',
        'CFOP', # <--- ADICIONADO
        'Data Documento'
    ]].copy()

    print(f"process_dataframe_for_dashboard concluído. {len(df_final)} linhas de entrada válidas.")
    return df_final


def calculate_supplier_abc(df_entrada):
    """Calcula a Curva ABC de fornecedores com base no Valor_Total."""
    if df_entrada is None or df_entrada.empty or 'Valor_Total' not in df_entrada.columns:
        return pd.DataFrame(columns=['CNPJ/CPF Participante', 'Nome Participante', 'Valor_Total', 'Percentual Acumulado', 'Curva_ABC'])

    if 'CNPJ/CPF Participante' not in df_entrada.columns:
        print("Erro: Coluna 'CNPJ/CPF Participante' não encontrada para calcular ABC.")
        return pd.DataFrame(columns=['CNPJ/CPF Participante', 'Nome Participante', 'Valor_Total', 'Percentual Acumulado', 'Curva_ABC'])

    supplier_summary = df_entrada.groupby(['CNPJ/CPF Participante', 'Nome Participante'])['Valor_Total'].sum().reset_index()
    supplier_summary = supplier_summary.sort_values(by='Valor_Total', ascending=False)
    
    total_value = supplier_summary['Valor_Total'].sum()
    if total_value == 0:
        supplier_summary['Percentual'] = 0.0
    else:
        supplier_summary['Percentual'] = (supplier_summary['Valor_Total'] / total_value) * 100
    
    supplier_summary['Percentual Acumulado'] = supplier_summary['Percentual'].cumsum()
    supplier_summary['Curva_ABC'] = np.where(supplier_summary['Percentual Acumulado'] <= 80, 'A',
                                          np.where(supplier_summary['Percentual Acumulado'] <= 95, 'B', 'C'))
    return supplier_summary

def calculate_ncm_summary(df_entrada):
    """Calcula o resumo de gastos por NCM, agrupando os menores em 'Outros'."""
    if df_entrada is None or df_entrada.empty or 'Valor_Total' not in df_entrada.columns or 'NCM' not in df_entrada.columns:
        return pd.DataFrame(columns=['NCM', 'Valor_Total'])

    ncm_summary = df_entrada.groupby('NCM')['Valor_Total'].sum().reset_index()
    ncm_summary = ncm_summary.sort_values(by='Valor_Total', ascending=False)
    
    top_n = 15 # Define quantos NCMs principais mostrar
    if len(ncm_summary) > top_n:
        outros = ncm_summary.iloc[top_n:]['Valor_Total'].sum()
        ncm_summary = ncm_summary.iloc[:top_n].copy()
        if outros > 0:
             outros_row = pd.DataFrame([{'NCM': 'Outros', 'Valor_Total': outros}])
             ncm_summary = pd.concat([ncm_summary, outros_row], ignore_index=True)
    return ncm_summary

# --- (NOVA FUNÇÃO) ---
def calculate_cfop_summary(df_entrada):
    """Calcula o resumo de gastos por CFOP, agrupando os menores em 'Outros'."""
    if df_entrada is None or df_entrada.empty or 'Valor_Total' not in df_entrada.columns or 'CFOP' not in df_entrada.columns:
        return pd.DataFrame(columns=['CFOP', 'Valor_Total'])

    cfop_summary = df_entrada.groupby('CFOP')['Valor_Total'].sum().reset_index()
    cfop_summary = cfop_summary.sort_values(by='Valor_Total', ascending=False)
    
    top_n = 15 # Define quantos CFOPs principais mostrar
    if len(cfop_summary) > top_n:
        outros = cfop_summary.iloc[top_n:]['Valor_Total'].sum()
        cfop_summary = cfop_summary.iloc[:top_n].copy()
        if outros > 0:
             outros_row = pd.DataFrame([{'CFOP': 'Outros', 'Valor_Total': outros}])
             ncm_summary = pd.concat([cfop_summary, outros_row], ignore_index=True)
    return cfop_summary
# --- (FIM DA NOVA FUNÇÃO) ---

def calculate_monthly_summary(df_entrada):
    """Calcula o resumo de gastos mensais."""
    if df_entrada is None or df_entrada.empty or 'Valor_Total' not in df_entrada.columns or 'Data Documento' not in df_entrada.columns:
        return pd.DataFrame(columns=['MesAno', 'Valor_Total', 'MesAno_Str'])

    # Garante que 'Data Documento' seja datetime antes de prosseguir
    if not pd.api.types.is_datetime64_any_dtype(df_entrada['Data Documento']):
        df_entrada['Data Documento'] = pd.to_datetime(df_entrada['Data Documento'], errors='coerce')

    df_entrada = df_entrada.dropna(subset=['Data Documento'])

    if df_entrada.empty:
         return pd.DataFrame(columns=['MesAno', 'Valor_Total', 'MesAno_Str'])

    # Agora é seguro usar '.dt'
    df_entrada['MesAno'] = df_entrada['Data Documento'].dt.to_period('M')
    monthly_summary = df_entrada.groupby('MesAno')['Valor_Total'].sum().reset_index()
    monthly_summary = monthly_summary.sort_values(by='MesAno')
    
    # Converte o período para string no formato MM/YYYY
    monthly_summary['MesAno_Str'] = monthly_summary['MesAno'].dt.strftime('%m/%Y')
    
    return monthly_summary