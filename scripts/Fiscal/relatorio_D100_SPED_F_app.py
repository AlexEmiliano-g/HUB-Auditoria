# ===== ANALISADOR DE XML DE CT-e (Modelo 57) v1.2 - COMPLETO E COMENTADO =====

import sys
import os
import pandas as pd
import xml.etree.ElementTree as ET
from datetime import datetime
from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout,
    QPushButton, QLabel, QFileDialog, QMessageBox, QListWidget
)
from PyQt6.QtGui import QFont
from PyQt6.QtCore import Qt

# ==============================================================================
# PARTE 1: LÓGICA DE PROCESSAMENTO DOS ARQUIVOS XML DE CT-e
# ==============================================================================

# Define o namespace padrão usado nos arquivos XML do CT-e para facilitar as buscas.
CTE_NAMESPACE = {'cte': 'http://www.portalfiscal.inf.br/cte'}

def safe_get(element, path, default=''):
    """
    Função auxiliar para extrair texto de um elemento XML de forma segura.
    Evita que o programa quebre se uma tag esperada não for encontrada.
    """
    if element is None: return default
    try:
        # Procura pelo caminho especificado dentro do elemento.
        found = element.find(path, CTE_NAMESPACE)
        # Retorna o texto do elemento se ele for encontrado e não for vazio.
        return found.text.strip() if found is not None and found.text is not None else default
    except AttributeError:
        # Retorna o valor padrão em caso de qualquer erro.
        return default

def safe_get_float(value_str):
    """
    Função auxiliar para converter uma string para número (float) de forma segura.
    Trata valores vazios e erros de conversão, retornando 0.0.
    """
    if not value_str: return 0.0
    try:
        # A conversão para float no Python usa ponto como separador decimal.
        return float(value_str)
    except (ValueError, TypeError):
        return 0.0

def get_specific_tax(imposto_element, tax_path_map):
    """
    Função robusta para extrair o primeiro valor encontrado para cada campo fiscal.
    Ela procura em diferentes subgrupos possíveis (ex: ICMS00, ICMS90, PISAliq, etc.).
    O '*' no caminho XPath atua como um "coringa", permitindo a busca dentro de qualquer tag filha.
    """
    tax_data = {}
    for key, paths in tax_path_map.items():
        tax_data[key] = '' # Inicializa a chave com valor em branco.
        if imposto_element is None: continue
        
        for path in paths:
            # O './/' garante a busca em qualquer nível abaixo do elemento de imposto.
            value = safe_get(imposto_element, path)
            if value:
                tax_data[key] = value
                break # Pega o primeiro valor não vazio que encontrar e para de procurar.
    return tax_data

def format_date_flexible(date_str, output_format='%d/%m/%Y'):
    """Formata uma data de múltiplos padrões (ISO ou AAAA-MM-DD) para o formato brasileiro."""
    if not date_str: return ''
    try:
        if 'T' in date_str:
             date_str_clean = date_str[:19]
             return datetime.fromisoformat(date_str_clean).strftime('%d/%m/%Y %H:%M:%S')
        else:
             return datetime.strptime(date_str, '%Y-%m-%d').strftime(output_format)
    except (ValueError, TypeError):
        return date_str

def processar_xmls_cte(lista_arquivos):
    """
    Função principal que lê múltiplos arquivos XML de CT-e, processa os dados
    e retorna um DataFrame consolidado com uma linha por NF-e vinculada.
    """
    all_rows_data = []
    error_log = []
    print(f"Iniciando processamento de {len(lista_arquivos)} arquivo(s) de CT-e...")

    for filepath in lista_arquivos:
        nome_arquivo = os.path.basename(filepath)
        try:
            print(f"  - Processando arquivo: {nome_arquivo}")
            tree = ET.parse(filepath)
            root = tree.getroot()
            
            infCte = root.find('.//cte:infCte', CTE_NAMESPACE)
            if infCte is None:
                raise ValueError("Estrutura do XML inválida: tag <infCte> não encontrada.")

            # Mapeia os principais grupos de tags do XML
            ide = infCte.find('cte:ide', CTE_NAMESPACE)
            emit = infCte.find('cte:emit', CTE_NAMESPACE)
            rem = infCte.find('cte:rem', CTE_NAMESPACE)
            dest = infCte.find('cte:dest', CTE_NAMESPACE)
            vPrest = infCte.find('cte:vPrest', CTE_NAMESPACE)
            imp = infCte.find('cte:imp', CTE_NAMESPACE)
            infCarga = infCte.find('.//cte:infCarga', CTE_NAMESPACE)
            
            chave_cte = infCte.get('Id', '').replace('CTe', '')

            # Coleta os dados do cabeçalho do CT-e
            header_data = {
                'CNPJ Emitente': safe_get(emit, 'cte:CNPJ'),
                'Nome Emitente': safe_get(emit, 'cte:xNome'),
                'IE Emitente': safe_get(emit, 'cte:IE'),
                'UF Emitente': safe_get(ide, 'cte:UFEnv'),
                'Número CT-e': safe_get(ide, 'cte:nCT'),
                'Série CT-e': safe_get(ide, 'cte:serie'),
                'Chave CT-e': chave_cte,
                'Data Emissão': format_date_flexible(safe_get(ide, 'cte:dhEmi')),
                'Natureza da Operação': safe_get(ide, 'cte:natOp'),
                'CFOP': safe_get(ide, 'cte:CFOP'),
                'UF Início': safe_get(ide, 'cte:UFIni'),
                'UF Fim': safe_get(ide, 'cte:UFFim'),
                'CNPJ Remetente': safe_get(rem, 'cte:CNPJ'),
                'Nome Remetente': safe_get(rem, 'cte:xNome'),
                'CNPJ Destinatário': safe_get(dest, 'cte:CNPJ') if dest is not None else '',
                'Nome Destinatário': safe_get(dest, 'cte:xNome') if dest is not None else '',
                'Valor Total Prestação': safe_get_float(safe_get(vPrest, 'cte:vTPrest')),
                'Valor Carga': safe_get_float(safe_get(infCarga, 'cte:vCarga')),
                'Produto Predominante': safe_get(infCarga, 'cte:proPred'),
                'Quantidade Carga': safe_get_float(safe_get(infCarga, 'cte:infQ/cte:qCarga')),
                'Unidade Carga': safe_get(infCarga, 'cte:infQ/cte:tpMed'),
            }
            
            # Define os mapas de busca para cada imposto
            icms_map = {'CST': ['.//cte:ICMS/*/cte:CST', './/cte:ICMS/*/cte:CSOSN'], 'vBC': ['.//cte:ICMS/*/cte:vBC'], 'pICMS': ['.//cte:ICMS/*/cte:pICMS'], 'vICMS': ['.//cte:ICMS/*/cte:vICMS']}
            pis_map = {'CST': ['.//cte:PIS/*/cte:CST'], 'vBC': ['.//cte:PIS/*/cte:vBC'], 'pPIS': ['.//cte:PIS/*/cte:pPIS'], 'vPIS': ['.//cte:PIS/*/cte:vPIS']}
            cofins_map = {'CST': ['.//cte:COFINS/*/cte:CST'], 'vBC': ['.//cte:COFINS/*/cte:vBC'], 'pCOFINS': ['.//cte:COFINS/*/cte:pCOFINS'], 'vCOFINS': ['.//cte:COFINS/*/cte:vCOFINS']}

            # Chama a função de extração para cada imposto
            icms_data = get_specific_tax(imp, icms_map)
            pis_data = get_specific_tax(imp, pis_map)
            cofins_data = get_specific_tax(imp, cofins_map)
            
            # Adiciona os dados de impostos ao cabeçalho
            header_data.update({
                'CST ICMS': icms_data.get('CST'),
                'ICMS - base de cálculo': safe_get_float(icms_data.get('vBC')),
                'ICMS - alíquota': safe_get_float(icms_data.get('pICMS')),
                'ICMS - valor': safe_get_float(icms_data.get('vICMS')),
                'CST PIS': pis_data.get('CST'),
                'PIS - base de cálculo': safe_get_float(pis_data.get('vBC')),
                'PIS - alíquota': safe_get_float(pis_data.get('pPIS')),
                'PIS - valor': safe_get_float(pis_data.get('vPIS')),
                'CST COFINS': cofins_data.get('CST'),
                'COFINS - base de cálculo': safe_get_float(cofins_data.get('vBC')),
                'COFINS - alíquota': safe_get_float(cofins_data.get('pCOFINS')),
                'COFINS - valor': safe_get_float(cofins_data.get('vCOFINS')),
            })
            
            infDoc = infCte.find('.//cte:infDoc', CTE_NAMESPACE)
            notas_vinculadas = infDoc.findall('cte:infNFe', CTE_NAMESPACE) if infDoc is not None else []
            
            if not notas_vinculadas:
                row = {**header_data, 'Chave NF-e Vinculada': 'N/A', 'Data Prevista Entrega': ''}
                all_rows_data.append(row)
            else:
                for nota in notas_vinculadas:
                    row = {
                        **header_data,
                        'Chave NF-e Vinculada': safe_get(nota, 'cte:chave'),
                        'Data Prevista Entrega': format_date_flexible(safe_get(nota, 'cte:dPrev'))
                    }
                    all_rows_data.append(row)

        except Exception as e:
            print(f"  -> ERRO ao processar {nome_arquivo}: {e}")
            error_log.append({'Arquivo': nome_arquivo, 'Motivo do Erro': str(e)})
            continue

    df_items = pd.DataFrame(all_rows_data)
    df_errors = pd.DataFrame(error_log)
    
    final_columns_order = [
        'CNPJ Emitente', 'Nome Emitente', 'IE Emitente', 'UF Emitente', 'Número CT-e', 'Série CT-e', 
        'Chave CT-e', 'Data Emissão', 'Natureza da Operação', 'CFOP', 'UF Início', 'UF Fim', 
        'CNPJ Remetente', 'Nome Remetente', 'CNPJ Destinatário', 'Nome Destinatário', 
        'Valor Total Prestação', 'CST ICMS', 'ICMS - base de cálculo', 'ICMS - alíquota', 'ICMS - valor',
        'CST PIS', 'PIS - base de cálculo', 'PIS - alíquota', 'PIS - valor',
        'CST COFINS', 'COFINS - base de cálculo', 'COFINS - alíquota', 'COFINS - valor',
        'Valor Carga', 'Produto Predominante', 'Quantidade Carga', 'Unidade Carga', 
        'Chave NF-e Vinculada', 'Data Prevista Entrega'
    ]
    if not df_items.empty:
        for col in final_columns_order:
            if col not in df_items.columns:
                df_items[col] = ''
        df_items = df_items[final_columns_order]
            
    return df_items, df_errors

# ==============================================================================
# PARTE 2: INTERFACE GRÁFICA (PyQt6)
# ==============================================================================

class CteAnalyzerApp(QWidget):
    def __init__(self):
        super().__init__()
        self.selected_files = []
        self.initUI()

    def initUI(self):
        self.setWindowTitle("Analisador de XML de CT-e (Modelo 57)")
        self.setGeometry(200, 200, 700, 450)
        
        main_layout = QVBoxLayout(self)
        main_layout.setSpacing(15)
        main_layout.setContentsMargins(20, 20, 20, 20)
        
        main_layout.addWidget(QLabel("1. Selecione um ou mais arquivos XML de CT-e:"))
        self.file_list_widget = QListWidget()
        main_layout.addWidget(self.file_list_widget)
        
        browse_button = QPushButton("Procurar Arquivos...")
        browse_button.clicked.connect(self.select_files)
        main_layout.addWidget(browse_button)
        
        self.run_button = QPushButton("🚀 Gerar Relatório Consolidado")
        self.run_button.setFont(QFont("Segoe UI", 12, QFont.Weight.Bold))
        self.run_button.clicked.connect(self.start_processing)
        main_layout.addWidget(self.run_button)
        
        self.status_label = QLabel("Pronto para começar.")
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        main_layout.addWidget(self.status_label)
        
        main_layout.addStretch()

    def select_files(self):
        filepaths, _ = QFileDialog.getOpenFileNames(self, "Selecione os arquivos XML de CT-e", "", "Arquivos XML (*.xml)")
        if filepaths:
            self.selected_files = filepaths
            self.file_list_widget.clear()
            for file_path in self.selected_files:
                self.file_list_widget.addItem(os.path.basename(file_path))

    def start_processing(self):
        if not self.selected_files:
            QMessageBox.warning(self, "Atenção", "Nenhum arquivo XML foi selecionado.")
            return
            
        self.run_button.setEnabled(False)
        self.status_label.setText("Processando múltiplos arquivos... Isso pode levar um momento.")
        QApplication.processEvents()

        try:
            df_items, df_errors = processar_xmls_cte(self.selected_files)
            
            if df_items.empty and not df_errors.empty:
                self.status_label.setText(f"❌ Todos os {len(df_errors)} arquivos falharam.")
                QMessageBox.warning(self, "Processamento Concluído", "Todos os arquivos selecionados continham erros e não puderam ser processados.")
            else:
                 self.status_label.setText("✅ Processamento concluído!")

            self.save_report(df_items, df_errors)
                 
        except Exception as e:
            self.status_label.setText(f"❌ Erro inesperado: {e}")
            QMessageBox.critical(self, "Erro Crítico", f"Ocorreu um erro inesperado durante o processamento.\n\nDetalhes: {e}")
        finally:
            self.run_button.setEnabled(True)

    def save_report(self, df_items, df_errors):
        default_filename = f"Relatorio_XML_CTe_{datetime.now().strftime('%Y%m%d')}.xlsx"
        save_path, _ = QFileDialog.getSaveFileName(self, "Salvar Relatório", default_filename, "Arquivos Excel (*.xlsx)")
        
        if save_path:
            try:
                self.status_label.setText("Salvando arquivo Excel...")
                QApplication.processEvents()
                
                with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
                    df_items.to_excel(writer, index=False, sheet_name="CTe_Analitico")
                    if not df_errors.empty:
                        df_errors.to_excel(writer, index=False, sheet_name="Erros_Processamento")
                
                msg = f"Relatório salvo em:\n{save_path}"
                if not df_errors.empty:
                    msg += f"\n\nAtenção: {len(df_errors)} arquivo(s) não puderam ser processados. Verifique a aba 'Erros_Processamento'."
                
                QMessageBox.information(self, "Sucesso", msg)
                
            except Exception as e:
                self.status_label.setText(f"❌ Erro ao salvar o arquivo: {e}")
                QMessageBox.critical(self, "Erro ao Salvar", f"Não foi possível salvar o arquivo.\n\nDetalhes: {e}")

# ==============================================================================
# PARTE 3: PONTO DE ENTRADA (MODIFICADO)
# ==============================================================================

def main():
    """Função principal que inicia esta aplicação específica."""
    # Reutiliza a instância da aplicação principal (do Hub) se ela existir.
    # Caso contrário, cria uma nova para execução independente.
    app = QApplication.instance() or QApplication(sys.argv)
    
    # Cria e exibe a janela da aplicação.
    window = CteAnalyzerApp()
    window.show()

    # Retorna a aplicação e a janela para o chamador poder gerenciá-las.
    return app, window

# Este bloco permite que o script ainda seja executável de forma independente para testes
if __name__ == '__main__':
    app, window = main()
    sys.exit(app.exec())