# ===== ANALISADOR DE XML DE NF-e v2.1 (VERSÃO COMPLETA E CORRIGIDA) =====

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
# PARTE 1: LÓGICA DE PROCESSAMENTO DOS ARQUIVOS XML
# ==============================================================================

NFE_NAMESPACE = {'nfe': 'http://www.portalfiscal.inf.br/nfe'}

def safe_get(element, path, default=''):
    """Extrai texto de um elemento XML de forma segura."""
    if element is None: return default
    try:
        found = element.find(path, NFE_NAMESPACE)
        return found.text.strip() if found is not None and found.text is not None else default
    except AttributeError:
        return default

def safe_get_float(element, path, default=0.0):
    """Extrai e converte um valor para float de forma segura."""
    value_str = safe_get(element, path, '0')
    try:
        return float(value_str)
    except (ValueError, TypeError):
        return default

def get_specific_tax(imposto_element, tax_path_map):
    """
    Extrai o primeiro valor encontrado para cada campo fiscal, procurando em diferentes
    subgrupos (ex: PISAliq, PISNT).
    """
    tax_data = {}
    for key, paths in tax_path_map.items():
        tax_data[key] = ''
        for path in paths:
            value = safe_get(imposto_element, path)
            if value:
                tax_data[key] = value
                break # Pega o primeiro que encontrar
    return tax_data

def processar_xmls_nfe(lista_arquivos):
    """
    Função principal que lê múltiplos arquivos XML de NF-e, processa os dados,
    e retorna dois DataFrames: um com os itens e outro com os erros.
    """
    all_items_data = []
    error_log = []
    print(f"Iniciando processamento de {len(lista_arquivos)} arquivo(s)...")

    for filepath in lista_arquivos:
        nome_arquivo = os.path.basename(filepath)
        try:
            print(f"  - Processando arquivo: {nome_arquivo}")
            tree = ET.parse(filepath)
            root = tree.getroot()
            infNFe = root.find('.//nfe:infNFe', NFE_NAMESPACE)
            
            if infNFe is None:
                raise ValueError("Estrutura do XML inválida: tag <infNFe> não encontrada.")

            emit = infNFe.find('nfe:emit', NFE_NAMESPACE)
            dest = infNFe.find('nfe:dest', NFE_NAMESPACE)
            ide = infNFe.find('nfe:ide', NFE_NAMESPACE)
            total = infNFe.find('nfe:total/nfe:ICMSTot', NFE_NAMESPACE)
            transp = infNFe.find('nfe:transp', NFE_NAMESPACE)
            infAdic = infNFe.find('nfe:infAdic', NFE_NAMESPACE)

            header_data = {
                'CNPJ emitente': safe_get(emit, 'nfe:CNPJ'),
                'Nome do emitente': safe_get(emit, 'nfe:xNome'),
                'Inscrição Estadual emitente': safe_get(emit, 'nfe:IE'),
                'UF emitente': safe_get(emit, 'nfe:enderEmit/nfe:UF'),
                'Número do Documento (NF-e)': safe_get(ide, 'nfe:nNF'),
                'Natureza da Operação': safe_get(ide, 'nfe:natOp'),
                'Valor total da operação': safe_get_float(total, 'nfe:vNF'),
                'CNPJ destinatário': safe_get(dest, 'nfe:CNPJ') or safe_get(dest, 'nfe:CPF'),
                'Nome do destinatario': safe_get(dest, 'nfe:xNome'),
                'UF destinatário': safe_get(dest, 'nfe:enderDest/nfe:UF'),
                'Referência de NF de devolução': safe_get(ide, 'nfe:NFref/nfe:refNFe'),
                'Observação do Fisco (<infAdFisco>)': safe_get(infAdic, 'nfe:infAdFisco'),
            }
            try:
                dh_sai_ent_str = safe_get(ide, 'nfe:dhSaiEnt') or safe_get(ide, 'nfe:dhEmi')
                header_data['Data de Saída'] = datetime.fromisoformat(dh_sai_ent_str).strftime('%d/%m/%Y %H:%M:%S') if dh_sai_ent_str else ''
            except (ValueError, TypeError):
                header_data['Data de Saída'] = ''
            
            valor_total_produtos = safe_get_float(total, 'nfe:vProd')
            valor_frete_total = safe_get_float(total, 'nfe:vFrete')

            for item in infNFe.findall('nfe:det', NFE_NAMESPACE):
                prod = item.find('nfe:prod', NFE_NAMESPACE)
                imposto = item.find('nfe:imposto', NFE_NAMESPACE)
                valor_produto_item = safe_get_float(prod, 'nfe:vProd')
                
                frete_rateado = 0.0
                if valor_frete_total > 0 and valor_total_produtos > 0:
                    frete_rateado = (valor_produto_item / valor_total_produtos) * valor_frete_total
                
                icms_map = {'CST': ['.//nfe:ICMS/*/nfe:CST', './/nfe:ICMS/*/nfe:CSOSN'], 'vBC': ['.//nfe:ICMS/*/nfe:vBC'], 'pICMS': ['.//nfe:ICMS/*/nfe:pICMS'], 'vICMS': ['.//nfe:ICMS/*/nfe:vICMS']}
                pis_map = {'CST': ['.//nfe:PIS/*/nfe:CST'], 'vBC': ['.//nfe:PIS/*/nfe:vBC'], 'pPIS': ['.//nfe:PIS/*/nfe:pPIS'], 'vPIS': ['.//nfe:PIS/*/nfe:vPIS']}
                cofins_map = {'CST': ['.//nfe:COFINS/*/nfe:CST'], 'vBC': ['.//nfe:COFINS/*/nfe:vBC'], 'pCOFINS': ['.//nfe:COFINS/*/nfe:pCOFINS'], 'vCOFINS': ['.//nfe:COFINS/*/nfe:vCOFINS']}
                ipi_map = {'CST': ['.//nfe:IPI/*/nfe:CST'], 'vBC': ['.//nfe:IPI/*/nfe:vBC'], 'pIPI': ['.//nfe:IPI/*/nfe:pIPI'], 'vIPI': ['.//nfe:IPI/*/nfe:vIPI']}

                icms_data = get_specific_tax(imposto, icms_map)
                pis_data = get_specific_tax(imposto, pis_map)
                cofins_data = get_specific_tax(imposto, cofins_map)
                ipi_data = get_specific_tax(imposto, ipi_map)

                item_row = {
                    **header_data,
                    'CFOP': safe_get(prod, 'nfe:CFOP'),
                    'Código do Produto': safe_get(prod, 'nfe:cProd'),
                    'Nome do Produto': safe_get(prod, 'nfe:xProd'),
                    'NCM': safe_get(prod, 'nfe:NCM'),
                    'CST': icms_data.get('CST'),
                    'CST PIS/COFINS': pis_data.get('CST'),
                    'CST IPI': ipi_data.get('CST'),
                    'Quantidade': safe_get_float(prod, 'nfe:qCom'),
                    'Unidade': safe_get(prod, 'nfe:uCom'),
                    'Valor total do produto': valor_produto_item,
                    'Valor unitário': safe_get_float(prod, 'nfe:vUnCom'),
                    'PIS – base de cálculo': safe_get_float(pis_data, 'vBC'),
                    'PIS – alíquota': safe_get_float(pis_data, 'pPIS'),
                    'PIS – valor': safe_get_float(pis_data, 'vPIS'),
                    'COFINS – base de cálculo': safe_get_float(cofins_data, 'vBC'),
                    'COFINS – alíquota': safe_get_float(cofins_data, 'pCOFINS'),
                    'COFINS – valor': safe_get_float(cofins_data, 'vCOFINS'),
                    'ICMS – base de cálculo': safe_get_float(icms_data, 'vBC'),
                    'ICMS – alíquota': safe_get_float(icms_data, 'pICMS'),
                    'ICMS – valor': safe_get_float(icms_data, 'vICMS'),
                    'IPI – base de cálculo': safe_get_float(ipi_data, 'vBC'),
                    'IPI – alíquota': safe_get_float(ipi_data, 'pIPI'),
                    'IPI – valor': safe_get_float(ipi_data, 'vIPI'),
                    'Frete': frete_rateado if frete_rateado > 0 else '',
                    'Informações adicionais do item (infAdProd)': safe_get(item, 'nfe:infAdProd'),
                }
                all_items_data.append(item_row)
        
        except Exception as e:
            print(f"  -> ERRO ao processar {nome_arquivo}: {e}")
            error_log.append({'Arquivo': nome_arquivo, 'Motivo do Erro': str(e)})
            continue

    df_items = pd.DataFrame(all_items_data)
    df_errors = pd.DataFrame(error_log)
    
    final_columns_order = [
        'CNPJ emitente', 'Nome do emitente', 'Número do Documento (NF-e)', 'Inscrição Estadual emitente', 'UF emitente', 
        'Natureza da Operação', 'CFOP', 'Valor total da operação', 'Data de Saída', 'CNPJ destinatário', 'Nome do destinatario',
        'UF destinatário', 'Código do Produto', 'Nome do Produto', 'NCM', 'CST', 'CST PIS/COFINS', 'CST IPI', 'Quantidade', 'Unidade', 
        'Valor total do produto', 'Valor unitário', 'PIS – base de cálculo', 'PIS – alíquota', 'PIS – valor', 
        'COFINS – base de cálculo', 'COFINS – alíquota', 'COFINS – valor', 'ICMS – base de cálculo', 'ICMS – alíquota', 
        'ICMS – valor', 'IPI – base de cálculo', 'IPI – alíquota', 'IPI – valor', 'Frete', 
        'Informações adicionais do item (infAdProd)', 'Referência de NF de devolução', 'Observação do Fisco (<infAdFisco>)'
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

class XmlAnalyzerApp(QWidget):
    def __init__(self):
        super().__init__()
        self.selected_files = []
        self.initUI()

    def initUI(self):
        self.setWindowTitle("Analisador de XML de NF-e (Modelo 55)")
        self.setGeometry(200, 200, 700, 450)
        main_layout = QVBoxLayout(self)
        main_layout.setSpacing(15)
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.addWidget(QLabel("1. Selecione um ou mais arquivos XML de NF-e:"))
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
        filepaths, _ = QFileDialog.getOpenFileNames(self, "Selecione os arquivos XML", "", "Arquivos XML (*.xml)")
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
            df_items, df_errors = processar_xmls_nfe(self.selected_files)
            if df_items.empty and not df_errors.empty:
                self.status_label.setText(f"❌ Todos os {len(df_errors)} arquivos falharam.")
                QMessageBox.warning(self, "Processamento Concluído", "Todos os arquivos selecionados continham erros e não puderam ser processados.")
                self.save_report(df_items, df_errors)
            else:
                 self.save_report(df_items, df_errors)
                 self.status_label.setText("✅ Processamento concluído!")
        except Exception as e:
            self.status_label.setText(f"❌ Erro inesperado: {e}")
            QMessageBox.critical(self, "Erro Crítico", f"Ocorreu um erro inesperado durante o processamento.\n\nDetalhes: {e}")
        finally:
            self.run_button.setEnabled(True)

    def save_report(self, df_items, df_errors):
        default_filename = f"Relatorio_XML_NFe_{datetime.now().strftime('%Y%m%d')}.xlsx"
        save_path, _ = QFileDialog.getSaveFileName(self, "Salvar Relatório", default_filename, "Arquivos Excel (*.xlsx)")
        if save_path:
            try:
                self.status_label.setText("Salvando arquivo Excel...")
                QApplication.processEvents()
                with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
                    df_items.to_excel(writer, index=False, sheet_name="Itens_NFe")
                    if not df_errors.empty:
                        df_errors.to_excel(writer, index=False, sheet_name="Erros_Processamento")
                msg = f"Relatório salvo em:\n{save_path}"
                if not df_errors.empty:
                    msg += f"\n\nAtenção: {len(df_errors)} arquivo(s) não puderam ser processados. Verifique a aba 'Erros_Processamento'."
                QMessageBox.information(self, "Sucesso", msg)
            except Exception as e:
                self.status_label.setText(f"❌ Erro ao salvar o arquivo: {e}")
                QMessageBox.critical(self, "Erro ao Salvar", f"Não foi possível salvar o arquivo.\n\nDetalhes: {e}")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = XmlAnalyzerApp()
    window.show()
    sys.exit(app.exec())