# Documentação de Arquitetura - HUB Auditoria

## 📋 Visão Geral

O **HUB Auditoria** é uma aplicação desktop modular desenvolvida em Python que centraliza ferramentas de automação para as áreas de **Auditoria e Contabilidade**. A aplicação segue um padrão de arquitetura baseado em um **hub central launcher** que gerencia múltiplos módulos temáticos organizados por setor.

### Informações Técnicas
- **Linguagem**: Python 3.x
- **Framework GUI**: PyQt6
- **Propósito**: Automação de processos contábeis e fiscais
- **Tipo**: Aplicação Desktop (Standalone)
- **Sistema Operacional**: Windows

---

## 🏗️ Arquitetura Geral

```
┌─────────────────────────────────────────────────────────────┐
│                      HUB Central (hub.py)                   │
│  - Splash Screen (Loading)                                  │
│  - Menu Lateral com Setores                                 │
│  - Lista de Programas por Setor                             │
│  - Log de Execução                                          │
└────────────────────┬────────────────────────────────────────┘
                     │
         ┌───────────┼───────────┬──────────────┐
         │           │           │              │
    ┌────▼────┐  ┌───▼───┐  ┌───▼────┐  ┌─────▼─────┐
    │Contabil │  │Fiscal │  │Análise │  │ IA (Beta) │
    │   idade │  │       │  │        │  │           │
    └────┬────┘  └───┬───┘  └───┬────┘  └─────┬─────┘
         │           │          │              │
    ┌────▼────┐  ┌───▼───────────┴──────────────────────┐
    │Tabulador│  │  Relatórios Fiscais (8 Tipos)       │
    │         │  │  - C100 (EFD-C, SPED-F)             │
    │         │  │  - C500 (EFD-C, SPED-F)             │
    │         │  │  - CIAP (SPED-F)                    │
    │         │  │  - CT (Controladoria)               │
    │         │  │  - D100 (SPED-F)                    │
    │         │  │  - F100 (EFD-C)                     │
    │         │  │  - XML NF-e                         │
    │         │  └────┬────────────────────────────────┘
    └──────────┘      │
                  ┌───▼───────┐
                  │ Análise    │
                  │ Fiscal     │
                  └────────────┘
```

---

## 📁 Estrutura de Diretórios

```
HUB-Auditoria/
├── hub.py                          # Aplicação Principal (Hub Launcher)
├── requirements.txt                # Dependências do Projeto
├── README.md                        # Documentação Inicial
├── ARQUITETURA.md                  # Este arquivo
├── DEPENDENCIAS.md                 # Documento de Dependências (gerado)
├── icone.ico                        # Ícone da Aplicação
├── logo.png                         # Logo Principal do HUB
├── logo_aber.png                    # Logo de Abertura (Splash)
│
└── scripts/                         # Diretório Principal de Módulos
    ├── Contabilidade/
    │   ├── tabulador_app.py         # Interface GUI do Tabulador
    │   ├── tabulador_logic.py       # Lógica de Processamento
    │   └── __pycache__/
    │
    ├── Fiscal/
    │   ├── relatorio_C100_EFD_C_app.py
    │   ├── relatorio_C100_SPED_F_app.py
    │   ├── relatorio_C500_EFD_C_app.py
    │   ├── relatorio_C500_SPED_F_app.py
    │   ├── relatorio_CIAP_SPED_F_app.py
    │   ├── relatorio_CT_app.py
    │   ├── relatorio_D100_SPED_F_app.py
    │   ├── relatorio_F100_EFD_C_app.py
    │   ├── relatorio_XML_app.py
    │   └── __pycache__/
    │
    ├── Analise/
    │   ├── analise_fiscal_app.py    # Interface GUI de Análise
    │   ├── analise_fiscal_logic.py  # Lógica de Análise
    │   └── __pycache__/
    │
    └── IA/
        ├── Texto_anonimo_app.py     # Módulo de Anonimização com IA (Beta)
        └── __pycache__/
```

---

## 🔄 Fluxo de Execução

### 1. Inicialização da Aplicação

```
┌──────────────────────────────────┐
│ Execução: python hub.py          │
└────────────┬─────────────────────┘
             │
             ▼
┌──────────────────────────────────┐
│ Splash Screen (Fade In)          │
│ Status: "Inicializando..."       │
└────────────┬─────────────────────┘
             │
             ▼
┌──────────────────────────────────┐
│ start_application(app, splash)   │
│ - Importa pandas                 │
│ - Importa openpyxl               │
│ - Inicializa Hub()               │
└────────────┬─────────────────────┘
             │
             ▼
┌──────────────────────────────────┐
│ Hub Window (Fade Out Splash)     │
│ - Menu Lateral (Setores)         │
│ - Área de Programas              │
│ - Log de Execução                │
└──────────────────────────────────┘
```

### 2. Seleção e Execução de Módulo

```
Usuário Seleciona Setor
    │
    ├─▶ Hub.load_scripts(setor)
    │   └─▶ Escaneia pasta: scripts/{setor}/
    │       └─▶ Identifica arquivos *_app.py
    │           └─▶ Popula lista de programas
    │
Usuário Seleciona Programa
    │
    └─▶ Hub.run_script()
        ├─▶ Importa módulo dinamicamente
        │   module_path = f"scripts.{setor}.{module_name}"
        │
        ├─▶ Executa function main()
        │   └─▶ Retorna (app, window)
        │
        └─▶ Exibe janela do programa
            └─▶ Log de Status
```

---

## 🧩 Componentes Principais

### 1. **Hub Central** (`hub.py`)

**Responsabilidades:**
- Gerenciamento da interface principal
- Carregamento dinâmico de módulos
- Controle do ciclo de vida da aplicação
- Renderização de Splash Screen

**Classes:**
- `Splash`: Janela de boas-vindas com animação de fade
- `Hub`: Janela principal com menu de setores e programas

**Métodos Principais:**
```python
Hub.load_scripts(item)      # Carrega scripts de um setor
Hub.run_script()            # Executa script selecionado
start_application(app, splash)  # Inicializa aplicação com loading
```

**Padrão Utilizado:** Lazy Loading + Dynamic Module Import

---

### 2. **Setor: Contabilidade**

#### Tabulador de Balancetes (`tabulador_app.py` + `tabulador_logic.py`)

**Objetivo:** Processar e tabular múltiplos arquivos de balancete de diferentes sistemas.

**Funcionalidades:**
- Seleção de múltiplos arquivos Excel
- Suporte a diferentes sistemas: Fuga, Uniair
- Processamento em batch
- Geração de relatórios consolidados

**Tecnologias:**
- pandas: Processamento de dados
- openpyxl: Manipulação de Excel
- PyQt6: Interface gráfica

**Fluxo:**
```
Selecionar Arquivos (CSV/XLSX)
    │
    ▼
Escolher Sistema (Fuga/Uniair)
    │
    ▼
processar_arquivos_selecionados()
    ├─ Lê cada arquivo
    ├─ Padroniza dados
    ├─ Consolida tabelas
    │
    ▼
Exportar Resultado (XLSX)
```

---

### 3. **Setor: Fiscal**

#### Módulos de Relatórios Fiscais (8 Tipos)

**Objetivo:** Gerar e analisar relatórios fiscais de diferentes formatos EFD-C e SPED-F.

| Módulo | Tipo | Formato | Descrição |
|--------|------|---------|-----------|
| `relatorio_C100_EFD_C_app.py` | C100 | EFD-C | Registro de Operações com ICMS |
| `relatorio_C100_SPED_F_app.py` | C100 | SPED-F | Registro de Operações com ICMS (Alimentar) |
| `relatorio_C500_EFD_C_app.py` | C500 | EFD-C | Registro de Consolidação |
| `relatorio_C500_SPED_F_app.py` | C500 | SPED-F | Registro de Consolidação (Alimentar) |
| `relatorio_CIAP_SPED_F_app.py` | CIAP | SPED-F | Controle de ICMS a Apropriação |
| `relatorio_CT_app.py` | CT | Controladoria | Relatórios de Controladoria |
| `relatorio_D100_SPED_F_app.py` | D100 | SPED-F | Documentos Fiscais (NFe, CTE) |
| `relatorio_F100_EFD_C_app.py` | F100 | EFD-C | Operações Internas |
| `relatorio_XML_app.py` | XML | NF-e | Análise de Nota Fiscal Eletrônica |

**Padrão de Implementação:**

Cada módulo de relatório segue o padrão:
```python
# relatorio_X_app.py
class RelatorioXApp(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()
    
    def initUI(self):
        # Constrói interface PyQt6
        pass
    
    def processar_dados(self):
        # Lógica de processamento
        pass

def main():
    app = QApplication(sys.argv)
    window = RelatorioXApp()
    window.show()
    return (app, window)
```

**XML Handler - Análise NF-e (`relatorio_XML_app.py`)**

Funcionalidades especializadas:
- Parse de arquivos XML de Nota Fiscal Eletrônica
- Extração segura de dados (tratamento de namespaces)
- Cálculo de impostos (ICMS, IPI, PIS, COFINS)
- Geração de relatórios consolidados em Excel

**Funções Utilitárias:**
```python
safe_get(element, path, default='')        # Extração segura de dados XML
safe_get_float(value_str)                  # Conversão segura para float
get_specific_tax(imposto_element, tax_map) # Cálculo de impostos específicos
```

---

### 4. **Setor: Análise**

#### Análise Fiscal (`analise_fiscal_app.py` + `analise_fiscal_logic.py`)

**Objetivo:** Análise comparativa e visualização de dados fiscais com gráficos interativos.

**Funcionalidades:**
- Carregamento de múltiplos arquivos de dados
- Análise comparativa entre períodos/entidades
- Visualização em gráficos (Matplotlib + PyQt6)
- Exportação de análises em PDF
- Interface com árvore de dados (TreeWidget)

**Arquitetura:**
- `analise_fiscal_app.py`: Camada de apresentação (PyQt6)
- `analise_fiscal_logic.py`: Camada de lógica (cálculos, processamento)

**Componentes UI:**
```
├─ QTabWidget (Abas)
│  ├─ Tab 1: Dados Brutos (QTreeWidget)
│  ├─ Tab 2: Gráficos (Matplotlib Canvas)
│  └─ Tab 3: Resumo Executivo
│
├─ Painel de Controle
│  ├─ Seletor de Arquivo
│  ├─ Filtros de Data
│  ├─ Menu de Opções (⋮)
│  └─ Botão de Exportação PDF
│
└─ Visualizador de Resultados
   └─ Log de Execução
```

**Integração Matplotlib:**
```python
matplotlib.use('Qt5Agg')  # Backend PyQt6
FigureCanvas          # Widget para renderizar gráficos
PdfPages              # Exportação para PDF
```

---

### 5. **Setor: IA** (Beta)

#### Anonimização de Texto (`Texto_anonimo_app.py`)

**Objetivo:** Processamento de textos com capacidades de anonimização baseadas em IA.

**Status:** Beta (Desenvolvimento)

**Tecnologias:** A definir (NLP, Transformers, ou serviços IA)

---

## 🔌 Dependências do Projeto

### Dependências Python

| Biblioteca | Versão | Propósito | Setor |
|-----------|--------|----------|-------|
| `pandas` | Latest | Processamento e manipulação de dados | Contabilidade, Fiscal, Análise |
| `PyQt6` | Latest | Framework GUI | Todos |
| `openpyxl` | Latest | Leitura/escrita de arquivos Excel | Contabilidade, Fiscal, Análise |
| `numpy` | Latest | Operações numéricas | Análise |
| `matplotlib` | Latest | Geração de gráficos | Análise |

**Arquivo:** `requirements.txt`
```
pandas
PyQt6
openpyxl
```

### Dependências Externas

- **Python 3.7+**: Runtime
- **Windows**: Sistema operacional alvo
- **Excel 2010+**: Para manipulação de arquivos (opcional)

### Namespace XML

- **NFe Namespace**: `http://www.portalfiscal.inf.br/nfe`
- Utilizado em: `relatorio_XML_app.py`

---

## 🎯 Padrões de Design Implementados

### 1. **Lazy Loading**
- Módulos são importados apenas quando necessário
- Reduz tempo de inicialização

```python
import importlib
script_module = importlib.import_module(module_path)
```

### 2. **Factory Pattern**
- Cada módulo expõe uma função `main()` que retorna a aplicação e janela

```python
def main():
    app = QApplication(sys.argv)
    window = RelatorioApp()
    window.show()
    return (app, window)
```

### 3. **Separação de Responsabilidades**
- Padrão MVC Light:
  - `*_app.py`: View (Interface GUI com PyQt6)
  - `*_logic.py`: Model/Logic (Processamento de dados)

### 4. **Resource Loading**
- Tratamento dinâmico de recursos (ícones, logos)

```python
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS  # PyInstaller
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)
```

---

## 🔐 Tratamento de Erros

### Estratégias Implementadas

1. **Try-Except com Contexto**
```python
try:
    script_module = importlib.import_module(module_path)
except ImportError as e:
    log.append(f"ERRO de Importação: {e}")
except AttributeError:
    log.append(f"ERRO: Script não possui função 'main()'")
except Exception as e:
    log.append(f"ERRO inesperado: {e}")
```

2. **Extração Segura de XML**
```python
def safe_get(element, path, default=''):
    try:
        found = element.find(path, NFE_NAMESPACE)
        return found.text.strip() if found is not None else default
    except AttributeError:
        return default
```

3. **Conversão Segura de Tipos**
```python
def safe_get_float(value_str):
    if not value_str: return 0.0
    try:
        return float(value_str)
    except (ValueError, TypeError):
        return 0.0
```

---

## 🎨 Estilização e Temas

### Stylesheet CSS (PyQt6)

```css
QWidget#HubWindow {
    background-color: #F8F9FA;
}

QWidget#LeftPanel {
    background-color: #FFFFFF;
    border-right: 1px solid #E0E0E0;
    border-radius: 10px;
}

QListWidget::item:selected {
    background-color: #E8F0FE;
    color: #1967D2;
    font-weight: bold;
}

QPushButton {
    background-color: #4285F4;
    color: white;
    border-radius: 5px;
}

QPushButton:hover {
    background-color: #1A73E8;
}
```

**Paleta de Cores:**
- Primária: #4285F4 (Azul Google)
- Secundária: #1967D2 (Azul Escuro)
- Fundo: #F8F9FA (Cinza Claro)
- Texto: #3C4043 (Cinza Escuro)
- Destaque: #E8F0FE (Azul Claro)

---

## 📊 Fluxo de Dados

### Contabilidade (Tabulador)
```
Arquivos Excel/CSV (Entrada)
        │
        ▼
sistema = [Fuga | Uniair]
        │
        ▼
pandas.read_excel() ou read_csv()
        │
        ▼
Normalização de Colunas
        │
        ▼
Merge/Consolidação
        │
        ▼
openpyxl.Workbook()
        │
        ▼
Arquivo Excel Consolidado (Saída)
```

### Fiscal (Relatórios)
```
Arquivo de Dados (EFD-C/SPED-F/XML)
        │
        ▼
Parser Específico (XML/CSV)
        │
        ▼
Extração de Informações
        │
        ▼
Cálculos Fiscais (Impostos)
        │
        ▼
DataFrame pandas
        │
        ▼
Formatação/Filtragem
        │
        ▼
Arquivo Excel/PDF (Saída)
```

### Análise
```
Dados Fiscais (Arquivo)
        │
        ▼
Carregamento em DataFrame
        │
        ▼
Análise Comparativa
        │
        ▼
Matplotlib Visualization
        │
        ▼
PdfPages Export
```

---

## 🚀 Pontos de Extensão

### Como Adicionar um Novo Módulo

1. **Criar pasta do setor** (se não existir):
   ```
   scripts/{NOVO_SETOR}/
   ```

2. **Implementar módulo** com padrão obrigatório:
   ```python
   # novo_modulo_app.py
   from PyQt6.QtWidgets import QApplication, QWidget
   
   class NovoModuloApp(QWidget):
       def __init__(self):
           super().__init__()
           self.initUI()
       
       def initUI(self):
           # Implementar interface
           pass
   
   def main():
       app = QApplication(sys.argv)
       window = NovoModuloApp()
       window.show()
       return (app, window)
   ```

3. **Registrar setor no Hub** (se novo):
   ```python
   # hub.py - linha 112
   for setor in ["Contabilidade", "Fiscal", "Analise", "NOVO_SETOR"]:
       self.sectors_list.addItem(QListWidgetItem(setor))
   ```

4. **Verificar carregamento dinâmico:**
   - Arquivo deve terminar em `_app.py`
   - Função `main()` obrigatória
   - Retornar tupla `(app, window)`

---

## 🔍 Debugging e Logs

### Sistema de Logging

- **Local:** QTextEdit no painel direito do Hub
- **Formato:** Console simples com histórico
- **Eventos Registrados:**
  - Inicialização de módulos
  - Erros de importação
  - Exceções não capturadas
  - Status de execução

```python
self.log.append(f">>> Carregando: {module_path}")
self.log.append(f"ERRO: {mensagem_erro}")
```

### Rastreamento de Erros

```python
import traceback
try:
    # código
except Exception as e:
    traceback.print_exc()  # Imprime stack trace completo
```

---

## 📈 Performance e Otimizações

### Implementadas

1. **Lazy Loading de Módulos**
   - Pandas e openpyxl importados apenas no `start_application()`
   - Reduz tempo de inicialização do Hub

2. **Cache de Recursos**
   - Logos e ícones carregados uma única vez
   - Armazenados em memória durante execução

3. **Processamento em Batch**
   - Tabulador processa múltiplos arquivos sequencialmente
   - Consolidação final em um arquivo

### Recomendações Futuras

1. Threading para processamento não-bloqueante
2. Compressão de dados em arquivos grandes
3. Cache de resultados intermediários
4. Índices em análises comparativas

---

## 📝 Convenções de Código

### Nomenclatura

- **Módulos:** `{descricao}_{tipo}.py`
  - `tabulador_app.py` (Interface)
  - `tabulador_logic.py` (Lógica)
  - `relatorio_XML_app.py` (App + Logic combinados)

- **Classes:** PascalCase + "App" sufixo
  - `TabuladorPythonApp`
  - `AnaliseFiscalApp`

- **Funções:** snake_case
  - `processar_arquivos_selecionados()`
  - `safe_get_float()`

- **Constantes:** UPPER_SNAKE_CASE
  - `LISTA_SISTEMAS`
  - `NFE_NAMESPACE`
  - `LOGO_HUB_PATH`

### Documentação

- Docstrings para funções complexas
- Comentários explicativos em algoritmos críticos
- Cabeçalhos de seção em arquivos grandes

---

## 🧪 Testes e Validação

### Estratégia de Teste Atual

- Testes manuais via interface GUI
- Validação de carregamento de módulos
- Testes de processamento de dados em pequenos datasets

### Recomendações

1. Implementar testes unitários (pytest)
2. Testes de integração entre módulos
3. Testes de regressão em processamento de dados
4. Testes de carga com múltiplos arquivos

---

## 📞 Suporte e Manutenção

### Responsáveis por Setor

| Setor | Módulo | Responsabilidade |
|-------|--------|------------------|
| Contabilidade | Tabulador | Processamento de balancetes |
| Fiscal | Relatórios (8x) | Geração de relatórios fiscais |
| Análise | Análise Fiscal | Visualização e comparação |
| IA | Anonimização | Processamento com IA (Beta) |

### Checklist de Manutenção

- [ ] Validar compatibilidade com novas versões de PyQt6
- [ ] Atualizar pandas para novas versões (breaking changes)
- [ ] Revisar tratamento de erros em cada módulo
- [ ] Documentar novos tipos de relatório fiscal
- [ ] Testar com arquivos reais de produção

---

## 📚 Referências e Recursos

### Documentação Oficial

- [PyQt6 Documentation](https://www.riverbankcomputing.com/static/Docs/PyQt6/)
- [Pandas Documentation](https://pandas.pydata.org/docs/)
- [OpenPyXL Documentation](https://openpyxl.readthedocs.io/)
- [Python XML Processing](https://docs.python.org/3/library/xml.etree.elementtree.html)

### Padrões Fiscais

- [Especificação EFD-C](http://sped.rfb.gov.br/) (Secretaria de Receita Federal)
- [Especificação SPED-F](http://www.fnde.gov.br/) (Alimentar)
- [NF-e Schema XML](https://www.nfe.fazenda.gov.br/)

### Tecnologias Utilizadas

- **Python 3.7+**
- **PyQt6** (GUI)
- **pandas** (Data Processing)
- **openpyxl** (Excel I/O)
- **matplotlib** (Visualization)
- **NumPy** (Numerical Operations)

---

## 🗂️ Versionamento e Histórico

| Versão | Data | Alterações |
|--------|------|-----------|
| 1.0 | Inicial | Estrutura base com Contabilidade, Fiscal, Análise |
| 2.0 | Planejado | Integração de IA, melhorias de performance |

---

**Documento Gerado:** 19/05/2026  
**Status:** Documentação Completa - Versão 1.0  
**Maintainer:** Tim do HUB Auditoria
