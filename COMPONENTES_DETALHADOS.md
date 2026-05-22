# Componentes Detalhados - HUB Auditoria

## 📐 Arquitetura de Componentes

### Diagrama de Camadas

```
┌─────────────────────────────────────────────────────────────┐
│                    User Interface Layer                      │
│  (PyQt6 Widgets - Desktop GUI)                              │
└────────────┬────────────────────────────────────────────────┘
             │
┌────────────▼────────────────────────────────────────────────┐
│                 Application Logic Layer                      │
│  (Business Logic, Data Processing)                          │
└────────────┬────────────────────────────────────────────────┘
             │
┌────────────▼────────────────────────────────────────────────┐
│                   Data Layer                                │
│  (DataFrame manipulation, Excel I/O, XML Parsing)          │
└────────────┬────────────────────────────────────────────────┘
             │
┌────────────▼────────────────────────────────────────────────┐
│              External Dependencies                           │
│  (pandas, openpyxl, matplotlib, PyQt6)                      │
└─────────────────────────────────────────────────────────────┘
```

---

## 🧩 Componentes por Camada

### CAMADA 1: Interface de Usuário (UI Layer)

#### 1.1 Hub Central (`hub.py`)

**Classe Principal:** `Hub(QWidget)`

```
Hub Window
├── Left Panel (240px)
│   ├── Logo (QLabel com imagem)
│   └── Sectors List (QListWidget)
│       ├── "Contabilidade"
│       ├── "Fiscal"
│       ├── "Analise"
│       └── "IA"
│
└── Right Panel (flex)
    ├── Scripts Title (QLabel)
    ├── Scripts List (QListWidget)
    ├── Run Button (QPushButton)
    ├── Log Title (QLabel)
    └── Log Area (QTextEdit - read-only)
```

**Responsabilidades:**
- Renderizar interface principal
- Gerenciar seleção de setores
- Carregar lista de scripts dinamicamente
- Executar scripts selecionados
- Exibir logs de execução
- Gerenciar referências a janelas abertas

**Métodos Públicos:**
```python
def load_scripts(item: QListWidgetItem) -> None
    """Carrega scripts para setor selecionado"""
    
def run_script() -> None
    """Executa script selecionado dinamicamente"""
    
def set_status(text: str) -> None
    """Atualiza status no log"""
```

#### 1.2 Splash Screen (`hub.py`)

**Classe Principal:** `Splash(QWidget)`

```
Splash Window
├── Background (Cinza #333333)
├── Logo (QPixmap - 400x300)
└── Status Label (QLabel com mensagem)
```

**Responsabilidades:**
- Exibir tela de boas-vindas durante inicialização
- Mostrar mensagens de status de carregamento
- Animar transição (fade in/out com QPropertyAnimation)
- Desaparecer após carregamento completo (2 segundos)

**Métodos Públicos:**
```python
def set_status(text: str) -> None
    """Atualiza mensagem de status"""
```

#### 1.3 Tabulador Interface (`scripts/Contabilidade/tabulador_app.py`)

**Classe Principal:** `TabuladorPythonApp(QWidget)`

```
Tabulador Window
├── Section 1: File Selection
│   ├── Label "1. Selecione um ou mais arquivos..."
│   ├── File List Widget (Multi-select)
│   └── Browse Button
│
├── Section 2: System Selection
│   ├── Label "2. Selecione o sistema:"
│   └── System Combo Box (Fuga, Uniair)
│
└── Section 3: Execution & Status
    ├── Run Button (Bold, 12pt)
    └── Status Display
```

**Responsabilidades:**
- Permitir seleção de múltiplos arquivos Excel/CSV
- Exibir lista de arquivos selecionados
- Selecionar sistema (Fuga/Uniair)
- Iniciar processamento
- Exibir status de progresso

#### 1.4 XML Relatório Interface (`scripts/Fiscal/relatorio_XML_app.py`)

**Classe Principal:** `RelatorioXMLApp(QWidget)` (implícita)

```
XML Analyzer Window
├── File Selection Panel
│   ├── Label "Selecione arquivo XML NF-e:"
│   ├── File Path Display
│   └── Browse Button
│
├── Data Display Panel
│   ├── Raw XML (opcional)
│   ├── Extracted Data (DataFrame view)
│   ├── Tax Summary
│   └── Totals
│
└── Export Panel
    ├── Format Selection (Excel/CSV)
    └── Export Button
```

**Responsabilidades:**
- Carregar e validar arquivos XML NF-e
- Extrair dados com segurança (namespace handling)
- Calcular impostos (ICMS, IPI, PIS, COFINS)
- Exibir dados estruturados
- Exportar resultados

#### 1.5 Análise Fiscal Interface (`scripts/Analise/analise_fiscal_app.py`)

**Classe Principal:** `AnaliseFiscalApp(QWidget)`

```
Análise Fiscal Window
├── Header
│   ├── Title Label
│   └── Menu Button (⋮)
│       ├── Exportar PDF
│       ├── Filtrar Dados
│       └── Configurações
│
├── Main Area (QTabWidget)
│   ├── Tab 1: Raw Data
│   │   ├── Tree Widget
│   │   │   ├── Período 1
│   │   │   ├─ Dados
│   │   │   ├─ Totais
│   │   │   ├── Período 2
│   │   │   └─ ...
│   │   └── Expand/Collapse controls
│   │
│   ├── Tab 2: Visualizações
│   │   ├── Graph 1 (Linhas)
│   │   ├── Graph 2 (Barras)
│   │   └── Graph 3 (Pizza)
│   │
│   └── Tab 3: Resumo Executivo
│       ├── KPI Cards
│       ├── Tendências
│       └── Recomendações
│
└── Footer
    ├── Status Bar
    └── Log de Execução
```

**Responsabilidades:**
- Carregar múltiplos arquivos de dados
- Fazer análise comparativa entre períodos
- Gerar visualizações interativas (Matplotlib)
- Exportar resultados em PDF
- Fornecer interface de filtros e buscas

---

### CAMADA 2: Lógica de Aplicação (Logic Layer)

#### 2.1 Hub Logic (Dinâmica)

**Funções Principais:**
```python
def start_application(app: QApplication, splash: Splash) -> None
    """
    Carrega módulos pesados e inicializa a aplicação.
    - Importa pandas
    - Importa openpyxl
    - Cria instância Hub
    - Exibe após 500ms
    """
    
def importlib.import_module(module_path: str) -> ModuleType
    """Carrega módulo dinâmico em tempo de execução"""
```

#### 2.2 Tabulador Logic (`scripts/Contabilidade/tabulador_logic.py`)

**Funções Principais:**
```python
def processar_arquivos_selecionados(
    arquivos: List[str],
    sistema: str
) -> pd.DataFrame
    """
    Processa múltiplos arquivos de balancete.
    
    Workflow:
    1. Validar arquivos (existência, formato)
    2. Ler cada arquivo em DataFrame
    3. Normalizar nomes de colunas
    4. Aplicar transformações específicas do sistema
    5. Consolidar em DataFrame único
    6. Retornar resultado consolidado
    """

def normalizar_colunas(df: pd.DataFrame, sistema: str) -> pd.DataFrame
    """Padroniza nomes de colunas conforme sistema"""

def aplicar_transformacoes(
    df: pd.DataFrame,
    sistema: str
) -> pd.DataFrame
    """Aplica transformações específicas do sistema"""

def consolidar_dados(
    dataframes: List[pd.DataFrame]
) -> pd.DataFrame
    """Consolida múltiplos DataFrames em um único"""

def exportar_resultado(
    df: pd.DataFrame,
    caminho_saida: str
) -> None
    """Exporta resultado para Excel com formatação"""
```

#### 2.3 XML Handler Logic (`scripts/Fiscal/relatorio_XML_app.py`)

**Funções Principais:**
```python
def safe_get(
    element: ET.Element,
    path: str,
    default: str = ''
) -> str
    """
    Extração segura de elemento XML.
    
    Benefícios:
    - Trata namespaces corretamente
    - Retorna valor padrão se não encontrar
    - Não causa exceções
    """

def safe_get_float(value_str: str) -> float
    """Conversão segura string → float"""

def get_specific_tax(
    imposto_element: ET.Element,
    tax_path_map: Dict[str, str]
) -> Dict[str, float]
    """Extrai impostos específicos (ICMS, IPI, etc)"""

def processar_nf(
    arquivo_xml: str
) -> pd.DataFrame
    """Processa uma NF-e completa"""

def consolidar_nfs(
    arquivos_xml: List[str]
) -> pd.DataFrame
    """Consolida múltiplas NF-es"""
```

#### 2.4 Análise Logic (`scripts/Analise/analise_fiscal_logic.py`)

**Funções Principais:**
```python
def carregar_dados(
    arquivo: str
) -> pd.DataFrame
    """Carrega dados de arquivo Excel/CSV"""

def fazer_analise_comparativa(
    df1: pd.DataFrame,
    df2: pd.DataFrame,
    coluna_comparacao: str
) -> Dict
    """Compara dois períodos/entidades"""

def gerar_grafico(
    dados: pd.DataFrame,
    tipo_grafico: str,
    titulo: str
) -> Figure
    """Gera gráfico usando Matplotlib"""

def exportar_pdf(
    figuras: List[Figure],
    caminho_saida: str
) -> None
    """Exporta figuras em PDF com PdfPages"""

def calcular_kpis(
    df: pd.DataFrame
) -> Dict[str, float]
    """Calcula KPIs (Key Performance Indicators)"""
```

---

### CAMADA 3: Dados (Data Layer)

#### 3.1 pandas Data Processing

**Operações Principais:**

```python
# Leitura
df = pd.read_excel('arquivo.xlsx', sheet_name='Planilha1')
df = pd.read_csv('arquivo.csv', delimiter=';', encoding='utf-8')

# Transformação
df['nova_coluna'] = df['coluna_existente'].apply(funcao)
df_filtrado = df[df['coluna'] > 100]
df_agrupado = df.groupby('categoria').sum()

# Consolidação
df_consolidado = pd.concat([df1, df2, df3])
df_merged = pd.merge(df1, df2, on='chave_comum')

# Validação
df_validado = df.dropna()
df_validado = df[df.notna().all(axis=1)]

# Exportação
df.to_excel('saida.xlsx', index=False, engine='openpyxl')
df.to_csv('saida.csv', index=False, encoding='utf-8')
```

#### 3.2 XML Data Extraction

```python
import xml.etree.ElementTree as ET

# Parse XML
tree = ET.parse('arquivo.xml')
root = tree.getroot()

# Namespace handling
ns = {'nfe': 'http://www.portalfiscal.inf.br/nfe'}
elemento = root.find('nfe:NFe', ns)

# Extração segura
texto = elemento.find('nfe:tag', ns).text if elemento else ''

# Conversão
valor = float(texto) if texto else 0.0
```

#### 3.3 Excel I/O (openpyxl)

```python
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

# Criar novo workbook
wb = Workbook()
ws = wb.active

# Adicionar dados
ws['A1'] = 'Cabeçalho'
ws['A2'] = 100

# Formatação
ws['A1'].font = Font(bold=True, size=12, color="FFFFFF")
ws['A1'].fill = PatternFill(start_color="4285F4", fill_type="solid")
ws['A1'].alignment = Alignment(horizontal='center', vertical='center')

# Dimensionar coluna
ws.column_dimensions['A'].width = 20

# Salvar
wb.save('resultado.xlsx')

# Carregar e modificar
wb_exist = load_workbook('existente.xlsx')
ws_exist = wb_exist.active
```

#### 3.4 Matplotlib Visualization

```python
from matplotlib.figure import Figure
from matplotlib.backends.backend_pdf import PdfPages
import matplotlib.pyplot as plt

# Criar figura
fig = Figure(figsize=(10, 6), dpi=100)
ax = fig.add_subplot(111)

# Plotar dados
ax.plot(x, y, label='Série 1', marker='o')
ax.bar(categorias, valores, label='Série 2')

# Customizar
ax.set_title('Título do Gráfico', fontsize=14, fontweight='bold')
ax.set_xlabel('Eixo X', fontsize=12)
ax.set_ylabel('Eixo Y', fontsize=12)
ax.legend()
ax.grid(True, alpha=0.3)

# Exportar PDF
with PdfPages('relatorio.pdf') as pdf:
    pdf.savefig(fig, bbox_inches='tight')
```

---

## 📊 Matriz de Responsabilidades

### Por Setor

| Setor | Módulo | Responsabilidade | Input | Output |
|-------|--------|------------------|-------|--------|
| **Contabilidade** | Tabulador | Consolidar balancetes | CSV/XLSX (múltiplos) | XLSX consolidado |
| **Fiscal** | C100 EFD-C | Relatório de operações | Arquivo EFD-C | XLSX relatório |
| **Fiscal** | C100 SPED-F | Relatório de operações | Arquivo SPED-F | XLSX relatório |
| **Fiscal** | C500 EFD-C | Consolidação | Arquivo EFD-C | XLSX consolidado |
| **Fiscal** | C500 SPED-F | Consolidação | Arquivo SPED-F | XLSX consolidado |
| **Fiscal** | CIAP SPED-F | ICMS Apropriação | Arquivo SPED-F | XLSX análise |
| **Fiscal** | CT | Controladoria | Dados CT | XLSX relatório |
| **Fiscal** | D100 SPED-F | Docs Fiscais | Arquivo SPED-F | XLSX relatório |
| **Fiscal** | F100 EFD-C | Operações Internas | Arquivo EFD-C | XLSX relatório |
| **Fiscal** | XML NF-e | Análise de NF-e | XML NF-e | XLSX com impostos |
| **Análise** | Análise Fiscal | Comparação de períodos | XLSX (múltiplos) | Gráficos + PDF |
| **IA** | Anonimização | Anonimizar textos | Texto | Texto anonimizado |

### Por Camada de Código

| Componente | UI | Lógica | Dados | Dependências |
|-----------|----|----|-------|--------------|
| hub.py | ✅ Principal | ✅ Importação dinâmica | ❌ | PyQt6 |
| tabulador_app.py | ✅ Principal | ❌ | ❌ | PyQt6, pandas |
| tabulador_logic.py | ❌ | ✅ Principal | ✅ | pandas, openpyxl |
| relatorio_XML_app.py | ✅ UI + Lógica | ✅ XML + Cálculos | ✅ | PyQt6, pandas, ElementTree |
| analise_fiscal_app.py | ✅ Principal | ✅ Integração | ✅ | PyQt6, pandas, matplotlib |
| analise_fiscal_logic.py | ❌ | ✅ Principal | ✅ | pandas, numpy |

---

## 🔄 Fluxo de Dados Detalhado

### Exemplo 1: Tabulador (Contabilidade)

```
┌────────────────────────────────┐
│ Usuário Seleciona Arquivos     │
│ ex: [balancete_01.xlsx,        │
│      balancete_02.xlsx]        │
└────────────┬───────────────────┘
             │
             ▼
┌────────────────────────────────┐
│ Usuário Seleciona Sistema      │
│ ex: "Fuga"                     │
└────────────┬───────────────────┘
             │
             ▼
┌─────────────────────────────────────────────────┐
│ tabulador_app.run_script()                      │
└────────────┬────────────────────────────────────┘
             │
             ▼
┌─────────────────────────────────────────────────┐
│ processar_arquivos_selecionados(                │
│   ["balancete_01.xlsx", "balancete_02.xlsx"],  │
│   "Fuga"                                        │
│ )                                               │
└────────────┬────────────────────────────────────┘
             │
    ┌────────┴────────┐
    │                 │
    ▼                 ▼
┌──────────────┐  ┌──────────────┐
│ Ler arquivo1 │  │ Ler arquivo2 │
│ → DataFrame  │  │ → DataFrame  │
└──────────────┘  └──────────────┘
    │                 │
    │     ┌───────────┘
    │     │
    ▼     ▼
┌──────────────────────┐
│ Normalizar Colunas   │
│ (Fuga → Padrão)      │
└──────────┬───────────┘
           │
    ┌──────┴──────┐
    │             │
    ▼             ▼
┌──────────┐  ┌──────────┐
│ DataFrame│  │ DataFrame│
│  Normal1 │  │  Normal2 │
└──────────┘  └──────────┘
    │             │
    │   ┌─────────┘
    │   │
    ▼   ▼
┌──────────────────────┐
│ Aplicar Transformações
│ (específicas Fuga)   │
└──────────┬───────────┘
           │
    ┌──────┴──────┐
    │             │
    ▼             ▼
┌──────────┐  ┌──────────┐
│ DataFrame│  │ DataFrame│
│ Tran1    │  │ Tran2    │
└──────────┘  └──────────┘
    │             │
    │   ┌─────────┘
    │   │
    ▼   ▼
┌──────────────────┐
│ Consolidar Dados │
│ pd.concat([...]) │
└──────────┬───────┘
           │
           ▼
┌──────────────────────┐
│ DataFrame Consolidado│
│ (Resultado Final)    │
└──────────┬───────────┘
           │
           ▼
┌──────────────────────┐
│ Exportar para Excel  │
│ → balancete_final.   │
│   xlsx               │
└──────────────────────┘
           │
           ▼
┌──────────────────────┐
│ Usuário Baixa Arquivo│
│ & Revisa Resultado   │
└──────────────────────┘
```

### Exemplo 2: XML NF-e (Fiscal)

```
┌────────────────────────────┐
│ Usuário Seleciona XML      │
│ ex: nfe.xml                │
└────────────┬───────────────┘
             │
             ▼
┌────────────────────────────────┐
│ relatorio_XML_app.processar()  │
└────────────┬───────────────────┘
             │
             ▼
┌────────────────────────────┐
│ ET.parse('nfe.xml')        │
│ → ElementTree root         │
└────────────┬───────────────┘
             │
             ▼
┌──────────────────────────────────────┐
│ Para cada transação na NF-e:         │
│  1. safe_get(tag, namespace)         │
│  2. Extrair campos principais        │
│  3. Extrair impostos                 │
└────────────┬──────────────────────────┘
             │
             ▼
┌──────────────────────┐
│ Consolidar em Dict   │
│ {                    │
│   'nfnumber': ...,   │
│   'total': ...,      │
│   'icms': ...,       │
│   'ipi': ...,        │
│   ...                │
│ }                    │
└────────────┬─────────┘
             │
             ▼
┌──────────────────────┐
│ Converter para       │
│ DataFrame            │
└────────────┬─────────┘
             │
             ▼
┌──────────────────────┐
│ Exportar para Excel  │
│ com formatação       │
└──────────────────────┘
```

### Exemplo 3: Análise Fiscal (Análise)

```
┌────────────────────────────┐
│ Usuário Carrega 2 Arquivos │
│ ex: [jan.xlsx, fev.xlsx]   │
└────────────┬───────────────┘
             │
    ┌────────┴────────┐
    │                 │
    ▼                 ▼
┌──────────────┐  ┌──────────────┐
│ pd.read_     │  │ pd.read_     │
│ excel        │  │ excel        │
│ (jan.xlsx)   │  │ (fev.xlsx)   │
└──────────────┘  └──────────────┘
    │                 │
    ▼                 ▼
┌──────────┐     ┌──────────┐
│ df_jan   │     │ df_fev   │
└──────────┘     └──────────┘
    │                 │
    │   ┌─────────────┘
    │   │
    ▼   ▼
┌──────────────────────┐
│ Fazer Análise        │
│ Comparativa          │
│ - Diferenças         │
│ - Percentuais        │
│ - Tendências         │
└──────────┬───────────┘
           │
    ┌──────┴───────────────┐
    │                      │
    ▼                      ▼
┌─────────────────┐  ┌──────────────┐
│ Calcular KPIs   │  │ Gerar Dados  │
│ - Total         │  │ para Gráficos│
│ - Média         │  │              │
│ - Variação      │  │              │
└─────────────────┘  └──────────────┘
    │                      │
    │                      ▼
    │              ┌──────────────────────┐
    │              │ Figure 1: Linhas     │
    │              │ Figure 2: Barras     │
    │              │ Figure 3: Pizza      │
    │              └──────────────────────┘
    │                      │
    │   ┌──────────────────┘
    │   │
    ▼   ▼
┌──────────────────────┐
│ Compilar Resultado:  │
│ - KPIs              │
│ - Figuras           │
│ - Resumo Executivo  │
└──────────┬───────────┘
           │
           ▼
┌──────────────────────────┐
│ Exportar PDF com PdfPages│
│ → relatorio_analise.pdf  │
└──────────────────────────┘
           │
           ▼
┌──────────────────────┐
│ Usuário Baixa PDF &  │
│ Revisa Análise       │
└──────────────────────┘
```

---

## 🔌 Extensibilidade

### Como Adicionar Novo Setor

```python
# 1. Criar pasta
# scripts/NovoSetor/

# 2. Registrar no hub.py
# Linha 112: adicionar "NovoSetor" à lista

# 3. Implementar módulo
# scripts/NovoSetor/novo_modulo_app.py

class NovoModuloApp(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()
    
    def initUI(self):
        # Implementar UI
        pass

def main():
    app = QApplication(sys.argv)
    window = NovoModuloApp()
    window.show()
    return (app, window)

# 4. Verificar carregamento
# Hub automaticamente descobrirá o novo módulo
```

### Como Adicionar Novo Relatório Fiscal

```python
# 1. Criar arquivo
# scripts/Fiscal/relatorio_NOVO_app.py

class RelatorioNOVOApp(QWidget):
    # Mesmo padrão das outras classes
    pass

def main():
    app = QApplication(sys.argv)
    window = RelatorioNOVOApp()
    window.show()
    return (app, window)

# 2. Hub descobrirá automaticamente
#    por convenção de nomeação (*_app.py)
```

---

## 📈 Escalabilidade

### Performance Atual
- Aplicação: ~100MB RAM
- Carregamento: ~2-3 segundos
- Processamento: ~1-5 segundos por arquivo

### Gargalos Identificados
- Importação pandas/openpyxl (mitigado com lazy loading)
- Processamento sequencial de múltiplos arquivos
- Matplotlib em gráficos complexos

### Melhorias Recomendadas
1. Threading para processamento não-bloqueante
2. Multiprocessing para múltiplos arquivos
3. Cache de resultados intermediários
4. Índices em análises comparativas
5. Compressão de dados em memória

---

## 🔐 Segurança de Componentes

### Validações Implementadas
- ✅ Extração segura de XML (safe_get)
- ✅ Conversão segura de tipos (safe_get_float)
- ✅ Try-except em importações dinâmicas
- ✅ Validação de formato de arquivo

### Recomendações
- [ ] Sanitização de entrada de usuário
- [ ] Validação de schema XML
- [ ] Detecção de anomalias em dados
- [ ] Auditoria de operações
- [ ] Criptografia de dados sensíveis

---

**Documento Gerado:** 19/05/2026  
**Versão:** 1.0  
**Status:** Completo
