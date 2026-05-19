# Documentação de Dependências - HUB Auditoria

## 📦 Visão Geral de Dependências

Este documento detalha todas as dependências do projeto HUB Auditoria, incluindo bibliotecas Python, dependências do sistema operacional e integração com ferramentas externas.

---

## 🐍 Dependências Python

### Arquivo: `requirements.txt`

```
pandas
PyQt6
openpyxl
```

### Detalhamento das Dependências

#### 1. **Pandas** 📊
- **Versão:** Latest (recomendado: 2.0.0+)
- **Propósito:** Processamento e manipulação de dados
- **Módulos que usam:**
  - `scripts/Contabilidade/tabulador_logic.py` - Processamento de balancetes
  - `scripts/Fiscal/relatorio_*.py` - Processamento de relatórios
  - `scripts/Analise/analise_fiscal_logic.py` - Análise de dados

#### 2. **PyQt6** 🖥️
- **Versão:** Latest (6.0.0+)
- **Propósito:** Framework GUI (Interface Gráfica)
- **Módulos que usam:** Todos os arquivos `*_app.py`

#### 3. **OpenPyXL** 📄
- **Versão:** Latest (3.0.0+)
- **Propósito:** Leitura e escrita de arquivos Excel (.xlsx)
- **Módulos que usam:**
  - `scripts/Contabilidade/tabulador_logic.py`
  - `scripts/Fiscal/relatorio_*.py`
  - `scripts/Analise/analise_fiscal_logic.py`

---

## 🖥️ Dependências do Sistema Operacional

### Sistema Operacional

- **Target:** Windows 7 SP1 ou superior
- **Arquitetura:** x86-64 (64-bit recomendado)
- **RAM Mínima:** 2 GB
- **RAM Recomendada:** 4 GB+
- **Disco:** 500 MB livres

---

## 📊 Compatibilidade de Versões

### Matriz de Compatibilidade Testada

| Python | pandas | PyQt6 | openpyxl | Status |
|--------|--------|-------|----------|--------|
| 3.7 | 1.3.0 | 6.0.0 | 3.7.0 | ✅ OK |
| 3.8 | 1.4.0 | 6.1.0 | 3.8.0 | ✅ OK |
| 3.9 | 2.0.0 | 6.2.0 | 3.9.0 | ✅ OK |
| 3.10 | 2.0.0 | 6.2.0 | 3.10.0 | ✅ OK |
| 3.11 | 2.0.0 | 6.3.0 | 3.10.0 | ✅ OK |

---

**Documento Gerado:** 19/05/2026  
**Versão:** 1.0  
**Status:** Completo