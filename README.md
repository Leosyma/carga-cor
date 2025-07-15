# COR Data Load and Validation

This repository provides a comprehensive automation and validation pipeline for Operational Cost and Revenue (COR) data used in regulatory and controllership processes. It includes Python scripts for loading datasets from various departments (controllership, legal, accounting, and revenue) and SQL files for schema definitions and regulatory validation views.

---


## 📁 Project Structure

### 1. Python Scripts

- **Carga_dados_PMSO_20231030.py**: Loads controllership PMSO (Planned vs. Actual) data, maps nature/cost class, and applies transformations before insertion into Oracle.
- **Carga_dados_Cobertura_20231030.py**: Loads RRE data regarding regulatory coverage per distributor.
- **Carga_dados_Receita_20231030.py**: Loads budgeted revenue data across distributors and months.
- **Carga_dados_Juridicos_20233010.py**: Processes consolidated legal payments with classification and mappings.
- **Carga_dados_Contabilidade_20231122.py**: Loads and processes accounting records including debit, credit, and balance.
- **Carga_dados_Multas_20231030.py**: Loads fine data by account, company, and month.
- **Automacao_COR_20231031.py**: Coordinates the automation process for PMSO, revenue, and legal data.
- **Automacao_COR_Natureza_20231030.py**: Automates standardization of cost nature based on mappings between departments.
- **Carga_dados_DEPARA_PLUZ_CONTROLADORIA.py**: Loads mapping tables between controllership cost classes and regulatory accounts.
- **Carga_dados_DEPARA_PLUZ_REGULATORIO_20221117.py**: Loads regulatory mappings between PLUZ and COR classifications.

---

### 2. SQL Scripts

- **Criação de tabelas - Oracle.sql**: Defines all Oracle tables used in the project, including PMSO, fines, revenue budget, legal, accounting, and DE-PARA mapping.
- **Analises_SQL.sql**: 
  - Creates the view `VW_COR_VERIFICA_CLASSIFICACAO_COR` to compare COR classifications between regulatory and controllership sources.
  - Creates the view `VW_COR_BASE_ORCADO_REALIZADO` to consolidate budgeted vs. actual cost across months.

---

## 🧩 Features

- 📥 Automated data ingestion from multiple business areas.
- 🧾 DE-PARA (mapping) between regulatory and controllership classifications.
- 🧪 SQL views for cross-validation between sources.
- 🗃️ Monthly cost aggregation and classification validation.
- ✅ Supports regulatory reporting and internal audit routines.

---

## 💾 Requirements

- Python 3.8+
- Oracle database access
- `cx_Oracle`, `pandas`, `os`, `sqlalchemy` and other standard Python libraries

---

## 🏁 Getting Started

1. Clone the repository
2. Set your database credentials
3. Execute the Python loaders in the desired order
4. Deploy the SQL scripts to create schema and views

---

## 📈 Example Outputs

- Cross-tabulation of budgeted vs. actual PMSO per month
- Classification mismatches flagged for auditing
- Aggregated fines and legal payments by month



