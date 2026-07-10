# 📊 PMO Executive Report Generator

> Automação em Excel VBA para geração de Dashboards Executivos no PowerPoint.

![Excel VBA](https://img.shields.io/badge/Excel-VBA-217346?logo=microsoftexcel&logoColor=white)
![PowerPoint](https://img.shields.io/badge/PowerPoint-Automation-B7472A?logo=microsoftpowerpoint&logoColor=white)
![Status](https://img.shields.io/badge/Status-Concluído-success)

---

## 📖 Sobre o projeto

O **PMO Executive Report Generator** é uma solução desenvolvida em **Excel VBA** para automatizar a criação de apresentações executivas em PowerPoint.

O sistema lê os dados de uma tabela estruturada do Excel (`tbProjetos`), realiza o processamento das informações e gera automaticamente dashboards profissionais para acompanhamento do portfólio de projetos.

Todo o relatório é criado com apenas um clique, reduzindo drasticamente o tempo gasto na preparação manual de apresentações para reuniões executivas.

---

# ✨ Principais funcionalidades

✔ Interface amigável (UserForm)

✔ Geração automática de apresentações em PowerPoint

✔ Filtros dinâmicos por período

- Pipeline Completo
- Próximos 30 dias
- Próximos 60 dias
- Próximos 90 dias

✔ Seleção dos slides que serão gerados

- Sumário Executivo
- Análise Regional

✔ Dashboard executivo

- Portfólio Ativo
- Projetos em andamento
- Projetos concluídos

✔ Pipeline de inaugurações por mês

✔ Distribuição por tipo de projeto

✔ Distribuição geográfica por estado

✔ Distribuição por bandeira

✔ Layout corporativo inspirado em apresentações executivas (McKinsey/Bain)

---

# 🖥 Interface

### Painel de geração

> *(<img width="1919" height="833" alt="image" src="https://github.com/user-attachments/assets/048685de-856e-4c4e-816b-fad4a1d82987" />)*


---

# 📊 Dashboard Gerado

### Executive Report

> *(adicione aqui o print do primeiro slide)*

![Dashboard](images/dashboard-status.png)

---

### Análise Regional

> *(adicione aqui o print do segundo slide)*

![Dashboard Regional](images/dashboard-regional.png)

---

# ⚙ Como funciona

O fluxo da aplicação é simples.

```text
Tabela tbProjetos
        │
        ▼
Leitura dos dados
        │
        ▼
Aplicação dos filtros
        │
        ▼
Processamento dos indicadores
        │
        ▼
Criação automática dos slides
        │
        ▼
PowerPoint aberto com relatório pronto
```

---

# 📋 Estrutura da tabela

O projeto utiliza uma tabela estruturada do Excel chamada:

```text
tbProjetos
```

Principais colunas utilizadas:

| Campo |
|-------|
| Projeto |
| Bandeira |
| Cidade |
| Estado |
| Início Obras |
| Término Obras |
| Inauguração |

---

# 📈 Indicadores gerados

O relatório apresenta automaticamente:

- Total do Portfólio
- Projetos em andamento
- Projetos concluídos
- Pipeline mensal
- Distribuição por tipo
- Distribuição por estado
- Distribuição por bandeira

---

# 💻 Tecnologias utilizadas

- Microsoft Excel
- VBA
- PowerPoint Object Library
- ListObjects
- Dictionaries (Scripting.Dictionary)
- COM Automation

---

# 📁 Estrutura do projeto

```
Projeto
│
├── Projeto_VBA_PMO.xlsm
├── GeradorDashboard.bas
├── frmGerador.frm
├── frmGerador.frx
├── images
│   ├── painel.png
│   ├── dashboard-status.png
│   └── dashboard-regional.png
└── README.md
```

---

# 🚀 Como executar

1. Abra o arquivo **Projeto_VBA_PMO.xlsm**

2. Habilite as macros.

3. Certifique-se de que a tabela possui o nome:

```
tbProjetos
```

4. Clique no botão **Painel de Relatório**.

5. Escolha:

- período do pipeline;
- slides desejados.

6. Clique em:

```
GERAR RELATÓRIO NO POWERPOINT
```

O PowerPoint será aberto automaticamente contendo o relatório executivo.

---

# 🎯 Objetivo

Este projeto foi desenvolvido para demonstrar a utilização do VBA na automação de processos corporativos, aplicando conceitos de PMO, Business Intelligence e automação do Microsoft Office.

Além da redução do tempo de elaboração dos relatórios, o projeto garante padronização visual e elimina tarefas repetitivas.

---

# 👨‍💻 Autor

**Davi Lima**

Analista de Planejamento e Expansão | PMO

Especializado em automação de processos com Excel VBA, Power BI e soluções para gerenciamento de projetos.
