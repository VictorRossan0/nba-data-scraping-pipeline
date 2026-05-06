# NBA Data Scraping Pipeline

Pipeline desenvolvido em **Python** para coleta automatizada de dados da NBA, processamento das informações e geração de planilhas consolidadas em Excel.

O projeto realiza scraping e organização de estatísticas esportivas, permitindo acompanhar classificações, líderes individuais, líderes por equipe, calendários e métricas avançadas da NBA em um único arquivo estruturado.

---

## 🎯 Objetivo do projeto

Este projeto foi criado para automatizar a coleta e consolidação de dados esportivos da NBA, reduzindo trabalho manual e facilitando análises estatísticas através de planilhas organizadas.

O foco técnico do projeto está em:

- Coleta automatizada de dados
- Processamento e organização de informações
- Geração automatizada de arquivos Excel
- Consolidação de múltiplas fontes de dados
- Estrutura modular para scraping esportivo

---

## ✨ Funcionalidades

- Extração automatizada de dados da NBA
- Coleta de calendário de jogos
- Coleta de classificação dos times
- Coleta de líderes estatísticos
- Coleta de líderes por equipe
- Processamento de double-double e triple-double
- Consolidação automática das informações
- Geração de planilha Excel final consolidada
- Ajuste automático de largura das colunas
- Organização dos dados em múltiplas abas

---

## 🛠️ Tecnologias utilizadas

- Python
- OpenPyXL
- Web Scraping
- Manipulação de Excel
- Estrutura modular em Python
- Automação de processamento de dados

---

## 🧱 Estrutura técnica

O projeto foi dividido em módulos responsáveis por diferentes tipos de coleta e processamento:

- `calendario_nba.py` → calendário e eventos
- `leaders_nba.py` → líderes estatísticos
- `leaders_teams_nba.py` → classificação de equipes
- `teams_nba_leaders.py` → líderes por equipe
- `doubled_tripled_nba.py` → estatísticas avançadas
- `main.py` → orquestração e consolidação dos dados

O sistema gera múltiplos arquivos intermediários em Excel e posteriormente consolida todas as informações em uma única planilha final.

---

## 📊 Fluxo de processamento

1. Coleta automatizada dos dados
2. Processamento das informações
3. Geração de arquivos Excel intermediários
4. Consolidação das abas
5. Ajuste automático de layout
6. Exportação do arquivo final

---

## 📁 Estrutura de saída

O pipeline gera arquivos como:

```txt
Excel/
├── informacoes_eventos.xlsx
├── classificacao_nba.xlsx
├── leaders_nba.xlsx
├── doubledouble.xlsx
├── tripledouble.xlsx
├── team_leaders_nba.xlsx
└── planilha_NBA.xlsx
```

---

## 🚀 Como executar o projeto

### Pré-requisitos

- Python 3.10+
- Pip

### Instalação

```bash
# Clone o repositório
git clone https://github.com/VictorRossan0/nba-data-scraping-pipeline.git

# Acesse a pasta
cd nba-data-scraping-pipeline

# Instale as dependências
pip install -r requirements.txt
```

### Execução

```bash
python main.py
```

Após a execução, os arquivos Excel serão gerados automaticamente dentro da pasta `Excel`.

---

## ⚙️ Possíveis melhorias futuras

- Exportação para banco de dados
- Dashboard interativo
- APIs esportivas em tempo real
- Integração com Streamlit
- Processamento assíncrono
- Agendamento automático de execução

---

## 📌 Observação

Este projeto foi desenvolvido para estudos, automação de coleta de dados e análise esportiva utilizando Python.

Nenhum dado sensível ou informação corporativa foi utilizado no repositório público.

---

## 👨‍💻 Autor

**Victor Rossano Couto do Amaral**

- GitHub: https://github.com/VictorRossan0
- LinkedIn: https://www.linkedin.com/in/victor-rossano-009b4556/
- Portfólio: https://victorrossano-dev.netlify.app