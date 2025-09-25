## 📑 CVM_Tool

O CVM_Tool é um software em Python desenvolvido para facilitar a coleta das decisões da Comissão de Valores Mobiliários (CVM) diretamente do site oficial.

Ele automatiza a busca, realiza o download dos PDFs e de seus anexos, e organiza tudo em pastas, otimizando o trabalho de advogados e profissionais da área jurídica que precisam dessas informações no dia a dia.


## 🚀 Funcionalidades

🔎 Busca automática por decisões da CVM

📂 Download em PDF das decisões

📎 Download dos anexos relacionados a cada decisão

🗂️ Organização dos arquivos em pastas estruturadas

🖥️ Interface gráfica amigável desenvolvida com Tkinter


## 🛠️ Tecnologias utilizadas

O projeto foi desenvolvido em Python com uso de bibliotecas de automação, scraping e manipulação de dados.

Principais bibliotecas externas utilizadas:

selenium

webdriver-manager

beautifulsoup4

pandas

requests

pyinstaller


## 📦 Instalação

Clone este repositório:
git clone https://github.com/deVerick06/CVM_Tool.git

Crie e ative um ambiente virtual (opcional, mas recomendado):
python -m venv venv
venv\Scripts\activate 

Instale as dependências:
pip install -r requirements.txt


## ▶️ Como usar

1- Use o pyinstaller para passar o arquivo .py para exe, e rode

2- Use a interface gráfica para:
-Selecionar o período ou critérios de busca
-Iniciar o download das decisões e anexos
-Acompanhar o progresso em tempo real

3- Os arquivos serão salvos em pastas organizadas por decisão.


## 🎯 Público-alvo

Este software foi criado especialmente para:

-Advogados

-Escritórios de advocacia

-Profissionais do setor jurídico que atuam com direito societário e mercado de capitais

-Pesquisadores interessados em decisões da CVM


## 📌 Próximos passos (Roadmap)

- Implementar filtros de busca avançados

- Exportar relatórios em Excel com metadados das decisões

- Versão web (futuro)

- Implementar um agente de IA

## 📜 Licença
Este projeto está licenciado sob a 
[![License: CC BY-NC-ND 4.0](https://img.shields.io/badge/License-CC%20BY--NC--ND%204.0-lightgrey.svg)](https://creativecommons.org/licenses/by-nc-nd/4.0/)


