# BM MOTOS

Aplicativo em Streamlit para apoiar a rotina de uma oficina de motos, com controle de estoque, registro de vendas e compras, caixa e gerenciamento de mecânicos.

## Sobre o projeto

O BM MOTOS nasceu de uma necessidade prática: concentrar em uma tela simples as informações que normalmente ficam espalhadas entre planilhas e anotações. A aplicação usa Python, Streamlit e Excel para permitir um protótipo rápido e fácil de operar em uma oficina pequena.

Este projeto é um protótipo de uso local. Ele ainda não deve ser tratado como um sistema multiusuário ou como uma solução pronta para operação crítica.

## Funcionalidades

- Consulta e atualização de itens do estoque.
- Registro de compras e vendas.
- Controle básico do caixa.
- Cadastro e consulta de mecânicos.
- Separação de acesso administrativo e acesso de mecânico.
- Persistência dos dados em uma planilha Excel.

## Como rodar localmente

### Requisitos

- Python 3.10 ou superior.
- Pip.
- Uma cópia da planilha de dados do projeto, sem dados reais de clientes.

### Instalação

```powershell
git clone [https://github.com/Gkirst/BMMOTOS.git](https://github.com/Gkirst/BMMOTOS.git)
cd BMMOTOS
python -m venv .venv
.\.venv\Scripts\activate
pip install -r requirements.txt
Copy-Item .env.example .env
