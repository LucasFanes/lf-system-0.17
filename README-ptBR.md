LF System

Sistema administrativo desenvolvido em Python para gerenciamento de contas, registro de compras, exportação de planilhas e integração com Google Sheets.

Sobre o projeto

O LF System é uma aplicação desenvolvida com o objetivo de praticar conceitos de desenvolvimento de software, programação orientada a objetos, manipulação de arquivos, bancos de dados SQLite e automação de processos.

O sistema permite:

Cadastro de contas
Registro de compras
Histórico de movimentações
Monitoramento de clientes
Exportação para Excel
Upload e download de planilhas do Google Sheets
Criação de backups em ZIP
Registro de logs para auditoria
Armazenamento persistente utilizando SQLite
Tecnologias utilizadas
Python 3
SQLite
OpenPyXL
EZSheets
PyMsgBox
Send2Trash
Programação Orientada a Objetos (POO)
Estrutura do Projeto
LF System
│
├── Main.py
├── Billing.py
├── Function.py
├── Logs.py
├── VERIFICAR_BANCO.py
│
├── Data/
├── Backups/
├── Spreadsheets/
└── Logs/
Funcionalidades
Gestão de contas
Criação de contas
Edição de informações
Consulta de registros
Monitoramento de clientes cadastrados
Gestão de compras
Registro de compras
Histórico completo
Exclusão por código NC
Consulta global de movimentações
Planilhas
Exportação automática para Excel
Integração com Google Sheets
Download em múltiplos formatos:
XLSX
CSV
PDF
ODS
TSV
Backup e Arquivos
Compactação ZIP
Cópia de arquivos
Cópia de diretórios
Exclusão segura para lixeira
Exclusão permanente com confirmação
Banco de Dados

O sistema utiliza SQLite para armazenamento local.

Tabelas principais:

accounts
Campo	Descrição
name	Nome da conta
code	Código interno
created_at	Data de criação
billing_items	Dados do cliente
purchases
Campo	Descrição
nc	Código da compra
account_name	Conta associada
item	Produto
price	Valor
date	Data
Como executar
Instalar dependências
pip install pyperclip
pip install pymsgbox
pip install send2trash
pip install openpyxl
pip install ezsheets
Executar o sistema
python Main.py
Objetivo Educacional

Este projeto foi desenvolvido para aprimorar conhecimentos em:

Desenvolvimento Python
Programação Orientada a Objetos
Banco de Dados SQLite
Manipulação de arquivos
Automação de tarefas
Estruturação de projetos
Integração com APIs e planilhas
Autor

Lucas Fanes

Estudante de Ciência da Computação e Engenharia da Computação na FMU.

Atualmente em busca de oportunidades de estágio e desenvolvimento de software para aplicar e expandir conhecimentos em programação, automação e tecnologia.

## Licença

Este projeto é disponibilizado exclusivamente para fins educacionais, acadêmicos e de portfólio.

É permitida a visualização, análise e estudo do código-fonte para aprendizado pessoal.

Não é permitida a utilização comercial deste software, nem sua redistribuição, revenda, sublicenciamento ou incorporação em produtos ou serviços comerciais sem autorização prévia e por escrito do autor.

© Lucas Dalla Rosa Fanes. Todos os direitos reservados.
READ.md
