# Salão Financeiro V2 (Flask + SQLite)

Versão **v2** do sistema de controle financeiro para salão de beleza, agora com cara de sistema de verdade e menos cara de protótipo apressado.

## O que entrou na v2

- **login** com usuário administrador -[telade login](image.png)
- **edição e exclusão** de serviços, colaboradores, atendimentos, despesas e compras
- **importação de planilha Excel pela interface** [tela principal](image-1.png)
- **dashboard com gráficos** (Chart.js)  [tela dash](image-2.png)
- **exportação mensal em Excel**
- configuração por variáveis de ambiente
- arquivos de apoio para **deploy** (`wsgi.py` + `Procfile`)
- tipos de serviços realizados, podendo ser cadastrado conforme necessidade [tela de serviços](image-3.png)
- registros  de despesas [tela de despesas ](image-4.png)
- relatorios de receitas entradas/saidas/comissões. exportar para excell os dadosde fechamento [tela de receitas](image-5.png)




## Stack

- **Python 3.11+**
- **Flask**
- **Flask-SQLAlchemy**
- **SQLite** (ou outro banco via `DATABASE_URL`)
- **openpyxl**
- **gunicorn** para deploy

## Estrutura

```text
salao_financeiro_v2_flask/
├─ app.py
├─ wsgi.py
├─ Procfile
├─ requirements.txt
├─ README.md
├─ instance/                  # banco SQLite criado automaticamente
├─ static/
│  └─ style.css
├─ templates/
│  ├─ base.html
│  ├─ login.html
│  ├─ index.html
│  ├─ services.html
│  ├─ collaborators.html
│  ├─ attendances.html
│  ├─ expenses.html
│  ├─ products.html
│  ├─ report.html
│  └─ import_excel.html
└─ scripts/
   └─ import_excel_template.py
```

## Como rodar

### 1) Criar ambiente virtual

No Windows:
```bash
python -m venv .venv
.venv\Scripts\activate
```

No Linux/macOS:
```bash
python -m venv .venv
source .venv/bin/activate
```

### 2) Instalar dependências
```bash
pip install -r requirements.txt
```

### 3) Executar
```bash
python app.py
```

Acesse:
- `http://127.0.0.1:5000`

## Login padrão

- **Usuário:** `admin`
- **Senha:** `admin123`

Você pode trocar isso antes da primeira execução:

No Windows (PowerShell):
```powershell
$env:ADMIN_USERNAME="ismael"
$env:ADMIN_PASSWORD="uma_senha_forte"
```

No Linux/macOS:
```bash
export ADMIN_USERNAME="ismael"
export ADMIN_PASSWORD="uma_senha_forte"
```

## Variáveis de ambiente úteis

- `SECRET_KEY` → chave da sessão
- `DATABASE_URL` → permite usar PostgreSQL/MySQL/SQLite em produção
- `ADMIN_USERNAME` → usuário inicial
- `ADMIN_PASSWORD` → senha inicial

## Rotas principais

- `/login` → login
- `/logout` → sair
- `/` → dashboard
- `/services` → serviços
- `/collaborators` → colaboradores
- `/attendances` → atendimentos
- `/expenses` → despesas
- `/products` → compras de produtos
- `/report` → fechamento mensal
- `/import-excel` → upload da planilha pela interface
- `/export/monthly.xlsx?year=2026&month=2` → exportação mensal
- `/seed-demo` → carrega dados de teste

## Como funciona o cálculo

No lançamento de um atendimento:

- Se o usuário informar `% comissão aplicada`, o sistema usa esse valor.
- Se não informar:
  - usa `% comissão do serviço`, se existir
  - senão usa `% padrão do colaborador`

Fórmulas:
- **Comissão = Valor cobrado × (% comissão / 100)**
- **Receita do salão = Valor cobrado - Comissão**
- **Resultado do mês = Receita do salão - Despesas - Compras de produtos**

## Importar pela interface

A tela **Importar Excel** aceita a planilha `.xlsx` com estas abas:

- `tbServicos`
- `tbColaboradores`
- `tbAtendimentos`
- `tbDespesas`
- `tbProdutos`

Se marcar **"Limpar dados atuais antes de importar"**, o sistema apaga os registros existentes e depois importa os novos.

## Importar via script

### Uso básico
```bash
python scripts/import_excel_template.py "Controle_Financeiro_Salao_Template_Melhorado.xlsx"
```

### Limpando dados antes de importar
```bash
python scripts/import_excel_template.py "Controle_Financeiro_Salao_Template_Melhorado.xlsx" --reset
```

## Deploy rápido

### Render / Railway / similares
O projeto já vem com:

- `wsgi.py`
- `Procfile`

Comando web:
```bash
gunicorn wsgi:app
```

## Próximos passos naturais

- múltiplos usuários e permissões
- reset de senha
- filtros avançados por período
- exportação em PDF
- API REST
- anexos e comprovantes
- integração com Power BI
- emissão de recibos

Essa v2 já é um salto real: sai do “funciona aqui em casa” e entra no território do sistema que aguenta rotina.
