# AtestoFacil

Sistema de previsão de pagamento de colaboradores (Django + DRF + PostgreSQL).

## Stack

- Django 5.2 LTS
- Django REST Framework
- PostgreSQL (produção) / SQLite (desenvolvimento local via `.env`)
- WeasyPrint (PDF; fallback textual se libs do sistema faltarem)
- Bootstrap 5

## Setup rápido

```bash
cd AtestoFacil
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
cp .env.example .env
# Ajuste DB_* no .env. Para PostgreSQL: DB_ENGINE=postgresql
python manage.py migrate
python manage.py createsuperuser
python manage.py runserver
```

Acesse: http://127.0.0.1:8000/

## Fluxo principal

1. Cadastre pessoa, empresa e 3 textos padrões.
2. Em **Serviços → Novo serviço**, informe dados, dias/horas e os 3 textos.
3. O sistema gera o número (`001/2026`) e os 3 PDFs.
4. Na tela do serviço, baixe ou imprima cada documento.

Remessas foram removidas do fluxo da interface: o serviço concentra a contratação e os documentos.

## Variáveis de texto padrão

Padrão principal: `{{ nome }}`, `{{ cpf }}`, `{{ chave_pix }}`, `{{ numero_documento }}`,
`{{ data_documento }}`, `{{ empresa }}`, `{{ dias_trabalhados }}`, `{{ motivo_contratacao }}`,
`{{ funcao }}`, `{{ total_horas }}`.

Alias suportado: `[ NOME ]`, `[ CPF ]`, etc.

## APIs

- `/api/pessoas/`
- `/api/empresas/`
- `/api/textos-padroes/`
- `/api/servicos/` (POST cria serviço + 3 documentos)
- `/api/documentos/`

Autenticação: sessão ou basic auth (usuário Django).

## Testes

```bash
python manage.py test
```

## Numeração

Sequência anual segura com `select_for_update` em `SequenciaAnual` (`SERVICO`).
Formato: `001/2026`.
