# Deploy — Lista de Presença RH com Streamlit + Supabase

Este pacote converte o app original de SQLite local para Supabase Postgres.

## Estrutura

```text
app.py
requirements.txt
.gitignore
.streamlit/secrets.toml.example
sql/schema_supabase.sql
```

## 1. Criar projeto no Supabase

1. Crie um projeto no Supabase.
2. Em Project Settings > Database > Connection string, copie a string do Pooler.
3. Para deploy em Streamlit Cloud, prefira o pooler de transaction/shared pooler.
4. Guarde a senha do banco. Ela entra apenas em secrets, nunca no GitHub.

## 2. Configurar banco

O app cria as tabelas automaticamente ao iniciar. Se quiser criar antes manualmente, rode `sql/schema_supabase.sql` no SQL Editor do Supabase.

## 3. Configurar segredo no Streamlit

No Streamlit Community Cloud, abra:

```text
App settings > Secrets
```

Cadastre:

```toml
SUPABASE_DB_URL = "postgresql://postgres.xxxxxxxxx:SUA_SENHA@aws-0-sa-east-1.pooler.supabase.com:6543/postgres"
```

Não suba `.streamlit/secrets.toml` para o GitHub.

## 4. Subir para GitHub

```bash
git init
git add .
git commit -m "deploy streamlit supabase"
git branch -M main
git remote add origin https://github.com/SEU_USUARIO/NOME_DO_REPOSITORIO.git
git push -u origin main
```

## 5. Deploy no Streamlit Community Cloud

1. Conecte sua conta GitHub.
2. Selecione o repositório.
3. Main file path: `app.py`.
4. Configure os Secrets.
5. Faça o deploy.

## Login inicial

O app cria dois usuários iniciais automaticamente:

- `gestor` / `1234`
- `operacao` / `1234`

Troque essas senhas no primeiro acesso pelo menu de gestão de acessos.

## Observação de governança

A connection string é uma credencial sensível. Trate como segredo corporativo. Não coloque no código, README público, print de tela ou commit.
