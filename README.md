# SEI Automation Platform - Django

Plataforma web em Django para automação de buscas no Sistema Eletrônico de Informações (SEI) com armazenamento em PostgreSQL.

## 🚀 Características

- **Interface Web Moderna**: Dashboard intuitivo com Bootstrap 5
- **Automação de Buscas**: Coleta automatizada de documentos do SEI
- **Processamento Assíncrono**: Tarefas executadas em background com Celery
- **Banco de Dados PostgreSQL**: Armazenamento persistente de buscas e documentos
- **Exportação para Excel**: Download de resultados em formato XLSX
- **Autenticação de Usuários**: Sistema de login e gerenciamento de usuários
- **Configurações Reutilizáveis**: Salve configurações de acesso para uso futuro

## 📋 Pré-requisitos

- Python 3.8+
- PostgreSQL 12+
- Redis (para Celery)
- Chrome/Chromium (para Selenium)

## 🔧 Instalação

### 1. Clone o repositório

```bash
git clone https://github.com/lab-dados-seges/sei_aspar.git
cd sei_aspar
```

### 2. Instale as dependências

```bash
pip install -r requirements.txt
```

### 3. Configure o PostgreSQL

Crie um banco de dados PostgreSQL:

```sql
CREATE DATABASE sei_database;
CREATE USER sei_user WITH PASSWORD 'sei_password';
ALTER ROLE sei_user SET client_encoding TO 'utf8';
ALTER ROLE sei_user SET default_transaction_isolation TO 'read committed';
ALTER ROLE sei_user SET timezone TO 'America/Sao_Paulo';
GRANT ALL PRIVILEGES ON DATABASE sei_database TO sei_user;
```

### 4. Configure as variáveis de ambiente

Edite o arquivo `sei_platform/settings.py` e ajuste as configurações do banco de dados:

```python
DATABASES = {
    'default': {
        'ENGINE': 'django.db.backends.postgresql',
        'NAME': 'sei_database',
        'USER': 'sei_user',
        'PASSWORD': 'sei_password',
        'HOST': 'localhost',
        'PORT': '5432',
    }
}
```

### 5. Execute as migrações

```bash
python manage.py makemigrations
python manage.py migrate
```

### 6. Crie um superusuário

```bash
python manage.py createsuperuser
```

### 7. Inicie o servidor Redis

```bash
redis-server
```

### 8. Inicie o Celery Worker (em outro terminal)

```bash
celery -A sei_platform worker -l info
```

### 9. Inicie o servidor Django

```bash
python manage.py runserver
```

Acesse a aplicação em: http://localhost:8000

## 📱 Uso

### Dashboard

Após fazer login, você verá o dashboard principal com:
- Total de buscas realizadas
- Total de documentos coletados
- Buscas recentes

### Nova Busca

1. Clique em "Nova Busca"
2. Preencha os campos:
   - URL do SEI
   - Usuário SEI
   - Senha
   - Órgão
   - Termos de pesquisa
   - Datas (opcional)
3. Clique em "Executar Busca"

A busca será processada em background e você pode acompanhar o status.

### Visualizar Resultados

- Acesse "Minhas Buscas" para ver todas as buscas
- Clique em "Ver" para ver os detalhes e documentos encontrados
- Clique em "Excel" para exportar os resultados

### Configurações

Salve configurações de acesso frequentes:
1. Acesse "Configurações"
2. Preencha os dados de acesso
3. Marque como "Ativa" para usar automaticamente em novas buscas

## 🏗️ Estrutura do Projeto

```
sei_aspar/
├── manage.py
├── sei_platform/           # Configurações do projeto Django
│   ├── settings.py
│   ├── urls.py
│   ├── celery.py
│   └── ...
├── sei_automation/         # App principal
│   ├── models.py          # Modelos de dados
│   ├── views.py           # Views/Controllers
│   ├── forms.py           # Formulários
│   ├── tasks.py           # Tarefas Celery
│   ├── admin.py           # Admin interface
│   ├── urls.py            # Rotas
│   └── templates/         # Templates HTML
├── requirements.txt
└── README.md
```

## 📊 Modelos de Dados

### BuscaSEI
Armazena informações sobre cada busca realizada:
- Usuário
- URL do SEI
- Termos de pesquisa
- Datas de filtro
- Status (pendente/em_andamento/concluída/erro)

### DocumentoSEI
Armazena os documentos encontrados:
- Número do processo
- Nome do documento
- Resumo
- Unidade
- Data de inclusão
- Links

### ConfiguracaoSEI
Armazena configurações de acesso:
- Nome da configuração
- Credenciais do SEI
- Órgão

## 🔐 Segurança

- Senhas devem ser armazenadas de forma criptografada
- Use HTTPS em produção
- Configure SECRET_KEY segura em produção
- Nunca commite credenciais no código

## 🛠️ Tecnologias Utilizadas

- **Django 5.2**: Framework web
- **PostgreSQL**: Banco de dados
- **Celery**: Processamento assíncrono
- **Redis**: Broker de mensagens
- **Selenium**: Automação web
- **Pandas**: Manipulação de dados
- **Bootstrap 5**: Interface UI

## 📝 Desenvolvimento

### Executar testes

```bash
python manage.py test
```

### Criar migrações após mudanças nos models

```bash
python manage.py makemigrations
python manage.py migrate
```

### Acessar o Admin

Acesse http://localhost:8000/admin com as credenciais de superusuário.

## 🤝 Contribuindo

Contribuições são bem-vindas! Por favor:
1. Fork o projeto
2. Crie uma branch para sua feature
3. Commit suas mudanças
4. Push para a branch
5. Abra um Pull Request

## 📄 Licença

Este projeto está sob a licença especificada no arquivo LICENSE.

## 👥 Autores

Lab-dados-SEGES

## 🐛 Problemas Conhecidos

- A senha SEI ainda não está sendo criptografada (TODO)
- Link shortener precisa de API key configurada

## 📞 Suporte

Para reportar problemas ou sugerir melhorias, abra uma issue no GitHub.
