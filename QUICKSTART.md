# Guia Rápido de Instalação - SEI Automation Platform

## 🚀 Início Rápido com Docker (Recomendado)

### Pré-requisitos
- Docker
- Docker Compose

### Passos

1. **Clone o repositório**
```bash
git clone https://github.com/lab-dados-seges/sei_aspar.git
cd sei_aspar
```

2. **Configurar variáveis de ambiente**
```bash
cp .env.example .env
# Edite o arquivo .env com suas configurações
```

3. **Iniciar os serviços**
```bash
docker-compose up -d
```

4. **Executar migrações**
```bash
docker-compose exec web python manage.py migrate
```

5. **Criar superusuário**
```bash
docker-compose exec web python manage.py createsuperuser
```

6. **Acessar a aplicação**
- Web: http://localhost:8000
- Admin: http://localhost:8000/admin

### Comandos úteis
```bash
# Ver logs
docker-compose logs -f web

# Parar serviços
docker-compose down

# Reiniciar serviços
docker-compose restart

# Executar comandos Django
docker-compose exec web python manage.py <comando>
```

---

## 💻 Instalação Manual

### Pré-requisitos
- Python 3.8+
- PostgreSQL 12+
- Redis
- Chrome/Chromium

### Passos

1. **Clone o repositório**
```bash
git clone https://github.com/lab-dados-seges/sei_aspar.git
cd sei_aspar
```

2. **Criar ambiente virtual**
```bash
python -m venv venv
source venv/bin/activate  # Linux/Mac
# ou
venv\Scripts\activate  # Windows
```

3. **Instalar dependências**
```bash
pip install -r requirements.txt
```

4. **Configurar PostgreSQL**
```sql
-- Entre no PostgreSQL como superusuário
CREATE DATABASE sei_database;
CREATE USER sei_user WITH PASSWORD 'sei_password';
ALTER ROLE sei_user SET client_encoding TO 'utf8';
ALTER ROLE sei_user SET default_transaction_isolation TO 'read committed';
ALTER ROLE sei_user SET timezone TO 'America/Sao_Paulo';
GRANT ALL PRIVILEGES ON DATABASE sei_database TO sei_user;
```

5. **Configurar variáveis de ambiente**
```bash
cp .env.example .env
# Edite o arquivo .env
```

6. **Executar migrações**
```bash
python manage.py migrate
```

7. **Criar superusuário**
```bash
python manage.py createsuperuser
```

8. **Coletar arquivos estáticos**
```bash
python manage.py collectstatic
```

9. **Iniciar serviços**

Em terminais separados:

```bash
# Terminal 1 - Redis
redis-server

# Terminal 2 - Celery Worker
celery -A sei_platform worker -l info

# Terminal 3 - Django
python manage.py runserver
```

10. **Acessar a aplicação**
- Web: http://localhost:8000
- Admin: http://localhost:8000/admin

---

## 🔧 Script Automatizado

Use o script de setup:
```bash
chmod +x setup.sh
./setup.sh
```

---

## 📝 Primeiro Uso

1. **Login**: Acesse http://localhost:8000 e faça login
2. **Configurar Acesso SEI**: Vá em "Configurações" e salve suas credenciais
3. **Nova Busca**: Clique em "Nova Busca" e preencha os parâmetros
4. **Aguardar**: A busca será processada em background
5. **Ver Resultados**: Acesse "Minhas Buscas" para ver os resultados
6. **Exportar**: Baixe os dados em Excel

---

## ❓ Solução de Problemas Comuns

### Erro ao conectar no PostgreSQL
```bash
# Verificar se PostgreSQL está rodando
sudo systemctl status postgresql
# ou
docker-compose ps
```

### Erro ao conectar no Redis
```bash
# Verificar se Redis está rodando
redis-cli ping
# Deve retornar: PONG
```

### Erro no Celery
```bash
# Ver logs do Celery
celery -A sei_platform worker -l debug
```

### Erro no Selenium
```bash
# Instalar Chrome
wget https://dl.google.com/linux/direct/google-chrome-stable_current_amd64.deb
sudo dpkg -i google-chrome-stable_current_amd64.deb

# Instalar ChromeDriver
sudo apt-get install chromium-chromedriver
```

---

## 📚 Documentação Completa

- **README.md**: Documentação técnica completa
- **GUIA_USO.md**: Guia detalhado de uso da plataforma
- **Admin Django**: Documentação de modelos e dados

---

## 🆘 Suporte

- Issues: https://github.com/lab-dados-seges/sei_aspar/issues
- Email: suporte@exemplo.com

---

## 📄 Licença

Ver arquivo LICENSE
