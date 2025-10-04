#!/bin/bash
# Script de configuração inicial da Plataforma SEI Automation

echo "============================================"
echo "SEI Automation Platform - Setup Script"
echo "============================================"
echo ""

# Cores para output
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

# Verificar Python
echo -e "${YELLOW}[1/8]${NC} Verificando Python..."
if ! command -v python3 &> /dev/null; then
    echo -e "${RED}Erro: Python 3 não encontrado. Por favor, instale Python 3.8+${NC}"
    exit 1
fi
PYTHON_VERSION=$(python3 --version)
echo -e "${GREEN}✓${NC} Python encontrado: $PYTHON_VERSION"

# Verificar PostgreSQL
echo -e "\n${YELLOW}[2/8]${NC} Verificando PostgreSQL..."
if ! command -v psql &> /dev/null; then
    echo -e "${RED}Aviso: PostgreSQL não encontrado. Por favor, instale PostgreSQL 12+${NC}"
else
    PSQL_VERSION=$(psql --version)
    echo -e "${GREEN}✓${NC} PostgreSQL encontrado: $PSQL_VERSION"
fi

# Verificar Redis
echo -e "\n${YELLOW}[3/8]${NC} Verificando Redis..."
if ! command -v redis-server &> /dev/null; then
    echo -e "${RED}Aviso: Redis não encontrado. Por favor, instale Redis${NC}"
else
    echo -e "${GREEN}✓${NC} Redis encontrado"
fi

# Instalar dependências Python
echo -e "\n${YELLOW}[4/8]${NC} Instalando dependências Python..."
pip install -r requirements.txt
if [ $? -eq 0 ]; then
    echo -e "${GREEN}✓${NC} Dependências instaladas com sucesso"
else
    echo -e "${RED}Erro ao instalar dependências${NC}"
    exit 1
fi

# Configurar banco de dados
echo -e "\n${YELLOW}[5/8]${NC} Configuração do banco de dados PostgreSQL"
echo "Por favor, execute os seguintes comandos no PostgreSQL:"
echo ""
echo "  CREATE DATABASE sei_database;"
echo "  CREATE USER sei_user WITH PASSWORD 'sei_password';"
echo "  ALTER ROLE sei_user SET client_encoding TO 'utf8';"
echo "  ALTER ROLE sei_user SET default_transaction_isolation TO 'read committed';"
echo "  ALTER ROLE sei_user SET timezone TO 'America/Sao_Paulo';"
echo "  GRANT ALL PRIVILEGES ON DATABASE sei_database TO sei_user;"
echo ""
read -p "Pressione ENTER quando terminar de configurar o PostgreSQL..."

# Executar migrações
echo -e "\n${YELLOW}[6/8]${NC} Executando migrações do Django..."
python manage.py migrate
if [ $? -eq 0 ]; then
    echo -e "${GREEN}✓${NC} Migrações executadas com sucesso"
else
    echo -e "${RED}Erro ao executar migrações. Verifique a configuração do PostgreSQL${NC}"
    exit 1
fi

# Criar superusuário
echo -e "\n${YELLOW}[7/8]${NC} Criar superusuário do Django"
echo "Por favor, forneça as credenciais para o administrador:"
python manage.py createsuperuser

# Coletar arquivos estáticos
echo -e "\n${YELLOW}[8/8]${NC} Coletando arquivos estáticos..."
python manage.py collectstatic --noinput
echo -e "${GREEN}✓${NC} Arquivos estáticos coletados"

# Instruções finais
echo -e "\n${GREEN}============================================${NC}"
echo -e "${GREEN}Setup concluído com sucesso!${NC}"
echo -e "${GREEN}============================================${NC}"
echo ""
echo "Para iniciar a aplicação, execute os seguintes comandos em terminais separados:"
echo ""
echo -e "${YELLOW}Terminal 1 - Redis:${NC}"
echo "  redis-server"
echo ""
echo -e "${YELLOW}Terminal 2 - Celery Worker:${NC}"
echo "  celery -A sei_platform worker -l info"
echo ""
echo -e "${YELLOW}Terminal 3 - Django Server:${NC}"
echo "  python manage.py runserver"
echo ""
echo "Depois acesse: http://localhost:8000"
echo ""
echo "Para acessar o admin: http://localhost:8000/admin"
echo ""
echo -e "${GREEN}Bom uso!${NC}"
