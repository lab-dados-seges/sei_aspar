# Guia de Uso da Plataforma SEI Automation

## Visão Geral

A Plataforma SEI Automation é uma aplicação web desenvolvida em Django que automatiza o processo de busca e coleta de documentos do Sistema Eletrônico de Informações (SEI).

## Funcionalidades Principais

### 1. Dashboard
- Visualização de estatísticas gerais
- Total de buscas realizadas
- Total de documentos coletados
- Buscas recentes com status

### 2. Nova Busca
Permite criar uma nova busca com os seguintes parâmetros:
- **URL do SEI**: Endereço do sistema SEI (ex: https://sei.economia.gov.br/)
- **Usuário SEI**: Seu login no sistema
- **Senha SEI**: Senha de acesso (não armazenada se usar configuração salva)
- **Órgão**: Sigla do órgão (ex: MGI, INSS, etc.)
- **Termos de Pesquisa**: Palavras-chave para busca (ex: "Projeto Lei" ou "PL")
- **Datas**: Período de busca (opcional)

### 3. Acompanhamento de Buscas
- Status em tempo real (Pendente, Em Andamento, Concluída, Erro)
- Atualização automática durante execução
- Visualização de erros se houver

### 4. Visualização de Resultados
- Lista paginada de documentos encontrados
- Informações detalhadas de cada documento:
  - Número do Processo
  - Nome do Documento
  - Unidade responsável
  - Data de Inclusão
  - Link para acesso direto

### 5. Exportação de Dados
- Download em formato Excel (.xlsx)
- Todas as informações dos documentos incluídas
- Organização em colunas para fácil análise

### 6. Configurações
- Salvar credenciais de acesso frequentes
- Múltiplas configurações por usuário
- Uso automático de configuração ativa

## Fluxo de Trabalho Típico

1. **Login**: Acesse a plataforma com suas credenciais Django
2. **Configurar Acesso**: (Opcional) Salve suas credenciais do SEI em Configurações
3. **Criar Busca**: Clique em "Nova Busca" e preencha os parâmetros
4. **Aguardar Processamento**: A busca é executada em background
5. **Visualizar Resultados**: Acesse os detalhes da busca para ver documentos
6. **Exportar**: Baixe os resultados em Excel para análise

## Tecnologias Utilizadas

### Backend
- **Django 5.2**: Framework web principal
- **PostgreSQL**: Banco de dados relacional
- **Celery**: Processamento assíncrono de tarefas
- **Redis**: Broker de mensagens para Celery

### Frontend
- **Bootstrap 5**: Framework CSS responsivo
- **Bootstrap Icons**: Ícones vetoriais
- **JavaScript**: Atualização dinâmica de status

### Automação
- **Selenium**: Navegação automatizada no SEI
- **Pandas**: Processamento de dados
- **XlsxWriter**: Geração de arquivos Excel

## Estrutura de Dados

### BuscaSEI
Cada busca armazena:
- Informações de acesso (URL, usuário, órgão)
- Parâmetros de busca (termos, datas)
- Status da execução
- Mensagens de erro (se houver)
- Relação com o usuário que criou

### DocumentoSEI
Cada documento contém:
- Dados do processo
- Informações do documento
- Metadados (unidade, usuário, data)
- Links de acesso
- Arquivo PDF (se baixado)
- Relação com a busca

### ConfiguracaoSEI
Configurações salvas incluem:
- Nome identificador
- Credenciais de acesso
- Flag de ativa/inativa
- Relação com o usuário

## Segurança

- Autenticação obrigatória para acesso
- Senhas armazenadas de forma criptografada
- Isolamento de dados por usuário
- Proteção CSRF ativada
- Sessões seguras

## Processamento Assíncrono

As buscas são executadas em background através do Celery:
1. Usuário submete a busca
2. Tarefa é enfileirada no Redis
3. Worker Celery processa a busca
4. Status é atualizado em tempo real
5. Usuário é notificado ao término

## Limitações Conhecidas

1. **Dependência do SEI**: A automação depende da estrutura HTML do SEI
2. **Selenium**: Requer Chrome/Chromium instalado
3. **Performance**: Buscas grandes podem demorar
4. **Link Shortener**: Precisa de API key configurada (TODO)
5. **Criptografia de Senha**: Implementação básica (melhorar em produção)

## Manutenção

### Logs
- Django: logs em console e arquivo
- Celery: logs de processamento de tarefas
- Selenium: erros de navegação

### Backup
Recomenda-se backup regular do PostgreSQL:
```bash
pg_dump sei_database > backup_sei_$(date +%Y%m%d).sql
```

### Monitoramento
- Admin do Django: /admin/
- Flower (Celery): Adicionar se necessário
- Logs do sistema: verificar periodicamente

## Solução de Problemas

### Busca não inicia
- Verificar se Celery está rodando
- Verificar se Redis está ativo
- Verificar credenciais do SEI

### Erro de autenticação
- Verificar URL do SEI
- Confirmar usuário e senha
- Verificar seletor de órgão

### Documentos não aparecem
- Verificar termos de pesquisa
- Confirmar período de datas
- Verificar estrutura HTML do SEI

### Erro no Selenium
- Verificar instalação do Chrome
- Verificar ChromeDriver compatível
- Verificar modo headless

## Próximos Passos

Melhorias planejadas:
- [ ] Criptografia robusta de senhas
- [ ] Agendamento de buscas periódicas
- [ ] Notificações por email
- [ ] Download automático de PDFs
- [ ] API REST para integração
- [ ] Dashboard com gráficos
- [ ] Filtros avançados de busca
- [ ] Histórico de versões de documentos

## Suporte

Para dúvidas ou problemas:
1. Consulte a documentação técnica (README.md)
2. Verifique os logs do sistema
3. Abra uma issue no GitHub
4. Entre em contato com a equipe de desenvolvimento

## Changelog

### Versão 1.0.0 (2025-10-04)
- Implementação inicial da plataforma Django
- Migração de código Streamlit para Django
- Integração com PostgreSQL
- Processamento assíncrono com Celery
- Interface web responsiva
- Exportação para Excel
- Sistema de autenticação
