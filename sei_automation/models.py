from django.db import models
from django.contrib.auth.models import User


class BuscaSEI(models.Model):
    """Model para armazenar as buscas realizadas no SEI"""
    STATUS_CHOICES = [
        ('pendente', 'Pendente'),
        ('em_andamento', 'Em Andamento'),
        ('concluida', 'Concluída'),
        ('erro', 'Erro'),
    ]
    
    usuario = models.ForeignKey(User, on_delete=models.CASCADE, related_name='buscas')
    url_sei = models.URLField(verbose_name='URL do SEI')
    usuario_sei = models.CharField(max_length=100, verbose_name='Usuário SEI')
    orgao = models.CharField(max_length=100, verbose_name='Órgão')
    termos_pesquisa = models.TextField(verbose_name='Termos de Pesquisa')
    data_inicio = models.DateField(verbose_name='Data Início', null=True, blank=True)
    data_fim = models.DateField(verbose_name='Data Fim', null=True, blank=True)
    status = models.CharField(max_length=20, choices=STATUS_CHOICES, default='pendente')
    data_criacao = models.DateTimeField(auto_now_add=True)
    data_atualizacao = models.DateTimeField(auto_now=True)
    mensagem_erro = models.TextField(blank=True, null=True)
    
    class Meta:
        verbose_name = 'Busca SEI'
        verbose_name_plural = 'Buscas SEI'
        ordering = ['-data_criacao']
    
    def __str__(self):
        return f"Busca {self.id} - {self.status} - {self.data_criacao.strftime('%d/%m/%Y %H:%M')}"


class DocumentoSEI(models.Model):
    """Model para armazenar os documentos encontrados nas buscas"""
    busca = models.ForeignKey(BuscaSEI, on_delete=models.CASCADE, related_name='documentos')
    numero_processo = models.CharField(max_length=100, verbose_name='Número do Processo')
    nome_documento = models.CharField(max_length=255, verbose_name='Nome do Documento')
    resumo = models.TextField(verbose_name='Resumo', blank=True)
    unidade = models.CharField(max_length=200, verbose_name='Unidade', blank=True)
    usuario_inclusao = models.CharField(max_length=100, verbose_name='Usuário de Inclusão', blank=True)
    data_inclusao = models.DateTimeField(verbose_name='Data de Inclusão', null=True, blank=True)
    link_original = models.URLField(verbose_name='Link Original', max_length=500)
    link_curto = models.URLField(verbose_name='Link Curto', max_length=500, blank=True, null=True)
    arquivo_pdf = models.FileField(upload_to='documentos_sei/', blank=True, null=True)
    data_criacao = models.DateTimeField(auto_now_add=True)
    
    class Meta:
        verbose_name = 'Documento SEI'
        verbose_name_plural = 'Documentos SEI'
        ordering = ['-data_criacao']
        indexes = [
            models.Index(fields=['numero_processo']),
            models.Index(fields=['nome_documento']),
        ]
    
    def __str__(self):
        return f"{self.numero_processo} - {self.nome_documento}"


class ConfiguracaoSEI(models.Model):
    """Model para armazenar configurações globais do SEI"""
    usuario = models.ForeignKey(User, on_delete=models.CASCADE)
    nome = models.CharField(max_length=100, verbose_name='Nome da Configuração')
    url_sei = models.URLField(verbose_name='URL do SEI', default='https://sei.economia.gov.br/')
    usuario_sei = models.CharField(max_length=100, verbose_name='Usuário SEI')
    senha_sei_encriptada = models.CharField(max_length=255, verbose_name='Senha SEI (Encriptada)')
    orgao = models.CharField(max_length=100, verbose_name='Órgão')
    ativo = models.BooleanField(default=True)
    data_criacao = models.DateTimeField(auto_now_add=True)
    data_atualizacao = models.DateTimeField(auto_now=True)
    
    class Meta:
        verbose_name = 'Configuração SEI'
        verbose_name_plural = 'Configurações SEI'
        ordering = ['-data_criacao']
    
    def __str__(self):
        return f"{self.nome} - {self.orgao}"

