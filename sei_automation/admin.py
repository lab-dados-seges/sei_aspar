from django.contrib import admin
from .models import BuscaSEI, DocumentoSEI, ConfiguracaoSEI


@admin.register(BuscaSEI)
class BuscaSEIAdmin(admin.ModelAdmin):
    list_display = ['id', 'usuario', 'orgao', 'status', 'data_criacao', 'data_atualizacao']
    list_filter = ['status', 'orgao', 'data_criacao']
    search_fields = ['usuario__username', 'orgao', 'termos_pesquisa']
    readonly_fields = ['data_criacao', 'data_atualizacao']
    
    fieldsets = (
        ('Informações Básicas', {
            'fields': ('usuario', 'url_sei', 'usuario_sei', 'orgao')
        }),
        ('Parâmetros de Busca', {
            'fields': ('termos_pesquisa', 'data_inicio', 'data_fim')
        }),
        ('Status', {
            'fields': ('status', 'mensagem_erro')
        }),
        ('Datas', {
            'fields': ('data_criacao', 'data_atualizacao')
        }),
    )


@admin.register(DocumentoSEI)
class DocumentoSEIAdmin(admin.ModelAdmin):
    list_display = ['numero_processo', 'nome_documento', 'unidade', 'data_inclusao', 'busca']
    list_filter = ['unidade', 'data_inclusao', 'data_criacao']
    search_fields = ['numero_processo', 'nome_documento', 'resumo', 'unidade']
    readonly_fields = ['data_criacao']
    
    fieldsets = (
        ('Informações do Documento', {
            'fields': ('busca', 'numero_processo', 'nome_documento', 'resumo')
        }),
        ('Metadata', {
            'fields': ('unidade', 'usuario_inclusao', 'data_inclusao')
        }),
        ('Links e Arquivos', {
            'fields': ('link_original', 'link_curto', 'arquivo_pdf')
        }),
        ('Datas', {
            'fields': ('data_criacao',)
        }),
    )


@admin.register(ConfiguracaoSEI)
class ConfiguracaoSEIAdmin(admin.ModelAdmin):
    list_display = ['nome', 'usuario', 'orgao', 'ativo', 'data_criacao']
    list_filter = ['ativo', 'orgao', 'data_criacao']
    search_fields = ['nome', 'orgao', 'usuario__username']
    readonly_fields = ['data_criacao', 'data_atualizacao']

