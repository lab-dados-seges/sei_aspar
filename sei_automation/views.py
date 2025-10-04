from django.shortcuts import render, redirect, get_object_or_404
from django.contrib.auth.decorators import login_required
from django.contrib import messages
from django.http import JsonResponse, HttpResponse
from django.core.paginator import Paginator
from .models import BuscaSEI, DocumentoSEI, ConfiguracaoSEI
from .forms import BuscaSEIForm, ConfiguracaoSEIForm
from .tasks import executar_busca_sei
import pandas as pd
from io import BytesIO


@login_required
def home(request):
    """View principal - dashboard"""
    buscas_recentes = BuscaSEI.objects.filter(usuario=request.user)[:5]
    total_buscas = BuscaSEI.objects.filter(usuario=request.user).count()
    total_documentos = DocumentoSEI.objects.filter(busca__usuario=request.user).count()
    
    context = {
        'buscas_recentes': buscas_recentes,
        'total_buscas': total_buscas,
        'total_documentos': total_documentos,
    }
    return render(request, 'sei_automation/home.html', context)


@login_required
def nova_busca(request):
    """View para criar uma nova busca"""
    if request.method == 'POST':
        form = BuscaSEIForm(request.POST)
        if form.is_valid():
            busca = form.save(commit=False)
            busca.usuario = request.user
            busca.status = 'pendente'
            busca.save()
            
            # Enfileirar tarefa Celery
            senha = form.cleaned_data.get('senha_sei', '')
            executar_busca_sei.delay(busca.id, senha)
            
            messages.success(request, 'Busca criada com sucesso! A execução foi iniciada.')
            return redirect('busca_detalhe', pk=busca.id)
    else:
        # Pré-preencher com configuração ativa do usuário
        config = ConfiguracaoSEI.objects.filter(usuario=request.user, ativo=True).first()
        initial = {}
        if config:
            initial = {
                'url_sei': config.url_sei,
                'usuario_sei': config.usuario_sei,
                'orgao': config.orgao,
            }
        form = BuscaSEIForm(initial=initial)
    
    return render(request, 'sei_automation/nova_busca.html', {'form': form})


@login_required
def busca_detalhe(request, pk):
    """View para ver detalhes de uma busca"""
    busca = get_object_or_404(BuscaSEI, pk=pk, usuario=request.user)
    documentos = busca.documentos.all()
    
    # Paginação
    paginator = Paginator(documentos, 25)
    page_number = request.GET.get('page')
    page_obj = paginator.get_page(page_number)
    
    context = {
        'busca': busca,
        'page_obj': page_obj,
        'total_documentos': documentos.count(),
    }
    return render(request, 'sei_automation/busca_detalhe.html', context)


@login_required
def lista_buscas(request):
    """View para listar todas as buscas"""
    buscas = BuscaSEI.objects.filter(usuario=request.user)
    
    # Filtros
    status = request.GET.get('status')
    if status:
        buscas = buscas.filter(status=status)
    
    # Paginação
    paginator = Paginator(buscas, 20)
    page_number = request.GET.get('page')
    page_obj = paginator.get_page(page_number)
    
    context = {
        'page_obj': page_obj,
        'status_choices': BuscaSEI.STATUS_CHOICES,
    }
    return render(request, 'sei_automation/lista_buscas.html', context)


@login_required
def exportar_excel(request, pk):
    """Exportar documentos de uma busca para Excel"""
    busca = get_object_or_404(BuscaSEI, pk=pk, usuario=request.user)
    documentos = busca.documentos.all()
    
    # Criar DataFrame
    data = []
    for doc in documentos:
        data.append({
            'Número do Processo': doc.numero_processo,
            'Documento': doc.nome_documento,
            'Resumo': doc.resumo,
            'Unidade': doc.unidade,
            'Usuário': doc.usuario_inclusao,
            'Data de Inclusão': doc.data_inclusao,
            'Link': doc.link_curto or doc.link_original,
        })
    
    df = pd.DataFrame(data)
    
    # Gerar Excel
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Documentos SEI')
    
    output.seek(0)
    response = HttpResponse(
        output.read(),
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
    response['Content-Disposition'] = f'attachment; filename=busca_{pk}_documentos.xlsx'
    return response


@login_required
def configuracoes(request):
    """View para gerenciar configurações"""
    configs = ConfiguracaoSEI.objects.filter(usuario=request.user)
    
    if request.method == 'POST':
        form = ConfiguracaoSEIForm(request.POST)
        if form.is_valid():
            config = form.save(commit=False)
            config.usuario = request.user
            # TODO: Encriptar senha antes de salvar
            config.senha_sei_encriptada = form.cleaned_data['senha_sei']
            config.save()
            messages.success(request, 'Configuração salva com sucesso!')
            return redirect('configuracoes')
    else:
        form = ConfiguracaoSEIForm()
    
    context = {
        'form': form,
        'configs': configs,
    }
    return render(request, 'sei_automation/configuracoes.html', context)


@login_required
def status_busca(request, pk):
    """API endpoint para verificar status da busca"""
    busca = get_object_or_404(BuscaSEI, pk=pk, usuario=request.user)
    return JsonResponse({
        'status': busca.status,
        'total_documentos': busca.documentos.count(),
        'mensagem_erro': busca.mensagem_erro,
    })

