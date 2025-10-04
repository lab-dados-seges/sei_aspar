from django.urls import path
from . import views

urlpatterns = [
    path('', views.home, name='home'),
    path('nova-busca/', views.nova_busca, name='nova_busca'),
    path('buscas/', views.lista_buscas, name='lista_buscas'),
    path('busca/<int:pk>/', views.busca_detalhe, name='busca_detalhe'),
    path('busca/<int:pk>/status/', views.status_busca, name='status_busca'),
    path('busca/<int:pk>/exportar/', views.exportar_excel, name='exportar_excel'),
    path('configuracoes/', views.configuracoes, name='configuracoes'),
]
