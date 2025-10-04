from django import forms
from .models import BuscaSEI, ConfiguracaoSEI


class BuscaSEIForm(forms.ModelForm):
    senha_sei = forms.CharField(
        widget=forms.PasswordInput(attrs={'class': 'form-control'}),
        label='Senha SEI',
        required=False,
        help_text='Deixe em branco para usar configuração salva'
    )
    
    class Meta:
        model = BuscaSEI
        fields = ['url_sei', 'usuario_sei', 'orgao', 'termos_pesquisa', 'data_inicio', 'data_fim']
        widgets = {
            'url_sei': forms.URLInput(attrs={'class': 'form-control'}),
            'usuario_sei': forms.TextInput(attrs={'class': 'form-control'}),
            'orgao': forms.TextInput(attrs={'class': 'form-control'}),
            'termos_pesquisa': forms.Textarea(attrs={'class': 'form-control', 'rows': 4}),
            'data_inicio': forms.DateInput(attrs={'class': 'form-control', 'type': 'date'}),
            'data_fim': forms.DateInput(attrs={'class': 'form-control', 'type': 'date'}),
        }


class ConfiguracaoSEIForm(forms.ModelForm):
    senha_sei = forms.CharField(
        widget=forms.PasswordInput(attrs={'class': 'form-control'}),
        label='Senha SEI'
    )
    
    class Meta:
        model = ConfiguracaoSEI
        fields = ['nome', 'url_sei', 'usuario_sei', 'orgao', 'ativo']
        widgets = {
            'nome': forms.TextInput(attrs={'class': 'form-control'}),
            'url_sei': forms.URLInput(attrs={'class': 'form-control'}),
            'usuario_sei': forms.TextInput(attrs={'class': 'form-control'}),
            'orgao': forms.TextInput(attrs={'class': 'form-control'}),
            'ativo': forms.CheckboxInput(attrs={'class': 'form-check-input'}),
        }
