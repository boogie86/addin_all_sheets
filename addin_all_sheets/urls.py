from django.urls import path

app_name = 'addin_all_sheets'

urlpatterns = [
     path('', views.index, name='index'),
]