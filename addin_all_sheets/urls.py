from django.urls import pathfrom . import views

app_name = 'addin_all_sheets'

urlpatterns = [
     path('', views.index, name='index'),
]