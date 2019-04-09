from django.contrib import admin
from django.urls import path, include

app_name = 'addin_all_sheets'

urlpatterns = [
    path('admin/', admin.site.urls),
    path('addin_all_sheets/', include('addin_all_sheets.urls', namespace='marian')),
]