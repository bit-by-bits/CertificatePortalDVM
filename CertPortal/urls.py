from django.urls import path
from .views import cert_portal, index
app_name = 'CertPortal'

urlpatterns = [
    path('', index, name='portal'),
    path('download', cert_portal, name='download'),
]
