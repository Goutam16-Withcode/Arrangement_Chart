# seating/urls.py

from django.urls import path
from . import views
from .views import seating_view, run_script

urlpatterns = [
    path('', views.seating_view, name='seating-home'),  
    path('run-script/', run_script, name='run_script'),  

]
