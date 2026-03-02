from django.urls import path
from . import views

urlpatterns = [
    path('', views.home, name='home'),
    path('download-excel/', views.download_excel, name='download_excel'),
    path('extract-one/', views.extract_one, name='extract_one'),
    path('generate-excel/', views.generate_excel, name='generate_excel'),
    path('clear-processed/', views.clear_processed_list, name='clear_processed'),
    path('advanced-analysis/', views.advanced_analysis, name='advanced_analysis'),
    path('advanced-analysis-one/', views.advanced_analysis_one, name='advanced_analysis_one'),
]
