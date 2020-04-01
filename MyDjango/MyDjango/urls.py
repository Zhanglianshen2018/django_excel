from django.urls import path
from .views import *

urlpatterns = [
    path('', index),
    path('upload/',upload),
    path('deal_excel/',deal_excel),

]
