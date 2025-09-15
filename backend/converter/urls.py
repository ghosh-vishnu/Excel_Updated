from django.urls import path
from . import views

urlpatterns = [
    path("api/upload/", views.upload_files, name="upload_files"),
    path("api/convert/", views.start_convert, name="start_convert"),
    path("api/progress/", views.progress, name="progress"),
    path("api/result/", views.result_file, name="result_file"),
    path("api/reset/", views.reset_job, name="reset_job"),
]
