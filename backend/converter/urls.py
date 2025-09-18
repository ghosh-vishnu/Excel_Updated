from django.urls import path
from . import views, auth_views

urlpatterns = [
    # Authentication URLs
    path("api/auth/login/", auth_views.login_view, name="login"),
    path("api/auth/logout/", auth_views.logout_view, name="logout"),
    path("api/auth/check/", auth_views.check_auth_view, name="check_auth"),
    
    # Converter URLs
    path("api/upload/", views.upload_files, name="upload_files"),
    path("api/convert/", views.start_convert, name="start_convert"),
    path("api/progress/", views.progress, name="progress"),
    path("api/result/", views.result_file, name="result_file"),
    path("api/reset/", views.reset_job, name="reset_job"),
]
