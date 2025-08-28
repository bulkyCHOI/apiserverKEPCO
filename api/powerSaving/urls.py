from django.urls import path
from .api import api  # api.py에서 정의한 api 객체 임포트

urlpatterns = [
    path("api/", api.urls),  # /api/ 경로로 API 엔드포인트 연결
]