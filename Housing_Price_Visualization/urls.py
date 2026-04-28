"""
URL configuration for Housing_Price_Visualization project.
"""
from django.contrib import admin
from django.urls import path, include

urlpatterns = [
    path('admin/', admin.site.urls),  # Django自带后台（保留）
    path('', include('house_price.urls')),  # 核心：加载你的所有页面路由
]