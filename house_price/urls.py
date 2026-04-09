# house_price/urls.py
from django.urls import path
from . import views

urlpatterns = [
    path('login/', views.user_login, name='login'),  # 登录页
    path('index/', views.index, name='index'),  # 普通用户首页
    path('logout/', views.user_logout, name='logout'),  # 退出登录
    path('', views.user_login, name='login'),  # 登录页
    path('register/', views.register, name='register'), # 注册页
    # 房价详情页
    path('house-detail/', views.house_detail, name='house_detail'),
    # 数据分析页
    path('data-analysis/', views.data_analysis, name='data_analysis'),


    # 后台管理路由
    path("admin_index/", views.admin_index, name="admin_index"),          # 后台首页
    path("admin_user_manage/", views.admin_user_manage, name="admin_user_manage"),  # 用户管理
    path("admin_user_edit_ajax/", views.admin_user_edit_ajax, name="admin_user_edit_ajax"),
    path("admin_data_manage/", views.admin_data_manage, name="admin_data_manage"),  # 数据管理
    path("admin_user_toggle_status_ajax/", views.admin_user_toggle_status_ajax, name="admin_user_toggle_status_ajax"),
    path("admin_user_delete_ajax/", views.admin_user_delete_ajax, name="admin_user_delete_ajax"),  # 删除路由
    path("admin_user_add_ajax/", views.admin_user_add_ajax, name="admin_user_add_ajax"),  # 管理员路由
    path("data_upload_ajax/", views.data_upload_ajax, name="data_upload_ajax"),  # 管理员路由
    path("data_edit_ajax/", views.data_edit_ajax, name="data_edit_ajax"),#数据修改
    path("data_delete_ajax/", views.data_delete_ajax, name="data_delete_ajax"),#数据删除
    path("data_export/", views.data_export, name="data_export"),#数据导入
    path("data_download_template/", views.data_download_template, name="data_download_template"),#数据模板下载
]