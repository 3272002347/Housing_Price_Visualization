from django.contrib import admin
from .models import HousePriceData,User


# 注册房价数据模型
@admin.register(HousePriceData)
class HousePriceDataAdmin(admin.ModelAdmin):
    list_display = ('city', 'area', 'house_type', 'area_size', 'total_price', 'create_time')
    search_fields = ('city', 'area', 'title')
    list_filter = ('city', 'house_type', 'decoration')
    readonly_fields = ('create_time',)

# 注册用户模型
@admin.register(User)
class UserAdmin(admin.ModelAdmin):
    list_display = ('username', 'role', 'is_active', 'date_joined')
    list_filter = ('role', 'is_active')
    search_fields = ('username',)