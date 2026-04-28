from django.db import models
from django.contrib.auth.models import AbstractUser

class HousePriceData(models.Model):
    city = models.CharField(max_length=50, verbose_name="城市")
    title = models.CharField(max_length=200, verbose_name="房源标题")
    area = models.CharField(max_length=100, verbose_name="区域/小区")
    house_type = models.CharField(max_length=50, verbose_name="户型")
    area_size = models.DecimalField(max_digits=10, decimal_places=2, verbose_name="面积(㎡)")
    orientation = models.CharField(max_length=50, blank=True, default="", verbose_name="朝向")
    decoration = models.CharField(max_length=20, blank=True, default="", verbose_name="装修情况")
    floor_info = models.CharField(max_length=100, blank=True, default="", verbose_name="楼层信息")
    house_age = models.IntegerField(null=True, blank=True, verbose_name="房龄(年)")
    structure_type = models.CharField(max_length=50, blank=True, default="", verbose_name="建筑结构")
    total_price = models.DecimalField(max_digits=10, decimal_places=2, verbose_name="总价(万元)")
    unit_price = models.DecimalField(max_digits=10, decimal_places=2, verbose_name="单价(元/平)")
    create_time = models.DateTimeField(auto_now_add=True, verbose_name="录入时间")
    class Meta:
        db_table = "house_price_data"
        verbose_name = "房价数据"
        verbose_name_plural = "房价数据"
        indexes = [
            models.Index(fields=["city"]),
            models.Index(fields=["area"]),
        ]
    def __str__(self):
        return f"{self.city}-{self.area}-{self.title}"

# 扩展Django自带用户模型，解决反向关联冲突
class User(AbstractUser):
    # 角色选项：1-普通用户，2-管理员
    ROLE_CHOICES = ((1, "普通用户"), (2, "管理员"))
    role = models.IntegerField(choices=ROLE_CHOICES, default=0, verbose_name="用户角色")

    # 关键：给冲突字段指定唯一的related_name
    groups = models.ManyToManyField(
        'auth.Group',
        verbose_name='groups',
        blank=True,
        related_name='house_price_user_groups',  # 唯一命名，避免冲突
        related_query_name='house_price_user',
    )
    user_permissions = models.ManyToManyField(
        'auth.Permission',
        verbose_name='user permissions',
        blank=True,
        related_name='house_price_user_permissions',  # 唯一命名，避免冲突
        related_query_name='house_price_user',
    )

    class Meta:
        db_table = "user"
        verbose_name = "用户"
        verbose_name_plural = "用户"

    def __str__(self):
        return self.username