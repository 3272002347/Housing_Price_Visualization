from datetime import datetime,timedelta
import urllib.parse

import json
from decimal import Decimal

import pandas as pd
from django.contrib.auth.decorators import login_required
from django.contrib.auth.hashers import make_password
from django.core.paginator import Paginator, PageNotAnInteger, EmptyPage
from django.db import transaction,models
from django.db.models import Avg, Max, Min, Count, F, ExpressionWrapper, DecimalField, Q,Case,When
from django.shortcuts import render, redirect, get_object_or_404
from django.contrib.auth import authenticate, login, logout
from django.contrib import messages
from django.http import HttpResponse,JsonResponse
from django.utils import timezone
from django.utils.safestring import mark_safe
from openpyxl import Workbook
from django.db.models.functions import ExtractYear, ExtractMonth, TruncMonth,Concat
from .models import User,HousePriceData
from .common import get_valid_chart_config  # 导入公共配置
from django.template.loader import render_to_string
from django.db.models import F, ExpressionWrapper, DecimalField, CharField, Value as V
from .utils import clean_area, clean_house_age, clean_total_price

try:
    # Django <4.0 用pytz
    import pytz
    shanghai_tz = pytz.timezone('Asia/Shanghai')
except ImportError:
    # Django >=4.0 用标准库zoneinfo
    from zoneinfo import ZoneInfo
    shanghai_tz = ZoneInfo('Asia/Shanghai')

class DecimalEncoder(json.JSONEncoder):
    def default(self, obj):
        if isinstance(obj, Decimal):
            # 把Decimal转为浮点数（前端ECharts需要数值类型）
            return float(obj)
        return super().default(obj)


def index(request):
    """首页 - 核心概览+趋势监控"""
    if not request.user.is_authenticated:
        return redirect('/login/')
    # 1. 读取图表配置（核心：动态获取配置）
    mode = request.GET.get("mode", "detailed")
    chart_config, current_mode = get_valid_chart_config(mode)
    trend_month = 6 if current_mode == 'simple' else 24
    # 2. 基础查询（保留有效数据）
    end_time = timezone.now()
    start_time = end_time - timedelta(days=trend_month*30)
    start_time = start_time.replace(hour=0, minute=0, second=0, microsecond=0)
    base_qs = HousePriceData.objects.filter(
        area_size__gt=0, total_price__gt=0,
        create_time__isnull=False, create_time__gte=start_time
    ).annotate(
        unit_price=ExpressionWrapper(
            (F('total_price') * 10000) / F('area_size'),
            output_field=DecimalField(max_digits=12, decimal_places=2)
        )
    )
    # 2. 核心指标
    total_houses = base_qs.count()  # 改用base_qs.count()，排除无效数据
    avg_unit_price = base_qs.aggregate(avg=Avg('unit_price'))['avg'] or 0
    avg_total_price = base_qs.aggregate(avg=Avg('total_price'))['avg'] or 0

    # 均价最高城市
    top_city_qs = base_qs.values('city').annotate(
        avg_price=Avg('unit_price')
    ).order_by('-avg_price')
    top_city = top_city_qs.first() or {}

    trend_data = []
    try:
        raw_data = base_qs.values('create_time', 'unit_price')
        # 手动分组：按年月统计
        month_group = {}
        for item in raw_data:
            utc_time = item['create_time']
            local_time = utc_time.astimezone(shanghai_tz)
            year = local_time.year
            month = local_time.month
            key = f"{year}-{month:02d}"  # 如2026-04
            # 累加数据
            if key not in month_group:
                month_group[key] = {'sum_price': 0, 'count': 0}
            month_group[key]['sum_price'] += float(item['unit_price'])
            month_group[key]['count'] += 1
        # 生成趋势数据并按配置的月份数截取（取最后N个月）
        sorted_keys = sorted(month_group.keys())
        recent_keys = sorted_keys[-trend_month:] if sorted_keys else []  # 动态截取
        # 步骤2：计算均价并生成trend_data
        for key in recent_keys:
            sum_price = month_group[key]['sum_price']
            count = month_group[key]['count']
            avg_price = round(sum_price / count, 2) if count > 0 else 0
            trend_data.append({
                'month': key,
                'avg_price': avg_price,
                'count': count
            })
    except Exception as e:
        print(f"趋势数据错误：{str(e)}")

    # 4. 城市TOP5
    city_top5_list = []
    try:
        city_top5_qs = base_qs.values('city').annotate(
            avg_price=Avg('unit_price'),
            count=Count('id')
        ).filter(count__gt=0, city__isnull=False).order_by('-avg_price')[:5]
        city_top5_list = list(city_top5_qs)
    except Exception as e:
        print(f"城市TOP5错误：{str(e)}")
    chart_config_js = chart_config.copy()
    # 把Python的True/False转成JS的true/false
    chart_config_js['show_label'] = True if chart_config_js['show_label'] else False
    chart_config_js['show_legend'] = True if chart_config_js['show_legend'] else False
    # 5. 上下文
    context = {
        'user': request.user,
        'total_houses': total_houses,
        'avg_unit_price': round(float(avg_unit_price), 2) if avg_unit_price else 0,
        'avg_total_price': round(float(avg_total_price), 2) if avg_total_price else 0,
        'top_city': top_city.get('city', '无'),
        'top_city_price': round(float(top_city.get('avg_price', 0)), 2) if top_city.get('avg_price') else 0,
        'trend_data': mark_safe(json.dumps(trend_data, cls=DecimalEncoder)),
        'city_top5': mark_safe(json.dumps([
            {
                'city': item.get('city', ''),
                'price': round(float(item.get('avg_price', 0)), 2) if item.get('avg_price') else 0,
                'count': item.get('count', 0)
            }
            for item in city_top5_list
        ], cls=DecimalEncoder)),
        'chart_config':  mark_safe(json.dumps(chart_config_js, cls=DecimalEncoder)),       # 模式配置（供前端渲染图表）
        'current_mode': current_mode,       # 当前模式（供前端高亮按钮）
        'trend_month': trend_month,

    }
    return render(request, 'index.html', context)


def data_analysis(request):
    """数据分析页 - 多维度图表数据接口+页面渲染（对齐index函数规范）"""
    if not request.user.is_authenticated:
        return redirect('/login/')

    # 1. 模式配置（对齐index逻辑）
    mode = request.GET.get("mode", "detailed")
    chart_config, current_mode = get_valid_chart_config(mode)
    trend_month = chart_config["trend_month"]

    # 2. 筛选参数
    selected_city = request.GET.get('city', 'all')

    # 3. 基础查询（统一时间处理逻辑，对齐index）
    end_time = timezone.now()
    start_time = end_time - timedelta(days=trend_month * 30)
    start_time = start_time.replace(hour=0, minute=0, second=0, microsecond=0)  # 统一时间精度
    base_qs = HousePriceData.objects.filter(
        area_size__gt=0,
        total_price__gt=0,
        create_time__isnull=False,
        create_time__gte=start_time
    ).annotate(
        unit_price=ExpressionWrapper(
            (F('total_price') * 10000) / F('area_size'),
            output_field=DecimalField(max_digits=12, decimal_places=2)
        )
    )

    # 4. 城市筛选
    if selected_city != 'all' and selected_city.strip():
        base_qs = base_qs.filter(city=selected_city.strip())

    # 5. 维度数据查询（异常捕获+数据清洗，对齐index）
    # 5.1 户型均价
    house_type_price = []
    try:
        house_type_price = list(base_qs.values('house_type').annotate(
            avg_price=Avg('unit_price'),
            count=Count('id')
        ).order_by('-avg_price'))
        house_type_price = [item for item in house_type_price if item['house_type'] and item['count'] > 0]
    except Exception as e:
        print(f"户型均价查询错误：{str(e)}")

    # 5.2 面积区间价格
    area_ranges = [(0, 50), (50, 70), (70, 90), (90, 110), (110, 130), (130, 150), (150, 200), (200, float('inf'))]
    area_range_price = []
    try:
        for min_area, max_area in area_ranges:
            if max_area == float('inf'):
                filter_qs = base_qs.filter(area_size__gte=min_area)
                range_name = f'{min_area}㎡以上'
            else:
                filter_qs = base_qs.filter(area_size__gte=min_area, area_size__lt=max_area)
                range_name = f'{min_area}-{max_area}㎡'

            avg_price = filter_qs.aggregate(avg=Avg('unit_price'))['avg'] or 0
            count = filter_qs.count()
            area_range_price.append({
                'range': range_name,
                'avg_price': round(avg_price, 2),
                'count': count
            })
    except Exception as e:
        print(f"面积区间查询错误：{str(e)}")

    # 5.3 装修类型价格
    decoration_price = []
    try:
        decoration_price = list(base_qs.values('decoration').annotate(
            avg_price=Avg('unit_price'),
            count=Count('id')
        ).order_by('-count'))
        decoration_price = [item for item in decoration_price if item['decoration'] and item['count'] > 0]
    except Exception as e:
        print(f"装修类型查询错误：{str(e)}")

    # 5.4 城市TOP10
    city_price_top10 = []
    try:
        if selected_city == 'all':
            city_price_top10 = list(base_qs.values('city').annotate(
                avg_price=Avg('unit_price'),
                count=Count('id')
            ).filter(count__gt=0, city__isnull=False).order_by('-avg_price')[:10])
    except Exception as e:
        print(f"城市TOP10查询错误：{str(e)}")

    # 5.5 单价区间数量
    price_ranges = [(0, 5000), (5000, 10000), (10000, 15000), (15000, 20000), (20000, 30000), (30000, 50000),
                    (50000, float('inf'))]
    price_range_count = []
    try:
        for min_price, max_price in price_ranges:
            if max_price == float('inf'):
                filter_qs = base_qs.filter(unit_price__gte=min_price)
                range_name = f'{min_price}元/㎡以上'
            else:
                filter_qs = base_qs.filter(unit_price__gte=min_price, unit_price__lt=max_price)
                range_name = f'{min_price}-{max_price}元/㎡'

            count = filter_qs.count()
            price_range_count.append({
                'range': range_name,
                'count': count
            })
    except Exception as e:
        print(f"单价区间查询错误：{str(e)}")

    # 6. 城市列表
    city_list = []
    try:
        city_list = list(
            HousePriceData.objects.values_list('city', flat=True)
            .distinct()
            .order_by('city')
        )
        city_list = [city for city in city_list if city]  # 过滤空值
    except Exception as e:
        print(f"城市列表查询错误：{str(e)}")

    # 7. 构造图表数据（统一Decimal格式化）
    chart_data = {
        'house_type': {
            'x': [item['house_type'] for item in house_type_price],
            'y': [round(float(item['avg_price']), 2) for item in house_type_price],
            'count': [item['count'] for item in house_type_price]
        },
        'area_range': {
            'x': [item['range'] for item in area_range_price],
            'y': [round(float(item['avg_price']), 2) for item in area_range_price],
            'count': [item['count'] for item in area_range_price]
        },
        'decoration': {
            'names': [item['decoration'] for item in decoration_price],
            'prices': [round(float(item['avg_price']), 2) for item in decoration_price],
            'counts': [item['count'] for item in decoration_price]
        },
        'city_top10': {
            'x': [item['city'] for item in city_price_top10],
            'y': [round(float(item['avg_price']), 2) for item in city_price_top10]
        },
        'price_range': {
            'x': [item['range'] for item in price_range_count],
            'y': [item['count'] for item in price_range_count]
        }
    }

    # 8. 配置JSON格式化（对齐index，转义布尔值）
    chart_config_js = chart_config.copy()
    chart_config_js['show_label'] = bool(chart_config_js['show_label'])
    chart_config_js['show_legend'] = bool(chart_config_js['show_legend'])

    # 9. 上下文（统一mark_safe+DecimalEncoder）
    context = {
        'user': request.user,
        'city_list': city_list,
        'selected_city': selected_city,
        'chart_data_json': mark_safe(json.dumps(chart_data, cls=DecimalEncoder, ensure_ascii=False)),
        'chart_config': mark_safe(json.dumps(chart_config_js, cls=DecimalEncoder)),  # 转JSON
        'current_mode': current_mode,
        'trend_month': trend_month
    }
    return render(request, 'data_analysis.html', context)
def house_detail(request):
    """房价详情页 - 支持分页、筛选"""
    # 未登录跳转登录页
    if not request.user.is_authenticated:
        return redirect('/login/')

    # 1. 获取筛选参数
    selected_city = request.GET.get('city', 'all')
    selected_house_type = request.GET.get('house_type', 'all')

    # 2. 基础查询（只取有效数据）
    queryset = HousePriceData.objects.filter(
        area_size__gt=0,
        total_price__gt=0,
        create_time__isnull=False
    ).annotate(
        # 计算单价（元/㎡）
        unit_price=ExpressionWrapper(
            (F('total_price') * 10000) / F('area_size'),
            output_field=DecimalField(max_digits=12, decimal_places=2)
        )
    ).order_by('-create_time')  # 按录入时间倒序

    # 3. 应用筛选条件
    if selected_city != 'all' and selected_city.strip():
        queryset = queryset.filter(city=selected_city.strip())
    if selected_house_type != 'all' and selected_house_type.strip():
        queryset = queryset.filter(house_type=selected_house_type.strip())

    # 4. 分页处理（每页10条）
    paginator = Paginator(queryset, 10)
    page = request.GET.get('page', 1)
    try:
        house_list = paginator.page(page)
    except PageNotAnInteger:
        house_list = paginator.page(1)
    except EmptyPage:
        house_list = paginator.page(paginator.num_pages)

    # 5. 获取筛选下拉框数据（城市/户型去重）
    city_list = list(
        HousePriceData.objects.values_list('city', flat=True)
        .distinct()
        .order_by('city')
    )
    house_type_list = list(
        HousePriceData.objects.values_list('house_type', flat=True)
        .distinct()
        .order_by('house_type')
    )

    # 6. 渲染页面
    context = {
        'user': request.user,
        'house_list': house_list,
        'city_list': city_list,
        'house_type_list': house_type_list,
        'selected_city': selected_city,
        'selected_house_type': selected_house_type
    }
    return render(request, 'house_detail.html', context)

#模板下载
@login_required
def data_download_template(request):
    if request.user.role != 2:
        messages.error(request, "仅管理员可访问")
        return redirect('/login/')
    """下载房价数据导入模板"""
    wb = Workbook()
    ws = wb.active
    ws.title = "模板"

    # 表头
    headers = ["标题", "区域", "户型", "面积", "朝向", "装修情况", "楼层", "房龄", "建筑结构", "总价"]
    ws.append(headers)

    # 示例数据
    sample_data = [
        ["时代倾城精装两房", "时代倾城", "2室2厅", "86.28平米", "南", "简装", "高楼层(共5年)", "5年", "板塔结合",
         "62万"],
        ["满五唯一三房", "时代倾城", "3室1厅", "87.34平米", "南", "精装", "低楼层(共5年)", "5年", "板塔结合", "69万"]
    ]
    for row in sample_data:
        ws.append(row)

    response = HttpResponse(content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


    filename = "房价数据批量导入模板.xlsx"
    encoded_filename = urllib.parse.quote(filename)
    response[
        "Content-Disposition"] = f"attachment; filename=\"{encoded_filename}\"; filename*=UTF-8''{encoded_filename}"

    wb.save(response)
    return response


#数据导出
@login_required
def data_export(request):
    if request.user.role != 2:
        messages.error(request, "仅管理员可访问")
        return redirect('/login/')
     # 1. 获取前端传递的筛选参数（和admin_data_manage完全一致）
    keyword = request.GET.get('keyword', '').strip()
    city = request.GET.get('city', '')
    decoration = request.GET.get('decoration', '')

    # 2. 构造筛选条件（和数据管理页面的筛选逻辑完全同步）
    queryset = HousePriceData.objects.all().order_by("id")
    # 关键词筛选：标题/小区名称模糊匹配
    if keyword:
        queryset = queryset.filter(Q(title__icontains=keyword) | Q(area__icontains=keyword))
    # 城市筛选
    if city:
        queryset = queryset.filter(city=city)
    # 装修类型筛选
    if decoration:
        queryset = queryset.filter(decoration=decoration)
        # 3. 生成Excel文件（导出筛选后的全部数据，不是仅当前分页）
    wb = Workbook()
    ws = wb.active
    ws.title = "房价数据导出"

    # 定义导出表头（包含用户关心的所有字段）
    headers = [
        "ID", "城市", "小区/标题", "区域", "户型", "面积(㎡)",
        "朝向", "装修情况", "楼层信息", "房龄", "建筑结构",
        "总价(万元)", "单价(元/㎡)", "录入时间"
    ]
    ws.append(headers)

    # 4. 填充筛选后的数据（计算单价，和前端展示逻辑一致）
    for house in queryset:
        # 计算单价：总价(万元)*10000 / 面积(㎡)，保留2位小数
        unit_price = 0.00
        if house.area_size > 0 and house.total_price > 0:
            unit_price = round((house.total_price * 10000) / house.area_size, 2)
        row_data = [
            house.id,
            house.city,
            house.title,
            house.area,
            house.house_type,
            round(house.area_size, 2),
            house.orientation if house.orientation else "未填写",
            house.decoration if house.decoration else "未填写",
            house.floor_info if house.floor_info else "未填写",
            house.house_age if house.house_age else "未填写",
            house.structure_type if house.structure_type else "未填写",
            round(house.total_price, 2),
            unit_price,
            house.create_time.strftime("%Y-%m-%d %H:%M:%S") if house.create_time else "未填写"
        ]
        ws.append(row_data)
    # 5. 构造下载响应（解决中文文件名+筛选标识）
    # 生成带筛选条件的文件名（方便用户识别）
    filter_suffix = ""
    if keyword:
        filter_suffix += f"_关键词-{keyword}"
    if city:
        filter_suffix += f"_城市-{city}"
    if decoration:
        filter_suffix += f"_装修-{decoration}"
    # 基础文件名：时间戳+筛选条件
    base_filename = f"房价数据导出_{datetime.now().strftime('%Y%m%d_%H%M%S')}{filter_suffix}.xlsx"
    # 中文文件名URL编码（兼容所有浏览器）
    encoded_filename = urllib.parse.quote(base_filename)

    # 设置响应头
    response = HttpResponse(content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    response[
        "Content-Disposition"] = f"attachment; filename=\"{encoded_filename}\"; filename*=UTF-8''{encoded_filename}"

    # 保存Excel到响应流
    wb.save(response)
    return response


#数据删除
@login_required
def data_delete_ajax(request):
    if request.user.role != 2:
        return JsonResponse({'success': False, 'message': '仅管理员可操作'})
    if request.method == 'POST':
        user_id = request.POST.get('id')
        try:
            house = get_object_or_404(HousePriceData, pk=user_id)
            house.delete()
            return JsonResponse({'success': True, 'message': '删除成功'})
        except Exception as e:
            return JsonResponse({'success': False, 'message': f'删除失败：{str(e)}'})
    return JsonResponse({'success': False, 'message': '请求方式错误'})
#数据修改
@login_required
def data_edit_ajax(request):
    if request.user.role != 2:
        return JsonResponse({'success': False, 'message': '仅管理员可操作'})
    if request.method == 'POST':
        try:
            house_id = request.POST.get('id')
            house = get_object_or_404(HousePriceData, pk=house_id)
            house.city = request.POST.get('city')
            house.area = request.POST.get('area')
            house.title = request.POST.get('community')  # 前端传的community对应title
            house.house_type = request.POST.get('house_type')
            house.area_size = request.POST.get('area_size')
            house.total_price = request.POST.get('total_price')
            house.decoration = request.POST.get('decoration')
            house.orientation = request.POST.get('orientation', '')
            house.floor_info = request.POST.get('floor_info', '')
            house.house_age = request.POST.get('house_age', '')
            house.structure_type = request.POST.get('structure_type', '')
            house.save()
            return JsonResponse({'success': True, 'message': '修改成功'})
        except Exception as e:
            return JsonResponse({'success': False, 'message': f'修改失败：{str(e)}'})
    return JsonResponse({'success': False, 'message': '请求方式错误'})


#数据上传
@login_required
def data_upload_ajax(request):
    if request.user.role!=2:
        messages.error(request,"仅管理员可访问")
        return redirect('/login/')
    file=request.FILES.get('excelFile')
    city=request.POST.get('city','').strip()
    if not file:
        return JsonResponse({
            "success": False,
            "message":"请选择要上传的Excel/CSV文件！"
        })
    #格式校验
    # 3. 文件格式校验
    allowed_extensions = [".xlsx", ".xls", ".csv"]
    file_name = file.name.lower()
    file_ext = None
    for ext in allowed_extensions:
        if file_name.endswith(ext):
            file_ext = ext
            break
    if not file_ext:
        return JsonResponse({
            "success": False,
            "message": f"文件格式错误！仅支持：{','.join(allowed_extensions)}"
        })
    try:
        # 根据文件格式读取
        if file_ext in [".xlsx", ".xls"]:
            df = pd.read_excel(file, engine="openpyxl" if file_ext == ".xlsx" else "xlrd")
        else:  # .csv
            df = pd.read_csv(file, encoding="utf-8")  # 中文编码

        # 重置索引，方便遍历行号
        df = df.reset_index(drop=True)
    except Exception as e:
        return JsonResponse({
            "success": False,
            "message": f"文件解析失败！错误：{str(e)}"
        })

        # 5. 校验Excel必要列（必须包含的字段）
    required_columns = ["标题", "区域", "户型", "面积", "总价"]
    missing_cols = [col for col in required_columns if col not in df.columns]
    if missing_cols:
        return JsonResponse({
            "success": False,
            "message": f"文件缺少必要列：{','.join(missing_cols)}，请下载模板后按格式填写！"
        })
    # 6. 数据清洗+批量入库（事务：要么全成功，要么全失败）
    success_count = 0  # 成功条数
    error_count = 0  # 失败条数
    error_details = []  # 失败详情（前10条，避免返回过长）

    try:
        with transaction.atomic():  # 开启数据库事务
            for row_idx, row in df.iterrows():
                # 行号：Excel中实际行号 = 索引+2（表头占1行，索引从0开始）
                excel_row_num = row_idx + 2

                # 清洗各字段
                area_size = clean_area(row["面积"])
                house_age = clean_house_age(row.get("房龄", ""))  # 可选列
                total_price = clean_total_price(row["总价"])

                # 校验核心字段（面积/总价必须有效）
                if not area_size:
                    error_count += 1
                    error_details.append(f"第{excel_row_num}行：面积格式错误（示例：89㎡/120.5平米）")
                    continue
                if not total_price:
                    error_count += 1
                    error_details.append(f"第{excel_row_num}行：总价格式错误（示例：120万/89.5）")
                    continue

                # 入库（可选列用get，避免KeyError）
                HousePriceData.objects.create(
                    city=city,
                    title=str(row["标题"]).strip(),
                    area=str(row["区域"]).strip(),
                    house_type=str(row["户型"]).strip(),
                    area_size=area_size,
                    orientation=str(row.get("朝向", "")).strip(),  # 可选列
                    decoration=str(row.get("装修情况", "")).strip(),  # 可选列
                    floor_info=str(row.get("楼层", "")).strip(),  # 可选列
                    house_age=house_age,
                    structure_type=str(row.get("建筑结构", "")).strip(),  # 可选列
                    total_price=total_price
                )
                success_count += 1
        # 7. 返回上传结果
        result_msg = f"上传完成！成功导入{success_count}条，失败{error_count}条"
        if error_details:
            # 只展示前10条错误，避免返回内容过长
            result_msg += f"\n错误详情：{'; '.join(error_details[:10])}"
            if len(error_details) > 10:
                result_msg += f"\n（共{len(error_details)}条错误，仅展示前10条）"

        return JsonResponse({
            "success": True,
            "message": result_msg,
            "success_count": success_count,
            "error_count": error_count
        })

    except Exception as e:
    # 事务回滚：任何异常都会触发回滚，保证数据不脏
        return JsonResponse({
            "success": False,
            "message": f"数据入库失败！错误：{str(e)}"
        })
#数据管理首页
@login_required
def admin_data_manage(request):
    if request.user.role!=2:
        messages.error(request,"仅管理员可访问")
        return redirect('/login/')
    keyword=request.GET.get('keyword',"").strip()
    city=request.GET.get('city',"")
    decoration=request.GET.get('decoration',"")
    sort_by = request.GET.get('sort', 'id')  # 默认按id排序
    sort_dir = request.GET.get('dir', 'asc')  # 默认升序
    # 允许排序的字段（防止SQL注入）
    allowed_sort_fields = ['area_size', 'unit_price', 'house_age', 'total_price', 'create_time', 'id']
    if sort_by not in allowed_sort_fields:
        sort_by = 'id'
    # 拼接排序方向
    if sort_dir == 'desc':
        order_by = f'-{sort_by}'
    else:
        order_by = sort_by
    # 2. 构造查询条件
    queryset = HousePriceData.objects.all().order_by("id")
    if keyword:
        # 模糊搜索小区名称或区域
        queryset = queryset.filter(
            Q(title__icontains=keyword) | Q(area__icontains=keyword)
        )
    if city:
        queryset = queryset.filter(city=city)
    if decoration:
        queryset = queryset.filter(decoration=decoration)
    allowed_sort_fields = ['area_size', 'house_age', 'total_price', 'create_time', 'id']
    # 校验排序字段合法性
    if sort_by not in allowed_sort_fields and sort_by != 'unit_price':
        sort_by = 'id'

    # 单价排序特殊处理（数据库层面计算）
    if sort_by == 'unit_price':
        queryset = queryset.annotate(
            calculated_unit_price=ExpressionWrapper(
                (F('total_price') * 10000) / F('area_size'),
                output_field=DecimalField(max_digits=10, decimal_places=2)
            )
        )
        #拼接排序方向
        order_by = f'{"-" if sort_dir == "desc" else ""}calculated_unit_price'
    else:
        # 普通字段排序
        order_by = f'{"-" if sort_dir == "desc" else ""}{sort_by}'

    queryset = queryset.order_by(order_by)
    # 3. 分页（每页10条）
    paginator = Paginator(queryset, 10)
    page_num = request.GET.get("page", 1)  # 当前页码
    try:
        page_obj = paginator.get_page(page_num)
    except PageNotAnInteger:
        page_obj = paginator.get_page(1)
    except EmptyPage:
        page_obj = paginator.get_page(paginator.num_pages)
    max_display = 5  # 最多显示5个页码
    current = page_obj.number
    total = paginator.num_pages
    page_range = []
    if total <= max_display:
        page_range = range(1, total + 1)
    else:
        # 显示前2页、当前页、后2页
        start = max(1, current - 2)
        end = min(total, current + 2)
        if start > 1:
            page_range.append(1)
            if start > 2:
                page_range.append('...')
        page_range.extend(range(start, end + 1))
        if end < total:
            if end < total - 1:
                page_range.append('...')
            page_range.append(total)
    house_data_list = []
    for house in page_obj:
        # 计算单价：总价(万元)*10000 / 面积(㎡)，保留2位小数
        unit_price = 0.00  # 默认值
        if house.area_size > 0 and house.total_price > 0:
            unit_price = round((house.total_price * 10000) / house.area_size, 2)

        # 把原对象的属性和计算的单价封装成字典
        house_data = {
            # 保留原有的所有字段
            "id": house.id,
            "city": house.city,
            "area": house.area,
            "title": house.title,
            "house_type": house.house_type,
            "area_size": round(house.area_size, 2),
            "orientation": house.orientation,
            "decoration": house.decoration,
            "floor_info": house.floor_info,
            "house_age": house.house_age,
            "structure_type": house.structure_type,
            "total_price":round(house.total_price, 2),
            "create_time": house.create_time,
            # 新增：后端计算好的单价
            "unit_price": unit_price
        }
        house_data_list.append(house_data)
    # 4. 传递数据到模板
    context = {
        "user": request.user,
        "page_obj": page_obj,  # 分页后的数据
        "keyword": keyword,  # 回显搜索词
        "house_data_list": house_data_list,
        "city_filter": city,  # 回显城市筛选
        "decoration": decoration,  # 回显装修筛选
        # 城市列表（用于筛选下拉框）
        "city_list": sorted(HousePriceData.objects.values_list("city", flat=True).distinct()),
        "page_range": page_range,  # 新增：精简后的页码列表
        "decoration_list":["精装", "简装", "毛坯"],
        "sort_by": sort_by,
        "sort_dir": sort_dir
    }
    return render(request, "admin_data_manage.html", context)
# AJAX版新增管理员
def admin_user_add_ajax(request):
    # 1. 权限验证：仅管理员可操作
    if not request.user.is_authenticated or request.user.role != 2:
        return JsonResponse({'success': False, 'message': '仅管理员可新增管理员！'})

    # 2. 仅处理AJAX的POST请求
    if request.method == 'POST' and request.headers.get('X-Requested-With') == 'XMLHttpRequest':
        username = request.POST.get('username', '').strip()
        password = request.POST.get('password', '')
        password2 = request.POST.get('password2', '')
        email = request.POST.get('email', '').strip()

        try:
            # 3. 后端二次校验
            # 用户名非空
            if not username:
                return JsonResponse({'success': False, 'message': '用户名不能为空！'})
            # 密码长度
            if len(password) < 6:
                return JsonResponse({'success': False, 'message': '密码长度不能少于6位！'})
            # 密码一致性
            if password != password2:
                return JsonResponse({'success': False, 'message': '两次输入的密码不一致！'})
            # 用户名是否已存在
            if User.objects.filter(username=username).exists():
                return JsonResponse({'success': False, 'message': '用户名已存在，请更换！'})

            # 4. 创建管理员用户
            User.objects.create(
                username=username,
                password=make_password(password),  # 密码加密存储（必须！）
                email=email,
                role=2,  # 2=管理员
                is_active=1  # 默认启用
            )

            # 5. 返回成功结果
            return JsonResponse({'success': True, 'message': f'管理员 {username} 创建成功！'})

        except Exception as e:
            return JsonResponse({'success': False, 'message': f'创建失败：{str(e)}'})

    # 非POST/AJAX请求
    return JsonResponse({'success': False, 'message': '请求方式错误！'})
# AJAX版删除用户（弹窗专用）
def admin_user_delete_ajax(request):
    # 1. 权限验证：仅管理员可操作
    if not request.user.is_authenticated or request.user.role != 2:
        return JsonResponse({'success': False, 'message': '仅管理员可删除用户！'})

    # 2. 仅处理AJAX的POST请求
    if request.method == 'POST' and request.headers.get('X-Requested-With') == 'XMLHttpRequest':
        user_id = request.POST.get('user_id')

        try:
            # 3. 查询要删除的用户（不存在则返回404）
            delete_user = get_object_or_404(User, pk=user_id)

            # 4. 禁止删除自己
            if delete_user.id == request.user.id:
                return JsonResponse({'success': False, 'message': '不能删除当前登录的管理员！'})

            # 5. 执行删除
            username = delete_user.username  # 保存用户名用于提示
            delete_user.delete()

            # 6. 返回成功结果
            return JsonResponse({'success': True, 'message': f'用户 {username} 已成功删除！'})

        except Exception as e:
            # 捕获异常（比如用户不存在、数据库错误等）
            return JsonResponse({'success': False, 'message': f'删除失败：{str(e)}'})

    # 非POST/AJAX请求
    return JsonResponse({'success': False, 'message': '请求方式错误！'})
# 切换用户状态（启用/禁用）
def admin_user_toggle_status_ajax(request):
    # 权限验证
    if not request.user.is_authenticated or request.user.role != 2:
        messages.error(request, "仅管理员可访问！")
        return redirect('/login/')
    if request.method == 'POST' and request.headers.get('X-Requested-With') == 'XMLHttpRequest':
        user_id = request.POST.get('user_id')

        try:
            user = get_object_or_404(User, pk=user_id)

            # 3. 禁止禁用自己
            if user.id == request.user.id:
                return JsonResponse({
                    'success': False,
                    'message': '不能禁用当前登录的管理员！'  # 错误信息返回前端
                })

            # 4. 切换状态
            user.is_active = not user.is_active
            user.save()
            status = "启用" if user.is_active else "禁用"

            return JsonResponse({
                'success': True,
                'message': f'用户 {user.username} 已{status}！'
            })

        except Exception as e:
            return JsonResponse({'success': False, 'message': f'操作失败：{str(e)}'})

        # 非POST/AJAX请求
    return JsonResponse({'success': False, 'message': '请求方式错误！'})
# 用户编辑
def admin_user_edit_ajax(request):
    if not request.user.is_authenticated or request.user.role != 2:
        return JsonResponse({'success': False, 'message': '仅管理员可操作！'})

    if request.method == 'POST' and request.headers.get('X-Requested-With') == 'XMLHttpRequest':
        user_id = request.POST.get('user_id')
        username = request.POST.get('username')
        email = request.POST.get('email', '')
        try:
            user = get_object_or_404(User, id=user_id)

            # 校验用户名是否重复（排除当前用户）
            if User.objects.filter(username=username).exclude(id=user_id).exists():
                return JsonResponse({'success': False, 'message': '用户名已存在！'})

            # 更新用户信息
            user.username = username
            user.email = email
            user.save()

            return JsonResponse({'success': True, 'message': '修改成功！'})
        except Exception as e:
            return JsonResponse({'success': False, 'message': f'修改失败：{str(e)}'})

    return JsonResponse({'success': False, 'message': '请求方式错误！'})

# 后台首页
def admin_index(request):
    # 1.验证是否登录
    if not request.user.is_authenticated:
        messages.error(request,'请先登录!')
        return redirect('/login/')
    # 2.验证是否是管理员
    if request.user.role !=2:
        messages.error(request,'仅管理员可访问后台!')
        return redirect('/index/')
    return render(request,'admin_index.html',{'user':request.user})
# 1. 用户管理页面
def admin_user_manage(request):
    # 权限验证
    if not request.user.is_authenticated or request.user.role != 2:
        messages.error(request, "仅管理员可访问！")
        return redirect('/login/')
    # 获取所有用户列表（供页面展示）
    users = User.objects.all().order_by('-date_joined')
    return render(request, 'admin_user_manage.html', {'user': request.user, 'users': users})


# 注册
def register(request):
    if request.method == "POST":
        username = request.POST['username']
        password = request.POST['password']
        password2 = request.POST['password2']
        if not username or not password or not password2:
            messages.error(request,'用户名和密码不能为空！',extra_tags='error')
            return render(request,'register.html')
        if len(password) < 6:
            messages.error(request,'密码长度不能少于6位！',extra_tags='error')
            return render(request,'register.html')
        if password != password2:
            messages.error(request,'两次输入的密码不一致！',extra_tags='error')
            return render(request,'register.html')
        if User.objects.filter(username=username).exists():
            messages.error(request,'用户名已存在',extra_tags='error')
            return render(request,'register.html')
        user = User.objects.create_user(username=username, password=password,role=1)
        messages.success(request,'注册成功！请登录')
        return redirect('/login/')
    return render(request,'register.html')

def user_login(request):
    if request.method == 'POST':
        username = request.POST.get('username')
        password = request.POST.get('password')
        user = authenticate(request, username=username, password=password)
        if user:
            login(request, user)
            if user.role==2:
                return redirect('/admin_index/')
            else:
                return redirect('/index/')
        else:
            messages.error(request,'账号或密码错误,请重新输入!')
            return render(request, 'login.html')
    else :
        return render(request, 'login.html')
def user_logout(request):
    logout(request)
    return redirect('/login/')
def admin_page(request):
    if not request.user.is_authenticated:
        return redirect('/login/')
    return render(request, 'admin_index.html',{'user':request.user})
