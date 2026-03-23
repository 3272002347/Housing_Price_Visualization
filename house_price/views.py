import pandas as pd
from django.contrib.auth.decorators import login_required
from django.contrib.auth.hashers import make_password
from django.core.paginator import Paginator, PageNotAnInteger, EmptyPage
from django.db import transaction
from django.db.models import Q
from django.shortcuts import render, redirect, get_object_or_404
from django.contrib.auth import authenticate, login, logout
from django.contrib import messages
from django.http import HttpResponse,JsonResponse
from openpyxl import Workbook

from .models import User,HousePriceData
from django.template.loader import render_to_string

from .utils import clean_area, clean_house_age, clean_total_price

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

    # 生成响应
    response = HttpResponse(content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    response["Content-Disposition"] = 'attachment; filename="房价数据导入模板.xlsx"'
    wb.save(response)
    return response

#数据下载
@login_required
def data_download_template(request):
    if request.user.role != 2:
        messages.error(request, "仅管理员可访问")
        return redirect('/login/')

#数据导出
@login_required
def data_export(request):
    if request.user.role != 2:
        messages.error(request, "仅管理员可访问")
        return redirect('/login/')

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
        unit_price = 0.0  # 默认值
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
            "area_size": house.area_size,
            "orientation": house.orientation,
            "decoration": house.decoration,
            "floor_info": house.floor_info,
            "house_age": house.house_age,
            "structure_type": house.structure_type,
            "total_price":house.total_price,
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
        "city": city,  # 回显城市筛选
        "decoration": decoration,  # 回显装修筛选
        # 城市列表（用于筛选下拉框）
        "city_list": HousePriceData.objects.values_list("city", flat=True).distinct(),
        "page_range": page_range,  # 新增：精简后的页码列表
        "decoration_list":["精装", "简装", "毛坯"],
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


# 3. 可视化配置页面（基础框架）
def admin_visual_config(request):
    # 权限验证
    if not request.user.is_authenticated or request.user.role != 2:
        messages.error(request, "仅管理员可访问！")
        return redirect('/login/')
    # 后续添加可视化配置逻辑
    return render(request, 'admin_visual_config.html', {'user': request.user})
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
def index(request):
    if not request.user.is_authenticated:
        return redirect('/login/')
    return render(request, 'index.html',{'user':request.user})
def user_logout(request):
    logout(request)
    return redirect('/login/')
def admin_page(request):
    if not request.user.is_authenticated:
        return redirect('/login/')
    return render(request, 'admin_index.html',{'user':request.user})
